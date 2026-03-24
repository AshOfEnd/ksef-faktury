"""
main.py
-------
Główny skrypt projektu KSeF API 2.0.
Pobiera faktury wystawione i otrzymane, zapisuje do Excel.

Uruchomienie:
    python main.py

Wymagana konfiguracja w pliku .env (skopiuj z .env.example).
"""

from datetime import datetime, timedelta, timezone
import logging
import os
import sys
import time
from pathlib import Path
from dotenv import load_dotenv
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from ksef_auth import KSeFAuth, KSeFAuthError
from ksef_invoices import SUBJECT_TYPE_LABELS, KSeFInvoices, KSeFInvoiceError
import pyodbc

# ------------------------------------------------------------------
# Konfiguracja logowania
# ------------------------------------------------------------------

if getattr(sys, 'frozen', False):
    # Ścieżka, gdy program jest uruchomiony jako .exe
    base_dir = Path(sys.executable).parent
else:
    # Ścieżka, gdy program jest uruchomiony jako skrypt .py
    base_dir = Path(__file__).parent

    
new_folder=base_dir/"logs"

new_folder.mkdir(parents=True,exist_ok=True)

log_filename=f"{datetime.now().strftime('%Y-%m-%d')}.log"
log_path=new_folder/log_filename
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
    handlers=
    [
        logging.FileHandler(log_path,encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ------------------------------------------------------------------
# Wczytaj .env
# ------------------------------------------------------------------
env_path = base_dir / ".env"
load_dotenv(dotenv_path=env_path)
tytul = "START PROGRAMU"
logger.info(tytul.center(60, "="))
logger.info(f"Szukam pliku .env w: {env_path.absolute()}")

NIP           = os.getenv("KSEF_NIP", "").strip()
TOKEN         = os.getenv("KSEF_TOKEN", "").strip()
ENV           = os.getenv("KSEF_ENV", "test").strip().lower()
DATE_FROM_STR = os.getenv("DATE_FROM", "2025-01-01").strip()
DATE_TO_STR   = os.getenv("DATE_TO",   "2025-12-31").strip()
PAGE_SIZE     = int(os.getenv("PAGE_SIZE", "100"))
AUTO_YESTERDAY = os.getenv("AUTO_YESTERDAY", "true").lower() == "true"
BASE_URLS = {
    "test": "https://api-test.ksef.mf.gov.pl/api/v2",
    "prod": "https://api.ksef.mf.gov.pl/api/v2",
}
BASE_URL = BASE_URLS.get(ENV, BASE_URLS["test"])
SQL_STRING=os.getenv("DB_CONNECTION_STRING","").strip()
OUTPUT_FILE = Path(__file__).parent


# ------------------------------------------------------------------
# Kolumny które chcemy pokazać w Excelu
# Klucze odpowiadają polom zwracanym przez KSeF API
# ------------------------------------------------------------------
INVOICE_COLUMNS = [
    ("_typ",          "Typ faktury"),
    ("ksefNumber",    "Numer KSeF"),
    ("invoiceNumber", "Numer faktury"),
    ("invoiceType",   "Rodzaj faktury"),
    ("issueDate",     "Data wystawienia"),
    ("invoicingDate", "Data przyjęcia w KSeF"),
    ("seller_name",   "Wystawca — nazwa"),
    ("seller_nip",    "Wystawca — NIP"),
    ("buyer_name",    "Nabywca — nazwa"),
    ("buyer_nip",     "Nabywca — NIP"),
    ("netAmount",     "Kwota netto"),
    ("vatAmount",     "Kwota VAT"),
    ("grossAmount",   "Kwota brutto"),
    ("currency",      "Waluta"),
]

def get_ksef_clients() -> list[dict]:
    """Pobiera listę NIP i TOKEN dla maksymalnie 5 firm z .env."""
    clients = []
    for i in range(1, 6):
        nip = os.getenv(f"KSEF_NIP_{i}", "").strip()
        token = os.getenv(f"KSEF_TOKEN_{i}", "").strip()
        if nip and token:
            clients.append({"nip": nip, "token": token, "id": i})
    return clients



def get_date_range() -> tuple [str,str]:
    logger.info(f"wartość AUTO_YESTERDAY= {AUTO_YESTERDAY}")
    if AUTO_YESTERDAY:
        yesterday=(datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        logger.info(f"TRYB: AUTO_YESTERDAY=true. pobieram dane z dnia: {yesterday}")
        return yesterday, yesterday
    
    else:
        d_from=DATE_FROM_STR
        d_to=DATE_TO_STR
        logger.info(f"TRYB RĘCZNY. pobieram dane z dnia: {d_from} -> {d_to} ")
        return d_from, d_to


def flatten_invoice(inv: dict) -> dict:
    """
    Spłaszcza zagnieżdżone pola faktury do jednopoziomowego słownika.
    KSeF API 2.0 zwraca seller.name, seller.nip, buyer.name, buyer.identifier.value itp.
    """
    flat: dict = {}

    for key, _ in INVOICE_COLUMNS:
        if key == "seller_name":
            flat[key] = (inv.get("seller") or {}).get("name", "")
        elif key == "seller_nip":
            flat[key] = (inv.get("seller") or {}).get("nip", "")
        elif key == "buyer_name":
            flat[key] = (inv.get("buyer") or {}).get("name", "")
        elif key == "buyer_nip":
            # buyer.identifier.value (NIP nabywcy zagnieżdżony głębiej)
            flat[key] = ((inv.get("buyer") or {}).get("identifier") or {}).get("value", "")
        else:
            flat[key] = inv.get(key, "")

    return flat



def format_date(date: str):

    dt = datetime.fromisoformat(date)
    dt_utc = dt.astimezone(timezone.utc).replace(tzinfo=None)

    return dt_utc


def save_to_sql(invoices: list[dict], subject_type: str, nip_firmy: str):
    """
    Zapisuje listę faktur do bazy SQL Server. 
    Wykorzystuje UNIQUE CONSTRAINT w bazie do ignorowania duplikatów.
    """
    if not invoices:
        return 0,0


    conn_str=SQL_STRING
    

    sql = """
    INSERT INTO T_faktury_from_ksef
    (TypPobrania, KSeFNumber, InvoiceNumber, invoiceType, IssueDate, invoicingDate, 
     seller_name, seller_nip, buyer_name, buyer_nip, NetAmount, vatAmount, GrossAmount, currency)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    added_count = 0
    skipped_count = 0

    try:
       
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            
            for inv in invoices:
          
                data = flatten_invoice(inv)
              
                vals = (
                    subject_type,                # TypPobrania
                    data.get('ksefNumber'),       # KSeFNumber
                    data.get('invoiceNumber'),    # InvoiceNumber
                    data.get('invoiceType'),      # invoiceType
                    data.get('issueDate'),        # IssueDate data wystawienia 
                    format_date(data.get('invoicingDate')),    # invoicingDate data przyjecia w ksef
                    data.get('seller_name'),      # seller_name
                    data.get('seller_nip'),       # seller_nip
                    data.get('buyer_name'),       # buyer_name
                    data.get('buyer_nip'),        # buyer_nip
                    data.get('netAmount'),        # NetAmount
                    data.get('vatAmount'),        # vatAmount
                    data.get('grossAmount'),      # GrossAmount
                    data.get('currency'),         # currency
                )

                try:
                    cursor.execute(sql, vals)
                    added_count += 1
                except pyodbc.IntegrityError:
                    # odrzucenie powtórek
                    ksef_ref = data.get('ksefNumber', 'NIEZNANY')
                    logger.warning(f"  [SQL] Pominięto powtórzoną fakturę: {ksef_ref}")
                    skipped_count += 1
                    continue
                except Exception as e:
                    logger.error(f"  [SQL] Błąd zapisu faktury {subject_type} {data.get('ksefNumber')}: {e}")
            
           
            conn.commit()
            
            if added_count > 0 or skipped_count > 0:
                logger.info(f"  [SQL] Firma {nip_firmy}: Nowe: {added_count}, Pominięte duplikaty: {skipped_count} ({subject_type})")

    except Exception as e:
        logger.error(f"  [SQL] Krytyczny błąd połączenia z bazą danych: {e}")

    return skipped_count, added_count



# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def main() -> None:

    logger.info("=" * 55)
    logger.info("  KSeF API 2.0 — Pobieranie faktur")
    logger.info("=" * 55)

    date_from_raw, date_to_raw = get_date_range()
    date_from = KSeFInvoices.to_iso(date_from_raw, end_of_day=False)
    date_to   = KSeFInvoices.to_iso(date_to_raw,   end_of_day=True)
    logger.info(f"Zakres dat ISO: {date_from}  →  {date_to}")
    logger.info("\n")

    sql_added=0
    sql_skipped=0
    global_stats = {subject: 0 for subject in SUBJECT_TYPE_LABELS.keys()}


    companies = get_ksef_clients()
    if not companies:
        logger.error("Brak skonfigurowanych firm (KSEF_NIP_1...) w pliku .env!")
        sys.exit(1)

    for co in companies:
        
        logger.info(f"=" * 55)
        logger.info(f" PRZETWARZANIE FIRMY: {co['nip']} (Klient #{co['id']})")
        logger.info("=" * 55)
        logger.info("\n")

        faktury_lacznie=0
        # ------------------------------------------------------------------
        # 1. Uwierzytelnienie
        # ------------------------------------------------------------------
        auth = KSeFAuth(nip=co['nip'], ksef_token=co['token'], env=ENV)
        try:
            auth.authenticate()
        except KSeFAuthError as e:
            logger.error(f"Błąd uwierzytelnienia: {e}")
            sys.exit(1)

        auth_headers = auth.get_auth_headers()
        time.sleep(1)  # krótkie opóźnienie po uwierzytelnieniu

    
        # auth przekazany do klienta — obsłuży wygaśnięcie tokena (401) automatycznie
        client = KSeFInvoices(base_url=BASE_URL, auth_headers=auth_headers, page_size=PAGE_SIZE, auth=auth)
        
        try:
            for subject,label in SUBJECT_TYPE_LABELS.items():

                logger.info(f"\n--- FAKTURY {label.upper()} ---")

                faktury = client.fetch_all(subject, date_from_raw, date_to_raw,auto_mode=AUTO_YESTERDAY)

                faktury_lacznie+=len(faktury)
                global_stats[subject]+=len(faktury)
                skipped,added=save_to_sql(faktury,label,co["nip"])
                sql_added+=added
                sql_skipped+=skipped

            logger.info(" ")
            logger.info(f"  Łącznie wystawione i otrzymane: {faktury_lacznie}")
            logger.info("=" * 55)
            logger.info("\n")

        except KSeFInvoiceError as e:
            logger.error(f"Błąd pobierania faktur: {e}")
            sys.exit(1)


    logger.info("="*55)
    logger.info("SESSION SUMMARY:")
    logger.info(f"faktur wystawionych (wszystkie firmy): {global_stats['Subject1']}")
    logger.info(f"faktur otrzymanych (wszystkie firmy): {global_stats['Subject2']}")
    logger.info(f"łącznie faktur: {global_stats['Subject2']+ global_stats['Subject1']}")
    logger.info(f"łącznie dodano do bazy: {sql_added}")
    logger.info(f"łącznie pominięto: {sql_skipped}")
    logger.info("="*55)
    logger.info("\n")
    logger.info("\n")
    sys.exit(0)



if __name__ == "__main__":
    main()
