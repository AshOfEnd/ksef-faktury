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

    
new_folder_base = base_dir / "logs"
new_folder_exports = new_folder_base / "exports"

new_folder_base.mkdir(parents=True, exist_ok=True)
new_folder_exports.mkdir(parents=True, exist_ok=True)

log_filename = f"{datetime.now().strftime('%Y-%m-%d')}.log"
log_path_base = new_folder_base / log_filename
log_path_exports = new_folder_exports / log_filename


def create_logger(name, log_path=None, tag="LOG", file_enabled=True):
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)

    formatter = logging.Formatter(
        f"%(asctime)s [{tag}] %(levelname)-8s %(message)s",
        datefmt="%H:%M:%S"
    )

    if not logger.handlers:
        if file_enabled and log_path:
            fh = logging.FileHandler(log_path, encoding='utf-8')
            fh.setFormatter(formatter)
            logger.addHandler(fh)

        sh = logging.StreamHandler()
        sh.setFormatter(formatter)
        logger.addHandler(sh)

    return logger


logger = create_logger("base", log_path_base, "BASE")


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
RANGE_EXPORT = os.getenv("RANGE_EXPORT", "false").lower()=="true"
BIWEEKLY_EXPORT=os.getenv("BIWEEKLY_EXPORT","false").lower()=="true"
BASE_URLS = {
    "test": "https://api-test.ksef.mf.gov.pl/api/v2",
    "prod": "https://api.ksef.mf.gov.pl/api/v2",
}
BASE_URL = BASE_URLS.get(ENV, BASE_URLS["test"])
SQL_STRING=os.getenv("DB_CONNECTION_STRING","").strip()
OUTPUT_FILE = Path(__file__).parent
logger.info("Zmienne wczytane pomyślnie")


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
def is_biweekly_sunday() -> bool:
    today = datetime(2026,4,5)

    #6=niedziela
    if today.weekday() != 6:
        return False

    
    week_number = today.isocalendar()[1]
    return week_number % 2 == 0

def do_export_logs() -> bool:

    if (is_biweekly_sunday() and BIWEEKLY_EXPORT) or RANGE_EXPORT:
        return True
    else:
        return False

logger_exports = create_logger("exports", log_path_exports, "EXPORT",file_enabled=do_export_logs())

def get_biweekly_range() -> tuple[str, str]:
    today = datetime.now()
    end = today
    start = end - timedelta(days=15)

    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


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
    if AUTO_YESTERDAY:
        yesterday=(datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        logger.info(f"TRYB: AUTO_YESTERDAY=true. pobieram dane z dnia: {yesterday}")
        return yesterday, yesterday
    else:
        d_from=DATE_FROM_STR
        d_to=DATE_TO_STR
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

def existing_invoices (dateFrom:str,dateTo:str ):
    
    
    conn_str=SQL_STRING
    sql="""
    SELECT [KSeFNumber], [invoicingDate]
FROM [CONNECT].[dbo].[T_faktury_from_ksef]
WHERE invoicingDate >= ?
  AND invoicingDate <= ?
ORDER BY invoicingDate DESC;
    """
    try:
       
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute(sql,(dateFrom,dateTo))
            rows=cursor.fetchall()
    except Exception as e:
        logger_exports.error(f"  [SQL] Krytyczny błąd połączenia z bazą danych: {e}")

    seen=set(row[0] for row in rows)

    
    return seen


def remove_duplicates(invoices: list[dict], seen:set) -> list[dict]:
    unique_invoices = []
    new=0
    duplicate=0
    for inv in invoices:
        ksef_num = inv.get("ksefNumber")
        if not ksef_num or ksef_num in seen:
            duplicate+=1
            logger_exports.info(f"znaleziono powtórkę na poziomie eliminacji: {ksef_num}")
            continue  
           

        seen.add(ksef_num) 
        unique_invoices.append(inv)
        new+=1
    return unique_invoices,new,duplicate



def save_to_sql(invoices: list[dict], subject_type: str, nip_firmy: str,log):
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
                    log.warning(f"  [SQL] Pominięto powtórzoną fakturę: {ksef_ref}")
                    skipped_count += 1
                    continue
                except Exception as e:
                    log.error(f"  [SQL] Błąd zapisu faktury {subject_type} {data.get('ksefNumber')}: {e}")
            
           
            conn.commit()
            
            if added_count > 0 or skipped_count > 0:
                log.info(f"  [SQL] Firma {nip_firmy}: Nowe: {added_count}, Pominięte duplikaty: {skipped_count} ({subject_type})")

    except Exception as e:
        log.error(f"  [SQL] Krytyczny błąd połączenia z bazą danych: {e}")

    return skipped_count, added_count



# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def main() -> None:
    logger.info("=" * 55)
    logger.info("  KSeF API 2.0 — Pobieranie faktur")
    logger.info("=" * 55)

    # -------------------------------------------------
    # Zbieramy aktywne tryby
    # -------------------------------------------------
    if RANGE_EXPORT and BIWEEKLY_EXPORT:
        
        logger_exports.warning(
            "RANGE_EXPORT=true i BIWEEKLY_EXPORT=true → używam tylko RANGE_EXPORT, "
            "BIWEEKLY_EXPORT zostanie pominięty")
        
        biweekly_enabled = False
    else:
        biweekly_enabled = BIWEEKLY_EXPORT


    ranges = []

    if AUTO_YESTERDAY:
        date_from_raw, date_to_raw = get_date_range()
        ranges.append(("AUTO", date_from_raw, date_to_raw))

    if RANGE_EXPORT:
        ranges.append(("RANGE", DATE_FROM_STR, DATE_TO_STR))

    if biweekly_enabled and is_biweekly_sunday():
        b_from, b_to = get_biweekly_range()
        logger_exports.info(f"TRYB BIWEEKLY: {b_from} -> {b_to}")
        ranges.append(("BIWEEKLY", b_from, b_to))

    if not ranges:
        logger.info("Brak aktywnego trybu — kończę")
        logger_exports.info("Brak aktywnego trybu — kończę")
        return

    seed = set()
    for mode, f_from, f_to in ranges:
        if mode in ("RANGE", "BIWEEKLY"):
            seed = existing_invoices(f_from, f_to)  # tylko raz
            logger_exports.info(f"Pobrano {len(seed)} istniejących faktur z bazy dla trybu {mode}. zakres: {f_from} -> {f_to}")
            break  # tylko pierwszy aktywny tryb
    # Statystyki globalne
    sql_added = 0
    sql_added_export=0
    sql_skipped = 0
    sql_skipped_export=0
    filtr_skipped=0
    filtr_added=0
    export_stats = {subject: 0 for subject in SUBJECT_TYPE_LABELS.keys()}
    meta_stats = {subject: 0 for subject in SUBJECT_TYPE_LABELS.keys()}

    companies = get_ksef_clients()
    if not companies:
        logger.error("Brak skonfigurowanych firm w pliku .env!")
        logger_exports.error("Brak skonfigurowanych firm w pliku .env!")
        sys.exit(1)


    for co in companies:
        logger.info(f"=" * 55)
        logger.info(f" PRZETWARZANIE FIRMY: {co['nip']} (Klient #{co['id']})")
        logger.info("=" * 55)
        if do_export_logs():
            logger_exports.info(f"=" * 55)
            logger_exports.info(f" PRZETWARZANIE FIRMY: {co['nip']} (Klient #{co['id']})")
            logger_exports.info("=" * 55)

        auth = KSeFAuth(nip=co['nip'], ksef_token=co['token'], env=ENV)
        try:
            auth.authenticate()
        except KSeFAuthError as e:
            logger.error(f"Błąd uwierzytelnienia {co['nip']}: {e}")
            continue

        client = KSeFInvoices(
            base_url=BASE_URL,
            auth_headers=auth.get_auth_headers(),
            page_size=PAGE_SIZE,
            auth=auth,
            logger_base=logger,
            logger_exports=logger_exports
        )
        
        try:
            for subject, label in SUBJECT_TYPE_LABELS.items():
                logger.info(f"\n--- FAKTURY {label.upper()} ---")
                if do_export_logs():
                    logger_exports.info(f"\n--- FAKTURY {label.upper()} ---")

            
                for mode, date_from_raw, date_to_raw in ranges:

                    logger.info(f"TRYB: {mode} | {date_from_raw} -> {date_to_raw}")
                    if do_export_logs():
                        logger_exports.info(f"TRYB: {mode} | {date_from_raw} -> {date_to_raw}")

                    exp_from_iso = KSeFInvoices.to_iso(date_from_raw, end_of_day=False)
                    exp_to_iso   = KSeFInvoices.to_iso(date_to_raw, end_of_day=True)

                    # --- PROCES 1: EKSPORT ---
                    if mode in ("RANGE", "BIWEEKLY"):
                        logger_exports.info(f"  [START] Eksport: {date_from_raw} -> {date_to_raw}")
                        
                        faktury_exp = client.fetch_all_exports(subject, exp_from_iso, exp_to_iso)
                        
                        if faktury_exp:
                            clean_faktury_exp, filtr_added,filtr_skipped=remove_duplicates(faktury_exp,seed) 
                            logger_exports.info(f"unikalne faktury: {filtr_added}, powtórzone: {filtr_skipped}")
                            added, skipped = save_to_sql(clean_faktury_exp, label, co["nip"],logger_exports)
                            logger_exports.info(f"  [STOP]  Eksport: przyjęto {added}, pominięto {skipped + filtr_skipped}")
                            
                            export_stats[subject] += len(faktury_exp)
                            sql_added_export += added
                            sql_skipped_export += skipped
                        else:
                            logger_exports.info("  [STOP]  Eksport: brak faktur.")

                    # --- PROCES 2: METADANE ---
                    if mode == "AUTO":
                        logger.info(f"  [START] Metadane: {date_from_raw} -> {date_to_raw}")
                        
                        faktury_meta = client.fetch_all(subject, date_from_raw, date_to_raw, auto_mode=True)
                        
                        if faktury_meta:
                            skipped, added = save_to_sql(faktury_meta, label, co["nip"],logger)
                            logger.info(f"  [STOP]  Metadane: przyjęto {added}, pominięto {skipped}")
                            
                            meta_stats[subject] += len(faktury_meta)
                            sql_added += added
                            sql_skipped += skipped
                        else:
                            logger.info("  [STOP]  Metadane: brak nowych faktur.")

            logger.info("=" * 55 + "\n")
            if do_export_logs(): logger_exports.info("=" * 55 + "\n")

        except KSeFInvoiceError as e:
            logger.error(f"Błąd krytyczny dla firmy {co['nip']}: {e}")
            logger_exports.error(f"Błąd krytyczny dla firmy {co['nip']}: {e}")
            
            continue

    # --- SESSION SUMMARY ---
    total_exp = sum(export_stats.values())
    total_meta = sum(meta_stats.values())

    
    if do_export_logs():
        logger_exports.info("="*55)
        logger_exports.info("SESSION SUMMARY:")
        logger_exports.info(f"łącznie pobrano rekordów: {total_exp}")
        logger_exports.info(f"łącznie dodano do bazy: {sql_added_export}")
        logger_exports.info(f"łącznie pominięto (duplikaty): {filtr_skipped+sql_skipped_export}")
    logger.info("-" * 25)
    logger.info(f"Łącznie pobranych rekordów:    {total_meta}")
    logger.info(f"Łącznie dodano do bazy:        {sql_added}")
    logger.info(f"Łącznie pominięto (duplikaty): {sql_skipped}")
    logger.info("="*55 + "\n")
    
    sys.exit(0)


if __name__ == "__main__":
    main()