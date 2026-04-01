"""
ksef_invoices.py
----------------
Moduł pobierania listy faktur z KSeF API 2.0.

Endpoint: POST /invoices/query/metadata
Limit produkcyjny: 20 req/h (sliding window).
Skrypt automatycznie dzieli długi zakres na okna 3-miesięczne i czeka
185 s między oknami, aby zmieścić się w limicie 20 req/h.
"""

import json
import logging
import time
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import requests
import io, zipfile, time
import xml.etree.ElementTree as ET
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives import padding

logger = logging.getLogger(__name__)

# Limit produkcyjny: 20 req/h → 3600/20 = 180 s/req; +5 s margines bezpieczeństwa
SLEEP_BETWEEN_WINDOWS = 185  # s

SUBJECT_TYPE_LABELS = {
    "Subject1": "Wystawione (sprzedaż)",  
    "Subject2": "Otrzymane (zakupy/koszty)",
}


class KSeFInvoiceError(Exception):
    pass


class KSeFInvoices:

    MAX_OFFSET = 10000

    def __init__(self, base_url: str, auth_headers: dict, page_size: int = 100, auth=None,logger_base=None,logger_exports=None):
        self.base_url = base_url
        self.auth_headers = auth_headers
        self.page_size = min(max(page_size, 1), 1000)
        self.auth = auth  # KSeFAuth — do auto-odświeżenia tokena przy 401
        self.logger_base=logger_base
        self.logger_exports=logger_exports



    def _initiate_export(self, subject_type: str, date_from: str, date_to: str) -> str:
        url = f"{self.base_url}/invoices/exports"
        enc_payload, aes_key, iv = self.auth.get_export_encryption_payload()
        self._last_aes_key = aes_key
        self._last_iv = iv
        
        body = {
            "encryption": enc_payload,
            "onlyMetadata": True,
            "filters": {
                "subjectType": subject_type,
                "dateRange": {
                    "dateType": "PermanentStorage",
                    "from": date_from,
                    "to": date_to
                }
            }
        }

        resp = requests.post(url, json=body, headers=self.auth_headers, timeout=60)
        
        
        if resp.status_code != 201:
            self.logger_exports.error(f"Błąd KSeF {resp.status_code}: {resp.text}")
            resp.raise_for_status()

        data = resp.json()
        token = data.get("referenceNumber")
        
        if not token:
            self.logger_exports.error(f"KSeF nie zwrócił tokena! Odpowiedź: {data}")
            raise KSeFInvoiceError("Brak acquisitionToken w odpowiedzi KSeF")

        self.logger_exports.info(f"Zlecono eksport, token: {token}")
        return token


    def _wait_for_export_ready(self, token: str) -> dict:
        url = f"{self.base_url}/invoices/exports/{token}"
        # 120 prób co 30s = 1 godzina czekania
        max_attempts = 120
        interval = 30
        
        for attempt in range(1, max_attempts + 1):
            resp = requests.get(url, headers=self.auth_headers, timeout=30)
            self._raise_for_status(resp, "Błąd sprawdzania statusu eksportu")

            data = resp.json()
            status = data.get("status", {}).get("code")

            if status == 200:
                self.logger_exports.info(f"Eksport {token} gotowy!")
                return data

            if status == 105:
                raise KSeFInvoiceError(f"Błąd eksportu (105): {data}")

            self.logger_exports.info(f"Export {token} w toku [{attempt}/{max_attempts}] (status: {status})")
            time.sleep(interval)

        raise KSeFInvoiceError(f"Timeout eksportu po 1h dla tokena {token}")

    
    def _download_metadata_file(self, export_data: dict) -> list[dict]:
       
        parts = export_data.get("package", {}).get("parts", [])
        if not parts:
            self.logger_exports.info("Brak faktur w tej paczce, zwracamy pustą listę")
            return []

        all_metadata = []

        for part in parts:
            url = part.get("url")
            resp = requests.get(url, timeout=120)
            self._raise_for_status(resp, "Błąd pobierania binarnej paczki")

            # Deszyfrowanie AES-256-CBC Twoim kluczem i IV
            cipher = Cipher(algorithms.AES(self._last_aes_key), modes.CBC(self._last_iv))
            decryptor = cipher.decryptor()
            decrypted_padded = decryptor.update(resp.content) + decryptor.finalize()
            
            # Zdjęcie paddingu PKCS7
            unpadder = padding.PKCS7(128).unpadder()
            decrypted_zip = unpadder.update(decrypted_padded) + unpadder.finalize()


            # Rozpakowanie ZIP i zamiana XML na listę słowników
            
            with zipfile.ZipFile(io.BytesIO(decrypted_zip)) as z:
                for file_name in z.namelist():
                    with z.open(file_name) as f:
                        raw_content = f.read()
                        self.logger_exports.info(f"Pobrano plik: {file_name} ({len(raw_content)} bajtów)")

                        # Skoro plik to _metadata.json, używamy json.loads
                        if file_name.endswith('.json'):
                            data = json.loads(raw_content)
                            
                            # KSeF w JSONie zazwyczaj ma listę faktur w polu 'invoices' lub 'invoiceList'
                            # Sprawdź strukturę, ale najczęściej jest to lista obiektów:
                            invoices = data.get("invoices") or data.get("invoiceList") or data
                            
                            if isinstance(invoices, list):
                                for inv in invoices:
                                    # Dodajemy surowy słownik do listy (Twoja baza to przemieli)
                                    # Upewnij się tylko, że pole 'acquisitionDate' tam jest
                                    all_metadata.append(inv)
                            else:
                                self.logger_exports.warning(f"Nieoczekiwana struktura JSON w {file_name}")
                        
                        elif file_name.endswith('.xml'):
                            # Tutaj zostawiasz stary kod XML na wszelki wypadek
                            pass 
        return all_metadata
    


    def _fetch_exports_window(self, subject_type: str, date_from: str, date_to: str) -> tuple[list[dict], str]:
        token = self._initiate_export(subject_type, date_from, date_to)
        export_data = self._wait_for_export_ready(token)
        metadata = self._download_metadata_file(export_data)

        
        if export_data.get("package", {}).get("isTruncated", False):
            new_start = export_data.get("package", {}).get("lastPermanentStorageDate")
            
        else:
            new_start = export_data.get("package", {}).get("permanentStorageHwmDate")

        return metadata, new_start

    

    def fetch_all_exports(self, subject_type: str, date_from: str, date_to: str) -> list[dict]:
        all_invoices = []
        current_start = date_from

        while current_start < date_to:
            batch, new_start = self._fetch_exports_window(subject_type, current_start, date_to)
            
            if not batch:
                break

            all_invoices.extend(batch)
            self.logger_exports.info(f"old:{current_start} -> {new_start} -> {date_to} ")
            if not new_start:
                break
            
            current_start = new_start
            if new_start == current_start:
                break
            
        return all_invoices



    def _query_page(self, subject_type: str, date_from: str, date_to: str,
                    page_offset: int) -> dict:
        """Pobiera jedną stronę wyników z /invoices/query/metadata"""
        url = f"{self.base_url}/invoices/query/metadata"
        params = {
            "pageSize": self.page_size,
            "pageOffset": page_offset,
        }
        body = {
            "subjectType": subject_type,
            "dateRange": {
                "dateType": "Invoicing", #typ poboru daty aktualnie: Data przyjęcia faktury w systemie KSeF (do dalszego przetrwarzania)
                "from": date_from,
                "to": date_to,
            },
        }
        for attempt in range(1, 6):
            resp = requests.post(url, json=body, params=params, headers=self.auth_headers, timeout=60)
            if resp.status_code == 401 and self.auth is not None:
                # Token wygasł — re-auth w miejscu, bez utraty postępu
                self.logger_base.warning(f"accessToken wygasł (próba {attempt}/5) — ponowne uwierzytelnienie...")
                self.auth.authenticate()
                self.auth_headers = self.auth.get_auth_headers()
                continue
            if resp.status_code == 429:
                # Czytaj Retry-After z nagłówka HTTP (standard); fallback: 185 s
                retry_after = int(resp.headers.get("Retry-After", SLEEP_BETWEEN_WINDOWS)) + 2
                self.logger_base.warning(
                    f"HTTP 429 — rate limit, czekam {retry_after}s "
                    f"(próba {attempt}/5)..."
                )
                time.sleep(retry_after)
                continue
            self._raise_for_status(resp, f"Błąd zapytania o faktury (offset={page_offset}, od={date_from})")
            data = resp.json()
            return data
        raise KSeFInvoiceError(
            f"Błąd zapytania o faktury (offset={page_offset}, od={date_from}) "
            f"— przekroczono limit prób"
        )

    def _fetch_window(self, subject_type: str, date_from: str, date_to: str) -> list[dict]:
        """Pobiera wszystkie strony dla jednego okna czasowego (max 3 miesiące)."""
        label = SUBJECT_TYPE_LABELS.get(subject_type, subject_type)
        all_invoices = []
        offset = 0

        while True:
            data = self._query_page(subject_type, date_from, date_to, offset)
            invoices = data.get("invoices", [])  # KSeF API 2.0: pole "invoices"

            if invoices and offset == 0:
                self.logger_base.info(f"  Okno {date_from[:10]}–{date_to[:10]}: znaleziono faktury ({label})")

            if not invoices:
                break

            for inv in invoices:
                inv["_typ"] = label

            all_invoices.extend(invoices)
            # offset += len(invoices) 
            offset += self.page_size #dla bezpiecznieństwa (API używa pageOffset a nie liczby zwróconych rekordów więc możliwe pominięcie faktury)

            # KSeF API 2.0: paginacja przez hasMore (nie totalCount)
            if not data.get("hasMore", False):
                break

            time.sleep(0.3)  # między stronami paginacji (limit 8 req/s jest OK)

        return all_invoices
    
    def fetch_windows_hours(self, subject_type, date_from, date_to):
        """jeżeli w danym dniu liczba faktur przekroczy limit 10k
          to dzieli je na coraz mniejsze zakresy godzinowe :2 """
        try:
            return self._fetch_window(subject_type,date_from,date_to)
        
        except KSeFInvoiceError as e: 
            if "10 000" not in str(e) and "Window too large" not in str(e):
                raise
        
            dt_from= datetime.strptime(date_from[:19], "%Y-%m-%dT%H:%M:%S")
            dt_to= datetime.strptime(date_to[:19], "%Y-%m-%dT%H:%M:%S")

            WINDOW_STEPS= [24,12,6,1] # fallback na wypadek gdyby dziennie było więcej niż 10k faktur na daną firmę 

            all_invoices=[]
            current=dt_from

            while current< dt_to:
                succes=False
                for step in WINDOW_STEPS: 
                    candidate_end=min(current + timedelta(hours=step), dt_to)
                    win_from= self.to_iso(current,end_of_day=False)
                    win_to = self.to_iso(candidate_end,end_of_day=True)

                    try:
                        invoices=self._fetch_window(subject_type,win_from,win_to)
                        if len(invoices) >=self.MAX_OFFSET:
                            raise KSeFInvoiceError("Window too large")
                        
                        self.logger_base.info(f"okno {win_from[:19]} -> {win_to[:19]}" f"({len(invoices)} faktur)")

                        all_invoices.extend(invoices)
                        current=candidate_end
                        succes=True
                        break
                    except KSeFInvoiceError as e:

                        if "10 000" not in str(e) and "Window too large" not in str(e):
                            raise

                        self.logger_base.warning(f"za dużo faktur w oknie {win_from[:19]} -> {win_to[:19]}")

                if not succes:
                    raise KSeFInvoiceError(f"Nie można zmiejszyć okna poniżej 1h: {current}")
            
            return all_invoices


    def _fetch_by_splitting(self, subject_type, date_from, date_to):
        """Rekurencyjne dzielenie zakresu dni na pół. tylko jeśli ilość faktur w zakresie >10k."""
        try:

            return self._fetch_window(subject_type, date_from, date_to)
            
        except KSeFInvoiceError as e:
            if "10 000" not in str(e) and "Window too large" not in str(e):
                raise
            

            self.logger_base.warning(f"przekroczono limit 10k. {date_from[:10]} - {date_to[:10]} stosuję podział na mniejsze okna.")
            dt_from = datetime.strptime(date_from[:10], "%Y-%m-%d")
            dt_to   = datetime.strptime(date_to[:10],   "%Y-%m-%d")
            
            if dt_from == dt_to:
                # awaryjny fallback, zmiana na tryb godzinowy
                self.logger_base.warning(f"Dzień {date_from[:10]} ma > 10k faktur. Przełączam na tryb godzinowy.")
                return self._fetch_by_hours(subject_type, date_from, date_to)

           
            mid = dt_from + (dt_to - dt_from) // 2
            left_to = self.to_iso(mid, end_of_day=True)
            right_from = self.to_iso(mid + timedelta(days=1), end_of_day=False)
            
            return self._fetch_by_splitting(subject_type, date_from, left_to) + \
                   self._fetch_by_splitting(subject_type, right_from, date_to)

    @staticmethod
    def _count_windows(dt_from: datetime, dt_to: datetime) -> int:
        count = 0
        start = dt_from
        while start <= dt_to:
            end = min(start + relativedelta(months=3) - timedelta(days=1), dt_to)
            count += 1
            start = end + timedelta(days=1)
        return count

    def fetch_all(self, subject_type: str, date_from: str, date_to: str, auto_mode: bool=True) -> list[dict]:
        """
        Pobiera wszystkie faktury w zakresie dat, automatycznie dzieląc
        na okna 3-miesięczne (limit API: 20 req/h).
        Między oknami czeka SLEEP_BETWEEN_WINDOWS sekund.
        """
        label = SUBJECT_TYPE_LABELS.get(subject_type, subject_type)

        dt_from = datetime.strptime(date_from[:10], "%Y-%m-%d")
        dt_to   = datetime.strptime(date_to[:10],   "%Y-%m-%d")

        total_windows = self._count_windows(dt_from, dt_to)
        eta_min = (total_windows * SLEEP_BETWEEN_WINDOWS) // 60

        self.logger_base.info(f"Pobieranie faktur: {label}")
        self.logger_base.info(
            f"Zakres: {date_from[:10]} → {date_to[:10]} | "
            f"{total_windows} okien 3-miesięcznych | "
            f"szac. czas: ~{eta_min} min"
        )

        all_invoices = []
        window_start = dt_from
        window_num   = 0

        while window_start <= dt_to:
            window_num += 1
            window_end = min(window_start + relativedelta(months=3) - timedelta(days=1), dt_to)

            w_from = self.to_iso(window_start.date(), end_of_day=False)
            w_to   = self.to_iso(window_end.date(),   end_of_day=True)

            if auto_mode:
                batch = self.fetch_windows_hours(subject_type, w_from, w_to)
            else:
                batch=self._fetch_by_splitting(subject_type,w_from,w_to)
                
            all_invoices.extend(batch)

            window_start = window_end + timedelta(days=1)

            if window_start <= dt_to:
                remaining_windows = total_windows - window_num
                remaining_min     = (remaining_windows * SLEEP_BETWEEN_WINDOWS) // 60
                self.logger_base.info(
                    f"  [{window_num}/{total_windows}] Czekam {SLEEP_BETWEEN_WINDOWS}s "
                    f"(limit 20 req/h) — pozostało ~{remaining_min} min..."
                )
                time.sleep(SLEEP_BETWEEN_WINDOWS)

        self.logger_base.info(f"✓ Łącznie pobrano: {len(all_invoices)} faktur ({label})")
        return all_invoices


    @staticmethod
    def to_iso(d: str | date | datetime, end_of_day: bool = False) -> str:
        if isinstance(d, str):
            d = datetime.strptime(d[:10], "%Y-%m-%d").date()
        if isinstance(d, date) and not isinstance(d, datetime):
            if end_of_day:
                d = datetime(d.year, d.month, d.day, 23, 59, 59)
            else:
                d = datetime(d.year, d.month, d.day, 0, 0, 0)
        return d.strftime("%Y-%m-%dT%H:%M:%S.000Z")

    @staticmethod
    def _raise_for_status(resp: requests.Response, context: str) -> None:
        if not resp.ok:
            try:
                detail = resp.json()
            except Exception:
                detail = resp.text[:500]
            raise KSeFInvoiceError(f"{context} — HTTP {resp.status_code}: {detail}")
