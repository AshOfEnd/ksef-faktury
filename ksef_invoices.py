"""
ksef_invoices.py
----------------
Moduł pobierania listy faktur z KSeF API 2.0.

Endpoint: POST /invoices/query/metadata
Limit produkcyjny: 20 req/h (sliding window).
Skrypt automatycznie dzieli długi zakres na okna 3-miesięczne i czeka
185 s między oknami, aby zmieścić się w limicie 20 req/h.
"""

import logging
import time
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
import requests

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

    def __init__(self, base_url: str, auth_headers: dict, page_size: int = 100, auth=None):
        self.base_url = base_url
        self.auth_headers = auth_headers
        self.page_size = min(max(page_size, 1), 1000)
        self.auth = auth  # KSeFAuth — do auto-odświeżenia tokena przy 401

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
                "dateType": "Invoicing",
                "from": date_from,
                "to": date_to,
            },
        }
        for attempt in range(1, 6):
            resp = requests.post(url, json=body, params=params, headers=self.auth_headers, timeout=60)
            if resp.status_code == 401 and self.auth is not None:
                # Token wygasł — re-auth w miejscu, bez utraty postępu
                logger.warning(f"accessToken wygasł (próba {attempt}/5) — ponowne uwierzytelnienie...")
                self.auth.authenticate()
                self.auth_headers = self.auth.get_auth_headers()
                continue
            if resp.status_code == 429:
                # Czytaj Retry-After z nagłówka HTTP (standard); fallback: 185 s
                retry_after = int(resp.headers.get("Retry-After", SLEEP_BETWEEN_WINDOWS)) + 2
                logger.warning(
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
                logger.info(f"  Okno {date_from[:10]}–{date_to[:10]}: znaleziono faktury ({label})")

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
                        
                        logger.info(f"okno {win_from[:19]} -> {win_to[:19]}" f"({len(invoices)} faktur)")

                        all_invoices.extend(invoices)
                        current=candidate_end
                        succes=True
                        break
                    except KSeFInvoiceError as e:

                        if "10 000" not in str(e) and "Window too large" not in str(e):
                            raise

                        logger.warning(f"za dużo faktur w oknie {win_from[:19]} -> {win_to[:19]}")

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
            

            logger.warning(f"przekroczono limit 10k. {date_from[:10]} - {date_to[:10]} stosuję podział na mniejsze okna.")
            dt_from = datetime.strptime(date_from[:10], "%Y-%m-%d")
            dt_to   = datetime.strptime(date_to[:10],   "%Y-%m-%d")
            
            if dt_from == dt_to:
                # awaryjny fallback, zmiana na tryb godzinowy
                logger.warning(f"Dzień {date_from[:10]} ma > 10k faktur. Przełączam na tryb godzinowy.")
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

        logger.info(f"Pobieranie faktur: {label}")
        logger.info(
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
                logger.info(
                    f"  [{window_num}/{total_windows}] Czekam {SLEEP_BETWEEN_WINDOWS}s "
                    f"(limit 20 req/h) — pozostało ~{remaining_min} min..."
                )
                time.sleep(SLEEP_BETWEEN_WINDOWS)

        logger.info(f"✓ Łącznie pobrano: {len(all_invoices)} faktur ({label})")
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
