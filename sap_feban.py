"""
SAP FEBAN Automation
Automates FEBAN transaction: enter company code, export to Excel,
flag matching +/- pairs, detect payment runs, and cross-check FBL3N.
"""

import sys
import time
import logging
from datetime import datetime
from typing import Optional

try:
    import win32com.client
except ImportError:
    sys.exit("Wymagane: pip install pywin32")

try:
    import pandas as pd
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Wymagane: pip install pandas openpyxl")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("sap_feban.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Konfiguracja
# ---------------------------------------------------------------------------

COMPANY_CODE = "2052"
OUTPUT_DIR = "."          # katalog zapisu pliku Excel
SAP_SESSION_INDEX = 0     # indeks sesji SAP (0 = pierwsza)


# ---------------------------------------------------------------------------
# Połączenie z SAP GUI
# ---------------------------------------------------------------------------

def get_sap_session(session_index: int = SAP_SESSION_INDEX):
    """
    Zwraca aktywną sesję SAP GUI.
    SAP Logon musi być uruchomiony i zalogowany przed wywołaniem.
    """
    log.info("Łączenie z SAP GUI Scripting...")
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
    except Exception:
        raise RuntimeError(
            "Nie można połączyć z SAP GUI. Upewnij się, że SAP Logon jest "
            "uruchomiony i włączony jest SAP GUI Scripting "
            "(Options → Accessibility & Scripting → Scripting → Enable)."
        )

    application = sap_gui_auto.GetScriptingEngine
    if application.Children.Count == 0:
        raise RuntimeError("Brak aktywnych połączeń SAP. Zaloguj się najpierw.")

    connection = application.Children(0)
    if connection.Children.Count <= session_index:
        raise RuntimeError(f"Brak sesji o indeksie {session_index}.")

    session = connection.Children(session_index)
    log.info("Połączono z sesją SAP: %s", session.Info.SystemName)
    return session


# ---------------------------------------------------------------------------
# Pomocnicze funkcje SAP
# ---------------------------------------------------------------------------

def run_transaction(session, tcode: str) -> None:
    """Przechodzi do podanej transakcji przez pole komend."""
    log.info("Uruchamianie transakcji: %s", tcode)
    session.StartTransaction(tcode)
    time.sleep(1)


def set_field(session, field_id: str, value: str) -> None:
    """Ustawia wartość pola po jego ID."""
    try:
        session.FindById(field_id).Text = value
    except Exception as exc:
        raise RuntimeError(f"Nie można ustawić pola '{field_id}': {exc}") from exc


def click_button(session, button_id: str) -> None:
    """Klika przycisk po jego ID."""
    try:
        session.FindById(button_id).Press()
    except Exception as exc:
        raise RuntimeError(f"Nie można kliknąć przycisku '{button_id}': {exc}") from exc


def press_enter(session) -> None:
    session.FindById("wnd[0]").SendVKey(0)   # VKey 0 = Enter
    time.sleep(0.5)


def press_f8(session) -> None:
    """F8 = Execute (uruchom)."""
    session.FindById("wnd[0]").SendVKey(8)
    time.sleep(1)


def get_statusbar_text(session) -> str:
    try:
        return session.FindById("wnd[0]/sbar").Text
    except Exception:
        return ""


def dismiss_popup(session) -> None:
    """Zamknij ewentualny popup (np. błąd, ostrzeżenie)."""
    try:
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
        time.sleep(0.3)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Krok 1 – Wejście do FEBAN i uzupełnienie ekranu selekcji
# ---------------------------------------------------------------------------

def open_feban(session, company_code: str = COMPANY_CODE) -> None:
    """
    Otwiera transakcję FEBAN i wpisuje kod spółki.
    Pola mogą się różnić w zależności od wersji SAP – dostosuj ID jeśli trzeba.
    """
    run_transaction(session, "FEBAN")
    time.sleep(1)

    # Sprawdź status po wejściu do transakcji
    status = get_statusbar_text(session)
    if "does not exist" in status.lower() or "nicht vorhanden" in status.lower():
        raise RuntimeError(f"Transakcja FEBAN niedostępna: {status}")

    log.info("Wpisywanie kodu spółki: %s", company_code)

    # Standardowe ID pól ekranu selekcji FEBAN
    # (wnd[0]/usr/ctxtRF05L-BUKRS = Company Code)
    _field_ids_to_try = [
        "wnd[0]/usr/ctxtRF05L-BUKRS",       # typowe w nowszych wersjach
        "wnd[0]/usr/txtRF05L-BUKRS",         # alternatywne
        "wnd[0]/usr/ctxtFEBAN-BUKRS",
    ]

    field_set = False
    for fid in _field_ids_to_try:
        try:
            session.FindById(fid).Text = company_code
            field_set = True
            log.info("Pole kodu spółki znalezione: %s", fid)
            break
        except Exception:
            continue

    if not field_set:
        # Metoda awaryjna: spróbuj użyć GuiTextField przez indeks
        try:
            _set_field_by_label(session, "Company Code", company_code)
            field_set = True
        except Exception:
            pass

    if not field_set:
        log.warning(
            "Nie udało się automatycznie ustawić kodu spółki. "
            "Ustaw go ręcznie i naciśnij Enter w SAP, a skrypt wznowi działanie."
        )
        input("Naciśnij ENTER tutaj po ręcznym ustawieniu kodu spółki w SAP...")

    press_enter(session)
    time.sleep(1)

    status = get_statusbar_text(session)
    log.info("Status po wpisaniu spółki: '%s'", status)


def _set_field_by_label(session, label_text: str, value: str) -> None:
    """
    Awaryjne wyszukiwanie pola po etykiecie – iteruje po wszystkich elementach okna.
    Używane gdy standardowe ID pola nie pasuje do wersji SAP.
    """
    main_wnd = session.FindById("wnd[0]")
    # GuiComponentCollection
    for i in range(main_wnd.Children.Count):
        child = main_wnd.Children(i)
        _search_and_set(child, label_text, value)


def _search_and_set(element, label_text: str, value: str, depth: int = 0) -> bool:
    if depth > 10:
        return False
    try:
        if hasattr(element, "Text") and label_text.lower() in str(getattr(element, "Text", "")).lower():
            element.Text = value
            return True
        for i in range(getattr(element, "Children", type("", (), {"Count": 0})()).Count):
            if _search_and_set(element.Children(i), label_text, value, depth + 1):
                return True
    except Exception:
        pass
    return False


# ---------------------------------------------------------------------------
# Główna funkcja wejściowa (Krok 1)
# ---------------------------------------------------------------------------

def step1_connect_and_open_feban() -> object:
    """
    Łączy z SAP i otwiera FEBAN z kodem spółki 2052.
    Zwraca sesję SAP do dalszego użycia.
    """
    session = get_sap_session()
    open_feban(session, COMPANY_CODE)
    log.info("=== Krok 1 zakończony: FEBAN otwarty, spółka %s wpisana ===", COMPANY_CODE)
    return session


# ---------------------------------------------------------------------------
# Punkt wejścia
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    try:
        session = step1_connect_and_open_feban()
        log.info("Skrypt gotowy do następnego kroku (eksport danych).")
    except RuntimeError as e:
        log.error("BŁĄD: %s", e)
        sys.exit(1)
    except KeyboardInterrupt:
        log.info("Przerwano przez użytkownika.")
        sys.exit(0)
