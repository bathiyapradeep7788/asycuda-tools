"""
trico_fast.py - Edge/Selenium nathuwa (requests) Trico boat note check eka.

Wada karana widiya:
  1. PHPSESSID cookie ekka session ekak hadanawa (browser eke login wela inna session eka)
  2. Container no ekak hama ekakma search request eka yawanawa
  3. Response eke text eka balala GREEN / RED

ONA DEWAL (me tika fill karanna):
  - PHPSESSID           : browser eke DevTools > Application > Cookies walin gannna
  - SEARCH_METHOD/URL/FIELD : DevTools > Network tab eke Search click karama yana request eken
                          (Fetch/XHR filter eka use karanna). Run karanna kalin DEBUG_DUMP=True
                          karala page eka save karala balannath puluwan.

SAVE: export.py wage - results memory eke + progress_log.csv eke thiyagena,
      save karanakota workbook eka ayeth load karala EKA parak save karanawa.
"""
from core.paths import WORK_DIR, DATA_XLSX, PDF_DIR, MERGED_DIR, LOGS_DIR, EXPORTERS_XLSX
import csv
import os
import random
import re
import shutil
import time
from datetime import datetime

import openpyxl
import requests
from openpyxl.styles import PatternFill

# ------------------------------------------------------------------
# SETTINGS
# ------------------------------------------------------------------
file_path = DATA_XLSX
SHEET_NAME = 'CDN DETAILS'
CONTAINER_COL = 24          # Column X
LAST_COLOR_COL = 24         # A..X pata karanawa

BASE = "https://s2.tricologi.net"
PAGE_URL = f"{BASE}/webuser/?option=etc&action=boatnote_check"

# Browser eke login wela inna session cookie eka. Expire unoth aluth ekak danna.
PHPSESSID = os.environ.get("TRICO_PHPSESSID", "")

# --- Search request eka (Network tab eken balala hariyatama fill karanna) ---
SEARCH_METHOD = "POST"
SEARCH_URL = f"{BASE}/webuser/"
SEARCH_PARAMS = {"option": "etc", "action": "boatnote_check_view", "req_type": "raw"}   # rnd eka hama parakma aluthen
SEARCH_FIELD = "cont_numebr"               # site eke field name eka (spelling eka ehemamai)
EXTRA_FIELDS = {}
EXTRA_HEADERS = {
    "X-Requested-With": "XMLHttpRequest",
    "Accept": "text/plain, */*; q=0.01",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "Origin": BASE,
}

PASS_TEXT = "Physical Boat Note received to Trico Logistics Boat Note Dispatcher from EFC Boat Note Office"
FAIL_TEXT = "Boat note not received yet."

DEBUG_DUMP = False   # True -> page eka + palamu response eka debug_html/ ekata save karanawa
TEST_MODE = False    # True -> palamu rows tikak witharai, Excel save karanne naha
TEST_ROWS = 3
SAVE_EVERY = 20
DELAY = 0.3
RETRIES = 2

GREEN = "92D050"
RED = "FF0000"

script_dir = LOGS_DIR
PROGRESS_LOG = os.path.join(script_dir, 'trico_progress_log.csv')

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36"),
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": PAGE_URL,
}


# ------------------------------------------------------------------
# HELPERS (Excel)
# ------------------------------------------------------------------
def make_backup():
    backup_dir = os.path.join(script_dir, 'backups')
    os.makedirs(backup_dir, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    path = os.path.join(backup_dir, f'data_backup_trico_{stamp}.xlsx')
    shutil.copy2(file_path, path)
    print(f"Backup ekak hadala: {path}")


def excel_is_open():
    folder, name = os.path.split(file_path)
    return os.path.exists(os.path.join(folder, '~$' + name))


def safe_save(wb):
    tmp_path = file_path[:-5] + '.tmp.xlsx'
    try:
        wb.save(tmp_path)
        os.replace(tmp_path, file_path)
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def flush_results(results):
    """Excel eka ayeth load karala results apply karala EKA parak save karanawa."""
    if not results:
        return True
    try:
        fresh = openpyxl.load_workbook(file_path)
        ws = fresh[SHEET_NAME]
        for row_idx, color in results.items():
            fill_style = PatternFill(start_color=color, end_color=color, fill_type="solid")
            for col in range(1, LAST_COLOR_COL + 1):
                ws.cell(row=row_idx, column=col).fill = fill_style
        safe_save(fresh)
        return True
    except PermissionError:
        print("  Save karanna bari una (Excel eka open da?). Passe ayeth try karanawa.")
    except Exception as e:
        print(f"  Save error: {type(e).__name__}: {e}. Passe ayeth try karanawa.")
    return False


def log_progress(row_idx, status, container):
    new_file = not os.path.exists(PROGRESS_LOG)
    with open(PROGRESS_LOG, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new_file:
            w.writerow(["time", "row", "status", "container"])
        w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), row_idx, status, container])


def is_green_cell(cell):
    """Kalin green (92D050 wage shade) nam True - e row eka skip karanawa."""
    f = cell.fill
    if f is None or f.fill_type != "solid":
        return False
    c = f.fgColor
    if c.type == "rgb" and isinstance(c.rgb, str) and len(c.rgb) >= 6:
        hexv = c.rgb[-6:]
        r, g, b = int(hexv[0:2], 16), int(hexv[2:4], 16), int(hexv[4:6], 16)
        return g >= 90 and g - r >= 25 and g - b >= 15
    return False


# ------------------------------------------------------------------
# HELPERS (Web)
# ------------------------------------------------------------------
class SessionExpired(Exception):
    pass


def looks_like_login(resp):
    low = resp.text[:5000].lower()
    return ("type=\"password\"" in low or "type='password'" in low
            or "option=login" in resp.url.lower())


def make_session():
    s = requests.Session()
    s.headers.update(HEADERS)
    s.cookies.set("PHPSESSID", PHPSESSID, domain="s2.tricologi.net", path="/")
    return s


def check_session(session):
    """Page eka load karala login thiyenawada balanawa."""
    r = session.get(PAGE_URL, timeout=20)
    r.raise_for_status()
    if looks_like_login(r):
        raise SessionExpired("PHPSESSID expire wela. Browser eken aluth cookie ekak aran script eke danna.")
    if DEBUG_DUMP:
        d = os.path.join(script_dir, 'debug_html')
        os.makedirs(d, exist_ok=True)
        with open(os.path.join(d, 'trico_page.html'), 'w', encoding='utf-8') as f:
            f.write(r.text)
        print(f"Page eka save kala: {d}\\trico_page.html")
    return r


def search_container(session, container):
    data = {SEARCH_FIELD: container, **EXTRA_FIELDS}
    params = {**SEARCH_PARAMS, "rnd": random.random()}
    if SEARCH_METHOD.upper() == "GET":
        resp = session.get(SEARCH_URL, params={**params, **data}, headers=EXTRA_HEADERS, timeout=30)
    else:
        resp = session.post(SEARCH_URL, params=params, data=data, headers=EXTRA_HEADERS, timeout=30)
    if resp.status_code >= 500:
        raise RuntimeError(f"Server error {resp.status_code}")
    resp.raise_for_status()
    if looks_like_login(resp):
        raise SessionExpired("Session eka expire wela.")

    text = resp.text
    if PASS_TEXT in text:
        return "GREEN", text
    if FAIL_TEXT in text:
        return "RED", text
    return "OTHER", text


def search_with_retry(session, container):
    last = None
    for attempt in range(1, RETRIES + 1):
        try:
            return search_container(session, container)
        except SessionExpired:
            raise
        except Exception as e:
            last = e
            time.sleep(1.5 * attempt)
    raise last


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def run_automation():
    if not os.path.exists(file_path):
        print(f"Error: {file_path} සොයාගත නොහැක!")
        return

    if not TEST_MODE:
        if excel_is_open():
            print("Excel eke data.xlsx open wela thiyenawa. Eka close karala ayeth run karanna.")
            return
        make_backup()
    else:
        print("*** TEST MODE: Excel eka wenas karanne naha ***")

    wb = openpyxl.load_workbook(file_path)   # kiyawanna witharai
    ws = wb[SHEET_NAME]

    session = make_session()
    try:
        check_session(session)
    except SessionExpired as e:
        print(e)
        return
    except Exception as e:
        print(f"Page eka load karanna bari una: {e}")
        return

    results = {}
    unsaved = 0
    tested = 0
    debug_saved = False

    try:
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=CONTAINER_COL)
            if is_green_cell(cell):
                continue
            if cell.value is None or not str(cell.value).strip():
                continue
            if TEST_MODE and tested >= TEST_ROWS:
                break

            container = str(cell.value).strip()

            try:
                status, html = search_with_retry(session, container)
            except SessionExpired as e:
                print(f"Row {row}: {e}")
                break
            except Exception as e:
                print(f"Row {row}: ERROR ({e}) - colour karanne naha")
                continue

            print(f"Row {row} | {container} -> {status}")

            if (TEST_MODE or DEBUG_DUMP) and not debug_saved:
                d = os.path.join(script_dir, 'debug_html')
                os.makedirs(d, exist_ok=True)
                out = os.path.join(d, f"trico_row_{row}_{status}.html")
                with open(out, "w", encoding="utf-8") as f:
                    f.write(html)
                print(f"   (response saved: {out}, length={len(html)})")
                debug_saved = True

            if status == "OTHER":
                print("   Pass/Fail text eka response eke naha. Response eka balanna (debug_html).")
                continue

            if TEST_MODE:
                tested += 1
            else:
                results[row] = GREEN if status == "GREEN" else RED
                log_progress(row, status, container)
                unsaved += 1
                if unsaved >= SAVE_EVERY and flush_results(results):
                    print(f"  (Excel ekata save kala: rows {len(results)})")
                    unsaved = 0

            time.sleep(DELAY)

    finally:
        if not TEST_MODE and results:
            if flush_results(results):
                print(f"Excel eka save kala (me run eke rows {len(results)}).")
            else:
                print("!!! Excel ekata save karanna bari una.")
                print(f"    Results tika me file eke thiyenawa: {PROGRESS_LOG}")
        print("Done!")


if __name__ == "__main__":
    run_automation()
