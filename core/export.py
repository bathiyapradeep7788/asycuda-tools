"""
Export_fast.py  -  Edge/Selenium nathuwa (requests) customs check eka.

Wada karana widiya:
  1. GET  /home       -> csrf_token eka gannawa
  2. POST /dashboard  -> officeCode, serial, cusdecNumber, cusdecYear, consigneeTIN, csrf_token
  3. Response eke "Export release" thiyenawanam GREEN, nathnam RED

SAVE FIX (important):
  openpyxl, image (logo) ekak thiyena workbook ekak EKA PARAKATA WADI GANAKA save karanna beri
  ("ValueError: I/O operation on closed file" + corrupt file). Eka nisa:
    - checks walin passe results memory eke (RESULTS) + progress_log.csv eke thiyaganna
    - save karanakota hama parama workbook eka AYETH LOAD karala, results apply karala,
      EKA parak witharai save karanawa (temp file -> replace)
"""
from core.paths import WORK_DIR, DATA_XLSX, PDF_DIR, MERGED_DIR, LOGS_DIR, EXPORTERS_XLSX
import colorsys
import csv
import os
import re
import shutil
import time
from datetime import datetime

import openpyxl
import pandas as pd
import requests
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import COLOR_INDEX

# ------------------------------------------------------------------
# SETTINGS
# ------------------------------------------------------------------
file_path = DATA_XLSX
script_dir = LOGS_DIR
SHEET_NAME = 'CUSDEC DETAILS'

BASE = "https://services.customs.gov.lk"

TEST_MODE = False    # True -> palamu rows tikak witharai check karanawa, Excel save karanne naha
TEST_ROWS = 3        # TEST_MODE eke check karana rows gana
SAVE_EVERY = 20      # real mode eke rows 20kata parak Excel ekata save karanawa
DELAY = 0.3          # rows athara seconds (site ekata wadi bara nodenna)
RETRIES = 2          # network error ekakin passe ayeth try karana gana

GREEN = "FF00FF00"
RED = "FFFF0000"

PROGRESS_LOG = os.path.join(script_dir, 'progress_log.csv')

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36"),
}


# ------------------------------------------------------------------
# HELPERS (Excel)
# ------------------------------------------------------------------
def get_match_name(text):
    words = str(text).strip().split()
    match_part = " ".join(words[:2]).lower() if len(words) >= 2 else str(text).lower()
    return re.sub(r'[^a-z0-9]', '', match_part)


def make_backup():
    backup_dir = os.path.join(script_dir, 'backups')
    os.makedirs(backup_dir, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(backup_dir, f'data_backup_{stamp}.xlsx')
    shutil.copy2(file_path, backup_path)
    print(f"Backup ekak hadala: {backup_path}")


def excel_is_open():
    folder, name = os.path.split(file_path)
    return os.path.exists(os.path.join(folder, '~$' + name))


def safe_save(wb):
    """Temp file ekakata save karala asal file eka replace karanawa. Fail unoth temp file eka ain karanawa."""
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
    """
    Excel eka AYETH load karala (fresh), results hama ekakma apply karala, EKA parak save karanawa.
    (openpyxl workbook ekak 2 parak save karanna bari image bug eka mehema walakwanawa.)
    Success nam True.
    """
    if not results:
        return True
    try:
        fresh = openpyxl.load_workbook(file_path)
        ws = fresh[SHEET_NAME]
        for row_idx, color in results.items():
            fill_style = PatternFill(start_color=color, end_color=color, fill_type="solid")
            for col in range(1, 31):
                ws.cell(row=row_idx, column=col).fill = fill_style
        safe_save(fresh)
        return True
    except PermissionError:
        print("  Save karanna bari una (Excel eka open da?). Passe ayeth try karanawa.")
    except Exception as e:
        print(f"  Save error: {type(e).__name__}: {e}. Passe ayeth try karanawa.")
    return False


def log_progress(row_idx, status, desc):
    new_file = not os.path.exists(PROGRESS_LOG)
    with open(PROGRESS_LOG, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new_file:
            w.writerow(["time", "row", "status", "details"])
        w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), row_idx, status, desc])


def build_tin_lookup(df_summary):
    lookup = []
    for _, s_row in df_summary.iterrows():
        name_clean = get_match_name(str(s_row.iloc[28]))   # AC Column
        tin = re.sub(r'[^0-9]', '', str(s_row.iloc[29]))   # AD Column
        if name_clean and tin:
            lookup.append((name_clean, tin))
    return lookup


def find_tin(lookup, target):
    for name_clean, tin in lookup:
        if target in name_clean or name_clean in target:
            return tin
    return ""


def clean_number(value):
    """47117.0 wage enawanam .0 ain karanawa"""
    s = str(value).strip()
    if re.fullmatch(r'\d+\.0', s):
        s = s[:-2]
    return s


# ------------------------------------------------------------------
# GREEN DETECTION (koi green shade ekak unath "kalin green" kiyala hithanawa)
# ------------------------------------------------------------------
THEME_NAMES = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3',
               'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']


def theme_color(theme_xml, idx, tint=0.0):
    """Excel theme colour (ex: 'Green, Accent 6') eka RRGGBB walata harawanawa (tint ekath ekka)."""
    if not theme_xml or idx is None or idx >= len(THEME_NAMES):
        return None
    if isinstance(theme_xml, bytes):
        theme_xml = theme_xml.decode('utf-8', 'ignore')
    name = THEME_NAMES[idx]
    block = re.search(rf'<a:{name}>(.*?)</a:{name}>', theme_xml, re.S)
    if not block:
        return None
    m = (re.search(r'lastClr="([0-9A-Fa-f]{6})"', block.group(1))
         or re.search(r'val="([0-9A-Fa-f]{6})"', block.group(1)))
    if not m:
        return None
    hexv = m.group(1)
    if tint:
        r, g, b = [int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4)]
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
        r, g, b = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)
        hexv = "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))
    return hexv.upper()


def fill_hex(cell, theme_xml):
    """Cell eke solid fill colour eka RRGGBB widiyata (fill nathnam None)."""
    f = cell.fill
    if f is None or f.fill_type != "solid":
        return None
    c = f.fgColor
    try:
        if c.type == "rgb" and isinstance(c.rgb, str):
            return c.rgb[-6:].upper()
        if c.type == "indexed" and isinstance(c.indexed, int) and c.indexed < len(COLOR_INDEX):
            return COLOR_INDEX[c.indexed][-6:].upper()
        if c.type == "theme":
            return theme_color(theme_xml, c.theme, c.tint or 0.0)
    except Exception:
        pass
    return None


def is_greenish(hexv):
    """00FF00, 92D050, 00B050, 70AD47, C6EFCE wage green shades walata True; red/yellow/blue/gray walata False."""
    if not hexv or len(hexv) != 6:
        return False
    r, g, b = int(hexv[0:2], 16), int(hexv[2:4], 16), int(hexv[4:6], 16)
    return g >= 90 and g - r >= 25 and g - b >= 15


def is_green_row(ws, row_idx, theme_xml):
    """Row eke A-D cells walin ekak hari green nam, kalin check karapu row ekak."""
    return any(is_greenish(fill_hex(ws.cell(row=row_idx, column=c), theme_xml)) for c in (1, 2, 3, 4))


def print_summary(ws, theme_xml):
    """Run eka patan ganna kalin: kiyadha skip wenne, kiyadha check wenne, A column eke colours mokakda."""
    hist, skip, todo = {}, 0, 0
    for r in range(2, ws.max_row + 1):
        a = ws.cell(row=r, column=1)
        if a.value is None or str(a.value).strip() == "":
            continue
        hist_key = fill_hex(a, theme_xml) or "no fill"
        hist[hist_key] = hist.get(hist_key, 0) + 1
        if is_green_row(ws, r, theme_xml):
            skip += 1
        else:
            todo += 1
    print(f"Rows: {skip + todo} | kalin GREEN (skip karanawa): {skip} | check karanna: {todo}")
    print("Column A colours:", ", ".join(f"{k}={v}" for k, v in sorted(hist.items(), key=lambda kv: -kv[1])))


# ------------------------------------------------------------------
# HELPERS (Web)
# ------------------------------------------------------------------
def extract_csrf(page_html):
    """<input name="csrf_token" value="..."> (value kalin/passe wenas unath hari) nathnam <meta name="csrf-token">"""
    for tag in re.findall(r'<input[^>]*>', page_html, re.I):
        if re.search(r'name=["\']csrf_token["\']', tag, re.I):
            m = re.search(r'value=["\']([^"\']+)["\']', tag, re.I)
            if m:
                return m.group(1)
    m = re.search(r'<meta[^>]*name=["\']csrf[-_]token["\'][^>]*content=["\']([^"\']+)["\']', page_html, re.I)
    if m:
        return m.group(1)
    raise ValueError("csrf_token eka /home page eke hambune naha")


def check_cusdec(session, off_code, cus_no, cus_year, tin):
    """
    Return (status, html)
      status = "GREEN" | "RED"   (site eken answer ekak awa)
      raises exception           (network/csrf awulak - Excel eka wenas karanna epa)
    """
    home = session.get(f"{BASE}/home", timeout=20)
    home.raise_for_status()
    csrf = extract_csrf(home.text)

    data = {
        "officeCode": off_code,
        "serial": "E",
        "cusdecNumber": cus_no,
        "cusdecYear": cus_year,
        "consigneeTIN": tin,
        "csrf_token": csrf,
    }
    resp = session.post(
        f"{BASE}/dashboard",
        data=data,
        headers={"Origin": BASE, "Referer": f"{BASE}/home"},
        timeout=30,
    )
    if resp.status_code >= 500:
        raise RuntimeError(f"Server error {resp.status_code}")

    return ("GREEN" if "Export release" in resp.text else "RED"), resp.text


def check_with_retry(session, *args):
    last_err = None
    for attempt in range(1, RETRIES + 1):
        try:
            return check_cusdec(session, *args)
        except Exception as e:
            last_err = e
            time.sleep(1.5 * attempt)
    raise last_err


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

    # Me workbook eka KIYAWANNA witharai (kisiwita save karanne naha)
    wb = openpyxl.load_workbook(file_path)
    sheet_details = wb[SHEET_NAME]
    df_summary = pd.read_excel(file_path, sheet_name='SUMMARY', skiprows=8)
    tin_lookup = build_tin_lookup(df_summary)

    theme_xml = wb.loaded_theme
    print_summary(sheet_details, theme_xml)

    debug_dir = os.path.join(script_dir, 'debug_html')
    if TEST_MODE:
        os.makedirs(debug_dir, exist_ok=True)

    session = requests.Session()
    session.headers.update(HEADERS)

    results = {}      # row_idx -> colour (me run eke hama result ekakma)
    unsaved = 0
    tested = 0

    try:
        for row_idx in range(2, sheet_details.max_row + 1):
            cell_a = sheet_details.cell(row=row_idx, column=1)

            # --- SKIP LOGIC (kalin green nam (koi green shade ekak hari) skip) ---
            if is_green_row(sheet_details, row_idx, theme_xml):
                continue
            if cell_a.value is None or str(cell_a.value).strip() == "":
                continue

            if TEST_MODE and tested >= TEST_ROWS:
                break

            # --- DATA ---
            off_code = str(cell_a.value).strip()
            cus_no = clean_number(sheet_details.cell(row=row_idx, column=2).value)
            date_raw = str(sheet_details.cell(row=row_idx, column=3).value)
            cus_year = date_raw.split('/')[-1].split(' ')[0] if '/' in date_raw else date_raw
            shipper_full_name = str(sheet_details.cell(row=row_idx, column=4).value)

            # --- TIN ---
            tin_val = find_tin(tin_lookup, get_match_name(shipper_full_name))
            if not tin_val:
                print(f"Row {row_idx}: TIN not found for {shipper_full_name}")
                continue

            # --- CHECK ---
            try:
                status, page_html = check_with_retry(session, off_code, cus_no, cus_year, tin_val)
            except Exception as e:
                # Network/csrf awulak: colour eka wenas karanne naha, passe ayeth run karanna puluwan
                print(f"Row {row_idx}: ERROR ({e}) - colour karanne naha")
                continue

            desc = f"{off_code} E {cus_no}/{cus_year} | TIN {tin_val}"
            print(f"Row {row_idx} | {desc} -> {status}")

            if TEST_MODE:
                tested += 1
                out = os.path.join(debug_dir, f"row_{row_idx}_{status}.html")
                with open(out, "w", encoding="utf-8") as f:
                    f.write(page_html)
                print(f"   (response saved: {out}, length={len(page_html)})")
            else:
                results[row_idx] = GREEN if status == "GREEN" else RED
                log_progress(row_idx, status, desc)     # save fail unath result eka ihala file eke thiyenawa
                unsaved += 1

                if unsaved >= SAVE_EVERY:
                    if flush_results(results):
                        print(f"  (Excel ekata save kala: rows {len(results)})")
                        unsaved = 0

            time.sleep(DELAY)

    finally:
        if not TEST_MODE and results:
            ok = flush_results(results)
            if ok:
                print(f"Excel eka save kala (me run eke rows {len(results)}).")
            else:
                print("!!! Excel ekata save karanna bari una.")
                print(f"    Results tika me file eke thiyenawa: {PROGRESS_LOG}")
                print("    Excel eka close karala script eka ayeth run karanna (eka rows ayeth check wenawa).")
        if TEST_MODE:
            print("\nTEST iwarai. Site eke green/red ekata samana nam TEST_MODE = False karala run karanna.")
        print("Done!")


if __name__ == "__main__":
    run_automation()