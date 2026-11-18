"""
boat2.py eke Excel (COM) version eka -> openpyxl + LibreOffice version.
Wada karana widiya (boat2.py ekatama samana):
  CUSDEC DETAILS eke pata nathi row hama ekakatama
    1. CDN DETAILS walin match wena rows hoyanawa (count check)
    2. b1 sheet eka fill karanawa
    3. b2 (B...), b3/b4 (P...), b5 (SS...) sheets PDF karanawa
    4. CDN rows + CUSDEC row eka pata (blue) karanawa
"""
import datetime as dt
import os
import re
import shutil
import tempfile
from copy import copy

import openpyxl
from openpyxl.styles import Alignment, PatternFill
from openpyxl.worksheet.print_settings import PrintArea
from openpyxl.worksheet.properties import PageSetupProperties
from PyPDF2 import PdfReader, PdfWriter

from core import office
from core.paths import DATA_XLSX, PDF_DIR, WORK_DIR

EXCEL_FILE = DATA_XLSX
PDF_SAVE_PATH = PDF_DIR

BLUE = "FF00B0F0"          # boat2.py eke BLUE_COLOR = 15773696 (BGR) = RGB 00B0F0
PAPER_A4 = 9
# boat2.py wage SS (b5) sheet eka A4 karanawa. (b5 eke data.xlsx eke paper size 9 = A4 kalinma save wela thiyenawa)
FORCE_A4 = True

# (sheet, prefix, print area, A4 ?)   -- boat2.py eke pdf_list ekatama samana
def pdf_plan(p_sheet):
    return [
        ("b2", "B", "A1:AA80", False),
        (p_sheet, "P", "A1:K57", False),
        ("b5", "SS", "A1:AE57", FORCE_A4),
    ]


# ---------------------------------------------------------------- helpers
def _blue_fill():
    return PatternFill(start_color=BLUE, end_color=BLUE, fill_type="solid")


def has_any_fill(cell):
    """boat2: Interior.Pattern != xlNone  (kalin pata karapu nam skip)"""
    f = cell.fill
    return bool(f is not None and f.fill_type not in (None, "none"))


def display_text(cell):
    """Excel .Text (cell eke pennana widiya)."""
    v = cell.value
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        fmt = (cell.number_format or "").lower()
        d = v if isinstance(v, dt.datetime) else dt.datetime(v.year, v.month, v.day)
        if "yy" in fmt:
            out = fmt.replace("yyyy", "%Y").replace("yy", "%y")
            out = re.sub(r"\bmm?\b", lambda m: "%m", out) if "d" in out and "mmm" not in out else out
            out = re.sub(r"\bdd?\b", "%d", out)
            try:
                return d.strftime(out)
            except Exception:
                pass
        return d.strftime("%d/%m/%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def clean(v):
    return "" if v is None else str(v)


def check_shipper_type(wb, current_shipper):
    ws = wb["SUMMARY"]
    current_shipper = str(current_shipper).upper()
    row = 9
    while True:
        short_val = str(ws.cell(row, 26).value or "").strip().upper()       # Z
        not_short_val = str(ws.cell(row, 27).value or "").strip().upper()   # AA
        if not short_val and not not_short_val:
            break
        if short_val and short_val in current_shipper:
            return "b3"
        if not_short_val and not_short_val in current_shipper:
            return "b4"
        row += 1
    return "b3"


def find_cdn_rows(cdn, ref_cus):
    found = []
    for j in range(2, cdn.max_row + 1):
        cdn_ref = str(cdn.cell(j, 4).value).strip()
        if ref_cus in cdn_ref or cdn_ref in ref_cus:
            found.append(j)
    return found


def fill_b1(wb, cus, i, cdn, found_rows, expected_count):
    """boat2.py eke b1 fill karana kotasa (cell-for-cell samana)."""
    b1 = wb["b1"]
    r = found_rows[0]
    shipper_raw = str(cdn.cell(r, 5).value or "")
    shipper_raw = shipper_raw.replace("\r\n", "\n").replace("\r", "\n")
    formatted = "\n".join(line.strip() for line in shipper_raw.split("\n") if line.strip())

    b1["D3"].value = formatted
    al = copy(b1["D3"].alignment)
    b1["D3"].alignment = Alignment(horizontal=al.horizontal, vertical=al.vertical, wrap_text=True,
                                   indent=al.indent, text_rotation=al.text_rotation,
                                   shrink_to_fit=al.shrink_to_fit)
    b1["D4"].value = cdn.cell(r, 6).value
    b1["D5"].value = cdn.cell(r, 23).value
    b1["D6"].value = f"{clean(cdn.cell(r, 3).value)} {clean(cdn.cell(r, 4).value)}"
    b1["D7"].value = cdn.cell(r, 9).value
    b1["D8"].value = f"{clean(cdn.cell(r, 7).value)} {clean(cdn.cell(r, 8).value)}"
    b1["D9"].value = cdn.cell(r, 16).value
    b1["D10"].value = cdn.cell(r, 15).value
    b1["D11"].value = cdn.cell(r, 17).value
    b1["D12"].value = cdn.cell(r, 18).value

    c1 = cdn.cell(r, 3).value
    b1["D2"].value = c1 if c1 not in [None, 0, "0"] else ""
    b1["D13"].value = expected_count

    b1["H7"].number_format = "@"
    b1["H7"].value = display_text(cus.cell(i, 3)).strip()
    b1["H8"].value = cdn.cell(r, 11).value
    b1["H10"].value = cus.cell(i, 20).value
    b1["H11"].value = cus.cell(i, 15).value

    for rr in range(15, 51):                       # D15:F50 ClearContents
        for cc in (4, 5, 6):
            b1.cell(rr, cc).value = None
    start = 15
    for idx in found_rows:
        b1.cell(start, 4).value = cdn.cell(idx, 24).value
        b1.cell(start, 5).value = cdn.cell(idx, 20).value
        b1.cell(start, 6).value = cdn.cell(idx, 28).value
        start += 1


def apply_print_setup(wb, plan):
    for sheet_name, _prefix, area, is_a4 in plan:
        ws = wb[sheet_name]
        ws.print_area = area
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.print_options.horizontalCentered = True
        ws.print_options.verticalCentered = True
        if is_a4:
            ws.page_setup.paperSize = PAPER_A4
            ws.page_setup.orientation = "portrait"


def render_pdfs(master_path, row_setup, plan, out_paths, work):
    """
    master_path eka fresh load karala (openpyxl image bug nisa hama parakma aluthen load),
    row_setup(wb) eken b1 fill karala, plan eke sheets witharak visible karala LibreOffice eken PDF hadanawa.
    PDF ekak = sheet ekak = page ekak (fit 1x1). Page gana hari nathnam sheet ekin eka convert karanawa.
    """
    names = [p[0] for p in plan]

    def build(visible, tag):
        wb = openpyxl.load_workbook(master_path)
        row_setup(wb)
        apply_print_setup(wb, plan)
        for ws in wb.worksheets:
            if ws.title in visible:
                ws.sheet_state = "visible"
            else:
                # LibreOffice: print area thiyena hidden sheet ekath print karanawa -> print area ain karanawa
                ws._print_area = PrintArea()
                ws.sheet_state = "hidden"
        wb.active = wb.sheetnames.index(visible[0])
        for ws in wb.worksheets:
            ws.sheet_view.tabSelected = (ws.title == visible[0])
        tmp = os.path.join(work, f"render_{tag}.xlsx")
        wb.save(tmp)
        return office.convert(tmp, "pdf", os.path.join(work, f"out_{tag}"))

    combined = build(names, "all")
    reader = PdfReader(combined)
    if len(reader.pages) == len(names):
        for page, dest in zip(reader.pages, out_paths):
            w = PdfWriter()
            w.add_page(page)
            with open(dest, "wb") as f:
                w.write(f)
        return "combined"

    # fallback: sheet ekin eka
    for n, dest in zip(names, out_paths):
        one = build([n], re.sub(r"\W", "_", n))
        if len(PdfReader(one).pages) != 1:
            raise RuntimeError(f"{n}: PDF eka page 1k nemei ({len(PdfReader(one).pages)}). Sheet eka fit wenne naha.")
        shutil.copy2(one, dest)
    return "single"


# ---------------------------------------------------------------- main
def process_and_pdf(on_mismatch="skip", overwrite=True, log=print):
    """
    on_mismatch : 'skip'  -> CDN row count ekata samana naththam row eka skip karanawa (boat2 eke "No")
                  'go'    -> ignore karala PDF hadanawa (boat2 eke "Yes")
    overwrite   : PDF eka kalin thiyenawanam uda liyanawada (boat2 eke ask_overwrite)
    """
    if not os.path.exists(EXCEL_FILE):
        log("Error: data.xlsx hambune naha!")
        return
    os.makedirs(PDF_SAVE_PATH, exist_ok=True)

    work = tempfile.mkdtemp(prefix="boat_", dir=WORK_DIR)
    state = os.path.join(work, "state.xlsx")
    shutil.copy2(EXCEL_FILE, state)          # original eka iwara unama witharai wenas wenne

    colored_cus, colored_cdn = set(), set()
    last_setup = None
    processed = 0
    try:
        wb = openpyxl.load_workbook(state)
        cus, cdn = wb["CUSDEC DETAILS"], wb["CDN DETAILS"]
        last_row = cus.max_row
        log(f"Mulu CUSDEC rows gana: {last_row}")
        consecutive_errors = 0

        for i in range(2, last_row + 1):
            cell = cus.cell(i, 2)
            if has_any_fill(cell):
                continue                       # kalin pata karala (boat2: skip)
            ref_cus = str(cell.value).strip()
            if not ref_cus or ref_cus == "None":
                continue
            shipper_name = str(cus.cell(i, 4).value or "").strip()
            try:
                expected = int(float(cus.cell(i, 9).value or 0))
            except Exception:
                expected = 0

            log(f"Process karanawa: {ref_cus} (row {i})")
            ref_name = ref_cus.replace("/", "_")
            found = find_cdn_rows(cdn, ref_cus)

            if len(found) != expected:
                msg = f"{ref_cus}: Expected CDN rows {expected}, found {len(found)}"
                if on_mismatch != "go":
                    log(f"  ! {msg} -> SKIP (Row count mismatch)")
                    continue
                log(f"  ! {msg} -> ignore karala PDF hadanawa")

            p_sheet = check_shipper_type(wb, shipper_name)
            log(f"Thorapu sheet eka: {p_sheet}")
            plan = pdf_plan(p_sheet)
            for sname, *_ in plan:
                if sname not in wb.sheetnames:
                    raise RuntimeError(f"Sheet '{sname}' hambune naha")

            targets = [os.path.join(PDF_SAVE_PATH, f"{prefix}{ref_name}.pdf") for _, prefix, _, _ in plan]
            todo = [(p, t) for p, t in zip(plan, targets) if overwrite or not os.path.exists(t)]
            if not todo:
                log("  PDF 3 ma kalin thiyenawa (overwrite off) - skip")
                continue

            def row_setup(w, _i=i, _found=found, _exp=expected):
                fill_b1(w, w["CUSDEC DETAILS"], _i, w["CDN DETAILS"], _found, _exp)

            try:
                mode = render_pdfs(state, row_setup, [p for p, _ in todo], [t for _, t in todo], work)
                consecutive_errors = 0
            except Exception as e:           # me row eka witharak fail; pata karanne naha, ayeth run karanna puluwan
                consecutive_errors += 1
                log(f"  ERROR {ref_cus}: {e} -> row eka skip (pata karanne naha)")
                if consecutive_errors >= 3:
                    log("  Row 3k pasu pasa fail unaa -> nawaththanawa.")
                    break
                continue
            for (_p, t) in todo:
                log(f"  PDF hadunaa: {os.path.basename(t)}")
            last_setup = (i, found, expected, plan)
            colored_cdn.update(found)
            colored_cus.add(i)
            processed += 1
            log(f"Iwarai: {ref_cus}")
    finally:
        # Cancel / error unath, iwara una rows walata pata + b1 save karanawa (PDF hadala thiyenawa nisa)
        try:
            if colored_cus:
                _final_save(state, work, last_setup, colored_cus, colored_cdn)
                log("data.xlsx update kala (pata karapu rows save kala).")
        finally:
            shutil.rmtree(work, ignore_errors=True)
    log(f"\nMulu process karapu gana: {processed}")
    log("Wade iwarai!")


def _final_save(state, work, last_setup, colored_cus, colored_cdn):
    """Fresh load -> colours + b1 (last row) + print setup -> eka parak save -> data.xlsx replace."""
    wb2 = openpyxl.load_workbook(state)
    cus2, cdn2 = wb2["CUSDEC DETAILS"], wb2["CDN DETAILS"]
    if last_setup:
        i, found, expected, plan = last_setup
        fill_b1(wb2, cus2, i, cdn2, found, expected)
        apply_print_setup(wb2, plan)
    fill = _blue_fill()
    for ws, rows in ((cdn2, colored_cdn), (cus2, colored_cus)):
        width = max(ws.max_column, 30)
        for idx in rows:
            for c in range(1, width + 1):
                ws.cell(idx, c).fill = copy(fill)
    tmp_final = os.path.join(work, "final.xlsx")
    wb2.save(tmp_final)
    os.replace(tmp_final, EXCEL_FILE)
