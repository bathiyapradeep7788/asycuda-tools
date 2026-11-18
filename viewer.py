"""Excel sheet ekak HTML table ekak widiyata (cell colours ekka) -- laptop ekath phone ekath wada."""
import html
import datetime as dt

import openpyxl
from openpyxl.utils import get_column_letter

from core.export import fill_hex          # theme/tint wala colours hariyata hoyanna (original script eke function eka)

DEFAULT_SHEETS = ["CUSDEC DETAILS", "CDN DETAILS", "SUMMARY", "extra partys", "b1"]

CSS = """
<style>
.xl-wrap{overflow:auto;max-height:72vh;border:1px solid #8884;border-radius:6px}
table.xl{border-collapse:collapse;font-size:12px;font-family:Calibri,Arial,sans-serif;white-space:nowrap}
table.xl th{position:sticky;top:0;background:#e9ecef;color:#222;border:1px solid #bbb;padding:2px 6px;z-index:2}
table.xl td{border:1px solid #d0d0d0;padding:2px 6px;color:#111;background:#fff;max-width:340px;overflow:hidden;text-overflow:ellipsis}
table.xl td.rn{background:#e9ecef;color:#555;position:sticky;left:0;text-align:right;z-index:1}
@media (max-width:640px){table.xl{font-size:11px}}
</style>
"""


def _fmt(v):
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def load(path):
    wb = openpyxl.load_workbook(path)          # fills (colours) ona nisa read_only naha
    return wb


def sheet_names(wb, all_sheets=False):
    if all_sheets:
        return wb.sheetnames
    return [n for n in DEFAULT_SHEETS if n in wb.sheetnames]


def search_rows(ws, text):
    """filter text eka thiyena row numbers (header row 1 ekath ekka)."""
    text = (text or "").strip().lower()
    rows = list(range(1, ws.max_row + 1))
    if not text:
        return rows
    out = [1]
    for r in range(2, ws.max_row + 1):
        for c in ws[r]:
            if c.value is not None and text in str(c.value).lower():
                out.append(r)
                break
    return out


def render(ws, wb, rows, max_cols=None):
    theme = wb.loaded_theme
    ncols = min(ws.max_column, max_cols or ws.max_column)
    parts = [CSS, '<div class="xl-wrap"><table class="xl"><thead><tr><th></th>']
    parts += [f"<th>{get_column_letter(c)}</th>" for c in range(1, ncols + 1)]
    parts.append("</tr></thead><tbody>")
    for r in rows:
        parts.append(f'<tr><td class="rn">{r}</td>')
        for c in range(1, ncols + 1):
            cell = ws.cell(r, c)
            css = []
            try:
                if cell.fill is not None and cell.fill.fill_type:
                    hx = fill_hex(cell, theme)
                    if hx and len(hx) >= 6:
                        css.append(f"background:#{hx[-6:]}")
            except Exception:
                pass
            if cell.font is not None and cell.font.b:
                css.append("font-weight:bold")
            style = f' style="{";".join(css)}"' if css else ""
            val = html.escape(_fmt(cell.value)).replace("\n", "<br>")
            parts.append(f"<td{style}>{val}</td>")
        parts.append("</tr>")
    parts.append("</tbody></table></div>")
    return "".join(parts)
