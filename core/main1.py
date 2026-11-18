from core.paths import WORK_DIR, DATA_XLSX, PDF_DIR, MERGED_DIR, LOGS_DIR, EXPORTERS_XLSX
import imaplib
import email
from email.header import decode_header  # Header decode කිරීම සඳහා එකතු කළා
import pdfplumber
import io
import os
import re
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

# ==========================================
# --- CONFIGURATIONS & PATHS ---
# ==========================================
GMAIL_USER = os.environ.get('GMAIL_USER', '')
GMAIL_APP_PASS = os.environ.get('GMAIL_APP_PASS', '')

EXCEL_FILE = DATA_XLSX
PDF_SAVE_PATH = PDF_DIR

GROSS_COL = 28  # CDN DETAILS -> Column AB (Gross weight)

if not os.path.exists(PDF_SAVE_PATH):
    os.makedirs(PDF_SAVE_PATH)

# ==========================================
# --- HELPER FUNCTIONS ---
# ==========================================
def clean_filename(text):
    """ගොනු නාමයේ ඇති අහිතකර ලකුණු ඉවත් කරයි"""
    return "".join(c for c in text if c.isalnum() or c in (' ', '.', '_', '-')).strip()


def parse_weight(value):
    """
    Weight එක number එකකට හරවනවා.
      22.222.00  -> 22222.00
      22,222.00  -> 22222.00
      22.222,00  -> 22222.00
      1.234.567.50 -> 1234567.50
    හරවන්න බැරි නම් None.
    """
    if value is None:
        return None
    s = str(value).upper().replace("KGS", "").replace("KG", "").replace(" ", "").strip()
    s = re.sub(r"[^\d.,]", "", s)
    if not s or not re.search(r"\d", s):
        return None

    last_dot, last_comma = s.rfind("."), s.rfind(",")

    if last_dot != -1 and last_comma != -1:
        # දෙකම තියෙනවා නම් අන්තිමට එන එක decimal point එක
        dec = "." if last_dot > last_comma else ","
        thou = "," if dec == "." else "."
        s = s.replace(thou, "").replace(dec, ".")
    elif s.count(".") > 1:
        # 22.222.00 -> අන්තිම තිත decimal, අනිත් ඒවා thousands
        head, tail = s.rsplit(".", 1)
        s = head.replace(".", "") + "." + tail
    elif s.count(",") > 1:
        s = s.replace(",", "")
    elif "," in s:
        head, tail = s.split(",")
        # 22,222 -> thousands | 22,50 -> decimal
        s = head + tail if len(tail) == 3 else head + "." + tail

    try:
        return float(s)
    except ValueError:
        return None


def get_unread_emails():
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(GMAIL_USER, GMAIL_APP_PASS)
        mail.select("inbox")
        status, messages = mail.search(None, '(UNSEEN FROM "bathiyapradeep@yahoo.com")')
        return mail, messages[0].split()
    except Exception as e:
        print(f"Login Error: {e}")
        return None, []

# ==========================================
# --- MAIN PROCESS ---
# ==========================================
def main_process():
    # --- STAGE 1: Email Extraction & IMMEDIATE Text Storage ---
    mail, email_ids = get_unread_emails()
    if not email_ids:
        print("අලුත් මේල් මුකුත් නැහැ මචං!")
        return
    else:
        try:
            book = load_workbook(EXCEL_FILE)
            for e_id in email_ids:
                _, msg_data = mail.fetch(e_id, "(RFC822)")
                msg = email.message_from_bytes(msg_data[0][1])
                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue

                    filename = part.get_filename()

                    # --- UTF-8 Filename Fix ---
                    if filename:
                        decode_list = decode_header(filename)
                        decoded_filename, encoding = decode_list[0]
                        if isinstance(decoded_filename, bytes):
                            decoded_filename = decoded_filename.decode(encoding if encoding else 'utf-8')
                        filename = clean_filename(decoded_filename)

                    if part.get_content_maintype() == 'application' and filename and 'pdf' in filename.lower():
                        pdf_data = part.get_payload(decode=True)
                        full_save_path = os.path.join(PDF_SAVE_PATH, filename)

                        with open(full_save_path, 'wb') as f:
                            f.write(pdf_data)
                        print(f"Saved PDF: {filename}")

                        with pdfplumber.open(io.BytesIO(pdf_data)) as pdf:
                            full_text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])

                        if "CARGO DISPATCH NOTE" in full_text.upper():
                            if 'ex cdn' in book.sheetnames:
                                sheet = book['ex cdn']
                                next_row = 1
                                while sheet.cell(row=next_row, column=1).value is not None:
                                    next_row += 1
                                sheet.cell(row=next_row, column=1).value = full_text + f"\nSOURCE_FILE:{filename}"

                        elif "GOODS DECLARATION" in full_text.upper() or "CUSDEC" in full_text.upper():
                            if 'ex cus' in book.sheetnames:
                                sheet = book['ex cus']
                                next_row = 1
                                while sheet.cell(row=next_row, column=1).value is not None:
                                    next_row += 1
                                sheet.cell(row=next_row, column=1).value = full_text

            book.save(EXCEL_FILE)
            print("පීඩීඑෆ් Extract කර Excel එකට දැමීම අවසන්!")
        except Exception as e:
            print(f"Email Process Error: {e}")

    # --- STAGE 2: Processing 'ex cdn' Data to 'CDN DETAILS' ---
    if not os.path.exists(EXCEL_FILE): return
    print("\nදැන් CDN Details ටික Extract කරනවා..."); current_year = datetime.now().year
    try:
        book = load_workbook(EXCEL_FILE)
        sheet_ex = book['ex cdn']
        sheet_det = book['CDN DETAILS']
        current_row_det = 2
        while sheet_det.cell(row=current_row_det, column=1).value is not None:
            current_row_det += 1
        for i in range(1, sheet_ex.max_row + 1):
            full_raw_text = sheet_ex.cell(row=i, column=1).value
            if full_raw_text:
                source_filename = ""
                if "SOURCE_FILE:" in full_raw_text:
                    parts = full_raw_text.split("SOURCE_FILE:")
                    full_text = parts[0]
                    source_filename = parts[1].strip()
                else:
                    full_text = full_raw_text
                lines = [l.strip() for l in full_text.split('\n') if l.strip()]
                data = {k: "" for k in ["cbex","yr","ser","num","ship","bl","l1","l2","cus_num","cdn_f","con_z","con_f","driver","loc","vessel_g","date_h","sec_p","col_n","port_o","vop_q","cop_r","con_x","size_t","type_aa","type_y","bl_u","v_60","weight_ab","desc_w"]}
                driver_idx = -1
                for k, line in enumerate(lines):
                    if "13. Location of Goods" in line and k > 0:
                        data["driver"] = lines[k-1]
                        driver_idx = k-1
                        break
                cdn_match = re.search(r"11\.a Seal No\. CDN No\.(.*?)2\.b Consignee", full_text, re.DOTALL)
                if cdn_match:
                    inner_text = cdn_match.group(1).strip()
                    found_cdn = re.findall(r"C\s*\d+", inner_text)
                    if found_cdn: data["cdn_f"] = found_cdn[0]
                for k, line in enumerate(lines):
                    if "Weight," in line and k+1 < len(lines) and "Kg" in lines[k+1]:
                        if k+2 < len(lines):
                            parts_x = lines[k+2].split()
                            if len(parts_x) >= 2: data["con_x"], data["size_t"] = parts_x[0], parts_x[1]
                    if "23. Type of container" in line and k > 0:
                        data["type_aa"] = lines[k-1]
                        if k > 1:
                            parts_yu = lines[k-2].split()
                            if len(parts_yu) >= 2: data["type_y"], data["bl_u"] = parts_yu[0], parts_yu[1]
                        if k > 2:
                            parts_v = lines[k-3].split()
                            if parts_v: data["v_60"] = parts_v[-1]
                        if k > 3:
                            w_line = lines[k-3]
                            w_clean = w_line.replace("KG", "").replace("kg", "").strip().split()
                            if len(w_clean) >= 2: data["weight_ab"] = w_clean[-2]
                            data["desc_w"] = w_line.split(data["weight_ab"])[0].strip() if data["weight_ab"] in w_line else w_line
                    h17 = "17. Signature, Designation and Date"
                    if h17 in line:
                        parts_after = line.split(h17)[0].strip().split()
                        if len(parts_after) >= 2: data["cop_r"], data["vop_q"] = parts_after[-1], parts_after[-2]
                        h_code = "Code Cont. Op.Code"
                        pat_o = f"{re.escape(h_code)}(.*?){re.escape(data['vop_q'])}"
                        m_o = re.search(pat_o, full_text, re.DOTALL)
                        if m_o: data["port_o"] = " ".join([l.strip() for l in m_o.group(1).split('\n') if l.strip()])
                    if "6. Name of certifying" in line and k+1 < len(lines):
                        data["sec_p"], data["col_n"] = lines[k+1].replace("COLOMBO", "").strip(), "COLOMBO"
                    if "conditions &" in line and k+1 < len(lines):
                        target_l = lines[k+1]
                        day_m = re.search(r"(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+([A-Za-z]+)\s+(\d+)", target_l)
                        if day_m:
                            data["vessel_g"] = target_l.split(day_m.group(1))[0].strip()
                            m_s, d_v = day_m.group(2), day_m.group(3).zfill(2)
                            months = {"Jan":"01","Feb":"02","Mar":"03","Apr":"04","May":"05","Jun":"06","Jul":"07","Aug":"08","Sep":"09","Oct":"10","Nov":"11","Dec":"12"}
                            data["date_h"] = f"{d_v}.{months.get(m_s, '01')}.{current_year}"
                    if "DECLARATION REFERENCE" in line and k+1 < len(lines):
                        ref = lines[k+1].split()
                        if len(ref) >= 4: data["cbex"], data["yr"], data["ser"], data["num"] = ref[0], ref[1], ref[2], ref[3]
                    if "1.a Shipper" in line and k+1 < len(lines):
                        p1 = lines[k+1].split()
                        if len(p1) >= 3:
                            data["bl"], data["l1"] = p1[-1], p1[-2]
                            ship_line1 = " ".join(p1[:-2])
                            if k+2 < len(lines) and all(x not in lines[k+2] for x in ["1.b", "2.b"]):
                                p2 = lines[k+2].split()
                                if len(p2) >= 2:
                                    data["l2"] = p2[-1]
                                    ship_line2 = " ".join(p2[:-1])
                                    data["ship"] = f"{ship_line1}\n{ship_line2}"
                                else: data["ship"] = ship_line1
                            else: data["ship"] = ship_line1
                    h2b = "2.b Consignee (Name and Address)"
                    if h2b in line:
                        pz = line.split(h2b)[1].split()
                        if pz: data["con_z"] = pz[0]
                        temp_f = []
                        f_start = line.split(data["con_z"])[-1].strip()
                        if f_start: temp_f.append(f_start)
                        for m in range(k+1, len(lines)):
                            if m >= driver_idx: break
                            if "12. Name of Driver" not in lines[m]: temp_f.append(lines[m])
                        data["con_f"] = "\n".join(temp_f).strip()
                    if "1.b Cusdec Numbers" in line and k+1 < len(lines): data["cus_num"] = lines[k+1]
                    if "13. Location of Goods" in line and k+1 < len(lines): data["loc"] = lines[k+1]

                cols = [data["cbex"], data["yr"], data["ser"], data["num"], data["ship"], data["con_f"], data["vessel_g"], data["date_h"], data["bl"], data["driver"], data["loc"], data["l1"], data["l2"], data["col_n"], data["port_o"], data["sec_p"], data["vop_q"], data["cop_r"], data["cus_num"], data["size_t"], data["bl_u"], data["v_60"], data["desc_w"], data["con_x"], data["type_y"], data["con_z"], data["type_aa"], data["weight_ab"], "", data["cdn_f"]]
                for idx, val in enumerate(cols, 1):
                    sheet_det.cell(row=current_row_det, column=idx).value = val

                # --- GROSS weight format fix: 22.222.00 -> 22,222.00 ---
                gross_cell = sheet_det.cell(row=current_row_det, column=GROSS_COL)
                gross_num = parse_weight(data["weight_ab"])
                if gross_num is not None:
                    gross_cell.value = gross_num
                    gross_cell.number_format = '#,##0.00'
                else:
                    print(f"Gross weight හරවන්න බැරි වුණා: '{data['weight_ab']}' (පේළිය {current_row_det})")

                loc_val = str(data["loc"]).upper()
                if any(x in loc_val for x in ["SLPA", "JCT", "ECT"]):
                    if source_filename:
                        file_to_del = os.path.join(PDF_SAVE_PATH, source_filename)
                        if os.path.exists(file_to_del):
                            os.remove(file_to_del)
                            print(f"Deleted CDN PDF: {source_filename}")
                current_row_det += 1
        book.save(EXCEL_FILE)
        print("CDN Details Update වුණා.")
    except Exception as e: print(f"CDN Error: {e}")

    # --- STAGE 3: Processing 'ex cus' Data to 'CUSDEC DETAILS' ---
    print("\nදැන් CusDec Details ටික Extract කරනවා...")
    try:
        book = load_workbook(EXCEL_FILE)
        sheet_ex_cus = book['ex cus']
        sheet_det_cus = book['CUSDEC DETAILS']
        curr_row_cus = 2
        while sheet_det_cus.cell(row=curr_row_cus, column=1).value is not None:
            curr_row_cus += 1
        for i in range(1, sheet_ex_cus.max_row + 1):
            txt = sheet_ex_cus.cell(row=i, column=1).value
            if txt:
                c_data = {k: "" for k in ["a","b","c","d","e","f","i","o","t","u"]}
                off_m = re.search(r"EX\s*1\s+([A-Z0-9]+)", txt)
                if off_m: c_data["a"] = off_m.group(1)
                bc_m = re.search(r"Load\s*List\s*[A-Z]\s*(\d+)\s*(\d{2}/\d{2}/\d{4})", txt)
                if bc_m: c_data["b"], c_data["c"] = bc_m.group(1), bc_m.group(2)
                d_m = re.search(r"Information(.*?)3\s*Forms", txt, re.DOTALL)
                if d_m: c_data["d"] = " ".join(d_m.group(1).replace("Customs Reference Number:", "").split()).strip()
                e_m = re.search(r"Financial\s*Settlement\s*TIN:(.*?)10\s*Cty\s*of\s*Last", txt, re.DOTALL)
                if e_m:
                    temp_e = re.sub(r'Order\s+granted', '', e_m.group(1), flags=re.IGNORECASE)
                    c_data["e"] = " ".join(temp_e.split()).strip()
                f_m = re.search(r"7\s*Declarant's\s*Sequence\s*Number\s*1\s+([\d,.]+)", txt)
                if f_m: c_data["f"] = f_m.group(1)
                i_m = re.search(r"Cty\.\s+(\d+)", txt)
                if i_m: c_data["i"] = i_m.group(1)
                o_m = re.search(r"([\d,.]+)\s+cod", txt)
                if o_m: c_data["o"] = o_m.group(1)
                t_m = re.search(r"36\s*Preference.*?a\s*LK\s*b\s*([\d,.]+)", txt, re.DOTALL)
                if t_m: c_data["t"] = t_m.group(1)
                u_m = re.search(r"39\s*Quota\s*Containers\s*No\(s\)\s+[\d.]+\s+[\d.]+\s+([\d,.]+)", txt)
                if u_m: c_data["u"] = u_m.group(1)
                cus_cols = {1: "a", 2: "b", 3: "c", 4: "d", 5: "e", 6: "f", 9: "i", 15: "o", 20: "t", 21: "u"}
                for col_idx, key in cus_cols.items():
                    sheet_det_cus.cell(row=curr_row_cus, column=col_idx).value = c_data[key]
                curr_row_cus += 1
        s_ex_cdn = book['ex cdn']
        s_ex_cus = book['ex cus']
        if s_ex_cdn.max_row >= 1: s_ex_cdn.delete_rows(1, s_ex_cdn.max_row)
        if s_ex_cus.max_row >= 1: s_ex_cus.delete_rows(1, s_ex_cus.max_row)
        book.save(EXCEL_FILE)
        print("දත්ත අලුතෙන් ඇඩ් වුණා සහ Raw Tabs පිරිසිදු කළා.")
        print("දැන් කියවපු මේල් ටික Trash එකට දානවා...")
        for e_id in email_ids:
            mail.store(e_id, '+X-GM-LABELS', '\\Trash')
        mail.expunge()
        mail.logout()
        print("වැඩේ සම්පූර්ණයෙන්ම ඉවරයි මචං!")
    except Exception as e: print(f"CusDec Error: {e}")

if __name__ == "__main__":
    main_process()