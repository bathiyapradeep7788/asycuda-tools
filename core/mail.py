from core.paths import WORK_DIR, DATA_XLSX, PDF_DIR, MERGED_DIR, LOGS_DIR, EXPORTERS_XLSX
import smtplib
import os
from email.message import EmailMessage
from openpyxl import load_workbook

# --- Configurations ---
GMAIL_USER = os.environ.get('GMAIL_USER', '')
GMAIL_APP_PASS = os.environ.get('GMAIL_APP_PASS', '')
EXCEL_FILE = DATA_XLSX
PDF_FOLDER_PATH = MERGED_DIR
TARGET_EMAIL = os.environ.get('TARGET_EMAIL', 'bathiyapradeep@yahoo.com')

# Attach වෙන පිළිවෙල (මේවා නැත්නම් අනිත් PDF පස්සෙන් යනවා)
ATTACH_ORDER = ["BOAT NOTES.pdf", "CUSDECS.pdf", "CDNS.pdf", "AMEND DOCS.pdf"]


def get_ordered_pdfs():
    files = [f for f in os.listdir(PDF_FOLDER_PATH) if f.lower().endswith('.pdf')]
    order = {name.upper(): idx for idx, name in enumerate(ATTACH_ORDER)}
    return sorted(files, key=lambda f: (order.get(f.upper(), len(order)), f.upper()))


def send_email_with_attachments():
    try:
        if not os.path.exists(EXCEL_FILE):
            print("Error: Excel file එක නැහැ!")
            return

        # 1. Excel එකෙන් දත්ත ගැනීම (AA3 - Body, AA4 - Subject)
        wb = load_workbook(EXCEL_FILE, data_only=True)
        ws = None
        for sn in wb.sheetnames:
            if sn.strip().upper() == "SUMMARY":
                ws = wb[sn]
                break

        if not ws:
            print("Error: 'SUMMARY' sheet එක නැහැ!")
            return

        body = ws['AA3'].value
        subject = ws['AA4'].value

        if not body or not subject:
            print("Error: AA3 හෝ AA4 හි දත්ත හිස්!")
            return

        # 2. Email එක සකස් කිරීම
        msg = EmailMessage()
        msg.set_content(str(body))
        msg['Subject'] = str(subject)
        msg['From'] = GMAIL_USER
        msg['To'] = TARGET_EMAIL

        # 3. Merged folder එකේ PDF ඔක්කොම attach කිරීම (AMEND DOCS.pdf එකත් එක්ක)
        pdf_files = []
        if os.path.exists(PDF_FOLDER_PATH):
            pdf_files = get_ordered_pdfs()

            if not pdf_files:
                print("Warning: Folder එකේ PDF ගොනු මුකුත් නැහැ!")

            if "AMEND DOCS.PDF" not in [f.upper() for f in pdf_files]:
                print("Warning: AMEND DOCS.pdf එක folder එකේ නැහැ!")

            for file_name in pdf_files:
                file_path = os.path.join(PDF_FOLDER_PATH, file_name)
                with open(file_path, 'rb') as f:
                    msg.add_attachment(
                        f.read(),
                        maintype='application',
                        subtype='pdf',
                        filename=file_name
                    )
                print(f"Attached: {file_name}")
        else:
            print("Error: PDF Folder එක හොයාගන්න බැහැ!")

        # 4. Gmail හරහා යැවීම
        print("Connecting to Gmail...")
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(GMAIL_USER, GMAIL_APP_PASS)
            smtp.send_message(msg)

        print(f"Success! මේල් එක {TARGET_EMAIL} ට යැව්වා.")

        # 5. යැව්වට පස්සේ merged folder එකේ PDF ඔක්කොම මකනවා (AMEND DOCS.pdf එකත්)
        if pdf_files:
            print("දැන් Merged Folder එකේ තිබුණ PDF ටික මකනවා...")
            for file_name in pdf_files:
                try:
                    os.remove(os.path.join(PDF_FOLDER_PATH, file_name))
                    print(f"Deleted: {file_name}")
                except Exception as del_e:
                    print(f"File එක මකන්න බැරි වුණා {file_name}: {del_e}")
            print("Folder එක Clear කිරීම අවසන්!")

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    send_email_with_attachments()