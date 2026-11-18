import os
from core.paths import WORK_DIR, DATA_XLSX, PDF_DIR, MERGED_DIR, LOGS_DIR, EXPORTERS_XLSX
import time
import random
import re
import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate, make_msgid

# ==========================================
# 1. CONFIGURATION / සැකසුම් (BREVO SMTP)
# ==========================================
EXCEL_FILE_PATH = EXPORTERS_XLSX
SHEET_NAME = "Exporters"

# Brevo SMTP Settings
SMTP_SERVER = "smtp-relay.brevo.com"
SMTP_PORT = 587  # TLS Port for Brevo

# Brevo Login Username
BREVO_LOGIN_USER = os.environ.get("BREVO_LOGIN_USER", "")

# Exporter ට පෙනෙන From Email Address එක
SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "export@bastellogistics.com")

# Brevo SMTP Key (Password)
SENDER_PASSWORD = os.environ.get("BREVO_SMTP_KEY", "")

# TEST MODE:
# True  -> Row 2 (පළමු exporter) ට පමණක් mail එක යවයි.
# False -> START_ROW සිට පහළට සියලුම Exporters ලාට auto mail යවයි.
TEST_MODE = False 

# START ROW:
START_ROW = 2 

# Logo Direct URL
LOGO_URL = "https://www.bastellogistics.com/assets/bastel.png"

# ==========================================
# 2. EMAIL CONTENT CREATOR (PLAIN + HTML)
# ==========================================
def get_email_content(exporter_name):
    subject = f"Cost-Effective Customs Clearing & Logistics Support for {exporter_name}"
    
    # 1. Plain Text Version
    plain_text = f"""
Dear {exporter_name} Team,

I hope this email finds you well.

I am reaching out from Bastel Pvt Ltd, a dedicated and licensed customs brokerage company based in Sri Lanka.

Our core expertise lies in seamless and fast customs clearing. However, we also understand the need for budget-friendly logistics. By leveraging our strong industry network, we are able to arrange and handle your air and sea freight requirements at highly discounted, fair, and transparent rates.

Our goal is simple: to get your work done at the most reasonable price without compromising on reliability.

Core Services:
- Licensed Customs Brokerage (Our Core Expertise)
- Cost-Effective Freight Handling (Arranging Sea/Air Freight at low rates)
- Full Trade Documentation (BL, COO, Phytosanitary Certs)
- Fair-Priced Transport Management & Port Pickups

Connect with us on WhatsApp for a quick quote: https://wa.me/94710940466

Best regards,
Export Management Team
BASTEL PRIVATE LIMITED
Email: export@bastellogistics.com
Phone/WhatsApp: +94 71 094 0466
Web: https://bastellogistics.com
Location: Colombo, Sri Lanka
    """

    # 2. Rich HTML Version
    html_body = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            body {{ font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif; background-color: #f1f5f9; margin: 0; padding: 25px 10px; color: #334155; }}
            .email-wrapper {{ max-width: 620px; margin: 0 auto; background: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 10px 25px rgba(0,0,0,0.06); border: 1px solid #e2e8f0; }}
            .header {{ background: #0b1f1c; padding: 25px 35px; text-align: left; border-bottom: 4px solid #14b8a6; }}
            .content {{ padding: 35px; line-height: 1.65; font-size: 14.5px; }}
            .greeting {{ font-size: 17px; font-weight: 700; color: #0f172a; margin-bottom: 16px; }}
            p {{ margin-bottom: 16px; color: #475569; }}
            .feature-list {{ padding: 0; margin: 20px 0; list-style: none; }}
            .feature-item {{ padding: 10px 0; border-bottom: 1px dashed #e2e8f0; display: flex; align-items: flex-start; }}
            .badge {{ background: #0f766e; color: #ffffff; font-size: 10px; font-weight: 700; padding: 3px 8px; border-radius: 4px; text-transform: uppercase; margin-right: 12px; min-width: 75px; text-align: center; display: inline-block; }}
            .cta-container {{ text-align: center; margin: 30px 0 15px 0; }}
            .cta-button {{ background: #25D366; color: #ffffff !important; text-decoration: none; padding: 14px 28px; border-radius: 8px; font-weight: 700; font-size: 14px; display: inline-block; box-shadow: 0 4px 12px rgba(37, 211, 102, 0.25); }}
            .signature-divider {{ border-top: 2px solid #f1f5f9; margin-top: 35px; padding-top: 25px; }}
            .sig-table {{ width: 100%; border-collapse: collapse; }}
            .sig-logo-cell {{ width: 110px; vertical-align: top; padding-right: 18px; text-align: center; }}
            .sig-info-cell {{ border-left: 3px solid #14b8a6; padding-left: 18px; vertical-align: top; }}
            .sig-name {{ font-size: 16px; font-weight: 700; color: #0f172a; margin: 0; }}
            .sig-title {{ font-size: 11.5px; font-weight: 700; color: #0d9488; letter-spacing: 0.8px; text-transform: uppercase; margin: 2px 0 10px 0; }}
            .sig-link {{ color: #0d9488; text-decoration: none; font-weight: 600; }}
            .footer {{ background: #f8fafc; padding: 18px 35px; text-align: center; font-size: 11.5px; color: #94a3b8; border-top: 1px solid #e2e8f0; }}
        </style>
    </head>
    <body>
    <div class="email-wrapper">
        <div class="header">
            <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                    <td style="vertical-align: middle;">
                        <a href="https://bastellogistics.com" target="_blank">
                            <img src="{LOGO_URL}" alt="Bastel Pvt Ltd" style="height: 45px; width: auto; max-height: 45px; display: block; border: 0;" />
                        </a>
                    </td>
                </tr>
            </table>
        </div>
        <div class="content">
            <div class="greeting">Dear {exporter_name} Team,</div>
            <p>I hope this email finds you well.</p>
            <p>I am reaching out from <strong>Bastel Pvt Ltd</strong>, a dedicated and licensed customs brokerage company based in Sri Lanka.</p>

            <table width="100%" cellpadding="0" cellspacing="0" style="margin: 20px 0;">
                <tr>
                    <td width="48%" style="background: #f0fdf4; border: 1px solid #bbf7d0; padding: 14px; border-radius: 8px; vertical-align: top;">
                        <div style="font-weight: 700; color: #166534; font-size: 13.5px; margin-bottom: 4px;">🤝 Customs Brokerage (Core)</div>
                        <div style="font-size: 12.5px; color: #15803d; line-height: 1.4;">Our primary expertise is fast, hassle-free customs clearance. We clear your cargo swiftly to avoid demurrage.</div>
                    </td>
                    <td width="4%"></td>
                    <td width="48%" style="background: #f0fdf4; border: 1px solid #bbf7d0; padding: 14px; border-radius: 8px; vertical-align: top;">
                        <div style="font-weight: 700; color: #166534; font-size: 13.5px; margin-bottom: 4px;">💰 Fair & Low Freight Rates</div>
                        <div style="font-size: 12.5px; color: #15803d; line-height: 1.4;">Though we are a brokerage, we leverage our network to arrange Sea/Air freight at highly discounted, transparent rates.</div>
                    </td>
                </tr>
            </table>

            <p>Our goal is simple: to handle your logistics needs at the most reasonable, fair price without compromising on reliability. Whether you need us just for clearing, or to handle the entire shipment at a low cost, we are ready to support you.</p>

            <p style="font-weight: 700; color: #0f172a; margin-top: 25px; margin-bottom: 10px;">Our Services & Commitment:</p>
            <ul class="feature-list">
                <li class="feature-item"><span class="badge">Clearing</span><div><strong>Licensed Customs Brokerage:</strong> Our core service. Swift and fully compliant clearance.</div></li>
                <li class="feature-item"><span class="badge">Freight</span><div><strong>Cost-Effective Freight:</strong> Arranging FCL/LCL shipments at very fair, competitive prices.</div></li>
                <li class="feature-item"><span class="badge">Docs</span><div><strong>Trade Documentation:</strong> Preparation of BL, COO, phytosanitary certs, and related docs.</div></li>
                <li class="feature-item"><span class="badge">Transport</span><div><strong>Affordable Transport:</strong> Island-wide haulage and port pickups at reasonable rates.</div></li>
            </ul>

            <p>If you have an upcoming shipment for <strong>{exporter_name}</strong> and are looking for a fair-priced alternative to your current agent, please reach out for a quick quote.</p>

            <div class="cta-container">
                <a href="https://wa.me/94710940466" class="cta-button">💬 Get a Quick Quote via WhatsApp (+94 71 094 0466)</a>
            </div>

            <div class="signature-divider">
                <table class="sig-table">
                    <tr>
                        <td class="sig-logo-cell">
                            <a href="https://bastellogistics.com" target="_blank">
                                <img src="{LOGO_URL}" alt="Bastel Logo" width="100" style="width: 100px; max-width: 100px; height: auto; display: block; border: 0;" />
                            </a>
                        </td>
                        <td class="sig-info-cell">
                            <div class="sig-name">Export Management Team</div>
                            <div class="sig-title">BASTEL PRIVATE LIMITED</div>
                            <table cellpadding="0" cellspacing="0" border="0" style="font-size: 12.5px; color: #475569;">
                                <tr><td style="padding-bottom: 4px; color: #0d9488; font-weight: bold; width: 20px;">📧</td><td style="padding-bottom: 4px;"><a href="mailto:export@bastellogistics.com" style="color: #334155; text-decoration: none; font-weight: 500;">export@bastellogistics.com</a></td></tr>
                                <tr><td style="padding-bottom: 4px; color: #0d9488; font-weight: bold;">📱</td><td style="padding-bottom: 4px;"><a href="https://wa.me/94710940466" style="color: #334155; text-decoration: none; font-weight: 500;">+94 71 094 0466</a></td></tr>
                                <tr><td style="padding-bottom: 4px; color: #0d9488; font-weight: bold;">🌐</td><td style="padding-bottom: 4px;"><a href="https://bastellogistics.com" target="_blank" class="sig-link">www.bastellogistics.com</a></td></tr>
                                <tr><td style="padding-bottom: 4px; color: #0d9488; font-weight: bold;">📍</td><td style="padding-bottom: 4px; color: #64748b;">Colombo, Sri Lanka</td></tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </div>
        </div>
        <div class="footer">
            © 2026 Bastel Pvt Ltd. All rights reserved. <br>
            Official Communications • <a href="mailto:export@bastellogistics.com" style="color: #94a3b8;">export@bastellogistics.com</a>
        </div>
    </div>
    </body>
    </html>
    """
    return subject, plain_text, html_body

def parse_emails_from_cell(cell_value):
    if not cell_value:
        return []
    raw_list = re.split(r'[,;\s]+', str(cell_value))
    valid_emails = []
    for item in raw_list:
        clean = item.strip()
        if '@' in clean and '.' in clean:
            valid_emails.append(clean)
    return valid_emails

# ==========================================
# 3. SCRIPT EXECUTION
# ==========================================
def send_emails():
    wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheet = wb[SHEET_NAME]
    
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(BREVO_LOGIN_USER, SENDER_PASSWORD)
        print("✓ Connected to Brevo SMTP server successfully!")
    except Exception as e:
        print(f"✗ Failed to connect to Brevo: {e}")
        return

    max_row = sheet.max_row
    
    if TEST_MODE:
        print("\n--- RUNNING IN TEST MODE (Row 2 Only) ---")
        end_row = 2
        start_row = 2
    else:
        print(f"\n--- RUNNING FROM ROW {START_ROW} to {max_row} ---")
        start_row = START_ROW
        end_row = max_row

    sent_count = 0

    for row in range(start_row, end_row + 1):
        exporter_name = sheet.cell(row=row, column=2).value
        val_g = sheet.cell(row=row, column=7).value
        val_n = sheet.cell(row=row, column=14).value

        if not exporter_name:
            print(f"Skipping Row {row}: Missing Exporter Name")
            continue

        exporter_name = str(exporter_name).strip()
        
        emails_g = parse_emails_from_cell(val_g)
        emails_n = parse_emails_from_cell(val_n)
        
        all_recipients = list(dict.fromkeys(emails_g + emails_n))

        if not all_recipients:
            print(f"Skipping Row {row} ({exporter_name}): No valid emails found")
            continue

        subject, plain_text, html_body = get_email_content(exporter_name)

        msg = MIMEMultipart('alternative')
        msg['From'] = f"Bastel Pvt Ltd <{SENDER_EMAIL}>"
        msg['To'] = ", ".join(all_recipients)
        msg['Subject'] = subject
        msg['Reply-To'] = SENDER_EMAIL
        msg['Date'] = formatdate(localtime=True)
        msg['Message-ID'] = make_msgid(domain="bastellogistics.com")

        msg.attach(MIMEText(plain_text, 'plain'))
        msg.attach(MIMEText(html_body, 'html'))

        try:
            server.sendmail(SENDER_EMAIL, all_recipients, msg.as_string())
            print(f"✓ Row {row}: Email sent to '{exporter_name}' -> Recipients: {', '.join(all_recipients)}")
            sent_count += 1
        except Exception as e:
            print(f"✗ Row {row}: Failed to send email to {all_recipients}. Error: {e}")

        if sent_count > 0 and sent_count % 25 == 0:
            print("\n--- Batch Limit Reached (25 emails). Pausing for 20 seconds... ---")
            time.sleep(20)
        else:
            time.sleep(random.uniform(5, 10))

    server.quit()
    print("\nProcess Completed!")

if __name__ == "__main__":
    send_emails()