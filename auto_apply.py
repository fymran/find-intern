import csv
import os
import smtplib
import imaplib
import time
from datetime import datetime
from email.message import EmailMessage
from docx import Document
from docx2pdf import convert

# ==============================================================================
# CONFIGURATION
# ==============================================================================
GMAIL_USER = "hafymran@gmail.com"
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "YOUR_16_CHAR_APP_PASSWORD")

# Set MODE = "DRAFT" to safely create drafts in your Gmail (Recommended first!)
# Set MODE = "SEND" to immediately dispatch emails to HR
MODE = "DRAFT"

LEADS_CSV = "leads.csv"
OUTPUT_DIR = "generated_sli02"

# Static PDF files to attach to every application
# Note: Put your Transcript PDF in this list as soon as UiTM issues it!
STATIC_ATTACHMENTS = [
    "SLI01_2025189891.pdf",
    "Muhammad_Hafiy_Imran_Resume.pdf",
    "DLI01_UiTM_Objectives_Scope.pdf",
    # "Official_Transcript_UiTM.pdf",  # <-- Uncomment once you download your transcript
]

os.makedirs(OUTPUT_DIR, exist_ok=True)

def replace_placeholders_in_docx(template_path, output_docx_path, replacements):
    doc = Document(template_path)
    for p in doc.paragraphs:
        full_text = p.text
        modified = False
        for key, val in replacements.items():
            if key in full_text:
                full_text = full_text.replace(key, val)
                modified = True
        if modified:
            # Preserve paragraph style while updating runs cleanly
            for i, run in enumerate(p.runs):
                if i == 0:
                    run.text = full_text
                else:
                    run.text = ""
    doc.save(output_docx_path)

def build_sli02_for_company(row):
    company = row["Company_Name"].strip()
    role = row["Target_Role"].strip()
    lang = row.get("Lang_Pref", "EN").strip().upper()
    hr_name = row.get("HR_Name", "").strip() or ("Pengurus Sumber Manusia" if lang == "BM" else "Hiring Manager")
    hr_role = row.get("HR_Role", "").strip() or ("Jabatan Sumber Manusia" if lang == "BM" else "Talent Acquisition Team")
    address = row.get("Company_Address", "").strip() or row.get("Location", "Selangor / Kuala Lumpur")

    safe_company = "".join(c if c.isalnum() else "_" for c in company).strip("_")
    temp_docx = os.path.join(OUTPUT_DIR, f"SLI02_{safe_company}.docx")
    final_pdf = os.path.join(OUTPUT_DIR, f"SLI02_{safe_company}.pdf")

    today_en = datetime.now().strftime("%d %B %Y")
    bulan_map = {
        1: "Januari", 2: "Februari", 3: "Mac", 4: "April", 5: "Mei", 6: "Jun",
        7: "Julai", 8: "Ogos", 9: "September", 10: "Oktober", 11: "November", 12: "Disember"
    }
    now = datetime.now()
    today_bm = f"{now.day} {bulan_map[now.month]} {now.year}"

    if lang == "BM":
        template = "SLI02_Template_BM.docx"
        replacements = {
            "{{TARIKH}}": today_bm,
            "{{NAMA_PEGAWAI}}": hr_name,
            "{{JAWATAN_PEGAWAI}}": hr_role,
            "{{NAMA_SYARIKAT}}": company,
            "{{ALAMAT_SYARIKAT}}": address,
            "{{JAWATAN_DIPOHON}}": role
        }
    else:
        template = "SLI02_Template_EN.docx"
        replacements = {
            "{{DATE}}": today_en,
            "{{HR_NAME}}": hr_name,
            "{{HR_ROLE}}": hr_role,
            "{{COMPANY_NAME}}": company,
            "{{COMPANY_ADDRESS}}": address,
            "{{SALUTATION}}": hr_name,
            "{{TARGET_ROLE}}": role
        }

    replace_placeholders_in_docx(template, temp_docx, replacements)
    convert(temp_docx, final_pdf)
    return final_pdf

def build_email_message(row, sli02_pdf_path):
    company = row["Company_Name"].strip()
    role = row["Target_Role"].strip()
    hr_email = row["HR_Email"].strip()
    lang = row.get("Lang_Pref", "EN").strip().upper()

    msg = EmailMessage()
    msg["From"] = f"Muhammad Hafiy Imran <{GMAIL_USER}>"
    msg["To"] = hr_email

    if lang == "BM":
        msg["Subject"] = f"Permohonan Latihan Industri (15 Mac - 18 Jun 2027): {role} - Muhammad Hafiy Imran"
        body = f"""Ke hadapan Pihak Pengurusan Sumber Manusia di {company},

Saya Muhammad Hafiy Imran Bin Khumaidi, pelajar tahun akhir Ijazah Sarjana Muda Sistem Maklumat (Kepujian) Kejuruteraan Sistem Pintar dari Universiti Teknologi MARA (UiTM) Shah Alam (CGPA: 3.82).

Tujuan e-mel ini adalah untuk memohon penempatan latihan industri bagi posisi {role} di {company} selama 14 minggu bermula 15 Mac 2027 hingga 18 Jun 2027.

Bersama-sama e-mel ini, saya lampirkan dokumen rasmi berikut untuk rujukan pihak Tuan/Puan:
1. Surat Pengenalan Rasmi UiTM (SLI-01)
2. Surat Permohonan Rasmi (SLI-02)
3. Resume Terkini
4. Objektif dan Skop Latihan Industri UiTM (DLI-01)

Segala perhatian dan pertimbangan daripada pihak Tuan/Puan amatlah saya hargai.

Sekian, terima kasih.

Yang benar,
Muhammad Hafiy Imran Bin Khumaidi
+60 13-502 5738
"""
    else:
        msg["Subject"] = f"Internship Application (15 Mar - 18 Jun 2027): {role} - Muhammad Hafiy Imran"
        body = f"""Dear Hiring Team at {company},

I am a final-year Bachelor of Information Systems (Hons.) Intelligent Systems Engineering student at Universiti Teknologi MARA (UiTM) Shah Alam (CGPA: 3.82). I am writing to apply for the {role} internship position at {company} for my 14-week university industrial training from 15 March 2027 to 18 June 2027.

Having previously developed automated incident escalation web tools using Flask and MySQL during my internship at TM Technology Services, alongside my final-year research building a multi-objective portfolio optimization engine in Python (NSGA-II), I am excited by the opportunity to contribute to your engineering and analytics initiatives.

Attached for your review are my required university placement documents:
1. Official University Placement Request Letter (SLI-01)
2. Formal Application Cover Letter (SLI-02)
3. Updated Resume
4. University Industrial Training Objectives & Scope (DLI-01)

Thank you for your time and consideration. I look forward to hearing from you.

Best regards,

Muhammad Hafiy Imran Bin Khumaidi
+60 13-502 5738
"""

    msg.set_content(body)

    # Attach SLI-02 + Static Attachments
    attachments = [sli02_pdf_path] + [f for f in STATIC_ATTACHMENTS if os.path.exists(f)]
    for path in attachments:
        with open(path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="pdf",
                filename=os.path.basename(path)
            )
    return msg

def process_leads():
    with open(LEADS_CSV, "r", encoding="utf-8") as f:
        reader = list(csv.DictReader(f))

    updated = False
    for row in reader:
        status = row.get("Status", "").strip()
        hr_email = row.get("HR_Email", "").strip()

        if status == "Ready" and hr_email:
            print(f"[*] Generating tailored SLI-02 PDF for {row['Company_Name']}...")
            sli02_pdf = build_sli02_for_company(row)
            msg = build_email_message(row, sli02_pdf)

            if MODE == "DRAFT":
                with imaplib.IMAP4_SSL("imap.gmail.com", 993) as imap:
                    imap.login(GMAIL_USER, GMAIL_APP_PASSWORD)
                    imap.append("[Gmail]/Drafts", "", imaplib.Time2Internaldate(time.time()), msg.as_bytes())
                row["Status"] = "Draft Created"
                print(f"    [✓] Created Gmail Draft for {row['Company_Name']} ({hr_email})")
            elif MODE == "SEND":
                with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                    smtp.login(GMAIL_USER, GMAIL_APP_PASSWORD)
                    smtp.send_message(msg)
                row["Status"] = f"Sent ({datetime.now().strftime('%Y-%m-%d')})"
                print(f"    [✓] Sent email to {row['Company_Name']} ({hr_email})")
            updated = True

    if updated:
        with open(LEADS_CSV, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=reader[0].keys())
            writer.writeheader()
            writer.writerows(reader)

if __name__ == "__main__":
    process_leads()
