import datetime
import os
import json
import gspread
import pandas as pd
import smtplib

from oauth2client.service_account import ServiceAccountCredentials
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# ---------------------------------------
# CONFIG
# ---------------------------------------
SPREADSHEET_ID = "1LdhvCL0-mEg66QI_83B_rXXWMXTMDrglGubR2gsEhF0"

EMAIL_FROM = os.getenv("EMAIL_FROM")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
EMAIL_TO = os.getenv("EMAIL_TO").split(",")

EMAIL_SUBJECT = "Officer Not Done List Report"

EMAIL_BODY = """
Dear user,

Please find attached the Officer not done report.

Regards,
Automation System
"""

# ---------------------------------------
# DATE
# ---------------------------------------
today = datetime.date.today()

month_name = today.strftime("%B")
year_short = today.strftime("%y")

SHEET_NAME = f"{month_name} {year_short}"


# ---------------------------------------
# GOOGLE CREDS
# ---------------------------------------
google_creds_json = os.getenv("GOOGLE_CREDS_JSON")

with open("credentials.json", "w") as f:
    json.dump(json.loads(google_creds_json), f)


# ---------------------------------------
# REPORT FUNCTION
# ---------------------------------------
def generate_report():

    month_start = today.replace(day=1)
    yesterday = today - datetime.timedelta(days=1)

    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        "credentials.json",
        scope
    )

    client = gspread.authorize(creds)

    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

    data = sheet.get_all_values()

    if not data:
        print("No data found.")
        return None

    df = pd.DataFrame(data[1:], columns=data[0])

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    filtered_df = df[
        (df["Date"] >= pd.to_datetime(month_start)) &
        (df["Date"] < pd.to_datetime(today)) &
        (df["Insp. Done"].str.upper() == "NO") &
        (df["Def. Submitted"].str.upper() == "NO")
    ]

    if filtered_df.empty:
        print("No pending records.")
        return None

    report_df = filtered_df[
        [
            "SN",
            "Date",
            "Name",
            "Designation",
            "Department",
            "Contact",
            "Remarks"
        ]
    ].copy()

    report_df.reset_index(drop=True, inplace=True)
    report_df["Sr. No"] = report_df.index + 1
    report_df["Date"] = report_df["Date"].dt.strftime("%Y-%m-%d")

    cols = [
        "Sr. No",
        "Date",
        "Name",
        "Designation",
        "Department",
        "Contact",
        "Remarks"
    ]

    report_df = report_df[cols]

    # -----------------------
    # WORD FILE
    # -----------------------
    doc = Document()

    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11)
    section.page_height = Inches(8.5)

    heading = doc.add_heading(
        f"Officer not done list ({month_start} to {yesterday})",
        level=1
    )

    run = heading.runs[0]
    run.bold = True
    run.underline = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(255, 0, 0)

    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table = doc.add_table(rows=1, cols=len(cols))
    table.style = "Table Grid"

    for i, c in enumerate(cols):
        table.rows[0].cells[i].text = c

    for _, row in report_df.iterrows():
        cells = table.add_row().cells
        for i, c in enumerate(cols):
            cells[i].text = str(row[c])

    filename = f"not_done_report_{today}.docx"

    doc.save(filename)

    return filename


# ---------------------------------------
# EMAIL FUNCTION
# ---------------------------------------
def send_email(file):

    msg = MIMEMultipart()

    msg["From"] = EMAIL_FROM
    msg["To"] = ",".join(EMAIL_TO)
    msg["Subject"] = EMAIL_SUBJECT

    msg.attach(MIMEText(EMAIL_BODY, "plain"))

    with open(file, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())

    encoders.encode_base64(part)

    part.add_header(
        "Content-Disposition",
        f"attachment; filename={file}"
    )

    msg.attach(part)

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(EMAIL_FROM, EMAIL_PASSWORD)
        server.sendmail(
            EMAIL_FROM,
            EMAIL_TO,
            msg.as_string()
        )

    print("Email sent successfully.")


# ---------------------------------------
# MAIN
# ---------------------------------------
if __name__ == "__main__":

    # Monday = 0
    if today.weekday() != 0:
        print("Today is not Monday. Exiting.")
        exit()

    print("Monday confirmed. Running...")

    file = generate_report()

    if file:
        send_email(file)
