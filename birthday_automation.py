"""Birthday Automation.

Reads contacts (Name, Email, Birthday) from an Excel file, emails everyone
whose birthday is today, and writes a report Excel file with the send
status (Sent / Failed / Skipped) for every contact.
"""

import os
import sys
import smtplib
import logging
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("birthday_automation.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)

CONTACTS_FILE = os.getenv("CONTACTS_FILE", "contacts.xlsx")
REPORT_FILE = os.getenv("REPORT_FILE", "birthday_report.xlsx")

SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")

DEFAULT_SUBJECT = "Happy Birthday, {name}! 🎂 - From Cholo Jai Dure Tour & Travels"
DEFAULT_MESSAGE = (
    "Dear {name},\n\n"
    "Wishing you a very Happy Birthday from all of us at Cholo Jai Dure Tour & Travels!\n"
    "May your new year of life be filled with joy, laughter and new adventures.\n\n"
    "As a small birthday gift, enjoy an exclusive discount on your next trip with us "
    "this month - because every birthday deserves a getaway!\n\n"
    "Warm wishes,\n"
    "Team Cholo Jai Dure Tour & Travels"
)

HTML_TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "birthday_email.html")
with open(HTML_TEMPLATE_PATH, "r", encoding="utf-8") as _f:
    HTML_TEMPLATE = _f.read()

REQUIRED_COLUMNS = {"Name", "Email", "Birthday"}
# Accepted birthday text formats in the Excel sheet (day-first, to avoid locale ambiguity use dd-mm-yyyy or ISO).
DATE_FORMATS = ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m", "%d/%m")


def load_contacts(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Contacts file not found: {path}")

    # Keep Birthday as the native Excel type (date or text) instead of forcing
    # everything to str, otherwise real date cells get mangled into strings.
    df = pd.read_excel(path, dtype={"Name": str, "Email": str})
    df.columns = [c.strip() for c in df.columns]

    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(f"Contacts file is missing required columns: {missing}")

    return df


def is_birthday_today(birthday_value, today=None):
    today = today or datetime.now()

    # Excel date cells arrive as pandas Timestamp/python datetime, not text.
    if isinstance(birthday_value, (datetime, pd.Timestamp)):
        return birthday_value.day == today.day and birthday_value.month == today.month

    if pd.isna(birthday_value):
        return False

    birthday_str = str(birthday_value).strip()
    for fmt in DATE_FORMATS:
        try:
            bday = datetime.strptime(birthday_str, fmt)
            return bday.day == today.day and bday.month == today.month
        except ValueError:
            continue
    return False


def build_email(name, subject_template, message_template):
    subject = subject_template.format(name=name)
    plain_body = message_template.format(name=name)

    message_html = "".join(f"<p style='margin:0 0 12px;'>{line}</p>" for line in plain_body.split("\n\n") if line.strip())
    html_body = HTML_TEMPLATE.format(name=name, message_html=message_html)

    return subject, plain_body, html_body


def send_email(to_email, subject, plain_body, html_body):
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        raise RuntimeError("SENDER_EMAIL / SENDER_PASSWORD not set (check your .env file)")

    msg = MIMEMultipart("alternative")
    msg["From"] = SENDER_EMAIL
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(plain_body, "plain"))
    msg.attach(MIMEText(html_body, "html"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, to_email, msg.as_string())


def process_contacts(contacts, today=None):
    today = today or datetime.now()
    results = []

    for _, row in contacts.iterrows():
        name = str(row.get("Name", "")).strip()
        email = str(row.get("Email", "")).strip()
        birthday = row.get("Birthday", "")

        subject_template = row.get("Subject")
        if "Subject" not in contacts.columns or pd.isna(subject_template) or not str(subject_template).strip():
            subject_template = DEFAULT_SUBJECT

        message_template = row.get("Personalised Message")
        if "Personalised Message" not in contacts.columns or pd.isna(message_template) or not str(message_template).strip():
            message_template = DEFAULT_MESSAGE

        status, error = "Skipped - Not Birthday", ""
        birthday_today = is_birthday_today(birthday, today)

        if not name or not email or email.lower() == "nan":
            status, error = "Skipped - Missing Data", "Name or Email missing"
            logger.warning("Row skipped due to missing name/email: %s", row.to_dict())
        elif birthday_today:
            try:
                subject, plain_body, html_body = build_email(name, subject_template, message_template)
                send_email(email, subject, plain_body, html_body)
                status = "Sent"
                logger.info("Birthday email sent to %s <%s>", name, email)
            except Exception as exc:
                status = "Failed"
                error = str(exc)
                logger.error("Failed to send email to %s <%s>: %s", name, email, error)

        results.append(
            {
                "Name": name,
                "Email": email,
                "Birthday": birthday,
                "Birthday Today": "Yes" if birthday_today else "No",
                "Message Sent": "Yes" if status == "Sent" else "No",
                "Status": status,
                "Error": error,
                "Checked At": today.strftime("%Y-%m-%d %H:%M:%S"),
            }
        )

    return pd.DataFrame(results)


def run():
    try:
        contacts = load_contacts(CONTACTS_FILE)
    except (FileNotFoundError, ValueError) as exc:
        logger.error(str(exc))
        sys.exit(1)

    report_df = process_contacts(contacts)

    # Report should only list today's birthdays, not every contact.
    todays_df = report_df[report_df["Birthday Today"] == "Yes"].reset_index(drop=True)
    todays_df.to_excel(REPORT_FILE, index=False)
    logger.info("Report written to %s", REPORT_FILE)

    sent = int((todays_df["Status"] == "Sent").sum())
    failed = int((todays_df["Status"] == "Failed").sum())
    logger.info("Summary: %d sent, %d failed (of %d birthdays today)", sent, failed, len(todays_df))


if __name__ == "__main__":
    run()
