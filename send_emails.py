import os
import time
from dotenv import load_dotenv
from utils import read_csv, read_html_template, connect_smtp, create_message
from email_config import SUBJECT

# Load env variables
load_dotenv()
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")
CSV_FILE = os.getenv("CSV_FILE")
TEMPLATE_FILE = os.getenv("TEMPLATE_FILE")
BATCH_SIZE = int(os.getenv("BATCH_SIZE", 50))
PAUSE_DURATION = int(os.getenv("PAUSE_DURATION", 60))

# Read participants and template
emails, names = read_csv(CSV_FILE)
html_template = read_html_template(TEMPLATE_FILE)

# Connect to SMTP
server = connect_smtp(SENDER_EMAIL, APP_PASSWORD)

# Send emails
for i, (recipient, name) in enumerate(zip(emails, names), start=1):
    body = html_template.replace("{name}", name)
    msg = create_message(SENDER_EMAIL, recipient, SUBJECT, body)

    try:
        server.send_message(msg)
        print(f"[{i}/{len(emails)}] Email sent to: {recipient}")
    except Exception as e:
        print(f"Failed to send email to {recipient}: {e}")

    if i % BATCH_SIZE == 0:
        print(f"Pausing for {PAUSE_DURATION} seconds to avoid Gmail limits...")
        time.sleep(PAUSE_DURATION)

server.quit()
print("All emails sent!")
