import os
import time
from dotenv import load_dotenv
from utils import read_csv, read_html_template, connect_smtp, create_message
from email_config import SUBJECT  # keep as-is

# === Load environment variables ===
load_dotenv()

SENDER_EMAIL = os.getenv("SENDER_EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")
CSV_FILE = os.getenv("CSV_FILE")
TEMPLATE_FILE = os.getenv("TEMPLATE_FILE")
BATCH_SIZE = int(os.getenv("BATCH_SIZE", 50))
PAUSE_DURATION = int(os.getenv("PAUSE_DURATION", 60))

# === Validate required environment variables ===
missing = [var for var in ["SENDER_EMAIL", "APP_PASSWORD", "CSV_FILE", "TEMPLATE_FILE"]
           if not os.getenv(var)]
if missing:
    raise EnvironmentError(f"❌ Missing required environment variables: {', '.join(missing)}")

# === Read input files ===
participants = read_csv(CSV_FILE)
html_template = read_html_template(TEMPLATE_FILE)

# === Connect to Gmail SMTP ===
server = connect_smtp(SENDER_EMAIL, APP_PASSWORD)
print(f"📨 Starting to send {len(participants)} progress report emails...\n")

# === Email sending loop ===
for i, participant in enumerate(participants, start=1):
    try:
        msg = create_message(
            sender=SENDER_EMAIL,
            recipient=participant["email"],
            subject=SUBJECT,
            html_template=html_template,
            participant=participant
        )

        server.send_message(msg)
        print(f"[{i}/{len(participants)}] ✅ Sent to: {participant['email']} ({participant['name']})")
        print(f"   → Progress: {participant['progress']}% ({participant['completed_labs']}/20 labs)\n")

    except Exception as e:
        print(f"[{i}/{len(participants)}] ❌ Failed to send to {participant['email']}: {e}")

    # Gmail anti-spam rate limiting
    if i % BATCH_SIZE == 0 and i < len(participants):
        print(f"\n⏸️ Pausing for {PAUSE_DURATION} seconds to avoid Gmail rate limits...\n")
        time.sleep(PAUSE_DURATION)

# === Wrap up ===
server.quit()
print("\n🎉 All emails have been processed successfully!")
