import os
import time
from dotenv import load_dotenv
from utils import read_csv, read_html_template, connect_smtp, create_message
from email_config import EVENTS

load_dotenv()

SENDER_EMAIL = os.getenv("SENDER_EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")
EVENT = os.getenv("EVENT")
BATCH_SIZE = int(os.getenv("BATCH_SIZE", 50))
PAUSE_DURATION = int(os.getenv("PAUSE_DURATION", 60))

missing = [v for v in ["SENDER_EMAIL", "APP_PASSWORD", "EVENT"] if not os.getenv(v)]
if missing:
    raise EnvironmentError(f"❌ Missing required environment variables: {', '.join(missing)}")

if EVENT not in EVENTS:
    raise ValueError(f"❌ Unknown event '{EVENT}'. Available: {', '.join(EVENTS)}")

TIER_LABELS = {
    "1": ("☁️ Cloud Champion", "tier-1"),
    "2": ("☁️ Cloud Explorer", "tier-2"),
    "3": ("☁️ Cloud Learner",  "tier-3"),
}

config = EVENTS[EVENT]
participants = read_csv(config["csv"])
html_template = read_html_template(config["template"])
subject = config["subject"]
event_dir = config["event_dir"]
attachments_by_field = config.get("attachments_by_field")

server = connect_smtp(SENDER_EMAIL, APP_PASSWORD)
print(f"📨 Sending '{EVENT}' emails to {len(participants)} participants...\n")

for i, participant in enumerate(participants, start=1):
    participant["name"] = participant.get("name", "").title()
    tier_label, tier_badge_class = TIER_LABELS.get(participant.get("tier", ""), ("", ""))
    participant["tier_label"] = tier_label
    participant["tier_badge_class"] = tier_badge_class

    try:
        attachments = None
        if attachments_by_field:
            field_val = participant.get(attachments_by_field["field"], "").strip()
            attachments = attachments_by_field["map"].get(field_val, [])

        msg = create_message(
            sender=SENDER_EMAIL,
            recipient=participant["email"],
            subject=subject,
            html_template=html_template,
            participant=participant,
            event_dir=event_dir,
            attachments=attachments,
        )
        server.send_message(msg)
        print(f"[{i}/{len(participants)}] ✅ Sent to: {participant['email']} ({participant['name']})")
    except Exception as e:
        print(f"[{i}/{len(participants)}] ❌ Failed: {participant['email']}: {e}")

    if i % BATCH_SIZE == 0 and i < len(participants):
        print(f"\n⏸️ Pausing {PAUSE_DURATION}s to avoid rate limits...\n")
        time.sleep(PAUSE_DURATION)

server.quit()
print("\n🎉 All emails processed!")
