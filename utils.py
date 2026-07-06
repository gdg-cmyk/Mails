import pandas as pd
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from string import Template


def read_csv(csv_file):
    df = pd.read_csv(csv_file)
    df.columns = [c.strip() for c in df.columns]

    # Support both column naming conventions
    name_col = "User Name" if "User Name" in df.columns else "name"
    email_col = "User Email" if "User Email" in df.columns else "email"

    if name_col not in df.columns or email_col not in df.columns:
        raise ValueError("CSV must contain name/email columns")

    participants = []
    for _, row in df.iterrows():
        email = str(row[email_col]).strip()
        if email == "—" or not email or email.lower() == "nan":
            continue
        participant = {"name": str(row[name_col]).strip(), "email": email}
        for col in df.columns:
            key = col.strip().lower().replace(" ", "_").replace("#_of_", "").replace("&_", "").replace("/", "_").replace("-", "_")
            if key not in participant:
                participant[key] = str(row[col]).strip() if pd.notna(row[col]) else ""
        participants.append(participant)

    return participants


def read_html_template(template_file):
    with open(template_file, "r", encoding="utf-8") as f:
        return f.read()


def connect_smtp(sender_email, app_password):
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender_email, app_password)
    return server


def _find_header_image(event_dir):
    for fname in os.listdir(event_dir):
        if fname.lower().endswith(".png"):
            return os.path.join(event_dir, fname)
    return None


def create_message(sender, recipient, subject, html_template, participant, event_dir, attachments=None):
    template = Template(html_template)
    body_html = template.safe_substitute(**participant)

    msg = MIMEMultipart("related")
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject

    msg_alt = MIMEMultipart("alternative")
    msg.attach(msg_alt)
    msg_alt.attach(MIMEText(body_html, "html", "utf-8"))

    script_dir = os.path.dirname(os.path.abspath(__file__))
    abs_event_dir = os.path.join(script_dir, event_dir)
    social_dir = os.path.join(script_dir, "social_icons")

    header_path = _find_header_image(abs_event_dir)
    if header_path:
        with open(header_path, "rb") as f:
            img = MIMEImage(f.read())
            img.add_header("Content-ID", "<header>")
            img.add_header("Content-Disposition", "inline", filename="header.png")
            msg.attach(img)
    else:
        print(f"⚠️ No header image found in: {abs_event_dir}")

    dark_theme_events = ["study_jams_goodies_distribution"]
    icon_suffix = "_white" if any(e in event_dir for e in dark_theme_events) else ""

    for icon in ["instagram", "google", "linkedin", "linktree"]:
        icon_path = os.path.join(social_dir, f"{icon}{icon_suffix}.png")
        if os.path.exists(icon_path):
            with open(icon_path, "rb") as f:
                img = MIMEImage(f.read())
                img.add_header("Content-ID", f"<{icon}>")
                img.add_header("Content-Disposition", "inline", filename=f"{icon}.png")
                msg.attach(img)
        else:
            print(f"⚠️ Missing icon: {icon_path}")

    for attachment_path in (attachments or []):
        abs_path = os.path.join(script_dir, attachment_path)
        if os.path.exists(abs_path):
            with open(abs_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", "attachment", filename=os.path.basename(abs_path))
            msg.attach(part)
        else:
            print(f"⚠️ Missing attachment: {abs_path}")

    return msg
