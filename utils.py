import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from string import Template
import os


def safe_int(value):
    """Safely convert any numeric-like value to int (handles '0.0', '', NaN, etc.)."""
    try:
        return int(float(str(value).strip() or 0))
    except (ValueError, TypeError):
        return 0


def read_csv(csv_file):
    """
    Reads participant progress CSV and returns a list of participant data.
    Expected columns:
    ['User Name', 'User Email', 'Access Code Redemption Status',
     'All Skill Badges & Games Completed',
     '# of Skill Badges Completed', '# of Arcade Games Completed']
    """
    df = pd.read_csv(csv_file)

    # Strip spaces and normalize column names
    df.columns = [c.strip() for c in df.columns]

    required_cols = [
        'User Name', 'User Email', 'Access Code Redemption Status',
        'All Skill Badges & Games Completed',
        '# of Skill Badges Completed', '# of Arcade Games Completed'
    ]

    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"❌ Missing column in CSV: {col}")

    participants = []
    for _, row in df.iterrows():
        # Extract and sanitize data
        skill_badges = safe_int(row['# of Skill Badges Completed'])
        arcade_games = safe_int(row['# of Arcade Games Completed'])
        completed_labs = skill_badges + arcade_games
        progress = round((completed_labs / 20) * 100)

        participants.append({
            "name": str(row['User Name']).strip(),
            "email": str(row['User Email']).strip(),
            "redemption_status": str(row['Access Code Redemption Status']).strip(),
            "all_completed": str(row['All Skill Badges & Games Completed']).strip(),
            "skill_badges": skill_badges,
            "arcade_games": arcade_games,
            "completed_labs": completed_labs,
            "progress": progress,
            "total_labs": completed_labs,
            "progress_percent": progress
        })

    return participants


def read_html_template(template_file):
    """Reads the HTML email template."""
    with open(template_file, "r", encoding="utf-8") as f:
        return f.read()


def connect_smtp(sender_email, app_password):
    """
    Connects to Gmail SMTP using App Password.
    Returns an active SMTP session.
    """
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender_email, app_password)
    return server


def create_message(sender, recipient, subject, html_template, participant):
    """
    Creates a personalized MIME email with inline images and variables from participant dict.
    """
    # Prepare HTML body with dynamic values
    template = Template(html_template)
    body_html = template.safe_substitute(
        name=participant["name"],
        redemption_status=participant["redemption_status"],
        all_completed=participant["all_completed"],
        skill_badges=participant["skill_badges"],
        arcade_games=participant["arcade_games"],
        completed_labs=participant["completed_labs"],
        progress=participant["progress"],
        total_labs=participant["total_labs"],
        progress_percent=participant["progress_percent"],
        progress_style=f"{participant['progress']}%"
    )

    # Create email container
    msg = MIMEMultipart('related')
    msg['From'] = sender
    msg['To'] = participant["email"]
    msg['Subject'] = subject

    # Attach HTML body
    msg_alt = MIMEMultipart('alternative')
    msg.attach(msg_alt)
    msg_alt.attach(MIMEText(body_html, 'html', 'utf-8'))

    # === Image Attachments ===
    script_dir = os.path.dirname(os.path.abspath(__file__))
    event_data_dir = os.path.join(script_dir, 'event_data')
    social_dir = os.path.join(script_dir, 'social_icons')

    # Header image path (adjust per campaign/state if needed)
    header_path = os.path.join(event_data_dir, 'access_code_claimed_yes', 'emailHeaderFinalWeek.png')
    if os.path.exists(header_path):
        with open(header_path, 'rb') as f:
            img = MIMEImage(f.read())
            img.add_header('Content-ID', '<header>')
            img.add_header('Content-Disposition', 'inline', filename='header.png')
            msg.attach(img)
    else:
        print(f"⚠️ Header image not found: {header_path}")

    # Attach social icons
    social_icons = ['instagram', 'google', 'linkedin', 'linktree']
    for icon in social_icons:
        icon_path = os.path.join(social_dir, f'{icon}.png')
        if os.path.exists(icon_path):
            with open(icon_path, 'rb') as f:
                img = MIMEImage(f.read())
                img.add_header('Content-ID', f'<{icon}>')
                img.add_header('Content-Disposition', 'inline', filename=f'{icon}.png')
                msg.attach(img)
        else:
            print(f"⚠️ Missing icon: {icon_path}")

    return msg
