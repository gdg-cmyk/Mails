import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from string import Template
import os

def read_csv(csv_file):
    df = pd.read_csv(csv_file)
    emails = df.iloc[:, 1].tolist()  # second column = emails
    names = df.iloc[:, 0].tolist() if df.shape[1] > 1 else ["Participant"]*len(emails)
    return emails, names

def read_html_template(template_file):
    with open(template_file, "r", encoding="utf-8") as f:
        return f.read()

def connect_smtp(sender_email, app_password):
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender_email, app_password)
    return server

def create_message(sender, recipient, subject, html_template, name):
    # Use string.Template for safe substitution
    template = Template(html_template)
    body_html = template.substitute(name=name)

    msg = MIMEMultipart('related')
    msg['From'] = sender
    msg['To'] = recipient
    msg['Subject'] = subject

    # Attach HTML body
    msg_alt = MIMEMultipart('alternative')
    msg.attach(msg_alt)
    msg_alt.attach(MIMEText(body_html, 'html'))

    # Attach header image
    with open(os.path.join('event_data', 'cloud_study_jams_bevy_meetup_header.png'), 'rb') as f:
        img = MIMEImage(f.read())
        img.add_header('Content-ID', '<header>')
        img.add_header('Content-Disposition', 'inline', filename='header.png')
        msg.attach(img)

    # Attach social icons
    script_dir = os.path.dirname(os.path.abspath(__file__))
    social_icons = ['instagram', 'google', 'linkedin', 'linktree']

    for icon in social_icons:
        path = os.path.join(script_dir, 'social_icons', f'{icon}.png')
        if os.path.exists(path):
            with open(path, 'rb') as f:
                img = MIMEImage(f.read())
                img.add_header('Content-ID', f'<{icon}>')
                img.add_header('Content-Disposition', 'inline', filename=f'{icon}.png')
                msg.attach(img)
        else:
            print(f"Path not found: {path}")


    return msg
