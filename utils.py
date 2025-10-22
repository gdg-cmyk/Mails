import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

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

def create_message(sender, recipient, subject, body_html):
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = recipient
    msg['Subject'] = subject
    msg.attach(MIMEText(body_html, 'html'))
    return msg
