import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Email account credentials
username = os.getenv('EMAIL_USER')  # Fetching from environment variable
password = os.getenv('EMAIL_PASS')  # Fetching from environment variable

# Email content
subject = 'Daily Email Report'
body = 'Please find the attached daily email report.'

# Path to the report
file_path = r'C:\my_reports\daily_report.csv' # Fix this path

def send_email(subject, body, to_email, from_email, password, file_path):
    # Create the email
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    # Attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))

    # Attach the file
    if file_path:
        with open(file_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(file_path)}")
        msg.attach(part)

    # Create SMTP session for sending the mail
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()  # Enable security
    server.login(from_email, password)  # Login with mail_id and password
    text = msg.as_string()
    server.sendmail(from_email, to_email, text)
    server.quit()
    print(f'Email sent to {to_email} successfully.')

# Send the email with the attached report
send_email(subject, body, username, username, password, file_path)
