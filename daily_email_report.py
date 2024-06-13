import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Email configuration
smtp_server = 'smtp.gmail.com'  # For Gmail
smtp_port = 587
sender_email = os.getenv('EMAIL_USER') # Fetching from environment variable
receiver_email = os.getenv('EMAIL_USER') # Set as my own email
password = os.getenv('EMAIL_PASS')  # Fetching from environment variable

# Email content
subject = 'Daily Report'
body = 'Please find the attached daily report.'

# File to attach
file_path = 'C:\Users\pcric\OneDrive\Desktop\daily_report.txt'

def create_report(file_path):
    # Generate the report content
    report_content = "This is your daily report. \n\n"

    # Write the report to a file
    with open(file_path, 'w') as report_file:
        report_file.write(report_content)

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
        attachment = open(file_path, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(file_path)}")
        msg.attach(part)

    # Create SMTP session for sending the mail
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()  # Enable security
    server.login(from_email, password)  # Login with mail_id and password
    text = msg.as_string()
    server.sendmail(from_email, to_email, text)
    server.quit()
    print(f'Email sent to {to_email} successfully.')

create_report(file_path)

# Send the email
send_email(subject, body, receiver_email, sender_email, password, file_path)
