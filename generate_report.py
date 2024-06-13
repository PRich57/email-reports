import imaplib
import email as email_lib
from email.header import decode_header
import pandas as pd
import os
from datetime import datetime, timedelta

# Email account credentials
username = os.getenv('EMAIL_USER')  # Fetching from environment variable
password = os.getenv('EMAIL_PASS')  # Fetching from environment variable

# Path to save the report
file_dir = 'C:\Users\pcric\OneDrive\Desktop\my_reports'
file_path = os.path.join(file_dir, 'daily_email_report.csv')

def create_report(file_path):
    # Ensure the directory exists
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    
    # Connect to the server
    imap = imaplib.IMAP4_SSL("imap.gmail.com")

    # Login to the account
    imap.login(username, password)

    # Select the mailbox you want to check
    imap.select("inbox")

    # Search for emails received today
    today = datetime.now().strftime("%d-%b-%Y")
    status, messages = imap.search(None, f'(SINCE {today})')

    # Convert messages to a list of email IDs
    messages = messages[0].split()

    # Initialize a list to store the email data
    email_data = []

    for msg_id in messages:
        # Fetch the email by ID
        status, msg = imap.fetch(msg_id, "(RFC822)")

        # Parse the email content
        for response_part in msg:
            if isinstance(response_part, tuple):
                msg = email_lib.message_from_bytes(response_part[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")
                from_ = msg.get("From")
                date_ = msg.get("Date")

                # Extract the email body
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if "attachment" not in content_disposition:
                            if content_type == "text/plain":
                                body = part.get_payload(decode=True).decode()
                                break
                else:
                    body = msg.get_payload(decode=True).decode()

                email_data.append([date_, from_, subject, body])

    # Create a DataFrame and save as CSV
    df = pd.DataFrame(email_data, columns=["Date", "From", "Subject", "Body"])
    df.to_csv(file_path, index=False)

    # Logout and close the connection
    imap.logout()

# Create the report file
create_report(file_path)
print(f"Report created at {file_path}")
