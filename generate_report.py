import os
import pickle
import imaplib
import email as email_lib
from email.header import decode_header
import pandas as pd
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# Fetch credentials path from environment variable
CREDENTIALS_JSON_FILE = os.getenv('GOOGLE_APPLICATION_CREDENTIALS')
TOKEN_PICKLE_FILE = 'token.pickle'
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def authenticate():
    creds = None
    if os.path.exists(TOKEN_PICKLE_FILE):
        with open(TOKEN_PICKLE_FILE, 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_JSON_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PICKLE_FILE, 'wb') as token:
            pickle.dump(creds, token)
    return creds

def fetch_emails(creds):
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    auth_string = 'user={}\1auth=Bearer {}\1\1'.format(creds.client_id, creds.token)
    mail.authenticate('XOAUTH2', lambda x: auth_string)
    mail.select("inbox")

    today = datetime.now().strftime("%d-%b-%Y")
    status, messages = mail.search(None, f'(SINCE {today})')
    messages = messages[0].split()

    email_data = []

    for msg_id in messages:
        status, msg = mail.fetch(msg_id, "(RFC822)")
        for response_part in msg:
            if isinstance(response_part, tuple):
                msg = email_lib.message_from_bytes(response_part[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")
                from_ = msg.get("From")
                date_ = msg.get("Date")

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
    
    return email_data

def create_report(file_path, email_data):
    df = pd.DataFrame(email_data, columns=["Date", "From", "Subject", "Body"])
    df.to_csv(file_path, index=False)

def main():
    file_dir = r'C:\Users\pcric\OneDrive\Desktop\my_reports'
    file_path = os.path.join(file_dir, 'daily_email_report.csv')
    if not os.path.exists(file_dir):
        os.makedirs(file_dir)
        print(f"Directory {file_dir} created.")
    else:
        print(f"Directory {file_dir} already exists.")
    
    creds = authenticate()
    email_data = fetch_emails(creds)
    create_report(file_path, email_data)
    print(f"Report created at {file_path}")

if __name__ == "__main__":
    main()
