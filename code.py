import os
import base64
import pdfplumber
import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Get environment variables
EMAIL = os.getenv('EMAIL')
PASSWORD = os.getenv('PASSWORD')
IMAP_SERVER = os.getenv('IMAP_SERVER')
IMAP_PORT = int(os.getenv('IMAP_PORT'))
DOWNLOAD_FOLDER = os.getenv('DOWNLOAD_FOLDER')
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def authenticate_gmail():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def download_latest_pdf(service):
    results = service.users().messages().list(userId='me', q='from:m.motallebi73@gmail.com').execute()
    messages = results.get('messages', [])

    if not messages:
        print("No messages found.")
        return None

    latest_message_id = messages[0]['id']
    message = service.users().messages().get(userId='me', id=latest_message_id).execute()

    for part in message['payload']['parts']:
        if part['filename'] and 'application/pdf' in part['mimeType']:
            if 'data' in part['body']:
                data = part['body']['data']
            else:
                att_id = part['body']['attachmentId']
                att = service.users().messages().attachments().get(userId='me', messageId=latest_message_id, id=att_id).execute()
                data = att['data']
            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
            file_path = os.path.join(DOWNLOAD_FOLDER, part['filename'])
            with open(file_path, 'wb') as f:
                f.write(file_data)
            return file_path
    return None

def extract_table_to_excel(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages):
            print(f"Extracting tables from page {page_number + 1}")
            table = page.extract_table()
            if table:
                print(f"Table found on page {page_number + 1}")
                df = pd.DataFrame(table[1:], columns=table[0])
                tables.append(df)
    
    if tables:
        combined_df = pd.concat(tables, ignore_index=True)
        excel_file_name = os.path.splitext(os.path.basename(pdf_path))[0] + '.xlsx'
        excel_file_path = os.path.join(DOWNLOAD_FOLDER, excel_file_name)

        os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            combined_df.to_excel(writer, sheet_name='Sheet1', index=False)
        print(f"Tables extracted and saved to {excel_file_path}.")
    else:
        print("No table found in the entire PDF.")

if __name__ == '__main__':
    creds = authenticate_gmail()
    service = build('gmail', 'v1', credentials=creds)

    pdf_path = download_latest_pdf(service)
    if pdf_path:
        extract_table_to_excel(pdf_path)
    else:
        print("No PDF found.")
