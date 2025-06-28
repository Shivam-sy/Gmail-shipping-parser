import os
import re
import email
import google.generativeai as genai
from bs4 import BeautifulSoup
from email import policy
from email.parser import BytesParser
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


SPREADSHEET_ID = 'Enter your SPREADSHEET_ID'
RANGE_NAME = 'Sheet1!A2'
EML_FOLDER = 'emls'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
GEMINI_API_KEY = 'Enter your GOOGLE_GEMINI_API_KEY'


genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")



def authenticate_sheets():
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
    return build('sheets', 'v4', credentials=creds)


def extract_from_eml(file_path):
    with open(file_path, 'rb') as f:
        msg = BytesParser(policy=policy.default).parse(f)

    sender = msg['From']
    date_raw = msg['Date']
    date_obj = email.utils.parsedate_to_datetime(date_raw)
    date = date_obj.strftime('%Y-%m-%d %H:%M:%S')

    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if "attachment" in str(part.get("Content-Disposition")):
                continue
            if content_type == "text/plain":
                body = part.get_payload(decode=True).decode(errors='ignore')
                break
            elif content_type == "text/html":
                html = part.get_payload(decode=True).decode(errors='ignore')
                soup = BeautifulSoup(html, "lxml")
                body = soup.get_text(separator=' ', strip=True)
    else:
        content_type = msg.get_content_type()
        if content_type == "text/plain":
            body = msg.get_payload(decode=True).decode(errors='ignore')
        elif content_type == "text/html":
            html = msg.get_payload(decode=True).decode(errors='ignore')
            soup = BeautifulSoup(html, "lxml")
            body = soup.get_text(separator=' ', strip=True)

    body = re.sub(r'\s+', ' ', body.replace('\t', ' ').replace('’', "'"))
    return sender, date, body


def ask_gemini_to_extract(body):
    prompt = f"""
You are a shipping data extraction assistant. Extract the following fields from the below email body:

- Origin Port
- Destination Port
- 20' Container Price
- 40' Container Price

Return the output in this format:
Origin: ...
Destination: ...
20GP: ...
40HC: ...

Email Content:
\"\"\"
{body}
\"\"\"
"""
    try:
        response = model.generate_content(prompt)
        text = response.text
        origin = re.search(r'Origin: (.+)', text)
        destination = re.search(r'Destination: (.+)', text)
        p20 = re.search(r'20GP: (.+)', text)
        p40 = re.search(r'40HC: (.+)', text)

        return {
            'origin': origin.group(1).strip() if origin else "UNKNOWN",
            'destination': destination.group(1).strip() if destination else "UNKNOWN",
            '20': p20.group(1).strip() if p20 else "NOT FOUND",
            '40': p40.group(1).strip() if p40 else "NOT FOUND"
        }
    except Exception as e:
        print(f"Gemini extraction failed: {e}")
        return {
            'origin': "ERROR",
            'destination': "ERROR",
            '20': "ERROR",
            '40': "ERROR"
        }


def extract_shipping_line(sender):
    domain = sender.split('@')[-1].split('.')[0]
    return domain.capitalize() + " Logistics"


def update_sheet(service, rows):
    sheet = service.spreadsheets()
    sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=RANGE_NAME,
        valueInputOption="RAW",
        body={"values": rows}
    ).execute()


def main():
    print("Reading emails...")
    if not os.path.exists(EML_FOLDER):
        print(f"Folder '{EML_FOLDER}' not found.")
        return

    service = authenticate_sheets()
    rows = []

    for file in os.listdir(EML_FOLDER):
        if not file.endswith('.eml'):
            continue

        path = os.path.join(EML_FOLDER, file)
        print(f"Processing: {file}")
        sender, date, body = extract_from_eml(path)
        shipping_line = extract_shipping_line(sender)
        result = ask_gemini_to_extract(body)

        status = "Partial" if "UNKNOWN" in result.values() or "NOT FOUND" in result.values() else "Complete"

        row = [
          sender,
          date,
          shipping_line.upper(),
          result['origin'].upper(),
          result['destination'].upper(),
          result['20'].upper(),
          result['40'].upper(),
        status
]

        print(f" {file}: {row}")
        rows.append(row)

    if rows:
        update_sheet(service, rows)
        print("Data pushed to Google Sheet!")
    else:
        print("No valid data found.")


if __name__ == '__main__':
    main()
