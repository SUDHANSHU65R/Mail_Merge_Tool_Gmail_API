import time
import base64
import pandas as pd
import pygsheets as pg
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from jinja2 import Environment, FileSystemLoader
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Define global variables
SCOPES = [
    "https://www.googleapis.com/auth/contacts.readonly",
    "https://www.googleapis.com/auth/chat.spaces.readonly",
    "https://www.googleapis.com/auth/chat",
    "https://www.googleapis.com/auth/chat.messages",
    "https://www.googleapis.com/auth/chat.messages.readonly",
    "https://www.googleapis.com/auth/chat.memberships",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.send",
]

SERVICE_FILES = ["sud_byjus_desk.json", "TeacherCommunity.json"]
TOKEN_FILES = ["tokenByj.json", "tc.json"]
HTML_TEMPLATE = "Template.html"
LOG_COLUMNS = ["Timestamp", "Name", "Email", "Attachment", "Subject", "Status", "Elapsed Time"]

def authenticate(token_file, client_secret):
    creds = None
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(client_secret, SCOPES)
            creds = flow.run_local_server(port=0, prompt="consent", authorization_prompt_message="")

        with open(token_file, "w") as token:
            token.write(creds.to_json())
    return creds

def load_sheet(url, title):
    gs = pg.authorize(custom_credentials=gkey)
    sheet = gs.open_by_url(url)
    return sheet.worksheet('title', title)

def send_email(gm, email, raw_message):
    try:
        sent_message = gm.users().messages().send(userId="me", body={'raw': raw_message}).execute()
        print(f"Email sent successfully!: {email}")
        return "Email sent successfully!"
    except Exception as error:
        print(f"Error sending email: {error}")
        return f"Error sending email: {error}"

def main():
    # Authenticate
    gkey = authenticate(TOKEN_FILES[1], SERVICE_FILES[1])
    gm = build('gmail', 'v1', credentials=gkey)

    # Load Sheets
    sheet = load_sheet("https://docs.google.com/spreadsheets/d/1ta-K8gMuNLctx18cq-JW9y6LfSV8J89dQeJTB9YYHcI/edit#gid=105697467", "HtmlMail")
    wksDF = sheet.get_as_df(start="A", end='H')

    # Read HTML template
    template_loader = FileSystemLoader(searchpath="./")
    env = Environment(loader=template_loader)
    template = env.get_template(HTML_TEMPLATE)

    # Initialize log DataFrame
    log_sheet = load_sheet("https://docs.google.com/spreadsheets/d/...", "Log")
    existing_logDF = log_sheet.get_as_df(start="A", end="G")
    logDF = existing_logDF if existing_logDF is not None else pd.DataFrame(columns=LOG_COLUMNS)

    # Send emails
    for index, row in wksDF.iterrows():
        if index < 1950:  # Limit for number of emails
            name = row["Name"]
            email = row["Email"]
            link = row["Attachment"]
            subject = row["Subject"]
            rendered_html = template.render(name=name, Link=link)

            message = MIMEMultipart()
            message['to'] = email
            message['cc'] = wksDF["CC"][0]
            message['subject'] = subject
            msg = MIMEText(rendered_html, 'html')
            message.attach(msg)
            raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')

            status = send_email(gm, email, raw_message)

            # Log email sending status
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            elapsed_time = time.time() - start_time
            log_entry = [timestamp, name, email, link, subject, status, elapsed_time]
            logDF = pd.concat([logDF, pd.DataFrame([log_entry], columns=LOG_COLUMNS)], ignore_index=True)
        else:
            print(f"Limit Reached at row number: {index+1, email}")
            break

    # Log information to the "Log" sheet
    log_sheet.clear()  # Clear existing data on the Log sheet
    log_sheet.set_dataframe(logDF, start=(1, 1))  # Write new log data

if __name__ == "__main__":
    start_time = time.time()
    main()
    print("This Window will Exit in 15 seconds...")
    time.sleep(15)
    print("Window will Exit Now")
    time.sleep(2)
