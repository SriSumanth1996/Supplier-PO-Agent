# gmail_reader_app.py

import os.path
import base64
import json
import streamlit as st
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Page config
st.set_page_config(page_title="Gmail Supplier Reader", layout="wide")
st.title("📧 Gmail Supplier Email Reader")
st.markdown("A Streamlit app to read supplier emails from Gmail.")

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'] 


def authenticate_gmail():
    """Authenticate with Gmail API using secrets stored in Streamlit."""
    from google.oauth2.credentials import Credentials

    CLIENT_ID = st.secrets["CLIENT_ID"]
    CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
    REFRESH_TOKEN = st.secrets["REFRESH_TOKEN"]
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'] 

    creds = Credentials(
        token=None,
        refresh_token=REFRESH_TOKEN,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET,
        token_uri="https://oauth2.googleapis.com/token", 
        scopes=SCOPES,
    )

    if creds.expired and creds.refresh_token:
        creds.refresh(Request())

    return build('gmail', 'v1', credentials=creds)


def get_email_body(msg_payload):
    """Extract the full body text from the email payload."""
    body = ""

    if 'parts' in msg_payload:
        for part in msg_payload['parts']:
            mime_type = part.get("mimeType")
            body_data = part.get("body", {}).get("data")
            if mime_type == "text/plain" and body_data:
                body = base64.urlsafe_b64decode(body_data).decode('utf-8')
                break
            elif mime_type == "text/html":
                import html2text
                html_body = base64.urlsafe_b64decode(body_data).decode('utf-8')
                body = html2text.html2text(html_body)
                break
            elif "parts" in part:
                body = get_email_body(part)
                if body:
                    break
    else:
        body_data = msg_payload.get("body", {}).get("data")
        if body_data:
            body = base64.urlsafe_b64decode(body_data).decode('utf-8')

    return body


def fetch_emails():
    try:
        service = authenticate_gmail()
        supplier_email = "supplier123.sample@gmail.com"

        with st.spinner("Fetching messages from Gmail..."):
            results = service.users().messages().list(userId='me', labelIds=['INBOX']).execute()
            messages = results.get('messages', [])

        if not messages:
            st.warning("No messages found in inbox.")
            return []

        output = []

        for message in messages[:5]:  # Check recent 5 emails
            msg = service.users().messages().get(userId='me', id=message['id']).execute()
            headers = msg['payload']['headers']
            subject = [h['value'] for h in headers if h['name'] == 'Subject']
            sender = [h['value'] for h in headers if h['name'] == 'From']
            date = [h['value'] for h in headers if h['name'] == 'Date']

            if not (subject and sender):
                continue

            subject = subject[0]
            sender = sender[0]

            if supplier_email in sender:
                body = get_email_body(msg['payload'])

                output.append({
                    "from": sender,
                    "subject": subject,
                    "received_at": date[0] if date else "Unknown",
                    "body": body
                })

        return output

    except Exception as e:
        st.error(f"Error fetching emails: {str(e)}")
        return []

# Main
if st.button("Fetch Latest Emails"):
    emails = fetch_emails()
    if emails:
        st.success(f"Found {len(emails)} email(s) from supplier(s).")

        # Convert list of dicts to DataFrame
        import pandas as pd
        df = pd.DataFrame(emails)

        # Optional: Clean up 'from' field to extract just email if needed
        df['email'] = df['from'].str.extract(r'<([^>]+)>')  # Extract email only

        # Display as a clean table
        st.dataframe(df)

        # Optional: Expand body for readability
        for idx, row in df.iterrows():
            with st.expander(f"Email {idx+1}: {row['subject']} from {row['from']}"):
                st.markdown(f"**Received at:** {row['received_at']}")
                st.markdown(f"**Body:**\n{row['body']}")

    else:
        st.info("No matching supplier emails found.")
