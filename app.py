import os.path
import base64
import json
import streamlit as st
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from groq import Groq
import pandas as pd

# Page config
st.set_page_config(page_title="Supplier Quote Parser", layout="wide")
st.title("📧 Supplier Email Quote Parser")
st.markdown("A Streamlit app that reads supplier emails and extracts structured quote data using Groq.")

# Groq Client Setup
GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
client = Groq(api_key=GROQ_API_KEY)

# Gmail Scopes
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']   

def authenticate_gmail():
    """Authenticate using secrets stored in Streamlit."""
    CLIENT_ID = st.secrets["CLIENT_ID"]
    CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
    REFRESH_TOKEN = st.secrets["REFRESH_TOKEN"]

    from google.oauth2.credentials import Credentials

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


qa_mapping = {
    "What is the product?": "product",
    "How many units?": "quantity",
    "What is the unit price?": "unit_price",
    "What is the total cost?": "total_cost",
    "What is the delivery time or method?": "delivery_method",
    "What is the Lead Time?": "lead_time",
    "What are the payment terms?": "payment_terms",
    "Is there any validity period for this quote?": "validity_period",
    "Where is the supplier located?": "supplier_place"  # NEW FIELD
}

def ask_groq(question, context):
    prompt = f"""
    Context: {context}

    Question: {question}

    Instructions:
    - Answer only based on the given context.
    - If the information is not present, respond with "Not present".
    - Do not explain, just return the answer directly.
    """

    try:
        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": prompt}],
            model="llama-3.3-70b-versatile",  # Or mixtral-8x7b-32768
            temperature=1,
            max_tokens=100,
            top_p=1.0,
            stop=None,
            stream=False,
        )
        answer = chat_completion.choices[0].message.content.strip()
        return answer if answer else "Not present"
    except Exception as e:
        return f"Error: {str(e)}"


def extract_quotation_data(context):
    response_json = {}
    for question, key in qa_mapping.items():
        answer = ask_groq(question, context)
        response_json[key] = answer
    return response_json


def main():
    if st.button("Fetch & Parse Supplier Emails"):
        with st.spinner("Authenticating with Gmail..."):
            service = authenticate_gmail()

        supplier_email = "supplier123.sample@gmail.com"

        with st.spinner("Fetching emails..."):
            results = service.users().messages().list(userId='me', labelIds=['INBOX']).execute()
            messages = results.get('messages', [])

        if not messages:
            st.warning("No messages found.")
            return

        df_columns = ["Received At", "Name", "Email", "Place", "Lead days", "Unit Cost", "Units", "Total Cost"]
        df_rows = []

        for message in messages[:5]:  # Check recent 5 emails
            msg = service.users().messages().get(userId='me', id=message['id']).execute()
            headers = msg['payload']['headers']
            sender = [h['value'] for h in headers if h['name'] == 'From'][0]

            if supplier_email in sender:
                received_at = [h['value'] for h in headers if h['name'] == 'Date'][0]
                subject = [h['value'] for h in headers if h['name'] == 'Subject'][0]
                body = get_email_body(msg['payload'])

                quotation_data = extract_quotation_data(body)

                name = sender.split("<")[0].strip() if "<" in sender else sender
                email_address = sender.split("<")[1][:-1] if "<" in sender else sender
                place = quotation_data.get("supplier_place", "Not present").strip("[]\"'") or "Unknown"
                lead_days = quotation_data.get("lead_time", "Not present").replace("–", "-") if quotation_data.get("lead_time") != "Not present" else "Not present"
                unit_cost = quotation_data.get("unit_price", "Not present").replace("$", "").strip()
                units = quotation_data.get("quantity", "Not present")
                total_cost = quotation_data.get("total_cost", "Not present").replace("$", "").strip()

                df_rows.append([
                    received_at,
                    name,
                    email_address,
                    place,
                    lead_days,
                    unit_cost,
                    units,
                    total_cost
                ])

        if df_rows:
            df = pd.DataFrame(df_rows, columns=df_columns)
            st.success("✅ Parsed supplier quotes:")
            st.dataframe(df)
        else:
            st.info("No matching supplier emails found.")
