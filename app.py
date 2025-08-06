from __future__ import print_function
import os.path
import base64
import email
import pandas as pd
from openai import OpenAI
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email.mime.text import MIMEText
import base64 as b64
from datetime import datetime, timedelta
import pytz
import re
import streamlit as st
from streamlit.logger import get_logger
from urllib.parse import urlparse, parse_qs
import json
import html2text
import uuid

logger = get_logger(__name__)

if 'OPENAI_API_KEY' in st.secrets:
    OPENAI_API_KEY = st.secrets['OPENAI_API_KEY']
else:
    try:
        with open("API.txt", "r") as f:
            OPENAI_API_KEY = f.read().strip()
    except FileNotFoundError:
        st.error("OpenAI API key not found. Please add it to secrets or API.txt file.")
        st.stop()

client = OpenAI(api_key=OPENAI_API_KEY)

SCOPES = [
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/calendar'
]

qa_mapping = {
    "What is the product or item being quoted?": "product",
    "How many units or quantity?": "quantity",
    "What is the unit price or price per piece?": "unit_price",
    "What is the total cost or total amount?": "total_cost",
    "What is the lead time or delivery time in days?": "lead_time",
    "What is the supplier's location, city, or place mentioned in the email signature?": "place",
    "What is the sender's personal name mentioned in the email signature?": "sender_name",
    "What is the company name mentioned in the email signature?": "company_name",
    "What is the contact phone number mentioned in the email signature?": "contact_number",
    "What is the sender's designation or job title mentioned in the email signature?": "designation"
}

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'gmail_service' not in st.session_state:
    st.session_state.gmail_service = None
if 'calendar_service' not in st.session_state:
    st.session_state.calendar_service = None
if 'processed_emails' not in st.session_state:
    st.session_state.processed_emails = []
if 'chat_messages' not in st.session_state:
    st.session_state.chat_messages = [
        {"role": "assistant",
         "content": "Ask me about supplier quotes or email details, e.g.:\n- 'Are there any partial quotes?'\n- 'Show details for SKF 6205 quotes'\n- 'What are the latest quotes under $10?'"}
    ]

def chatbot_response(prompt):
    st.sidebar.header("🔍 Supplier Quotation Assistant")
    for msg in st.session_state.chat_messages:
        st.sidebar.chat_message(msg["role"]).write(msg["content"])
    if prompt:
        st.session_state.chat_messages.append({"role": "user", "content": prompt})
        st.sidebar.chat_message("user").write(prompt)
        with st.sidebar.chat_message("assistant"):
            with st.spinner("Analyzing..."):
                try:
                    response = generate_response(prompt, st.session_state.processed_emails)
                    st.write(response)
                    st.session_state.chat_messages.append({"role": "assistant", "content": response})
                except Exception as e:
                    response = f"Error: {str(e)}"
                    st.write(response)
                    st.session_state.chat_messages.append({"role": "assistant", "content": response})

def generate_response(query, processed_emails):
    email_context = []
    for email in processed_emails:
        qd = email['quotation_data']
        email_summary = {
            "email_address": email['email_address'],
            "subject": email['subject'],
            "classification": email['final_classification'],
            "product": qd.get('product', 'Not present'),
            "quantity": qd.get('quantity', 'Not present'),
            "unit_price": qd.get('unit_price', 'Not present'),
            "total_cost": qd.get('total_cost', 'Not present'),
            "lead_time": qd.get('lead_time', 'Not present'),
            "sender_name": qd.get('sender_name', 'Not present'),
            "company_name": qd.get('company_name', 'Not present'),
            "place": qd.get('place', 'Not present'),
            "contact_number": qd.get('contact_number', 'Not present'),
            "designation": qd.get('designation', 'Not present'),
            "meeting_status": get_meeting_status(email.get('meeting_details'), email.get('meeting_result'))
        }
        email_context.append(email_summary)
    prompt = f"""
    You are a supplier quotation assistant. Answer the user's query concisely and precisely based on the processed email data.
    USER QUERY: "{query}"
    PROCESSED EMAILS: {json.dumps(email_context, indent=2)}
    GUIDELINES:
    - Provide a direct, concise response (max 100 words).
    - Focus on the query's intent (e.g., partial quotes, specific products, price ranges).
    - If the query asks about partial quotes, list emails classified as "Quotation Partially Received" with missing fields.
    - If specific products or price ranges are mentioned, summarize relevant email data.
    - If no relevant data is found, state so clearly.
    - Do not use a product catalog; rely only on email data.
    - Avoid technical jargon unless requested.
    RESPONSE:
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=150
        )
        return response.choices[0].message.content.strip() or "No relevant information found."
    except Exception as e:
        return f"Error processing query: {str(e)}"

def authenticate_gmail_and_calendar():
    creds = None
    if 'REFRESH_TOKEN' in st.secrets:
        creds = Credentials(
            token=None,
            refresh_token=st.secrets['REFRESH_TOKEN'],
            token_uri='https://oauth2.googleapis.com/token',
            client_id=st.secrets['CLIENT_ID'],
            client_secret=st.secrets['CLIENT_SECRET'],
            scopes=SCOPES
        )
        try:
            creds.refresh(Request())
            gmail_service = build('gmail', 'v1', credentials=creds)
            calendar_service = build('calendar', 'v3', credentials=creds)
            return gmail_service, calendar_service
        except Exception as e:
            st.warning(f"Could not refresh token: {e}")
    flow = InstalledAppFlow.from_client_config(
        {
            "installed": {
                "client_id": st.secrets['CLIENT_ID'],
                "client_secret": st.secrets['CLIENT_SECRET'],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": ["http://localhost:8501", "http://localhost"]
            }
        },
        SCOPES
    )
    if not st.runtime.exists():
        auth_url, _ = flow.authorization_url(prompt='consent')
        st.markdown(f"""
        1. Click [this link]({auth_url}) to authorize
        2. You'll be redirected to a localhost URL that won't work
        3. Copy the entire URL from your browser's address bar
        4. Paste it below
        """)
        code_url = st.text_input("Paste the redirect URL here:")
        if code_url:
            try:
                code = parse_qs(urlparse(code_url).query)['code'][0]
                flow.fetch_token(code=code)
                creds = flow.credentials
                st.session_state['temp_creds'] = {
                    'token': creds.token,
                    'refresh_token': creds.refresh_token,
                    'token_uri': creds.token_uri,
                    'client_id': creds.client_id,
                    'client_secret': creds.client_secret,
                    'scopes': creds.scopes
                }
                gmail_service = build('gmail', 'v1', credentials=creds)
                calendar_service = build('calendar', 'v3', credentials=creds)
                return gmail_service, calendar_service
            except Exception as e:
                st.error(f"Authentication failed: {e}")
                return None, None
        return None, None
    try:
        creds = flow.run_local_server(port=0)
        gmail_service = build('gmail', 'v1', credentials=creds)
        calendar_service = build('calendar', 'v3', credentials=creds)
        return gmail_service, calendar_service
    except Exception as e:
        st.error(f"Authentication failed: {e}")
        return None, None

def get_email_body(msg_payload):
    body = ""
    if 'parts' in msg_payload:
        for part in msg_payload['parts']:
            mime_type = part.get("mimeType")
            body_data = part.get("body", {}).get("data")
            if mime_type == "text/plain" and body_data:
                body = base64.urlsafe_b64decode(body_data).decode('utf-8')
                break
            elif mime_type == "text/html" and body_data:
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

def ask_openai(question, context):
    prompt = f"""
    You are a specialized Purchase Order (PO) and supplier quotation data extraction assistant. Your task is to analyze business emails from suppliers and extract specific information accurately.
    EMAIL CONTENT TO ANALYZE:
    {context}
    QUESTION: {question}
    EXTRACTION GUIDELINES:
    1. PRODUCTS/ITEMS:
       - Look for product names, item descriptions, part numbers, SKUs, model numbers
       - Include brand names if mentioned
       - Be specific about the product (e.g., "SKF Deep Groove Ball Bearing 6205-2RS" not just "bearing")
    2. QUANTITIES/UNITS:
       - Extract numerical quantities (e.g., "5 units", "100 pieces", "2 boxes")
       - Return just the number with unit type (e.g., "50 pieces")
    3. UNIT PRICE:
       - Extract ONLY the price per single unit/item as a clean number with currency
       - Look for currency symbols ($, €, ₹, etc.) and amounts
       - Return ONLY the price (e.g., "₹180" or "$25.50")
       - Look for terms like: "unit price", "per piece", "each", "price per unit"
       - If you need to calculate unit price from total and quantity, do the math and return only the result
       - NEVER return explanatory text, only the price value
    4. TOTAL COST:
       - Extract final/total amount for the entire order
       - Look for terms like: "total", "total amount", "grand total", "final cost", "total price"
       - Include currency symbol (e.g., "$1275.00", "₹60000")
       - Return ONLY the amount with currency, no descriptive text
    5. LEAD TIME:
       - Extract time frames for delivery or production in DAYS
       - Look for: "delivery time", "lead time", "shipping time", "ready in", "available in"
       - Convert to days if given in weeks/months (e.g., "2 weeks" = "14 days")
       - Return format e.g., "7 days", "10-15 days", "21 days"
    6. SUPPLIER PLACE/LOCATION:
       - Look for supplier's city, state, or location in email signature (e.g., "Mumbai", "Kolkata"). Return only that value.
    7. SENDER NAME:
       - Extract personal name from email signature (e.g., "Rakshan"). Return only that value.
    8. COMPANY NAME:
       - Extract company or organization name from signature (e.g., "TamilNadu Bearing Industries Ltd."). Return only that value.
    9. CONTACT NUMBER:
       - Extract phone number (e.g., "+91 7661598752")
    10. DESIGNATION:
       - Extract sender's designation or job title (e.g., "Sales Manager")
    RESPONSE RULES:
    - Extract ONLY explicitly stated info.
    - These responses are fed into a tabular format. So, just return what is asked for.
    - Do not hallucinate or take from the examples given above in this prompt
    - If not found, respond with "Not present"
    - Keep original format for non-price fields
    - Don't assume or guess, only state exact extracted values
    ANSWER:
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=150
        )
        return response.choices[0].message.content.strip() or "Not present"
    except Exception as e:
        return f"Error: {str(e)}"

def classify_email_intent(context):
    prompt = f"""
    You are an email classification assistant specialized in analyzing supplier/business emails.
    EMAIL CONTENT TO ANALYZE:
    {context}
    Your task is to classify this email into ONE of these three categories:
    1. "New Business Connection" - If the email is:
       - Introduction email from new supplier
       - Company introduction or capability presentation
       - General business development outreach
       - Marketing or promotional content
       - Request for partnership or collaboration
    2. "Quotation Received" - If the email contains ALL four essential elements:
       - Product details
       - Clear pricing information (unit price or total cost)
       - Specific quantities mentioned
       - Lead time or delivery information
    3. "Quotation Partially Received" - If the email contains quotation elements but is missing ANY of these:
       - Product details OR
       - Unit price OR
       - Quantities OR
       - Lead time
    CLASSIFICATION RULES:
    - Focus on the PRIMARY intent of the email
    - Be specific and choose only ONE category
    - Meeting requests should be detected separately (not as a classification)
    - In "Quotation Partially Received", return "Not Present" for the elements that are missing
    RESPOND WITH ONLY THE CLASSIFICATION CATEGORY NAME (exactly as written above):
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=50
        )
        classification = response.choices[0].message.content.strip()
        valid_classifications = ["Quotation Received", "Quotation Partially Received", "New Business Connection"]
        if classification not in valid_classifications:
            return "Unknown"
        return classification
    except Exception as e:
        return "Unknown"

def extract_meeting_details(context):
    ist = pytz.timezone('Asia/Kolkata')
    now_ist = datetime.now(ist)
    current_ist_iso = now_ist.isoformat()
    prompt = f"""
    You are a meeting scheduling assistant. Read the email below and extract a clear, structured datetime for a proposed meeting.
    Email Content:
    {context}
    Current datetime (IST): {current_ist_iso}
    Task:
    1. Determine if the sender intends to set up a meeting. Reply with "Yes" or "No".
    2. If yes, infer the proposed date and time, even if partial (e.g., "6th at 4PM", "Monday morning").
       - Convert into a full datetime string in ISO 8601 format (e.g., "2025-08-06T16:00:00+05:30").
       - Assume Indian Standard Time (IST).
       - If the date (e.g., "6th") has already passed this month, infer the next month's same date.
       - Use the provided current datetime to resolve relative references.
    3. If no date or time is found, respond with "Not specified".
    Output format:
    Meeting Intent: Yes/No
    Proposed Datetime: <ISO 8601 timestamp> or Not specified
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=150
        )
        reply = response.choices[0].message.content.strip()
        meeting_intent, proposed_datetime = "No", "Not specified"
        for line in reply.splitlines():
            if line.startswith("Meeting Intent:"):
                meeting_intent = line.split(":", 1)[1].strip()
            elif line.startswith("Proposed Datetime:"):
                proposed_datetime = line.split(":", 1)[1].strip()
        return {
            "meeting_intent": meeting_intent,
            "proposed_datetime": proposed_datetime
        }
    except Exception as e:
        return {
            "meeting_intent": "No",
            "proposed_datetime": "Not specified"
        }

def extract_quotation_data(context, classification):
    if classification in ["New Business Connection", "Unknown"]:
        business_qa_mapping = {
            "What is the supplier's location, city, or place mentioned in the email signature?": "place",
            "What is the sender's personal name mentioned in the email signature?": "sender_name",
            "What is the company name mentioned in the email signature?": "company_name",
            "What is the contact phone number mentioned in the email signature?": "contact_number",
            "What is the sender's designation or job title mentioned in the email signature?": "designation"
        }
        return {key: ask_openai(question, context) for question, key in business_qa_mapping.items()}
    else:
        return {key: ask_openai(question, context) for question, key in qa_mapping.items()}

def get_final_classification(quotation_data, initial_classification):
    if initial_classification in ["New Business Connection", "Unknown"]:
        return initial_classification
    required_fields = ["product", "quantity", "unit_price", "lead_time"]
    missing_fields = [field for field in required_fields if quotation_data.get(field, "Not present") == "Not present"]
    if missing_fields:
        return "Quotation Partially Received"
    return "Quotation Received"

def calculate_unit_price_if_missing(quotation_data):
    if quotation_data.get("quantity", "Not present") == "Not present":
        return quotation_data
    unit_cost = quotation_data.get("unit_price", "Not present")
    total_cost = quotation_data.get("total_cost", "Not present")
    units = quotation_data.get("quantity", "Not present")
    if unit_cost == "Not present" and total_cost != "Not present":
        try:
            total_num = float(''.join([c for c in total_cost if c.isdigit() or c == '.']))
            quantity_num = float(''.join([c for c in units if c.isdigit() or c == '.']))
            currency_symbol = ''.join([c for c in total_cost if not (c.isdigit() or c == '.' or c.isspace())])
            if quantity_num > 0:
                unit_price_calculated = total_num / quantity_num
                quotation_data["unit_price"] = f"{currency_symbol}{unit_price_calculated:.2f}"
        except Exception:
            pass
    return quotation_data

def calculate_total_cost_if_missing(quotation_data):
    if quotation_data.get("quantity", "Not present") == "Not present":
        return quotation_data
    unit_cost = quotation_data.get("unit_price", "Not present")
    total_cost = quotation_data.get("total_cost", "Not present")
    units = quotation_data.get("quantity", "Not present")
    if total_cost == "Not present" and unit_cost != "Not present":
        try:
            unit_price_num = float(''.join([c for c in unit_cost if c.isdigit() or c == '.']))
            quantity_num = float(''.join([c for c in units if c.isdigit() or c == '.']))
            currency_symbol = ''.join([c for c in unit_cost if not (c.isdigit() or c == '.' or c.isspace())])
            total_cost = f"{currency_symbol}{unit_price_num * quantity_num:.2f}"
            quotation_data["total_cost"] = total_cost
        except Exception:
            pass
    return quotation_data

def send_reply(service, thread_id, to_email, subject, body):
    message = MIMEText(body)
    message['to'] = to_email
    message['subject'] = f"Re: {subject}"
    raw = b64.urlsafe_b64encode(message.as_bytes()).decode()
    message = {
        'raw': raw,
        'threadId': thread_id
    }
    try:
        sent = service.users().messages().send(userId="me", body=message).execute()
        return True, f"Reply sent to {to_email}"
    except Exception as e:
        return False, f"Error sending reply: {e}"

def check_calendar_conflict(calendar_service, start_time, end_time):
    try:
        events_result = calendar_service.events().list(
            calendarId='primary',
            timeMin=start_time.isoformat(),
            timeMax=end_time.isoformat(),
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        events = events_result.get('items', [])
        for event in events:
            event_start = datetime.fromisoformat(
                event['start'].get('dateTime', event['start'].get('date')).replace('Z', '+00:00'))
            event_end = datetime.fromisoformat(
                event['end'].get('dateTime', event['end'].get('date')).replace('Z', '+00:00'))
            if not (end_time <= event_start or start_time >= event_end):
                return True, event.get('summary', 'Existing meeting')
        return False, None
    except Exception as e:
        print(f"Error checking calendar conflict: {e}")
        return False, None

def schedule_meeting(calendar_service, quotation_data, email_address, proposed_datetime=None, classification="Unknown"):
    try:
        ist = pytz.timezone('Asia/Kolkata')
        now = datetime.now(ist)
        if proposed_datetime:
            start_hour = proposed_datetime.hour
            if start_hour < 9 or start_hour >= 17:
                return None, "outside_business_hours"
            start_time = proposed_datetime
            end_time = start_time + timedelta(minutes=30)
            if start_time < now:
                return None, "past_time"
            has_conflict, conflicting_event = check_calendar_conflict(calendar_service, start_time, end_time)
            if has_conflict:
                return None, "conflict"
        else:
            return None, "no_specific_time"
        company_name = quotation_data.get('company_name', 'Unknown Company')
        if company_name == "Not present":
            company_name = "Unknown Company"
        if classification == "New Business Connection":
            description = f"""
            Meeting to discuss potential business collaboration with {company_name}.
            Contact: {quotation_data.get('sender_name', 'Not present')} ({quotation_data.get('contact_number', 'Not present')})
            Designation: {quotation_data.get('designation', 'Not present')}
            Location: {quotation_data.get('place', 'Not present')}
            Email: {email_address}
            """
        else:
            description = f"""
            Meeting to discuss quotation for {quotation_data.get('product', 'Unknown Product')}.
            Quantity: {quotation_data.get('quantity', 'Not present')}
            Unit Price: {quotation_data.get('unit_price', 'Not present')}
            Total Cost: {quotation_data.get('total_cost', 'Not present')}
            Lead Time: {quotation_data.get('lead_time', 'Not present')}
            Contact: {quotation_data.get('sender_name', 'Not present')} ({quotation_data.get('contact_number', 'Not present')})
            Location: {quotation_data.get('place', 'Not present')}
            Email: {email_address}
            """
        event = {
            'summary': f"Supplier Meeting: {company_name}",
            'description': description,
            'start': {
                'dateTime': start_time.isoformat(),
                'timeZone': 'Asia/Kolkata',
            },
            'end': {
                'dateTime': end_time.isoformat(),
                'timeZone': 'Asia/Kolkata',
            },
            'attendees': [
                {'email': email_address},
                {'email': 'srivenkatasumanthpisapati@gmail.com'}
            ],
            'reminders': {
                'useDefault': False,
                'overrides': [
                    {'method': 'email', 'minutes': 24 * 60},
                    {'method': 'popup', 'minutes': 10},
                ],
            },
        }
        event = calendar_service.events().insert(calendarId='primary', body=event, sendUpdates='all').execute()
        return event, "scheduled"
    except Exception as e:
        print(f"Error scheduling meeting: {e}")
        return None, "error"

def parse_new_datetime(instructions):
    ist = pytz.timezone('Asia/Kolkata')
    now = datetime.now(ist)
    prompt = f"""
    You are a datetime parsing assistant. Extract a proposed meeting datetime from the following instructions.
    Instructions: {instructions}
    Current datetime (IST): {now.isoformat()}
    Task:
    - Extract a clear datetime in ISO 8601 format (e.g., "2025-08-06T16:00:00+05:30") from the instructions.
    - Assume Indian Standard Time (IST).
    - If the date has passed this month, infer the next month.
    - If no valid datetime is found, return "Not specified".
    Output: <ISO 8601 timestamp> or "Not specified"
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=50
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return "Not specified"

def get_meeting_date_time(meeting_details):
    """Extract date and time separately from meeting details"""
    if not meeting_details or meeting_details.get("meeting_intent") != "Yes":
        return "Not Requested", "Not Requested"
    if meeting_details.get("proposed_datetime") == "Not specified":
        return "Not Specified", "Not Specified"
    try:
        dt = datetime.fromisoformat(meeting_details["proposed_datetime"])
        date_str = dt.strftime("%Y-%m-%d")
        time_str = dt.strftime("%H:%M")
        return date_str, time_str
    except:
        return "Not Specified", "Not Specified"

def get_reply_body(classification, quotation_data, sender_name, meeting_details=None, meeting_result=None, instructions=""):
    ist = pytz.timezone('Asia/Kolkata')
    # Base message based on classification
    if classification == "Quotation Received":
        base_message = f"""Dear {sender_name or 'Supplier'},
Thank you for your quotation. We have received the full details regarding your product and pricing."""
    elif classification == "Quotation Partially Received":
        missing_items = []
        if quotation_data.get("product", "Not present") == "Not present":
            missing_items.append("product details")
        if quotation_data.get("unit_price", "Not present") == "Not present":
            missing_items.append("unit price")
        if quotation_data.get("quantity", "Not present") == "Not present":
            missing_items.append("quantity")
        if quotation_data.get("lead_time", "Not present") == "Not present":
            missing_items.append("lead time")
        base_message = f"""Dear {sender_name or 'Supplier'},
Thank you for your quotation. We have reviewed the information provided, however, we need additional details to complete our evaluation."""
        if missing_items:
            base_message += f"\n\nPlease provide the following missing information:"
            for i, item in enumerate(missing_items, 1):
                base_message += f"\n{i}. {item.title()}"
    elif classification == "New Business Connection":
        base_message = f"""Dear {sender_name or 'Supplier'},
Thank you for introducing your company and sharing your offerings with us."""
    else:
        base_message = f"""Dear {sender_name or 'Supplier'},
Thank you for your email."""
    # Process instructions if provided
    if instructions.strip():
        try:
            # Use GPT to interpret instructions and generate appropriate meeting text
            prompt = f"""
            You are a professional email assistant. Based on the following context and instructions, generate appropriate meeting-related text for a business email.
            Email Classification: {classification}
            Original Meeting Details: {meeting_details}
            Instructions from User: "{instructions}"
            Guidelines:
            1. If instructions contain a new meeting time:
               - Politely explain we can't meet at original time (if applicable)
               - Propose the new time
               - Mention calendar invite will be sent
            2. If instructions request someone to join:
               - Politely request their presence
               - Explain why if reason is given
            3. For other instructions:
               - Incorporate naturally into the email
            4. Keep tone professional and polite
            Respond ONLY with the text to be inserted in the email (no headings or markers).
            """
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=200
            )
            meeting_text = "\n\n" + response.choices[0].message.content.strip()
        except Exception as e:
            meeting_text = f"\n\nAdditional Instructions: {instructions}"
    else:
        meeting_text = ""
    base_message += meeting_text
    # Standard closing based on classification
    if classification == "Quotation Received":
        base_message += "\n\nLooking forward to your response.\n\nBest regards,\nDr. Saravanan Kesavan\nBITSoM"
    elif classification == "Quotation Partially Received":
        base_message += "\n\nOnce we receive the complete information, we will be able to proceed with our evaluation.\n\nThank you for your cooperation.\n\nBest regards,\nDr. Saravanan Kesavan\nBITSoM"
    elif classification == "New Business Connection":
        base_message += "\n\nWe will keep your information on record and reach out to you when opportunities arise.\n\nWarm regards,\nDr. Saravanan Kesavan\nBITSoM"
    else:
        base_message += "\n\nBest regards,\nDr. Saravanan Kesavan\nBITSoM"
    return base_message

def get_meeting_status(meeting_details, meeting_result):
    if not meeting_details or meeting_details.get("meeting_intent") != "Yes":
        return "No Meeting Requested"
    if meeting_result:
        status = meeting_result[1]
        if status == "scheduled":
            return "Meeting Scheduled"
        elif status == "conflict":
            return "Schedule Conflict"
        elif status == "outside_business_hours":
            return "Outside Business Hours"
        elif status == "past_time":
            return "Time Already Passed"
        elif status == "no_specific_time":
            return "No Specific Time"
        elif status == "incomplete_details":
            return "Incomplete Details"
        elif status == "parse_error":
            return "Time Parse Error"
        else:
            return "Error Occurred"
    else:
        return "Meeting Requested"

def create_quotation_received_table(emails):
    if not emails:
        return pd.DataFrame()
    data = []
    for email in emails:
        qd = email['quotation_data']
        meeting_status = get_meeting_status(email.get('meeting_details'), email.get('meeting_result'))
        meeting_date, meeting_time = get_meeting_date_time(email.get('meeting_details'))
        data.append({
            'Sender Name': qd.get('sender_name', 'Not present'),
            'Company': qd.get('company_name', 'Not present'),
            'Email': email['email_address'],
            'Product': qd.get('product', 'Not present').split(':')[-1].strip() if ':' in qd.get('product', 'Not present') else qd.get('product', 'Not present'),
            'Quantity': qd.get('quantity', 'Not present'),
            'Unit Price': qd.get('unit_price', 'Not present'),
            'Total Cost': qd.get('total_cost', 'Not present'),
            'Lead Time': qd.get('lead_time', 'Not present'),
            'Location': qd.get('place', 'Not present'),
            'Contact': qd.get('contact_number', 'Not present'),
            'Meeting Status': meeting_status,
            'Date of Meeting': meeting_date,
            'Time of Meeting': meeting_time,
            'Instructions': '',
            'Send': False
        })
    df = pd.DataFrame(data)
    return df

def create_quotation_partial_table(emails):
    if not emails:
        return pd.DataFrame()
    data = []
    for email in emails:
        qd = email['quotation_data']
        meeting_status = get_meeting_status(email.get('meeting_details'), email.get('meeting_result'))
        meeting_date, meeting_time = get_meeting_date_time(email.get('meeting_details'))
        missing_fields = []
        if qd.get('product', 'Not present') == 'Not present':
            missing_fields.append('Product')
        if qd.get('quantity', 'Not present') == 'Not present':
            missing_fields.append('Quantity')
        if qd.get('unit_price', 'Not present') == 'Not present':
            missing_fields.append('Unit Price')
        if qd.get('lead_time', 'Not present') == 'Not present':
            missing_fields.append('Lead Time')
        data.append({
            'Sender Name': qd.get('sender_name', 'Not present'),
            'Company': qd.get('company_name', 'Not present'),
            'Email': email['email_address'],
            'Product': qd.get('product', 'Not present').split(':')[-1].strip() if ':' in qd.get('product', 'Not present') else qd.get('product', 'Not present'),
            'Quantity': qd.get('quantity', 'Not present'),
            'Unit Price': qd.get('unit_price', 'Not present'),
            'Total Cost': qd.get('total_cost', 'Not present'),
            'Lead Time': qd.get('lead_time', 'Not present'),
            'Location': qd.get('place', 'Not present'),
            'Contact': qd.get('contact_number', 'Not present'),
            'Missing Fields': ', '.join(missing_fields) if missing_fields else 'None',
            'Meeting Status': meeting_status,
            'Date of Meeting': meeting_date,
            'Time of Meeting': meeting_time,
            'Instructions': '',
            'Send': False
        })
    df = pd.DataFrame(data)
    return df

def create_business_connection_table(emails):
    if not emails:
        return pd.DataFrame()
    data = []
    for email in emails:
        qd = email['quotation_data']
        meeting_status = get_meeting_status(email.get('meeting_details'), email.get('meeting_result'))
        meeting_date, meeting_time = get_meeting_date_time(email.get('meeting_details'))
        data.append({
            'Sender Name': qd.get('sender_name', 'Not present'),
            'Company': qd.get('company_name', 'Not present'),
            'Email': email['email_address'],
            'Designation': qd.get('designation', 'Not present'),
            'Location': qd.get('place', 'Not present'),
            'Contact': qd.get('contact_number', 'Not present'),
            'Meeting Status': meeting_status,
            'Date of Meeting': meeting_date,
            'Time of Meeting': meeting_time,
            'Instructions': '',
            'Send': False
        })
    df = pd.DataFrame(data)
    return df

def send_replies_for_emails(service, calendar_service, emails, df):
    success_count = 0
    error_count = 0
    selected_emails = [(email, row) for email, row in zip(emails, df.itertuples(index=False)) if getattr(row, 'Send')]
    if not selected_emails:
        st.warning("No emails selected to send replies.")
        return
    progress_bar = st.progress(0)
    status_text = st.empty()
    for i, (email_data, row) in enumerate(selected_emails):
        progress = (i + 1) / len(selected_emails)
        progress_bar.progress(progress)
        status_text.text(f'Sending reply {i + 1} of {len(selected_emails)}...')
        instructions = getattr(row, 'Instructions')
        meeting_details = email_data.get('meeting_details', {})
        # Process meeting scheduling if instructions contain time information
        if instructions.strip() and meeting_details.get('meeting_intent') == "Yes":
            try:
                # Extract datetime from instructions
                ist = pytz.timezone('Asia/Kolkata')
                current_time_ist = datetime.now(ist).isoformat()
                prompt = f"""
                Extract a clear meeting datetime from these instructions. If found, convert to ISO 8601 format in IST.
                Instructions: "{instructions}"
                Current Time (IST): {current_time_ist}
                Respond ONLY with the ISO timestamp or "Not specified" if no time found.
                """
                response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.3,
                    max_tokens=50
                )
                proposed_dt_str = response.choices[0].message.content.strip()
                if proposed_dt_str != "Not specified":
                    proposed_dt = datetime.fromisoformat(proposed_dt_str)
                    event, status = schedule_meeting(
                        calendar_service,
                        email_data['quotation_data'],
                        email_data['email_address'],
                        proposed_dt,
                        email_data['final_classification']
                    )
                    email_data['meeting_result'] = (event, status)
            except Exception as e:
                email_data['meeting_result'] = (None, "parse_error")
        # Generate reply body
        reply_body = get_reply_body(
            email_data['final_classification'],
            email_data['quotation_data'],
            email_data['quotation_data'].get('sender_name'),
            meeting_details,
            email_data.get('meeting_result'),
            instructions
        )
        # Send the reply
        success, message = send_reply(
            service,
            email_data['thread_id'],
            email_data['email_address'],
            email_data['subject'],
            reply_body
        )
        if success:
            success_count += 1
        else:
            error_count += 1
    progress_bar.progress(1.0)
    status_text.text('Bulk reply complete!')
    if success_count > 0:
        st.success(f"Successfully sent {success_count} replies!")
    if error_count > 0:
        st.error(f"Failed to send {error_count} replies.")

def display_classification_tables(processed_emails):
    if not processed_emails:
        st.warning("No emails processed yet.")
        return
    quotation_received = [e for e in processed_emails if e['final_classification'] == 'Quotation Received']
    quotation_partial = [e for e in processed_emails if e['final_classification'] == 'Quotation Partially Received']
    business_connection = [e for e in processed_emails if e['final_classification'] == 'New Business Connection']
    unknown = [e for e in processed_emails if e['final_classification'] == 'Unknown']
    tabs = st.sidebar.radio("Select View", ["Quotations", "New Business Connections"])
    if tabs == "Quotations":
        st.header("Complete Quotations Received")
        if quotation_received:
            df_complete = create_quotation_received_table(quotation_received)
            edited_df_complete = st.data_editor(
                df_complete,
                column_config={
                    "Send": st.column_config.CheckboxColumn("Send", default=False),
                    "Instructions": st.column_config.TextColumn(
                        "Instructions",
                        help="Specify any additional instructions for the reply."
                    )
                },
                use_container_width=True,
                num_rows="dynamic",
                hide_index=True
            )
            csv_complete = edited_df_complete.to_csv(index=False)
            st.download_button(
                label="Download Complete Quotations CSV",
                data=csv_complete,
                file_name=f"complete_quotations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
            if st.button("Send Replies for Selected Complete Quotations"):
                send_replies_for_emails(st.session_state.gmail_service, st.session_state.calendar_service,
                                        quotation_received, edited_df_complete)
        else:
            st.info("No complete quotations found in the processed emails.")
        st.header("Partial Quotations Received")
        if quotation_partial:
            df_partial = create_quotation_partial_table(quotation_partial)
            edited_df_partial = st.data_editor(
                df_partial,
                column_config={
                    "Send": st.column_config.CheckboxColumn("Send", default=False),
                    "Instructions": st.column_config.TextColumn(
                        "Instructions",
                        help="Specify any additional instructions for the reply."
                    )
                },
                use_container_width=True,
                num_rows="dynamic",
                hide_index=True
            )
            csv_partial = edited_df_partial.to_csv(index=False)
            st.download_button(
                label="Download Partial Quotations CSV",
                data=csv_partial,
                file_name=f"partial_quotations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
            if st.button("Send Replies for Selected Partial Quotations"):
                send_replies_for_emails(st.session_state.gmail_service, st.session_state.calendar_service,
                                        quotation_partial, edited_df_partial)
        else:
            st.info("No partial quotations found in the processed emails.")
    elif tabs == "New Business Connections":
        st.header("New Business Connections")
        if business_connection:
            df_business = create_business_connection_table(business_connection)
            edited_df_business = st.data_editor(
                df_business,
                column_config={
                    "Send": st.column_config.CheckboxColumn("Send", default=False),
                    "Instructions": st.column_config.TextColumn(
                        "Instructions",
                        help="Specify any additional instructions for the reply."
                    )
                },
                use_container_width=True,
                num_rows="dynamic",
                hide_index=True
            )
            csv_business = edited_df_business.to_csv(index=False)
            st.download_button(
                label="Download Business Connections CSV",
                data=csv_business,
                file_name=f"business_connections_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
            if st.button("Send Replies for Selected Business Connections"):
                send_replies_for_emails(st.session_state.gmail_service, st.session_state.calendar_service,
                                        business_connection, edited_df_business)
        else:
            st.info("No new business connection emails found in the processed emails.")
    if unknown:
        st.header("Unknown/Other Classifications")
        st.warning(f"Found {len(unknown)} emails that could not be properly classified:")
        for email in unknown:
            st.write(f"- {email['email_address']}: {email['subject']}")

def process_emails(gmail_service, calendar_service, num_emails=5):
    results = gmail_service.users().messages().list(
        userId='me',
        q='category:primary',
        labelIds=['INBOX']
    ).execute()
    messages = results.get('messages', [])
    if not messages:
        st.warning("No messages found in inbox.")
        return
    processed_emails = []
    progress_bar = st.progress(0)
    status_text = st.empty()
    for i, message in enumerate(messages[:num_emails]):
        progress = (i + 1) / num_emails
        progress_bar.progress(progress)
        status_text.text(f'Processing email {i + 1} of {num_emails}...')
        msg = gmail_service.users().messages().get(userId='me', id=message['id']).execute()
        headers = msg['payload']['headers']
        sender = [h['value'] for h in headers if h['name'] == 'From'][0]
        subject = [h['value'] for h in headers if h['name'] == 'Subject'][0]
        thread_id = msg['threadId']
        body = get_email_body(msg['payload'])
        meeting_details = extract_meeting_details(body)
        initial_classification = classify_email_intent(body)
        quotation_data = extract_quotation_data(body, initial_classification)
        if initial_classification not in ["New Business Connection", "Unknown"]:
            quotation_data = calculate_unit_price_if_missing(quotation_data)
            quotation_data = calculate_total_cost_if_missing(quotation_data)
        final_classification = get_final_classification(quotation_data, initial_classification)
        name = sender.split("<")[0].strip() if "<" in sender else sender
        email_address = sender.split("<")[1][:-1] if "<" in sender else sender
        reply_body = get_reply_body(
            final_classification,
            quotation_data,
            quotation_data.get("sender_name"),
            meeting_details,
            None
        )
        processed_emails.append({
            "email_address": email_address,
            "subject": subject,
            "final_classification": final_classification,
            "quotation_data": quotation_data,
            "meeting_details": meeting_details,
            "meeting_result": None,
            "reply_body": reply_body,
            "thread_id": thread_id
        })
    progress_bar.progress(1.0)
    status_text.text('Processing complete!')
    return processed_emails

def main():
    st.set_page_config(page_title="Supplier Quotation Processor", layout="wide")
    st.title("Supplier Quotation Processing System")
    st.markdown(
        "This application processes supplier emails, extracts quotation details, classifies them, and responds to queries concisely.")
    st.sidebar.header("Authentication")
    if not st.session_state.authenticated:
        if st.sidebar.button("Authenticate with Google"):
            with st.spinner("Authenticating..."):
                try:
                    gmail_service, calendar_service = authenticate_gmail_and_calendar()
                    if gmail_service and calendar_service:
                        st.session_state.gmail_service = gmail_service
                        st.session_state.calendar_service = calendar_service
                        st.session_state.authenticated = True
                        st.sidebar.success("Authentication successful!")
                        st.rerun()
                    else:
                        st.sidebar.error("Authentication failed. Please try again.")
                except Exception as e:
                    st.sidebar.error(f"Authentication error: {str(e)}")
    else:
        st.sidebar.success("Authenticated with Google")
        if st.sidebar.button("Logout"):
            st.session_state.authenticated = False
            st.session_state.gmail_service = None
            st.session_state.calendar_service = None
            st.session_state.processed_emails = []
            st.session_state.chat_messages = []
            st.rerun()
    prompt = st.sidebar.chat_input("Ask about supplier quotes or email details...")
    chatbot_response(prompt)
    if not st.session_state.authenticated:
        st.warning("Please authenticate with Google to continue.")
        return
    st.header("Process Emails")
    col1, col2 = st.columns([2, 1])
    with col1:
        num_emails = st.slider("Number of emails to process", 1, 20, 5)
    with col2:
        st.write("")
        process_button = st.button("Process Latest Emails", type="primary")
    if process_button:
        with st.spinner("Processing emails..."):
            try:
                st.session_state.processed_emails = process_emails(
                    st.session_state.gmail_service,
                    st.session_state.calendar_service,
                    num_emails
                )
                st.success(f"Successfully processed {len(st.session_state.processed_emails)} emails!")
            except Exception as e:
                st.error(f"Error processing emails: {str(e)}")
    if st.session_state.processed_emails:
        display_classification_tables(st.session_state.processed_emails)

if __name__ == '__main__':
    main()
