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
if 'selected_replies_complete' not in st.session_state:
    st.session_state.selected_replies_complete = {}
if 'selected_meetings_complete' not in st.session_state:
    st.session_state.selected_meetings_complete = {}
if 'selected_replies_partial' not in st.session_state:
    st.session_state.selected_replies_partial = {}
if 'selected_meetings_partial' not in st.session_state:
    st.session_state.selected_meetings_partial = {}

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
         patrimony4. Paste it below
        """
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
            elif 'parts' in part:  # Support for nested parts
                body = get_email_body(part)
                if body:
                    break
    else:
        # If no 'parts' field exists, handle the single-part email
        mime_type = msg_payload.get("mimeType")
        body_data = msg_payload.get("body", {}).get("data")
        if mime_type == "text/plain" and body_data:
            body = base64.urlsafe_b64decode(body_data).decode('utf-8')
        elif mime_type == "text/html" and body_data:
            html_body = base64.urlsafe_b64decode(body_data).decode('utf-8')
            body = html2text.html2text(html_body)

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
       - Return just the number with unit type (e.g., "50 Pieces")
    3. UNIT PRICE:
       - Extract ONLY the price per single unit/item as a clean number with currency
       - Look for currency symbols ($, €, ₹, etc.) and amounts
       - Return ONLY the price (e.g., "₹180" or "$25.50" - NO descriptive text like Dollars, Rupees, INR, Euros etc.)
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
       - Extract sender's designation or job title (e.g., "Sales Manager") of the person who is sending the mail
    RESPONSE RULES:
    - Extract ONLY explicitly stated info.
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
       - No specific quotation or pricing information provided
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
    RESPOND WITH ONLY THE CLASSIFICATION CATEGORY NAME (exactly as written above):
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.5,
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
       - If the date (e.g., "6th") has already passed this month, infer the next month.
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
    product = quotation_data.get("product", "Not present")
    lead_time = quotation_data.get("lead_time", "Not present")
    unit_price = quotation_data.get("unit_price", "Not present")
    quantity = quotation_data.get("quantity", "Not present")
    key_elements = [product, lead_time, unit_price, quantity]
    missing_elements = [elem for elem in key_elements if elem == "Not present"]
    if len(missing_elements) == 4:
        return "New Business Connection"
    elif len(missing_elements) > 0:
        return "Quotation Partially Received"
    else:
        return "Quotation Received"

def calculate_unit_price_if_missing(quotation_data):
    unit_cost = quotation_data.get("unit_price", "Not present")
    total_cost = quotation_data.get("total_cost", "Not present")
    units = quotation_data.get("quantity", "Not present")
    if unit_cost == "Not present" and total_cost != "Not present" and units != "Not present":
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
    unit_cost = quotation_data.get("unit_price", "Not present")
    total_cost = quotation_data.get("total_cost", "Not present")
    units = quotation_data.get("quantity", "Not present")
    if total_cost == "Not present" and unit_cost != "Not present" and units != "Not present":
        try:
            unit_price_num = float(''.join([c for c in unit_cost if c.isdigit() or c == '.']))
            quantity_num = float(''.join([c for c in units if c.isdigit() or c == '.']))
            currency_symbol = ''.join([c for c in unit_cost if not (c.isdigit() or c == '.' or c.isspace())])
            total_cost = f"{currency_symbol}{unit_price_num * quantity_num:.2f}"
            quotation_data["total_cost"] = total_cost
        except Exception:
            quotation_data["total_cost"] = "Calculation failed"
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

def get_reply_body(classification, quotation_data, sender_name, meeting_details=None, meeting_result=None, meeting_response=None):
    meeting_text = ""
    if meeting_details and meeting_details.get("meeting_intent") == "Yes":
        if meeting_response == "Yes":
            if meeting_details.get("proposed_datetime") == "Not specified":
                meeting_text = "\n\nRegarding your meeting request, we are open to scheduling a meeting. Please provide specific date and time preferences within our business hours (9:00 AM to 5:00 PM IST)."
            elif meeting_result and meeting_result[1] == "scheduled":
                meeting_text = "\n\nWe have scheduled the meeting as requested. Please check your calendar for the invitation."
            elif meeting_result and meeting_result[1] == "conflict":
                if meeting_details.get("proposed_datetime") != "Not specified":
                    try:
                        proposed_dt = datetime.fromisoformat(meeting_details["proposed_datetime"])
                        formatted_date = proposed_dt.strftime("%B %d, %Y")
                        formatted_time = proposed_dt.strftime("%I:%M %p")
                        meeting_text = f"\n\nYou requested a meeting for {formatted_date} at {formatted_time}, but we have a scheduling conflict at that time. Please suggest an alternative time that works for you."
                    except:
                        meeting_text = "\n\nYou requested a meeting, but we have a scheduling conflict at the proposed time. Please suggest an alternative time that works for you."
                else:
                    meeting_text = "\n\nYou requested a meeting, but we have a scheduling conflict. Please suggest an alternative time that works for you."
            elif meeting_result and meeting_result[1] == "outside_business_hours":
                if meeting_details.get("proposed_datetime") != "Not specified":
                    try:
                        proposed_dt = datetime.fromisoformat(meeting_details["proposed_datetime"])
                        formatted_date = proposed_dt.strftime("%B %d, %Y")
                        formatted_time = proposed_dt.strftime("%I:%M %p")
                        meeting_text = f"\n\nYou requested a meeting for {formatted_date} at {formatted_time}, but this falls outside our business hours (9:00 AM to 5:00 PM IST). Please suggest a time within business hours."
                    except:
                        meeting_text = "\n\nThe proposed meeting time falls outside our business hours (9:00 AM to 5:00 PM IST). Please suggest a time within business hours."
                else:
                    meeting_text = "\n\nThe proposed meeting time falls outside our business hours (9:00 AM to 5:00 PM IST). Please suggest a time within business hours."
            elif meeting_result and meeting_result[1] == "past_time":
                if meeting_details.get("proposed_datetime") != "Not specified":
                    try:
                        proposed_dt = datetime.fromisoformat(meeting_details["proposed_datetime"])
                        formatted_date = proposed_dt.strftime("%B %d, %Y")
                        formatted_time = proposed_dt.strftime("%I:%M %p")
                        meeting_text = f"\n\nYou requested a meeting for {formatted_date} at {formatted_time}, but this time has already passed. Please suggest a future date and time."
                    except:
                        meeting_text = "\n\nThe proposed meeting time has already passed. Please suggest a future date and time."
                else:
                    meeting_text = "\n\nThe proposed meeting time has already passed. Please suggest a future date and time."
            else:
                meeting_text = "\n\nWe are open to scheduling a meeting. Please provide specific date and time preferences within our business hours (9:00 AM to 5:00 PM IST)."
        elif meeting_response == "No":
            meeting_text = "\n\nThank you for your meeting request. At this time, we will proceed with written communication and will reach out if a meeting is needed."
        else:
            meeting_text = "\n\nRegarding your meeting request, please provide specific date and time preferences for scheduling."
    
    if classification == "Quotation Received":
        base_message = f"""Dear {sender_name or 'Supplier'},
Thank you for your quotation. We have received the full details regarding your product and pricing.{meeting_text}
Looking forward to your response.
Best regards,
Dr. Saravanan Kesavan
BITSoM"""
        return base_message
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
        base_message += f"{meeting_text}\n\nOnce we receive the complete information, we will be able to proceed with our evaluation.\n\nThank you for your cooperation.\n\nBest regards,\nDr. Saravanan Kesavan\nBITSoM"
        return base_message
    elif classification == "New Business Connection":
        base_message = f"""Dear {sender_name or 'Supplier'},
Thank you for introducing your company and sharing your offerings with us.{meeting_text}
We will keep your information on record and reach out to you when opportunities arise.
Warm regards,
Dr. Saravanan Kesavan
BITSoM"""
        return base_message
    else:
        return ""

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
            'Meeting Status': get_meeting_status(email.get('meeting_details'), email.get('meeting_result')),
            'Reply': False,
            'Meeting': None
        })
    return pd.DataFrame(data)

def create_quotation_partial_table(emails):
    if not emails:
        return pd.DataFrame()
    data = []
    for email in emails:
        qd = email['quotation_data']
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
            'Unit Priceflood': qd.get('unit_price', 'Not present'),
            'Total Cost': qd.get('total_cost', 'Not present'),
            'Lead Time': qd.get('lead_time', 'Not present'),
            'Location': qd.get('place', 'Not present'),
            'Contact': qd.get('contact_number', 'Not present'),
            'Missing Fields': ', '.join(missing_fields) if missing_fields else 'None',
            'Meeting Status': get_meeting_status(email.get('meeting_details'), email.get('meeting_result')),
            'Reply': False,
            'Meeting': None
        })
    return pd.DataFrame(data)

def create_business_connection_table(emails):
    if not emails:
        return pd.DataFrame()
    data = []
    for email in emails:
        qd = email['quotation_data']
        data.append({
            'Sender Name': qd.get('sender_name', 'Not present'),
            'Company': qd.get('company_name', 'Not present'),
            'Email': email['email_address'],
            'Designation': qd.get('designation', 'Not present'),
            'Location': qd.get('place', 'Not present'),
            'Contact': qd.get('contact_number', 'Not present'),
            'Meeting Status': get_meeting_status(email.get('meeting_details'), email.get('meeting_result')),
            'Reply': False,
            'Meeting': None
        })
    return pd.DataFrame(data)

def send_replies_for_emails(service, emails, selected_replies, selected_meetings):
    success_count = 0
    error_count = 0
    progress_bar = st.progress(0)
    status_text = st.empty()
    selected_emails_list = [
        email for email in emails
        if selected_replies.get(email['email_address'], False)
    ]
    for i, email_data in enumerate(selected_emails_list):
        progress = (i + 1) / len(selected_emails_list)
        progress_bar.progress(progress)
        status_text.text(f'Sending reply {i + 1} of {len(selected_emails_list)}...')
        meeting_response = selected_meetings.get(email_data['email_address'], None)
        reply_body = get_reply_body(
            email_data['final_classification'],
            email_data['quotation_data'],
            email_data['quotation_data'].get('sender_name'),
            email_data.get('meeting_details'),
            email_data.get('meeting_result'),
            meeting_response
        )
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

def schedule_meetings_for_emails(calendar_service, emails, selected_emails):
    success_count = 0
    error_count = 0
    progress_bar = st.progress(0)
    status_text = st.empty()
    selected_emails_list = [email for email in emails if selected_emails.get(email['email_address'], False)]
    for i, email_data in enumerate(selected_emails_list):
        progress = (i + 1) / len(selected_emails_list)
        progress_bar.progress(progress)
        status_text.text(f'Scheduling meeting {i + 1} of {len(selected_emails_list)}...')
        proposed_datetime = None
        if email_data.get('meeting_details', {}).get('proposed_datetime') != "Not specified":
            try:
                proposed_datetime = datetime.fromisoformat(email_data['meeting_details']['proposed_datetime'])
            except Exception:
                proposed_datetime = None
        scheduled_event, status = schedule_meeting(
            calendar_service,
            email_data['quotation_data'],
            email_data['email_address'],
            proposed_datetime,
            email_data['final_classification']
        )
        if status == "scheduled":
            email_data['meeting_result'] = (scheduled_event, status)
            email_data['reply_body'] = get_reply_body(
                email_data['final_classification'],
                email_data['quotation_data'],
                email_data['quotation_data'].get('sender_name'),
                email_data['meeting_details'],
                email_data['meeting_result'],
                "Yes"
            )
            success_count += 1
        else:
            error_count += 1
    progress_bar.progress(1.0)
    status_text.text('Bulk meeting scheduling complete!')
    if success_count > 0:
        st.success(f"Successfully scheduled {success_count} meetings!")
    if error_count > 0:
        st.error(f"Failed to schedule {error_count} meetings.")

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
        scheduled_event = None
        meeting_result = None
        if meeting_details and meeting_details.get("meeting_intent") == "Yes":
            if meeting_details.get("proposed_datetime") != "Not specified":
                try:
                    proposed_datetime = datetime.fromisoformat(meeting_details["proposed_datetime"])
                except Exception as e:
                    proposed_datetime = None
                if proposed_datetime:
                    scheduled_event, status = schedule_meeting(
                        calendar_service,
                        quotation_data,
                        email_address,
                        proposed_datetime,
                        final_classification
                    )
                    meeting_result = (scheduled_event, status)
                else:
                    meeting_result = (None, "parse_error")
            else:
                meeting_result = (None, "incomplete_details")
        reply_body = get_reply_body(
            final_classification,
            quotation_data,
            quotation_data.get("sender_name"),
            meeting_details,
            meeting_result
        )
        processed_emails.append({
            "email_address": email_address,
            "subject": subject,
            "final_classification": final_classification,
            "quotation_data": quotation_data,
            "meeting_details": meeting_details,
            "meeting_result": meeting_result,
            "reply_body": reply_body,
            "thread_id": thread_id
        })
    progress_bar.progress(1.0)
    status_text.text('Processing complete!')
    return processed_emails

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
            st.write("Select emails to send replies or respond to meeting requests:")
            col1, col2 = st.columns([1, 1])
            with col1:
                select_all_replies = st.checkbox("Select All Replies (Complete)", key="select_all_replies_complete")
            with col2:
                select_all_meetings = st.checkbox("Select All Meetings (Complete)", key="select_all_meetings_complete")
            for i, row in df_complete.iterrows():
                email = row['Email']
                meeting_intent = row['Meeting Status'] != "No Meeting Requested"
                col1, col2 = st.columns([1, 1])
                with col1:
                    default_reply = select_all_replies or st.session_state.selected_replies_complete.get(email, False)
                    st.session_state.selected_replies_complete[email] = st.checkbox(
                        f"Reply to {email}", value=default_reply, key=f"reply_complete_{i}"
                    )
                with col2:
                    if meeting_intent:
                        default_meeting = select_all_meetings or st.session_state.selected_meetings_complete.get(email, None)
                        options = ["Yes", "No", None]
                        st.session_state.selected_meetings_complete[email] = st.selectbox(
                            f"Meeting for {email}", options, index=options.index(default_meeting) if default_meeting in options else 2,
                            key=f"meeting_complete_{i}", help="Meeting request detected"
                        )
                    else:
                        st.session_state.selected_meetings_complete[email] = None
                        st.write("No meeting request")
            st.dataframe(df_complete.drop(columns=['Reply', 'Meeting']), use_container_width=True)
            csv_complete = df_complete.drop(columns=['Reply', 'Meeting']).to_csv(index=False)
            st.download_button(
                label="Download Complete Quotations CSV",
                data=csv_complete,
                file_name=f"complete_quotations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
            if st.button("Send Replies"):
                send_replies_for_emails(
                    st.session_state.gmail_service,
                    quotation_received,
                    st.session_state.selected_replies_complete,
                    st.session_state.selected_meetings_complete
                )
        else:
            st.info("No complete quotations found in the processed emails.")
        st.header("Partial Quotations Received")
        if quotation_partial:
            df_partial = create_quotation_partial_table(quotation_partial)
            st.write("Select emails to send replies or respond to meeting requests:")
            col1, col2 = st.columns([1, 1])
            with col1:
                select_all_replies = st.checkbox("Select All Replies (Partial)", key="select_all_replies_partial")
            with col2:
                select_all_meetings = st.checkbox("Select All Meetings (Partial)", key="select_all_meetings_partial")
            for i, row in df_partial.iterrows():
                email = row['Email']
                meeting_intent = row['Meeting Status'] != "No Meeting Requested"
                col1, col2 = st.columns([1, 1])
                with col1:
                    default_reply = select_all_replies or st.session_state.selected_replies_partial.get(email, False)
                    st.session_state.selected_replies_partial[email] = st.checkbox(
                        f"Reply to {email}", value=default_reply, key=f"reply_partial_{i}"
                    )
                with col2:
                    if meeting_intent:
                        default_meeting = select_all_meetings or st.session_state.selected_meetings_partial.get(email, None)
                        options = ["Yes", "No", None]
                        st.session_state.selected_meetings_partial[email] = st.selectbox(
                            f"Meeting for {email}", options, index=options.index(default_meeting) if default_meeting in options else 2,
                            key=f"meeting_partial_{i}", help="Meeting request detected"
                        )
                    else:
                        st.session_state.selected_meetings_partial[email] = None
                        st.write("No meeting request")
            st.dataframe(df_partial.drop(columns=['Reply', 'Meeting']), use_container_width=True)
            csv_partial = df_partial.drop(columns=['Reply', 'Meeting']).to_csv(index=False)
            st.download_button(
                label="Download Partial Quotations CSV",
                data=csv_partial,
                file_name=f"partial_quotations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
            if st.button("Send Replies for Partial Quotations"):
                send_replies_for_emails(
                    st.session_state.gmail_service,
                    quotation_partial,
                    st.session_state.selected_replies_partial,
                    st.session_state.selected_meetings_partial
                )
        else:
            st.info("No partial quotations found in the processed emails.")
    elif tabs == "New Business Connections":
        st.header("New Business Connections")
        if business_connection:
            df_business = create_business_connection_table(business_connection)
            st.write("Select emails to send replies or respond to meeting requests:")
            col1, col2 = st.columns([1, 1])
            with col1:
                select_all_replies = st.checkbox("Select All Replies (Business)", key="select_all_replies_business")
            with col2:
                select_all_meetings = st.checkbox("Select All Meetings (Business)", key="select_all_meetings_business")
            for i, row in df_business.iterrows():
                email = row['Email']
                meeting_intent = row['Meeting Status'] != "No Meeting Requested"
                col1, col2 = st.columns([1, 1])
                with col1:
                    default_reply = select_all_replies or st.session_state.selected_replies_complete.get(email, False)
                    st.session_state.selected_replies_complete[email] = st.checkbox(
                        f"Reply to {email}", value=default_reply, key=f"reply_business_{i}"
                    )
                with col2:
                    if meeting_intent:
                        default_meeting = select_all_meetings or st.session_state.selected_meetings_complete.get(email, None)
                        options = ["Yes", "No", None]
                        st.session_state.selected_meetings_complete[email] = st.selectbox(
                            f"Meeting for {email}", options, index=options.index(default_meeting) if default_meeting in options else 2,
                            key=f"meeting_business_{i}", help="Meeting request detected"
                        )
                    else:
                        st.session_state.selected_meetings_complete[email] = None
                        st.write("No meeting request")
            st.dataframe(df_business.drop(columns=['Reply', 'Meeting']), use_container_width=True)
            csv_business = df_business.drop(columns=['Reply', 'Meeting']).to_csv(index=False)
            st.download_button(
                label="Download Business Connections CSV",
                data=csv_business,
                file_name=f"business_connections_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
            if st.button("Send Replies to New Business Connections"):
                send_replies_for_emails(
                    st.session_state.gmail_service,
                    business_connection,
                    st.session_state.selected_replies_complete,
                    st.session_state.selected_meetings_complete
                )
        else:
            st.info("No new business connection emails found in the processed emails.")
    if unknown:
        st.header("Unknown/Other Classifications")
        st.warning(f"Found {len(unknown)} emails that could not be properly classified:")
        for email in unknown:
            st.write(f"- {email['email_address']}: {email['subject']}")

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
            st.session_state.selected_replies_complete = {}
            st.session_state.selected_meetings_complete = {}
            st.session_state.selected_replies_partial = {}
            st.session_state.selected_meetings_partial = {}
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
