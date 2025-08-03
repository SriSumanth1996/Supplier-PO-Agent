from __future__ import print_function
import os.path
import base64
import email
import pandas as pd
import html2text
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

# Set up logging
logger = get_logger(__name__)

# Load OpenAI API key from secrets or file
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

# Gmail read+write and Calendar scope
SCOPES = [
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/calendar'
]

# QA mapping for extraction
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

# Initialize session state
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'gmail_service' not in st.session_state:
    st.session_state.gmail_service = None
if 'calendar_service' not in st.session_state:
    st.session_state.calendar_service = None
if 'processed_emails' not in st.session_state:
    st.session_state.processed_emails = []

def authenticate_gmail_and_calendar():
    creds = None
    
    # First try using refresh token from secrets
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
    
    # If no valid creds, initiate OAuth flow
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
    
    # For environments without browser access
    if not st.runtime.exists():
        auth_url, _ = flow.authorization_url(prompt='consent')
        st.markdown(f"""
        ### Authentication Required
        1. Click [this link]({auth_url}) to authorize
        2. You'll be redirected to a localhost URL that won't work
        3. Copy the **entire URL** from your browser's address bar
        4. Paste it below
        """)
        
        code_url = st.text_input("Paste the redirect URL here:")
        if code_url:
            try:
                code = parse_qs(urlparse(code_url).query)['code'][0]
                flow.fetch_token(code=code)
                creds = flow.credentials
                
                # Store refresh token for future use (in production you'd save this securely)
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
    
    # For local environments with browser access
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
       - Be specific about the product (e.g., "iPhone 15 Pro 256GB" not just "phone")
    2. QUANTITIES/UNITS:
       - Extract numerical quantities (e.g., "5 units", "100 pieces", "2 boxes")
       - Return just the number with unit type (e.g., "50 pieces")
    3. UNIT PRICE:
       - Extract ONLY the price per single unit/item as a clean number with currency
       - Look for currency symbols ($, €, ₹, etc.) and amounts
       - Return ONLY the price (e.g., "₹180" or "$25.50" - NO descriptive text)
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
       - Look for supplier's city, state, or location in email signature
    7. SENDER NAME:
       - Extract personal name from email signature (e.g., "Ramesh Patel")
    8. COMPANY NAME:
       - Extract company or organization name from signature (e.g., "Shree Fasteners Pvt. Ltd.")
    9. CONTACT NUMBER:
       - Extract phone number (e.g., "+91 98765 43210")
    10. DESIGNATION:
       - Extract sender's designation or job title (e.g., "Sales Manager") of the person who is sending the mail
    RESPONSE RULES:
    - Extract ONLY explicitly stated info.
    - Do not hallucinate or take from the examples given above in this prompt
    - If not found, respond with "Not present"
    - Keep original format for non-price fields
    - Don't assume or guess, only state exact extracted values
    ANSWER:"""
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
    RESPOND WITH ONLY THE CLASSIFICATION CATEGORY NAME (exactly as written above):"""
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
    """Use GPT to extract and interpret proposed meeting datetime directly"""
    # Get current IST time
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
    """Send reply to sender using Gmail API"""
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
    """Check if there's a conflict in the calendar for the given time slot"""
    try:
        events_result = calendar_service.events().list(
            calendarId='primary',
            timeMin=start_time.isoformat(),
            timeMax=end_time.isoformat(),
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        events = events_result.get('items', [])

        # Check if any existing event conflicts with the proposed time
        for event in events:
            event_start = datetime.fromisoformat(
                event['start'].get('dateTime', event['start'].get('date')).replace('Z', '+00:00'))
            event_end = datetime.fromisoformat(
                event['end'].get('dateTime', event['end'].get('date')).replace('Z', '+00:00'))

            # If there's any overlap, it's a conflict
            if not (end_time <= event_start or start_time >= event_end):
                return True, event.get('summary', 'Existing meeting')

        return False, None
    except Exception as e:
        print(f"Error checking calendar conflict: {e}")
        return False, None

def schedule_meeting(calendar_service, quotation_data, email_address, proposed_datetime=None, classification="Unknown"):
    """Schedule a 30-minute meeting with supplier with dynamic event description based on classification"""
    try:
        ist = pytz.timezone('Asia/Kolkata')
        now = datetime.now(ist)

        if proposed_datetime:
            # Ensure proposed time is during business hours (9 AM to 5 PM IST)
            start_hour = proposed_datetime.hour
            if start_hour < 9 or start_hour >= 17:
                # If outside business hours, don't auto-schedule
                return None, "outside_business_hours"

            start_time = proposed_datetime
            end_time = start_time + timedelta(minutes=30)

            # If the proposed time is in the past, don't auto-schedule
            if start_time < now:
                return None, "past_time"

            # Check for calendar conflicts
            has_conflict, conflicting_event = check_calendar_conflict(calendar_service, start_time, end_time)
            if has_conflict:
                return None, "conflict"
        else:
            # If no specific time proposed, don't auto-schedule
            return None, "no_specific_time"

        # Get company name or use "Unknown Company" if not present
        company_name = quotation_data.get('company_name', 'Unknown Company')
        if company_name == "Not present":
            company_name = "Unknown Company"

        # Dynamic event description based on classification
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

def get_reply_body(classification, quotation_data, sender_name, meeting_details=None, meeting_result=None):
    """Generate appropriate reply body based on classification and meeting details"""
    if classification == "Quotation Received":
        base_message = f"""Dear {sender_name or 'Supplier'},

Thank you for your quotation. We have received the full details regarding your product and pricing."""

        if meeting_details and meeting_details.get("meeting_intent") == "Yes":
            if meeting_details.get("proposed_datetime") == "Not specified":
                meeting_text = "\n\nYou have requested a meeting but did not specify complete date and time details. Please provide your preferred date and time for the meeting."
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
                meeting_text = "\n\nRegarding your meeting request, please provide specific date and time preferences for scheduling."

            base_message += meeting_text

        base_message += "\n\nLooking forward to your response.\n\nBest regards,\nDr. Saravanan Kesavan\nBITSoM"
        return base_message

    elif classification == "Quotation Partially Received":
        # Identify missing information
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

        if meeting_details and meeting_details.get("meeting_intent") == "Yes":
            if meeting_details.get("proposed_datetime") == "Not specified":
                meeting_text = "\n\nYou have requested a meeting but did not specify complete date and time details. Please provide your preferred date and time for the meeting."
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
                meeting_text = "\n\nRegarding your meeting request, please provide specific date and time preferences for scheduling."

            base_message += meeting_text

        base_message += "\n\nOnce we receive the complete information, we will be able to proceed with our evaluation.\n\nThank you for your cooperation.\n\nBest regards,\nDr. Saravanan Kesavan\nBITSoM"
        return base_message

    elif classification == "New Business Connection":
        base_message = f"""Dear {sender_name or 'Supplier'},

Thank you for introducing your company and sharing your offerings with us."""

        if meeting_details and meeting_details.get("meeting_intent") == "Yes":
            if meeting_details.get("proposed_datetime") == "Not specified":
                meeting_text = "\n\nYou have requested a meeting but did not specify complete date and time details. Please provide your preferred date and time for the meeting."
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
                meeting_text = "\n\nRegarding your meeting request, please provide specific date and time preferences for scheduling."

            base_message += meeting_text

        base_message += "\n\nWe will keep your information on record and reach out to you when opportunities arise.\n\nWarm regards,\nDr. Saravanan Kesavan\nBITSoM"
        return base_message

    else:
        return ""

def display_email_data(email_address, final_classification, quotation_data, meeting_details=None):
    st.subheader(f"Email from: {email_address}")
    
    # Classification badge
    if final_classification == "Quotation Received":
        st.success(f"Classification: {final_classification}")
    elif final_classification == "Quotation Partially Received":
        st.warning(f"Classification: {final_classification}")
    elif final_classification == "New Business Connection":
        st.info(f"Classification: {final_classification}")
    else:
        st.error(f"Classification: {final_classification}")

    if meeting_details and meeting_details.get("meeting_intent") == "Yes":
        st.subheader("Meeting Details")
        st.write(f"Meeting Requested: Yes")
        if meeting_details.get("proposed_datetime") != "Not specified":
            try:
                proposed_dt = datetime.fromisoformat(meeting_details["proposed_datetime"])
                formatted_date = proposed_dt.strftime("%B %d, %Y")
                formatted_time = proposed_dt.strftime("%I:%M %p")
                st.write(f"Proposed Time: {formatted_date} at {formatted_time}")
            except:
                st.write("Proposed Time: Not specified")
        else:
            st.write("Proposed Time: Not specified")

    # Sender information
    st.subheader("Sender Information")
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Name:** {quotation_data.get('sender_name', 'Not present')}")
        st.write(f"**Designation:** {quotation_data.get('designation', 'Not present')}")
    with col2:
        st.write(f"**Company:** {quotation_data.get('company_name', 'Not present')}")
        st.write(f"**Contact:** {quotation_data.get('contact_number', 'Not present')}")
    st.write(f"**Location:** {quotation_data.get('place', 'Not present')}")

    if final_classification in ["Quotation Received", "Quotation Partially Received"]:
        st.subheader("Quotation Details")
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"**Product:** {quotation_data.get('product', 'Not present')}")
            st.write(f"**Quantity:** {quotation_data.get('quantity', 'Not present')}")
        with col2:
            st.write(f"**Unit Price:** {quotation_data.get('unit_price', 'Not present')}")
            st.write(f"**Total Cost:** {quotation_data.get('total_cost', 'Not present')}")
        st.write(f"**Lead Time:** {quotation_data.get('lead_time', 'Not present')}")

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
    
    for message in messages[:num_emails]:
        msg = gmail_service.users().messages().get(userId='me', id=message['id']).execute()
        headers = msg['payload']['headers']
        sender = [h['value'] for h in headers if h['name'] == 'From'][0]
        subject = [h['value'] for h in headers if h['name'] == 'Subject'][0]
        thread_id = msg['threadId']
        body = get_email_body(msg['payload'])

        # Extract meeting details first
        meeting_details = extract_meeting_details(body)
        initial_classification = classify_email_intent(body)

        quotation_data = extract_quotation_data(body, initial_classification)

        if initial_classification not in ["New Business Connection", "Unknown"]:
            quotation_data = calculate_unit_price_if_missing(quotation_data)
            quotation_data = calculate_total_cost_if_missing(quotation_data)

        final_classification = get_final_classification(quotation_data, initial_classification)

        name = sender.split("<")[0].strip() if "<" in sender else sender
        email_address = sender.split("<")[1][:-1] if "<" in sender else sender

        # Handle meeting scheduling and replies
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

    return processed_emails

def main():
    st.set_page_config(page_title="Supplier Quotation Processor", layout="wide")
    st.title("Supplier Quotation Processing System")
    st.markdown("""
    This application processes supplier emails, extracts quotation details, classifies them, 
    and can automatically respond or schedule meetings.
    """)

    # Authentication section
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
            st.experimental_rerun()

    if not st.session_state.authenticated:
        st.warning("Please authenticate with Google to continue.")
        return

    # Main processing section
    st.header("Process Emails")
    num_emails = st.slider("Number of emails to process", 1, 10, 3)
    
    if st.button("Process Latest Emails"):
        with st.spinner("Processing emails..."):
            try:
                st.session_state.processed_emails = process_emails(
                    st.session_state.gmail_service,
                    st.session_state.calendar_service,
                    num_emails
                )
            except Exception as e:
                st.error(f"Error processing emails: {str(e)}")

    # Display processed emails
    if st.session_state.processed_emails:
        st.header("Processed Emails")
        for i, email_data in enumerate(st.session_state.processed_emails):
            with st.expander(f"{i+1}. {email_data['subject']} - {email_data['email_address']}"):
                display_email_data(
                    email_data['email_address'],
                    email_data['final_classification'],
                    email_data['quotation_data'],
                    email_data['meeting_details']
                )

                # Show reply and actions
                st.subheader("Proposed Reply")
                st.text_area("Reply Content", email_data['reply_body'], height=200, key=f"reply_{i}")

                col1, col2 = st.columns(2)
                with col1:
                    if st.button(f"Send Reply to {email_data['email_address']}", key=f"send_{i}"):
                        with st.spinner("Sending reply..."):
                            success, message = send_reply(
                                st.session_state.gmail_service,
                                email_data['thread_id'],
                                email_data['email_address'],
                                email_data['subject'],
                                email_data['reply_body']
                            )
                            if success:
                                st.success(message)
                            else:
                                st.error(message)
                with col2:
                    if email_data['meeting_result'] and email_data['meeting_result'][0]:
                        st.info(f"Meeting scheduled: {email_data['meeting_result'][1]}")
                    elif email_data['meeting_details'] and email_data['meeting_details'].get('meeting_intent') == 'Yes':
                        st.warning(f"Meeting not scheduled: {email_data['meeting_result'][1] if email_data['meeting_result'] else 'No time specified'}")

if __name__ == '__main__':
    main()
