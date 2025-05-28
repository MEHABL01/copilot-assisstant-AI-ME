import streamlit as st
import requests
from dotenv import load_dotenv
import os

# Load credentials from .env file
load_dotenv()
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

# Function to refresh access token using client credentials flow
def refresh_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": CLIENT_ID,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials"
    }
    response = requests.post(url, headers=headers, data=data)
    return response.json().get("access_token")

ACCESS_TOKEN = refresh_access_token()

# Helper to make Graph API calls
def graph_get(endpoint):
    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}"}
    response = requests.get(f"https://graph.microsoft.com/v1.0{endpoint}", headers=headers)
    return response.json().get("value", [])

# User Profile
def get_user_profile():
    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}"}
    response = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers)
    return response.json()

# Send Email
def send_email(subject, body, recipients):
    url = "https://graph.microsoft.com/v1.0/me/sendMail"
    headers = {"Authorization": f"Bearer {ACCESS_TOKEN}", "Content-Type": "application/json"}
    email = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": r.strip()}} for r in recipients]
        }
    }
    return requests.post(url, headers=headers, json=email).status_code

# Streamlit UI
st.title("Copilot Assistant Dashboard")

# User Profile
st.header("👤 User Profile")
profile = get_user_profile()
st.write(f"Name: {profile.get('displayName')}")
st.write(f"Email: {profile.get('mail') or profile.get('userPrincipalName')}")
st.write(f"Job Title: {profile.get('jobTitle')}")
st.write("---")

# Calendar Events
st.header("📅 Calendar Events")
for event in graph_get("/me/events"):
    st.subheader(event.get("subject", "No Subject"))
    st.write(f"Start: {event['start']['dateTime']}")
    st.write(f"End: {event['end']['dateTime']}")
    st.write("---")

# Emails
st.header("📧 Unread Emails")
for email in graph_get("/me/mailFolders/Inbox/messages?$filter=isRead eq false&$top=10"):
    st.subheader(email.get("subject", "No Subject"))
    st.write(f"From: {email['from']['emailAddress']['name']}")
    st.write(f"Received: {email['receivedDateTime']}")
    st.write("---")

# Send Email
st.header("✉️ Send Email")
with st.form("send_email_form"):
    subject = st.text_input("Subject")
    body = st.text_area("Body")
    recipients = st.text_input("Recipients (comma-separated)")
    if st.form_submit_button("Send"):
        status = send_email(subject, body, recipients.split(","))
        if status == 202:
            st.success("Email sent successfully!")
        else:
            st.error("Failed to send email.")

# Teams Chats
st.header("💬 Teams Chats")
for chat in graph_get("/me/chats"):
    st.subheader(chat.get("topic", "No Topic"))
    st.write(f"Chat ID: {chat['id']}")
    st.write("---")

# OneDrive Files
st.header("📁 OneDrive Files")
for file in graph_get("/me/drive/root/children"):
    st.subheader(file.get("name"))
    st.write(f"Last Modified: {file['lastModifiedDateTime']}")
    st.write("---")

# To Do Tasks
st.header("✅ Microsoft To Do Lists")
for task_list in graph_get("/me/todo/lists"):
    st.subheader(task_list.get("displayName"))
    st.write(f"ID: {task_list['id']}")
    st.write("---")

# Bookings
st.header("📘 Bookings Businesses")
for booking in graph_get("/solutions/bookingBusinesses"):
    st.subheader(booking.get("displayName"))
    st.write(f"Business Type: {booking.get('businessType')}")
    st.write("---")

# Refresh
if st.button("🔄 Refresh"):
    st.experimental_rerun()
