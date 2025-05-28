import streamlit as st
import requests
from dotenv import load_dotenv
import os

# Load credentials from .env file
load_dotenv()
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")
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

# Refresh the access token
ACCESS_TOKEN = refresh_access_token()

# Function to get user profile
def get_user_profile(token):
    url = "https://graph.microsoft.com/v1.0/me"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    return response.json()

# Function to get calendar events
def get_calendar_events(token):
    url = "https://graph.microsoft.com/v1.0/me/events"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    return response.json().get("value", [])

# Function to get unread emails
def get_unread_emails(token):
    url = "https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages?$filter=isRead eq false&$top=10"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    return response.json().get("value", [])

# Function to send email
def send_email(token, subject, body, recipients):
    url = "https://graph.microsoft.com/v1.0/me/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    email = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body
            },
            "toRecipients": [{"emailAddress": {"address": recipient}} for recipient in recipients]
        }
    }
    response = requests.post(url, headers=headers, json=email)
    return response.status_code

# Function to get Teams chat messages
def get_teams_messages(token):
    url = "https://graph.microsoft.com/v1.0/me/chats"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    return response.json().get("value", [])

# Function to get OneDrive files
def get_onedrive_files(token):
    url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    return response.json().get("value", [])

# Function to get Microsoft To Do tasks
def get_todo_tasks(token):
    url = "https://graph.microsoft.com/v1.0/me/todo/lists"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    return response.json().get("value", [])

# Streamlit UI
st.title("Assistant Dashboard")

# Display user profile
user_profile = get_user_profile(ACCESS_TOKEN)
st.header("👤 User Profile")
st.write(f"Name: {user_profile.get('displayName', 'N/A')}")
st.write(f"Email: {user_profile.get('mail', 'N/A')}")
st.write(f"Job Title: {user_profile.get('jobTitle', 'N/A')}")
st.write("---")

# Display calendar events
st.header("📅 Calendar Events")
for event in get_calendar_events(ACCESS_TOKEN):
    st.subheader(event.get("subject", "No Subject"))
    st.write(f"Start: {event['start']['dateTime']}")
    st.write(f"End: {event['end']['dateTime']}")
    st.write(f"Location: {event['location']['displayName']}")
    st.write("---")

# Display unread emails
st.header("📧 Unread Emails")
for email in get_unread_emails(ACCESS_TOKEN):
    st.subheader(email.get("subject", "No Subject"))
    st.write(f"From: {email['from']['emailAddress']['name']}")
    st.write(f"Received: {email['receivedDateTime']}")
    st.write("---")

# Email sending form
st.header("✉️ Send Email")
with st.form(key='send_email_form'):
    subject = st.text_input("Subject")
    body = st.text_area("Body")
    recipients = st.text_input("Recipients (comma-separated)")
    submit_button = st.form_submit_button(label='Send Email')
    if submit_button:
        recipient_list = [email.strip() for email in recipients.split(",")]
        status_code = send_email(ACCESS_TOKEN, subject, body, recipient_list)
        if status_code == 202:
            st.success("Email sent successfully!")
        else:
            st.error("Failed to send email.")

# Display Teams chat messages
st.header("💬 Teams Chat Messages")
for chat in get_teams_messages(ACCESS_TOKEN):
    st.subheader(chat.get("topic", "No Topic"))
    st.write(f"Last Updated: {chat['lastUpdatedDateTime']}")
    st.write("---")

# Display OneDrive files
st.header("📁 OneDrive Files")
for file in get_onedrive_files(ACCESS_TOKEN):
    st.subheader(file.get("name", "No Name"))
    st.write(f"Last Modified: {file['lastModifiedDateTime']}")
    st.write("---")

# Display Microsoft To Do tasks
st.header("✅ Microsoft To Do Tasks")
for task in get_todo_tasks(ACCESS_TOKEN):
    st.subheader(task.get("title", "No Title"))
    st.write("---")

# Refresh button
if st.button("🔄 Refresh"):
    st.experimental_rerun()
