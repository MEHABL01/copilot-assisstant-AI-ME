
import streamlit as st
import requests

# Function to get access token
def get_access_token():
    # Replace with your actual token retrieval logic
    return "YOUR_ACCESS_TOKEN"

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

# Function to get Microsoft To Do tasks
def get_todo_tasks(token):
    url = "https://graph.microsoft.com/v1.0/me/todo/lists"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers)
    return response.json().get("value", [])

# Streamlit UI
st.title("Assistant Dashboard")

token = get_access_token()

st.header("📅 Calendar Events")
for event in get_calendar_events(token):
    st.subheader(event.get("subject", "No Subject"))
    st.write(f"Start: {event['start']['dateTime']}")
    st.write(f"End: {event['end']['dateTime']}")
    st.write(f"Location: {event['location']['displayName']}")
    st.write("---")

st.header("📧 Unread Emails")
for email in get_unread_emails(token):
    st.subheader(email.get("subject", "No Subject"))
    st.write(f"From: {email['from']['emailAddress']['name']}")
    st.write(f"Received: {email['receivedDateTime']}")
    st.write("---")

st.header("✅ Microsoft To Do Tasks")
for task in get_todo_tasks(token):
    st.subheader(task.get("title", "No Title"))
    st.write("---")

if st.button("🔄 Refresh"):
    st.rerun()
