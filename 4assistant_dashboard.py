import streamlit as st
import requests
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

# Authentication and token management
def get_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    response = requests.post(url, headers=headers, data=data)
    response.raise_for_status()
    return response.json()["access_token"]

# Function to get user profile
def get_user_profile(token):
    url = "https://graph.microsoft.com/v1.0/me"
    headers = {
        "Authorization": f"Bearer {token}"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

# Function to read emails
def read_emails(token):
    url = "https://graph.microsoft.com/v1.0/me/messages"
    headers = {
        "Authorization": f"Bearer {token}"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

# Function to manage calendar events
def manage_calendar_events(token):
    url = "https://graph.microsoft.com/v1.0/me/events"
    headers = {
        "Authorization": f"Bearer {token}"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

# Function to read Teams chat messages
def read_teams_messages(token):
    url = "https://graph.microsoft.com/v1.0/me/chats"
    headers = {
        "Authorization": f"Bearer {token}"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

# Function to access OneDrive/SharePoint files
def access_files(token):
    url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
    headers = {
        "Authorization": f"Bearer {token}"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

# Function to manage tasks
def manage_tasks(token):
    url = "https://graph.microsoft.com/v1.0/me/todo/lists"
    headers = {
        "Authorization": f"Bearer {token}"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

# Function to manage bookings
def manage_bookings(token):
    url = "https://graph.microsoft.com/v1.0/me/bookingsBusinesses"
    headers = {
        "Authorization": f"Bearer {token}"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

# Streamlit UI
st.title("Assistant Dashboard")

token = get_token()

st.header("User Profile")
user_profile = get_user_profile(token)
st.write(user_profile)

st.header("Emails")
emails = read_emails(token)
st.write(emails)

st.header("Calendar Events")
calendar_events = manage_calendar_events(token)
st.write(calendar_events)

st.header("Teams Messages")
teams_messages = read_teams_messages(token)
st.write(teams_messages)

st.header("Files")
files = access_files(token)
st.write(files)

st.header("Tasks")
tasks = manage_tasks(token)
st.write(tasks)

st.header("Bookings")
bookings = manage_bookings(token)
st.write(bookings)

# Instructions for deployment
st.header("Deployment Instructions")
st.write("""
1. Create a .env file in the same folder with:
   TENANT_ID=your_tenant_id_here
   CLIENT_ID=your_client_id_here
   CLIENT_SECRET=your_client_secret_here
2. Install dependencies:
   pip install streamlit requests python-dotenv
3. Run the app:
   python -m streamlit run assistant_dashboard.py
""")
