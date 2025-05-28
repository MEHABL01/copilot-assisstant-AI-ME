
import streamlit as st
import requests
import json
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()

TENANT_ID = os.getenv('TENANT_ID')
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')

# Function to get access token
def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    data = {
        'grant_type': 'client_credentials',
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scope': 'https://graph.microsoft.com/.default'
    }
    response = requests.post(url, headers=headers, data=data)
    return response.json().get('access_token')

# Function to make Microsoft Graph API requests
def make_graph_api_request(endpoint):
    access_token = get_access_token()
    url = f"https://graph.microsoft.com/v1.0{endpoint}"
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    response = requests.get(url, headers=headers)
    return response.json()

# Streamlit UI
st.title("Personal AI Copilot")

# Email Section
st.header("Emails")
if st.button("Read Emails"):
    emails = make_graph_api_request("/me/messages")
    st.write(emails)

# Calendar Section
st.header("Calendar")
if st.button("Read Calendar Events"):
    events = make_graph_api_request("/me/events")
    st.write(events)

# Teams Chat Section
st.header("Teams Chats")
if st.button("Read Teams Chats"):
    chats = make_graph_api_request("/me/chats")
    st.write(chats)

# OneDrive/SharePoint Files Section
st.header("OneDrive/SharePoint Files")
if st.button("Read OneDrive/SharePoint Files"):
    files = make_graph_api_request("/me/drive/root/children")
    st.write(files)

# Tasks Section
st.header("Tasks")
if st.button("Read Tasks"):
    tasks = make_graph_api_request("/me/todo/lists")
    st.write(tasks)

# Bookings Section
st.header("Bookings")
if st.button("Read Bookings"):
    bookings = make_graph_api_request("/solutions/bookingBusinesses")
    st.write(bookings)

# User Profile Section
st.header("User Profile")
if st.button("Read User Profile"):
    profile = make_graph_api_request("/me")
    st.write(profile)

# Deployment Instructions
st.header("Deployment Instructions")
st.write("""
1. **Create a `.env` file** in the same folder with the following content:
   TENANT_ID=your_tenant_id_here
   CLIENT_ID=your_client_id_here
   CLIENT_SECRET=your_client_secret_here

2. **Install dependencies**:

   pip install streamlit requests python-dotenv

3. **Run the app**:

   streamlit run assistant_dashboard.py
""")
