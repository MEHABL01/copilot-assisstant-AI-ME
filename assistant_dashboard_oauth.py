
import streamlit as st
import requests
import json
import os
from dotenv import load_dotenv
from requests_oauthlib import OAuth2Session

# Load environment variables
load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REDIRECT_URI = "http://localhost:8501"

# OAuth2 Authorization Code Flow
AUTH_BASE_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
SCOPE = ["Mail.ReadWrite", "Calendars.ReadWrite", "Chat.ReadWrite", "Files.ReadWrite.Selected", "Tasks.ReadWrite", "Bookings.ReadWrite", "User.Read", "offline_access"]

# Initialize OAuth2 session
oauth = OAuth2Session(CLIENT_ID, redirect_uri=REDIRECT_URI, scope=SCOPE)

# Streamlit app
st.title("Personal AI Copilot")

# Login button
if "oauth_token" not in st.session_state:
    authorization_url, state = oauth.authorization_url(AUTH_BASE_URL)
    st.write("Please [login to Microsoft](%s) to authorize." % authorization_url)
else:
    oauth.token = st.session_state["oauth_token"]

    # Function to refresh token
    def refresh_token():
        extra = {
            'client_id': CLIENT_ID,
            'client_secret': CLIENT_SECRET,
        }
        oauth.refresh_token(TOKEN_URL, **extra)

    # Function to make API calls
    def make_api_call(endpoint):
        response = oauth.get(endpoint)
        if response.status_code == 401:
            refresh_token()
            response = oauth.get(endpoint)
        return response.json()

    # Display user profile
    user_profile = make_api_call("https://graph.microsoft.com/v1.0/me")
    st.write("### User Profile")
    st.json(user_profile)

    # Display emails
    emails = make_api_call("https://graph.microsoft.com/v1.0/me/messages")
    st.write("### Emails")
    st.json(emails)

    # Display calendar events
    events = make_api_call("https://graph.microsoft.com/v1.0/me/events")
    st.write("### Calendar Events")
    st.json(events)

    # Display Teams chats
    chats = make_api_call("https://graph.microsoft.com/v1.0/me/chats")
    st.write("### Teams Chats")
    st.json(chats)

    # Display OneDrive files
    files = make_api_call("https://graph.microsoft.com/v1.0/me/drive/root/children")
    st.write("### OneDrive/SharePoint Files")
    st.json(files)

    # Display tasks
    tasks = make_api_call("https://graph.microsoft.com/v1.0/me/todo/lists")
    st.write("### Tasks")
    st.json(tasks)

    # Display bookings
    bookings = make_api_call("https://graph.microsoft.com/v1.0/solutions/bookingBusinesses")
    st.write("### Bookings")
    st.json(bookings)

# Save token to session state
if oauth.authorized:
    st.session_state["oauth_token"] = oauth.token
