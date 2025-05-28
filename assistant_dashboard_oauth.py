import streamlit as st
import requests
import json
import os
from dotenv import load_dotenv
from requests_oauthlib import OAuth2Session
import logging

# Load environment variables
load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REDIRECT_URI = "http://localhost:8501"

# OAuth2 Authorization Code Flow
AUTH_BASE_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
SCOPE = [
    "User.Read",
    "offline_access"
]  # Reduced the SCOPE list to include only permissions that do not require admin consent.

# Initialize OAuth2 session
oauth = OAuth2Session(CLIENT_ID, redirect_uri=REDIRECT_URI, scope=SCOPE)

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Streamlit app
st.title("Personal AI Copilot")

# Login button
if "oauth_token" not in st.session_state:
    logger.info("User not logged in. Generating authorization URL.")
    authorization_url, state = oauth.authorization_url(AUTH_BASE_URL)
    st.write("Please [login to Microsoft](%s) to authorize." % authorization_url)
else:
    logger.info("User already logged in. Using existing token.")
    oauth.token = st.session_state["oauth_token"]

    # Function to refresh token
    def refresh_token():
        extra = {
            'client_id': CLIENT_ID,
            'client_secret': CLIENT_SECRET,
        }
        oauth.refresh_token(TOKEN_URL, **extra)

    # Add error handling for API calls
    def make_api_call(endpoint):
        try:
            logger.debug(f"Making API call to: {endpoint}")
            response = oauth.get(endpoint)
            if response.status_code == 401:
                logger.warning("Token expired. Refreshing token.")
                refresh_token()
                response = oauth.get(endpoint)
            response.raise_for_status()
            logger.debug(f"API call successful: {endpoint}")
            return response.json()
        except requests.exceptions.RequestException as e:
            logger.error(f"Error during API call to {endpoint}: {e}")
            st.error(f"Error fetching data: {e}")
            return {}

    # Improve UI elements
    st.sidebar.title("Navigation")
    options = st.sidebar.radio("Select a feature:", ["User Profile", "Emails", "Calendar Events", "Teams Chats", "Files", "Tasks"])

    # Adjusted the OAuth2 session initialization to dynamically request scopes
    # and handle admin consent issues more gracefully.

    # Function to dynamically request additional scopes
    def request_additional_scopes(scopes):
        try:
            logger.info(f"Requesting additional scopes: {scopes}")
            additional_oauth = OAuth2Session(CLIENT_ID, redirect_uri=REDIRECT_URI, scope=scopes)
            authorization_url, state = additional_oauth.authorization_url(AUTH_BASE_URL)
            st.write("Please [login to Microsoft](%s) to authorize additional permissions." % authorization_url)
        except Exception as e:
            logger.error(f"Error requesting additional scopes: {e}")
            st.error(f"Error requesting additional permissions: {e}")

    # Updated navigation options to request additional permissions dynamically
    if options == "User Profile":
        user_profile = make_api_call("https://graph.microsoft.com/v1.0/me")
        st.write("### User Profile")
        st.json(user_profile)
    elif options == "Emails":
        if "Mail.Read" not in SCOPE:
            request_additional_scopes(["Mail.Read"])
        else:
            emails = make_api_call("https://graph.microsoft.com/v1.0/me/messages")
            st.write("### Emails")
            st.json(emails)
    elif options == "Calendar Events":
        if "Calendars.Read" not in SCOPE:
            request_additional_scopes(["Calendars.Read"])
        else:
            events = make_api_call("https://graph.microsoft.com/v1.0/me/events")
            st.write("### Calendar Events")
            st.json(events)
    elif options == "Teams Chats":
        chats = make_api_call("https://graph.microsoft.com/v1.0/me/chats")
        st.write("### Teams Chats")
        st.json(chats)
    elif options == "Files":
        files = make_api_call("https://graph.microsoft.com/v1.0/me/drive/root/children")
        st.write("### OneDrive/SharePoint Files")
        st.json(files)
    elif options == "Tasks":
        tasks = make_api_call("https://graph.microsoft.com/v1.0/me/todo/lists")
        st.write("### Tasks")
        st.json(tasks)

# Save token to session state
if oauth.authorized:
    st.session_state["oauth_token"] = oauth.token
