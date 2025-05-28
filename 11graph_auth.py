from dotenv import load_dotenv
import os
from msal import ConfidentialClientApplication

# Load environment variables
load_dotenv()

tenant_id = os.getenv("TENANT_ID")
client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")

authority = f"https://login.microsoftonline.com/{tenant_id}"
scopes = ["https://graph.microsoft.com/.default"]

app = ConfidentialClientApplication(
    client_id,
    authority=authority,
    client_credential=client_secret
)

result = app.acquire_token_for_client(scopes=scopes)

if "access_token" in result:
    print("✅ Access token acquired!")
    print(result["access_token"])
else:
    print("❌ Failed to acquire token:")
    print(result.get("error"))
    print(result.get("error_description"))
