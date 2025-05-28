{
    "chunks": [
        {
            "type": "txt",
            "chunk_number": 1,
            "lines": [
                {
                    "line_number": 1,
                    "text": "from dotenv import load_dotenv"
                },
                {
                    "line_number": 2,
                    "text": "import os"
                },
                {
                    "line_number": 3,
                    "text": "from msal import ConfidentialClientApplication"
                },
                {
                    "line_number": 4,
                    "text": ""
                },
                {
                    "line_number": 5,
                    "text": "# Load environment variables"
                },
                {
                    "line_number": 6,
                    "text": "load_dotenv()"
                },
                {
                    "line_number": 7,
                    "text": ""
                },
                {
                    "line_number": 8,
                    "text": "tenant_id = os.getenv(\"TENANT_ID\")"
                },
                {
                    "line_number": 9,
                    "text": "client_id = os.getenv(\"CLIENT_ID\")"
                },
                {
                    "line_number": 10,
                    "text": "client_secret = os.getenv(\"CLIENT_SECRET\")"
                },
                {
                    "line_number": 11,
                    "text": ""
                },
                {
                    "line_number": 12,
                    "text": "authority = f\"https://login.microsoftonline.com/{tenant_id}\""
                },
                {
                    "line_number": 13,
                    "text": "scopes = [\"https://graph.microsoft.com/.default\"]"
                },
                {
                    "line_number": 14,
                    "text": ""
                },
                {
                    "line_number": 15,
                    "text": "app = ConfidentialClientApplication("
                },
                {
                    "line_number": 16,
                    "text": "client_id,"
                },
                {
                    "line_number": 17,
                    "text": "authority=authority,"
                },
                {
                    "line_number": 18,
                    "text": "client_credential=client_secret"
                },
                {
                    "line_number": 19,
                    "text": ")"
                },
                {
                    "line_number": 20,
                    "text": ""
                },
                {
                    "line_number": 21,
                    "text": "result = app.acquire_token_for_client(scopes=scopes)"
                },
                {
                    "line_number": 22,
                    "text": ""
                },
                {
                    "line_number": 23,
                    "text": "if \"access_token\" in result:"
                },
                {
                    "line_number": 24,
                    "text": "print(\"\u2705 Access token acquired!\")"
                },
                {
                    "line_number": 25,
                    "text": "print(result[\"access_token\"])"
                },
                {
                    "line_number": 26,
                    "text": "else:"
                },
                {
                    "line_number": 27,
                    "text": "print(\"\u274c Failed to acquire token:\")"
                },
                {
                    "line_number": 28,
                    "text": "print(result.get(\"error\"))"
                },
                {
                    "line_number": 29,
                    "text": "print(result.get(\"error_description\"))"
                }
            ],
            "token_count": 57
        }
    ]
}