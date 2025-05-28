from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient
import json
from dotenv import set_key
import os

# Replace with your actual Key Vault URL and secret name
key_vault_url = "https://MyCopilotAssistant-ME.vault.azure.net/"
secret_name = "COPILOT-PERSONAL-AI-ME-Secret"

# Authenticate and create the client
credential = DefaultAzureCredential()
client = SecretClient(vault_url=key_vault_url, credential=credential)

# Retrieve the secret
retrieved_secret = client.get_secret(secret_name)
secret_value = retrieved_secret.value

# Parse the JSON string
try:
    secret_data = json.loads(secret_value)
    tenant_id = secret_data.get("tenant_id")
    client_secret = secret_data.get("client_secret")

    print("✅ Secret retrieved successfully:")
    print(f"Tenant ID: {tenant_id}")
    print(f"Client Secret: {client_secret}")

    # Write to .env file
    env_file = '.env'
    if not os.path.exists(env_file):
        open(env_file, 'a').close()

    set_key(env_file, 'TENANT_ID', tenant_id)
    set_key(env_file, 'CLIENT_SECRET', client_secret)

    print(f"✅ Tenant ID and Client Secret written to {env_file} successfully.")
except json.JSONDecodeError:
    print("❌ Failed to parse secret value as JSON.")
