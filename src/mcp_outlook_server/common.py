import os, logging
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP
from office365.graph_client import GraphClient
import requests

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[logging.FileHandler('mcp_outlook.log'), logging.StreamHandler()])
logger = logging.getLogger('mcp_outlook')

# Helper function for recipient formatting
def _fmt(addrs): return [{"emailAddress": {"address": a}} for a in addrs]

load_dotenv()
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID = os.getenv('TENANT_ID')

if not CLIENT_ID or not CLIENT_SECRET or not TENANT_ID:
	raise ValueError("Missing required environment variables: CLIENT_ID, CLIENT_SECRET, and TENANT_ID must be set")

mcp = FastMCP(name="mcp_outlook",
              instructions="This server provides tools to interact with Outlook emails. "
                          "Each tool requires a user_email parameter to specify which mailbox to access.")
graph_client = GraphClient(tenant=TENANT_ID).with_client_secret(client_id=CLIENT_ID, client_secret=CLIENT_SECRET)

def _get_graph_access_token():
    # Get a fresh token using client credentials (same as graph_client)
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    resp = requests.post(token_url, data=data)
    resp.raise_for_status()
    return resp.json()["access_token"]