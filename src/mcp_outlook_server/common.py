import os, logging
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP
from office365.graph_client import GraphClient

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler('mcp_outlook.log'), logging.StreamHandler()]
)
logger = logging.getLogger('mcp_outlook')

# Helper function for recipient formatting
def _fmt(addrs):
    return [{"emailAddress": {"address": a}} for a in addrs]

# Load environment variables
load_dotenv()

# Configuration
ID_CLIENT = os.getenv('ID_CLIENT')
APP_SECRET = os.getenv('APP_SECRET')
TENANT_ID = os.getenv('TENANT_ID')
USER_EMAIL = os.getenv('USER_EMAIL')

# Initialize MCP server
mcp = FastMCP(
    name="mcp_outlook",
    instructions=f"This server provides tools to interact with Outlook emails for the user {USER_EMAIL}."
)

# Initial outlook context
graph_client = GraphClient(tenant=TENANT_ID).with_client_secret(
    client_id=ID_CLIENT,
    client_secret=APP_SECRET
)