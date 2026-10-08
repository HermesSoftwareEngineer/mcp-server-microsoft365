import os
from urllib.parse import urlparse

from dotenv import load_dotenv

load_dotenv()

PUBLIC_URL = os.environ["PUBLIC_URL"].rstrip("/")
PUBLIC_HOST = urlparse(PUBLIC_URL).netloc
MCP_PATH = "/mcp"
RESOURCE_URL = PUBLIC_URL + MCP_PATH

HOST = os.environ.get("HOST", "127.0.0.1")
PORT = int(os.environ.get("PORT", "8000"))

AZURE_CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
AZURE_CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
AZURE_TENANT_ID = os.environ["AZURE_TENANT_ID"]
MS_REDIRECT_URI = PUBLIC_URL + "/auth/microsoft/callback"

ALLOWED_USERS = {
    u.strip().lower() for u in os.environ["ALLOWED_USERS"].split(",") if u.strip()
}
if not ALLOWED_USERS:
    raise RuntimeError("ALLOWED_USERS está vazio: defina os e-mails autorizados a usar o servidor.")

DB_PATH = os.environ.get("DB_PATH", "data/m365.db")
TIMEZONE = os.environ.get("TIMEZONE", "America/Sao_Paulo")

MCP_SCOPE = "m365"
