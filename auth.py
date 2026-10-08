"""OAuth2 Delegated flow para Microsoft Graph API.

Uso inicial:
    python auth.py --setup

Após isso o refresh token fica salvo em TOKEN_FILE e o servidor
renova o access token automaticamente.
"""

import argparse
import asyncio
import json
import os
import time
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from threading import Thread

import httpx
from dotenv import load_dotenv

load_dotenv()

CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
TENANT_ID = os.environ["AZURE_TENANT_ID"]
TOKEN_FILE = os.environ.get("TOKEN_FILE", ".tokens.json")

REDIRECT_URI = "http://localhost:8400/callback"

SCOPES = [
    "Files.ReadWrite",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Contacts.ReadWrite",
    "Tasks.ReadWrite",
    "Chat.ReadWrite",
    "Sites.ReadWrite.All",
    "offline_access",
]

AUTH_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

_auth_code: str | None = None


class _CallbackHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        global _auth_code
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        if "code" in params:
            _auth_code = params["code"][0]
            body = b"<html><body><h2>Autorizado! Pode fechar esta aba.</h2></body></html>"
        else:
            body = b"<html><body><h2>Erro na autorizacao.</h2></body></html>"
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args):
        pass


def _load_tokens() -> dict:
    try:
        with open(TOKEN_FILE) as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def _save_tokens(tokens: dict) -> None:
    with open(TOKEN_FILE, "w") as f:
        json.dump(tokens, f, indent=2)


def _exchange_code(code: str) -> dict:
    resp = httpx.post(
        TOKEN_URL,
        data={
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "code": code,
            "redirect_uri": REDIRECT_URI,
            "grant_type": "authorization_code",
            "scope": " ".join(SCOPES),
        },
    )
    resp.raise_for_status()
    data = resp.json()
    data["expires_at"] = time.time() + data.get("expires_in", 3600) - 60
    return data


def _refresh(refresh_token: str) -> dict:
    resp = httpx.post(
        TOKEN_URL,
        data={
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "refresh_token": refresh_token,
            "grant_type": "refresh_token",
            "scope": " ".join(SCOPES),
        },
    )
    resp.raise_for_status()
    data = resp.json()
    data["expires_at"] = time.time() + data.get("expires_in", 3600) - 60
    return data


async def get_access_token() -> str:
    """Retorna um access token válido, fazendo refresh se necessário."""
    tokens = _load_tokens()
    if not tokens:
        raise RuntimeError(
            "Tokens não encontrados. Execute: python auth.py --setup"
        )

    if time.time() >= tokens.get("expires_at", 0):
        refresh_token = tokens.get("refresh_token")
        if not refresh_token:
            raise RuntimeError(
                "Refresh token ausente. Execute: python auth.py --setup"
            )
        tokens = await asyncio.to_thread(_refresh, refresh_token)
        _save_tokens(tokens)

    return tokens["access_token"]


def setup():
    """Inicia o fluxo de autorização e salva os tokens."""
    global _auth_code

    params = {
        "client_id": CLIENT_ID,
        "response_type": "code",
        "redirect_uri": REDIRECT_URI,
        "scope": " ".join(SCOPES),
        "response_mode": "query",
    }
    url = AUTH_URL + "?" + urllib.parse.urlencode(params)

    server = HTTPServer(("localhost", 8400), _CallbackHandler)
    thread = Thread(target=server.handle_request, daemon=True)
    thread.start()

    print(f"\nAbrindo navegador para autorização...\n{url}\n")
    webbrowser.open(url)

    thread.join(timeout=120)
    server.server_close()

    if not _auth_code:
        raise RuntimeError("Código de autorização não recebido. Tente novamente.")

    print("Código recebido. Trocando por tokens...")
    tokens = _exchange_code(_auth_code)
    _save_tokens(tokens)
    print(f"Tokens salvos em '{TOKEN_FILE}'. Setup concluído.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--setup", action="store_true", help="Executar fluxo OAuth inicial")
    args = parser.parse_args()
    if args.setup:
        setup()
    else:
        parser.print_help()
