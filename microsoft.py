"""OAuth2 com a Microsoft (Entra ID) e gestão dos tokens da Graph API por usuário."""

import asyncio
import time
from urllib.parse import urlencode

import httpx
from mcp.server.mcpserver.exceptions import ToolError

import config
import store

AUTHORITY = f"https://login.microsoftonline.com/{config.AZURE_TENANT_ID}/oauth2/v2.0"

SCOPES = [
    "offline_access",
    "User.Read",
    "Files.ReadWrite",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Contacts.ReadWrite",
    "Tasks.ReadWrite",
    "Chat.ReadWrite",
    "ChannelMessage.Send",
    "Sites.ReadWrite.All",
]

_refresh_locks: dict[str, asyncio.Lock] = {}


class MicrosoftReauthRequired(ToolError):
    pass


def authorize_url(state: str) -> str:
    params = {
        "client_id": config.AZURE_CLIENT_ID,
        "response_type": "code",
        "redirect_uri": config.MS_REDIRECT_URI,
        "response_mode": "query",
        "scope": " ".join(SCOPES),
        "state": state,
        "prompt": "select_account",
    }
    return f"{AUTHORITY}/authorize?{urlencode(params)}"


async def _token_request(data: dict[str, str]) -> dict:
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(
            f"{AUTHORITY}/token",
            data={
                "client_id": config.AZURE_CLIENT_ID,
                "client_secret": config.AZURE_CLIENT_SECRET,
                "scope": " ".join(SCOPES),
                **data,
            },
        )
    body = resp.json()
    if resp.status_code != 200:
        raise MicrosoftReauthRequired(
            f"Microsoft recusou o token ({body.get('error')}): {body.get('error_description', '')[:300]}"
        )
    return body


async def exchange_code(code: str) -> dict:
    return await _token_request(
        {"grant_type": "authorization_code", "code": code, "redirect_uri": config.MS_REDIRECT_URI}
    )


async def fetch_profile(access_token: str) -> dict:
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.get(
            "https://graph.microsoft.com/v1.0/me?$select=id,userPrincipalName,mail,displayName",
            headers={"Authorization": f"Bearer {access_token}"},
        )
    resp.raise_for_status()
    return resp.json()


def is_allowed(account: dict) -> bool:
    emails = {(account.get(k) or "").lower() for k in ("userPrincipalName", "upn", "mail")}
    return bool(emails & config.ALLOWED_USERS)


def save_tokens(subject: str, tokens: dict, profile: dict | None = None) -> None:
    current = store.get("ms_token", subject) or {}
    store.put(
        "ms_token",
        subject,
        {
            "access_token": tokens["access_token"],
            # A Microsoft nem sempre devolve um refresh token novo; mantém o anterior nesse caso.
            "refresh_token": tokens.get("refresh_token") or current.get("refresh_token"),
            "expires_at": time.time() + int(tokens.get("expires_in", 3600)) - 120,
            "upn": (profile or {}).get("userPrincipalName") or current.get("upn"),
            "mail": (profile or {}).get("mail") or current.get("mail"),
        },
        subject=subject,
    )


async def get_valid_token(subject: str, force_refresh: bool = False) -> str:
    lock = _refresh_locks.setdefault(subject, asyncio.Lock())
    async with lock:
        tokens = store.get("ms_token", subject)
        if tokens is None:
            store.delete_owner_tokens(subject)
            raise MicrosoftReauthRequired(
                "Conta Microsoft não vinculada. Reconecte o conector no Claude."
            )
        if not force_refresh and time.time() < tokens["expires_at"]:
            return tokens["access_token"]
        try:
            refreshed = await _token_request(
                {"grant_type": "refresh_token", "refresh_token": tokens["refresh_token"]}
            )
        except MicrosoftReauthRequired:
            # Sessão Microsoft revogada/expirada: invalida os tokens do MCP para o Claude refazer o login.
            store.delete("ms_token", subject)
            store.delete_owner_tokens(subject)
            raise MicrosoftReauthRequired(
                "A sessão Microsoft expirou ou foi revogada. Reconecte o conector no Claude."
            )
        save_tokens(subject, refreshed)
        return refreshed["access_token"]
