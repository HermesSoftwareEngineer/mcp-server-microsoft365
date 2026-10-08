"""Cliente HTTP da Microsoft Graph API, autenticado como o usuário da requisição MCP atual."""

from typing import Any
from urllib.parse import quote

import httpx
from mcp.server.auth.middleware.auth_context import get_access_token
from mcp.server.mcpserver.exceptions import ToolError

import microsoft

BASE_URL = "https://graph.microsoft.com/v1.0"

_client = httpx.AsyncClient(timeout=60, follow_redirects=True)


def odata_str(value: str) -> str:
    """Escapa um literal de string OData (aspas simples duplicadas)."""
    return value.replace("'", "''")


def drive_path(path: str) -> str:
    return quote(path.strip("/"), safe="/")


def _current_subject() -> str:
    token = get_access_token()
    if token is None or token.subject is None:
        raise ToolError("Requisição sem usuário autenticado.")
    return token.subject


async def _request(method: str, url: str, **kwargs: Any) -> httpx.Response:
    subject = _current_subject()
    if not url.startswith("https://"):
        url = BASE_URL + url
    extra_headers = kwargs.pop("headers", {})

    for attempt in range(2):
        token = await microsoft.get_valid_token(subject, force_refresh=attempt > 0)
        resp = await _client.request(
            method, url, headers={"Authorization": f"Bearer {token}", **extra_headers}, **kwargs
        )
        if resp.status_code != 401:
            break
    _raise(resp)
    return resp


async def get(path: str, params: dict | None = None, headers: dict | None = None) -> Any:
    return (await _request("GET", path, params=params, headers=headers or {})).json()


async def get_all(path: str, params: dict | None = None, headers: dict | None = None) -> list:
    """GET com paginação automática via @odata.nextLink."""
    items: list = []
    url: str | None = path
    while url:
        data = (await _request("GET", url, params=params, headers=headers or {})).json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None  # o nextLink já carrega os parâmetros
    return items


async def get_content(path: str) -> bytes:
    return (await _request("GET", path)).content


async def post(path: str, body: dict, headers: dict | None = None) -> Any:
    resp = await _request("POST", path, json=body, headers=headers or {})
    return resp.json() if resp.content else {}


async def patch(path: str, body: dict) -> Any:
    resp = await _request("PATCH", path, json=body)
    return resp.json() if resp.content else {}


async def delete(path: str) -> None:
    await _request("DELETE", path)


async def put_bytes(path: str, data: bytes, content_type: str = "application/octet-stream") -> Any:
    resp = await _request("PUT", path, content=data, headers={"Content-Type": content_type})
    return resp.json()


def _raise(resp: httpx.Response) -> None:
    if resp.status_code < 400:
        return
    try:
        error = resp.json().get("error", {})
        detail = f"{error.get('code', '')}: {error.get('message', '')}"
    except ValueError:
        detail = resp.text[:500]
    if resp.status_code == 404:
        raise ToolError(f"Recurso não encontrado ({detail})")
    if resp.status_code == 403:
        raise ToolError(
            f"Permissão negada pela Microsoft ({detail}). Verifique se o app no Entra ID "
            "tem a permissão delegada necessária e se o consentimento foi concedido."
        )
    raise ToolError(f"Graph API {resp.status_code} ({detail})")
