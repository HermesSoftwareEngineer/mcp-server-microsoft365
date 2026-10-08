"""Cliente HTTP base para a Microsoft Graph API."""

from typing import Any

import httpx

from auth import get_access_token

BASE_URL = "https://graph.microsoft.com/v1.0"


async def _headers() -> dict:
    token = await get_access_token()
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


async def get(path: str, params: dict | None = None) -> Any:
    async with httpx.AsyncClient() as client:
        resp = await client.get(
            BASE_URL + path,
            headers=await _headers(),
            params=params,
            timeout=30,
        )
        _raise(resp)
        return resp.json()


async def get_all(path: str, params: dict | None = None) -> list:
    """GET com paginação automática via @odata.nextLink."""
    items: list = []
    url: str | None = BASE_URL + path
    p = params
    async with httpx.AsyncClient() as client:
        while url:
            resp = await client.get(
                url,
                headers=await _headers(),
                params=p,
                timeout=30,
            )
            _raise(resp)
            data = resp.json()
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            p = None  # nextLink já carrega os parâmetros
    return items


async def post(path: str, body: dict) -> Any:
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            BASE_URL + path,
            headers=await _headers(),
            json=body,
            timeout=30,
        )
        _raise(resp)
        if resp.status_code == 204:
            return {}
        return resp.json()


async def patch(path: str, body: dict) -> Any:
    async with httpx.AsyncClient() as client:
        resp = await client.patch(
            BASE_URL + path,
            headers=await _headers(),
            json=body,
            timeout=30,
        )
        _raise(resp)
        if resp.status_code == 204:
            return {}
        return resp.json()


async def delete(path: str) -> None:
    async with httpx.AsyncClient() as client:
        resp = await client.delete(
            BASE_URL + path,
            headers=await _headers(),
            timeout=30,
        )
        _raise(resp)


async def put_bytes(path: str, data: bytes, content_type: str = "application/octet-stream") -> Any:
    headers = await _headers()
    headers["Content-Type"] = content_type
    async with httpx.AsyncClient() as client:
        resp = await client.put(BASE_URL + path, headers=headers, content=data, timeout=60)
        _raise(resp)
        return resp.json()


def _raise(resp: httpx.Response) -> None:
    if resp.status_code == 404:
        raise ValueError(f"Recurso não encontrado: {resp.url}")
    if resp.status_code == 403:
        raise PermissionError(
            f"Permissão negada. Verifique os scopes do app Azure. URL: {resp.url}"
        )
    if resp.status_code >= 400:
        try:
            detail = resp.json().get("error", {}).get("message", resp.text)
        except Exception:
            detail = resp.text
        raise RuntimeError(f"Graph API {resp.status_code}: {detail}")
