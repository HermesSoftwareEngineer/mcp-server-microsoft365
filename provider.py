"""Servidor de autorização OAuth 2.1 que o Claude usa para se conectar.

Fluxo: Claude -> /authorize (aqui) -> login Microsoft -> /auth/microsoft/callback (aqui)
-> código emitido por este servidor -> Claude troca em /token (PKCE validado pelo SDK).
"""

import html
import logging
import secrets
import time

from mcp.server.auth.provider import (
    AccessToken,
    AuthorizationCode,
    AuthorizationParams,
    AuthorizeError,
    RefreshToken,
    TokenError,
    construct_redirect_uri,
)
from mcp.shared.auth import OAuthClientInformationFull, OAuthToken
from starlette.requests import Request
from starlette.responses import HTMLResponse, RedirectResponse, Response

import config
import microsoft
import store

log = logging.getLogger(__name__)

PENDING_TTL = 600
CODE_TTL = 300
ACCESS_TTL = 3600
REFRESH_TTL = 90 * 24 * 3600


class M365OAuthProvider:
    async def get_client(self, client_id: str) -> OAuthClientInformationFull | None:
        data = store.get("client", client_id)
        return OAuthClientInformationFull.model_validate(data) if data else None

    async def register_client(self, client_info: OAuthClientInformationFull) -> None:
        store.put("client", client_info.client_id, client_info.model_dump(mode="json"))

    async def authorize(self, client: OAuthClientInformationFull, params: AuthorizationParams) -> str:
        if params.resource and params.resource.rstrip("/") != config.RESOURCE_URL:
            raise AuthorizeError("invalid_target", f"Recurso desconhecido: {params.resource}")
        state = secrets.token_urlsafe(32)
        store.put(
            "pending",
            state,
            {"client_id": client.client_id, "params": params.model_dump(mode="json")},
            expires_at=time.time() + PENDING_TTL,
        )
        return microsoft.authorize_url(state)

    async def handle_microsoft_callback(self, request: Request) -> Response:
        state = request.query_params.get("state", "")
        pending = store.pop("pending", state)
        if pending is None:
            return _error_page("Sessão de login inválida ou expirada. Tente conectar novamente pelo Claude.")
        params = AuthorizationParams.model_validate(pending["params"])

        def back_to_client(**query: str | None) -> RedirectResponse:
            return RedirectResponse(
                construct_redirect_uri(str(params.redirect_uri), state=params.state, **query),
                status_code=302,
            )

        if "error" in request.query_params:
            return back_to_client(
                error="access_denied", error_description=request.query_params.get("error_description")
            )

        try:
            tokens = await microsoft.exchange_code(request.query_params["code"])
            profile = await microsoft.fetch_profile(tokens["access_token"])
        except Exception:
            log.exception("Falha ao concluir login Microsoft")
            return back_to_client(error="server_error", error_description="Falha no login Microsoft")

        if not microsoft.is_allowed(profile):
            log.warning("Login recusado para %s", profile.get("userPrincipalName"))
            return back_to_client(
                error="access_denied", error_description="Usuário não autorizado neste servidor"
            )

        subject = profile["id"]
        microsoft.save_tokens(subject, tokens, profile)

        code = secrets.token_urlsafe(32)
        auth_code = AuthorizationCode(
            code=code,
            scopes=params.scopes or [config.MCP_SCOPE],
            expires_at=time.time() + CODE_TTL,
            client_id=pending["client_id"],
            code_challenge=params.code_challenge,
            redirect_uri=params.redirect_uri,
            redirect_uri_provided_explicitly=params.redirect_uri_provided_explicitly,
            resource=config.RESOURCE_URL,
            subject=subject,
        )
        store.put(
            "auth_code",
            store.hash_secret(code),
            auth_code.model_dump(mode="json", exclude={"code"}),
            expires_at=auth_code.expires_at,
        )
        return back_to_client(code=code)

    async def load_authorization_code(
        self, client: OAuthClientInformationFull, authorization_code: str
    ) -> AuthorizationCode | None:
        data = store.get("auth_code", store.hash_secret(authorization_code))
        if data is None or data["client_id"] != client.client_id:
            return None
        return AuthorizationCode(code=authorization_code, **data)

    async def exchange_authorization_code(
        self, client: OAuthClientInformationFull, authorization_code: AuthorizationCode
    ) -> OAuthToken:
        store.delete("auth_code", store.hash_secret(authorization_code.code))
        return self._issue_tokens(client.client_id, authorization_code.subject, authorization_code.scopes)

    async def load_refresh_token(
        self, client: OAuthClientInformationFull, refresh_token: str
    ) -> RefreshToken | None:
        data = store.get("refresh_token", store.hash_secret(refresh_token))
        if data is None or data["client_id"] != client.client_id:
            return None
        return RefreshToken(token=refresh_token, **data)

    async def exchange_refresh_token(
        self, client: OAuthClientInformationFull, refresh_token: RefreshToken, scopes: list[str]
    ) -> OAuthToken:
        store.delete("refresh_token", store.hash_secret(refresh_token.token))
        return self._issue_tokens(client.client_id, refresh_token.subject, scopes or refresh_token.scopes)

    async def load_access_token(self, token: str) -> AccessToken | None:
        data = store.get("access_token", store.hash_secret(token))
        if data is None:
            return None
        account = store.get("ms_token", data["subject"])
        if account is None or not microsoft.is_allowed(account):
            return None
        return AccessToken(token=token, **data)

    async def revoke_token(self, token: AccessToken | RefreshToken) -> None:
        if token.subject:
            store.delete_owner_tokens(token.subject, token.client_id)

    def _issue_tokens(self, client_id: str, subject: str | None, scopes: list[str]) -> OAuthToken:
        if subject is None:
            raise TokenError("invalid_grant", "Concessão sem usuário associado")
        now = time.time()
        access = secrets.token_urlsafe(32)
        refresh = secrets.token_urlsafe(48)
        common = {"client_id": client_id, "scopes": scopes, "resource": config.RESOURCE_URL, "subject": subject}
        store.put(
            "access_token",
            store.hash_secret(access),
            {**common, "expires_at": int(now + ACCESS_TTL)},
            expires_at=now + ACCESS_TTL,
            client_id=client_id,
            subject=subject,
        )
        store.put(
            "refresh_token",
            store.hash_secret(refresh),
            {**common, "expires_at": int(now + REFRESH_TTL)},
            expires_at=now + REFRESH_TTL,
            client_id=client_id,
            subject=subject,
        )
        store.purge_expired()
        return OAuthToken(
            access_token=access,
            token_type="Bearer",
            expires_in=ACCESS_TTL,
            scope=" ".join(scopes),
            refresh_token=refresh,
        )


def _error_page(message: str) -> HTMLResponse:
    return HTMLResponse(
        f"<!doctype html><meta charset='utf-8'><title>Erro</title>"
        f"<p style='font-family:sans-serif'>{html.escape(message)}</p>",
        status_code=400,
    )
