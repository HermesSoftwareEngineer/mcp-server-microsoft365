"""MCP Server Microsoft 365 — servidor HTTP (Streamable HTTP) protegido por OAuth 2.1.

Endpoint MCP: {PUBLIC_URL}/mcp
"""

import logging

from mcp.server.auth.settings import AuthSettings, ClientRegistrationOptions, RevocationOptions
from mcp.server.mcpserver import MCPServer
from mcp.server.transport_security import TransportSecuritySettings
from starlette.requests import Request
from starlette.responses import PlainTextResponse, Response

import config
from modules import calendar, contacts, drive, excel, mail, tasks, teams
from provider import M365OAuthProvider

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")

provider = M365OAuthProvider()

mcp = MCPServer(
    "microsoft365",
    instructions=(
        "Acesso ao Microsoft 365 do usuário: e-mails e agenda do Outlook, OneDrive/SharePoint, "
        "planilhas Excel, contatos, Microsoft To Do e Teams. Todas as ações são feitas em nome "
        f"do usuário autenticado. Datas sem fuso horário são interpretadas em {config.TIMEZONE}."
    ),
    auth_server_provider=provider,
    auth=AuthSettings(
        issuer_url=config.PUBLIC_URL,
        resource_server_url=config.RESOURCE_URL,
        validate_token_resource=True,
        required_scopes=[config.MCP_SCOPE],
        client_registration_options=ClientRegistrationOptions(
            enabled=True,
            valid_scopes=[config.MCP_SCOPE],
            default_scopes=[config.MCP_SCOPE],
        ),
        revocation_options=RevocationOptions(enabled=True),
    ),
)


@mcp.custom_route("/auth/microsoft/callback", methods=["GET"])
async def microsoft_callback(request: Request) -> Response:
    return await provider.handle_microsoft_callback(request)


@mcp.custom_route("/health", methods=["GET"])
async def health(_: Request) -> Response:
    return PlainTextResponse("ok")


for module in (drive, excel, mail, calendar, contacts, tasks, teams):
    module.register(mcp)


if __name__ == "__main__":
    mcp.run(
        transport="streamable-http",
        host=config.HOST,
        port=config.PORT,
        streamable_http_path=config.MCP_PATH,
        stateless_http=True,
        transport_security=TransportSecuritySettings(
            allowed_hosts=[config.PUBLIC_HOST, "127.0.0.1:*", "localhost:*"],
            allowed_origins=[config.PUBLIC_URL, "https://claude.ai"],
        ),
    )
