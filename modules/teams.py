"""Microsoft Teams."""

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

import graph


def register(mcp: MCPServer) -> None:

    @mcp.tool()
    async def teams_list_chats(limit: int = 20) -> list[dict]:
        """Lista os chats do Teams mais recentes, com participantes e prévia da última mensagem.

        Args:
            limit: Máximo de chats (1-50). Padrão: 20.
        """
        data = await graph.get(
            "/me/chats",
            {
                "$top": max(1, min(limit, 50)),
                "$expand": "members,lastMessagePreview",
                "$orderby": "lastMessagePreview/createdDateTime desc",
            },
        )
        return [
            {
                "id": c["id"],
                "type": c.get("chatType"),
                "topic": c.get("topic"),
                "members": [m.get("displayName") for m in c.get("members", [])],
                "lastMessage": (c.get("lastMessagePreview") or {}).get("body", {}).get("content"),
                "lastMessageAt": (c.get("lastMessagePreview") or {}).get("createdDateTime"),
                "webUrl": c.get("webUrl"),
            }
            for c in data.get("value", [])
        ]

    @mcp.tool()
    async def teams_send_message(
        body: str,
        chat_id: str | None = None,
        team_id: str | None = None,
        channel_id: str | None = None,
        html: bool = False,
    ) -> str:
        """Envia uma mensagem em um chat ou em um canal do Teams.

        Informe `chat_id` (chat individual/grupo) OU `team_id` + `channel_id` (canal).

        Args:
            body: Texto da mensagem.
            chat_id: ID do chat (de teams_list_chats).
            team_id: ID do time.
            channel_id: ID do canal.
            html: True se o corpo for HTML.
        """
        message = {"body": {"contentType": "html" if html else "text", "content": body}}
        if chat_id:
            await graph.post(f"/chats/{chat_id}/messages", message)
            return "Mensagem enviada no chat."
        if team_id and channel_id:
            await graph.post(f"/teams/{team_id}/channels/{channel_id}/messages", message)
            return "Mensagem enviada no canal."
        raise ToolError("Informe 'chat_id' ou 'team_id' + 'channel_id'.")
