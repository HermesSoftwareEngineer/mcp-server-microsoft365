"""Microsoft Teams via Microsoft Graph API."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

import graph


def register(mcp: FastMCP) -> None:

    @mcp.tool()
    async def teams_list_chats(limit: int = 20) -> list[dict]:
        """Lista as conversas (chats) recentes do Teams.

        Args:
            limit: Número máximo de conversas. Padrão: 20.
        """
        data = await graph.get(
            "/me/chats",
            {
                "$top": limit,
                "$select": "id,chatType,topic,lastUpdatedDateTime",
                "$expand": "members($select=displayName,email)",
                "$orderby": "lastUpdatedDateTime desc",
            },
        )
        chats = []
        for c in data.get("value", []):
            members = [
                m.get("displayName") or m.get("email", "")
                for m in c.get("members", [])
            ]
            chats.append({
                "id": c["id"],
                "type": c.get("chatType"),
                "topic": c.get("topic"),
                "members": members,
                "lastUpdated": c.get("lastUpdatedDateTime"),
            })
        return chats

    @mcp.tool()
    async def teams_send_message(
        body: str,
        chat_id: str | None = None,
        team_id: str | None = None,
        channel_id: str | None = None,
        html: bool = False,
    ) -> str:
        """Envia uma mensagem em um chat ou canal do Teams.

        Forneça 'chat_id' para chats diretos/grupos, ou 'team_id' + 'channel_id' para canais.

        Args:
            body: Corpo da mensagem.
            chat_id: ID do chat (retornado por teams_list_chats). Use este OU team_id+channel_id.
            team_id: ID do time. Use junto com channel_id.
            channel_id: ID do canal.
            html: Se True, o corpo é HTML. Padrão: False.
        """
        content_type = "html" if html else "text"
        message_body = {"body": {"contentType": content_type, "content": body}}

        if chat_id:
            await graph.post(f"/chats/{chat_id}/messages", message_body)
            return f"Mensagem enviada no chat '{chat_id}'."
        elif team_id and channel_id:
            await graph.post(
                f"/teams/{team_id}/channels/{channel_id}/messages", message_body
            )
            return f"Mensagem enviada no canal '{channel_id}' do time '{team_id}'."
        else:
            raise ValueError("Forneça 'chat_id' ou 'team_id' + 'channel_id'.")
