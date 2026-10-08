"""Outlook Mail via Microsoft Graph API."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

import graph


def register(mcp: FastMCP) -> None:

    @mcp.tool()
    async def mail_list(
        folder: str = "inbox",
        from_address: str | None = None,
        subject: str | None = None,
        since: str | None = None,
        limit: int = 20,
    ) -> list[dict]:
        """Lista e-mails com filtros opcionais.

        Args:
            folder: Pasta do Outlook: 'inbox', 'sentitems', 'drafts', 'deleteditems'. Padrão: 'inbox'.
            from_address: Filtrar por endereço do remetente.
            subject: Filtrar por texto no assunto (contém).
            since: Filtrar por data mínima, ex: '2025-01-01'.
            limit: Número máximo de e-mails. Padrão: 20.
        """
        filters = []
        if from_address:
            filters.append(f"from/emailAddress/address eq '{from_address}'")
        if subject:
            filters.append(f"contains(subject, '{subject}')")
        if since:
            filters.append(f"receivedDateTime ge {since}T00:00:00Z")

        params: dict = {
            "$top": limit,
            "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview",
            "$orderby": "receivedDateTime desc",
        }
        if filters:
            params["$filter"] = " and ".join(filters)

        data = await graph.get(f"/me/mailFolders/{folder}/messages", params)
        return [
            {
                "id": m["id"],
                "subject": m.get("subject"),
                "from": m.get("from", {}).get("emailAddress", {}).get("address"),
                "received": m.get("receivedDateTime"),
                "isRead": m.get("isRead"),
                "preview": m.get("bodyPreview"),
            }
            for m in data.get("value", [])
        ]

    @mcp.tool()
    async def mail_get(message_id: str) -> dict:
        """Lê um e-mail completo incluindo o corpo.

        Args:
            message_id: ID da mensagem retornado por mail_list ou mail_search.
        """
        m = await graph.get(
            f"/me/messages/{message_id}",
            {"$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body"},
        )
        return {
            "id": m["id"],
            "subject": m.get("subject"),
            "from": m.get("from", {}).get("emailAddress", {}).get("address"),
            "to": [r["emailAddress"]["address"] for r in m.get("toRecipients", [])],
            "cc": [r["emailAddress"]["address"] for r in m.get("ccRecipients", [])],
            "received": m.get("receivedDateTime"),
            "bodyType": m.get("body", {}).get("contentType"),
            "body": m.get("body", {}).get("content"),
        }

    @mcp.tool()
    async def mail_send(
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        html: bool = False,
    ) -> str:
        """Envia um e-mail.

        Args:
            to: Lista de destinatários, ex: ['fulano@empresa.com'].
            subject: Assunto do e-mail.
            body: Corpo do e-mail.
            cc: Lista de destinatários em cópia (opcional).
            html: Se True, o corpo é HTML. Padrão: False.
        """
        message: dict = {
            "subject": subject,
            "body": {"contentType": "HTML" if html else "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": a}} for a in to],
        }
        if cc:
            message["ccRecipients"] = [{"emailAddress": {"address": a}} for a in cc]

        await graph.post("/me/sendMail", {"message": message, "saveToSentItems": True})
        return f"E-mail enviado para {', '.join(to)}."

    @mcp.tool()
    async def mail_reply(message_id: str, body: str, html: bool = False) -> str:
        """Responde a um e-mail.

        Args:
            message_id: ID da mensagem a responder.
            body: Corpo da resposta.
            html: Se True, o corpo é HTML. Padrão: False.
        """
        await graph.post(
            f"/me/messages/{message_id}/reply",
            {"message": {"body": {"contentType": "HTML" if html else "Text", "content": body}}},
        )
        return "Resposta enviada."

    @mcp.tool()
    async def mail_move(message_id: str, destination_folder: str) -> str:
        """Move um e-mail para outra pasta.

        Args:
            message_id: ID da mensagem.
            destination_folder: Nome da pasta destino, ex: 'deleteditems', 'archive'.
        """
        await graph.post(
            f"/me/messages/{message_id}/move",
            {"destinationId": destination_folder},
        )
        return f"Mensagem movida para '{destination_folder}'."

    @mcp.tool()
    async def mail_delete(message_id: str) -> str:
        """Deleta permanentemente um e-mail.

        Args:
            message_id: ID da mensagem.
        """
        await graph.delete(f"/me/messages/{message_id}")
        return "E-mail deletado."

    @mcp.tool()
    async def mail_search(query: str, limit: int = 20) -> list[dict]:
        """Busca e-mails por texto livre (assunto, remetente, corpo).

        Args:
            query: Texto a buscar.
            limit: Número máximo de resultados. Padrão: 20.
        """
        data = await graph.get(
            "/me/messages",
            {
                "$search": f'"{query}"',
                "$top": limit,
                "$select": "id,subject,from,receivedDateTime,bodyPreview",
            },
        )
        return [
            {
                "id": m["id"],
                "subject": m.get("subject"),
                "from": m.get("from", {}).get("emailAddress", {}).get("address"),
                "received": m.get("receivedDateTime"),
                "preview": m.get("bodyPreview"),
            }
            for m in data.get("value", [])
        ]
