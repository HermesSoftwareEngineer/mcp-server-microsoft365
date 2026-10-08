"""Outlook — e-mail."""

from mcp.server.mcpserver import MCPServer

import graph
from graph import odata_str


def _summary(m: dict) -> dict:
    return {
        "id": m["id"],
        "subject": m.get("subject"),
        "from": m.get("from", {}).get("emailAddress", {}).get("address"),
        "received": m.get("receivedDateTime"),
        "isRead": m.get("isRead"),
        "preview": m.get("bodyPreview"),
    }


def _recipients(addresses: list[str]) -> list[dict]:
    return [{"emailAddress": {"address": a}} for a in addresses]


def register(mcp: MCPServer) -> None:

    @mcp.tool()
    async def mail_list(
        folder: str = "inbox",
        from_address: str | None = None,
        subject: str | None = None,
        since: str | None = None,
        unread_only: bool = False,
        limit: int = 20,
    ) -> list[dict]:
        """Lista e-mails de uma pasta, do mais recente para o mais antigo.

        Args:
            folder: Pasta: 'inbox', 'sentitems', 'drafts', 'deleteditems', 'archive' ou ID de pasta.
            from_address: Filtra pelo e-mail do remetente.
            subject: Filtra por trecho do assunto.
            since: Data mínima de recebimento, ex: '2025-01-01'.
            unread_only: Só e-mails não lidos.
            limit: Máximo de e-mails (1-100). Padrão: 20.
        """
        # A Graph exige que a propriedade do $orderby apareça primeiro no $filter.
        filters = [f"receivedDateTime ge {since or '1900-01-01'}T00:00:00Z"]
        if from_address:
            filters.append(f"from/emailAddress/address eq '{odata_str(from_address)}'")
        if subject:
            filters.append(f"contains(subject, '{odata_str(subject)}')")
        if unread_only:
            filters.append("isRead eq false")

        data = await graph.get(
            f"/me/mailFolders/{folder}/messages",
            {
                "$top": max(1, min(limit, 100)),
                "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview",
                "$filter": " and ".join(filters),
                "$orderby": "receivedDateTime desc",
            },
        )
        return [_summary(m) for m in data.get("value", [])]

    @mcp.tool()
    async def mail_get(message_id: str) -> dict:
        """Lê um e-mail completo, com corpo em texto e lista de anexos.

        Args:
            message_id: ID da mensagem (de mail_list ou mail_search).
        """
        m = await graph.get(
            f"/me/messages/{message_id}",
            {
                "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments",
                "$expand": "attachments($select=name,contentType,size)",
            },
            headers={"Prefer": 'outlook.body-content-type="text"'},
        )
        return {
            "id": m["id"],
            "subject": m.get("subject"),
            "from": m.get("from", {}).get("emailAddress", {}).get("address"),
            "to": [r["emailAddress"]["address"] for r in m.get("toRecipients", [])],
            "cc": [r["emailAddress"]["address"] for r in m.get("ccRecipients", [])],
            "received": m.get("receivedDateTime"),
            "body": m.get("body", {}).get("content"),
            "attachments": [
                {"name": a.get("name"), "type": a.get("contentType"), "size": a.get("size")}
                for a in m.get("attachments", [])
            ],
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
            to: Destinatários, ex: ['fulano@empresa.com'].
            subject: Assunto.
            body: Corpo.
            cc: Destinatários em cópia (opcional).
            html: True se o corpo for HTML. Padrão: texto simples.
        """
        message: dict = {
            "subject": subject,
            "body": {"contentType": "HTML" if html else "Text", "content": body},
            "toRecipients": _recipients(to),
        }
        if cc:
            message["ccRecipients"] = _recipients(cc)
        await graph.post("/me/sendMail", {"message": message, "saveToSentItems": True})
        return f"E-mail enviado para {', '.join(to)}."

    @mcp.tool()
    async def mail_reply(message_id: str, body: str, reply_all: bool = False) -> str:
        """Responde a um e-mail.

        Args:
            message_id: ID da mensagem.
            body: Texto da resposta.
            reply_all: True para responder a todos.
        """
        action = "replyAll" if reply_all else "reply"
        await graph.post(f"/me/messages/{message_id}/{action}", {"comment": body})
        return "Resposta enviada."

    @mcp.tool()
    async def mail_move(message_id: str, destination_folder: str) -> str:
        """Move um e-mail para outra pasta.

        Args:
            message_id: ID da mensagem.
            destination_folder: 'inbox', 'archive', 'deleteditems', 'junkemail' ou ID de pasta.
        """
        await graph.post(f"/me/messages/{message_id}/move", {"destinationId": destination_folder})
        return f"Mensagem movida para '{destination_folder}'."

    @mcp.tool()
    async def mail_delete(message_id: str) -> str:
        """Exclui um e-mail (vai para 'Itens Excluídos').

        Args:
            message_id: ID da mensagem.
        """
        await graph.delete(f"/me/messages/{message_id}")
        return "E-mail excluído."

    @mcp.tool()
    async def mail_search(query: str, limit: int = 20) -> list[dict]:
        """Busca e-mails em todas as pastas por texto (assunto, remetente, corpo).

        Args:
            query: Texto a buscar.
            limit: Máximo de resultados (1-100). Padrão: 20.
        """
        data = await graph.get(
            "/me/messages",
            {
                "$search": f'"{query.replace(chr(34), "")}"',
                "$top": max(1, min(limit, 100)),
                "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview",
            },
        )
        return [_summary(m) for m in data.get("value", [])]
