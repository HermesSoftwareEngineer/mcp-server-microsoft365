"""Contatos do Outlook via Microsoft Graph API."""

from mcp.server.mcpserver import MCPServer

import graph


def register(mcp: MCPServer) -> None:

    @mcp.tool()
    async def contacts_list(limit: int = 50) -> list[dict]:
        """Lista os contatos do Outlook.

        Args:
            limit: Número máximo de contatos. Padrão: 50.
        """
        data = await graph.get(
            "/me/contacts",
            {
                "$top": limit,
                "$select": "id,displayName,emailAddresses,mobilePhone,businessPhones,jobTitle,companyName",
                "$orderby": "displayName",
            },
        )
        return [
            {
                "id": c["id"],
                "name": c.get("displayName"),
                "email": c.get("emailAddresses", [{}])[0].get("address")
                if c.get("emailAddresses")
                else None,
                "mobile": c.get("mobilePhone"),
                "phone": c.get("businessPhones", [None])[0],
                "title": c.get("jobTitle"),
                "company": c.get("companyName"),
            }
            for c in data.get("value", [])
        ]

    @mcp.tool()
    async def contacts_get(contact_id: str) -> dict:
        """Retorna os detalhes completos de um contato.

        Args:
            contact_id: ID do contato retornado por contacts_list.
        """
        c = await graph.get(f"/me/contacts/{contact_id}")
        return {
            "id": c["id"],
            "name": c.get("displayName"),
            "firstName": c.get("givenName"),
            "lastName": c.get("surname"),
            "emails": c.get("emailAddresses", []),
            "mobile": c.get("mobilePhone"),
            "businessPhones": c.get("businessPhones", []),
            "title": c.get("jobTitle"),
            "company": c.get("companyName"),
            "department": c.get("department"),
            "notes": c.get("personalNotes"),
        }

    @mcp.tool()
    async def contacts_create(
        name: str,
        email: str | None = None,
        mobile: str | None = None,
        phone: str | None = None,
        title: str | None = None,
        company: str | None = None,
        first_name: str | None = None,
        last_name: str | None = None,
    ) -> dict:
        """Cria um novo contato no Outlook.

        Args:
            name: Nome de exibição do contato.
            email: Endereço de e-mail (opcional).
            mobile: Telefone celular (opcional).
            phone: Telefone comercial (opcional).
            title: Cargo (opcional).
            company: Empresa (opcional).
            first_name: Primeiro nome (opcional).
            last_name: Sobrenome (opcional).
        """
        contact: dict = {"displayName": name}
        if first_name:
            contact["givenName"] = first_name
        if last_name:
            contact["surname"] = last_name
        if email:
            contact["emailAddresses"] = [{"address": email, "name": name}]
        if mobile:
            contact["mobilePhone"] = mobile
        if phone:
            contact["businessPhones"] = [phone]
        if title:
            contact["jobTitle"] = title
        if company:
            contact["companyName"] = company

        result = await graph.post("/me/contacts", contact)
        return {"id": result["id"], "name": result.get("displayName")}

    @mcp.tool()
    async def contacts_update(
        contact_id: str,
        name: str | None = None,
        email: str | None = None,
        mobile: str | None = None,
        phone: str | None = None,
        title: str | None = None,
        company: str | None = None,
    ) -> dict:
        """Atualiza os dados de um contato.

        Args:
            contact_id: ID do contato.
            name: Novo nome de exibição (opcional).
            email: Novo e-mail (opcional).
            mobile: Novo celular (opcional).
            phone: Novo telefone comercial (opcional).
            title: Novo cargo (opcional).
            company: Nova empresa (opcional).
        """
        updates: dict = {}
        if name is not None:
            updates["displayName"] = name
        if email is not None:
            updates["emailAddresses"] = [{"address": email}]
        if mobile is not None:
            updates["mobilePhone"] = mobile
        if phone is not None:
            updates["businessPhones"] = [phone]
        if title is not None:
            updates["jobTitle"] = title
        if company is not None:
            updates["companyName"] = company

        await graph.patch(f"/me/contacts/{contact_id}", updates)
        return {"id": contact_id, "updated": list(updates.keys())}

    @mcp.tool()
    async def contacts_delete(contact_id: str) -> str:
        """Deleta um contato.

        Args:
            contact_id: ID do contato.
        """
        await graph.delete(f"/me/contacts/{contact_id}")
        return f"Contato '{contact_id}' deletado."
