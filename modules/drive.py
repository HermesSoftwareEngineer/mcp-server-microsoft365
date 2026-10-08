"""OneDrive / SharePoint — ferramentas MCP."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

import graph


def register(mcp: FastMCP) -> None:

    @mcp.tool()
    async def drive_list_files(folder_path: str = "/") -> list[dict]:
        """Lista arquivos e pastas em um caminho do OneDrive.

        Args:
            folder_path: Caminho relativo à raiz, ex: '/Documentos/Projetos'. Use '/' para a raiz.
        """
        if folder_path.strip("/") == "":
            path = "/me/drive/root/children"
        else:
            encoded = folder_path.strip("/")
            path = f"/me/drive/root:/{encoded}:/children"

        items = await graph.get_all(
            path,
            params={"$select": "id,name,size,lastModifiedDateTime,file,folder"},
        )
        return [
            {
                "id": i["id"],
                "name": i["name"],
                "type": "pasta" if "folder" in i else "arquivo",
                "size": i.get("size"),
                "modified": i.get("lastModifiedDateTime"),
            }
            for i in items
        ]

    @mcp.tool()
    async def drive_get_file(item_id: str) -> str:
        """Lê o conteúdo de um arquivo de texto, JSON ou CSV do OneDrive.

        Args:
            item_id: ID do item retornado por drive_list_files ou drive_search.
        """
        import httpx
        from auth import get_access_token

        token = await get_access_token()
        async with httpx.AsyncClient() as client:
            resp = await client.get(
                f"https://graph.microsoft.com/v1.0/me/drive/items/{item_id}/content",
                headers={"Authorization": f"Bearer {token}"},
                follow_redirects=True,
                timeout=60,
            )
            if resp.status_code == 404:
                raise ValueError(f"Arquivo '{item_id}' não encontrado.")
            resp.raise_for_status()
            return resp.text

    @mcp.tool()
    async def drive_upload_file(folder_path: str, filename: str, content: str) -> dict:
        """Cria ou substitui um arquivo no OneDrive.

        Args:
            folder_path: Caminho da pasta de destino, ex: '/Documentos'. Use '' para a raiz.
            filename: Nome do arquivo, ex: 'relatorio.txt'.
            content: Conteúdo em texto puro (UTF-8).
        """
        path_clean = folder_path.strip("/")
        if path_clean:
            api_path = f"/me/drive/root:/{path_clean}/{filename}:/content"
        else:
            api_path = f"/me/drive/root:/{filename}:/content"

        result = await graph.put_bytes(api_path, content.encode("utf-8"), "text/plain")
        return {"id": result["id"], "name": result["name"], "webUrl": result.get("webUrl")}

    @mcp.tool()
    async def drive_delete_file(item_id: str) -> str:
        """Deleta um arquivo ou pasta do OneDrive.

        Args:
            item_id: ID do item a deletar.
        """
        await graph.delete(f"/me/drive/items/{item_id}")
        return f"Item '{item_id}' deletado com sucesso."

    @mcp.tool()
    async def drive_move_file(
        item_id: str, destination_path: str, new_name: str | None = None
    ) -> dict:
        """Move ou renomeia um arquivo/pasta no OneDrive.

        Args:
            item_id: ID do item a mover.
            destination_path: Caminho da pasta de destino, ex: '/Arquivos/2025'.
            new_name: Novo nome do arquivo (opcional).
        """
        dest = await graph.get(f"/me/drive/root:/{destination_path.strip('/')}")
        body: dict = {"parentReference": {"id": dest["id"]}}
        if new_name:
            body["name"] = new_name
        result = await graph.patch(f"/me/drive/items/{item_id}", body)
        return {"id": result["id"], "name": result["name"]}

    @mcp.tool()
    async def drive_search(query: str) -> list[dict]:
        """Busca arquivos no OneDrive por nome ou conteúdo.

        Args:
            query: Texto a buscar.
        """
        data = await graph.get(f"/me/drive/root/search(q='{query}')")
        return [
            {
                "id": i["id"],
                "name": i["name"],
                "parentPath": i.get("parentReference", {}).get("path", ""),
                "modified": i.get("lastModifiedDateTime"),
            }
            for i in data.get("value", [])
        ]
