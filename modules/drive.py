"""OneDrive / SharePoint."""

from urllib.parse import quote

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

import graph
from graph import drive_path, odata_str


def register(mcp: MCPServer) -> None:

    @mcp.tool()
    async def drive_list_files(folder_path: str = "/") -> list[dict]:
        """Lista arquivos e pastas em um caminho do OneDrive.

        Args:
            folder_path: Caminho relativo à raiz, ex: '/Documentos/Projetos'. Use '/' para a raiz.
        """
        path = drive_path(folder_path)
        url = "/me/drive/root/children" if not path else f"/me/drive/root:/{path}:/children"
        items = await graph.get_all(
            url, params={"$select": "id,name,size,lastModifiedDateTime,file,folder,webUrl"}
        )
        return [
            {
                "id": i["id"],
                "name": i["name"],
                "type": "pasta" if "folder" in i else "arquivo",
                "size": i.get("size"),
                "modified": i.get("lastModifiedDateTime"),
                "webUrl": i.get("webUrl"),
            }
            for i in items
        ]

    @mcp.tool()
    async def drive_get_file(item_id: str) -> str:
        """Lê o conteúdo de um arquivo de texto (txt, md, JSON, CSV etc.) do OneDrive.

        Args:
            item_id: ID do item retornado por drive_list_files ou drive_search.
        """
        content = await graph.get_content(f"/me/drive/items/{item_id}/content")
        try:
            return content.decode("utf-8")
        except UnicodeDecodeError:
            raise ToolError(
                "O arquivo não é texto UTF-8 (provavelmente binário). Para .xlsx use excel_read_sheet."
            )

    @mcp.tool()
    async def drive_upload_file(folder_path: str, filename: str, content: str) -> dict:
        """Cria ou substitui um arquivo de texto no OneDrive (até 4 MB).

        Args:
            folder_path: Caminho da pasta de destino, ex: '/Documentos'. Use '/' para a raiz.
            filename: Nome do arquivo, ex: 'relatorio.txt'.
            content: Conteúdo em texto (UTF-8).
        """
        path = drive_path(f"{folder_path.strip('/')}/{filename}")
        result = await graph.put_bytes(
            f"/me/drive/root:/{path}:/content", content.encode("utf-8"), "text/plain"
        )
        return {"id": result["id"], "name": result["name"], "webUrl": result.get("webUrl")}

    @mcp.tool()
    async def drive_delete_file(item_id: str) -> str:
        """Move um arquivo ou pasta do OneDrive para a lixeira.

        Args:
            item_id: ID do item a deletar.
        """
        await graph.delete(f"/me/drive/items/{item_id}")
        return f"Item '{item_id}' enviado para a lixeira."

    @mcp.tool()
    async def drive_move_file(
        item_id: str, destination_path: str, new_name: str | None = None
    ) -> dict:
        """Move e/ou renomeia um arquivo ou pasta no OneDrive.

        Args:
            item_id: ID do item a mover.
            destination_path: Pasta de destino, ex: '/Arquivos/2025'. Use '/' para a raiz.
            new_name: Novo nome (opcional; mantém o atual se omitido).
        """
        path = drive_path(destination_path)
        dest = await graph.get("/me/drive/root" if not path else f"/me/drive/root:/{path}")
        body: dict = {"parentReference": {"id": dest["id"]}}
        if new_name:
            body["name"] = new_name
        result = await graph.patch(f"/me/drive/items/{item_id}", body)
        return {"id": result["id"], "name": result["name"], "webUrl": result.get("webUrl")}

    @mcp.tool()
    async def drive_search(query: str) -> list[dict]:
        """Busca arquivos no OneDrive por nome ou conteúdo.

        Args:
            query: Texto a buscar.
        """
        q = quote(odata_str(query), safe="")
        data = await graph.get(f"/me/drive/root/search(q='{q}')")
        return [
            {
                "id": i["id"],
                "name": i["name"],
                "parentPath": i.get("parentReference", {}).get("path", ""),
                "modified": i.get("lastModifiedDateTime"),
                "webUrl": i.get("webUrl"),
            }
            for i in data.get("value", [])
        ]
