"""Excel (.xlsx) via Workbook API da Graph."""

import re
from urllib.parse import quote

from mcp.server.mcpserver import MCPServer

import graph
from graph import odata_str


def _sheet(item_id: str, sheet_name: str) -> str:
    return f"/me/drive/items/{item_id}/workbook/worksheets/{quote(sheet_name, safe='')}"


def _range(item_id: str, sheet_name: str, address: str) -> str:
    return f"{_sheet(item_id, sheet_name)}/range(address='{quote(odata_str(address), safe=':')}')"


def register(mcp: MCPServer) -> None:

    @mcp.tool()
    async def excel_read_sheet(item_id: str, sheet_name: str) -> dict:
        """Lê todas as células usadas de uma aba do Excel.

        Args:
            item_id: ID do arquivo .xlsx no OneDrive.
            sheet_name: Nome da aba (worksheet).
        """
        data = await graph.get(f"{_sheet(item_id, sheet_name)}/usedRange(valuesOnly=true)")
        return {"address": data.get("address"), "values": data.get("values", [])}

    @mcp.tool()
    async def excel_append_row(item_id: str, sheet_name: str, values: list) -> dict:
        """Adiciona uma linha logo abaixo da última linha usada da aba, a partir da coluna A.

        Args:
            item_id: ID do arquivo .xlsx.
            sheet_name: Nome da aba.
            values: Valores da nova linha, ex: ['João', 30, 'SP'].
        """
        used = await graph.get(
            f"{_sheet(item_id, sheet_name)}/usedRange(valuesOnly=true)",
            params={"$select": "address,values"},
        )
        is_empty = used.get("values") in (None, [[""]], [[None]])
        last_row = 0 if is_empty else int(re.search(r"(\d+)$", used["address"]).group(1))
        target = last_row + 1
        address = f"A{target}:{_col_letter(len(values))}{target}"
        await graph.patch(_range(item_id, sheet_name, address), {"values": [values]})
        return {"address": address}

    @mcp.tool()
    async def excel_delete_row(item_id: str, sheet_name: str, row_index: int) -> str:
        """Remove uma linha de dados e desloca as de baixo para cima.

        Args:
            item_id: ID do arquivo .xlsx.
            sheet_name: Nome da aba.
            row_index: Índice 0-based da linha de dados, sem contar o cabeçalho
                (0 = primeira linha após o cabeçalho, ou seja, a linha 2 do Excel).
        """
        excel_row = row_index + 2
        await graph.post(
            f"{_range(item_id, sheet_name, f'{excel_row}:{excel_row}')}/delete", {"shift": "Up"}
        )
        return f"Linha {excel_row} do Excel removida."

    @mcp.tool()
    async def excel_update_cell(
        item_id: str, sheet_name: str, cell: str, value: str | int | float | bool
    ) -> dict:
        """Atualiza o valor de uma célula. Textos iniciados com '=' são gravados como fórmula.

        Args:
            item_id: ID do arquivo .xlsx.
            sheet_name: Nome da aba.
            cell: Endereço da célula, ex: 'B5'.
            value: Novo valor.
        """
        await graph.patch(_range(item_id, sheet_name, cell), {"values": [[value]]})
        return {"cell": cell, "value": value}


def _col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s
