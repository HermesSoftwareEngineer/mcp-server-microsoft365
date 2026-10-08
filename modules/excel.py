"""Excel (.xlsx) via Microsoft Graph Workbook API."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

import graph


def register(mcp: FastMCP) -> None:

    @mcp.tool()
    async def excel_read_sheet(item_id: str, sheet_name: str) -> list[list]:
        """Lê todas as linhas usadas de uma aba do Excel.

        Args:
            item_id: ID do arquivo .xlsx no OneDrive.
            sheet_name: Nome da aba (worksheet).
        """
        data = await graph.get(
            f"/me/drive/items/{item_id}/workbook/worksheets/{sheet_name}/usedRange"
        )
        return data.get("values", [])

    @mcp.tool()
    async def excel_append_row(item_id: str, sheet_name: str, values: list) -> dict:
        """Adiciona uma linha ao final dos dados de uma aba.

        Args:
            item_id: ID do arquivo .xlsx.
            sheet_name: Nome da aba.
            values: Lista de valores a inserir, ex: ['João', 30, 'SP'].
        """
        used = await graph.get(
            f"/me/drive/items/{item_id}/workbook/worksheets/{sheet_name}/usedRange"
        )
        row_count = used.get("rowCount", 0)
        col_count = max(len(values), used.get("columnCount", 1))

        def col_letter(n: int) -> str:
            s = ""
            while n > 0:
                n, r = divmod(n - 1, 26)
                s = chr(65 + r) + s
            return s

        target_row = row_count + 1
        end_col = col_letter(col_count)
        address = f"A{target_row}:{end_col}{target_row}"

        await graph.patch(
            f"/me/drive/items/{item_id}/workbook/worksheets/{sheet_name}/range(address='{address}')",
            {"values": [values]},
        )
        return {"address": address, "rowsAdded": 1}

    @mcp.tool()
    async def excel_delete_row(item_id: str, sheet_name: str, row_index: int) -> str:
        """Deleta uma linha pelo índice (0-based, sem contar o cabeçalho).

        Args:
            item_id: ID do arquivo .xlsx.
            sheet_name: Nome da aba.
            row_index: Índice da linha de dados (0 = primeira linha após o cabeçalho).
        """
        # Linha 1 é o cabeçalho; dados começam na linha 2
        excel_row = row_index + 2
        await graph.post(
            f"/me/drive/items/{item_id}/workbook/worksheets/{sheet_name}/rows/{excel_row}/delete",
            {},
        )
        return f"Linha {row_index} (Excel linha {excel_row}) deletada."

    @mcp.tool()
    async def excel_update_cell(
        item_id: str, sheet_name: str, cell: str, value: str | int | float
    ) -> dict:
        """Atualiza o valor de uma célula específica.

        Args:
            item_id: ID do arquivo .xlsx.
            sheet_name: Nome da aba.
            cell: Endereço da célula, ex: 'B5'.
            value: Novo valor da célula.
        """
        await graph.patch(
            f"/me/drive/items/{item_id}/workbook/worksheets/{sheet_name}/range(address='{cell}')",
            {"values": [[value]]},
        )
        return {"cell": cell, "value": value}
