"""Microsoft To Do via Microsoft Graph API."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

import graph


def register(mcp: FastMCP) -> None:

    @mcp.tool()
    async def tasks_list_lists() -> list[dict]:
        """Lista todas as listas de tarefas do Microsoft To Do."""
        data = await graph.get_all(
            "/me/todo/lists",
            {"$select": "id,displayName,isOwner,isShared"},
        )
        return [
            {
                "id": lst["id"],
                "name": lst.get("displayName"),
                "isOwner": lst.get("isOwner"),
                "isShared": lst.get("isShared"),
            }
            for lst in data
        ]

    @mcp.tool()
    async def tasks_list(list_id: str, include_completed: bool = False) -> list[dict]:
        """Lista as tarefas de uma lista específica.

        Args:
            list_id: ID da lista retornado por tasks_list_lists.
            include_completed: Se True, inclui tarefas concluídas. Padrão: False.
        """
        params: dict = {
            "$select": "id,title,status,dueDateTime,importance,body,createdDateTime"
        }
        if not include_completed:
            params["$filter"] = "status ne 'completed'"

        data = await graph.get_all(f"/me/todo/lists/{list_id}/tasks", params)
        return [
            {
                "id": t["id"],
                "title": t.get("title"),
                "status": t.get("status"),
                "due": t.get("dueDateTime", {}).get("dateTime") if t.get("dueDateTime") else None,
                "importance": t.get("importance"),
                "notes": t.get("body", {}).get("content") if t.get("body") else None,
                "created": t.get("createdDateTime"),
            }
            for t in data
        ]

    @mcp.tool()
    async def tasks_create(
        list_id: str,
        title: str,
        due_date: str | None = None,
        notes: str | None = None,
        importance: str = "normal",
    ) -> dict:
        """Cria uma nova tarefa em uma lista.

        Args:
            list_id: ID da lista onde criar a tarefa.
            title: Título da tarefa.
            due_date: Data de vencimento, ex: '2025-10-31' (opcional).
            notes: Notas/descrição (opcional).
            importance: Prioridade: 'low', 'normal' ou 'high'. Padrão: 'normal'.
        """
        task: dict = {"title": title, "importance": importance}
        if due_date:
            task["dueDateTime"] = {
                "dateTime": f"{due_date}T00:00:00",
                "timeZone": "America/Sao_Paulo",
            }
        if notes:
            task["body"] = {"contentType": "text", "content": notes}

        result = await graph.post(f"/me/todo/lists/{list_id}/tasks", task)
        return {"id": result["id"], "title": result.get("title"), "status": result.get("status")}

    @mcp.tool()
    async def tasks_update(
        list_id: str,
        task_id: str,
        title: str | None = None,
        status: str | None = None,
        due_date: str | None = None,
        notes: str | None = None,
        importance: str | None = None,
    ) -> dict:
        """Atualiza uma tarefa existente.

        Args:
            list_id: ID da lista que contém a tarefa.
            task_id: ID da tarefa.
            title: Novo título (opcional).
            status: Novo status: 'notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred' (opcional).
            due_date: Nova data de vencimento, ex: '2025-10-31' (opcional).
            notes: Novas notas (opcional).
            importance: Nova prioridade: 'low', 'normal', 'high' (opcional).
        """
        updates: dict = {}
        if title is not None:
            updates["title"] = title
        if status is not None:
            updates["status"] = status
        if importance is not None:
            updates["importance"] = importance
        if due_date is not None:
            updates["dueDateTime"] = {
                "dateTime": f"{due_date}T00:00:00",
                "timeZone": "America/Sao_Paulo",
            }
        if notes is not None:
            updates["body"] = {"contentType": "text", "content": notes}

        await graph.patch(f"/me/todo/lists/{list_id}/tasks/{task_id}", updates)
        return {"id": task_id, "updated": list(updates.keys())}

    @mcp.tool()
    async def tasks_delete(list_id: str, task_id: str) -> str:
        """Deleta uma tarefa.

        Args:
            list_id: ID da lista.
            task_id: ID da tarefa.
        """
        await graph.delete(f"/me/todo/lists/{list_id}/tasks/{task_id}")
        return f"Tarefa '{task_id}' deletada."
