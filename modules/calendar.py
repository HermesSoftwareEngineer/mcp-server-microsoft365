"""Agenda (Calendar) via Microsoft Graph API."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

import graph


def register(mcp: FastMCP) -> None:

    @mcp.tool()
    async def calendar_list_events(start: str, end: str) -> list[dict]:
        """Lista eventos da agenda em um intervalo de datas.

        Args:
            start: Data/hora de início, ex: '2025-10-01' ou '2025-10-01T08:00:00'.
            end: Data/hora de fim, ex: '2025-10-31' ou '2025-10-31T23:59:59'.
        """
        if "T" not in start:
            start += "T00:00:00"
        if "T" not in end:
            end += "T23:59:59"

        data = await graph.get(
            "/me/calendarView",
            {
                "startDateTime": start,
                "endDateTime": end,
                "$select": "id,subject,start,end,location,organizer,isAllDay",
                "$orderby": "start/dateTime",
                "$top": 100,
            },
        )
        return [
            {
                "id": e["id"],
                "subject": e.get("subject"),
                "start": e.get("start", {}).get("dateTime"),
                "end": e.get("end", {}).get("dateTime"),
                "isAllDay": e.get("isAllDay"),
                "location": e.get("location", {}).get("displayName"),
                "organizer": e.get("organizer", {}).get("emailAddress", {}).get("address"),
            }
            for e in data.get("value", [])
        ]

    @mcp.tool()
    async def calendar_get_event(event_id: str) -> dict:
        """Retorna os detalhes completos de um evento.

        Args:
            event_id: ID do evento retornado por calendar_list_events.
        """
        e = await graph.get(f"/me/events/{event_id}")
        return {
            "id": e["id"],
            "subject": e.get("subject"),
            "start": e.get("start"),
            "end": e.get("end"),
            "isAllDay": e.get("isAllDay"),
            "location": e.get("location", {}).get("displayName"),
            "body": e.get("body", {}).get("content"),
            "organizer": e.get("organizer", {}).get("emailAddress", {}).get("address"),
            "attendees": [
                {
                    "email": a["emailAddress"]["address"],
                    "name": a["emailAddress"].get("name"),
                    "status": a.get("status", {}).get("response"),
                }
                for a in e.get("attendees", [])
            ],
        }

    @mcp.tool()
    async def calendar_create_event(
        title: str,
        start: str,
        end: str,
        attendees: list[str] | None = None,
        location: str | None = None,
        body: str | None = None,
        is_all_day: bool = False,
    ) -> dict:
        """Cria um novo evento na agenda.

        Args:
            title: Título do evento.
            start: Data/hora de início, ex: '2025-10-15T14:00:00'.
            end: Data/hora de fim, ex: '2025-10-15T15:00:00'.
            attendees: Lista de e-mails dos participantes (opcional).
            location: Local do evento (opcional).
            body: Descrição do evento (opcional).
            is_all_day: Se True, cria como evento de dia inteiro. Padrão: False.
        """
        event: dict = {
            "subject": title,
            "start": {"dateTime": start, "timeZone": "America/Sao_Paulo"},
            "end": {"dateTime": end, "timeZone": "America/Sao_Paulo"},
            "isAllDay": is_all_day,
        }
        if attendees:
            event["attendees"] = [
                {"emailAddress": {"address": a}, "type": "required"} for a in attendees
            ]
        if location:
            event["location"] = {"displayName": location}
        if body:
            event["body"] = {"contentType": "Text", "content": body}

        result = await graph.post("/me/events", event)
        return {"id": result["id"], "subject": result["subject"], "webLink": result.get("webLink")}

    @mcp.tool()
    async def calendar_update_event(
        event_id: str,
        title: str | None = None,
        start: str | None = None,
        end: str | None = None,
        location: str | None = None,
        body: str | None = None,
    ) -> dict:
        """Atualiza campos de um evento existente.

        Args:
            event_id: ID do evento.
            title: Novo título (opcional).
            start: Nova data/hora de início, ex: '2025-10-15T14:00:00' (opcional).
            end: Nova data/hora de fim (opcional).
            location: Novo local (opcional).
            body: Nova descrição (opcional).
        """
        updates: dict = {}
        if title is not None:
            updates["subject"] = title
        if start is not None:
            updates["start"] = {"dateTime": start, "timeZone": "America/Sao_Paulo"}
        if end is not None:
            updates["end"] = {"dateTime": end, "timeZone": "America/Sao_Paulo"}
        if location is not None:
            updates["location"] = {"displayName": location}
        if body is not None:
            updates["body"] = {"contentType": "Text", "content": body}

        await graph.patch(f"/me/events/{event_id}", updates)
        return {"id": event_id, "updated": list(updates.keys())}

    @mcp.tool()
    async def calendar_delete_event(event_id: str) -> str:
        """Deleta um evento da agenda.

        Args:
            event_id: ID do evento.
        """
        await graph.delete(f"/me/events/{event_id}")
        return f"Evento '{event_id}' deletado."

    @mcp.tool()
    async def calendar_find_free_slots(
        attendees: list[str],
        duration_minutes: int,
        start: str,
        end: str,
    ) -> list[dict]:
        """Encontra horários livres entre os participantes.

        Args:
            attendees: Lista de e-mails dos participantes.
            duration_minutes: Duração desejada da reunião em minutos.
            start: Data/hora de início da janela de busca, ex: '2025-10-15T08:00:00'.
            end: Data/hora de fim da janela de busca, ex: '2025-10-15T18:00:00'.
        """
        body = {
            "attendees": [
                {"emailAddress": {"address": a}, "type": "required"} for a in attendees
            ],
            "timeConstraint": {
                "timeslots": [
                    {
                        "start": {"dateTime": start, "timeZone": "America/Sao_Paulo"},
                        "end": {"dateTime": end, "timeZone": "America/Sao_Paulo"},
                    }
                ]
            },
            "meetingDuration": f"PT{duration_minutes}M",
            "returnSuggestionReasons": True,
            "minimumAttendeePercentage": 100,
        }
        data = await graph.post("/me/findMeetingTimes", body)
        return [
            {
                "start": s["meetingTimeSlot"]["start"]["dateTime"],
                "end": s["meetingTimeSlot"]["end"]["dateTime"],
                "confidence": s.get("confidence"),
            }
            for s in data.get("meetingTimeSuggestions", [])
        ]
