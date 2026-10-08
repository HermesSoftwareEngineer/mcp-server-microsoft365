"""Outlook — agenda. Datas sem fuso são interpretadas no fuso de TIMEZONE."""

from datetime import datetime
from zoneinfo import ZoneInfo

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

import config
import graph

TZ_HEADER = {"Prefer": f'outlook.timezone="{config.TIMEZONE}"'}


def _with_offset(value: str, end_of_day: bool = False) -> str:
    if "T" not in value:
        value += "T23:59:59" if end_of_day else "T00:00:00"
    dt = datetime.fromisoformat(value)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=ZoneInfo(config.TIMEZONE))
    return dt.isoformat()


def _when(value: str) -> dict:
    return {"dateTime": value, "timeZone": config.TIMEZONE}


def _summary(e: dict) -> dict:
    return {
        "id": e["id"],
        "subject": e.get("subject"),
        "start": e.get("start", {}).get("dateTime"),
        "end": e.get("end", {}).get("dateTime"),
        "isAllDay": e.get("isAllDay"),
        "location": e.get("location", {}).get("displayName"),
        "organizer": e.get("organizer", {}).get("emailAddress", {}).get("address"),
        "onlineMeetingUrl": (e.get("onlineMeeting") or {}).get("joinUrl"),
    }


def register(mcp: MCPServer) -> None:

    @mcp.tool()
    async def calendar_list_events(start: str, end: str) -> list[dict]:
        """Lista eventos (incluindo ocorrências de recorrentes) entre duas datas.

        Args:
            start: Início, ex: '2025-10-01' ou '2025-10-01T08:00:00'.
            end: Fim, ex: '2025-10-31' ou '2025-10-31T18:00:00'.
        """
        events = await graph.get_all(
            "/me/calendarView",
            {
                "startDateTime": _with_offset(start),
                "endDateTime": _with_offset(end, end_of_day=True),
                "$select": "id,subject,start,end,location,organizer,isAllDay,onlineMeeting",
                "$orderby": "start/dateTime",
                "$top": 100,
            },
            headers=TZ_HEADER,
        )
        return [_summary(e) for e in events]

    @mcp.tool()
    async def calendar_get_event(event_id: str) -> dict:
        """Detalha um evento, com descrição e participantes.

        Args:
            event_id: ID do evento.
        """
        e = await graph.get(
            f"/me/events/{event_id}",
            headers={"Prefer": f'{TZ_HEADER["Prefer"]}, outlook.body-content-type="text"'},
        )
        return {
            **_summary(e),
            "body": e.get("body", {}).get("content"),
            "attendees": [
                {
                    "email": a["emailAddress"]["address"],
                    "name": a["emailAddress"].get("name"),
                    "response": a.get("status", {}).get("response"),
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
        teams_meeting: bool = False,
    ) -> dict:
        """Cria um evento. Se houver participantes, eles recebem o convite.

        Args:
            title: Título.
            start: Início, ex: '2025-10-15T14:00:00' (para dia inteiro: '2025-10-15T00:00:00').
            end: Fim, ex: '2025-10-15T15:00:00' (para dia inteiro: o dia seguinte às 00:00).
            attendees: E-mails dos participantes (opcional).
            location: Local (opcional).
            body: Descrição (opcional).
            is_all_day: Evento de dia inteiro.
            teams_meeting: Cria reunião online do Teams.
        """
        event: dict = {
            "subject": title,
            "start": _when(start),
            "end": _when(end),
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
        if teams_meeting:
            event["isOnlineMeeting"] = True
            event["onlineMeetingProvider"] = "teamsForBusiness"

        result = await graph.post("/me/events", event)
        return {**_summary(result), "webLink": result.get("webLink")}

    @mcp.tool()
    async def calendar_update_event(
        event_id: str,
        title: str | None = None,
        start: str | None = None,
        end: str | None = None,
        location: str | None = None,
        body: str | None = None,
    ) -> dict:
        """Atualiza campos de um evento. Só os campos informados são alterados.

        Args:
            event_id: ID do evento.
            title: Novo título.
            start: Novo início, ex: '2025-10-15T14:00:00'.
            end: Novo fim.
            location: Novo local.
            body: Nova descrição.
        """
        updates: dict = {}
        if title is not None:
            updates["subject"] = title
        if start is not None:
            updates["start"] = _when(start)
        if end is not None:
            updates["end"] = _when(end)
        if location is not None:
            updates["location"] = {"displayName": location}
        if body is not None:
            updates["body"] = {"contentType": "Text", "content": body}
        if not updates:
            raise ToolError("Informe ao menos um campo para atualizar.")

        await graph.patch(f"/me/events/{event_id}", updates)
        return {"id": event_id, "updated": list(updates)}

    @mcp.tool()
    async def calendar_delete_event(event_id: str) -> str:
        """Exclui um evento. Se você for o organizador, os participantes recebem o cancelamento.

        Args:
            event_id: ID do evento.
        """
        await graph.delete(f"/me/events/{event_id}")
        return f"Evento '{event_id}' excluído."

    @mcp.tool()
    async def calendar_find_free_slots(
        attendees: list[str],
        duration_minutes: int,
        start: str,
        end: str,
    ) -> list[dict]:
        """Sugere horários em que todos os participantes (e você) estão livres.

        Args:
            attendees: E-mails dos participantes.
            duration_minutes: Duração da reunião em minutos.
            start: Início da janela de busca, ex: '2025-10-15T08:00:00'.
            end: Fim da janela de busca, ex: '2025-10-17T18:00:00'.
        """
        data = await graph.post(
            "/me/findMeetingTimes",
            {
                "attendees": [
                    {"emailAddress": {"address": a}, "type": "required"} for a in attendees
                ],
                "timeConstraint": {
                    "activityDomain": "work",
                    "timeSlots": [{"start": _when(start), "end": _when(end)}],
                },
                "meetingDuration": f"PT{duration_minutes}M",
                "minimumAttendeePercentage": 100,
                "maxCandidates": 20,
            },
            headers=TZ_HEADER,
        )
        return [
            {
                "start": s["meetingTimeSlot"]["start"]["dateTime"],
                "end": s["meetingTimeSlot"]["end"]["dateTime"],
                "timeZone": s["meetingTimeSlot"]["start"].get("timeZone"),
                "confidence": s.get("confidence"),
            }
            for s in data.get("meetingTimeSuggestions", [])
        ] or [{"message": data.get("emptySuggestionsReason", "Nenhum horário livre encontrado.")}]
