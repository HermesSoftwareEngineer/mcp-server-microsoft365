"""MCP Server Microsoft 365 — ponto de entrada.

Inicia o servidor FastMCP e registra as ferramentas de todos os módulos.

Uso:
    python main.py               # inicia o servidor MCP (stdio)
    python auth.py --setup       # primeiro login OAuth (executar uma vez)
"""

from dotenv import load_dotenv

load_dotenv()

from mcp.server.fastmcp import FastMCP

from modules import calendar, contacts, drive, excel, mail, tasks, teams

mcp = FastMCP(
    "microsoft365",
    instructions=(
        "Servidor MCP para Microsoft 365. Gerencie e-mails (Outlook), agenda, "
        "OneDrive, Excel, contatos, tarefas (To Do) e Teams via Microsoft Graph API. "
        "Todas as operações agem em nome do usuário autenticado."
    ),
)

drive.register(mcp)
excel.register(mcp)
mail.register(mcp)
calendar.register(mcp)
contacts.register(mcp)
tasks.register(mcp)
teams.register(mcp)

if __name__ == "__main__":
    mcp.run()
