# MCP Server Microsoft 365

Servidor MCP remoto (Streamable HTTP) que dá ao Claude acesso ao Microsoft 365 do usuário:
Outlook (e-mail e agenda), OneDrive/SharePoint, Excel, contatos, Microsoft To Do e Teams.

Ele é conectado no Claude como **conector personalizado** e se autentica com **OAuth 2.1**:

```
Claude ──/authorize──▶ este servidor ──▶ login Microsoft (Entra ID)
                                               │
Claude ◀──código──── /auth/microsoft/callback ◀┘
Claude ──/token (PKCE)──▶ token deste servidor ──▶ chamadas em /mcp
```

- O servidor é o servidor de autorização do conector: registro dinâmico de cliente, PKCE, refresh token com rotação e revogação.
- Os tokens da Microsoft ficam no servidor (SQLite), associados a cada usuário. O Claude nunca os vê.
- Só os e-mails listados em `ALLOWED_USERS` conseguem concluir o login.
- Se a sessão Microsoft expirar ou for revogada, o servidor invalida os tokens do conector e o Claude pede para reconectar.

## Ferramentas (35)

| Módulo | Ferramentas |
|---|---|
| OneDrive / SharePoint | `drive_list_files`, `drive_get_file`, `drive_upload_file`, `drive_delete_file`, `drive_move_file`, `drive_search` |
| Excel | `excel_read_sheet`, `excel_append_row`, `excel_delete_row`, `excel_update_cell` |
| E-mail | `mail_list`, `mail_get`, `mail_send`, `mail_reply`, `mail_move`, `mail_delete`, `mail_search` |
| Agenda | `calendar_list_events`, `calendar_get_event`, `calendar_create_event`, `calendar_update_event`, `calendar_delete_event`, `calendar_find_free_slots` |
| Contatos | `contacts_list`, `contacts_get`, `contacts_create`, `contacts_update`, `contacts_delete` |
| To Do | `tasks_list_lists`, `tasks_list`, `tasks_create`, `tasks_update`, `tasks_delete` |
| Teams | `teams_list_chats`, `teams_send_message` |

## 1. Registrar o app no Microsoft Entra ID

1. Em [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID → Registros de aplicativo → Novo registro**.
   - Tipos de conta: **Somente contas deste diretório organizacional**.
   - URI de redirecionamento: plataforma **Web**, `https://SEU_DOMINIO/auth/microsoft/callback`.
2. Anote o **ID do aplicativo (cliente)** e o **ID do diretório (locatário)**.
3. **Certificados e segredos → Novo segredo do cliente**. Copie o **valor** (ele só aparece uma vez).
4. **Permissões de API → Adicionar permissão → Microsoft Graph → Permissões delegadas**:

   `offline_access`, `User.Read`, `Files.ReadWrite`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`,
   `Contacts.ReadWrite`, `Tasks.ReadWrite`, `Chat.ReadWrite`, `ChannelMessage.Send`, `Sites.ReadWrite.All`

5. Se o tenant não permite que usuários consintam sozinhos, um administrador precisa clicar em
   **Conceder consentimento do administrador**.

## 2. Instalar na VPS

Pré-requisitos: Python 3.11+, um subdomínio apontando para a VPS (registro DNS A/AAAA) e portas 80/443 abertas.

```bash
sudo useradd --system --home /opt/m365-mcp --shell /usr/sbin/nologin m365mcp
sudo git clone https://github.com/HermesSoftwareEngineer/mcp-server-microsoft365.git /opt/m365-mcp
cd /opt/m365-mcp
sudo python3 -m venv .venv
sudo .venv/bin/pip install -r requirements.txt
sudo cp .env.example .env
sudo mkdir -p data
sudo chown -R m365mcp:m365mcp /opt/m365-mcp
sudo chmod 600 .env
```

Edite o `.env`:

| Variável | Valor |
|---|---|
| `PUBLIC_URL` | URL pública HTTPS, sem barra no final. Ex.: `https://m365.seudominio.com.br` |
| `AZURE_CLIENT_ID` | ID do aplicativo (cliente) |
| `AZURE_CLIENT_SECRET` | Valor do segredo do cliente |
| `AZURE_TENANT_ID` | ID do locatário ou `officetorc.onmicrosoft.com` |
| `ALLOWED_USERS` | E-mails autorizados, separados por vírgula |
| `HOST` / `PORT` | Onde o servidor escuta localmente (padrão `127.0.0.1:8000`) |
| `DB_PATH` | Banco SQLite com clientes e tokens (padrão `data/m365.db`) |
| `TIMEZONE` | Fuso usado para datas sem fuso (padrão `America/Sao_Paulo`) |

### Serviço systemd

```bash
sudo cp deploy/m365-mcp.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now m365-mcp
sudo journalctl -u m365-mcp -f
```

### HTTPS com proxy reverso

O Claude acessa o servidor pela internet, então ele precisa de HTTPS com certificado válido.

**Caddy** (emite o certificado sozinho): troque o domínio em `deploy/Caddyfile`, copie para
`/etc/caddy/Caddyfile` e rode `sudo systemctl reload caddy`.

**Nginx** (com certificado do certbot):

```nginx
server {
    listen 443 ssl;
    server_name m365.seudominio.com.br;
    # ssl_certificate ... (certbot)

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Forwarded-Proto https;
        proxy_http_version 1.1;
        proxy_buffering off;
        proxy_read_timeout 300s;
    }
}
```

O proxy tem que repassar o cabeçalho `Host` original. O servidor recusa (HTTP 421) qualquer host que
não seja o de `PUBLIC_URL`.

Teste: `curl https://SEU_DOMINIO/health` deve responder `ok`, e
`curl https://SEU_DOMINIO/.well-known/oauth-authorization-server` deve devolver o JSON de metadados.

## 3. Conectar no Claude

1. No Claude: **Configurações → Conectores → Adicionar conector personalizado**.
2. Nome: `Microsoft 365`. URL: `https://SEU_DOMINIO/mcp`.
3. Deixe Client ID e Client Secret em branco (o Claude se registra sozinho).
4. Clique em **Conectar**, faça login com a conta Microsoft autorizada e aceite as permissões.

## Desenvolvimento local

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt   # Windows: .venv\Scripts\pip
cp .env.example .env    # use PUBLIC_URL=http://localhost:8000
python main.py
```

Para o login Microsoft funcionar localmente, adicione também `http://localhost:8000/auth/microsoft/callback`
como URI de redirecionamento do app no Entra ID.

## Segurança

- `data/m365.db` guarda os refresh tokens da Microsoft em texto puro. O serviço roda com usuário próprio e `UMask=0077`; não exponha nem faça backup desse arquivo em local inseguro.
- Tokens e códigos emitidos para o Claude são guardados só como hash SHA-256.
- Para cortar o acesso de alguém: remova o e-mail de `ALLOWED_USERS` e reinicie o serviço. A lista é checada em toda requisição, então as sessões já abertas param na hora.
