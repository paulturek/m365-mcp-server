# M365 MCP Server v2.0.0

A Model Context Protocol (MCP) server providing comprehensive access to Microsoft 365 services with multi-user OAuth 2.0 authentication.

## Features

- **🔐 Multi-User OAuth 2.0**: Web-based login flow — each user authenticates independently
- **🔄 Automatic Token Refresh**: Encrypted token persistence with silent refresh via MSAL
- **📧 Outlook**: Read, send, list emails and calendar events
- **📁 OneDrive**: Browse, upload, download, delete, share files and folders
- **🌐 SharePoint**: Access sites, document libraries, search content
- **📊 Excel**: Read/write ranges, get workbook info, create charts
- **📄 Office Docs**: Get document content, convert to PDF
- **👥 Teams**: List teams/channels, send channel messages
- **✅ To Do**: List task lists, create/update tasks
- **👤 Users**: Get profiles, list/search directory users

## Architecture

```
┌──────────────────────────────────────────────────────────────────┐
│                    M365 MCP Server v2.0.0                        │
├──────────────────────────────────────────────────────────────────┤
│  ┌────────────────────────────────────────────────────────────┐  │
│  │             OAuth 2.0 Web Flow (MSAL)                      │  │
│  │  • /auth/login?user_id=<email>  → Microsoft consent        │  │
│  │  • /auth/callback               → token capture            │  │
│  │  • /auth/status?user_id=<email> → check auth state         │  │
│  │  • /auth/revoke?user_id=<email> → sign out                 │  │
│  │  • Pluggable token store (file / PostgreSQL)               │  │
│  │  • Encrypted at rest (Fernet / AES-128)                    │  │
│  └────────────────────────────────────────────────────────────┘  │
│                              │                                   │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │             MCP JSON-RPC 2.0 Dispatcher                    │  │
│  │  • POST /mcp  → tools/call, tools/list, initialize        │  │
│  │  • GET  /mcp  → tool manifest                              │  │
│  │  • Bearer token guard (MCP_BEARER_TOKEN)                   │  │
│  └────────────────────────────────────────────────────────────┘  │
│                              │                                   │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │              HTTP Client (httpx)                            │  │
│  │  • Microsoft Graph API (graph.microsoft.com/v1.0)          │  │
│  │  • Per-request access token injection                      │  │
│  │  • Pagination, error handling, file download/upload        │  │
│  └────────────────────────────────────────────────────────────┘  │
│                              │                                   │
│  ┌────────┬─────────┬────────┼────────┬────────┬─────────────┐  │
│  │Outlook │OneDrive │SharePt │ Excel  │ Teams  │  To Do      │  │
│  │  Mail  │ Files   │ Sites  │Workbook│Channel │  Tasks      │  │
│  │Calendar│ Folders │ Search │ Ranges │Messages│  Lists      │  │
│  ├────────┴─────────┴────────┴────────┴────────┴─────────────┤  │
│  │  Users (profiles, directory)  │  Office Docs (convert)    │  │
│  └───────────────────────────────┴───────────────────────────┘  │
└──────────────────────────────────────────────────────────────────┘
```

## Quick Start

### 1. Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) → Microsoft Entra ID → App registrations
2. Click **New registration**
3. Configure:
   - Name: `M365 MCP Server`
   - Supported account types: **Single tenant** (or multi-tenant based on your needs)
   - Redirect URI: **Web** → `https://<your-domain>/auth/callback`
4. Copy the **Application (client) ID** and **Directory (tenant) ID**
5. Under **Certificates & secrets**, create a new **Client secret** and copy the value

### 2. Configure API Permissions

Add these **delegated** permissions under API Permissions:

| API | Permission | Purpose |
|-----|------------|---------|
| Microsoft Graph | `User.Read` | User profile |
| Microsoft Graph | `Mail.ReadWrite` | Read/write emails |
| Microsoft Graph | `Calendars.ReadWrite` | Calendar access |
| Microsoft Graph | `Files.ReadWrite.All` | OneDrive / SharePoint files |
| Microsoft Graph | `Sites.ReadWrite.All` | SharePoint sites |
| Microsoft Graph | `ChannelMessage.Send` | Send Teams channel messages |
| Microsoft Graph | `Tasks.ReadWrite` | To Do tasks |
| Microsoft Graph | `offline_access` | Refresh tokens |

Click **Grant admin consent** if required by your organization.

### 3. Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_CLIENT_ID` | ✅ | Application (client) ID |
| `AZURE_CLIENT_SECRET` | ✅ | Client secret value |
| `AZURE_TENANT_ID` | ✅ | Directory (tenant) ID |
| `OAUTH_REDIRECT_URI` | ✅ | `https://<your-domain>/auth/callback` |
| `TOKEN_ENCRYPTION_KEY` | ✅ | Fernet key (see below) |
| `MCP_BEARER_TOKEN` | Recommended | Protects `POST /mcp` endpoint |
| `TOKEN_STORE_BACKEND` | ❌ | `file` (default) or `postgresql` |
| `DATABASE_URL` | ❌ | PostgreSQL connection string (if using pg token store) |

Generate a `TOKEN_ENCRYPTION_KEY`:
```bash
python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
```

### 4. Installation

```bash
git clone https://github.com/paulturek/m365-mcp-server.git
cd m365-mcp-server
pip install -e .
cp .env.example .env
# Edit .env with your Azure App credentials
```

### 5. Run the Server

```bash
python -m m365_mcp
```

The server starts on `http://0.0.0.0:8080` with endpoints:
- `POST /mcp` — MCP JSON-RPC 2.0 handler
- `GET /mcp` — Tool manifest
- `GET /health` — Health check
- `/auth/*` — OAuth flow

## Authentication Flow

```
┌──────────┐     ┌──────────────┐     ┌─────────────┐
│  User /  │────►│  /auth/login │────►│  Microsoft  │
│  Agent   │     │  ?user_id=   │     │  Consent    │
└──────────┘     └──────────────┘     └──────┬──────┘
                                             │
                 ┌──────────────┐            │
                 │ /auth/       │◄───────────┘
                 │  callback    │
                 └──────┬───────┘
                        │ Token stored (encrypted)
                        ▼
                 ┌──────────────┐
                 │ Tools ready  │
                 │ for user_id  │
                 └──────────────┘
```

1. Direct user to `/auth/login?user_id=user@company.com`
2. User completes Microsoft consent
3. Callback captures tokens, encrypts and stores them
4. All tool calls include `user_id` for per-user token resolution
5. Tokens refresh silently — re-auth only if idle for ~90 days

## Available Tools (32)

### OneDrive (6)
| Tool | Description |
|------|-------------|
| `onedrive_list_files` | List files and folders |
| `onedrive_download_file` | Download file content |
| `onedrive_upload_file` | Upload a file |
| `onedrive_delete_item` | Delete file or folder |
| `onedrive_create_folder` | Create a new folder |
| `onedrive_share_item` | Create sharing link |

### Excel (4)
| Tool | Description |
|------|-------------|
| `excel_get_workbook_info` | Get workbook metadata |
| `excel_read_range` | Read data from range |
| `excel_write_range` | Write data to range |
| `excel_create_chart` | Create a chart |

### Outlook (4)
| Tool | Description |
|------|-------------|
| `outlook_list_mail` | List emails from any folder |
| `outlook_send_mail` | Send a new email |
| `outlook_list_calendar_events` | List upcoming events |
| `outlook_create_event` | Create new event |

### SharePoint (5)
| Tool | Description |
|------|-------------|
| `sharepoint_list_sites` | List accessible sites |
| `sharepoint_list_items` | List items in a library |
| `sharepoint_download_file` | Download from SharePoint |
| `sharepoint_upload_file` | Upload to SharePoint |
| `sharepoint_search` | Search SharePoint content |

### Teams (3)
| Tool | Description |
|------|-------------|
| `teams_list_teams` | List joined teams |
| `teams_list_channels` | List channels in team |
| `teams_send_message` | Send channel message |

### To Do (4)
| Tool | Description |
|------|-------------|
| `todo_list_task_lists` | List task lists |
| `todo_list_tasks` | List tasks in a list |
| `todo_create_task` | Create a new task |
| `todo_update_task` | Update existing task |

### Users (4)
| Tool | Description |
|------|-------------|
| `users_get_me` | Get current user profile |
| `users_get_user` | Get specific user profile |
| `users_list_users` | List directory users |
| `users_search` | Search users by name/email |

### Office Documents (2)
| Tool | Description |
|------|-------------|
| `docs_get_content` | Get document content |
| `docs_convert` | Convert document to PDF |

## Railway Deployment

### Manual Deployment

1. Create a new Railway project
2. Connect your GitHub repository (`paulturek/m365-mcp-server`, branch: `main`)
3. Railway auto-detects the Dockerfile
4. Add environment variables:
   - `AZURE_CLIENT_ID`
   - `AZURE_CLIENT_SECRET`
   - `AZURE_TENANT_ID`
   - `OAUTH_REDIRECT_URI` = `https://<railway-domain>/auth/callback`
   - `TOKEN_ENCRYPTION_KEY`
   - `MCP_BEARER_TOKEN`
5. Deploy — health check hits `GET /health`

### Token Storage on Railway

If your Railway project includes a PostgreSQL service, set:
```
TOKEN_STORE_BACKEND=postgresql
DATABASE_URL=postgresql://...  (auto-provided by Railway)
```

Tokens will be stored in an encrypted `oauth_tokens` table, surviving container restarts.

Without PostgreSQL, file-based storage works but tokens are lost on redeploy.

## Security

- **MCP Bearer Guard**: `POST /mcp` requires `Authorization: Bearer <MCP_BEARER_TOKEN>` when configured
- **Token Encryption**: All OAuth tokens encrypted at rest with Fernet (AES-128-CBC)
- **Per-User Isolation**: Each `user_id` has independent token storage and Graph API access
- **Delegated Permissions**: All API calls run in user context — no application-level over-privilege
- **No Secrets in Code**: All credentials via environment variables

## Development

```bash
pip install -e ".[dev]"
pytest
mypy src/
ruff check src/
```

## License

MIT License — See [LICENSE](LICENSE) for details.
