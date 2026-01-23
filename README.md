# M365 MCP Server

[![Deploy on Railway](https://railway.app/button.svg)](https://railway.app/template/m365-mcp)

A Model Context Protocol (MCP) server providing comprehensive access to Microsoft 365 services with automatic OAuth token refresh.

## Features

- **🔐 Automatic Authentication**: MSAL-based OAuth 2.0 with encrypted token persistence
- **🔄 Silent Token Refresh**: Tokens refresh automatically, users stay authenticated for months
- **📧 Outlook Mail**: Read, send, reply, search emails
- **📅 Calendar**: List, create, update, delete events with Teams meeting support
- **📁 OneDrive**: Browse, search, upload, download files
- **🌐 SharePoint**: Access sites, document libraries, and lists
- **📊 Excel**: Read/write ranges, tables, worksheets
- **📄 Word & PowerPoint**: Download, upload, convert to PDF
- **👥 Teams**: List teams/channels, send messages, manage chats
- **📈 Power BI**: Access workspaces, reports, datasets, trigger refreshes

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                      M365 MCP Server                            │
├─────────────────────────────────────────────────────────────────┤
│  ┌───────────────────────────────────────────────────────────┐  │
│  │              Token Manager (MSAL)                         │  │
│  │  • Encrypted persistent cache (AES-256)                   │  │
│  │  • Auto-refresh ~5 min before expiry                      │  │
│  │  • Device Code + Auth Code PKCE flows                     │  │
│  └───────────────────────────────────────────────────────────┘  │
│                              │                                  │
│  ┌───────────────────────────▼───────────────────────────────┐  │
│  │              HTTP Clients (httpx)                         │  │
│  │  • Microsoft Graph API (graph.microsoft.com)              │  │
│  │  • Power BI API (api.powerbi.com)                         │  │
│  └───────────────────────────────────────────────────────────┘  │
│                              │                                  │
│  ┌────────┬────────┬────────┼────────┬────────┬────────────┐   │
│  │Outlook │OneDrive│SharePt │ Excel  │ Teams  │  PowerBI   │   │
│  │  Mail  │ Files  │ Sites  │Workbook│Channel │  Reports   │   │
│  │Calendar│        │ Lists  │ Ranges │ Chats  │  Datasets  │   │
│  └────────┴────────┴────────┴────────┴────────┴────────────┘   │
└─────────────────────────────────────────────────────────────────┘
```

## Quick Start

### 1. Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) → Microsoft Entra ID → App registrations
2. Click **New registration**
3. Configure:
   - Name: `M365 MCP Server`
   - Supported account types: Choose based on your needs
   - Redirect URI: Leave blank for device code flow
4. Copy the **Application (client) ID**

### 2. Configure API Permissions

Add these **delegated** permissions under API Permissions:

| API | Permission | Purpose |
|-----|------------|----------|
| Microsoft Graph | `User.Read` | User profile |
| Microsoft Graph | `Mail.ReadWrite` | Read/write emails |
| Microsoft Graph | `Mail.Send` | Send emails |
| Microsoft Graph | `Calendars.ReadWrite` | Calendar access |
| Microsoft Graph | `Files.ReadWrite.All` | OneDrive/SharePoint files |
| Microsoft Graph | `Sites.ReadWrite.All` | SharePoint sites |
| Microsoft Graph | `Team.ReadBasic.All` | List Teams |
| Microsoft Graph | `Channel.ReadBasic.All` | List channels |
| Microsoft Graph | `Chat.ReadWrite` | Read/write chats |
| Microsoft Graph | `ChannelMessage.Send` | Send channel messages |
| Microsoft Graph | `offline_access` | Refresh tokens |
| Power BI Service | `Report.Read.All` | Read reports |
| Power BI Service | `Dataset.ReadWrite.All` | Manage datasets |

### 3. Installation

```bash
# Clone the repository
git clone https://github.com/paulturek/m365-mcp-server.git
cd m365-mcp-server

# Install dependencies
pip install -e .

# Configure environment
cp .env.example .env
# Edit .env with your Azure App credentials
```

### 4. Run the Server

```bash
# Run directly
python -m m365_mcp.server

# Or use the entry point
m365-mcp
```

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `M365_CLIENT_ID` | ✅ | Azure AD Application (client) ID |
| `M365_TENANT_ID` | ❌ | Tenant ID (default: `common`) |
| `M365_CLIENT_SECRET` | ❌ | Client secret for confidential client flow |
| `M365_TOKEN_CACHE_PATH` | ❌ | Custom token cache location |
| `M365_CACHE_ENCRYPTION_KEY` | ❌ | Custom encryption key for token cache |

## Available Tools

### Authentication
| Tool | Description |
|------|-------------|
| `m365_authenticate` | Start device code authentication flow |
| `m365_auth_status` | Check current authentication status |
| `m365_logout` | Sign out and clear cached tokens |

### Outlook Mail
| Tool | Description |
|------|-------------|
| `outlook_list_messages` | List emails from any folder |
| `outlook_get_message` | Get full email content by ID |
| `outlook_send_email` | Send a new email |
| `outlook_reply` | Reply to an email |
| `outlook_list_folders` | List mail folders |

### Calendar
| Tool | Description |
|------|-------------|
| `calendar_list_events` | List upcoming events |
| `calendar_get_event` | Get event details |
| `calendar_create_event` | Create new event (with Teams option) |
| `calendar_update_event` | Update existing event |
| `calendar_delete_event` | Delete an event |
| `calendar_find_meeting_times` | Find available meeting times |

### OneDrive
| Tool | Description |
|------|-------------|
| `onedrive_list_files` | List files and folders |
| `onedrive_search` | Search for files |
| `onedrive_download` | Download file content |
| `onedrive_upload` | Upload a file |
| `onedrive_create_folder` | Create a new folder |
| `onedrive_delete` | Delete file or folder |
| `onedrive_share` | Create sharing link |

### SharePoint
| Tool | Description |
|------|-------------|
| `sharepoint_search_sites` | Search for sites |
| `sharepoint_list_drives` | List document libraries |
| `sharepoint_list_items` | List files in a library |
| `sharepoint_list_lists` | List SharePoint lists |
| `sharepoint_get_list_items` | Get list items |
| `sharepoint_create_list_item` | Create list item |

### Excel
| Tool | Description |
|------|-------------|
| `excel_list_worksheets` | List worksheets in workbook |
| `excel_read_range` | Read data from range |
| `excel_write_range` | Write data to range |
| `excel_get_tables` | List tables in workbook |
| `excel_add_table_rows` | Add rows to a table |

### Teams
| Tool | Description |
|------|-------------|
| `teams_list_teams` | List joined teams |
| `teams_list_channels` | List channels in team |
| `teams_send_channel_message` | Send message to channel |
| `teams_list_chats` | List user's chats |
| `teams_send_chat_message` | Send message to chat |

### Power BI
| Tool | Description |
|------|-------------|
| `powerbi_list_workspaces` | List workspaces |
| `powerbi_list_reports` | List reports |
| `powerbi_list_datasets` | List datasets |
| `powerbi_refresh_dataset` | Trigger dataset refresh |
| `powerbi_get_refresh_history` | Get refresh history |

## Railway Deployment

### One-Click Deploy

[![Deploy on Railway](https://railway.app/button.svg)](https://railway.app/new/template?template=https://github.com/paulturek/m365-mcp-server)

### Manual Deployment

1. Create a new Railway project
2. Connect your GitHub repository
3. Add environment variables:
   - `M365_CLIENT_ID`
   - `M365_TENANT_ID`
4. Deploy!

The server will be available at your Railway-provided URL.

## How Token Refresh Works

MSAL handles token lifecycle automatically:

1. **Initial Auth**: User completes device code flow, receives access + refresh tokens
2. **Token Cache**: Tokens encrypted with AES-256 and persisted to disk
3. **Silent Refresh**: When access token expires (~1 hour), MSAL uses refresh token
4. **Refresh Token Lifetime**: ~90 days, extended on each use
5. **Re-auth Required**: Only if refresh token expires (rare with regular use)

```
┌─────────────┐    ┌─────────────┐    ┌─────────────┐
│  Access     │    │  Refresh    │    │  Re-auth    │
│  Token      │───►│  Token      │───►│  Required   │
│  (~1 hour)  │    │  (~90 days) │    │  (if idle)  │
└─────────────┘    └─────────────┘    └─────────────┘
       │                  │
       │   Auto-refresh   │
       └──────────────────┘
```

## Security Best Practices

1. **Least Privilege**: Only request permissions your app actually needs
2. **Token Encryption**: Tokens encrypted at rest with AES-256
3. **No Secrets in Code**: All credentials via environment variables
4. **Delegated Permissions**: User context ensures proper access control

## Development

```bash
# Install dev dependencies
pip install -e ".[dev]"

# Run tests
pytest

# Type checking
mypy src/

# Linting
ruff check src/
```

## License

MIT License - See [LICENSE](LICENSE) for details.

## Contributing

Contributions welcome! Please read our contributing guidelines first.
