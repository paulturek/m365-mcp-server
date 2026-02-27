# m365-mcp-server

A production-grade **Model Context Protocol (MCP) server** for Microsoft 365, built with FastAPI. Exposes 68 tools covering Outlook, Calendar, OneDrive, SharePoint, Teams, Excel, To Do, Users/Directory, and Office Docs — all via Microsoft Graph API.

---

## Overview

| Property | Value |
|---|---|
| Version | 2.0.1 |
| MCP Protocol | `2024-11-05` |
| Transport | HTTP (JSON-RPC 2.0) at `/mcp` |
| Auth | MSAL Device Code Flow (per-user tokens, stored in PostgreSQL) |
| Tools | **68** |

---

## Quick Start

### 1. Prerequisites

- Python 3.11+
- PostgreSQL database (for token storage)
- Azure App Registration with the permissions listed below

### 2. Environment Variables

```env
# Azure AD
AZURE_CLIENT_ID=<your-app-client-id>
AZURE_TENANT_ID=<your-tenant-id>

# Token storage
DATABASE_URL=postgresql://user:pass@host:5432/dbname

# MCP security
MCP_BEARER_TOKEN=<your-secret-bearer-token>

# Domain normalization (appended to bare usernames)
USER_EMAIL_DOMAIN=bolthousefresh.com
```

### 3. Run

```bash
pip install -e .
python -m m365_mcp
```

Server starts on `http://0.0.0.0:8080`. MCP endpoint: `POST /mcp`.

### 4. Health Check

```
GET /health
```

---

## Authentication

The server uses **MSAL Device Code Flow**. On first use per user:

1. Call `auth_start_device_login` with `user_id` (email)
2. User visits the URL shown and enters the device code
3. Call `auth_check_device_login` to confirm completion
4. All subsequent tool calls use the cached token automatically (auto-refreshed)

**Azure App Registration requirements:**
- "Allow public client flows" must be set to **Yes**
- Admin consent must be granted for all required permissions (see below)

---

## Required Azure AD Permissions

All are **Delegated** permissions on Microsoft Graph:

| Permission | Scope |
|---|---|
| `User.Read` | Sign in and read user profile |
| `User.Read.All` | Read all users' profiles (directory lookup, manager, direct reports) |
| `Mail.ReadWrite` | Read, send, update, delete mail |
| `Mail.Send` | Send mail |
| `Calendars.ReadWrite` | Read and write calendar events |
| `Files.ReadWrite.All` | Full OneDrive access |
| `Sites.ReadWrite.All` | SharePoint sites, lists, and files |
| `Tasks.ReadWrite` | To Do lists and tasks |
| `Team.ReadBasic.All` | List joined teams |
| `Channel.ReadBasic.All` | List channels |
| `ChannelMessage.Read.All` | Read channel messages |
| `ChannelMessage.Send` | Send channel messages |
| `Chat.ReadWrite` | List and read chats |
| `ChatMessage.Send` | Send chat messages |

After adding permissions, click **"Grant admin consent"** in the Azure portal.

---

## Tool Reference

### Authentication (3 tools)

| Tool | Description |
|---|---|
| `auth_status` | Check authentication status for a user |
| `auth_start_device_login` | Initiate device code flow — returns URL and code for user to authenticate |
| `auth_check_device_login` | Poll for device code completion and store token |

---

### Outlook — Mail (8 tools)

| Tool | Key Parameters | Notes |
|---|---|---|
| `outlook_list_mail` | `user_id`, `folder`, `top`, `filter`, `search` | `search` and `filter` are mutually exclusive. `search` does keyword matching; `filter` supports `eq`, `ne`, `startsWith()`, `isRead eq true/false`. `contains()` is NOT supported. |
| `outlook_get_message` | `user_id`, `message_id` | Returns full body content |
| `outlook_send_mail` | `user_id`, `to[]`, `subject`, `body`, `cc[]`, `content_type` | HTML or plain text |
| `outlook_update_message` | `user_id`, `message_id`, `is_read`, `importance`, `categories[]`, `flag` | Mark read/unread, set importance, manage categories and follow-up flags |
| `outlook_delete_message` | `user_id`, `message_id` | Moves to Deleted Items |
| `outlook_move_message` | `user_id`, `message_id`, `destination_folder` | Use folder ID or well-known name (`inbox`, `archive`, `deleteditems`, `drafts`) |
| `outlook_reply_mail` | `user_id`, `message_id`, `comment`, `reply_all` | Reply or reply-all |
| `outlook_forward_mail` | `user_id`, `message_id`, `to[]`, `comment` | Forward to new recipients |
| `outlook_list_mail_folders` | `user_id`, `top` | Lists all mail folders with item counts |

---

### Outlook — Calendar (4 tools)

| Tool | Key Parameters | Notes |
|---|---|---|
| `outlook_list_calendar_events` | `user_id`, `top`, `start_datetime`, `end_datetime` | Use `start_datetime` + `end_datetime` for bounded range (calendarView); omit for next N events |
| `outlook_create_event` | `user_id`, `subject`, `start`, `end`, `timezone`, `body`, `attendees[]`, `location`, `is_online_meeting` | Creates Teams meeting if `is_online_meeting: true` |
| `outlook_update_event` | `user_id`, `event_id`, `subject`, `start`, `end`, `timezone`, `body`, `location`, `attendees[]`, `is_online_meeting` | Partial update — only provided fields are changed |
| `outlook_delete_event` | `user_id`, `event_id` | Permanently deletes the event |

---

### OneDrive (10 tools)

| Tool | Key Parameters | Notes |
|---|---|---|
| `onedrive_list_files` | `user_id`, `path` | Default path is `/` (root) |
| `onedrive_download_file` | `user_id`, `item_path` or `item_id` | Returns metadata + pre-authenticated download URL |
| `onedrive_upload_file` | `user_id`, `path`, `content` | `content` is base64-encoded |
| `onedrive_delete_item` | `user_id`, `item_path` or `item_id` | Moves to recycle bin |
| `onedrive_create_folder` | `user_id`, `folder_name`, `parent_path` | Conflict behavior: rename |
| `onedrive_share_item` | `user_id`, `item_path` or `item_id`, `type`, `scope` | `type`: view/edit/embed; `scope`: anonymous/organization/users |
| `onedrive_move_item` | `user_id`, `item_id` or `item_path`, `destination_path` or `destination_id` | Resolves destination folder ID automatically |
| `onedrive_rename_item` | `user_id`, `item_id` or `item_path`, `new_name` | |
| `onedrive_copy_item` | `user_id`, `item_id` or `item_path`, `destination_path` or `destination_id`, `new_name` | Async operation — returns immediately, copy completes in background |
| `onedrive_search` | `user_id`, `query`, `top` | Searches file names and content |

All paths are URL-encoded automatically — spaces and special characters are handled correctly.

---

### SharePoint (11 tools)

| Tool | Key Parameters | Notes |
|---|---|---|
| `sharepoint_list_sites` | `user_id`, `search` | Keyword search optional |
| `sharepoint_get_site` | `user_id`, `hostname`, `site_path` | e.g. `contoso.sharepoint.com`, `/sites/TeamSite` |
| `sharepoint_list_items` | `user_id`, `site_id`, `drive_id`, `path` | Lists document library contents |
| `sharepoint_download_file` | `user_id`, `site_id`, `item_id`, `drive_id` | Returns download URL |
| `sharepoint_upload_file` | `user_id`, `site_id`, `path`, `content`, `drive_id` | `content` is base64-encoded |
| `sharepoint_search` | `user_id`, `query`, `top` | Cross-site content search via Graph search API |
| `sharepoint_list_lists` | `user_id`, `site_id` | Lists all SP lists including custom lists and libraries |
| `sharepoint_list_list_items` | `user_id`, `site_id`, `list_id`, `top`, `expand_fields` | Returns list items with column values |
| `sharepoint_create_list_item` | `user_id`, `site_id`, `list_id`, `fields` | `fields` is a key/value object of column names |
| `sharepoint_update_list_item` | `user_id`, `site_id`, `list_id`, `item_id`, `fields` | Partial update — only provided fields changed |
| `sharepoint_delete_list_item` | `user_id`, `site_id`, `list_id`, `item_id` | |

---

### Teams (7 tools)

| Tool | Key Parameters | Notes |
|---|---|---|
| `teams_list_teams` | `user_id` | Lists all joined teams |
| `teams_list_channels` | `user_id`, `team_id` | Lists channels in a team |
| `teams_send_message` | `user_id`, `team_id`, `channel_id`, `message`, `content_type` | HTML or text |
| `teams_list_channel_messages` | `user_id`, `team_id`, `channel_id`, `top` | Requires `ChannelMessage.Read.All` |
| `teams_list_chats` | `user_id`, `top`, `include_members` | `include_members: true` requires `ChatMember.Read.All` — defaults to false |
| `teams_list_chat_messages` | `user_id`, `chat_id`, `top` | Lists messages in a 1:1 or group chat |
| `teams_send_chat_message` | `user_id`, `chat_id`, `message`, `content_type` | Requires `ChatMessage.Send` |

---

### To Do (8 tools)

| Tool | Key Parameters | Notes |
|---|---|---|
| `todo_list_task_lists` | `user_id` | Lists all task lists |
| `todo_list_tasks` | `user_id`, `list_id`, `top`, `filter` | OData filter supported |
| `todo_create_task` | `user_id`, `list_id`, `title`, `body`, `due_date`, `importance` | `due_date` format: `YYYY-MM-DD` |
| `todo_update_task` | `user_id`, `list_id`, `task_id`, `title`, `status`, `importance`, `due_date`, `body` | Partial update |
| `todo_complete_task` | `user_id`, `list_id`, `task_id` | Convenience wrapper — sets status to `completed` |
| `todo_delete_task` | `user_id`, `list_id`, `task_id` | |
| `todo_create_list` | `user_id`, `display_name` | Creates a new task list |
| `todo_delete_list` | `user_id`, `list_id` | |

Task `status` values: `notStarted`, `inProgress`, `completed`, `waitingOnOthers`, `deferred`

---

### Excel (7 tools)

Operates on Excel workbooks stored in OneDrive. Provide either `item_path` (OneDrive path) or `item_id`.

| Tool | Key Parameters | Notes |
|---|---|---|
| `excel_get_workbook_info` | `user_id`, `item_path` or `item_id` | Returns worksheets and named ranges |
| `excel_read_range` | `user_id`, `item_path` or `item_id`, `worksheet`, `range` | e.g. `range: "A1:D10"` |
| `excel_write_range` | `user_id`, `item_path` or `item_id`, `worksheet`, `range`, `values` | `values` is a 2D array |
| `excel_create_chart` | `user_id`, `item_path` or `item_id`, `worksheet`, `chart_type`, `source_range`, `chart_name` | Chart types: `ColumnClustered`, `Pie`, `Line`, etc. |
| `excel_add_table_rows` | `user_id`, `item_path` or `item_id`, `table_name`, `values` | Appends rows to an existing named table |
| `excel_create_worksheet` | `user_id`, `item_path` or `item_id`, `name` | |
| `excel_delete_worksheet` | `user_id`, `item_path` or `item_id`, `worksheet` | |

Paths with spaces (e.g. `My Documents/Budget 2026.xlsx`) are handled correctly via URL encoding.

---

### Users / Directory (7 tools)

| Tool | Key Parameters | Notes |
|---|---|---|
| `users_get_me` | `user_id` | Returns profile of the authenticated user |
| `users_get_user` | `user_id`, `target` | Look up any user by UPN or object ID |
| `users_list_users` | `user_id`, `top`, `filter` | OData filter supported |
| `users_search` | `user_id`, `query`, `top` | `startsWith` search on displayName and mail |
| `users_get_manager` | `user_id`, `target` | Omit `target` for current user's manager |
| `users_get_direct_reports` | `user_id`, `target` | Omit `target` for current user's direct reports |
| `users_get_photo` | `user_id`, `target`, `size` | Returns photo metadata + download endpoint. Sizes: `48x48` through `648x648` |

---

### Office Docs (2 tools)

| Tool | Key Parameters | Notes |
|---|---|---|
| `docs_get_content` | `user_id`, `item_path` or `item_id` | Extracts text content from Word/PDF documents |
| `docs_convert` | `user_id`, `item_path` or `item_id`, `format` | Convert documents (e.g. docx → pdf) |

---

## Architecture

```
POST /mcp  (JSON-RPC 2.0)
     │
     ▼
 __main__.py  ←  tool registry (auto-discovers TOOLS + HANDLERS from each module)
     │
     ├── tools/auth.py          (3 tools)
     ├── tools/outlook.py       (13 tools)
     ├── tools/onedrive.py      (10 tools)
     ├── tools/sharepoint.py    (11 tools)
     ├── tools/teams.py         (7 tools)
     ├── tools/todo.py          (8 tools)
     ├── tools/excel.py         (7 tools)
     ├── tools/users.py         (7 tools)
     └── tools/office_docs.py   (2 tools)
          │
          ▼
     clients/graph_client.py   ←  httpx-based Graph API client
          │
          ▼
     auth/oauth_web.py         ←  token retrieval + refresh
     auth/device_code.py       ←  MSAL device code flow
     auth/token_store_pg.py    ←  PostgreSQL token persistence
```

---

## Known Graph API Constraints

| Constraint | Detail |
|---|---|
| `$search` + `$filter` | Mutually exclusive on mail endpoints — use one or the other |
| `$search` + `$orderby` | Not supported — search results use relevance ranking |
| `contains()` in `$filter` | Not supported on mail — use `$search` for keyword matching |
| `$expand=members` on `/me/chats` | Requires `ChatMember.Read.All` — use `include_members: true` only if that permission is granted |
| OneDrive copy | Async — returns 202 immediately; copy completes in background |

---

## Deployment

The server is designed for Railway, Docker, or any container platform.

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY . .
RUN pip install -e .
CMD ["python", "-m", "m365_mcp"]
```

Set all environment variables in your platform's config. The server binds to `0.0.0.0:8080` by default.

---

## License

MIT
