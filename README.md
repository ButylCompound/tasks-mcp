# todo-mcp

An MCP (Model Context Protocol) server that gives AI assistants access to **Microsoft To Do**, **Microsoft Planner**, and **Outlook Calendars** via the Microsoft Graph API. Supports multiple accounts (personal and work), encrypted token storage, and common planning/task/calendar operations.

## Features

- List task lists and tasks across any number of accounts
- Create, update, complete, and delete tasks
- Get all pending tasks across every list in one call
- List Planner plans, buckets, and tasks
- Create, update, complete, and delete Planner tasks
- List Outlook calendars and events (events older than one month are filtered out by default)
- Create, update, and delete Outlook calendar events for user and group calendars
- Manage Planner task details and checklist items
- Multi-account support (e.g. separate `work` and `personal` aliases)
- Refresh tokens stored encrypted with AES-256-GCM on disk
- Non-blocking device code authentication flow (URL returned immediately)

## Prerequisites

- **Node.js** 18 or later
- An **Azure app registration** with the Microsoft Graph `Tasks.ReadWrite`, `Calendars.ReadWrite`, `Group.ReadWrite.All`, `User.ReadBasic.All`, and `offline_access` delegated permissions (see below)

## 1. Create an Azure App Registration

1. Go to [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Name it anything (e.g. `My Todo MCP`)
3. Supported account types: choose **"Accounts in any organizational directory and personal Microsoft accounts"** (enables both work and personal accounts)
4. No redirect URI needed — click **Register**
5. Copy the **Application (client) ID** — you'll need it below
6. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**
  - Add: `Tasks.ReadWrite`, `Calendars.ReadWrite`, `Group.ReadWrite.All`, `User.ReadBasic.All`, `offline_access`
  - Click **Grant admin consent** (or users will be prompted on first sign-in)
7. Go to **Authentication** → **Advanced settings** → enable **"Allow public client flows"** (required for device code)
8. Go to **Manifest** → set `"requestedAccessTokenVersion": 2` in the `api` section → **Save**

## 2. Install

```bash
git clone <repo-url> todo-mcp
cd todo-mcp
npm install
```

## 3. Configure

### Option A — Claude Desktop / Cowork (`claude_desktop_config.json`)

Add this to `%APPDATA%\Claude\claude_desktop_config.json` (Windows) or `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS):

```json
{
  "mcpServers": {
    "todo-mcp": {
      "command": "node",
      "args": ["C:/path/to/todo-mcp/index.js"],
      "env": {
        "TODO_MCP_CLIENT_ID": "<your-client-id>",
        "TODO_MCP_MASTER_KEY": "<random-base64-key>"
      }
    }
  }
}
```

Generate a secure master key (PowerShell):
```powershell
[System.Convert]::ToBase64String([System.Security.Cryptography.RandomNumberGenerator]::GetBytes(32))
```

Or on macOS/Linux:
```bash
openssl rand -base64 32
```

### Option B — VS Code (`.vscode/mcp.json`)

The included `.vscode/mcp.json` prompts for the master key securely on first use. Set `TODO_MCP_CLIENT_ID` in your environment or add it to the `env` block.

## 4. Authenticate Accounts

Once the server is running, use the `authenticate_account` tool from your AI assistant. It returns a device code URL immediately:

```
authenticate_account(account="work", tenant_id="<your-org-tenant-id>")
authenticate_account(account="personal", tenant_id="consumers")
```

Go to the URL shown, enter the code, and sign in. The token is saved automatically in the background.

For **personal Microsoft accounts** (Outlook, Hotmail, Live), always pass `tenant_id="consumers"`.  
For **work/school accounts**, pass your organization's tenant ID (a UUID), or `"organizations"` to accept any.

> Microsoft Planner is available for work/school Microsoft 365 accounts. Personal Microsoft accounts do not support Planner APIs.

## Environment Variables

| Variable | Required | Description |
|---|---|---|
| `TODO_MCP_MASTER_KEY` | Yes | Base64 key used to encrypt stored refresh tokens |
| `TODO_MCP_CLIENT_ID` | Yes* | Azure app client ID (*or hardcode in `auth.js`) |
| `TODO_MCP_TENANT_ID` | No | Default tenant ID (defaults to `"common"`) |
| `TODO_MCP_DEFAULT_ACCOUNT` | No | Default account alias (defaults to `"default"`) |
| `TODO_MCP_CONFIG_PATH` | No | Override path for the config file (defaults to `~/.todo-mcp-config.json`) |
| `TODO_MCP_SCOPE` | No | OAuth scope override (defaults to `Tasks.ReadWrite Calendars.ReadWrite Group.ReadWrite.All User.ReadBasic.All offline_access`) |

## Available Tools

| Tool | Description |
|---|---|
| `list_accounts` | List all stored account aliases |
| `authenticate_account` | Start device code auth for a new or existing account |
| `set_default_account` | Change which account is used when none is specified |
| `delete_account` | Remove a stored account |
| `list_todo_lists` | List all task lists for an account |
| `get_tasks` | Get tasks from a specific list (defaults to pending; completed omitted unless explicitly requested) |
| `get_all_pending_tasks` | Get all non-completed tasks across every list |
| `create_task` | Create a task with optional due date, importance, and body |
| `update_task` | Update a task's title, due date, importance, or body |
| `complete_task` | Mark a task as completed |
| `delete_task` | Delete a task |
| `list_planner_plans` | List Planner plans available to the user or a specific group |
| `list_planner_buckets` | List buckets in a Planner plan |
| `list_employee_ids` | List employee IDs with names for assignment lookups |
| `get_user_planner_tasks` | Get pending tasks assigned to a specific user from plans visible to the signed-in user (optionally filtered by plan) |
| `get_planner_tasks` | Get tasks in a Planner plan (defaults to incomplete; completed omitted unless explicitly requested) |
| `get_all_pending_planner_tasks` | Get all not-completed Planner tasks assigned to the user |
| `create_planner_task` | Create a Planner task with optional assignees/categories |
| `update_planner_task` | Update Planner task fields, assignees, categories, progress |
| `complete_planner_task` | Mark a Planner task as completed |
| `delete_planner_task` | Delete a Planner task |
| `get_planner_task_details` | Get Planner task details (description, checklist, references) |
| `update_planner_task_details` | Update Planner task details |
| `add_planner_checklist_item` | Add a checklist item to a Planner task |
| `update_planner_checklist_item` | Update a Planner checklist item |
| `delete_planner_checklist_item` | Delete a Planner checklist item |
| `list_calendars` | List Outlook calendars for the signed-in user |
| `list_group_calendars` | List Outlook calendars from accessible Microsoft 365 groups |
| `list_calendar_events` | List Outlook calendar events (filters out events older than one month) |
| `create_calendar_event` | Create an Outlook calendar event |
| `update_calendar_event` | Update an Outlook calendar event |
| `delete_calendar_event` | Delete an Outlook calendar event |

All data tools accept an optional `account` parameter to target a specific alias. If omitted, the default account is used.

## Token Storage

Refresh tokens are stored in `~/.todo-mcp-config.json` encrypted with AES-256-GCM using a key derived from `TODO_MCP_MASTER_KEY`. The file is created with mode `0600` (owner read/write only). **Never commit this file.**
