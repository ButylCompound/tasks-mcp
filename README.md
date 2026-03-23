# todo-mcp

An MCP (Model Context Protocol) server that gives AI assistants access to **Microsoft To Do** via the Microsoft Graph API. Supports multiple accounts (personal and work), encrypted token storage, and all common task operations.

## Features

- List task lists and tasks across any number of accounts
- Create, update, complete, and delete tasks
- Get all pending tasks across every list in one call
- Multi-account support (e.g. separate `work` and `personal` aliases)
- Refresh tokens stored encrypted with AES-256-GCM on disk
- Non-blocking device code authentication flow (URL returned immediately)

## Prerequisites

- **Node.js** 18 or later
- An **Azure app registration** with the Microsoft Graph `Tasks.Read`, `Tasks.ReadWrite`, and `offline_access` delegated permissions (see below)

## 1. Create an Azure App Registration

1. Go to [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Name it anything (e.g. `My Todo MCP`)
3. Supported account types: choose **"Accounts in any organizational directory and personal Microsoft accounts"** (enables both work and personal accounts)
4. No redirect URI needed — click **Register**
5. Copy the **Application (client) ID** — you'll need it below
6. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**
   - Add: `Tasks.Read`, `Tasks.ReadWrite`, `offline_access`
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

## Environment Variables

| Variable | Required | Description |
|---|---|---|
| `TODO_MCP_MASTER_KEY` | Yes | Base64 key used to encrypt stored refresh tokens |
| `TODO_MCP_CLIENT_ID` | Yes* | Azure app client ID (*or hardcode in `auth.js`) |
| `TODO_MCP_TENANT_ID` | No | Default tenant ID (defaults to `"common"`) |
| `TODO_MCP_DEFAULT_ACCOUNT` | No | Default account alias (defaults to `"default"`) |
| `TODO_MCP_CONFIG_PATH` | No | Override path for the config file (defaults to `~/.todo-mcp-config.json`) |

## Available Tools

| Tool | Description |
|---|---|
| `list_accounts` | List all stored account aliases |
| `authenticate_account` | Start device code auth for a new or existing account |
| `set_default_account` | Change which account is used when none is specified |
| `delete_account` | Remove a stored account |
| `list_todo_lists` | List all task lists for an account |
| `get_tasks` | Get tasks from a specific list (filterable by status) |
| `get_all_pending_tasks` | Get all non-completed tasks across every list |
| `create_task` | Create a task with optional due date, importance, and body |
| `update_task` | Update a task's title, due date, importance, or body |
| `complete_task` | Mark a task as completed |
| `delete_task` | Delete a task |

All data tools accept an optional `account` parameter to target a specific alias. If omitted, the default account is used.

## Token Storage

Refresh tokens are stored in `~/.todo-mcp-config.json` encrypted with AES-256-GCM using a key derived from `TODO_MCP_MASTER_KEY`. The file is created with mode `0600` (owner read/write only). **Never commit this file.**
