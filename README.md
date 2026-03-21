[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/ryaker-outlook-mcp-badge.png)](https://mseep.ai/app/ryaker-outlook-mcp)

# M365 Assistant MCP Server

A comprehensive MCP (Model Context Protocol) server that connects AI assistants with Microsoft 365 services through the Microsoft Graph API and Power Automate API. Works with any MCP-compatible client including Claude Desktop, VS Code (GitHub Copilot), Cursor, Windsurf, and **OpenWebUI**.

## Supported Services

- **Outlook** - Email, calendar, folders, and rules
- **OneDrive** - Files, folders, search, and sharing
- **Power Automate** - Flows, environments, and run history

## Directory Structure

```
├── index.js                 # Main entry point (all tools)
├── server.js                # Shared MCP server factory
├── server-email.js          # Email-only server entry point
├── server-calendar.js       # Calendar-only server entry point
├── server-onedrive.js       # OneDrive-only server entry point
├── server-power-automate.js # Power Automate-only server entry point
├── config.js                # Configuration settings
├── auth/                    # Authentication modules
│   ├── index.js             # Authentication exports
│   ├── token-manager.js     # Token storage and refresh (Graph + Flow)
│   └── tools.js             # Auth-related tools
├── calendar/                # Calendar functionality
│   ├── index.js             # Calendar exports
│   ├── list.js              # List events
│   ├── create.js            # Create event
│   ├── delete.js            # Delete event
│   ├── cancel.js            # Cancel event
│   ├── accept.js            # Accept event
│   └── decline.js           # Decline event
├── email/                   # Email functionality
│   ├── index.js             # Email exports
│   ├── list.js              # List emails
│   ├── search.js            # Search emails
│   ├── read.js              # Read email
│   ├── send.js              # Send email
│   ├── draft.js             # Create email draft
│   ├── mark-as-read.js      # Mark email read/unread
│   ├── attachments.js       # List/download attachments
│   └── folder-utils.js      # Folder lookup utilities
├── folder/                  # Folder functionality
│   ├── index.js             # Folder exports
│   ├── list.js              # List folders
│   ├── create.js            # Create folder
│   └── move.js              # Move emails
├── rules/                   # Email rules functionality
│   ├── index.js             # Rules exports
│   ├── list.js              # List rules
│   └── create.js            # Create rule
├── onedrive/                # OneDrive functionality
│   ├── index.js             # OneDrive exports
│   ├── list.js              # List files/folders
│   ├── search.js            # Search files
│   ├── download.js          # Get download URL
│   ├── upload.js            # Simple upload (<4MB)
│   ├── upload-large.js      # Chunked upload (>4MB)
│   ├── share.js             # Create sharing link
│   └── folder.js            # Create/delete folders
├── power-automate/          # Power Automate functionality
│   ├── index.js             # Power Automate exports
│   ├── flow-api.js          # Flow API client
│   ├── list-environments.js # List environments
│   ├── list-flows.js        # List flows
│   ├── run-flow.js          # Trigger flow
│   ├── list-runs.js         # Run history
│   └── toggle-flow.js       # Enable/disable flow
└── utils/                   # Utility functions
    ├── graph-api.js         # Microsoft Graph API helper
    ├── odata-helpers.js     # OData query building
    ├── html-sanitizer.js    # HTML email sanitization
    ├── metadata-sanitizer.js # Metadata sanitization
    └── mock-data.js         # Test mode data
```

## Features

- **Authentication**: OAuth 2.0 authentication with Microsoft Graph API (+ Flow API for Power Automate)
- **Email Management**: List, search, read, send, draft, and organize emails with attachment support
- **Calendar Management**: List, create, accept, decline, and delete calendar events
- **OneDrive Integration**: List, search, upload, download, and share files
- **Power Automate**: List environments/flows, trigger flows, view run history
- **Modular Structure**: Clean separation of concerns for maintainability
- **Test Mode**: Simulated responses for testing without real API calls

## Available Tools

Each tool category below corresponds to a **split server entry point** that can be enabled independently. See [Selective Service Configuration](#selective-service-configuration-split-servers) for details.

### Authentication (included in all servers)
| Tool | Description |
|------|-------------|
| `about` | Server info and capabilities |
| `authenticate` | Authenticate with Microsoft Graph API |
| `check-auth-status` | Check authentication status |

### Email (`server-email.js`)
| Tool | Description |
|------|-------------|
| `list-emails` | List recent emails from inbox |
| `search-emails` | Search emails with filters |
| `read-email` | Read email content |
| `send-email` | Send a new email |
| `draft-email` | Create and save an email draft |
| `mark-as-read` | Mark email as read/unread |
| `list-email-attachments` | List all attachments for a specific email |
| `download-email-attachment` | Download a specific attachment from an email |
| `download-email-attachments` | Download all attachments from an email (optionally save to disk) |
| `list-folders` | List mail folders |
| `create-folder` | Create mail folder |
| `move-emails` | Move emails between folders |
| `list-rules` | List inbox rules |
| `create-rule` | Create inbox rule |
| `edit-rule-sequence` | Change rule execution order |

### Calendar (`server-calendar.js`)
| Tool | Description |
|------|-------------|
| `list-events` | List calendar events |
| `create-event` | Create calendar event |
| `accept-event` | Accept event invitation |
| `decline-event` | Decline event invitation |
| `cancel-event` | Cancel a calendar event |
| `delete-event` | Delete calendar event |

### OneDrive (`server-onedrive.js`)
| Tool | Description |
|------|-------------|
| `onedrive-list` | List files in a path |
| `onedrive-search` | Search files by query |
| `onedrive-download` | Get download URL |
| `onedrive-upload` | Upload small file (<4MB) |
| `onedrive-upload-large` | Chunked upload (>4MB) |
| `onedrive-share` | Create sharing link |
| `onedrive-create-folder` | Create folder |
| `onedrive-delete` | Delete file or folder |

### Power Automate (`server-power-automate.js`)
| Tool | Description |
|------|-------------|
| `flow-list-environments` | List Power Platform environments |
| `flow-list` | List flows in environment |
| `flow-run` | Trigger a manual flow |
| `flow-list-runs` | Get flow run history |
| `flow-toggle` | Enable/disable a flow |

## Quick Start

1. **Install dependencies**: `npm install`
2. **Azure setup**: Register app in Azure Portal (see detailed steps below)
3. **Configure environment**: Copy `.env.example` to `.env` and add your Azure credentials
4. **Configure your MCP client**: Add the server to Claude Desktop, VS Code, Cursor, Windsurf, or any MCP-compatible client
5. **Start auth server**: `npm run auth-server`
6. **Authenticate**: Use the authenticate tool in your MCP client to get the OAuth URL
7. **Start using**: Access your M365 data through your AI assistant!

## Installation

### Prerequisites
- Node.js 14.0.0 or higher
- npm or yarn package manager
- Azure account for app registration

### Install Dependencies

```bash
npm install
```

## Azure App Registration & Configuration

### App Registration

1. Open [Azure Portal](https://portal.azure.com/)
2. Search for "App registrations"
3. Click "New registration"
4. Name: "M365 MCP Server"
5. Account type: "Accounts in any organizational directory and personal Microsoft accounts"
6. Redirect URI: Web → `http://localhost:3333/auth/callback`
7. Click "Register"
8. Copy the "Application (client) ID" for your `.env` file

### App Permissions

1. Go to "API permissions" under Manage
2. Click "Add a permission" → "Microsoft Graph" → "Delegated permissions"
3. Add these permissions:
   - `offline_access`
   - `User.Read`
   - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.Read`, `Calendars.ReadWrite`
   - `Files.Read`, `Files.ReadWrite`
4. Click "Add permissions"

**For Power Automate** (optional):
- Requires additional Azure AD configuration with Flow API scope
- See Power Automate section below for details

### Client Secret

1. Go to "Certificates & secrets" → "Client secrets"
2. Click "New client secret"
3. Add description and select expiration
4. **Copy the VALUE** (not the Secret ID)

## Configuration

### 1. Environment Variables

```bash
cp .env.example .env
```

Edit `.env`:
```bash
# Get these values from Azure Portal > App Registrations > Your App
MS_CLIENT_ID=your-application-client-id-here
MS_CLIENT_SECRET=your-client-secret-VALUE-here
MS_TENANT_ID=your-tenant-id-here
USE_TEST_MODE=false
```

**Important Notes:**
- Use `MS_CLIENT_ID` and `MS_CLIENT_SECRET` in the `.env` file
- Set `MS_TENANT_ID` for single-tenant apps to avoid `/common` endpoint errors
- For MCP client configs, you'll use `OUTLOOK_CLIENT_ID` and `OUTLOOK_CLIENT_SECRET`
- Always use the client secret **VALUE**, never the Secret ID

### 2. MCP Client Configuration

This server works with any MCP-compatible client. Below are setup instructions for popular platforms.

> **Stdio vs HTTP transports**: The default entry point (`index.js`) uses **stdio** transport — each AI client manages the process directly. For **OpenWebUI** (and any browser-based or remote client), use the HTTP server (`sse-server.js` / `npm run start:http`) which exposes a Streamable HTTP endpoint at `/mcp`. See the [OpenWebUI section](#openwebui) below for full setup instructions.

#### Generic MCP Client (Command + ENV UI)

Many MCP clients provide a simple form with a **Command** field and **ENV** key/value pairs. Use the following values:

**Command:**
```
node /path/to/outlook-mcp/index.js
```

**Environment Variables:**
| Key | Value |
|-----|-------|
| `OUTLOOK_CLIENT_ID` | Your Azure application client ID |
| `OUTLOOK_CLIENT_SECRET` | Your Azure client secret VALUE |
| `USE_TEST_MODE` | `false` |

#### Claude Desktop

Add to your Claude Desktop config (`claude_desktop_config.json`). A ready-to-edit sample is provided in [`claude-config-sample.json`](claude-config-sample.json).

```json
{
  "mcpServers": {
    "m365-assistant": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/index.js"],
      "env": {
        "USE_TEST_MODE": "false",
        "OUTLOOK_CLIENT_ID": "your-client-id",
        "OUTLOOK_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

#### VS Code (GitHub Copilot)

Add to `.vscode/mcp.json` in your workspace, or open the user-level config via the Command Palette → `MCP: Open User Configuration`. A ready-to-edit sample is provided in [`vscode-config-sample.json`](vscode-config-sample.json).

```json
{
  "servers": {
    "m365-assistant": {
      "type": "stdio",
      "command": "node",
      "args": ["/path/to/outlook-mcp/index.js"],
      "env": {
        "USE_TEST_MODE": "false",
        "OUTLOOK_CLIENT_ID": "your-client-id",
        "OUTLOOK_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

#### Cursor

Add to `~/.cursor/mcp.json` (global) or `.cursor/mcp.json` (project-level), or go to **Settings → Tools & MCP → Add Custom MCP**. A ready-to-edit sample is provided in [`cursor-config-sample.json`](cursor-config-sample.json).

```json
{
  "mcpServers": {
    "m365-assistant": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/index.js"],
      "env": {
        "USE_TEST_MODE": "false",
        "OUTLOOK_CLIENT_ID": "your-client-id",
        "OUTLOOK_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

#### OpenWebUI

OpenWebUI (≥ 0.6.6) supports MCP natively over **Streamable HTTP**. Instead of the stdio-based `index.js`, run the dedicated HTTP server (`sse-server.js`) which serves all tools over HTTP on port 3001 (configurable).

**Step 1 — Start the HTTP server**

```bash
# Copy and edit the .env file with your Microsoft credentials
cp .env.example .env   # or create one manually

MS_CLIENT_ID=your-client-id
MS_CLIENT_SECRET=your-client-secret
MS_TENANT_ID=common       # or your specific tenant ID

# Start the server (default: http://0.0.0.0:3001)
npm run start:http
```

> To change the port: `MCP_HTTP_PORT=8080 npm run start:http`

**Step 2 — Authenticate with Microsoft**

Open `http://localhost:3001/auth` in your browser to complete the OAuth flow. Once authenticated, the token is stored and all MCP requests will be automatically authorized.

**Step 3 — Add the server to OpenWebUI**

1. Go to **Admin Panel → Settings → Tools → Add MCP Server**
2. Fill in the form:
   - **Name**: `M365 Assistant`
   - **URL**: `http://your-server-address:3001/mcp`
   - **Type**: `MCP (Streamable HTTP)`
3. Click **Save**.

The M365 tools will now appear in OpenWebUI's tool list and can be used in chat sessions.

> **Note on origins**: By default the server allows requests from any origin. In a production deployment, set `MCP_ALLOWED_ORIGIN=https://your-openwebui-host` to restrict cross-origin access.

#### Windsurf

Add to `~/.codeium/windsurf/mcp_config.json`, or go to **Settings → Cascade → MCP Servers → Add custom server**. A ready-to-edit sample is provided in [`windsurf-config-sample.json`](windsurf-config-sample.json).

```json
{
  "mcpServers": {
    "m365-assistant": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/index.js"],
      "env": {
        "USE_TEST_MODE": "false",
        "OUTLOOK_CLIENT_ID": "your-client-id",
        "OUTLOOK_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

### Selective Service Configuration (Split Servers)

By default, `index.js` loads all 37 tools. If you only need a subset of Microsoft 365 services, you can use the **focused server entry points** instead. This reduces context window usage and avoids exposing unnecessary tools to the AI assistant.

| Entry Point | Server Name | Tools | Definition Size | ~Tokens |
|-------------|-------------|:-----:|:---------:|:-------:|
| `index.js` | m365-assistant | 37 | 11.7 KB | ~2,934 |
| `server-email.js` | m365-email | 18 | 6.1 KB | ~1,530 |
| `server-calendar.js` | m365-calendar | 9 | 2.1 KB | ~533 |
| `server-onedrive.js` | m365-onedrive | 11 | 3.1 KB | ~770 |
| `server-power-automate.js` | m365-power-automate | 8 | 1.8 KB | ~459 |

> **Context window note:** The "~Tokens" column shows the approximate token cost of just the tool definitions (at ~4 chars/token). Your model's context window must have room for these definitions *plus* the conversation. Recommended minimum context window: **8,192 tokens** for all tools, **4,096 tokens** for any single split server.

**To enable only the services you need**, add separate MCP server entries to your client config — one for each service. Simply omit any services you don't want. For example, to enable only Email and Calendar in Claude Desktop:

```json
{
  "mcpServers": {
    "m365-email": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/server-email.js"],
      "env": {
        "OUTLOOK_CLIENT_ID": "your-client-id",
        "OUTLOOK_CLIENT_SECRET": "your-client-secret"
      }
    },
    "m365-calendar": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/server-calendar.js"],
      "env": {
        "OUTLOOK_CLIENT_ID": "your-client-id",
        "OUTLOOK_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

Each server shares the same `.env` file, credentials, and token storage — no extra configuration needed. You can also start individual servers via npm scripts:

```bash
npm run start:email           # Email, folders, rules only
npm run start:calendar        # Calendar only
npm run start:onedrive        # OneDrive only
npm run start:power-automate  # Power Automate only
npm start                     # All tools (default)
```

## Authentication

### Graph API (Outlook + OneDrive)

1. Start auth server: `npm run auth-server`
2. Use the `authenticate` tool in your MCP client
3. Visit the provided URL and sign in
4. Tokens saved to `~/.outlook-mcp-tokens.json`

### Power Automate (Optional)

Power Automate requires a separate token with the Flow API scope. Configure additional Azure AD permissions for `https://service.flow.microsoft.com//.default` scope.

**Limitations:**
- Only solution-aware flows are accessible
- Only manual trigger flows can be run via API
- Requires environment ID for most operations

## Troubleshooting

### Common Issues

**"Cannot find module"**
```bash
npm install
```

**"Port 3333 in use"**
```bash
npx kill-port 3333
npm run auth-server
```

**"Invalid client secret" (AADSTS7000215)**
- Use the secret **VALUE**, not the Secret ID

**"Authentication required"**
- Delete `~/.outlook-mcp-tokens.json` and re-authenticate

**Agent only sees some tools / missing email or calendar tools**
- This server exposes 37 tools. By default, all tools are returned in a single `tools/list` response (no pagination).
- Some MCP clients (e.g., Claude.ai) do not follow `nextCursor` pagination, so the default is to return everything at once.
- If you previously set `MCP_TOOLS_PAGE_SIZE` to a small number (e.g., 10), the client may only see the first page of tools. Remove the variable or set it to `0` to disable pagination:
  ```bash
  # In your .env file or environment:
  MCP_TOOLS_PAGE_SIZE=0
  ```
- **Minimum recommended context window: 8,192 tokens** (works well for all 37 tools)
- **Minimum functional context window: 4,096 tokens** (tools load but very little room for conversation)
- If your agent is missing tools, increase the model's context window or max tool tokens setting in your client configuration.

## Testing

```bash
# Run with MCP Inspector
npm run inspect

# Run in test mode (mock data)
npm run test-mode

# Run Jest tests
npm test
```

## Extending the Server

1. Create new module directory
2. Implement tool handlers in separate files
3. Export tool definitions from module index
4. Import and add to `TOOLS` array in `index.js`
