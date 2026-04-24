# msgraphmcp

[![CI](https://github.com/DustHoff/msgraphmcp/actions/workflows/ci.yml/badge.svg)](https://github.com/DustHoff/msgraphmcp/actions/workflows/ci.yml)
[![Docker](https://github.com/DustHoff/msgraphmcp/actions/workflows/docker.yml/badge.svg)](https://github.com/DustHoff/msgraphmcp/actions/workflows/docker.yml)
[![GHCR](https://img.shields.io/badge/ghcr.io-msgraphmcp-blue?logo=github)](https://github.com/DustHoff/msgraphmcp/pkgs/container/msgraphmcp)

**MCP Server for the Microsoft Graph API** — runs as a container, exposes **140+ tools** across all major Microsoft 365 workloads to any MCP-compatible client such as [Claude Code](https://claude.ai/code), and supports four authentication modes: **authorization code + PKCE** (delegated, recommended for Kubernetes/HTTP), **client secret** (app-only), **client certificate** (app-only, recommended for production), and **device code** (interactive, local use).

---

## Table of Contents

- [Architecture](#architecture)
- [Prerequisites](#prerequisites)
- [Azure App Registration](#azure-app-registration)
- [Quick Start](#quick-start)
  - [Local (Node.js)](#local-nodejs)
  - [Docker](#docker)
  - [Claude Code Integration](#claude-code-integration)
- [Environment Variables](#environment-variables)
- [Authentication Flow](#authentication-flow)
- [Tool Reference](#tool-reference)
  - [Users](#users)
  - [Mail](#mail)
  - [Calendar](#calendar)
  - [OneDrive / Files](#onedrive--files)
  - [Groups](#groups)
  - [Microsoft Teams](#microsoft-teams)
  - [Contacts](#contacts)
  - [To Do / Tasks](#to-do--tasks)
  - [SharePoint Sites](#sharepoint-sites)
  - [Intune — Apps](#intune--apps)
  - [Intune — Device Configurations](#intune--device-configurations)
  - [Intune — Settings Catalog](#intune--settings-catalog)
  - [Intune — Compliance Policies](#intune--compliance-policies)
  - [Intune — Managed Devices](#intune--managed-devices)
  - [Intune — Device Diagnostics](#intune--device-diagnostics)
  - [Intune — Notification Templates](#intune--notification-templates)
  - [Auth](#auth)
- [Development](#development)
- [CI/CD and Docker Registry](#cicd-and-docker-registry)
- [Security Notes](#security-notes)

---

## Architecture

```
Claude Code (MCP client)           Another MCP client
       │  stdio / HTTP                    │  HTTP
       ▼                                  ▼
  msgraphmcp (Node.js / TypeScript)
  ├── auth/TokenManager     Four auth modes — per-session in HTTP mode
  ├── graph/GraphClient     Axios wrapper, user-stamped logs, auto-pagination, single retry on 401
  └── tools/
      ├── users · mail · calendar · files · groups
      ├── teams · contacts · tasks · sites
      └── intune (apps · device config · settings catalog · compliance · managed devices)
       │  HTTPS
       ▼
  Microsoft Graph API  (https://graph.microsoft.com/v1.0)

HTTP mode auth flow (authorization code, per-session):
  1. MCP client connects  →  server creates session with isolated TokenManager
  2. Tool call without auth  →  401 + loginUrl (/auth/login?token=<one-time-token>)
  3. Browser → GET /auth/login?token=<one-time-token>  →  Microsoft Login
  4. GET /auth/callback  →  token stored in that session's TokenManager only
  5. Subsequent tool calls from the same session use that user's token exclusively

The login token is a single-use 256-bit value that maps server-side to the
MCP session — the session ID itself never appears in any URL, browser
history, or proxy log.
```

In HTTP mode every MCP session gets its own isolated `TokenManager` — tokens are never shared between sessions. Each user authenticates independently. Tokens are kept in-memory for the lifetime of the session; no cross-user token bleed is possible even when multiple clients are connected simultaneously.

---

## Prerequisites

| Requirement | Version |
|---|---|
| Node.js | ≥ 20 |
| npm | ≥ 10 |
| Docker | ≥ 24 (optional) |
| Microsoft 365 / Azure AD tenant | — |
| Azure App Registration | see below |

---

## Azure App Registration

1. Open [Azure Portal → App registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) → **New registration**.
2. **Name**: `msgraphmcp` (or any name)
3. **Supported account types**: *Accounts in this organizational directory only* (or *multitenant* if needed)
4. **Redirect URI** — depends on the auth mode you plan to use:
   - **Authorization code flow (Mode A):** select **Web** → enter your callback URL, e.g. `https://msgraph.example.com/auth/callback`
   - **Device code flow (Mode D):** select **Mobile and desktop applications** → `https://login.microsoftonline.com/common/oauth2/nativeclient`
   - **App-only (Modes B/C):** no redirect URI needed
5. Click **Register** — copy the **Application (client) ID** and **Directory (tenant) ID**.
6. **Add permissions** — the type depends on the auth mode you choose:

   - **Authorization code / device code (delegated):** Go to **API Permissions → Add a permission → Microsoft Graph → Delegated permissions** and add the permissions below, then click **Grant admin consent**.
   - **Client secret / certificate (app-only):** Go to **API Permissions → Add a permission → Microsoft Graph → Application permissions** and add the same permissions, then click **Grant admin consent**.  
     > With app-only auth, `userId: "me"` does not resolve — use explicit UPNs or object IDs in all tool calls.

| Permission | Purpose |
|---|---|
| `User.ReadWrite.All` | Manage users |
| `Group.ReadWrite.All` | Manage groups |
| `GroupMember.ReadWrite.All` | Manage group membership |
| `Mail.ReadWrite` | Read/write mailboxes |
| `Mail.Send` | Send email |
| `Calendars.ReadWrite` | Manage calendars & events |
| `Files.ReadWrite.All` | OneDrive CRUD |
| `Sites.ReadWrite.All` | SharePoint CRUD |
| `Tasks.ReadWrite` | Microsoft To Do |
| `Contacts.ReadWrite` | Contacts |
| `Team.ReadWrite.All` | Teams management |
| `Channel.ReadWrite.All` | Teams channels |
| `ChannelMessage.Send` | Send Teams messages |
| `Directory.ReadWrite.All` | Directory objects |
| `DeviceManagementApps.ReadWrite.All` | Intune apps |
| `DeviceManagementConfiguration.ReadWrite.All` | Intune device configs |
| `DeviceManagementManagedDevices.ReadWrite.All` | Managed devices |
| `DeviceManagementServiceConfig.ReadWrite.All` | Intune service config |

7. Click **Grant admin consent for \<your tenant\>**.

> **Tip:** You can restrict the scope by setting the `GRAPH_SCOPES` environment variable to only the permissions you actually need.

---

## Quick Start

### Local (Node.js)

```bash
git clone https://github.com/DustHoff/msgraphmcp.git
cd msgraphmcp
npm install
npm run build

export AZURE_CLIENT_ID="your-client-id"
export AZURE_TENANT_ID="your-tenant-id"
export TOKEN_CACHE_PATH="$HOME/.msgraphmcp/tokens.json"

node dist/index.js
```

On first run the device code authentication prompt appears on **stderr**:

```
============================================================
To sign in, use a web browser to open the page
https://microsoft.com/devicelogin and enter the code XXXXXXXX to authenticate.
============================================================
```

After successful authentication the MCP server is ready and keeps the token refreshed automatically.

### Docker

```bash
cp .env.example .env
# Edit .env with your AZURE_CLIENT_ID and AZURE_TENANT_ID

docker-compose up
```

`docker-compose.yml` mounts a named volume (`token-cache`) at `/data` so the token cache survives container restarts.

**First run (one-time device code):**

```bash
docker-compose run --rm msgraphmcp
# Follow the authentication prompt on screen, then Ctrl+C
# Subsequent runs will use the cached refresh token
docker-compose up -d
```

### Claude Code Integration

Add the server to your project's `.claude/settings.json`:

```jsonc
{
  "mcpServers": {
    "msgraphmcp": {
      "command": "node",
      "args": ["dist/index.js"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_TENANT_ID": "your-tenant-id",
        "TOKEN_CACHE_PATH": "/path/to/tokens.json"
      }
    }
  }
}
```

Or using the pre-built Docker image from GHCR:

```jsonc
{
  "mcpServers": {
    "msgraphmcp": {
      "command": "docker",
      "args": [
        "run", "--rm", "-i",
        "-v", "msgraphmcp-tokens:/data",
        "-e", "AZURE_CLIENT_ID=your-client-id",
        "-e", "AZURE_TENANT_ID=your-tenant-id",
        "ghcr.io/DustHoff/msgraphmcp:latest"
      ]
    }
  }
}
```

Restart Claude Code after editing the config — the server appears in the MCP tools panel.

---

## Environment Variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `AZURE_CLIENT_ID` | **Yes** | — | App Registration client ID |
| `AZURE_TENANT_ID` | No | `common` | Tenant ID or `common` for multi-tenant |
| `AZURE_CLIENT_SECRET` | No | — | Client secret — required for **Mode A** (auth code) and **Mode B** (app-only) |
| `AZURE_REDIRECT_URI` | No | — | **Auth mode A** — full callback URL, e.g. `https://msgraph.example.com/auth/callback`. When set together with `AZURE_CLIENT_SECRET`, activates the authorization code flow |
| `AZURE_CLIENT_CERTIFICATE_PATH` | No | — | **Auth mode C** — path to a PEM private key file; must be set together with `THUMBPRINT` |
| `AZURE_CLIENT_CERTIFICATE_THUMBPRINT` | No | — | **Auth mode C** — SHA-256 certificate thumbprint (64 hex chars) from the App Registration |
| `GRAPH_SCOPES` | No | all scopes | Space-separated delegated scopes (used in auth code and device code modes) |
| `TOKEN_CACHE_PATH` | No | `/data/tokens.json` | Path to the MSAL token cache file |
| `PORT` | No | — | When set, the server listens on HTTP (Kubernetes mode); otherwise uses stdio |
| `LOG_LEVEL` | No | `info` | Log verbosity: `debug`, `info`, `warn`, `error` |
| `MAX_SESSIONS` | No | `50` | Maximum concurrent MCP sessions in HTTP mode. New connections beyond this limit receive HTTP 503. |
| `SESSION_IDLE_TIMEOUT_MINUTES` | No | `60` | Minutes of inactivity before an MCP session is automatically closed and removed. |

---

## Authentication Flow

Four modes are selected automatically by which environment variables are set:

| Mode | Env vars set | Type | User context |
|---|---|---|---|
| **A — Authorization Code** | `AZURE_CLIENT_SECRET` + `AZURE_REDIRECT_URI` | Delegated | Yes — full `userId: "me"` support |
| **B — Client Secret** | `AZURE_CLIENT_SECRET` (no redirect URI) | App-only | No — use explicit UPNs/object IDs |
| **C — Client Certificate** | `AZURE_CLIENT_CERTIFICATE_PATH` + `THUMBPRINT` | App-only | No |
| **D — Device Code** | none of the above | Delegated | Yes |

### Mode A — Authorization Code + PKCE (delegated, recommended for HTTP/Kubernetes)

Set `AZURE_CLIENT_SECRET` **and** `AZURE_REDIRECT_URI`. Each MCP session authenticates independently — tokens are isolated per session and kept in-memory. Suitable for containers with Entra ID Conditional Access — CA compliance is evaluated against the **user's browser device**, not the container. Multiple users can be authenticated simultaneously.

**Prerequisites:**
- Register `AZURE_REDIRECT_URI` (e.g. `https://msgraph.example.com/auth/callback`) as a **Web** redirect URI in the Entra ID app registration.
- Grant **Delegated** permissions (not Application) + admin consent.

#### Authentication flow

```
Step 1 — Connect your MCP client (Claude Code or other)
  POST /mcp  (initialize)
  └─► Server creates a session with its own isolated TokenManager
  └─► Session ID returned in response header:  mcp-session-id: <uuid>

Step 2 — Authenticate (once per session / pod restart)
  Tool call without auth → 401 response containing a fresh loginUrl
  GET /auth/login?token=<one-time-token>
  └─► Server resolves the one-time token → session, invalidates the token
  └─► Server generates PKCE code_verifier + S256 challenge
  └─► Redirects browser → Microsoft login page
  └─► User authenticates + consents
  └─► Microsoft redirects → GET /auth/callback?code=...&state=...
  └─► Server exchanges code and stores tokens in that session's TokenManager only
  └─► Green success page shown — browser can be closed

Step 3 — Tool calls
  └─► acquireTokenSilent() — uses the session's in-memory refresh token
  └─► On 401 from Graph: single retry with fresh token, then error
```

#### How to start the login flow

A fresh `loginUrl` is embedded in every 401 response to an unauthenticated
tool call, and is also returned by the `get_login_url` MCP tool:

```json
{
  "error": "Unauthorized",
  "loginUrl": "https://msgraph.example.com/auth/login?token=3941f8c7dca11388..."
}
```

Open the `loginUrl` directly in a browser to authenticate. The token is
single-use and expires after 15 minutes — each tool call / `get_login_url`
invocation mints a new one bound to the current MCP session. The session
ID itself is never exposed in URLs.

#### Health check

The `/health` endpoint intentionally returns only aggregate counts — no
session IDs, no UPNs — so that anyone who can reach the endpoint cannot
enumerate active users or sessions.

```
GET /health
→ {
    "status": "ok",
    "service": "msgraphmcp",
    "authMode": "authorization-code",
    "sessions": 2,
    "authenticatedSessions": 1
  }
```

#### Token lifetime

Tokens are kept **in-memory** for the session's lifetime. They survive within a running pod but are lost on pod restart — each session must re-authenticate after a restart. The `loginUrl` in the 401 response makes re-authentication a single browser visit.

#### Kubernetes deployment

Uncomment the **Option A** block in `k8s/deployment.yaml` and set `AZURE_REDIRECT_URI` to your ingress hostname:

```yaml
- name: AZURE_CLIENT_SECRET
  valueFrom:
    secretKeyRef:
      name: msgraphmcp-azure
      key: AZURE_CLIENT_SECRET
- name: AZURE_REDIRECT_URI
  value: https://msgraph.example.com/auth/callback
```

### Mode B — Client Secret (app-only)

Set `AZURE_CLIENT_SECRET` without `AZURE_REDIRECT_URI`. Requires **Application** permissions + admin consent. Recommended for fully automated deployments where no user context is needed.

```
Every request
  └─► acquireTokenByClientCredential({ scopes: ['.default'] })
  └─► Entra ID returns a short-lived access token (no refresh token stored)
  └─► Device compliance CA policies are NOT evaluated — safe in containers
```

### Mode C — Client Certificate (app-only, recommended for production)

Set `AZURE_CLIENT_CERTIFICATE_PATH` + `AZURE_CLIENT_CERTIFICATE_THUMBPRINT`. Same flow as Mode B but uses a certificate assertion instead of a shared secret — no secret rotation required.

```
Every request
  └─► acquireTokenByClientCredential with cert assertion
  └─► Entra ID returns a short-lived access token
  └─► Device compliance CA policies are NOT evaluated
```

**Kubernetes setup:**

```bash
# 1. Generate a self-signed key + cert (or use your PKI)
openssl req -x509 -newkey rsa:2048 -keyout tls.key -out tls.crt -days 365 -nodes -subj "/CN=msgraphmcp"

# 2. Get the SHA-256 thumbprint (64 hex chars)
openssl x509 -in tls.crt -fingerprint -sha256 -noout | tr -d ':' | sed 's/.*=//'

# 3. Upload tls.crt to the App Registration → Certificates & secrets → Certificates

# 4. Store the private key as a Kubernetes Secret
kubectl create secret generic msgraphmcp-client-cert --from-file=tls.key -n msgraphmcp
```

Then uncomment the `client-cert` volume in `k8s/deployment.yaml` and set `AZURE_CLIENT_CERTIFICATE_THUMBPRINT` in `k8s/secret.yaml`.

### Mode D — Device Code (delegated, local / interactive)

No secret or certificate configured. Suitable for local use and Claude Code stdio integration. **Not recommended in containers under Entra ID Conditional Access device-compliance policies** — token refresh from a non-enrolled host will be rejected.

```
First run
  └─► Device code prompt (stderr) → user visits URL, enters code
  └─► MSAL receives tokens → writes cache to TOKEN_CACHE_PATH

Subsequent runs / token expiry
  └─► acquireTokenSilent() uses cached refresh token
  └─► CA compliance evaluated against the container → may fail

401 from Graph API
  └─► Interceptor retries once with fresh token, then throws error
```

The refresh token typically lasts **90 days**. If it expires, the device code prompt appears again on next start.

---

## Tool Reference

All tools accept `userId: string` parameters that default to `"me"` (the signed-in user) unless noted.  
Parameters marked *optional* can be omitted.

### Users

| Tool | Description | Key Parameters |
|---|---|---|
| `list_users` | List directory users | `filter`, `select`, `top`, `search` |
| `get_user` | Get a single user | `userId` (**required**), `select` |
| `create_user` | Create a new user | `displayName`, `userPrincipalName`, `mailNickname`, `password` |
| `update_user` | Update user properties | `userId`, `displayName`, `jobTitle`, `department`, … |
| `delete_user` | Delete a user | `userId` |
| `get_user_member_of` | Get groups/roles the user belongs to | `userId` |
| `reset_user_password` | Reset a user's password | `userId`, `newPassword`, `forceChangePasswordNextSignIn` |

**Example — create user:**
```
create_user displayName="Alice Müller" userPrincipalName="alice@contoso.com" mailNickname="alice" password="P@ssw0rd!"
```

---

### Mail

| Tool | Description | Key Parameters |
|---|---|---|
| `list_messages` | List messages in a folder | `userId`, `folderId` (default `inbox`), `filter`, `top`, `search` |
| `get_message` | Get a specific message | `userId`, `messageId` |
| `send_mail` | Send an email | `subject`, `body`, `toRecipients[]`, `ccRecipients[]`, `bccRecipients[]` |
| `reply_to_message` | Reply to a message | `userId`, `messageId`, `comment` |
| `forward_message` | Forward a message | `userId`, `messageId`, `toRecipients[]`, `comment` |
| `delete_message` | Delete a message | `userId`, `messageId` |
| `move_message` | Move to another folder | `userId`, `messageId`, `destinationFolderId` |
| `list_mail_folders` | List mail folders | `userId`, `includeHiddenFolders` |
| `create_mail_folder` | Create a new folder | `userId`, `displayName`, `parentFolderId` |

**Well-known folder IDs:** `inbox`, `sentitems`, `drafts`, `deleteditems`, `archive`, `junkemail`

**Example — send mail:**
```
send_mail subject="Meeting tomorrow" body="Hi Bob,\nSee you at 10:00." toRecipients=[{address:"bob@contoso.com"}]
```

---

### Calendar

| Tool | Description | Key Parameters |
|---|---|---|
| `list_calendars` | List all calendars | `userId` |
| `create_calendar` | Create a calendar | `userId`, `name`, `color` |
| `list_events` | List events (or calendar view) | `userId`, `calendarId`, `startDateTime`, `endDateTime`, `filter`, `top` |
| `get_event` | Get a specific event | `userId`, `eventId` |
| `create_event` | Create an event | `subject`, `startDateTime`, `endDateTime`, `attendees[]`, `location`, `isOnlineMeeting` |
| `update_event` | Update an event | `userId`, `eventId`, + any field |
| `delete_event` | Delete an event | `userId`, `eventId` |

**Example — create Teams meeting:**
```
create_event subject="Sprint Review" startDateTime="2024-06-14T14:00:00" endDateTime="2024-06-14T15:00:00" isOnlineMeeting=true attendees=[{address:"bob@contoso.com",type:"required"}]
```

---

### OneDrive / Files

| Tool | Description | Key Parameters |
|---|---|---|
| `list_drive_items` | List items in a folder | `userId`, `itemPath` (default `/`), `top` |
| `get_drive_item` | Get metadata | `userId`, `itemPath` or `itemId` |
| `create_drive_folder` | Create a folder | `userId`, `parentPath`, `folderName`, `conflictBehavior` |
| `upload_drive_file` | Upload text file (≤ 4 MB) | `userId`, `filePath`, `content`, `conflictBehavior` |
| `delete_drive_item` | Delete an item | `userId`, `itemPath` or `itemId` |
| `copy_drive_item` | Copy an item | `userId`, `itemId`, `destinationParentId`, `newName` |
| `search_drive` | Search OneDrive | `userId`, `query`, `top` |
| `list_shared_with_me` | List shared items | `userId` |

---

### Groups

| Tool | Description | Key Parameters |
|---|---|---|
| `list_groups` | List groups | `filter`, `select`, `top`, `search` |
| `get_group` | Get a group | `groupId`, `select` |
| `create_group` | Create M365 or Security group | `displayName`, `mailNickname`, `groupType` (`Microsoft365`\|`Security`) |
| `update_group` | Update group | `groupId`, `displayName`, `description`, `visibility` |
| `delete_group` | Delete a group | `groupId` |
| `list_group_members` | List members | `groupId`, `select` |
| `add_group_member` | Add a member | `groupId`, `memberId` |
| `remove_group_member` | Remove a member | `groupId`, `memberId` |
| `list_group_owners` | List owners | `groupId` |
| `add_group_owner` | Add an owner | `groupId`, `ownerId` |

---

### Microsoft Teams

| Tool | Description | Key Parameters |
|---|---|---|
| `list_joined_teams` | List teams the user belongs to | `userId` |
| `get_team` | Get team details | `teamId` |
| `create_team` | Create a new team | `displayName`, `description`, `visibility`, `template` |
| `list_channels` | List channels | `teamId` |
| `get_channel` | Get a channel | `teamId`, `channelId` |
| `create_channel` | Create a channel | `teamId`, `displayName`, `membershipType` |
| `delete_channel` | Delete a channel | `teamId`, `channelId` |
| `list_channel_messages` | List messages | `teamId`, `channelId`, `top` |
| `send_channel_message` | Post a message | `teamId`, `channelId`, `content`, `contentType` |
| `reply_to_channel_message` | Reply to a message | `teamId`, `channelId`, `messageId`, `content` |
| `list_team_members` | List members | `teamId` |
| `add_team_member` | Add a member | `teamId`, `userId`, `roles` |

---

### Contacts

| Tool | Description | Key Parameters |
|---|---|---|
| `list_contacts` | List contacts | `userId`, `filter`, `select`, `top` |
| `get_contact` | Get a contact | `userId`, `contactId` |
| `create_contact` | Create a contact | `userId`, `givenName`, `surname`, `emailAddresses[]`, `businessPhones[]` |
| `update_contact` | Update a contact | `userId`, `contactId`, + any field |
| `delete_contact` | Delete a contact | `userId`, `contactId` |

---

### To Do / Tasks

| Tool | Description | Key Parameters |
|---|---|---|
| `list_todo_lists` | List task lists | `userId` |
| `create_todo_list` | Create a task list | `userId`, `displayName` |
| `delete_todo_list` | Delete a task list | `userId`, `listId` |
| `list_tasks` | List tasks in a list | `userId`, `listId`, `filter`, `top` |
| `create_task` | Create a task | `userId`, `listId`, `title`, `dueDateTime`, `importance`, `reminderDateTime` |
| `update_task` | Update a task | `userId`, `listId`, `taskId`, `status`, `title`, `importance` |
| `complete_task` | Mark as completed | `userId`, `listId`, `taskId` |
| `delete_task` | Delete a task | `userId`, `listId`, `taskId` |

**Task status values:** `notStarted`, `inProgress`, `completed`, `waitingOnOthers`, `deferred`

---

### SharePoint Sites

| Tool | Description | Key Parameters |
|---|---|---|
| `list_sites` | List sites | `filter`, `top` |
| `get_site` | Get a site | `siteId` or (`hostname` + `sitePath`) |
| `search_sites` | Search by keyword | `query` |
| `list_site_lists` | List site lists/libraries | `siteId` |
| `get_site_list` | Get a list | `siteId`, `listId` |
| `list_site_list_items` | List list items | `siteId`, `listId`, `filter`, `top`, `expand` |
| `get_site_list_item` | Get an item (with fields) | `siteId`, `listId`, `itemId` |
| `create_site_list_item` | Create an item | `siteId`, `listId`, `fields` (key-value object) |
| `update_site_list_item` | Update item fields | `siteId`, `listId`, `itemId`, `fields` |
| `delete_site_list_item` | Delete an item | `siteId`, `listId`, `itemId` |

---

### Intune — Apps

| Tool | Description | Key Parameters |
|---|---|---|
| `list_intune_apps` | List managed apps | `filter`, `appType`, `select`, `top` |
| `get_intune_app` | Get app details | `appId`, `select` |
| `create_intune_web_app` | Add a web shortcut app | `displayName`, `publisher`, `appUrl` |
| `create_intune_store_app` | Add a store app | `displayName`, `publisher`, `storeType` (`windowsStore`\|`iosStore`\|`androidStore`), `appStoreUrl` |
| `update_intune_app` | Update app metadata (incl. Win32 detection `rules`) | `appId`, `displayName`, `description`, `isFeatured`, `rules[]`, … |
| `delete_intune_app` | Delete an app | `appId` |
| `upload_win32_lob_app` | Upload a `.intunewin` Win32 LOB package. Source is **exactly one** of: (a) `filePath`, (b) `fileUrl`, or (c) a OneDrive reference (`oneDriveItemPath` **or** `oneDriveItemId`, optionally with `oneDriveUserId` — defaults to `"me"`). | `filePath` \| `fileUrl` \| (`oneDriveItemPath` \| `oneDriveItemId` [+ `oneDriveUserId`]), `displayName`, `publisher`, `description`, `installCommandLine`, `uninstallCommandLine`, `setupFilePath`, `applicableArchitectures`, `minimumSupportedWindowsRelease`, `runAsAccount`, `deviceRestartBehavior` |
| `list_intune_app_relationships` | List supersedence / dependency links | `appId` |
| `set_intune_app_relationships` | Set supersedence / dependency links (replaces all) | `appId`, `relationships[]` |
| `list_intune_app_assignments` | List assignments | `appId` |
| `assign_intune_app` | Assign to groups (replaces all existing) | `appId`, `assignments[]` (`groupId`, `intent`, `filterMode`, `filterId`) |
| `get_intune_app_install_status` | Per-device install status (Reports API) | `appId`, `top`, `includeUserStatuses` |
| `list_intune_app_protection_policies` | List MAM / App Protection policies | `platform` (`ios`\|`android`\|`all`) |

**Assignment intent values:** `available`, `required`, `uninstall`, `availableWithoutEnrollment`

**Win32 LOB upload notes:**

- Provide **exactly one** source: `filePath` (server-side absolute path), `fileUrl` (HTTP/HTTPS URL) or a OneDrive reference (`oneDriveItemPath` or `oneDriveItemId`, optionally with `oneDriveUserId` — defaults to `"me"`).
- `fileUrl` is fetched server-side with an SSRF guard (http/https only; loopback, link-local, RFC1918 ranges, and cloud-metadata hostnames are blocked).
- OneDrive sources resolve the DriveItem via Graph (`$select=id,name,size,file,folder,@microsoft.graph.downloadUrl`), reject folders, and stream the short-lived pre-authenticated download URL through the same bounded writer.
- Downloads are capped at 2 GB (applies to `fileUrl` and OneDrive sources alike).
- After upload, set detection/requirement rules via `update_intune_app` (field `rules[]`) and assign to groups via `assign_intune_app`.

**Example — upload from OneDrive:**
```
upload_win32_lob_app
  oneDriveItemPath="/Apps/myapp.intunewin"
  displayName="My App"
  publisher="Contoso"
  installCommandLine="setup.exe /S"
  uninstallCommandLine="setup.exe /U"
  setupFilePath="setup.exe"
```

**Example — assign app as required:**
```
assign_intune_app appId="00000000-0000-0000-0000-000000000001" assignments=[{groupId:"grp-id",intent:"required"}]
```

---

### Intune — Device Configurations

| Tool | Description | Key Parameters |
|---|---|---|
| `list_device_configurations` | List config profiles | `filter`, `select`, `top` |
| `get_device_configuration` | Get a profile | `configId` |
| `create_device_configuration` | Create a profile | `displayName`, `odataType`, `settings` (key-value) |
| `update_device_configuration` | Update a profile | `configId`, `displayName`, `settings` |
| `delete_device_configuration` | Delete a profile | `configId` |
| `assign_device_configuration` | Assign to groups | `configId`, `assignments[]` (`groupId`, `intent: include\|exclude`) |
| `get_device_configuration_assignments` | List assignments | `configId` |
| `get_device_configuration_device_status` | Per-device status | `configId`, `top` |

**Common `odataType` values:**

| Platform | `@odata.type` |
|---|---|
| Windows 10 (general) | `#microsoft.graph.windows10GeneralConfiguration` |
| Windows 10 (endpoint protection) | `#microsoft.graph.windows10EndpointProtectionConfiguration` |
| iOS | `#microsoft.graph.iosGeneralDeviceConfiguration` |
| Android | `#microsoft.graph.androidGeneralDeviceConfiguration` |
| macOS | `#microsoft.graph.macOSGeneralDeviceConfiguration` |

**Example — create Windows 10 password policy:**
```
create_device_configuration
  displayName="Windows Password Policy"
  odataType="#microsoft.graph.windows10GeneralConfiguration"
  settings={"passwordRequired":true,"passwordMinimumLength":12,"passwordRequiredType":"alphanumeric"}
```

---

### Intune — Settings Catalog

The Settings Catalog is the modern replacement for device configuration profiles and supports a broader set of settings.

| Tool | Description | Key Parameters |
|---|---|---|
| `list_configuration_policies` | List catalog policies | `filter`, `top` |
| `get_configuration_policy` | Get policy + settings | `policyId` |
| `create_configuration_policy` | Create a policy | `name`, `platforms`, `technologies`, `settings[]` |
| `update_configuration_policy` | Update name/description | `policyId`, `name`, `description` |
| `delete_configuration_policy` | Delete a policy | `policyId` |
| `assign_configuration_policy` | Assign to groups | `policyId`, `assignments[]` |
| `search_settings_catalog` | Find available settings | `keyword`, `platform`, `top` |

**Workflow — create a BitLocker policy:**
```
1. search_settings_catalog keyword="BitLocker" platform="windows10"
   → note the settingDefinitionId values

2. create_configuration_policy
     name="BitLocker Encryption"
     platforms="windows10"
     technologies="mdm"
     settings=[{
       id:"0",
       settingInstance:{
         "@odata.type":"#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance",
         "settingDefinitionId":"device_vendor_msft_bitlocker_requiredeviceencryption",
         "choiceSettingValue":{"value":"device_vendor_msft_bitlocker_requiredeviceencryption_1","children":[]}
       }
     }]
```

---

### Intune — Compliance Policies

| Tool | Description | Key Parameters |
|---|---|---|
| `list_compliance_policies` | List compliance policies | `filter`, `top` |
| `get_compliance_policy` | Get a policy | `policyId` |
| `create_compliance_policy` | Create a policy | `displayName`, `odataType`, `settings`, `scheduledActionsForRule` |
| `update_compliance_policy` | Update a policy | `policyId`, `displayName`, `settings` |
| `delete_compliance_policy` | Delete a policy | `policyId` |
| `assign_compliance_policy` | Assign to groups | `policyId`, `assignments[]` |
| `get_compliance_policy_device_status` | Per-device status | `policyId`, `top` |

**Common `odataType` values:**

| Platform | `@odata.type` |
|---|---|
| Windows 10 | `#microsoft.graph.windows10CompliancePolicy` |
| iOS | `#microsoft.graph.iosCompliancePolicy` |
| Android | `#microsoft.graph.androidCompliancePolicy` |
| macOS | `#microsoft.graph.macOSCompliancePolicy` |

**Non-compliance actions:** `block`, `retire`, `wipe`, `notification`, `pushNotification`

---

### Intune — Managed Devices

| Tool | Description | Key Parameters |
|---|---|---|
| `list_managed_devices` | List enrolled devices | `filter`, `select`, `top` |
| `get_managed_device` | Get device details | `deviceId` |
| `sync_managed_device` | Trigger policy sync | `deviceId` |
| `restart_managed_device` | Reboot a device | `deviceId` |
| `shutdown_managed_device` | Shut down a Windows device | `deviceId` |
| `lock_managed_device` | Remote lock (PIN/password required to unlock) | `deviceId` |
| `set_managed_device_name` | Rename a Windows device (≤ 15 chars) | `deviceId`, `deviceName` |
| `retire_managed_device` | Unenroll (keep personal data) | `deviceId` |
| `wipe_managed_device` | Factory reset (**destructive**) | `deviceId`, `keepEnrollmentData`, `keepUserData` |
| `disable_managed_device` | Disable a device — keeps enrollment, blocks access | `deviceId` |
| `reenable_managed_device` | Re-enable a previously disabled device | `deviceId` |
| `send_device_notification` | Push a Company Portal notification | `deviceId`, `notificationTitle`, `notificationBody` |
| `windows_defender_scan` | Trigger a Defender antivirus scan | `deviceId`, `quickScan` |
| `windows_defender_update_signatures` | Force Defender signature update | `deviceId` |
| `rotate_bitlocker_keys` | Rotate BitLocker recovery key (escrowed to Entra) | `deviceId` |
| `rotate_local_admin_password` | Rotate Windows LAPS local admin password | `deviceId` |
| `trigger_proactive_remediation` | Run a Proactive Remediation script on demand | `deviceId`, `scriptId` |
| `get_device_compliance_overview` | Tenant-wide compliance stats | — |

> **Warning:** `wipe_managed_device` erases all data on the device. Use with extreme caution.

**Useful filter examples for `list_managed_devices`:**
```
filter="operatingSystem eq 'Windows'"
filter="complianceState eq 'noncompliant'"
filter="contains(deviceName,'DESKTOP')"
```

---

### Intune — Device Diagnostics

| Tool | Description | Key Parameters |
|---|---|---|
| `collect_device_diagnostics` | Start a "Collect diagnostics" remote action | `deviceId` |
| `list_device_diagnostics` | List / get status of diagnostic collection requests | `deviceId`, `requestId` (optional) |
| `download_device_diagnostics` | Generate a SAS download URL for a completed ZIP | `deviceId`, `requestId` |

Workflow:
1. `collect_device_diagnostics` → returns a `requestId`.
2. Poll `list_device_diagnostics` until `status: "completed"`.
3. `download_device_diagnostics` → returns `{ value: "<sas-url>" }` pointing at the ZIP archive.

---

### Intune — Notification Templates

| Tool | Description | Key Parameters |
|---|---|---|
| `list_notification_templates` | List templates | `top` |
| `get_notification_template` | Get a template + localized messages | `templateId` |
| `create_notification_template` | Create a template (via beta) | `displayName`, `defaultLocale`, `brandingOptions`, `roleScopeTagIds` |
| `update_notification_template` | Update template metadata | `templateId`, `displayName`, `brandingOptions`, … |
| `delete_notification_template` | Delete a template | `templateId` |
| `add_notification_template_message` | Add / update a localized message | `templateId`, `locale`, `subject`, `messageTemplate`, `isDefault` |
| `send_notification_template_test` | Send a test email using the default locale | `templateId` |

> `create_notification_template` is routed through the `beta` API because the v1.0 endpoint returns 400 for all write operations on this resource.

---

### Auth

| Tool | Description | Key Parameters |
|---|---|---|
| `get_login_url` | Returns authentication status and a fresh one-time login URL when sign-in is required. Call this first whenever another tool returns an auth error. | — |

---

## Development

```bash
# Install dependencies
npm install

# Start in watch mode (ts-node, no compile step)
npm run dev

# Build production bundle (esbuild, ~10 ms)
npm run build

# Run unit tests
npm test

# Run tests with coverage report
npm run test:coverage

# TypeScript type check only (no emit)
npm run typecheck
```

### Project Structure

```
msgraphmcp/
├── src/
│   ├── index.ts                  Entry point — creates MCP server, connects stdio transport
│   ├── auth/
│   │   └── TokenManager.ts       MSAL public client, device code flow, silent refresh
│   ├── graph/
│   │   └── GraphClient.ts        Axios wrapper: auth header injection, pagination, error handling
│   └── tools/
│       ├── shared.ts             URL / OData / HTML escaping helpers
│       ├── auth.ts               1 tool  (get_login_url)
│       ├── users.ts              7 tools
│       ├── mail.ts               9 tools
│       ├── calendar.ts           7 tools
│       ├── files.ts              8 tools
│       ├── groups.ts             10 tools
│       ├── teams.ts              12 tools
│       ├── contacts.ts           5 tools
│       ├── tasks.ts              8 tools
│       ├── sites.ts              10 tools
│       └── intune.ts             63 tools (apps · win32 LOB · config · settings catalog · compliance · devices · diagnostics · notification templates)
├── tests/
│   ├── helpers/
│   │   ├── MockMcpServer.ts      Captures tool registrations for unit testing
│   │   └── mockGraphClient.ts    Jest mock for GraphClient
│   ├── auth/
│   │   └── TokenManager.test.ts
│   ├── graph/
│   │   └── GraphClient.test.ts
│   └── tools/
│       ├── users.test.ts … intune.test.ts
├── .github/
│   └── workflows/
│       ├── ci.yml                Test + build on push/PR
│       └── docker.yml            Build multi-arch image → ghcr.io
├── Dockerfile
├── docker-compose.yml
└── .env.example
```

---

## CI/CD and Docker Registry

### GitHub Actions

| Workflow | Trigger | Jobs |
|---|---|---|
| `ci.yml` | push to `main`/`develop`, PR to `main` | typecheck → test (with coverage) → build → smoke test |
| `docker.yml` | push to `main`, version tags `v*.*.*`, manual | test (gate) → build multi-arch image → push to GHCR → provenance attestation |

### Docker image tags

| Tag | When created |
|---|---|
| `latest` | Every push to `main` |
| `1.2.3` | When a `v1.2.3` tag is pushed |
| `1.2` | When a `v1.2.x` tag is pushed |
| `sha-abc1234` | Every push (traceability) |

### Pulling from GHCR

```bash
docker pull ghcr.io/DustHoff/msgraphmcp:latest
```

### Setup in your repository

1. **Replace `DustHoff`** in the badge URLs above with your GitHub username or organisation.
2. Push to a GitHub repository — Actions will run automatically.
3. For private packages: go to **Settings → Packages → msgraphmcp → Package settings → Change visibility** if needed.

---

## Security Notes

- **Per-session token isolation (Mode A):** each MCP session in HTTP mode holds its own in-memory token. Tokens are never shared between sessions — a compromised session cannot access another user's data. Tokens are lost on pod restart; re-authentication is a single browser visit via the `loginUrl` in the 401 response.
- **One-time login tokens:** the `loginUrl` carries a single-use 256-bit token (not the MCP session ID). It is consumed on first visit, expires after 15 minutes, and binds server-side to exactly one session — so even if a login URL is disclosed via browser history, a proxy log, or an email it cannot be replayed.
- **HTML escaping on the auth pages:** the `/auth/callback` error page escapes all five HTML-significant characters (`&`, `<`, `>`, `"`, `'`) — an attacker cannot reflect a crafted `error_description` (including HTML numeric entities like `&#60;script&#62;`) into executable markup.
- **URL-path safety:** all user-supplied opaque ids passed to Graph (`groupId`, `teamId`, `channelId`, `appId`, `configId`, `policyId`, `templateId`, `deviceId`, `siteId`, `listId`, `itemId`, `memberId`, `ownerId`, `userId`) are percent-encoded before they are embedded in Graph URLs, so a tool argument cannot smuggle extra path segments, query strings, or fragments into a request. SharePoint composite site ids (`hostname,guid,guid`) keep their commas intact.
- **SSRF guard on Win32 LOB upload:** `upload_win32_lob_app` with a `fileUrl` argument rejects non-http(s) schemes and blocks loopback, link-local, RFC1918, cloud-metadata and multicast ranges before issuing the download. OneDrive sources reuse the same guard against the short-lived pre-authenticated download URL returned by Graph. The download size is capped at 2 GB in all cases.
- **Request-body limit:** the HTTP server caps incoming request bodies at 4 MB and destroys the socket immediately on overflow to prevent OOM via a large JSON-RPC payload.
- **Session limits:** `MAX_SESSIONS` (default 50) caps concurrent sessions; `SESSION_IDLE_TIMEOUT_MINUTES` (default 60) reaps idle sessions so abandoned connections do not leak memory.
- **Sensitive-value redaction in debug logs:** values of keys matching `/password|secret|token|credential|private[-_]?key|apikey/i` are replaced with `***REDACTED***` before any request/response body is logged.
- **Token cache file** (`tokens.json`) is only written in device-code (Mode D) and stdio mode. Written with `mode 0o600`; mount as a restricted Docker volume; do not bake into images.
- The image runs as the **non-root `node` user** (see `Dockerfile`).
- **Secrets** (`AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, certificate thumbprint) — use Kubernetes Secrets or an external vault; never commit real values to the repository. `secrets.txt`, `tokens.json`, and `data/` are excluded via both `.gitignore` and `.dockerignore`.
- **Prefer authorization code (Mode A) over device code (Mode D)** for HTTP deployments — tokens are tied to the user's browser device so CA compliance is evaluated correctly. Device code refresh from a non-enrolled container is rejected by device-compliance CA policies.
- **Prefer client certificate (Mode C) over client secret (Mode B)** for app-only deployments — certificates are not transmitted over the wire and can be rotated without application downtime.
- **Authorization code flow:** `/auth/login?token=<one-time-token>` and `/auth/callback` must be reachable by the authenticating browser but do not need to be internet-facing — internal DNS is sufficient. Requests without a valid, unexpired token are rejected with 400.
- **Graph API logging** includes the authenticated user's UPN (`"user": "alice@contoso.com"`) in every log entry, making all MS Graph calls attributable to a specific user identity. Set `LOG_LEVEL=warn` in production to suppress URL logging if UPNs in logs are a concern.
- Scope down `GRAPH_SCOPES` (delegated modes) or grant only the required permissions (app-only modes) for your use case.
- `wipe_managed_device` is irreversible — consider requiring explicit confirmation in your workflows.
- See [`SECURITY-NOTICE.md`](SECURITY-NOTICE.md) for the full security assessment including dependency risk analysis.
