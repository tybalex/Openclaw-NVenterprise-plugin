# NVIDIA Enterprise Plugin for OpenClaw

Standalone plugin that adds NVIDIA enterprise integrations to [OpenClaw](https://github.com/openclaw/openclaw) — no fork required.

## Features

- **NVIDIA Inference API** — Claude Sonnet 4.6 via `inference-api.nvidia.com`
- **Azure AD SSO** — browser-based login with auto token refresh, per-session cookies
- **Glean Search** — enterprise content search (Confluence, Slack, GDrive, Jira, SharePoint) via SSA service credentials
- **People Search** — Microsoft Graph people search and profiles
- **Outlook Email** — read-only inbox access (list, read, search)
- **NFD Desk Booking** — NVIDIA flex desk reservations
- **Meeting Rooms** — room search (meet.nvidia.com) + availability/booking (MS Graph)
- **Employee Info** — Helios directory API (manager chain, direct reports)

## Quick Start

```bash
# 1. Clone and build the plugin
git clone https://github.com/tybalex/Openclaw-NVenterprise-plugin.git
cd Openclaw-NVenterprise-plugin
pnpm install
pnpm build

# 2. Install into stock OpenClaw
cd ~/openclaw
pnpm openclaw plugins install ~/Openclaw-NVenterprise-plugin --link

# 3. Run setup (writes ~/.openclaw/openclaw.json)
cd ~/Openclaw-NVenterprise-plugin && pnpm nvidia:setup

# 4. Set env vars and run
export NVIDIA_API_KEY="..."
cd ~/openclaw && pnpm openclaw gateway run
```

### Setup Options

The setup script accepts flags to customize the config:

```bash
# Default: port 3000, auth none
pnpm nvidia:setup

# Custom port
pnpm nvidia:setup -- --port 8080

# Custom auth mode (none, token, password, trusted-proxy)
pnpm nvidia:setup -- --auth token

# Both
pnpm nvidia:setup -- --port 8080 --auth token
```

You can also run setup directly without cloning:

```bash
npx tsx https://raw.githubusercontent.com/tybalex/Openclaw-NVenterprise-plugin/main/scripts/nvidia-setup.ts
```

### Minimal Setup (model only, no enterprise tools)

If you just need the NVIDIA model provider without Azure AD or enterprise tools:

```bash
export NVIDIA_API_KEY="your-key"
cd ~/openclaw && pnpm openclaw gateway run
```

The plugin skips the Azure AD login gate and enterprise tools when their env vars aren't configured.

## Environment Variables

### Required — Model Provider
```
NVIDIA_API_KEY              # NVIDIA inference API key
```

### Azure AD (Outlook, People, NFD, Meeting Rooms)
```
AZURE_AD_CLIENT_ID          # Azure AD app client ID
AZURE_AD_CLIENT_SECRET      # Azure AD app client secret
AZURE_AD_TENANT_ID          # Azure AD tenant ID
AZURE_AD_SCOPES             # OAuth scopes (default: openid email profile offline_access api://<client-id>/User.Access)
```

### Glean Search (service credentials)
```
ECS_CONTENT_SEARCH_URL      # ECS search API endpoint
STARFLEET_CLIENT_ID         # Starfleet SSA client ID
STARFLEET_CLIENT_SECRET     # Starfleet SSA client secret
STARFLEET_TOKEN_URL         # Starfleet token endpoint
STARFLEET_SCOPE             # SSA scopes (default: content:search content:retrieve)
```

### NFD Desk Booking
```
NFD_SCOPE                   # Azure AD scope for NFD API
NFD_API_BASE_URL            # NFD API base URL (default: https://nfd-dev.nvidia.com)
```

### Meeting Rooms
```
GRAPH_SCOPES                # MS Graph scopes (default: Calendars.ReadWrite Group.Read.All)
MEET_API_URL                # Meet API URL (default: https://meet.nvidia.com/api/v1)
```

### Employee Info
```
HELIOS_API_KEY              # Helios directory API key
HELIOS_BASE_URL             # Helios base URL (default: https://helios-api.nvidia.com/api)
```

## How It Works

### Authentication

On first visit to `http://localhost:3000`, the plugin redirects to Azure AD login. After authentication:

- A session cookie is set (per-browser)
- The refresh token is stored in-memory for OBO token exchange
- Tokens auto-refresh 5 minutes before expiry
- Tools use the tokens automatically — no manual token management

Glean search uses separate SSA (service-to-service) credentials and works without browser login.

### Architecture

This is a **pure plugin** — zero changes to OpenClaw core:

- Model provider configured via `nvidia-setup.ts` (writes config)
- OIDC/SSO implemented via `registerHttpRoute` (plugin HTTP routes)
- Tools registered via `registerTool` (plugin tool API)
- Auth gate via prefix route on `/` (redirects unauthenticated browsers)

Pull from upstream OpenClaw anytime — no merge conflicts.

### Plugin HTTP Routes

| Route | Purpose |
|-------|---------|
| `GET /azure-ad/login` | Start Azure AD OAuth flow |
| `GET /api/auth/callback/nvlogin` | OAuth callback (registered in Azure AD) |
| `GET /azure-ad/status` | JSON login status |
| `GET /azure-ad/logout` | Clear session |

### Tool Availability

Each tool auto-enables when its env vars are set, and silently skips otherwise.

| Tool | Needs Browser Login | Needs Internal Network |
|------|--------------------|-----------------------|
| Glean Search | No (SSA only) | Yes (`enterprise-content-intelligence.nvidia.com`) |
| People Search | Yes (Azure AD) | No (`graph.microsoft.com`) |
| Outlook Email | Yes (Azure AD) | No (`graph.microsoft.com`) |
| NFD Desk | Yes (Azure AD) | Yes (`nfd-dev.nvidia.com`) |
| Meeting Rooms | Yes (Azure AD) | Partial (`meet.nvidia.com` + `graph.microsoft.com`) |
| Employee Info | No (API key) | Yes (`helios-api.nvidia.com`) |

## VM Startup Script

```bash
#!/bin/bash
sudo apt update && sudo apt upgrade -y
curl -fsSL https://deb.nodesource.com/setup_22.x | sudo bash -
sudo apt install -y nodejs
sudo npm install -g pnpm

# Stock OpenClaw
cd "$HOME" && git clone https://github.com/openclaw/openclaw
cd openclaw && pnpm install && pnpm ui:build && pnpm build

# NVIDIA Enterprise Plugin
cd "$HOME" && git clone https://github.com/tybalex/Openclaw-NVenterprise-plugin
cd Openclaw-NVenterprise-plugin && pnpm install && pnpm build

# Install plugin + setup config
cd "$HOME/openclaw" && pnpm openclaw plugins install "$HOME/Openclaw-NVenterprise-plugin" --link
node --import tsx "$HOME/Openclaw-NVenterprise-plugin/scripts/nvidia-setup.ts"
```

Then to run:
```bash
source ~/.env  # or however you inject secrets
cd ~/openclaw && pnpm openclaw gateway run
```

## Development

```bash
pnpm install
pnpm build        # compile TypeScript
pnpm dev          # watch mode
pnpm nvidia:setup        # write NVIDIA config to ~/.openclaw/openclaw.json
```
