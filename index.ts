/**
 * NVIDIA Enterprise plugin.
 *
 * Provides enterprise tools and Azure AD login for Microsoft Graph tools.
 * Run `pnpm nvidia:setup` to configure, then visit http://localhost:3000
 * to authenticate via Azure AD.
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { definePluginEntry, type AnyAgentTool } from "openclaw/plugin-sdk/core";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const FAVICON_PATH = path.join(__dirname, "..", "assets", "favicon.svg");
import { createEmployeeInfoTool } from "./src/employee-info.js";
import { createGleanSearchTool } from "./src/glean-search.js";
import { createMeetingRoomTool } from "./src/meeting-room.js";
import { createNfdDeskTool } from "./src/nfd-desk.js";
import {
  getSSOToken,
  getAzureRefreshToken,
  handleLogin,
  handleCallback,
  handleTokenLogin,
  handleStatus,
  handleLogout,
  handleAuthGate,
} from "./src/oidc.js";
import { createOutlookAvailabilityTool } from "./src/outlook-availability.js";
import { createOutlookCalendarTool } from "./src/outlook-calendar.js";
import { createOutlookEmailTool } from "./src/outlook-email.js";
import { createPeopleSearchTool } from "./src/people-search.js";
import { createTeamsChatTool } from "./src/teams-chat.js";

// =============================================================================
// Plugin Entry
// =============================================================================

export default definePluginEntry({
  id: "nvidia-enterprise",
  name: "NVIDIA Enterprise Plugin",
  description:
    "NVIDIA enterprise tools + Azure AD SSO for Outlook, Calendar, Teams, People, NFD, Meeting rooms, Glean, Employee info",
  register(api) {
    // -------------------------------------------------------------------------
    // 1. Custom favicon (overrides stock OpenClaw favicon)
    // -------------------------------------------------------------------------
    let faviconCache: Buffer | null = null;
    api.registerHttpRoute({
      path: "/favicon.svg",
      auth: "plugin",
      handler: (req, res) => {
        try {
          if (!faviconCache) {
            faviconCache = fs.readFileSync(FAVICON_PATH);
          }
          res.writeHead(200, { "Content-Type": "image/svg+xml", "Cache-Control": "public, max-age=3600" });
          res.end(faviconCache);
          return true;
        } catch {
          return false; // fall through to stock favicon
        }
      },
    });

    // -------------------------------------------------------------------------
    // 2. Auth gate: redirect unauthenticated browser requests to Azure AD
    // -------------------------------------------------------------------------
    api.registerHttpRoute({
      path: "/",
      auth: "plugin",
      match: "prefix",
      handler: (req, res) => handleAuthGate(req, res),
    });

    // -------------------------------------------------------------------------
    // 2. Azure AD OAuth routes
    // -------------------------------------------------------------------------
    api.registerHttpRoute({
      path: "/azure-ad/login",
      auth: "plugin",
      handler: (req, res) => {
        handleLogin(req, res);
      },
    });
    api.registerHttpRoute({
      path: "/api/auth/callback/nvlogin",
      auth: "plugin",
      handler: async (req, res) => {
        await handleCallback(req, res);
      },
    });
    api.registerHttpRoute({
      path: "/api/auth/token-login",
      auth: "plugin",
      handler: async (req, res) => {
        await handleTokenLogin(req, res);
      },
    });
    api.registerHttpRoute({
      path: "/azure-ad/status",
      auth: "plugin",
      handler: (req, res) => {
        handleStatus(req, res);
      },
    });
    api.registerHttpRoute({
      path: "/azure-ad/logout",
      auth: "plugin",
      handler: (req, res) => {
        handleLogout(req, res);
      },
    });

    // -------------------------------------------------------------------------
    // 3. Enterprise tools
    // -------------------------------------------------------------------------
    const tools: Array<AnyAgentTool | null> = [
      createGleanSearchTool({ getSSOToken }),
      createPeopleSearchTool({ getRefreshToken: getAzureRefreshToken }),
      createOutlookEmailTool({ getRefreshToken: getAzureRefreshToken }),
      createOutlookCalendarTool({ getRefreshToken: getAzureRefreshToken }),
      createOutlookAvailabilityTool({ getRefreshToken: getAzureRefreshToken }),
      createTeamsChatTool({ getRefreshToken: getAzureRefreshToken }),
      createNfdDeskTool({ getRefreshToken: getAzureRefreshToken }),
      createMeetingRoomTool({ getRefreshToken: getAzureRefreshToken }),
      createEmployeeInfoTool(),
    ];

    for (const tool of tools) {
      if (tool) {
        api.registerTool(tool);
      }
    }
  },
});
