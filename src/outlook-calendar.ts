/**
 * Outlook Calendar tool (read-only).
 *
 * Fetches calendar events via Microsoft Graph API for a time window.
 * Uses Azure AD refresh token -> silent acquisition -> OBO for Graph Calendars.Read scope.
 */

import { Type } from "@sinclair/typebox";
import { jsonResult, readNumberParam, readStringParam } from "openclaw/plugin-sdk/agent-runtime";
import { stringEnum, type AnyAgentTool } from "openclaw/plugin-sdk/core";
import { acquireDownstreamToken, isAzureOBOConfigured } from "./azure-obo.js";

// =============================================================================
// Configuration
// =============================================================================

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const CALENDAR_SCOPES = "https://graph.microsoft.com/Calendars.Read";

const DEFAULT_TIMEOUT_MS = 30_000;
const DEFAULT_EVENT_COUNT = 20;
const MAX_EVENT_COUNT = 50;

// =============================================================================
// Actions
// =============================================================================

const CALENDAR_ACTIONS = ["list_events"] as const;

// =============================================================================
// Schema
// =============================================================================

const OutlookCalendarSchema = Type.Object({
  action: stringEnum(CALENDAR_ACTIONS, {
    description: "Action to perform: list_events (get calendar events for a date range).",
  }),
  start_time: Type.Optional(
    Type.String({
      description:
        "Start of time window (ISO 8601). Defaults to now. Example: 2026-03-23T00:00:00Z",
    }),
  ),
  end_time: Type.Optional(
    Type.String({
      description:
        "End of time window (ISO 8601). Defaults to 24h after start_time. Example: 2026-03-24T00:00:00Z",
    }),
  ),
  count: Type.Optional(
    Type.Number({
      description: `Max events to return (default: ${DEFAULT_EVENT_COUNT}, max: ${MAX_EVENT_COUNT}).`,
      minimum: 1,
      maximum: MAX_EVENT_COUNT,
    }),
  ),
});

// =============================================================================
// Graph API Helpers
// =============================================================================

async function graphGet(
  token: string,
  path: string,
  headers?: Record<string, string>,
): Promise<unknown> {
  const url = `${GRAPH_BASE_URL}${path}`;
  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...headers,
    },
    signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
  });

  if (!res.ok) {
    const errorData = (await res.json().catch(() => ({}))) as {
      error?: { message?: string; code?: string };
    };
    const message = errorData?.error?.message ?? `Graph API error: ${res.status}`;
    throw new Error(message);
  }

  return (await res.json()) as unknown;
}

function toISOClean(d: Date): string {
  return d.toISOString().replace(/\.\d{3}Z$/, "Z");
}

// =============================================================================
// Response Types
// =============================================================================

type GraphEvent = {
  subject?: string;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  organizer?: { emailAddress?: { name?: string; address?: string } };
  bodyPreview?: string;
  webLink?: string;
  isAllDay?: boolean;
  isCancelled?: boolean;
  location?: { displayName?: string };
  attendees?: Array<{
    emailAddress?: { name?: string; address?: string };
    status?: { response?: string };
  }>;
};

type GraphEventList = {
  value?: GraphEvent[];
};

type GraphMailboxSettings = {
  timeZone?: string;
};

// =============================================================================
// Action Handler
// =============================================================================

function formatEvent(ev: GraphEvent) {
  return {
    subject: ev.subject ?? "(No subject)",
    start: ev.start?.dateTime,
    end: ev.end?.dateTime,
    organizer: ev.organizer?.emailAddress?.name ?? "",
    organizerEmail: ev.organizer?.emailAddress?.address ?? "",
    location: ev.location?.displayName ?? "",
    preview: ev.bodyPreview ?? "",
    webUrl: ev.webLink ?? "",
    isAllDay: ev.isAllDay ?? false,
    isCancelled: ev.isCancelled ?? false,
    attendees: (ev.attendees ?? []).map((a) => ({
      name: a.emailAddress?.name ?? "",
      email: a.emailAddress?.address ?? "",
      response: a.status?.response ?? "",
    })),
  };
}

async function handleListEvents(
  graphToken: string,
  params: Record<string, unknown>,
): Promise<unknown> {
  const count = readNumberParam(params, "count", { integer: true }) ?? DEFAULT_EVENT_COUNT;

  const start = params.start_time ? new Date(String(params.start_time)) : new Date();
  const end = params.end_time
    ? new Date(String(params.end_time))
    : new Date(start.getTime() + 24 * 60 * 60 * 1000);

  // Get user timezone for proper calendar display
  let tz = "UTC";
  try {
    const settings = (await graphGet(graphToken, "/me/mailboxSettings")) as GraphMailboxSettings;
    tz = settings.timeZone || "UTC";
  } catch {
    // Fall back to UTC
  }

  const startStr = toISOClean(start);
  const endStr = toISOClean(end);
  const path = `/me/calendar/calendarView?startDateTime=${encodeURIComponent(startStr)}&endDateTime=${encodeURIComponent(endStr)}&$top=${Math.min(count, MAX_EVENT_COUNT)}&$orderby=start/dateTime`;

  const data = (await graphGet(graphToken, path, {
    Prefer: `outlook.timezone="${tz}"`,
  })) as GraphEventList;

  const events = (data.value ?? []).map(formatEvent);
  return { timezone: tz, count: events.length, events };
}

// =============================================================================
// Tool Factory
// =============================================================================

export function createOutlookCalendarTool(options: {
  getRefreshToken: () => string | null | Promise<string | null>;
  enabled?: boolean;
}): AnyAgentTool | null {
  if (options.enabled === false) {
    return null;
  }

  if (!isAzureOBOConfigured()) {
    return null;
  }

  return {
    label: "Outlook Calendar",
    name: "outlook_calendar",
    description:
      "Read Outlook calendar events via Microsoft Graph. Use to list upcoming meetings and events for a date range. Requires NVIDIA SSO authentication. This tool is read-only.",
    parameters: OutlookCalendarSchema,
    execute: async (_toolCallId, args) => {
      const params = args as Record<string, unknown>;
      const action = readStringParam(params, "action", { required: true });

      try {
        const refreshToken = await options.getRefreshToken();
        if (!refreshToken) {
          return jsonResult({
            error: "not_authenticated",
            message: "Outlook calendar tool requires NVIDIA SSO authentication. Please log in first.",
          });
        }

        const tokenResult = await acquireDownstreamToken(refreshToken, CALENDAR_SCOPES);
        if (!tokenResult.ok) {
          return jsonResult({
            error: "token_exchange_failed",
            message: `Failed to acquire Graph token: ${"error" in tokenResult ? tokenResult.error : "unknown"}`,
          });
        }
        const graphToken = tokenResult.accessToken;

        let result: unknown;
        switch (action) {
          case "list_events":
            result = await handleListEvents(graphToken, params);
            break;
          default:
            return jsonResult({
              error: "invalid_action",
              message: `Unknown action: ${action}. Valid: ${CALENDAR_ACTIONS.join(", ")}`,
            });
        }

        return jsonResult({ action, result });
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        return jsonResult({ error: "outlook_calendar_error", message });
      }
    },
  } as AnyAgentTool;
}
