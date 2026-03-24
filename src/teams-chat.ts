/**
 * Teams Chat tool (read-only).
 *
 * Fetches Microsoft Teams chats and messages via Graph API.
 * Uses Azure AD refresh token -> silent acquisition -> OBO for Graph Chat.Read scope.
 *
 * Key features:
 * - Batch API for efficient message fetching (20 chats per batch)
 * - Engagement window filtering (only chats where user participated recently)
 * - HTML stripping from message bodies
 */

import { Type } from "@sinclair/typebox";
import { jsonResult, readNumberParam, readStringParam } from "openclaw/plugin-sdk/agent-runtime";
import { stringEnum, type AnyAgentTool } from "openclaw/plugin-sdk/core";
import { acquireDownstreamToken, isAzureOBOConfigured } from "./azure-obo.js";

// =============================================================================
// Configuration
// =============================================================================

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const TEAMS_SCOPES = "https://graph.microsoft.com/Chat.Read";

const DEFAULT_TIMEOUT_MS = 60_000;
const MAX_CHATS = 50;
const BATCH_SIZE = 20;
const ENGAGEMENT_WINDOW_DAYS = 7;
const MAX_WORDS_PER_CHAT = 500;

// =============================================================================
// Actions
// =============================================================================

const TEAMS_ACTIONS = ["list_chats"] as const;

// =============================================================================
// Schema
// =============================================================================

const TeamsChatSchema = Type.Object({
  action: stringEnum(TEAMS_ACTIONS, {
    description: "Action to perform: list_chats (get recent Teams chats with messages).",
  }),
  start_time: Type.Optional(
    Type.String({
      description:
        "Only include messages after this time (ISO 8601). Defaults to 7 days ago. Example: 2026-03-16T00:00:00Z",
    }),
  ),
  limit: Type.Optional(
    Type.Number({
      description: `Max number of chats to fetch (default: ${MAX_CHATS}, max: ${MAX_CHATS}).`,
      minimum: 1,
      maximum: MAX_CHATS,
    }),
  ),
});

// =============================================================================
// Graph API Helpers
// =============================================================================

async function graphGet(token: string, path: string): Promise<unknown> {
  const url = `${GRAPH_BASE_URL}${path}`;
  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
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

async function graphBatchPost(
  token: string,
  requests: Array<{ id: string; method: string; url: string }>,
): Promise<Array<{ id: string; status: number; body?: unknown }>> {
  const res = await fetch(`${GRAPH_BASE_URL}/$batch`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ requests }),
    signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
  });

  if (!res.ok) {
    throw new Error(`Graph batch API error: ${res.status}`);
  }

  const data = (await res.json()) as { responses?: Array<{ id: string; status: number; body?: unknown }> };
  return data.responses ?? [];
}

function toISOClean(d: Date): string {
  return d.toISOString().replace(/\.\d{3}Z$/, "Z");
}

function stripHtml(html: string): string {
  return html.replace(/<[^>]+>/g, "").trim();
}

// =============================================================================
// Response Types
// =============================================================================

type GraphChat = {
  id?: string;
  topic?: string;
  chatType?: string;
  webUrl?: string;
};

type GraphChatList = {
  value?: GraphChat[];
};

type GraphMessage = {
  createdDateTime?: string;
  from?: { user?: { id?: string; displayName?: string } };
  body?: { content?: string; contentType?: string };
};

type GraphMessageList = {
  value?: GraphMessage[];
};

// =============================================================================
// Action Handler
// =============================================================================

async function handleListChats(
  graphToken: string,
  params: Record<string, unknown>,
): Promise<unknown> {
  const limit = readNumberParam(params, "limit", { integer: true }) ?? MAX_CHATS;
  const chatLimit = Math.min(limit, MAX_CHATS);

  const start = params.start_time
    ? new Date(String(params.start_time))
    : new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  const startStr = toISOClean(start);

  const engagementCutoff = new Date(Date.now() - ENGAGEMENT_WINDOW_DAYS * 24 * 60 * 60 * 1000);

  // Get current user ID for engagement filtering
  let currentUserId: string | null = null;
  try {
    const me = (await graphGet(graphToken, "/me")) as { id?: string };
    currentUserId = me.id ?? null;
  } catch {
    // Skip engagement filter if /me fails
  }

  // Fetch chat list
  const chatsData = (await graphGet(graphToken, `/me/chats?$top=${chatLimit}`)) as GraphChatList;
  const chats = chatsData.value ?? [];
  if (chats.length === 0) {
    return { count: 0, chats: [] };
  }

  // Fetch messages in batches
  const items: Array<{
    topic: string;
    chatType: string;
    messages: string;
    webUrl: string;
  }> = [];

  for (let i = 0; i < chats.length; i += BATCH_SIZE) {
    const batch = chats.slice(i, i + BATCH_SIZE);
    const requests = batch.map((c, idx) => ({
      id: String(idx),
      method: "GET",
      url: `/me/chats/${c.id}/messages?$filter=lastModifiedDateTime gt ${startStr}&$top=50`,
    }));

    const responses = await graphBatchPost(graphToken, requests);

    for (let j = 0; j < batch.length; j++) {
      const chat = batch[j];
      const resp = responses.find((r) => parseInt(r.id, 10) === j);
      const msgs = resp?.status === 200 ? ((resp.body as GraphMessageList)?.value ?? []) : [];
      if (msgs.length === 0) continue;

      // Only include chats where the user sent a message in the engagement window
      if (currentUserId) {
        const userEngaged = msgs.some((m: GraphMessage) => {
          const senderId = m.from?.user?.id;
          if (senderId !== currentUserId) return false;
          const sentAt = new Date(m.createdDateTime ?? 0);
          return sentAt >= engagementCutoff;
        });
        if (!userEngaged) continue;
      }

      const topic = chat.topic || (chat.chatType === "oneOnOne" ? "Personal chat" : chat.id ?? "");
      const webUrl = chat.webUrl ?? "";

      // Sort messages chronologically and format
      const sortedMsgs = [...msgs]
        .sort(
          (a, b) =>
            new Date(a.createdDateTime ?? 0).getTime() -
            new Date(b.createdDateTime ?? 0).getTime(),
        )
        .map((m) => {
          const from = m.from?.user?.displayName ?? "";
          const body = stripHtml(m.body?.content ?? "");
          return `${m.createdDateTime}: ${from} ${body}`;
        });

      // Truncate to max words, keeping the most recent messages
      const fullStr = sortedMsgs.join("\n");
      const words = fullStr.split(/\s+/);
      const messagesStr =
        (webUrl ? `[Source URL: ${webUrl}]\n` : "") +
        (words.length > MAX_WORDS_PER_CHAT
          ? words.slice(-MAX_WORDS_PER_CHAT).join(" ")
          : fullStr);

      items.push({
        topic,
        chatType: chat.chatType ?? "",
        messages: messagesStr,
        webUrl,
      });
    }
  }

  return { count: items.length, chats: items };
}

// =============================================================================
// Tool Factory
// =============================================================================

export function createTeamsChatTool(options: {
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
    label: "Teams Chat",
    name: "teams_chat",
    description:
      "Read Microsoft Teams chats and messages via Graph API. Lists recent chats with message history, filtered to chats where you actively participated. Requires NVIDIA SSO authentication. This tool is read-only.",
    parameters: TeamsChatSchema,
    execute: async (_toolCallId, args) => {
      const params = args as Record<string, unknown>;
      const action = readStringParam(params, "action", { required: true });

      try {
        const refreshToken = await options.getRefreshToken();
        if (!refreshToken) {
          return jsonResult({
            error: "not_authenticated",
            message: "Teams chat tool requires NVIDIA SSO authentication. Please log in first.",
          });
        }

        const tokenResult = await acquireDownstreamToken(refreshToken, TEAMS_SCOPES);
        if (!tokenResult.ok) {
          return jsonResult({
            error: "token_exchange_failed",
            message: `Failed to acquire Graph token: ${"error" in tokenResult ? tokenResult.error : "unknown"}`,
          });
        }
        const graphToken = tokenResult.accessToken;

        let result: unknown;
        switch (action) {
          case "list_chats":
            result = await handleListChats(graphToken, params);
            break;
          default:
            return jsonResult({
              error: "invalid_action",
              message: `Unknown action: ${action}. Valid: ${TEAMS_ACTIONS.join(", ")}`,
            });
        }

        return jsonResult({ action, result });
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        return jsonResult({ error: "teams_chat_error", message });
      }
    },
  } as AnyAgentTool;
}
