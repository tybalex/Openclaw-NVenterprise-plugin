/**
 * Microsoft 365 Copilot tool.
 *
 * Queries M365 Copilot via Graph beta API to get AI-generated answers
 * based on the user's Microsoft 365 data (emails, chats, files, etc.).
 *
 * API flow:
 *   1. POST /beta/copilot/conversations  → create conversation
 *   2. POST /beta/copilot/conversations/{id}/chat → send question
 *   3. Return the assistant's reply
 *
 * Uses Azure AD refresh token -> OBO for Copilot-specific scopes.
 * NOTE: This uses the Graph beta API which may change without notice.
 * Requires M365 Copilot license.
 */

import { Type } from "@sinclair/typebox";
import { jsonResult, readStringParam } from "openclaw/plugin-sdk/agent-runtime";
import { stringEnum, type AnyAgentTool } from "openclaw/plugin-sdk/core";
import { acquireDownstreamToken, isAzureOBOConfigured } from "./azure-obo.js";

// =============================================================================
// Configuration
// =============================================================================

const COPILOT_BASE = "https://graph.microsoft.com/beta/copilot";
const COPILOT_SCOPES = "https://graph.microsoft.com/Chat.ReadWrite";
const HTTP_TIMEOUT_MS = 240_000; // 4 minutes — Copilot can be slow

// =============================================================================
// Schema
// =============================================================================

const COPILOT_ACTIONS = ["query"] as const;

const MSCopilotSchema = Type.Object({
  action: stringEnum(COPILOT_ACTIONS, {
    description: "Action to perform: query (ask Microsoft 365 Copilot a question).",
  }),
  question: Type.String({
    description:
      "The question to ask Microsoft 365 Copilot. It can answer based on your emails, Teams chats, calendar, OneDrive files, SharePoint, etc.",
  }),
});

// =============================================================================
// Action Handler
// =============================================================================

async function handleQuery(graphToken: string, params: Record<string, unknown>): Promise<unknown> {
  const question = readStringParam(params, "question", { required: true });

  const headers = {
    Authorization: `Bearer ${graphToken}`,
    "Content-Type": "application/json",
  };

  // Step 1: Create conversation
  const createRes = await fetch(`${COPILOT_BASE}/conversations`, {
    method: "POST",
    headers,
    body: JSON.stringify({}),
    signal: AbortSignal.timeout(HTTP_TIMEOUT_MS),
  });

  if (!createRes.ok) {
    const text = await createRes.text();
    return { error: `Failed to create Copilot conversation (${createRes.status}): ${text}` };
  }

  const { id: conversationId } = (await createRes.json()) as { id?: string };
  if (!conversationId) {
    return { error: "Copilot returned empty conversation ID" };
  }

  // Step 2: Send question
  const chatRes = await fetch(`${COPILOT_BASE}/conversations/${conversationId}/chat`, {
    method: "POST",
    headers,
    body: JSON.stringify({
      message: { text: question.trim() },
      locationHint: {
        timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone || "America/Chicago",
      },
    }),
    signal: AbortSignal.timeout(HTTP_TIMEOUT_MS),
  });

  if (!chatRes.ok) {
    const text = await chatRes.text();
    return { error: `Copilot chat failed (${chatRes.status}): ${text}` };
  }

  const chatData = (await chatRes.json()) as { messages?: Array<{ text?: string }> };

  // Step 3: Extract reply
  const messages = chatData.messages;
  if (Array.isArray(messages) && messages.length > 1 && messages[1]?.text) {
    return { reply: messages[1].text };
  }
  if (Array.isArray(messages) && messages.length > 0) {
    const last = messages[messages.length - 1];
    if (last?.text) return { reply: last.text };
  }

  return { error: "No response from Microsoft Copilot" };
}

// =============================================================================
// Tool Factory
// =============================================================================

export function createMSCopilotTool(options: {
  getRefreshToken: () => string | null | Promise<string | null>;
  enabled?: boolean;
}): AnyAgentTool | null {
  if (options.enabled === false) return null;
  if (!isAzureOBOConfigured()) return null;

  return {
    label: "Microsoft Copilot",
    name: "ms_copilot",
    description:
      "Query Microsoft 365 Copilot to get AI-generated answers based on your M365 data (Outlook emails, Teams chats, calendar, OneDrive files, SharePoint). Requires NVIDIA SSO authentication and M365 Copilot license. Uses beta API.",
    parameters: MSCopilotSchema,
    execute: async (_toolCallId, args) => {
      const params = args as Record<string, unknown>;
      const action = readStringParam(params, "action", { required: true });

      try {
        const refreshToken = await options.getRefreshToken();
        if (!refreshToken) return jsonResult({ error: "not_authenticated", message: "Please log in first." });

        const tokenResult = await acquireDownstreamToken(refreshToken, COPILOT_SCOPES);
        if (!tokenResult.ok) {
          return jsonResult({ error: "token_exchange_failed", message: `Failed to acquire token: ${"error" in tokenResult ? tokenResult.error : "unknown"}` });
        }

        let result: unknown;
        switch (action) {
          case "query":
            result = await handleQuery(tokenResult.accessToken, params);
            break;
          default:
            return jsonResult({ error: "invalid_action", message: `Unknown action: ${action}` });
        }

        return jsonResult({ action, result });
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        return jsonResult({ error: "ms_copilot_error", message });
      }
    },
  } as AnyAgentTool;
}
