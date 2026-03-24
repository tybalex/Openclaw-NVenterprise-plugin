/**
 * Todo Sync tool.
 *
 * Syncs todo items to/from OneDrive as a JSON file (personal_assistant/todos.json).
 * Upserts by ID, with schema validation matching the Python Pydantic models.
 *
 * Uses Azure AD refresh token -> OBO for Graph Files.ReadWrite scope.
 */

import { Type } from "@sinclair/typebox";
import { jsonResult, readStringParam } from "openclaw/plugin-sdk/agent-runtime";
import { stringEnum, type AnyAgentTool } from "openclaw/plugin-sdk/core";
import { acquireDownstreamToken, isAzureOBOConfigured } from "./azure-obo.js";
import crypto from "node:crypto";

// =============================================================================
// Configuration
// =============================================================================

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const TODOS_SCOPES = "https://graph.microsoft.com/Files.ReadWrite";
const TODOS_PATH = "personal_assistant/todos.json";
const DEFAULT_TIMEOUT_MS = 30_000;

// =============================================================================
// Todo Types (mirrors Python Pydantic models)
// =============================================================================

const TODO_STATUSES = ["pending", "in_progress", "completed", "cancelled"] as const;
const TODO_SOURCES = ["slack", "teams", "outlook", "event", "other", "manual"] as const;
const TODO_FOLLOW_UPS = ["no_follow_up_needed", "draft_email", "draft_meeting", "draft_message"] as const;

interface TodoItem {
  id: string;
  title: string | null;
  description: string | null;
  status: string;
  priority: number;
  due_date: string | null;
  created_at: string;
  updated_at: string;
  tags: string[];
  project_name: string | null;
  source_type: string | null;
  source_url: string | null;
  source_title: string | null;
  source_requestor: string | null;
  source_timestamp: string | null;
  data_index: number;
  version: number;
  history: any[];
  additional_fields: Record<string, unknown>;
  follow_up_action: string | null;
  follow_up_title: string | null;
  follow_up_content: string | null;
  follow_up_metadata: { thread_id: string; created_at: string; action_str: string };
  requestors: string[];
}

interface TodosFile {
  todos: Record<string, TodoItem>;
  last_updated: string | null;
  last_generated: string | null;
}

// =============================================================================
// Schema
// =============================================================================

const TODO_ACTIONS = ["sync_todos", "fetch_todos", "get_summaries"] as const;

const TodoSyncSchema = Type.Object({
  action: stringEnum(TODO_ACTIONS, {
    description:
      "Action: sync_todos (upsert items), fetch_todos (get full JSON), get_summaries (concise list for dedup).",
  }),
  todo_items: Type.Optional(
    Type.Array(Type.Any(), {
      description:
        "Array of todo objects to sync (for sync_todos). Each must have an id. Fields: id, title, description, status, priority, due_date, tags, source_type, source_url, requestors, follow_up_action, etc.",
    }),
  ),
});

// =============================================================================
// Coercion Helpers
// =============================================================================

function asStr(v: unknown): string {
  return typeof v === "string" ? v : "";
}

function asStrNull(v: unknown): string | null {
  return typeof v === "string" ? v : null;
}

function asStrArr(v: unknown): string[] {
  return Array.isArray(v) ? v.filter((x): x is string => typeof x === "string") : [];
}

function coerceStatus(v: unknown): string {
  if (typeof v === "string" && (TODO_STATUSES as readonly string[]).includes(v.toLowerCase())) return v.toLowerCase();
  return "pending";
}

function coercePriority(v: unknown): number {
  if (typeof v === "string") {
    if (v.toLowerCase() === "high") return 2;
    if (v.toLowerCase() === "medium") return 1;
    if (v.toLowerCase() === "low") return 0;
  }
  if (typeof v === "number" && [0, 1, 2].includes(v)) return v;
  return 1;
}

function coerceSource(v: unknown): string | null {
  if (v == null) return "manual";
  if (typeof v === "string" && (TODO_SOURCES as readonly string[]).includes(v.toLowerCase())) return v.toLowerCase();
  if (typeof v === "string" && v.toLowerCase() === "email") return "outlook";
  return "other";
}

function coerceFollowUp(v: unknown): string | null {
  if (v == null) return "draft_message";
  if (typeof v === "string" && (TODO_FOLLOW_UPS as readonly string[]).includes(v.toLowerCase())) return v.toLowerCase();
  return "draft_message";
}

function normalizeTodoItem(raw: Record<string, unknown>, existing?: Partial<TodoItem>): TodoItem {
  const now = new Date().toISOString();
  return {
    id: asStr(raw.id) || asStr(existing?.id) || crypto.randomUUID(),
    title: asStrNull(raw.title) ?? asStrNull(existing?.title) ?? null,
    description: asStrNull(raw.description) ?? asStrNull(existing?.description) ?? "",
    status: coerceStatus(raw.status ?? existing?.status),
    priority: coercePriority(raw.priority ?? existing?.priority),
    due_date: asStrNull(raw.due_date) ?? asStrNull(existing?.due_date) ?? null,
    created_at: asStr(existing?.created_at) || asStr(raw.created_at) || now,
    updated_at: now,
    tags: asStrArr(raw.tags).length ? asStrArr(raw.tags) : asStrArr(existing?.tags),
    project_name: asStrNull(raw.project_name) ?? asStrNull(existing?.project_name) ?? "",
    source_type: coerceSource(raw.source_type ?? existing?.source_type),
    source_url: asStrNull(raw.source_url) ?? asStrNull(existing?.source_url) ?? "",
    source_title: asStrNull(raw.source_title) ?? asStrNull(existing?.source_title) ?? "",
    source_requestor: asStrNull(raw.source_requestor) ?? asStrNull(existing?.source_requestor) ?? "",
    source_timestamp: asStrNull(raw.source_timestamp) ?? asStrNull(existing?.source_timestamp) ?? "",
    data_index: typeof raw.data_index === "number" ? raw.data_index : (existing?.data_index ?? -1),
    version: ((existing?.version ?? 0) as number) + 1,
    history: Array.isArray(existing?.history) ? existing.history : [],
    additional_fields: (raw.additional_fields && typeof raw.additional_fields === "object" ? raw.additional_fields : existing?.additional_fields ?? {}) as Record<string, unknown>,
    follow_up_action: coerceFollowUp(raw.follow_up_action ?? existing?.follow_up_action),
    follow_up_title: asStrNull(raw.follow_up_title) ?? asStrNull(existing?.follow_up_title) ?? "",
    follow_up_content: asStrNull(raw.follow_up_content) ?? asStrNull(existing?.follow_up_content) ?? "",
    follow_up_metadata: {
      thread_id: asStr((raw.follow_up_metadata as any)?.thread_id) || asStr(existing?.follow_up_metadata?.thread_id) || "",
      created_at: asStr((raw.follow_up_metadata as any)?.created_at) || asStr(existing?.follow_up_metadata?.created_at) || now,
      action_str: asStr((raw.follow_up_metadata as any)?.action_str) || asStr(existing?.follow_up_metadata?.action_str) || "",
    },
    requestors: asStrArr(raw.requestors).length ? asStrArr(raw.requestors) : asStrArr(existing?.requestors),
  };
}

function normalizeTodosFile(raw: unknown): TodosFile {
  const obj = (raw && typeof raw === "object" && !Array.isArray(raw) ? raw : {}) as Record<string, unknown>;
  return {
    todos: (obj.todos && typeof obj.todos === "object" && !Array.isArray(obj.todos) ? obj.todos : {}) as Record<string, TodoItem>,
    last_updated: asStrNull(obj.last_updated),
    last_generated: asStrNull(obj.last_generated),
  };
}

// =============================================================================
// OneDrive Helpers
// =============================================================================

async function readTodosFromOneDrive(token: string): Promise<TodosFile> {
  const url = `${GRAPH_BASE_URL}/me/drive/root:/${TODOS_PATH}:/content`;
  const res = await fetch(url, {
    method: "GET",
    headers: { Authorization: `Bearer ${token}` },
    signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
  });

  if (res.status === 404) {
    return { todos: {}, last_updated: null, last_generated: null };
  }
  if (!res.ok) throw new Error(`Read failed (${res.status}): ${await res.text()}`);

  const data = await res.json();
  return normalizeTodosFile(data);
}

async function writeTodosToOneDrive(token: string, todosFile: TodosFile): Promise<void> {
  const url = `${GRAPH_BASE_URL}/me/drive/root:/${TODOS_PATH}:/content`;
  const res = await fetch(url, {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(todosFile, null, 2),
    signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
  });
  if (!res.ok) throw new Error(`Write failed (${res.status}): ${await res.text()}`);
}

// =============================================================================
// Action Handlers
// =============================================================================

async function handleSyncTodos(token: string, params: Record<string, unknown>): Promise<unknown> {
  const items = params.todo_items;
  if (!Array.isArray(items) || items.length === 0) {
    return { error: "todo_items array is required and must not be empty." };
  }

  // Read existing
  const todosFile = await readTodosFromOneDrive(token);

  // Merge (upsert by id)
  let added = 0;
  let updated = 0;
  for (const raw of items) {
    if (!raw || typeof raw !== "object") continue;
    const rawObj = raw as Record<string, unknown>;
    const id = asStr(rawObj.id) || crypto.randomUUID();
    rawObj.id = id;

    const existing = todosFile.todos[id];
    todosFile.todos[id] = normalizeTodoItem(rawObj, existing);

    if (existing) updated++;
    else added++;
  }

  todosFile.last_updated = new Date().toISOString();

  // Write back
  await writeTodosToOneDrive(token, todosFile);

  return {
    added,
    updated,
    total: Object.keys(todosFile.todos).length,
  };
}

async function handleFetchTodos(token: string): Promise<unknown> {
  const todosFile = await readTodosFromOneDrive(token);
  return {
    total: Object.keys(todosFile.todos).length,
    last_updated: todosFile.last_updated,
    todos: todosFile.todos,
  };
}

async function handleGetSummaries(token: string): Promise<unknown> {
  const todosFile = await readTodosFromOneDrive(token);
  const summaries = Object.values(todosFile.todos).map((t) => ({
    id: t.id,
    title: t.title,
    status: t.status,
    priority: t.priority,
    source_type: t.source_type,
    due_date: t.due_date,
  }));
  return { total: summaries.length, summaries };
}

// =============================================================================
// Tool Factory
// =============================================================================

export function createTodoSyncTool(options: {
  getRefreshToken: () => string | null | Promise<string | null>;
  enabled?: boolean;
}): AnyAgentTool | null {
  if (options.enabled === false) return null;
  if (!isAzureOBOConfigured()) return null;

  return {
    label: "Todo Sync",
    name: "todo_sync",
    description:
      "Sync todo items to/from OneDrive (personal_assistant/todos.json). Use to persist generated todos, fetch existing todos, or get concise summaries for deduplication. Requires NVIDIA SSO authentication.",
    parameters: TodoSyncSchema,
    execute: async (_toolCallId, args) => {
      const params = args as Record<string, unknown>;
      const action = readStringParam(params, "action", { required: true });

      try {
        const refreshToken = await options.getRefreshToken();
        if (!refreshToken) return jsonResult({ error: "not_authenticated", message: "Please log in first." });

        const tokenResult = await acquireDownstreamToken(refreshToken, TODOS_SCOPES);
        if (!tokenResult.ok) {
          return jsonResult({ error: "token_exchange_failed", message: `Failed to acquire token: ${"error" in tokenResult ? tokenResult.error : "unknown"}` });
        }
        const graphToken = tokenResult.accessToken;

        let result: unknown;
        switch (action) {
          case "sync_todos":
            result = await handleSyncTodos(graphToken, params);
            break;
          case "fetch_todos":
            result = await handleFetchTodos(graphToken);
            break;
          case "get_summaries":
            result = await handleGetSummaries(graphToken);
            break;
          default:
            return jsonResult({ error: "invalid_action", message: `Unknown action: ${action}. Valid: ${TODO_ACTIONS.join(", ")}` });
        }

        return jsonResult({ action, result });
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        return jsonResult({ error: "todo_sync_error", message });
      }
    },
  } as AnyAgentTool;
}
