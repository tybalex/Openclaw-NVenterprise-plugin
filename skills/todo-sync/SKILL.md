---
name: todo-sync
description: Sync generated todo items to Microsoft OneDrive (personal_assistant/todos.json). Use after generating a todo list so items are persisted, or when the user asks to save/sync/push todos to OneDrive.
---

# Todo Sync to OneDrive

## When to Use

Use this skill when:
- You have just generated a todo list and need to persist the items to OneDrive
- The user explicitly asks to sync, save, or push their todos to OneDrive
- The user asks to see what todos are currently stored on OneDrive

**After every todo generation, you MUST call `todo_sync` with action `sync_todos` to persist the items.** Do not skip this step.

## Tools

### 1. Sync todos: `todo_sync` (action: `sync_todos`)

Upserts an array of todo items into the user's OneDrive file at `personal_assistant/todos.json`. New items are added; existing items (matched by `id`) are updated.

#### Parameters

- **todo_items** (REQUIRED, array): Array of todo item objects. Each item must have at least `id` and `title`.

#### Item fields

| Field | Type | Required | Description |
|---|---|---|---|
| `id` | string | Yes | UUID for the todo item (generate one per item, e.g. `crypto.randomUUID()` style) |
| `title` | string | Yes | Short task description |
| `description` | string | No | Longer context |
| `status` | string | No | `"pending"` (default), `"in_progress"`, `"completed"`, `"cancelled"` |
| `priority` | number | No | `0` = LOW, `1` = NORMAL (default), `2` = HIGH |
| `source_type` | string | No | `"slack"`, `"teams"`, `"outlook"`, `"event"`, `"other"`, `"manual"` |
| `source_title` | string | No | Title of the source item |
| `source_timestamp` | string | No | Timestamp of the source (YYYY-MM-DD HH:MM:SS) |
| `project_name` | string | No | Associated project name |
| `requestors` | array of strings | No | People who requested this action |
| `follow_up_action` | string | No | `"no_follow_up_needed"`, `"draft_email"`, `"draft_meeting"`, `"draft_message"` |
| `follow_up_title` | string | No | Subject for follow-up |
| `follow_up_content` | string | No | Body text for follow-up |

#### How to call

After generating todos with the todo-generation skill, take the `todo_items` array from the JSON code block you produced and pass it directly:

```json
{
  "todo_items": [
    {
      "id": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
      "title": "Review Q4 compliance report",
      "description": "Alice asked for review by Friday",
      "status": "pending",
      "priority": 2,
      "source_type": "outlook",
      "source_title": "RE: Q4 Compliance Review",
      "project_name": "Compliance"
    }
  ]
}
```

### 2. Fetch todos JSON: `todo_sync` (action: `fetch_todos`)

Returns the current contents of `personal_assistant/todos.json` from OneDrive.

#### Parameters

- **force_refresh** (optional, boolean): Bypass local cache and fetch directly from OneDrive. Default `false`.

#### How to call

```json
{}
```

or to bypass cache:

```json
{
  "force_refresh": true
}
```

### 3. Pull todos summaries: `todo_sync` (action: `get_summaries`)

Returns a concise, LLM-friendly text summary of existing todo items from OneDrive — just project names, titles, and statuses. Use this **before** generating new todos to avoid duplicates and reuse consistent project names.

#### Parameters

- **force_refresh** (optional, boolean): Bypass local cache and fetch directly from OneDrive. Default `false`.

#### How to call

```json
{}
```

## Workflow: Todo Generation + Sync

1. Generate todos using the **todo-generation** skill (fetch data, produce JSON code block).
2. Immediately after outputting the JSON code block, call `todo_sync` (action: `sync_todos`) with `{"todo_items": [...]}` where the array contains the items from your JSON block.
3. **Check the tool result.** If it succeeded, reply with one short confirmation (e.g. "Your todos have been synced to OneDrive."). If it failed (isError is true), tell the user the sync failed and include the error message. **Never claim success if the tool returned an error.**
4. Do NOT repeat, re-list, re-summarize, or re-format the todos — they are already displayed above in the rich UI.

## Important Notes

- Every todo item **must** have a unique `id`. Use UUID v4 format.
- The tool merges by `id` — calling it again with the same `id` updates that item rather than creating a duplicate.
- The tool reuses the Microsoft Graph authentication. Requires NVIDIA SSO authentication.
- The OneDrive file path is always `personal_assistant/todos.json` — do not change this.
