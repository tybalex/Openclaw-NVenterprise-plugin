---
name: teams-chat
description: Fetch recent Microsoft Teams chats and messages. Use when the user asks about their Teams chats or messages from the last N days.
---

# Teams Chat

## When to Use

Use this skill when the user asks to:
- See their recent Teams chats or messages
- Get Teams messages from the last N days
- Review Teams conversation history

## Tool

Call the **teams_chat** tool with action **list_chats**. Requires NVIDIA SSO authentication.

### Parameters

- **start_time**: Start of the window in ISO 8601. Default: 7 days ago.
- **end_time** (optional): End of the window in ISO 8601.
- **limit**: Max number of chats to fetch (must not exceed 50).

### Example

For "get my Teams messages from the last 7 days", call with no arguments or set start_time to 7 days ago in ISO format.

### Output

The tool returns a JSON object with `items` (array of chats: id, topic, chatType, messages, webUrl, source_type) and `count`. If the user is not authenticated, an error message is returned.
