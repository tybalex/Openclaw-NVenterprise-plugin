---
name: outlook-emails
description: Fetch recent Outlook emails (Inbox and Sent) for a time window. Use when the user asks to see their emails, review inbox, or get email from the last N days.
---

# Outlook Emails

## When to Use

Use this skill when the user asks to:
- See their recent emails, inbox, or sent messages
- Get emails from the last N days
- Review Outlook mail for a time period

## Tool

Call the **outlook_email** tool with action **list_emails**. Requires NVIDIA SSO authentication.

### Parameters

- **start_time** (optional): Start of the window in ISO 8601 (e.g. `2025-02-01T00:00:00Z`). Default: 7 days before end_time.
- **end_time** (optional): End of the window in ISO 8601. Default: now.
- **limit** (optional): Max number of email threads to return (default 50, max 500).

### Example

For "get my emails from the last 7 days", call with no arguments (defaults to last 7 days) or set:
- `start_time`: 7 days ago in ISO format
- `end_time`: now in ISO format

### Output

The tool returns a JSON object with `items` (array of email threads: subject, preview, sender, thread_content, webUrl) and `count`. If the user is not authenticated, an error message is returned.
