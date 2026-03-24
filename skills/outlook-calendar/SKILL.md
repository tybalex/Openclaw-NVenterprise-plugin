---
name: outlook-calendar
description: Fetch Outlook calendar events for a time window. Use when the user asks about their calendar, meetings, or events.
---

# Outlook Calendar

## When to Use

Use this skill when the user asks to:
- See their calendar or upcoming meetings
- Get events for a date or time range
- Review Outlook calendar for a window

## Tool

Call the **outlook_calendar** tool with action **list_events**. Requires NVIDIA SSO authentication.

### Parameters

- **start_time** (optional): Start of the window in ISO 8601. Default: now.
- **end_time** (optional): End of the window in ISO 8601. Default: 24 hours after start_time.
- **limit** (optional): Max number of events (default 50).

### Example

For "what's on my calendar today", set start_time and end_time to today's date range in ISO format.

### Output

The tool returns a JSON object with `items` (array of events: subject, start, end, organizer, preview, webUrl) and `count`. If the user is not authenticated, an error message is returned.
