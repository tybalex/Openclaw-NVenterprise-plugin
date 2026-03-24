---
name: outlook-availability
description: Find available meeting times for people using their Outlook calendars. Returns Teams meeting deeplinks for easy scheduling. Use when the user asks to find when someone is available, schedule a meeting, or find free time slots.
---

# Outlook Availability

## When to Use

Use this skill when the user asks to:
- Find when someone is available for a meeting
- Check availability of multiple people
- Schedule a meeting and find the best time
- Find free time slots for a group
- See when people can meet
- Get a link to schedule a Teams meeting

## Tool

Call the **outlook_availability** tool with action **find_meeting_times**. Requires NVIDIA SSO authentication.

### Parameters

- **attendees** (required): Array of email addresses or names of people to check availability for. At least one attendee is required. If a name is provided (not an email), the tool will automatically search for the person's email address using Microsoft Graph API.
- **stakeholder** (optional): Name or email of the key stakeholder among the attendees. When specified, results are sorted to prioritize times when the stakeholder is free, and meeting rooms are selected based on the stakeholder's most frequent meeting locations (so the room is most convenient for them). The stakeholder should also be included in the `attendees` array.
- **start_time** (optional): Start of the search window in ISO 8601 format (Do not append timezone suffix if the user does not mention). Default: current time.
- **end_time** (optional): End of the search window in ISO 8601 format (Do not append timezone suffix if the user does not mention). Default: 7 days from start_time.
- **meeting_duration** (optional): Duration of the meeting in minutes. Default: 30 minutes. Must be between 5 and 1440 (24 hours).
- **is_organizer_optional** (optional): Whether the organizer is optional. Default: false.
- **return_suggestion_reasons** (optional): Whether to return reasons for each suggestion. Default: false.


### Example

For "find when john@example.com and jane@example.com are available", call with:
```json
{
  "attendees": ["john@example.com", "jane@example.com"]
}
```

For "find when John Smith and Jane Doe are available", call with:
```json
{
  "attendees": ["John Smith", "Jane Doe"]
}
```

For "find when the team is available, prioritizing VP Jane Doe's schedule", call with:
```json
{
  "attendees": ["John Smith", "Jane Doe", "Bob Lee"],
  "stakeholder": "Jane Doe"
}
```

For "find available times next week for a 1-hour meeting", call with:
```json
{
  "attendees": ["person@example.com"],
  "start_time": "2024-01-15T00:00:00Z",
  "end_time": "2024-01-21T23:59:59Z",
  "meeting_duration": 60
}
```

### Output

The tool returns a JSON object with `items` (array of available time suggestions). Each suggestion includes:
- **start**: Start time of the available slot (ISO 8601)
- **end**: End time of the available slot (ISO 8601)
- **confidence**: Confidence score (0.0 to 1.0)
- **score**: Quality score (0.0 to 1.0)
- **attendeeAvailability**: Array showing each attendee's availability status (free, tentative, busy, oof, workingElsewhere, unknown)
- **locations**: Suggested locations (if any)
- **suggestionReason**: Reason for the suggestion (if return_suggestion_reasons is true)
- **teamsDeeplink**: Teams meeting deeplink URL that opens Outlook/Teams with pre-filled meeting details (start time, end time, attendees). Users can click this link to schedule the meeting without manually entering details.


If the user is not authenticated, an error message is returned.
