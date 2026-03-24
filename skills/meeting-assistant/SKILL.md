---
name: meeting-assistant
description: Smart meeting scheduling assistant that gathers topic context via MS Copilot, identifies key people, finds available meeting times with room suggestions, and composes a ready-to-send meeting deeplink with the full pre-read included. Use when the user wants to schedule or prepare a meeting, especially when attendees or context need to be discovered.
---

# Meeting Assistant

**You can gather project context, identify relevant people via Microsoft Copilot, check calendar availability with meeting room suggestions, and compose a ready-to-send meeting invite with the full pre-read included.** This skill combines multiple tools to provide an intelligent meeting scheduling workflow.

## When to Use

Use this skill when the user asks to:
- Schedule or prepare a meeting (e.g. "prepare a meeting about Project Phoenix", "set up a meeting with the project leads")
- Find who they should meet with about a topic and when they're free
- Identify important contacts related to a project/topic and schedule time with them
- Set up a meeting with "the team", "stakeholders", or other vague groups
- Find availability for people involved in a recent discussion, email thread, or project
- Schedule a follow-up meeting based on recent communications
- Prepare a meeting with context and pre-read materials

## Workflow (Three Steps)

### Step 1: Generate Pre-Read Summary from All Sources

Gather project-related context to produce a pre-read summary that attendees can review ahead of the meeting.

Call all four tools **in parallel** — they are independent and can be called simultaneously:

1. **ms_copilot** (action: `query`) — broad synthesis question covering emails, Teams, OneDrive, and SharePoint:
   ```json
   {
     "question": "Summarize the latest status, recent decisions, open action items, and any relevant documents related to [topic from user's request]. Include key discussion points from emails and Teams chats in the past 2 weeks."
   }
   ```

2. **outlook_email** (action: `list_emails`) — all recent email threads (2-week window):
   ```json
   {
     "start_time": "<ISO 8601 timestamp 2 weeks ago>",
     "end_time": "<ISO 8601 timestamp now>",
     "limit": 50
   }
   ```

3. **teams_chat** (action: `list_chats`) — recent Teams chats the user participated in (2-week window):
   ```json
   {
     "start_time": "<ISO 8601 timestamp 2 weeks ago>",
     "end_time": "<ISO 8601 timestamp now>",
     "limit": 50
   }
   ```

4. **data--fetch_slack_messages** — recent Slack DMs, mentions, and channel messages (2-week window):
   ```json
   {
     "start_time": "<ISO 8601 timestamp 2 weeks ago>",
     "end_time": "<ISO 8601 timestamp now>",
     "limit": 50
   }
   ```

**Important notes on timestamps**: Do not append timezone suffix to ISO 8601 timestamps (e.g. use `2026-02-26T10:00:00` not `2026-02-26T10:00:00Z`).

**After all tools return**, scan each result set for content relevant to the meeting topic — look for keyword matches in subject lines, message content, and email bodies. Extract items that mention the meeting topic, related projects, or involved people.

**Synthesize from all available sources**:
- Use ms_copilot results as the high-level synthesis and for document references
- Use Outlook email results for threaded conversation context and decisions (preserve source URLs where relevant)
- Use Teams chat results for recent real-time discussion context
- Use Slack results for channel discussions, mentions, and DM context
- If ms_copilot returns vague or empty results, the direct data tools provide the reliable baseline
- If any single tool fails or returns no relevant results, continue with what's available — do not fail the whole step

#### Compile and Display the Pre-Read

After the tool returns, format the output as a rich markdown pre-read summary inside a `card` fenced block. This renders the pre-read as a styled card with **Copy** (plain text) and **Copy Markdown** buttons:

````card
## Pre-Read: [Meeting Topic]

### Recent Context
- [Key status updates, decisions, and action items from Copilot]

### Relevant Documents
- [Any documents, links, or references mentioned by Copilot]

### Key Risks *(include only if Copilot surfaces any blockers, concerns, or unresolved issues)*
- [Risk or blocker, if any]

### Suggested Agenda
1. [Agenda item derived from open items / recent discussions]
2. [Agenda item derived from document content]
3. [Agenda item for decisions needed]
````

**Critical format rules**:
- Use exactly **three backticks** (` ``` `) followed by `card` to open the fence
- Place ALL pre-read content **inside** the fence — between the opening ` ```card ` and the closing ` ``` `
- Do NOT output the pre-read text outside or after the fence markers
- The fence MUST contain actual content, not be left empty

Present the pre-read summary to the user and **explicitly ask for confirmation** before proceeding.

#### Confirm the Pre-Read Summary with the User

You **MUST** present the compiled pre-read summary to the user and wait for their confirmation before proceeding to Step 2. Ask the user in a natural, conversational way. For example:

> Here's the pre-read summary I've put together for the meeting:
>
> [Pre-read summary]
>
> Does this look right, or do you want anything changed before I identify the attendees?

You may also say:

> If this looks good, I'll move on to identifying the core attendees. If not, tell me what you'd like to change.

**Do NOT proceed to Step 2 until the user confirms the pre-read summary.** The user may:
- **Confirm** the summary as-is → proceed to Step 2.
- **Request edits** (e.g. "add the timeline discussion", "remove the budget section") → update the summary accordingly and present the revised version for confirmation again.
- **Add context** the user knows but Copilot missed → incorporate it into the summary and confirm again.
- **Ask to regenerate** with a different focus → call **ms_copilot** again with an adjusted question (and optionally re-call the direct data tools) and compile a new summary.

Once confirmed, use the **approved version** of the pre-read summary for all subsequent steps.

### Step 2: Identify People with MS Copilot

Call the **ms_copilot** tool (action: `query`) to ask Microsoft 365 Copilot who the **core people** are for this meeting topic. Copilot has access to the user's Outlook emails, Teams chats, calendar, OneDrive, and SharePoint — so it can identify the right attendees based on context.

**Important**: Only include **core people** who are directly involved and essential to the meeting topic. Avoid adding peripheral contacts or people only loosely related — keep the attendee list focused and small to make scheduling easier and meetings more productive.

**Craft a specific question** based on the user's request. Examples:

- User says "schedule a meeting about Project Phoenix" → ask Copilot: "Who are the core people directly involved in Project Phoenix? Only include people who are essential to the discussion. List their full names and email addresses."
- User says "set up a sync with my team" → ask Copilot: "Who are the core members of my immediate team that I interact with most frequently? List their full names and email addresses."
- User says "follow up on the budget discussion" → ask Copilot: "Who are the core people directly involved in recent budget discussions? Only include essential participants. List their full names and email addresses."

```json
{
  "question": "Who are the core people directly involved in [topic from user's request]? Only include people who are essential to the discussion — do not include peripheral contacts. Please list their full names and email addresses."
}
```

**After getting the Copilot response**, extract the names and/or email addresses from the reply. You **MUST** present the identified people to the user and wait for their confirmation before proceeding to Step 3.

#### Confirm Attendees with the User

Present the attendee list in a clear format and explicitly ask the user to confirm, add, or remove people. **Also ask if any attendee is a key stakeholder** whose schedule and location should be prioritized. Keep the phrasing natural and concise. For example:

> Based on your recent communications, these look like the core attendees for this meeting:
>
> 1. **Alice Smith** — alice@example.com
> 2. **Bob Jones** — bob@example.com
> 3. **Carol Lee** — carol@example.com
>
> Do you want to keep this list, or add/remove anyone?
>
> Also, is there a key stakeholder I should prioritize when looking for time and a meeting room?

**Do NOT proceed to Step 3 until the user confirms the attendee list.** The user may:
- **Confirm** the list as-is → proceed to Step 3.
- **Add** additional people (by name or email) → update the list and confirm again.
- **Remove** people from the list → update the list and confirm again.
- **Replace** the entire list with their own attendees → use the user-provided list and proceed.
- **Designate a stakeholder** → note the stakeholder name and pass it as the `stakeholder` parameter in Step 3.

### Step 3: Find Availability, Meeting Rooms, and Generate Deeplinks

#### Collect Meeting Preferences Before Searching

Before calling the tool, ask the user two quick questions in one message. Mention the defaults so they can just confirm or override:

> How long should the meeting be — **30 minutes** works for most syncs, or did you have something else in mind? And any timing preferences (e.g. this week, next Monday afternoon, avoid Fridays)? I'll search the next 7 days by default.

If the user already mentioned duration or timing earlier in the conversation, use those values and skip asking. Use the user's answer to set `meeting_duration`, `start_time`, and `end_time` before calling the tool.

Once preferences are confirmed, call the **outlook_availability** tool (action: `find_meeting_times`) to find common free time slots, available meeting rooms, and generate deeplinks.

**Important**: The meeting search must account for the current user / organizer as well, not just the other attendees. Do **not** assume the user is free at any time. By default, keep `is_organizer_optional` as `false` so the tool includes the organizer's calendar when finding times. Only set `is_organizer_optional: true` if the user explicitly says their attendance is optional or flexible.

#### Compose the `subject` and `body`

Before calling the tool, compose the meeting subject and body:

- **subject**: A descriptive meeting title based on the user's topic (e.g. `"Q3 Budget Review Sync"`).
- **body**: The full pre-read summary from Step 1 as plain text:

```
Pre-Read: [Meeting Topic]

Recent Context:
- [Key points from Step 1]

Open Action Items:
- [Items from Step 1]

Suggested Agenda:
1. [Agenda items from Step 1]
```

#### Parameters

- **attendees** (required): Array of attendee **names** (not email addresses) confirmed in Step 2. At least one attendee is required. Always pass full names — the tool will automatically resolve each name to the corresponding email address using Microsoft Graph API.
- **stakeholder** (optional): The **name** of the key stakeholder among the attendees. When specified, the tool will: (1) prioritize time slots where the stakeholder is free, (2) find meeting rooms based on the stakeholder's most frequent meeting locations (so the room is most convenient for them). The stakeholder must also be included in the `attendees` array.
- **subject** (required): The meeting subject/title to include in the deeplinks.
- **body** (required): The full pre-read summary from Step 1.
- **start_time** (optional): Start of the search window in ISO 8601 format (Do not append timezone suffix if the user does not mention). Default: current time.
- **end_time** (optional): End of the search window in ISO 8601 format (Do not append timezone suffix if the user does not mention). Default: 7 days from start_time.
- **meeting_duration** (optional): Duration of the meeting in minutes. Default: 30 minutes. Must be between 5 and 1440 (24 hours).
- **is_organizer_optional** (optional): Whether the organizer is optional. Default: false. Leave this as `false` unless the user explicitly says their attendance is optional.
- **return_suggestion_reasons** (optional): Whether to return reasons for each suggestion. Default: false.

#### Example

```json
{
  "attendees": ["Alice Smith", "Bob Jones", "Carol Lee"],
  "stakeholder": "Carol Lee",
  "subject": "Q3 Budget Review Sync",
  "body": "Pre-Read: Q3 Budget Review\n\nRecent Context:\n- Budget allocation finalized last week\n- Pending approval from finance team\n\nSuggested Agenda:\n1. Review Q3 forecast\n2. Discuss open items",
  "meeting_duration": 30
}
```

#### Output

The tool returns a JSON object with:
- **items**: Array of available time suggestions. Each suggestion includes:
  - **start**: Start time of the available slot (ISO 8601)
  - **end**: End time of the available slot (ISO 8601)
  - **confidence**: Confidence score (0.0 to 1.0)
  - **score**: Quality score (0.0 to 1.0)
  - **attendeeAvailability**: Array showing each attendee's availability status (free, tentative, busy, oof, workingElsewhere, unknown)
  - **locations**: Suggested locations (if any)
  - **availableRooms**: Array of available meeting rooms for this time slot (name, email, address)
  - **teamsDeeplink**: Pre-filled Outlook deeplink URL containing the subject, body (with full pre-read), attendees, and an available meeting room (if found)
  - **suggestionReason**: Reason for the suggestion (if return_suggestion_reasons is true)
- **frequentMeetingRooms**: Most frequently used meeting room buildings across attendees
- **roomsInFrequentBuildings**: All rooms in the most frequent building(s)
- **ambiguousAttendees** *(only present when disambiguation is needed)*: Array of `{ query, candidates }` where each `candidate` has `{ email, displayName, jobTitle?, department? }`

#### Handling Ambiguous Attendees

If the tool returns `ambiguousAttendees` (and `items` is empty), **do not proceed** — one or more names matched multiple people. You must:

1. Present the candidates to the user in a clear list, e.g.:
   > "**Bob** matched multiple people — which one did you mean?
   > 1. Bob Smith \<bob.smith@nvidia.com\> — Senior Engineer (Networking)
   > 2. Bob Jones \<bob.jones@nvidia.com\> — Product Manager (AI)"
2. Wait for the user to confirm the correct person.
3. Re-call **outlook_availability** (action: `find_meeting_times`) with the confirmed **email address** (not the name) for that attendee.

#### No Slots Found

If `items` is empty and there are no `ambiguousAttendees`, do not silently fail. Instead:

1. Tell the user no common free time was found in the search window.
2. Offer concrete next steps — pick one or more that fit the situation:
   - **Extend the window**: re-call the tool with a wider `end_time` (e.g. +7 more days)
   - **Mark someone optional**: ask the user if any attendee's presence is flexible, then re-call with `is_organizer_optional: true` or by removing that attendee
   - **Shorten the meeting**: a shorter duration (e.g. 30 → 15 min) often opens up more slots — re-call with reduced `meeting_duration`
   - **Let the user decide**: if all options are exhausted, tell the user and let them pick a time manually

#### Present the Final Result

Show the user the available time slots. For each slot include:
- The time range
- Available meeting rooms (if any)
- The **teamsDeeplink** as a clickable link

Make it clear that **clicking a link opens a pre-filled Outlook/Teams invite — the user still needs to review it and click Send**. The meeting is not booked until they do. Example phrasing:

> Here are the available times. Click any link to open a pre-filled meeting invite in Outlook — just review and hit **Send** to book it.

The invite will be pre-filled with:
- The meeting subject
- All confirmed attendees (plus an available meeting room)
- The full pre-read summary in the invite body

## Important Notes

- **Always confirm the pre-read summary** with the user after Step 1 before proceeding to Step 2.
- **Always confirm attendees** with the user after Step 2 before proceeding to Step 3.
- If the user already provides specific names or emails, you may skip Step 2.
- If the user explicitly says they don't need a pre-read, skip Step 1 and go directly to Step 2 to identify people and then Step 3 for availability (leave the deeplink body empty or use a short description).
- If MS Copilot cannot identify people (e.g. vague request), ask the user for clarification.
- Requires NVIDIA SSO authentication. If the user is not authenticated, an error message is returned from MS Copilot and availability tools.

## Example Conversation Flow

**User**: "Can you help me schedule a meeting about the Q3 budget review?"

1. Call all four tools **in parallel**:
   - **ms_copilot** (action: `query`): `{"question": "Summarize the latest status, recent decisions, open action items, and any relevant documents related to the Q3 budget review. Include key discussion points from emails and Teams chats in the past 2 weeks."}`
   - **outlook_email** (action: `list_emails`): `{"start_time": "<2 weeks ago>", "end_time": "<now>", "limit": 50}`
   - **teams_chat** (action: `list_chats`): `{"start_time": "<2 weeks ago>", "end_time": "<now>", "limit": 50}`
   - **data--fetch_slack_messages**: `{"start_time": "<2 weeks ago>", "end_time": "<now>", "limit": 50}`
2. Scan all results for content related to "Q3 budget review", synthesize into the pre-read card, then display it and ask for confirmation in natural language before moving on.
3. User confirms the summary (or requests revisions — update and re-confirm until approved).
4. Call **ms_copilot** (action: `query`) with: `{"question": "Who are the key people I've been communicating with about the Q3 budget review? List their full names and email addresses."}`
5. Present attendees to the user and ask, in natural language, whether they want to keep the list or add/remove anyone. Also ask whether there is a key stakeholder whose schedule should be prioritized.
6. User confirms attendees and designates Carol Lee as the stakeholder → Ask: "How long should the meeting be, and do you have any timing preferences?"
7. User says "30 minutes, this week if possible" → Call **outlook_availability** (action: `find_meeting_times`):
   ```json
   {
     "attendees": ["Alice Smith", "Bob Jones", "Carol Lee"],
     "stakeholder": "Carol Lee",
     "subject": "Q3 Budget Review Sync",
     "body": "Pre-Read: Q3 Budget Review\n\nRecent Context:\n- Budget allocation finalized\n\nSuggested Agenda:\n1. Review forecast\n2. Discuss open items",
     "meeting_duration": 30
   }
   ```
8. Present available slots with deeplinks — tell the user to click a link to open the pre-filled Outlook invite, then review and hit Send to book it.

### Organizer Availability Rule

- Always account for the current user's availability when finding times.
- Do **not** describe the results as "common free time" unless the organizer is included too.
- Do **not** assume the user is available simply because they did not list themselves as an attendee.
- Keep `is_organizer_optional` at its default of `false` unless the user explicitly says they are optional for the meeting.
