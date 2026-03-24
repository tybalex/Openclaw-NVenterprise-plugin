/**
 * Outlook Availability / Meeting Time Finder tool.
 *
 * Multi-step scheduling workflow:
 * 1. Resolve attendee names to email addresses (fuzzy search via /me/people, /users)
 * 2. Find available meeting times (/me/findMeetingTimes)
 * 3. Find frequent meeting room locations from attendee calendars
 * 4. Check room availability (/me/calendar/getSchedule)
 * 5. Generate Outlook meeting deeplinks with room + attendees
 *
 * Uses Azure AD refresh token -> OBO for Graph Calendars.Read, People.Read, User.Read.All scopes.
 */

import { Type } from "@sinclair/typebox";
import { jsonResult, readNumberParam, readStringParam, readStringArrayParam } from "openclaw/plugin-sdk/agent-runtime";
import { stringEnum, type AnyAgentTool } from "openclaw/plugin-sdk/core";
import { acquireDownstreamToken, isAzureOBOConfigured } from "./azure-obo.js";

// =============================================================================
// Configuration
// =============================================================================

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const AVAILABILITY_SCOPES =
  "https://graph.microsoft.com/Calendars.Read https://graph.microsoft.com/People.Read https://graph.microsoft.com/User.Read.All";

const DEFAULT_TIMEOUT_MS = 60_000;
const MEET_API_URL = process.env.MEET_API_URL ?? "https://meet.nvidia.com/api/v1";

// =============================================================================
// Actions & Schema
// =============================================================================

const AVAILABILITY_ACTIONS = ["find_meeting_times"] as const;

const OutlookAvailabilitySchema = Type.Object({
  action: stringEnum(AVAILABILITY_ACTIONS, {
    description: "Action to perform: find_meeting_times (find available times for attendees).",
  }),
  attendees: Type.Union([Type.Array(Type.String()), Type.String()], {
    description:
      'Attendees: email addresses or names to search. Array or comma-separated string. Example: ["alice@nvidia.com", "Bob Smith"]',
  }),
  stakeholder: Type.Optional(
    Type.String({
      description:
        "Key stakeholder name or email — schedule and room selection will be prioritized for this person.",
    }),
  ),
  start_time: Type.Optional(
    Type.String({ description: "Start of search window (ISO 8601). Defaults to now." }),
  ),
  end_time: Type.Optional(
    Type.String({ description: "End of search window (ISO 8601). Defaults to 7 days from start." }),
  ),
  meeting_duration: Type.Optional(
    Type.Number({ description: "Duration in minutes (default: 30, min: 5, max: 1440).", minimum: 5, maximum: 1440 }),
  ),
  is_organizer_optional: Type.Optional(
    Type.Boolean({ description: "Whether the organizer is optional (default: false)." }),
  ),
  subject: Type.Optional(
    Type.String({ description: "Meeting subject for the deeplink (default: 'Meeting')." }),
  ),
  body: Type.Optional(
    Type.String({ description: "Meeting body/description for the deeplink." }),
  ),
});

// =============================================================================
// Graph API Helpers
// =============================================================================

async function graphGet(
  token: string,
  path: string,
  headers?: Record<string, string>,
): Promise<any> {
  const url = `${GRAPH_BASE_URL}${path}`;
  const res = await fetch(url, {
    method: "GET",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...headers },
    signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API ${res.status}: ${text || res.statusText}`);
  }
  return res.json();
}

async function graphPost(
  token: string,
  path: string,
  body: object,
  headers?: Record<string, string>,
): Promise<any> {
  const url = `${GRAPH_BASE_URL}${path}`;
  const res = await fetch(url, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...headers },
    body: JSON.stringify(body),
    signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API ${res.status}: ${text || res.statusText}`);
  }
  return res.json();
}

// =============================================================================
// Timezone Helpers
// =============================================================================

const WINDOWS_TO_IANA: Record<string, string> = {
  "Pacific Standard Time": "America/Los_Angeles",
  "Mountain Standard Time": "America/Denver",
  "Central Standard Time": "America/Chicago",
  "Eastern Standard Time": "America/New_York",
  "Atlantic Standard Time": "America/Halifax",
  "Alaskan Standard Time": "America/Anchorage",
  "Hawaiian Standard Time": "Pacific/Honolulu",
  "GMT Standard Time": "Europe/London",
  "Central European Standard Time": "Europe/Paris",
  "Eastern European Standard Time": "Europe/Bucharest",
  "Russian Standard Time": "Europe/Moscow",
  "China Standard Time": "Asia/Shanghai",
  "Japan Standard Time": "Asia/Tokyo",
  "India Standard Time": "Asia/Kolkata",
  "Australian Eastern Standard Time": "Australia/Sydney",
  "AUS Eastern Standard Time": "Australia/Sydney",
  "New Zealand Standard Time": "Pacific/Auckland",
  UTC: "UTC",
};

function convertWindowsToIANA(windowsTz: string): string {
  if (WINDOWS_TO_IANA[windowsTz]) return WINDOWS_TO_IANA[windowsTz];
  const lower = windowsTz.toLowerCase();
  for (const [key, value] of Object.entries(WINDOWS_TO_IANA)) {
    if (key.toLowerCase() === lower) return value;
  }
  if (lower.includes("pacific")) return "America/Los_Angeles";
  if (lower.includes("eastern")) return "America/New_York";
  if (lower.includes("central")) return "America/Chicago";
  return "UTC";
}

function toISOInTimezone(d: Date, timeZone: string): string {
  const ianaTimezone = timeZone.includes("/") ? timeZone : convertWindowsToIANA(timeZone);
  const formatter = new Intl.DateTimeFormat("en-CA", {
    timeZone: ianaTimezone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  });
  const parts = formatter.formatToParts(d);
  const get = (type: string) => parts.find((p) => p.type === type)!.value;
  return `${get("year")}-${get("month")}-${get("day")}T${get("hour")}:${get("minute")}:${get("second")}`;
}

// =============================================================================
// Person Resolution
// =============================================================================

interface PersonCandidate {
  email: string;
  displayName: string;
  jobTitle?: string;
  department?: string;
}

interface AmbiguousAttendee {
  query: string;
  candidates: PersonCandidate[];
}

function extractCandidateFromPerson(person: any): PersonCandidate | null {
  let email: string | null = null;
  if (person.scoredEmailAddresses?.length > 0) email = person.scoredEmailAddresses[0].address || null;
  else if (person.emailAddresses?.length > 0) email = person.emailAddresses[0].address || null;
  if (!email) return null;
  return { email, displayName: person.displayName || email, jobTitle: person.jobTitle, department: person.department };
}

function extractCandidateFromUser(user: any): PersonCandidate | null {
  const email = user.mail || user.userPrincipalName || null;
  if (!email) return null;
  return { email, displayName: user.displayName || email, jobTitle: user.jobTitle, department: user.department };
}

async function searchUserCandidates(query: string, graphToken: string, maxCandidates = 5): Promise<PersonCandidate[]> {
  const trimmed = query.trim();

  // Exact email — resolve directly
  if (trimmed.includes("@")) {
    try {
      const user = await graphGet(graphToken, `/users/${encodeURIComponent(trimmed.toLowerCase())}`);
      const c = extractCandidateFromUser(user);
      return c ? [c] : [];
    } catch {
      // fall through
    }
  }

  // Primary: /me/people $search
  try {
    const searchQuery = trimmed.replace(/"/g, '\\"');
    const url = `/me/people?$search="${encodeURIComponent(searchQuery)}"&$top=${maxCandidates}`;
    const res = await graphGet(graphToken, url, { ConsistencyLevel: "eventual" });
    const candidates = ((res.value || []) as any[]).map(extractCandidateFromPerson).filter(Boolean) as PersonCandidate[];
    if (candidates.length > 0) return candidates;
  } catch {
    // fall through
  }

  // Fallback: /me/people $filter
  try {
    const escaped = trimmed.replace(/'/g, "''");
    const filter = `contains(displayName,'${escaped}') or contains(givenName,'${escaped}') or contains(surname,'${escaped}')`;
    const res = await graphGet(graphToken, `/me/people?$filter=${encodeURIComponent(filter)}&$top=${maxCandidates}`);
    const candidates = ((res.value || []) as any[]).map(extractCandidateFromPerson).filter(Boolean) as PersonCandidate[];
    if (candidates.length > 0) return candidates;
  } catch {
    // fall through
  }

  // Fallback: /users directory
  try {
    const escaped = trimmed.replace(/'/g, "''");
    const filter = `contains(displayName,'${escaped}') or contains(givenName,'${escaped}') or contains(surname,'${escaped}') or contains(mail,'${escaped}')`;
    const res = await graphGet(graphToken, `/users?$filter=${encodeURIComponent(filter)}&$top=${maxCandidates}`);
    const candidates = ((res.value || []) as any[]).map(extractCandidateFromUser).filter(Boolean) as PersonCandidate[];
    if (candidates.length > 0) return candidates;
  } catch {
    // fall through
  }

  return [];
}

// =============================================================================
// Meeting Room Intelligence
// =============================================================================

async function fetchAttendeeMeetingRooms(
  attendeeEmail: string,
  graphToken: string,
  lookbackDays = 90,
  isCurrentUser = false,
): Promise<any[]> {
  try {
    const now = new Date();
    const startStr = new Date(now.getTime() - lookbackDays * 24 * 60 * 60 * 1000).toISOString();
    const endStr = now.toISOString();

    const prefix = isCurrentUser ? "/me" : `/users/${encodeURIComponent(attendeeEmail)}`;
    const url = `${prefix}/calendar/calendarView?startDateTime=${encodeURIComponent(startStr)}&endDateTime=${encodeURIComponent(endStr)}&$top=500`;
    const events = await graphGet(graphToken, url);

    const rooms: any[] = [];
    for (const event of events.value || []) {
      for (const loc of event.locations || []) {
        if (loc.uniqueIdType === "private" || loc.locationType === "default") continue;
        const name = (loc.displayName || "").toLowerCase();
        if (name.includes("microsoft teams") || name === "microsoft teams meeting") continue;
        if (
          loc.locationType === "conferenceRoom" ||
          name.includes("room") ||
          name.includes("conference") ||
          name.includes("meeting")
        ) {
          rooms.push({
            displayName: loc.displayName,
            address: loc.address,
            locationType: loc.locationType,
            uniqueId: loc.uniqueId,
            uniqueIdType: loc.uniqueIdType,
          });
        }
      }
    }
    return rooms;
  } catch {
    return [];
  }
}

function extractBuildingName(displayName?: string, uniqueId?: string): string | null {
  const id = displayName || uniqueId || "";
  if (!id) return null;
  const match = id.match(/^([A-Za-z0-9-]+?)-(\d+)[-\s]/);
  if (match?.[1]) return match[1];
  const fallback = id.match(/^([A-Za-z-]+?)(?=-\d)/);
  if (fallback?.[1]) return fallback[1];
  return null;
}

async function findMostFrequentMeetingRooms(
  attendees: string[],
  graphToken: string,
  currentUserEmail: string | null,
  lookbackDays = 90,
  topN = 5,
): Promise<any[]> {
  const promises = attendees.map((email) => fetchAttendeeMeetingRooms(email, graphToken, lookbackDays, false));
  if (currentUserEmail) {
    promises.push(fetchAttendeeMeetingRooms(currentUserEmail, graphToken, lookbackDays, true));
  }

  const results = await Promise.allSettled(promises);
  const counts = new Map<string, { buildingName: string; count: number; sampleRooms: any[] }>();

  for (const result of results) {
    if (result.status !== "fulfilled") continue;
    for (const room of result.value) {
      const building = extractBuildingName(room.displayName, room.uniqueId);
      if (!building) continue;
      const entry = counts.get(building);
      if (entry) {
        entry.count++;
        if (entry.sampleRooms.length < 3 && !entry.sampleRooms.some((r: any) => r.uniqueId === room.uniqueId)) {
          entry.sampleRooms.push(room);
        }
      } else {
        counts.set(building, { buildingName: building, count: 1, sampleRooms: [room] });
      }
    }
  }

  return Array.from(counts.values())
    .sort((a, b) => b.count - a.count)
    .slice(0, topN)
    .map((item) => ({ buildingName: item.buildingName, frequency: item.count, sampleRooms: item.sampleRooms }));
}

async function fetchAllRoomsFromAPI(): Promise<any[]> {
  try {
    const res = await fetch(`${MEET_API_URL}/room`, {
      method: "GET",
      headers: { "Content-Type": "application/json" },
      signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
    });
    if (!res.ok) return [];
    const data = await res.json();
    return Array.isArray(data) ? data : [];
  } catch {
    return [];
  }
}

function filterRoomsByBuilding(rooms: any[], buildingNames: string[]): any[] {
  if (!buildingNames.length || !rooms.length) return [];

  const normalized = buildingNames.map((b) => {
    const n = b.toLowerCase().trim();
    return { original: n, noDashes: n.replace(/-/g, "") };
  });

  return rooms.filter((room) => {
    const name = (room.name || "").toLowerCase();
    const address = (room.address || room.email || "").toLowerCase();
    const email = (room.email || "").toLowerCase();
    const location = (room.location || "").toLowerCase();
    const all = `${name} ${address} ${location}`;

    return normalized.some(({ original, noDashes }) => {
      const pattern = original.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const regex = new RegExp(`(^|[-_])${pattern}([-_]|@|$)`, "i");
      return (
        name.startsWith(original) ||
        regex.test(name) ||
        regex.test(address) ||
        regex.test(email) ||
        location.includes(original) ||
        all.includes(original) ||
        (noDashes.length > 3 && all.includes(noDashes))
      );
    });
  });
}

async function checkRoomsAvailability(
  rooms: any[],
  startTime: string,
  endTime: string,
  graphToken: string,
): Promise<Array<{ room: any; available: boolean; error?: string }>> {
  if (!rooms.length) return [];

  const roomEmails = rooms.map((r) => r.email || r.address).filter(Boolean) as string[];
  if (!roomEmails.length) return rooms.map((room) => ({ room, available: false, error: "No email" }));

  const BATCH_SIZE = 2;
  const startISO = new Date(startTime).toISOString();
  const endISO = new Date(endTime).toISOString();
  const results: Array<{ room: any; available: boolean; error?: string }> = [];

  for (let i = 0; i < roomEmails.length; i += BATCH_SIZE) {
    const batchEmails = roomEmails.slice(i, i + BATCH_SIZE);
    const batchRooms = rooms.slice(i, i + BATCH_SIZE);

    if (i > 0) await new Promise((r) => setTimeout(r, 1000));

    try {
      const response = await graphPost(graphToken, "/me/calendar/getSchedule", {
        schedules: batchEmails,
        startTime: { dateTime: startISO, timeZone: "UTC" },
        endTime: { dateTime: endISO, timeZone: "UTC" },
        availabilityViewInterval: 30,
      });

      const schedules = response.value || [];
      const scheduleMap = new Map<string, any>();
      schedules.forEach((s: any) => scheduleMap.set((s.scheduleId || "").toLowerCase(), s));

      for (let j = 0; j < batchRooms.length; j++) {
        const room = batchRooms[j];
        const email = (room.email || room.address || "").toLowerCase();
        const schedule = scheduleMap.get(email);
        if (!schedule) {
          results.push({ room, available: false, error: "Schedule not found" });
          continue;
        }
        const view = schedule.availabilityView || "";
        const available = !view || view.split("").every((s: string) => s === "0");
        results.push({ room, available });
      }
    } catch {
      // Fall back to marking batch as unavailable
      for (const room of batchRooms) {
        results.push({ room, available: false, error: "Check failed" });
      }
    }
  }

  return results;
}

// =============================================================================
// Deeplink Generation
// =============================================================================

function plainTextToHtml(text: string): string {
  let html = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
  html = html.replace(/\*\*(.+?)\*\*/g, "<b>$1</b>");
  html = html.replace(/^#{1,3}\s+(.+)$/gm, "<b>$1</b>");
  html = html.replace(/\[([^\]]+)\]\((https?:\/\/[^)]+)\)/g, '<a href="$2">$1</a>');
  html = html.replace(/(^|[\s>])(https?:\/\/\S+)/g, '$1<a href="$2">$2</a>');
  html = html.replace(/^- (.+)$/gm, "&bull; $1");
  html = html.replace(/\n/g, "<br>");
  return html;
}

function generateTeamsMeetingDeeplink(
  startTime: string,
  endTime: string,
  attendees: string[],
  subject?: string,
  durationMinutes?: number,
  body?: string,
): string {
  const startISO = new Date(startTime).toISOString().replace(/\.\d{3}Z$/, "Z");
  const endISO = new Date(endTime).toISOString().replace(/\.\d{3}Z$/, "Z");
  const params = new URLSearchParams();
  params.append("subject", subject || "Meeting");
  params.append("startdt", startISO);
  params.append("enddt", endISO);
  params.append("allday", "false");
  params.append("body", plainTextToHtml(body || `Meeting scheduled for ${durationMinutes || 30} minutes.`));
  if (attendees.length > 0) params.append("to", attendees.join(","));
  return `https://outlook.office.com/calendar/deeplink/compose?${params.toString()}`;
}

// =============================================================================
// Main Handler
// =============================================================================

async function handleFindMeetingTimes(
  graphToken: string,
  params: Record<string, unknown>,
): Promise<unknown> {
  // Parse attendees
  let attendeeQueries: string[] = [];
  const raw = params.attendees;
  if (Array.isArray(raw)) {
    attendeeQueries = raw.filter((a) => typeof a === "string" && a.trim().length > 0) as string[];
  } else if (typeof raw === "string") {
    attendeeQueries = raw.split(",").map((s) => s.trim()).filter(Boolean);
  }
  if (!attendeeQueries.length) {
    return { error: "At least one attendee is required." };
  }

  // Resolve attendees
  const attendees: string[] = [];
  const unresolved: string[] = [];
  const ambiguous: AmbiguousAttendee[] = [];

  for (const query of attendeeQueries) {
    if (query.includes("@")) {
      attendees.push(query.trim().toLowerCase());
    } else {
      const candidates = await searchUserCandidates(query, graphToken);
      if (candidates.length === 0) unresolved.push(query);
      else if (candidates.length === 1) attendees.push(candidates[0].email);
      else ambiguous.push({ query, candidates });
    }
  }

  // Return ambiguous matches for LLM disambiguation
  if (ambiguous.length > 0) {
    const lines = ambiguous.map(({ query, candidates }) => {
      const options = candidates
        .map((c, i) => `  ${i + 1}. ${c.displayName} <${c.email}>${c.jobTitle ? ` — ${c.jobTitle}` : ""}${c.department ? ` (${c.department})` : ""}`)
        .join("\n");
      return `"${query}" matched ${candidates.length} people:\n${options}`;
    });
    return {
      ambiguousAttendees: ambiguous,
      error: `Some attendee names matched multiple people. Please confirm which person you meant and re-call with their email address.\n\n${lines.join("\n\n")}`,
    };
  }

  if (!attendees.length) {
    return { error: unresolved.length ? `Could not find: ${unresolved.join(", ")}` : "No valid attendees." };
  }

  // Resolve stakeholder
  let stakeholderEmail: string | null = null;
  const stakeholderRaw = readStringParam(params, "stakeholder");
  if (stakeholderRaw) {
    if (stakeholderRaw.includes("@")) {
      stakeholderEmail = stakeholderRaw.toLowerCase();
    } else {
      const candidates = await searchUserCandidates(stakeholderRaw, graphToken);
      if (candidates.length === 1) stakeholderEmail = candidates[0].email;
      else if (candidates.length > 1) {
        const options = candidates.map((c, i) => `  ${i + 1}. ${c.displayName} <${c.email}>${c.jobTitle ? ` — ${c.jobTitle}` : ""}`).join("\n");
        return {
          ambiguousAttendees: [{ query: stakeholderRaw, candidates }],
          error: `Stakeholder "${stakeholderRaw}" matched multiple people:\n${options}`,
        };
      }
    }
    if (stakeholderEmail && !attendees.some((a) => a.toLowerCase() === stakeholderEmail!.toLowerCase())) {
      attendees.push(stakeholderEmail);
    }
  }

  // Timezone
  let windowsTimezone = "UTC";
  try {
    const settings = await graphGet(graphToken, "/me/mailboxSettings");
    windowsTimezone = settings.timeZone || "UTC";
  } catch {
    // default UTC
  }

  // Time window
  const now = new Date();
  const start = new Date(String(params.start_time || now.toISOString()));
  const end = new Date(String(params.end_time || new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000).toISOString()));
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return { error: "Invalid date format." };
  if (start >= end) return { error: "start_time must be before end_time." };

  const meetingDuration = Math.min(Math.max(readNumberParam(params, "meeting_duration", { integer: true }) ?? 30, 5), 1440);
  const subject = readStringParam(params, "subject");
  const body = readStringParam(params, "body");

  // Get current user email
  let currentUserEmail: string | null = null;
  try {
    const me = await graphGet(graphToken, "/me");
    currentUserEmail = me.mail || me.userPrincipalName || null;
  } catch {
    // continue
  }

  // Room intelligence (parallel)
  const roomSearchAttendees = stakeholderEmail ? [stakeholderEmail] : attendees;
  const frequentRoomsPromise = findMostFrequentMeetingRooms(roomSearchAttendees, graphToken, currentUserEmail, 90, 5);

  // Find meeting times
  const requestBody: any = {
    attendees: attendees.map((email) => ({ emailAddress: { address: email }, type: "required" })),
    timeConstraint: {
      activityDomain: "work",
      timeslots: [
        {
          start: { dateTime: toISOInTimezone(start, windowsTimezone), timeZone: windowsTimezone },
          end: { dateTime: toISOInTimezone(end, windowsTimezone), timeZone: windowsTimezone },
        },
      ],
    },
    meetingDuration: `PT${meetingDuration}M`,
    isOrganizerOptional: params.is_organizer_optional ?? false,
    returnSuggestionReasons: true,
    maxCandidates: 50,
  };

  const response = await graphPost(graphToken, "/me/findMeetingTimes", requestBody, {
    Prefer: `outlook.timezone="${windowsTimezone}"`,
  });

  const suggestions = response.meetingTimeSuggestions || [];
  const ianaTimezone = convertWindowsToIANA(windowsTimezone);

  // Map and filter to business hours
  let items = suggestions
    .map((s: any) => {
      const slot = s.meetingTimeSlot;
      const startDT = slot?.start?.dateTime;
      const endDT = slot?.end?.dateTime;
      const deeplink =
        startDT && endDT ? generateTeamsMeetingDeeplink(startDT, endDT, attendees, subject, meetingDuration, body) : undefined;
      return {
        start: startDT,
        end: endDT,
        confidence: s.confidence,
        score: s.score,
        attendeeAvailability: s.attendeeAvailability?.map((a: any) => ({
          attendee: a.attendee?.emailAddress?.address,
          availability: a.availability,
        })),
        locations: s.locations || [],
        suggestionReason: s.suggestionReason,
        teamsDeeplink: deeplink,
      };
    })
    .filter((item: any) => {
      if (!item.start) return false;
      const formatter = new Intl.DateTimeFormat("en-US", { timeZone: ianaTimezone, hour: "numeric", hour12: false });
      const hour = parseInt(formatter.format(new Date(item.start)), 10);
      return hour >= 8 && hour < 18;
    });

  // Stakeholder priority sort
  if (stakeholderEmail && items.length > 1) {
    const priority: Record<string, number> = { free: 0, tentative: 1, workingElsewhere: 2, unknown: 3, busy: 4, oof: 5 };
    items.sort((a: any, b: any) => {
      const aAvail = (a.attendeeAvailability || []).find((av: any) => av.attendee?.toLowerCase() === stakeholderEmail!.toLowerCase());
      const bAvail = (b.attendeeAvailability || []).find((av: any) => av.attendee?.toLowerCase() === stakeholderEmail!.toLowerCase());
      const ap = priority[aAvail?.availability || "unknown"] ?? 3;
      const bp = priority[bAvail?.availability || "unknown"] ?? 3;
      if (ap !== bp) return ap - bp;
      return (b.confidence || 0) - (a.confidence || 0);
    });
  }

  // Room availability
  let frequentMeetingRooms: any[] = [];
  try {
    frequentMeetingRooms = await frequentRoomsPromise;
  } catch {
    // continue without rooms
  }

  let roomsInBuilding: any[] = [];
  if (frequentMeetingRooms.length > 0) {
    const building = frequentMeetingRooms[0]?.buildingName;
    if (building) {
      const allRooms = await fetchAllRoomsFromAPI();
      roomsInBuilding = filterRoomsByBuilding(allRooms, [building]);
    }
  }

  // Check room availability for top 5 time slots
  if (roomsInBuilding.length > 0 && items.length > 0) {
    const maxSlots = Math.min(items.length, 5);
    for (let i = 0; i < maxSlots; i++) {
      const slot = items[i];
      if (!slot?.start || !slot?.end) continue;
      try {
        if (i > 0) await new Promise((r) => setTimeout(r, 1000));
        const availability = await checkRoomsAvailability(roomsInBuilding, slot.start, slot.end, graphToken);
        const available = availability
          .filter((ra) => ra.available)
          .map((ra) => ({ name: ra.room.name, email: ra.room.email || ra.room.address, location: ra.room.location }))
          .filter((r) => r.email?.includes("@"));

        slot.availableRooms = available;

        // Update deeplink with first available room
        if (available.length > 0 && available[0].email) {
          slot.teamsDeeplink = generateTeamsMeetingDeeplink(
            slot.start,
            slot.end,
            [...attendees, available[0].email],
            subject,
            meetingDuration,
            body,
          );
        }
      } catch {
        // continue without room info for this slot
      }
    }
  }

  return {
    items,
    frequentMeetingRooms,
    roomsInFrequentBuildings: roomsInBuilding,
    ...(unresolved.length ? { unresolvedAttendees: unresolved } : {}),
    ...(stakeholderEmail
      ? { stakeholder: stakeholderEmail, stakeholderNote: "Results sorted by stakeholder availability; rooms based on stakeholder's frequent locations." }
      : {}),
  };
}

// =============================================================================
// Tool Factory
// =============================================================================

export function createOutlookAvailabilityTool(options: {
  getRefreshToken: () => string | null | Promise<string | null>;
  enabled?: boolean;
}): AnyAgentTool | null {
  if (options.enabled === false) return null;
  if (!isAzureOBOConfigured()) return null;

  return {
    label: "Outlook Availability",
    name: "outlook_availability",
    description:
      "Find available meeting times for multiple attendees using Microsoft Graph. Resolves names to emails, finds common free slots, suggests meeting rooms based on attendee history, and generates Outlook calendar deeplinks. Requires NVIDIA SSO authentication.",
    parameters: OutlookAvailabilitySchema,
    execute: async (_toolCallId, args) => {
      const params = args as Record<string, unknown>;
      const action = readStringParam(params, "action", { required: true });

      try {
        const refreshToken = await options.getRefreshToken();
        if (!refreshToken) {
          return jsonResult({ error: "not_authenticated", message: "Please log in first." });
        }

        const tokenResult = await acquireDownstreamToken(refreshToken, AVAILABILITY_SCOPES);
        if (!tokenResult.ok) {
          return jsonResult({
            error: "token_exchange_failed",
            message: `Failed to acquire Graph token: ${"error" in tokenResult ? tokenResult.error : "unknown"}`,
          });
        }

        let result: unknown;
        switch (action) {
          case "find_meeting_times":
            result = await handleFindMeetingTimes(tokenResult.accessToken, params);
            break;
          default:
            return jsonResult({ error: "invalid_action", message: `Unknown action: ${action}` });
        }

        return jsonResult({ action, result });
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        return jsonResult({ error: "outlook_availability_error", message });
      }
    },
  } as AnyAgentTool;
}
