// ====================================================
// Calendar operations (Outlook COM via PowerShell)
// ====================================================

import type { CalendarEvent, OutlookRunner } from "../types.js";
import { buildScript } from "../ps/preamble.js";
import { asArray, expectOk } from "./common.js";
import type { CalendarArgs } from "../util/validate.js";
import {
  addDays,
  normalizeLocalDateTime,
  startOfDay,
  toRestrictFormat,
} from "../util/dates.js";

const DEFAULT_LIMIT = 10;
const DEFAULT_DAYS = 7;

// Window query over the calendar (with recurrence expansion).
const BODY_WINDOW = String.raw`
$cal = $ns.GetDefaultFolder($olFolderCalendar)
$limit = if ($payload.limit) { [int]$payload.limit } else { 10 }
$items = $cal.Items
# Outlook requires sorting by [Start] BEFORE enabling IncludeRecurrences;
# the reverse order throws and/or fails to expand recurring occurrences.
$items.Sort('[Start]')
$items.IncludeRecurrences = $true
$restricted = $items.Restrict([string]$payload.filter)
$result = @(); $count = 0
foreach ($e in $restricted) {
  try { $result += (Convert-Event $e $false); $count++ } catch {}
  if ($count -ge $limit) { break }
}
Write-McpResult -Obj @($result)
`;

const BODY_SEARCH = String.raw`
$cal = $ns.GetDefaultFolder($olFolderCalendar)
$limit = if ($payload.limit) { [int]$payload.limit } else { 10 }
$term = ([string]$payload.searchTerm).ToLower()
$items = $cal.Items
$items.Sort('[Start]')
$result = @(); $count = 0
foreach ($e in $items) {
  try {
    $hay = (([string]$e.Subject) + ' ' + ([string]$e.Location)).ToLower()
    if ($hay.Contains($term)) { $result += (Convert-Event $e $false); $count++ }
  } catch {}
  if ($count -ge $limit) { break }
}
Write-McpResult -Obj @($result)
`;

const BODY_CREATE = String.raw`
$evt = $app.CreateItem($olAppointmentItem)
$evt.Subject = [string]$payload.subject
$evt.Start = [datetime]::Parse($payload.start, [Globalization.CultureInfo]::InvariantCulture)
$evt.End = [datetime]::Parse($payload.end, [Globalization.CultureInfo]::InvariantCulture)
if ($payload.location) { $evt.Location = [string]$payload.location }
if ($payload.body) { $evt.Body = [string]$payload.body }
if ($payload.attendees) {
  $evt.MeetingStatus = $olMeeting
  foreach ($a in ([string]$payload.attendees -split ',')) {
    $addr = $a.Trim()
    if ($addr) { [void]$evt.Recipients.Add($addr) }
  }
  if (-not $evt.Recipients.ResolveAll()) {
    $bad = @()
    foreach ($r in $evt.Recipients) { if (-not $r.Resolved) { $bad += [string]$r.Name } }
    throw "Could not resolve attendee(s): $($bad -join ', ')"
  }
  $evt.Send()
  Write-McpResult -Obj @{ status = 'invited'; subject = [string]$evt.Subject }
} else {
  $evt.Save()
  Write-McpResult -Obj @{ status = 'created'; entryId = [string]$evt.EntryID; subject = [string]$evt.Subject }
}
`;

const BODY_UPDATE = String.raw`
$evt = $ns.GetItemFromID($payload.entryId)
if ($payload.subject) { $evt.Subject = [string]$payload.subject }
if ($payload.start) { $evt.Start = [datetime]::Parse($payload.start, [Globalization.CultureInfo]::InvariantCulture) }
if ($payload.end) { $evt.End = [datetime]::Parse($payload.end, [Globalization.CultureInfo]::InvariantCulture) }
if ($payload.location) { $evt.Location = [string]$payload.location }
if ($payload.body) { $evt.Body = [string]$payload.body }
$evt.Save()
Write-McpResult -Obj @{ status = 'updated'; entryId = [string]$evt.EntryID }
`;

const BODY_DELETE = String.raw`
$evt = $ns.GetItemFromID($payload.entryId)
$evt.Delete()
Write-McpResult -Obj @{ status = 'deleted' }
`;

function restrictFilter(start: Date, end: Date): string {
  return `[Start] >= '${toRestrictFormat(start)}' AND [Start] < '${toRestrictFormat(end)}'`;
}

export async function todayEvents(runner: OutlookRunner, args: CalendarArgs): Promise<CalendarEvent[]> {
  const now = new Date();
  const start = startOfDay(now);
  const end = addDays(start, 1);
  const res = await runner.run(buildScript(BODY_WINDOW), {
    filter: restrictFilter(start, end),
    limit: args.limit ?? DEFAULT_LIMIT,
  });
  return asArray<CalendarEvent>(expectOk(res));
}

export async function upcomingEvents(runner: OutlookRunner, args: CalendarArgs): Promise<CalendarEvent[]> {
  const days = args.days ?? DEFAULT_DAYS;
  const now = new Date();
  const end = addDays(startOfDay(now), days);
  const res = await runner.run(buildScript(BODY_WINDOW), {
    filter: restrictFilter(now, end),
    limit: args.limit ?? DEFAULT_LIMIT,
  });
  return asArray<CalendarEvent>(expectOk(res));
}

export async function searchEvents(runner: OutlookRunner, args: CalendarArgs): Promise<CalendarEvent[]> {
  const res = await runner.run(buildScript(BODY_SEARCH), {
    searchTerm: args.searchTerm,
    limit: args.limit ?? DEFAULT_LIMIT,
  });
  return asArray<CalendarEvent>(expectOk(res));
}

export async function createEvent(runner: OutlookRunner, args: CalendarArgs): Promise<unknown> {
  const res = await runner.run(buildScript(BODY_CREATE), {
    subject: args.subject,
    start: normalizeLocalDateTime(args.start!),
    end: normalizeLocalDateTime(args.end!),
    location: args.location,
    body: args.body,
    attendees: args.attendees,
  });
  return expectOk(res);
}

export async function updateEvent(runner: OutlookRunner, args: CalendarArgs): Promise<unknown> {
  const res = await runner.run(buildScript(BODY_UPDATE), {
    entryId: args.entryId,
    subject: args.subject,
    start: args.start ? normalizeLocalDateTime(args.start) : undefined,
    end: args.end ? normalizeLocalDateTime(args.end) : undefined,
    location: args.location,
    body: args.body,
  });
  return expectOk(res);
}

export async function deleteEvent(runner: OutlookRunner, args: CalendarArgs): Promise<unknown> {
  const res = await runner.run(buildScript(BODY_DELETE), { entryId: args.entryId });
  return expectOk(res);
}
