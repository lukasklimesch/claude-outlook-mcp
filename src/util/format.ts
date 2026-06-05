// ====================================================
// Human-readable formatting of results for MCP responses
// ====================================================

import type {
  AttachmentInfo,
  CalendarEvent,
  ContactInfo,
  MailDetail,
  MailFolderInfo,
  MailSummary,
} from "../types.js";

function clip(s: string | undefined, n: number): string {
  if (!s) return "";
  const flat = String(s).replace(/\r\n/g, "\n");
  return flat.length > n ? flat.slice(0, n) + "…" : flat;
}

export function formatMailList(emails: MailSummary[], heading: string): string {
  if (!emails || emails.length === 0) return `No emails found ${heading}`.trim();
  const lines = emails.map((e, i) => {
    const flags = [
      e.unread ? "UNREAD" : null,
      e.hasAttachments ? `📎${e.attachmentCount}` : null,
    ]
      .filter(Boolean)
      .join(" ");
    return (
      `${i + 1}. [${e.received}] ${flags}\n` +
      `   From: ${e.senderName}${e.senderEmail ? ` <${e.senderEmail}>` : ""}\n` +
      `   Subject: ${e.subject || "(no subject)"}\n` +
      `   EntryID: ${e.entryId}\n` +
      (e.preview ? `   ${clip(e.preview, 200)}\n` : "")
    );
  });
  return `Found ${emails.length} email(s) ${heading}\n\n${lines.join("\n")}`.trim();
}

export function formatMailDetail(e: MailDetail): string {
  const parts = [
    `Subject: ${e.subject || "(no subject)"}`,
    `From: ${e.senderName}${e.senderEmail ? ` <${e.senderEmail}>` : ""}`,
    e.to ? `To: ${e.to}` : null,
    e.cc ? `Cc: ${e.cc}` : null,
    `Received: ${e.received}`,
    `Unread: ${e.unread ? "yes" : "no"}`,
    `EntryID: ${e.entryId}`,
  ].filter(Boolean);

  let out = parts.join("\n");
  if (e.attachments && e.attachments.length > 0) {
    out +=
      `\n\nAttachments (${e.attachments.length}):\n` +
      e.attachments
        .map((a) => `  [${a.index}] ${a.fileName} (${formatBytes(a.size)})`)
        .join("\n");
  }
  out += `\n\n--- Body ---\n${e.body || "(empty)"}`;
  return out;
}

export function formatEvents(events: CalendarEvent[], heading: string): string {
  if (!events || events.length === 0) return `No events found ${heading}`.trim();
  const lines = events.map((ev, i) => {
    return (
      `${i + 1}. ${ev.subject || "(no title)"}\n` +
      `   When: ${ev.start} → ${ev.end}${ev.isAllDay ? " (all day)" : ""}\n` +
      `   Where: ${ev.location || "(no location)"}\n` +
      (ev.organizer ? `   Organizer: ${ev.organizer}\n` : "") +
      `   EntryID: ${ev.entryId}\n`
    );
  });
  return `Found ${events.length} event(s) ${heading}\n\n${lines.join("\n")}`.trim();
}

export function formatContacts(contacts: ContactInfo[], heading: string): string {
  if (!contacts || contacts.length === 0)
    return `No contacts found ${heading}`.trim();
  const lines = contacts.map((c, i) => {
    return (
      `${i + 1}. ${c.fullName}\n` +
      `   Email: ${c.email || "(none)"}\n` +
      `   Phone: ${c.phone || "(none)"}\n` +
      (c.company ? `   Company: ${c.company}\n` : "") +
      (c.jobTitle ? `   Title: ${c.jobTitle}\n` : "")
    );
  });
  return `Found ${contacts.length} contact(s) ${heading}\n\n${lines.join("\n")}`.trim();
}

export function formatFolders(folders: MailFolderInfo[]): string {
  if (!folders || folders.length === 0) return "No mail folders found.";
  const lines = folders.map((f) => {
    const counts =
      f.unread !== undefined || f.total !== undefined
        ? ` (${f.unread ?? "?"} unread / ${f.total ?? "?"} total)`
        : "";
    return `  ${f.path || f.name}${counts}`;
  });
  return `Found ${folders.length} folder(s):\n${lines.join("\n")}`;
}

export function formatAttachments(list: AttachmentInfo[]): string {
  if (!list || list.length === 0) return "This message has no attachments.";
  const lines = list.map(
    (a) => `  [${a.index}] ${a.fileName} (${formatBytes(a.size)})`,
  );
  return `Found ${list.length} attachment(s):\n${lines.join("\n")}`;
}

export function formatSavedFiles(paths: string[]): string {
  if (!paths || paths.length === 0) return "No attachments were saved.";
  return `Saved ${paths.length} file(s):\n${paths.map((p) => `  ${p}`).join("\n")}`;
}

export function formatBytes(bytes: number | undefined): string {
  if (bytes === undefined || bytes === null || Number.isNaN(bytes))
    return "unknown size";
  const units = ["B", "KB", "MB", "GB"];
  let n = bytes;
  let u = 0;
  while (n >= 1024 && u < units.length - 1) {
    n /= 1024;
    u++;
  }
  return `${u === 0 ? n : n.toFixed(1)} ${units[u]}`;
}
