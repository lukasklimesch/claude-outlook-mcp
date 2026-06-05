// ====================================================
// MCP tool schema definitions (Windows Outlook)
// ====================================================

import type { Tool } from "@modelcontextprotocol/sdk/types.js";

export const OUTLOOK_MAIL_TOOL: Tool = {
  name: "outlook_mail",
  description:
    "Read, search, send and manage email in the Microsoft Outlook desktop app on Windows (via the Outlook COM object model). Operations: unread, read, search, get (full message by EntryID), send, draft, reply, forward, move, delete, mark (read/unread), flag, folders.",
  inputSchema: {
    type: "object",
    properties: {
      operation: {
        type: "string",
        description: "The mail operation to perform.",
        enum: [
          "unread",
          "read",
          "search",
          "get",
          "send",
          "draft",
          "reply",
          "forward",
          "move",
          "delete",
          "mark",
          "flag",
          "folders",
        ],
      },
      folder: {
        type: "string",
        description:
          "Folder name for unread/read/search (e.g. 'Inbox', 'Sent Items', 'Drafts', or any custom folder). Defaults to Inbox.",
      },
      limit: { type: "number", description: "Max items to return (default 10)." },
      searchTerm: {
        type: "string",
        description: "Text to match in subject/sender/body (required for search).",
      },
      entryId: {
        type: "string",
        description:
          "Outlook EntryID of a message (required for get/reply/forward/move/delete/mark/flag). Returned by unread/read/search.",
      },
      to: { type: "string", description: "Recipient address(es), semicolon-separated (send/forward)." },
      cc: { type: "string", description: "CC address(es) (optional)." },
      bcc: { type: "string", description: "BCC address(es) (optional)." },
      subject: { type: "string", description: "Subject line (send/draft)." },
      body: { type: "string", description: "Message body (send/draft/reply/forward)." },
      isHtml: {
        type: "boolean",
        description: "Treat body as HTML (default false).",
      },
      attachments: {
        type: "array",
        items: { type: "string" },
        description: "Absolute Windows file paths to attach (send/draft).",
      },
      replyAll: { type: "boolean", description: "Reply to all recipients (reply)." },
      targetFolder: { type: "string", description: "Destination folder name (move)." },
      read: { type: "boolean", description: "Mark as read (true) or unread (false) — for mark." },
      flag: { type: "boolean", description: "Set (true) or clear (false) a follow-up flag — for flag." },
    },
    required: ["operation"],
  },
};

export const OUTLOOK_CALENDAR_TOOL: Tool = {
  name: "outlook_calendar",
  description:
    "View and manage the Microsoft Outlook calendar on Windows. Operations: today, upcoming, search, create, update, delete. Creating with attendees sends a meeting invitation.",
  inputSchema: {
    type: "object",
    properties: {
      operation: {
        type: "string",
        enum: ["today", "upcoming", "search", "create", "update", "delete"],
        description: "The calendar operation to perform.",
      },
      searchTerm: { type: "string", description: "Text to match in event subject/location (search)." },
      limit: { type: "number", description: "Max events to return (default 10)." },
      days: { type: "number", description: "Days ahead for 'upcoming' (default 7)." },
      entryId: { type: "string", description: "Outlook EntryID of the event (update/delete)." },
      subject: { type: "string", description: "Event title (create; optional for update)." },
      start: { type: "string", description: "Start time, ISO 8601 local e.g. 2026-06-05T14:00:00 (create/update)." },
      end: { type: "string", description: "End time, ISO 8601 local (create/update)." },
      location: { type: "string", description: "Event location (optional)." },
      body: { type: "string", description: "Event description (optional)." },
      attendees: {
        type: "string",
        description: "Comma-separated attendee email addresses; presence turns the event into a meeting invite (create).",
      },
    },
    required: ["operation"],
  },
};

export const OUTLOOK_CONTACTS_TOOL: Tool = {
  name: "outlook_contacts",
  description:
    "List, search and create contacts in Microsoft Outlook on Windows. Operations: list, search, create.",
  inputSchema: {
    type: "object",
    properties: {
      operation: {
        type: "string",
        enum: ["list", "search", "create"],
        description: "The contacts operation to perform.",
      },
      searchTerm: { type: "string", description: "Text to match in name/email/company (search)." },
      limit: { type: "number", description: "Max contacts to return (default 25)." },
      fullName: { type: "string", description: "Contact full name (create)." },
      email: { type: "string", description: "Primary email address (create)." },
      phone: { type: "string", description: "Business phone number (create)." },
      company: { type: "string", description: "Company name (create)." },
      jobTitle: { type: "string", description: "Job title (create)." },
    },
    required: ["operation"],
  },
};

export const OUTLOOK_ATTACHMENTS_TOOL: Tool = {
  name: "outlook_attachments",
  description:
    "Inspect and download (save to disk) attachments from an Outlook message on Windows. Operations: list, download. Use outlook_mail get/unread/search first to obtain the message EntryID.",
  inputSchema: {
    type: "object",
    properties: {
      operation: {
        type: "string",
        enum: ["list", "download"],
        description: "list attachment metadata, or download (save) attachments to disk.",
      },
      entryId: { type: "string", description: "Outlook EntryID of the message (required)." },
      saveDir: {
        type: "string",
        description: "Directory to save into (download). Defaults to the user's Downloads folder.",
      },
      index: {
        type: "number",
        description: "1-based attachment index to download a single attachment (optional).",
      },
      fileName: {
        type: "string",
        description: "Download only the attachment with this exact file name (optional).",
      },
    },
    required: ["operation", "entryId"],
  },
};

export const OUTLOOK_FILES_TOOL: Tool = {
  name: "outlook_files",
  description:
    "Read, write and inspect files on the Windows host filesystem. Useful for staging attachments before sending, reading downloaded attachments, or saving content. Operations: read, write, info, list, delete. Binary content is base64; text is UTF-8.",
  inputSchema: {
    type: "object",
    properties: {
      operation: {
        type: "string",
        enum: ["read", "write", "info", "list", "delete"],
        description: "The filesystem operation to perform.",
      },
      path: { type: "string", description: "Absolute path to the file or directory." },
      content: {
        type: "string",
        description: "Content to write (write). Text (UTF-8) or base64 — see encoding.",
      },
      encoding: {
        type: "string",
        enum: ["auto", "text", "base64"],
        description:
          "Content encoding. On read, 'auto' (default) detects binary→base64 vs text→UTF-8. On write, content is UTF-8 unless you pass 'base64' explicitly (auto/text never guess, so text is never corrupted).",
      },
      overwrite: {
        type: "boolean",
        description: "Allow overwriting an existing file on write (default true).",
      },
      maxBytes: {
        type: "number",
        description: "Cap bytes returned by read (default 5,000,000); larger files are truncated.",
      },
    },
    required: ["operation", "path"],
  },
};

export const ALL_TOOLS: Tool[] = [
  OUTLOOK_MAIL_TOOL,
  OUTLOOK_CALENDAR_TOOL,
  OUTLOOK_CONTACTS_TOOL,
  OUTLOOK_ATTACHMENTS_TOOL,
  OUTLOOK_FILES_TOOL,
];
