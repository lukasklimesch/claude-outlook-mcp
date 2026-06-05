// ====================================================
// Shared types for the Windows Outlook MCP server
// ====================================================

/**
 * The normalized result of running a PowerShell/Outlook-COM script.
 *
 * Every script emitted by this server prints exactly one JSON object on
 * stdout shaped like `{ ok, data }` or `{ ok, error, category }`. The
 * `parsePsOutput` helper turns raw stdout into this structure.
 */
export interface PsResult<T = unknown> {
  ok: boolean;
  data?: T;
  error?: string;
  /** Coarse error class, e.g. "outlook", "parse", "empty", "spawn". */
  category?: string;
}

/**
 * Abstraction over "run a PowerShell script that drives Outlook COM and
 * give me back parsed JSON". The production implementation spawns
 * `powershell.exe`; tests inject a fake so the entire server can be
 * exercised on any platform.
 */
export interface OutlookRunner {
  run<T = unknown>(
    script: string,
    payload?: Record<string, unknown>,
  ): Promise<PsResult<T>>;
}

// ---- Domain shapes returned to the model -------------------------------

export interface MailSummary {
  entryId: string;
  subject: string;
  senderName: string;
  senderEmail: string;
  to?: string;
  cc?: string;
  received: string;
  unread: boolean;
  hasAttachments: boolean;
  attachmentCount: number;
  size?: number;
  preview?: string;
}

export interface MailDetail extends MailSummary {
  body: string;
  attachments: AttachmentInfo[];
}

export interface AttachmentInfo {
  index: number;
  fileName: string;
  size: number;
  type: number;
}

export interface CalendarEvent {
  entryId: string;
  subject: string;
  start: string;
  end: string;
  location: string;
  organizer?: string;
  body?: string;
  isAllDay?: boolean;
}

export interface ContactInfo {
  entryId?: string;
  fullName: string;
  email: string;
  phone: string;
  company?: string;
  jobTitle?: string;
}

export interface MailFolderInfo {
  name: string;
  path: string;
  unread?: number;
  total?: number;
}
