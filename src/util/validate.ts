// ====================================================
// Input validation & typed argument shapes
// ====================================================
//
// Each MCP tool call is validated here before any PowerShell is generated.
// Validators throw a {@link ValidationError} with an actionable message that
// is surfaced to the model, mirroring (and extending) the type guards in the
// macOS original.

export class ValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ValidationError";
  }
}

function asObject(args: unknown): Record<string, unknown> {
  if (typeof args !== "object" || args === null || Array.isArray(args)) {
    throw new ValidationError("Arguments must be an object.");
  }
  return args as Record<string, unknown>;
}

function reqStr(o: Record<string, unknown>, key: string, op: string): string {
  const v = o[key];
  if (typeof v !== "string" || v.trim() === "") {
    throw new ValidationError(`"${key}" is required for the "${op}" operation.`);
  }
  return v;
}

function optStr(o: Record<string, unknown>, key: string): string | undefined {
  const v = o[key];
  if (v === undefined || v === null) return undefined;
  if (typeof v !== "string") {
    throw new ValidationError(`"${key}" must be a string.`);
  }
  return v;
}

function optNum(o: Record<string, unknown>, key: string): number | undefined {
  const v = o[key];
  if (v === undefined || v === null) return undefined;
  const n = typeof v === "string" ? Number(v) : v;
  if (typeof n !== "number" || !Number.isFinite(n)) {
    throw new ValidationError(`"${key}" must be a number.`);
  }
  return n;
}

function optBool(o: Record<string, unknown>, key: string): boolean | undefined {
  const v = o[key];
  if (v === undefined || v === null) return undefined;
  if (typeof v === "boolean") return v;
  if (v === "true") return true;
  if (v === "false") return false;
  throw new ValidationError(`"${key}" must be a boolean.`);
}

function reqBool(o: Record<string, unknown>, key: string, op: string): boolean {
  const v = optBool(o, key);
  if (v === undefined) {
    throw new ValidationError(`"${key}" is required for the "${op}" operation.`);
  }
  return v;
}

function optStrArray(
  o: Record<string, unknown>,
  key: string,
): string[] | undefined {
  const v = o[key];
  if (v === undefined || v === null) return undefined;
  if (!Array.isArray(v) || v.some((x) => typeof x !== "string")) {
    throw new ValidationError(`"${key}" must be an array of strings.`);
  }
  return v as string[];
}

function oneOf<T extends string>(
  o: Record<string, unknown>,
  key: string,
  allowed: readonly T[],
): T {
  const v = o[key];
  if (typeof v !== "string" || !allowed.includes(v as T)) {
    throw new ValidationError(
      `"${key}" must be one of: ${allowed.join(", ")}.`,
    );
  }
  return v as T;
}

// ---- Mail ---------------------------------------------------------------

export const MAIL_OPERATIONS = [
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
] as const;
export type MailOperation = (typeof MAIL_OPERATIONS)[number];

export interface MailArgs {
  operation: MailOperation;
  folder?: string;
  limit?: number;
  searchTerm?: string;
  entryId?: string;
  to?: string;
  cc?: string;
  bcc?: string;
  subject?: string;
  body?: string;
  isHtml?: boolean;
  attachments?: string[];
  replyAll?: boolean;
  targetFolder?: string;
  read?: boolean;
  flag?: boolean;
}

export function validateMailArgs(args: unknown): MailArgs {
  const o = asObject(args);
  const operation = oneOf(o, "operation", MAIL_OPERATIONS);
  const base: MailArgs = {
    operation,
    folder: optStr(o, "folder"),
    limit: optNum(o, "limit"),
    searchTerm: optStr(o, "searchTerm"),
    entryId: optStr(o, "entryId"),
    to: optStr(o, "to"),
    cc: optStr(o, "cc"),
    bcc: optStr(o, "bcc"),
    subject: optStr(o, "subject"),
    body: optStr(o, "body"),
    isHtml: optBool(o, "isHtml"),
    attachments: optStrArray(o, "attachments"),
    replyAll: optBool(o, "replyAll"),
    targetFolder: optStr(o, "targetFolder"),
    read: optBool(o, "read"),
    flag: optBool(o, "flag"),
  };

  switch (operation) {
    case "search":
      reqStr(o, "searchTerm", operation);
      break;
    case "get":
      reqStr(o, "entryId", operation);
      break;
    case "send":
    case "draft":
      reqStr(o, "to", operation);
      reqStr(o, "subject", operation);
      reqStr(o, "body", operation);
      break;
    case "reply":
      reqStr(o, "entryId", operation);
      reqStr(o, "body", operation);
      break;
    case "forward":
      reqStr(o, "entryId", operation);
      reqStr(o, "to", operation);
      break;
    case "move":
      reqStr(o, "entryId", operation);
      reqStr(o, "targetFolder", operation);
      break;
    case "delete":
      reqStr(o, "entryId", operation);
      break;
    case "mark":
      reqStr(o, "entryId", operation);
      base.read = reqBool(o, "read", operation);
      break;
    case "flag":
      reqStr(o, "entryId", operation);
      base.flag = reqBool(o, "flag", operation);
      break;
  }
  return base;
}

// ---- Calendar -----------------------------------------------------------

export const CALENDAR_OPERATIONS = [
  "today",
  "upcoming",
  "search",
  "create",
  "update",
  "delete",
] as const;
export type CalendarOperation = (typeof CALENDAR_OPERATIONS)[number];

export interface CalendarArgs {
  operation: CalendarOperation;
  searchTerm?: string;
  limit?: number;
  days?: number;
  entryId?: string;
  subject?: string;
  start?: string;
  end?: string;
  location?: string;
  body?: string;
  attendees?: string;
}

export function validateCalendarArgs(args: unknown): CalendarArgs {
  const o = asObject(args);
  const operation = oneOf(o, "operation", CALENDAR_OPERATIONS);
  const base: CalendarArgs = {
    operation,
    searchTerm: optStr(o, "searchTerm"),
    limit: optNum(o, "limit"),
    days: optNum(o, "days"),
    entryId: optStr(o, "entryId"),
    subject: optStr(o, "subject"),
    start: optStr(o, "start"),
    end: optStr(o, "end"),
    location: optStr(o, "location"),
    body: optStr(o, "body"),
    attendees: optStr(o, "attendees"),
  };

  switch (operation) {
    case "search":
      reqStr(o, "searchTerm", operation);
      break;
    case "create":
      reqStr(o, "subject", operation);
      reqStr(o, "start", operation);
      reqStr(o, "end", operation);
      break;
    case "update":
      reqStr(o, "entryId", operation);
      break;
    case "delete":
      reqStr(o, "entryId", operation);
      break;
  }
  return base;
}

// ---- Contacts -----------------------------------------------------------

export const CONTACT_OPERATIONS = ["list", "search", "create"] as const;
export type ContactOperation = (typeof CONTACT_OPERATIONS)[number];

export interface ContactArgs {
  operation: ContactOperation;
  searchTerm?: string;
  limit?: number;
  fullName?: string;
  email?: string;
  phone?: string;
  company?: string;
  jobTitle?: string;
}

export function validateContactArgs(args: unknown): ContactArgs {
  const o = asObject(args);
  const operation = oneOf(o, "operation", CONTACT_OPERATIONS);
  const base: ContactArgs = {
    operation,
    searchTerm: optStr(o, "searchTerm"),
    limit: optNum(o, "limit"),
    fullName: optStr(o, "fullName"),
    email: optStr(o, "email"),
    phone: optStr(o, "phone"),
    company: optStr(o, "company"),
    jobTitle: optStr(o, "jobTitle"),
  };

  switch (operation) {
    case "search":
      reqStr(o, "searchTerm", operation);
      break;
    case "create":
      reqStr(o, "fullName", operation);
      break;
  }
  return base;
}

// ---- Attachments (Outlook side) ----------------------------------------

export const ATTACHMENT_OPERATIONS = ["list", "download"] as const;
export type AttachmentOperation = (typeof ATTACHMENT_OPERATIONS)[number];

export interface AttachmentArgs {
  operation: AttachmentOperation;
  entryId: string;
  saveDir?: string;
  index?: number;
  fileName?: string;
}

export function validateAttachmentArgs(args: unknown): AttachmentArgs {
  const o = asObject(args);
  const operation = oneOf(o, "operation", ATTACHMENT_OPERATIONS);
  const entryId = reqStr(o, "entryId", operation);
  return {
    operation,
    entryId,
    saveDir: optStr(o, "saveDir"),
    index: optNum(o, "index"),
    fileName: optStr(o, "fileName"),
  };
}

// ---- Files (local filesystem side) -------------------------------------

export const FILE_OPERATIONS = [
  "read",
  "write",
  "info",
  "list",
  "delete",
] as const;
export type FileOperation = (typeof FILE_OPERATIONS)[number];

export interface FileArgs {
  operation: FileOperation;
  path: string;
  content?: string;
  encoding?: "auto" | "text" | "base64";
  overwrite?: boolean;
  maxBytes?: number;
}

export function validateFileArgs(args: unknown): FileArgs {
  const o = asObject(args);
  const operation = oneOf(o, "operation", FILE_OPERATIONS);
  const path = reqStr(o, "path", operation);
  const encoding = o.encoding === undefined ? undefined : oneOf(o, "encoding", ["auto", "text", "base64"] as const);
  const base: FileArgs = {
    operation,
    path,
    content: optStr(o, "content"),
    encoding,
    overwrite: optBool(o, "overwrite"),
    maxBytes: optNum(o, "maxBytes"),
  };
  if (operation === "write" && base.content === undefined) {
    throw new ValidationError(`"content" is required for the "write" operation.`);
  }
  return base;
}
