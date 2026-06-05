// ====================================================
// MCP server: tool registration & request dispatch
// ====================================================

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import type { OutlookRunner } from "./types.js";
import { ALL_TOOLS } from "./tools.js";
import {
  validateAttachmentArgs,
  validateCalendarArgs,
  validateContactArgs,
  validateFileArgs,
  validateMailArgs,
} from "./util/validate.js";
import {
  formatAttachments,
  formatContacts,
  formatEvents,
  formatFolders,
  formatMailDetail,
  formatMailList,
  formatSavedFiles,
} from "./util/format.js";
import * as mail from "./outlook/mail.js";
import * as calendar from "./outlook/calendar.js";
import * as contacts from "./outlook/contacts.js";
import * as attachments from "./outlook/attachments.js";
import {
  deleteFileOp,
  fileInfoOp,
  listDirOp,
  readFileOp,
  writeFileOp,
  type ReadResult,
  type DirEntry,
} from "./files.js";
import { formatBytes } from "./util/format.js";

export const SERVER_NAME = "Outlook MCP Tool (Windows)";
export const SERVER_VERSION = "2.0.0";

function errMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

function formatReadResult(r: ReadResult): string {
  const header =
    `File: ${r.path}\n` +
    `Encoding: ${r.encoding} | Size: ${formatBytes(r.size)} | Returned: ${formatBytes(r.bytesReturned)}` +
    (r.truncated ? " (truncated)" : "") +
    `\n\n--- Content ---\n`;
  return header + r.content;
}

function formatDirList(path: string, entries: DirEntry[]): string {
  if (entries.length === 0) return `Directory is empty: ${path}`;
  const lines = entries.map((e) => {
    const tag = e.type === "directory" ? "<DIR> " : "      ";
    const size = e.type === "file" ? formatBytes(e.size) : "";
    return `  ${tag}${e.name}${size ? `  (${size})` : ""}`;
  });
  return `Contents of ${path} (${entries.length}):\n${lines.join("\n")}`;
}

/**
 * Validate, route and format a single tool call. Exported so the full request
 * flow can be exercised in tests with a fake {@link OutlookRunner}.
 */
export async function dispatch(
  runner: OutlookRunner,
  name: string,
  args: Record<string, unknown>,
): Promise<string> {
  switch (name) {
    case "outlook_mail": {
      const a = validateMailArgs(args);
      switch (a.operation) {
        case "unread":
          return formatMailList(
            await mail.getUnread(runner, a),
            a.folder ? `in "${a.folder}"` : "in Inbox",
          );
        case "read":
          return formatMailList(
            await mail.readMail(runner, a),
            a.folder ? `in "${a.folder}"` : "in Inbox",
          );
        case "search":
          return formatMailList(
            await mail.searchMail(runner, a),
            `matching "${a.searchTerm}"${a.folder ? ` in "${a.folder}"` : ""}`,
          );
        case "get":
          return formatMailDetail(await mail.getMail(runner, a));
        case "send":
          await mail.sendMail(runner, a);
          return `Email sent to ${a.to} (subject: "${a.subject}").`;
        case "draft": {
          const r = (await mail.createDraft(runner, a)) as { entryId?: string };
          return `Draft saved${r?.entryId ? ` (EntryID: ${r.entryId})` : ""}.`;
        }
        case "reply":
          await mail.replyMail(runner, a);
          return `Reply ${a.replyAll ? "(all) " : ""}sent.`;
        case "forward":
          await mail.forwardMail(runner, a);
          return `Message forwarded to ${a.to}.`;
        case "move":
          await mail.moveMail(runner, a);
          return `Message moved to "${a.targetFolder}".`;
        case "delete":
          await mail.deleteMail(runner, a);
          return "Message deleted (moved to Deleted Items).";
        case "mark":
          await mail.markMail(runner, a);
          return `Message marked as ${a.read ? "read" : "unread"}.`;
        case "flag":
          await mail.flagMail(runner, a);
          return `Follow-up flag ${a.flag ? "set" : "cleared"}.`;
        case "folders":
          return formatFolders(await mail.listFolders(runner));
      }
      break;
    }

    case "outlook_calendar": {
      const a = validateCalendarArgs(args);
      switch (a.operation) {
        case "today":
          return formatEvents(await calendar.todayEvents(runner, a), "for today");
        case "upcoming":
          return formatEvents(
            await calendar.upcomingEvents(runner, a),
            `in the next ${a.days ?? 7} day(s)`,
          );
        case "search":
          return formatEvents(
            await calendar.searchEvents(runner, a),
            `matching "${a.searchTerm}"`,
          );
        case "create": {
          const r = (await calendar.createEvent(runner, a)) as {
            status?: string;
            entryId?: string;
          };
          return r?.status === "invited"
            ? `Meeting invitation sent for "${a.subject}".`
            : `Event "${a.subject}" created${r?.entryId ? ` (EntryID: ${r.entryId})` : ""}.`;
        }
        case "update":
          await calendar.updateEvent(runner, a);
          return "Event updated.";
        case "delete":
          await calendar.deleteEvent(runner, a);
          return "Event deleted.";
      }
      break;
    }

    case "outlook_contacts": {
      const a = validateContactArgs(args);
      switch (a.operation) {
        case "list":
          return formatContacts(await contacts.listContacts(runner, a), "");
        case "search":
          return formatContacts(
            await contacts.searchContacts(runner, a),
            `matching "${a.searchTerm}"`,
          );
        case "create": {
          const c = await contacts.createContact(runner, a);
          return `Contact "${c.fullName}" created.`;
        }
      }
      break;
    }

    case "outlook_attachments": {
      const a = validateAttachmentArgs(args);
      if (a.operation === "list") {
        return formatAttachments(await attachments.listAttachments(runner, a));
      }
      return formatSavedFiles(await attachments.downloadAttachments(runner, a));
    }

    case "outlook_files": {
      const a = validateFileArgs(args);
      switch (a.operation) {
        case "read":
          return formatReadResult(await readFileOp(a));
        case "write": {
          const r = await writeFileOp(a);
          return `Wrote ${formatBytes(r.bytesWritten)} (${r.encoding}) to ${r.path}.`;
        }
        case "info":
          return JSON.stringify(await fileInfoOp(a), null, 2);
        case "list":
          return formatDirList(a.path, await listDirOp(a));
        case "delete": {
          const r = await deleteFileOp(a);
          return `Deleted ${r.path}.`;
        }
      }
      break;
    }

    default:
      throw new Error(`Unknown tool: ${name}`);
  }

  // Unreachable for valid operations (exhaustive switches above).
  throw new Error(`Unsupported operation for tool ${name}.`);
}

/** Build a configured MCP server bound to the given Outlook runner. */
export function createServer(runner: OutlookRunner): Server {
  const server = new Server(
    { name: SERVER_NAME, version: SERVER_VERSION },
    { capabilities: { tools: {} } },
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: ALL_TOOLS,
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;
    try {
      const text = await dispatch(runner, name, (args ?? {}) as Record<string, unknown>);
      return { content: [{ type: "text", text }], isError: false };
    } catch (err) {
      return {
        content: [{ type: "text", text: `Error: ${errMessage(err)}` }],
        isError: true,
      };
    }
  });

  return server;
}
