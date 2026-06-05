import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { dispatch, createServer } from "../src/server.js";
import { ALL_TOOLS } from "../src/tools.js";
import type { PsResult } from "../src/types.js";
import { FakeRunner } from "./helpers.js";

const ok = (data: unknown): PsResult => ({ ok: true, data });

const sampleSummary = {
  entryId: "ENTRY-1",
  subject: "Quarterly report",
  senderName: "Jane Doe",
  senderEmail: "jane@example.com",
  received: "2026-06-05T09:00:00",
  unread: true,
  hasAttachments: true,
  attachmentCount: 1,
  size: 2048,
  preview: "Here is the report",
};

const sampleDetail = {
  ...sampleSummary,
  body: "Full body",
  attachments: [{ index: 1, fileName: "a.pdf", size: 1024, type: 1 }],
};

/**
 * Return only the operation body — the part of the script AFTER the shared
 * preamble bootstrap. The preamble defines every Convert-* helper, so we must
 * match against the body to tell operations apart.
 */
function opBody(script: string): string {
  const marker = "$ns = $app.GetNamespace('MAPI')";
  const idx = script.indexOf(marker);
  return idx >= 0 ? script.slice(idx + marker.length) : script;
}

/** A responder that returns realistic shapes based on the operation body. */
function smartResponder(script: string, payload?: Record<string, unknown>): PsResult {
  const body = opBody(script);
  if (body.includes("SaveAsFile")) return ok(["C:\\Users\\Joe\\Downloads\\a.pdf"]);
  if (body.includes("CreateItem($olContactItem)"))
    return ok({ fullName: payload?.fullName, email: payload?.email, phone: payload?.phone });
  if (body.includes("CreateItem($olAppointmentItem)"))
    return ok(payload?.attendees ? { status: "invited" } : { status: "created", entryId: "EV-NEW" });
  if (body.includes("Convert-MailDetail")) return ok(sampleDetail);
  if (body.includes("Convert-MailSummary")) return ok([sampleSummary]);
  if (body.includes("Add-FolderInfo"))
    return ok([{ name: "Inbox", path: "Mailbox\\Inbox", unread: 2, total: 9 }]);
  if (body.includes("Convert-Attachment"))
    return ok([{ index: 1, fileName: "a.pdf", size: 1024, type: 1 }]);
  if (body.includes("Convert-Contact"))
    return ok([{ fullName: "Jane Doe", email: "jane@x.com", phone: "555" }]);
  if (body.includes("Convert-Event"))
    return ok([{ entryId: "EV1", subject: "Standup", start: "s", end: "e", location: "Room" }]);
  return ok({}); // action operations (send/draft/move/etc.)
}

describe("dispatch — mail", () => {
  test("unread formats a list and targets the Inbox by default", async () => {
    const r = new FakeRunner(smartResponder);
    const out = await dispatch(r, "outlook_mail", { operation: "unread" });
    expect(out).toContain("Found 1 email(s) in Inbox");
    expect(out).toContain("ENTRY-1");
    expect(r.last().script).toContain("Restrict('[UnRead] = True')");
  });

  test("get returns full detail", async () => {
    const r = new FakeRunner(smartResponder);
    const out = await dispatch(r, "outlook_mail", { operation: "get", entryId: "ENTRY-1" });
    expect(out).toContain("--- Body ---");
    expect(out).toContain("Full body");
    expect(r.last().payload?.entryId).toBe("ENTRY-1");
  });

  test("send wires the right COM calls and confirms", async () => {
    const r = new FakeRunner(smartResponder);
    const out = await dispatch(r, "outlook_mail", {
      operation: "send",
      to: "bob@example.com",
      subject: "Hello",
      body: "Hi there",
      attachments: ["C:/files/report.pdf"],
    });
    expect(out).toContain("Email sent to bob@example.com");
    const { script, payload } = r.last();
    expect(script).toContain("$app.CreateItem($olMailItem)");
    expect(script).toContain("$mail.Send()");
    expect(script).toContain("$mail.Attachments.Add");
    expect(payload?.attachments).toEqual(["C:/files/report.pdf"]);
  });

  test("folders lists folder tree", async () => {
    const r = new FakeRunner(smartResponder);
    const out = await dispatch(r, "outlook_mail", { operation: "folders" });
    expect(out).toContain("Mailbox\\Inbox");
    expect(out).toContain("2 unread");
  });
});

describe("dispatch — injection safety", () => {
  test("malicious input never appears in the generated PowerShell", async () => {
    const r = new FakeRunner(smartResponder);
    const evil = '"; Remove-Item C:\\ -Recurse -Force #';
    await dispatch(r, "outlook_mail", {
      operation: "send",
      to: "a@b.com",
      subject: evil,
      body: evil,
    });
    const { script, payload } = r.last();
    // The dangerous text is carried as DATA, not woven into the script source.
    expect(script).not.toContain("Remove-Item");
    expect(script).toContain("$payload.subject");
    expect(payload?.subject).toBe(evil);
  });
});

describe("dispatch — calendar & contacts", () => {
  test("create meeting (with attendees) sends an invite", async () => {
    const r = new FakeRunner(smartResponder);
    const out = await dispatch(r, "outlook_calendar", {
      operation: "create",
      subject: "Sync",
      start: "2026-06-05T14:00:00",
      end: "2026-06-05T15:00:00",
      attendees: "a@x.com,b@x.com",
    });
    expect(out).toContain("Meeting invitation sent");
    expect(r.last().script).toContain("$app.CreateItem($olAppointmentItem)");
  });

  test("today builds a date Restrict filter", async () => {
    const r = new FakeRunner(smartResponder);
    const out = await dispatch(r, "outlook_calendar", { operation: "today" });
    expect(out).toContain("Standup");
    expect(String(r.last().payload?.filter)).toContain("[Start] >=");
  });

  test("contacts create confirms by name", async () => {
    const r = new FakeRunner(smartResponder);
    const out = await dispatch(r, "outlook_contacts", { operation: "create", fullName: "Jane Doe" });
    expect(out).toContain('Contact "Jane Doe" created');
  });
});

describe("dispatch — attachments", () => {
  test("download saves files and reports paths", async () => {
    const r = new FakeRunner(smartResponder);
    const out = await dispatch(r, "outlook_attachments", {
      operation: "download",
      entryId: "ENTRY-1",
    });
    expect(out).toContain("Saved 1 file");
    expect(out).toContain("a.pdf");
    expect(r.last().script).toContain("SaveAsFile");
  });
});

describe("dispatch — local files (real fs)", () => {
  let dir = "";
  beforeAll(async () => {
    dir = await mkdtemp(join(tmpdir(), "outlook-mcp-dispatch-"));
  });
  afterAll(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  test("write then read through the dispatcher", async () => {
    const r = new FakeRunner(smartResponder);
    const p = join(dir, "out.txt");
    const w = await dispatch(r, "outlook_files", {
      operation: "write",
      path: p,
      content: "from-claude",
    });
    expect(w).toContain("to " + p);
    const read = await dispatch(r, "outlook_files", { operation: "read", path: p });
    expect(read).toContain("from-claude");
    // File ops must not touch the PowerShell runner.
    expect(r.calls.length).toBe(0);
  });
});

describe("dispatch — error handling", () => {
  test("propagates Outlook errors from the runner", async () => {
    const r = new FakeRunner(() => ({ ok: false, error: "Outlook not running", category: "outlook" }));
    await expect(dispatch(r, "outlook_mail", { operation: "unread" })).rejects.toThrow(
      "Outlook not running",
    );
  });

  test("rejects invalid operations during validation", async () => {
    const r = new FakeRunner(smartResponder);
    await expect(dispatch(r, "outlook_mail", { operation: "frobnicate" })).rejects.toThrow();
    await expect(dispatch(r, "unknown_tool", {})).rejects.toThrow("Unknown tool");
  });
});

describe("createServer", () => {
  test("exposes exactly the five tools with unique names", () => {
    const names = ALL_TOOLS.map((t) => t.name);
    expect(names).toEqual([
      "outlook_mail",
      "outlook_calendar",
      "outlook_contacts",
      "outlook_attachments",
      "outlook_files",
    ]);
    expect(new Set(names).size).toBe(names.length);
  });

  test("builds a server bound to the runner without throwing", () => {
    const server = createServer(new FakeRunner(smartResponder));
    expect(typeof server.connect).toBe("function");
  });
});
