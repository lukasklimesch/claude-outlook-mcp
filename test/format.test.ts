import { describe, expect, test } from "bun:test";
import {
  formatAttachments,
  formatBytes,
  formatContacts,
  formatEvents,
  formatFolders,
  formatMailDetail,
  formatMailList,
  formatSavedFiles,
} from "../src/util/format.js";
import type {
  AttachmentInfo,
  CalendarEvent,
  ContactInfo,
  MailDetail,
  MailFolderInfo,
  MailSummary,
} from "../src/types.js";

const sampleMail: MailSummary = {
  entryId: "ENTRY-1",
  subject: "Quarterly report",
  senderName: "Jane Doe",
  senderEmail: "jane@example.com",
  received: "2026-06-05T09:00:00",
  unread: true,
  hasAttachments: true,
  attachmentCount: 2,
  size: 2048,
  preview: "Here is the report",
};

describe("formatMailList", () => {
  test("includes count, flags, sender and EntryID", () => {
    const out = formatMailList([sampleMail], "in Inbox");
    expect(out).toContain("Found 1 email(s) in Inbox");
    expect(out).toContain("UNREAD");
    expect(out).toContain("📎2");
    expect(out).toContain("jane@example.com");
    expect(out).toContain("ENTRY-1");
  });

  test("empty list message", () => {
    expect(formatMailList([], "in Inbox")).toBe("No emails found in Inbox");
  });
});

describe("formatMailDetail", () => {
  test("renders headers, attachments and body", () => {
    const detail: MailDetail = {
      ...sampleMail,
      to: "me@example.com",
      body: "Full body text",
      attachments: [{ index: 1, fileName: "a.pdf", size: 1024, type: 1 }],
    };
    const out = formatMailDetail(detail);
    expect(out).toContain("Subject: Quarterly report");
    expect(out).toContain("Attachments (1)");
    expect(out).toContain("a.pdf");
    expect(out).toContain("--- Body ---");
    expect(out).toContain("Full body text");
  });
});

describe("formatEvents / formatContacts / formatFolders / formatAttachments", () => {
  test("events", () => {
    const ev: CalendarEvent = {
      entryId: "EV1",
      subject: "Standup",
      start: "2026-06-05T09:00:00",
      end: "2026-06-05T09:15:00",
      location: "Room 1",
    };
    const out = formatEvents([ev], "for today");
    expect(out).toContain("Standup");
    expect(out).toContain("Room 1");
    expect(out).toContain("EV1");
  });

  test("contacts", () => {
    const c: ContactInfo = { fullName: "Jane Doe", email: "j@x.com", phone: "555" };
    expect(formatContacts([c], "")).toContain("Jane Doe");
  });

  test("folders with counts", () => {
    const f: MailFolderInfo = { name: "Inbox", path: "Mailbox\\Inbox", unread: 3, total: 10 };
    const out = formatFolders([f]);
    expect(out).toContain("Mailbox\\Inbox");
    expect(out).toContain("3 unread");
  });

  test("attachments and saved files", () => {
    const a: AttachmentInfo[] = [{ index: 1, fileName: "x.pdf", size: 2048, type: 1 }];
    expect(formatAttachments(a)).toContain("x.pdf");
    expect(formatAttachments([])).toContain("no attachments");
    expect(formatSavedFiles(["C:/d/x.pdf"])).toContain("Saved 1 file");
  });
});

describe("formatBytes", () => {
  test("scales units", () => {
    expect(formatBytes(512)).toBe("512 B");
    expect(formatBytes(2048)).toBe("2.0 KB");
    expect(formatBytes(5 * 1024 * 1024)).toBe("5.0 MB");
    expect(formatBytes(undefined)).toBe("unknown size");
  });
});
