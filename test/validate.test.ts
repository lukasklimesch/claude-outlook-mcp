import { describe, expect, test } from "bun:test";
import {
  ValidationError,
  validateAttachmentArgs,
  validateCalendarArgs,
  validateContactArgs,
  validateFileArgs,
  validateMailArgs,
} from "../src/util/validate.js";

describe("validateMailArgs", () => {
  test("accepts a valid unread request", () => {
    const a = validateMailArgs({ operation: "unread", folder: "Inbox", limit: 5 });
    expect(a.operation).toBe("unread");
    expect(a.limit).toBe(5);
  });

  test("requires searchTerm for search", () => {
    expect(() => validateMailArgs({ operation: "search" })).toThrow(ValidationError);
  });

  test("requires to/subject/body for send", () => {
    expect(() => validateMailArgs({ operation: "send", to: "a@b.com" })).toThrow();
    const ok = validateMailArgs({
      operation: "send",
      to: "a@b.com",
      subject: "Hi",
      body: "Body",
      attachments: ["C:/x.pdf"],
    });
    expect(ok.attachments).toEqual(["C:/x.pdf"]);
  });

  test("requires entryId + boolean read for mark", () => {
    expect(() => validateMailArgs({ operation: "mark", entryId: "X" })).toThrow();
    const ok = validateMailArgs({ operation: "mark", entryId: "X", read: true });
    expect(ok.read).toBe(true);
  });

  test("coerces stringified numbers and booleans", () => {
    const a = validateMailArgs({ operation: "read", limit: "7", isHtml: "true" } as any);
    expect(a.limit).toBe(7);
    expect(a.isHtml).toBe(true);
  });

  test("rejects unknown operation and bad attachment types", () => {
    expect(() => validateMailArgs({ operation: "nope" })).toThrow();
    expect(() =>
      validateMailArgs({ operation: "send", to: "a", subject: "b", body: "c", attachments: [1] as any }),
    ).toThrow();
  });
});

describe("validateCalendarArgs", () => {
  test("create requires subject/start/end", () => {
    expect(() => validateCalendarArgs({ operation: "create", subject: "x" })).toThrow();
    const ok = validateCalendarArgs({
      operation: "create",
      subject: "Sync",
      start: "2026-06-05T14:00:00",
      end: "2026-06-05T15:00:00",
    });
    expect(ok.subject).toBe("Sync");
  });

  test("update/delete require entryId", () => {
    expect(() => validateCalendarArgs({ operation: "update" })).toThrow();
    expect(() => validateCalendarArgs({ operation: "delete" })).toThrow();
  });
});

describe("validateContactArgs", () => {
  test("create requires fullName", () => {
    expect(() => validateContactArgs({ operation: "create" })).toThrow();
    expect(validateContactArgs({ operation: "create", fullName: "Jane" }).fullName).toBe("Jane");
  });
});

describe("validateAttachmentArgs", () => {
  test("requires entryId", () => {
    expect(() => validateAttachmentArgs({ operation: "list" })).toThrow();
    const a = validateAttachmentArgs({ operation: "download", entryId: "E", index: 2 });
    expect(a.index).toBe(2);
  });
});

describe("validateFileArgs", () => {
  test("write requires content", () => {
    expect(() => validateFileArgs({ operation: "write", path: "C:/x.txt" })).toThrow();
    const a = validateFileArgs({
      operation: "write",
      path: "C:/x.txt",
      content: "hi",
      encoding: "text",
    });
    expect(a.encoding).toBe("text");
  });

  test("read requires only a path and validates encoding enum", () => {
    expect(validateFileArgs({ operation: "read", path: "C:/x" }).path).toBe("C:/x");
    expect(() => validateFileArgs({ operation: "read", path: "C:/x", encoding: "rot13" } as any)).toThrow();
  });
});
