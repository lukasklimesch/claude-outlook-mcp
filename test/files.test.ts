import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  deleteFileOp,
  fileInfoOp,
  listDirOp,
  looksBinary,
  readFileOp,
  writeFileOp,
} from "../src/files.js";

let dir = "";

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "outlook-mcp-files-"));
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

describe("write + read (text)", () => {
  test("round-trips UTF-8 text", async () => {
    const p = join(dir, "note.txt");
    const w = await writeFileOp({ operation: "write", path: p, content: "héllo ✓", encoding: "text" });
    expect(w.bytesWritten).toBeGreaterThan(0);

    const r = await readFileOp({ operation: "read", path: p });
    expect(r.encoding).toBe("text");
    expect(r.content).toBe("héllo ✓");
    expect(r.truncated).toBe(false);
  });
});

describe("write + read (binary/base64)", () => {
  test("round-trips binary via base64 and auto-detects it on read", async () => {
    const p = join(dir, "blob.bin");
    const original = Buffer.from([0, 1, 2, 250, 255, 13, 10]);
    const w = await writeFileOp({
      operation: "write",
      path: p,
      content: original.toString("base64"),
      encoding: "base64",
    });
    expect(w.encoding).toBe("base64");

    const r = await readFileOp({ operation: "read", path: p, encoding: "auto" });
    expect(r.encoding).toBe("base64"); // NUL byte => detected as binary
    expect(Buffer.from(r.content, "base64").equals(original)).toBe(true);
  });

  test("looksBinary detects NUL bytes", () => {
    expect(looksBinary(Buffer.from([65, 0, 66]))).toBe(true);
    expect(looksBinary(Buffer.from("plain text"))).toBe(false);
  });

  test("a base64-looking word is written verbatim under auto/text (no corruption)", async () => {
    // "metadata" is 8 chars of the base64 alphabet; it must NOT be decoded.
    const p = join(dir, "word.txt");
    await writeFileOp({ operation: "write", path: p, content: "metadata" });
    const r = await readFileOp({ operation: "read", path: p });
    expect(r.content).toBe("metadata");
  });
});

describe("read truncation", () => {
  test("caps returned bytes and flags truncation", async () => {
    const p = join(dir, "big.txt");
    await writeFileOp({ operation: "write", path: p, content: "A".repeat(1000), encoding: "text" });
    const r = await readFileOp({ operation: "read", path: p, maxBytes: 100 });
    expect(r.truncated).toBe(true);
    expect(r.bytesReturned).toBe(100);
    expect(r.size).toBe(1000);
  });
});

describe("overwrite guard", () => {
  test("refuses to overwrite when overwrite=false", async () => {
    const p = join(dir, "guard.txt");
    await writeFileOp({ operation: "write", path: p, content: "first" });
    await expect(
      writeFileOp({ operation: "write", path: p, content: "second", overwrite: false }),
    ).rejects.toThrow();
  });
});

describe("info / list / delete", () => {
  test("info reports file metadata", async () => {
    const p = join(dir, "meta.txt");
    await writeFileOp({ operation: "write", path: p, content: "data" });
    const info = await fileInfoOp({ operation: "info", path: p });
    expect(info.isFile).toBe(true);
    expect(info.isDirectory).toBe(false);
    expect(info.size).toBe(4);
  });

  test("list enumerates directory entries", async () => {
    const entries = await listDirOp({ operation: "list", path: dir });
    const names = entries.map((e) => e.name);
    expect(names).toContain("note.txt");
    expect(entries.find((e) => e.name === "note.txt")?.type).toBe("file");
  });

  test("delete removes a file and refuses directories", async () => {
    const p = join(dir, "temp-del.txt");
    await writeFileOp({ operation: "write", path: p, content: "x" });
    const res = await deleteFileOp({ operation: "delete", path: p });
    expect(res.deleted).toBe(true);
    await expect(deleteFileOp({ operation: "delete", path: dir })).rejects.toThrow();
  });
});
