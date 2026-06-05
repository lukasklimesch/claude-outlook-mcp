// ====================================================
// Local filesystem tools: read / write / info / list / delete
// ====================================================
//
// These operate on the host filesystem (Windows in production) and are the
// "read, write, download" file primitives. They are pure Node/`fs` code with
// no Outlook dependency, so they execute — and are unit-tested — on any OS.
// Binary content is exchanged as base64; text as UTF-8.

import {
  mkdir,
  open,
  readFile,
  readdir,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import { dirname } from "node:path";
import type { FileArgs } from "./util/validate.js";

const DEFAULT_MAX_BYTES = 5_000_000;

function trimQuotes(p: string): string {
  let s = p.trim();
  if (
    (s.startsWith('"') && s.endsWith('"')) ||
    (s.startsWith("'") && s.endsWith("'"))
  ) {
    s = s.slice(1, -1);
  }
  return s;
}

/** Heuristic: treat content with NUL bytes as binary. */
export function looksBinary(buf: Buffer): boolean {
  const n = Math.min(buf.length, 8000);
  for (let i = 0; i < n; i++) {
    if (buf[i] === 0) return true;
  }
  return false;
}

export interface ReadResult {
  path: string;
  encoding: "text" | "base64";
  size: number;
  bytesReturned: number;
  truncated: boolean;
  content: string;
}

export async function readFileOp(args: FileArgs): Promise<ReadResult> {
  const path = trimQuotes(args.path);
  const st = await stat(path);
  if (!st.isFile()) throw new Error(`Not a file: ${path}`);

  const cap = args.maxBytes && args.maxBytes > 0 ? args.maxBytes : DEFAULT_MAX_BYTES;
  let buf: Buffer;
  let truncated = false;

  if (st.size > cap) {
    const fh = await open(path, "r");
    try {
      const tmp = Buffer.alloc(cap);
      const { bytesRead } = await fh.read(tmp, 0, cap, 0);
      buf = tmp.subarray(0, bytesRead);
    } finally {
      await fh.close();
    }
    truncated = true;
  } else {
    buf = await readFile(path);
  }

  const requested = args.encoding ?? "auto";
  const actual: "text" | "base64" =
    requested === "auto" ? (looksBinary(buf) ? "base64" : "text") : requested;

  return {
    path,
    encoding: actual,
    size: st.size,
    bytesReturned: buf.length,
    truncated,
    content: actual === "base64" ? buf.toString("base64") : buf.toString("utf8"),
  };
}

export interface WriteResult {
  path: string;
  encoding: "text" | "base64";
  bytesWritten: number;
}

export async function writeFileOp(args: FileArgs): Promise<WriteResult> {
  const path = trimQuotes(args.path);
  const content = args.content ?? "";
  // Writing is unambiguous and must never guess: only an explicit "base64"
  // request decodes. "auto"/"text" write the bytes verbatim as UTF-8, so a
  // plain word that merely looks like base64 is never silently corrupted.
  const actual: "text" | "base64" = args.encoding === "base64" ? "base64" : "text";

  if (args.overwrite === false) {
    let exists = false;
    try {
      await stat(path);
      exists = true;
    } catch {
      exists = false;
    }
    if (exists) {
      throw new Error(`Refusing to overwrite existing file: ${path}`);
    }
  }

  const buf =
    actual === "base64"
      ? Buffer.from(content, "base64")
      : Buffer.from(content, "utf8");

  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, buf);
  return { path, encoding: actual, bytesWritten: buf.length };
}

export interface FileInfoResult {
  path: string;
  isFile: boolean;
  isDirectory: boolean;
  size: number;
  modified: string;
  created: string;
}

export async function fileInfoOp(args: FileArgs): Promise<FileInfoResult> {
  const path = trimQuotes(args.path);
  const st = await stat(path);
  return {
    path,
    isFile: st.isFile(),
    isDirectory: st.isDirectory(),
    size: st.size,
    modified: st.mtime.toISOString(),
    created: st.birthtime.toISOString(),
  };
}

export interface DirEntry {
  name: string;
  type: "file" | "directory" | "other";
  size: number;
}

export async function listDirOp(args: FileArgs): Promise<DirEntry[]> {
  const path = trimQuotes(args.path);
  const entries = await readdir(path, { withFileTypes: true });
  const out: DirEntry[] = [];
  for (const e of entries) {
    let size = 0;
    let type: DirEntry["type"] = "other";
    if (e.isFile()) type = "file";
    else if (e.isDirectory()) type = "directory";
    if (type === "file") {
      try {
        size = (await stat(`${path}/${e.name}`)).size;
      } catch {
        size = 0;
      }
    }
    out.push({ name: e.name, type, size });
  }
  return out;
}

export async function deleteFileOp(args: FileArgs): Promise<{ path: string; deleted: boolean }> {
  const path = trimQuotes(args.path);
  const st = await stat(path);
  if (st.isDirectory()) {
    throw new Error(
      `Path is a directory, not a file: ${path}. Directory deletion is not supported.`,
    );
  }
  await rm(path, { force: false });
  return { path, deleted: true };
}
