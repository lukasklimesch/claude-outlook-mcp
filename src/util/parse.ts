// ====================================================
// Robust parsing of PowerShell stdout into PsResult
// ====================================================
//
// The macOS original parsed AppleScript records with a fragile
// `/\{([^}]+)\}/g` regex. We instead emit real JSON from PowerShell
// (`ConvertTo-Json`) and parse it here, with defensive fallbacks for the
// occasional case where a warning or progress line leaks onto stdout.

import type { PsResult } from "../types.js";

function truncate(s: string, n: number): string {
  return s.length > n ? s.slice(0, n) + "…" : s;
}

function tryJson(text: string): unknown | undefined {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

/**
 * Return the balanced JSON span starting at `start` (whose char is `{` or
 * `[`), or undefined if it never balances. String contents (including escaped
 * quotes) are skipped so braces inside strings don't confuse the counter.
 */
function balancedSpan(text: string, start: number): string | undefined {
  const open = text[start];
  const close = open === "{" ? "}" : "]";
  let depth = 0;
  let inString = false;
  let escaped = false;

  for (let i = start; i < text.length; i++) {
    const ch = text[i];
    if (inString) {
      if (escaped) escaped = false;
      else if (ch === "\\") escaped = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') inString = true;
    else if (ch === open) depth++;
    else if (ch === close) {
      depth--;
      if (depth === 0) return text.slice(start, i + 1);
    }
  }
  return undefined;
}

/**
 * Recover a JSON value embedded in noisy output (e.g. a PowerShell warning or
 * progress line preceding the result). Every `{`/`[` is tried as a start; the
 * LONGEST substring that parses as valid JSON wins, so a stray bracket in the
 * preamble text cannot hijack parsing of the real result object.
 */
export function extractTrailingJson(text: string): unknown | undefined {
  let best: unknown;
  let bestLen = 0;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch !== "{" && ch !== "[") continue;
    const span = balancedSpan(text, i);
    if (span === undefined || span.length <= bestLen) continue;
    const parsed = tryJson(span);
    if (parsed !== undefined) {
      best = parsed;
      bestLen = span.length;
    }
  }
  return best;
}

/**
 * Normalize an already-parsed value into a {@link PsResult}.
 *
 * PowerShell's `ConvertTo-Json` collapses a single-element array to a bare
 * object, so we accept either an object with an `ok` field or wrap an
 * arbitrary value as a successful result.
 */
function normalize(value: unknown): PsResult {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    const obj = value as Record<string, unknown>;
    if (typeof obj.ok === "boolean") {
      return {
        ok: obj.ok,
        data: obj.data,
        error: typeof obj.error === "string" ? obj.error : undefined,
        category: typeof obj.category === "string" ? obj.category : undefined,
      };
    }
  }
  return { ok: true, data: value };
}

/**
 * Parse raw PowerShell stdout into a structured {@link PsResult}.
 * Never throws — malformed output becomes `{ ok: false, category: "parse" }`.
 */
export function parsePsOutput(stdout: string | null | undefined): PsResult {
  const trimmed = (stdout ?? "").replace(/^﻿/, "").trim();
  if (!trimmed) {
    return {
      ok: false,
      error:
        "No output received from PowerShell. Is Microsoft Outlook installed and is this a Windows host?",
      category: "empty",
    };
  }

  const direct = tryJson(trimmed);
  if (direct !== undefined) return normalize(direct);

  const extracted = extractTrailingJson(trimmed);
  if (extracted !== undefined) return normalize(extracted);

  return {
    ok: false,
    error: `Could not parse Outlook response as JSON: ${truncate(trimmed, 600)}`,
    category: "parse",
  };
}
