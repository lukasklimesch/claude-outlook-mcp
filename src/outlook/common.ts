// Small helpers shared by the Outlook domain modules.

import type { PsResult } from "../types.js";

/**
 * Coerce a PowerShell `ConvertTo-Json` value into an array. Windows PowerShell
 * unwraps single-element arrays to a bare object and renders empty arrays as
 * null; this normalizes both back to a real array.
 */
export function asArray<T>(value: unknown): T[] {
  if (Array.isArray(value)) return value as T[];
  if (value === undefined || value === null) return [];
  return [value as T];
}

/** Return the data from a successful result, or throw with its error message. */
export function expectOk<T>(result: PsResult<T>): T {
  if (!result.ok) {
    throw new Error(result.error || "Outlook operation failed.");
  }
  return result.data as T;
}
