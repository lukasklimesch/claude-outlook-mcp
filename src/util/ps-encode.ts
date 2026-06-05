// ====================================================
// Safe encoding of user input for PowerShell execution
// ====================================================
//
// The single most important security/robustness property of this server is
// that *user-supplied data is never interpolated into PowerShell source*.
// Instead, every value (recipients, subject, body, file paths, search terms,
// ...) is serialized to a JSON payload and handed to PowerShell out-of-band.
// The script reads it back with `ConvertFrom-Json` and references typed
// properties (`$payload.subject`). This eliminates an entire class of
// script-injection bugs that the naive `replace(/"/g, '\\"')` approach in
// the macOS/AppleScript original is vulnerable to.

/** Environment variable carrying the inline JSON payload (small payloads). */
export const PAYLOAD_ENV = "OUTLOOK_MCP_PAYLOAD";

/** Environment variable carrying a path to a temp file with the payload. */
export const PAYLOAD_FILE_ENV = "OUTLOOK_MCP_PAYLOAD_FILE";

/** Override the PowerShell executable (e.g. point at a specific pwsh.exe). */
export const PWSH_ENV = "OUTLOOK_MCP_PWSH";

/**
 * Payloads larger than this (in characters of JSON) are written to a temp
 * file instead of an environment variable. A single Windows environment
 * variable tops out around 32 KiB; we stay comfortably under that and spill
 * large HTML bodies / many attachments to a file.
 */
export const MAX_INLINE_PAYLOAD = 30_000;

export interface EncodedPayload {
  /** Canonical JSON for the payload (always defined). */
  json: string;
  /** True when the payload should be passed via an env var; false => temp file. */
  inline: boolean;
}

/**
 * Serialize a payload object to JSON and decide whether it can be passed
 * inline via an environment variable or must spill to a temp file.
 */
export function encodePayload(
  payload: Record<string, unknown> | undefined,
  maxInline: number = MAX_INLINE_PAYLOAD,
): EncodedPayload {
  const json = JSON.stringify(payload ?? {});
  return { json, inline: json.length <= maxInline };
}

/**
 * Choose which PowerShell executable to invoke.
 *
 * Precedence:
 *   1. An explicit override (from {@link PWSH_ENV}).
 *   2. PowerShell 7+ (`pwsh.exe`) when detected on PATH — the modern engine.
 *   3. Windows PowerShell 5.1 (`powershell.exe`) — always present on Windows.
 *
 * Kept pure (no filesystem access) so it is trivially unit-testable; the
 * caller supplies the detection result.
 */
export function choosePowerShellCommand(opts: {
  override?: string | undefined;
  pwshAvailable?: boolean;
}): string {
  const override = opts.override?.trim();
  if (override) return override;
  if (opts.pwshAvailable) return "pwsh.exe";
  return "powershell.exe";
}

/** Standard non-interactive args for running a script file safely. */
export function powerShellArgs(scriptPath: string): string[] {
  return [
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    scriptPath,
  ];
}

/**
 * Build the environment overlay used to hand the payload to PowerShell.
 * Returns the variables to merge into the child process environment.
 */
export function payloadEnv(
  encoded: EncodedPayload,
  tempFilePath?: string,
): Record<string, string> {
  if (encoded.inline) {
    return { [PAYLOAD_ENV]: encoded.json };
  }
  if (!tempFilePath) {
    throw new Error(
      "Payload exceeds the inline limit but no temp file path was provided.",
    );
  }
  return { [PAYLOAD_FILE_ENV]: tempFilePath };
}
