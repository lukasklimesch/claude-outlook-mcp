// ====================================================
// Production PowerShell runner (Outlook COM bridge)
// ====================================================

import { spawn } from "node:child_process";
import { mkdtempSync, writeFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { randomUUID } from "node:crypto";
import type { OutlookRunner, PsResult } from "./types.js";
import {
  choosePowerShellCommand,
  encodePayload,
  payloadEnv,
  powerShellArgs,
  PWSH_ENV,
} from "./util/ps-encode.js";
import { parsePsOutput } from "./util/parse.js";

export interface PowerShellRunnerOptions {
  /** Explicit executable, else honored from env, else powershell.exe. */
  command?: string;
  /** Hard timeout for a single script execution. */
  timeoutMs?: number;
  /** Platform override (testing). Defaults to process.platform. */
  platform?: NodeJS.Platform;
}

const DEFAULT_TIMEOUT_MS = (() => {
  const fromEnv = Number(process.env.OUTLOOK_MCP_TIMEOUT_MS);
  return Number.isFinite(fromEnv) && fromEnv > 0 ? fromEnv : 60_000;
})();

/**
 * Runs generated PowerShell scripts that drive the Outlook COM object model
 * and returns parsed JSON. User input is delivered via environment variable
 * (or a temp file for large payloads) — never interpolated into the script.
 */
export class PowerShellRunner implements OutlookRunner {
  private readonly command: string;
  private readonly timeoutMs: number;
  private readonly platform: NodeJS.Platform;

  constructor(opts: PowerShellRunnerOptions = {}) {
    this.command =
      opts.command ??
      choosePowerShellCommand({ override: process.env[PWSH_ENV] });
    this.timeoutMs = opts.timeoutMs ?? DEFAULT_TIMEOUT_MS;
    this.platform = opts.platform ?? process.platform;
  }

  async run<T = unknown>(
    script: string,
    payload?: Record<string, unknown>,
  ): Promise<PsResult<T>> {
    if (this.platform !== "win32") {
      return {
        ok: false,
        category: "platform",
        error:
          "This MCP server requires Windows with Microsoft Outlook (desktop) " +
          `installed. Detected platform: ${this.platform}.`,
      } as PsResult<T>;
    }

    const dir = mkdtempSync(join(tmpdir(), "outlook-mcp-"));
    const scriptPath = join(dir, `op-${randomUUID()}.ps1`);
    const encoded = encodePayload(payload);
    let payloadFilePath: string | undefined;

    try {
      // UTF-8 with BOM so Windows PowerShell reads non-ASCII script text right.
      writeFileSync(scriptPath, "﻿" + script, { encoding: "utf8" });

      let env: Record<string, string>;
      if (encoded.inline) {
        env = payloadEnv(encoded);
      } else {
        payloadFilePath = join(dir, `payload-${randomUUID()}.json`);
        writeFileSync(payloadFilePath, encoded.json, { encoding: "utf8" });
        env = payloadEnv(encoded, payloadFilePath);
      }

      const { stdout, stderr, code, timedOut } = await this.spawn(
        scriptPath,
        env,
      );

      if (timedOut) {
        return {
          ok: false,
          category: "timeout",
          error: `Outlook operation timed out after ${this.timeoutMs}ms.`,
        } as PsResult<T>;
      }

      const result = parsePsOutput(stdout);
      // When we couldn't get a structured JSON result (no output, or junk on
      // stdout), prefer PowerShell's own stderr diagnostics — they carry the
      // actionable cause (e.g. the COM error) instead of truncated noise.
      if (
        !result.ok &&
        (result.category === "empty" || result.category === "parse") &&
        stderr.trim()
      ) {
        return {
          ok: false,
          category: code === 0 ? "outlook" : "spawn",
          error: stderr.trim(),
        } as PsResult<T>;
      }
      return result as PsResult<T>;
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }

  private spawn(
    scriptPath: string,
    extraEnv: Record<string, string>,
  ): Promise<{
    stdout: string;
    stderr: string;
    code: number | null;
    timedOut: boolean;
  }> {
    return new Promise((resolve) => {
      const child = spawn(this.command, powerShellArgs(scriptPath), {
        env: { ...process.env, ...extraEnv },
        windowsHide: true,
        stdio: ["ignore", "pipe", "pipe"],
      });

      let stdout = "";
      let stderr = "";
      let timedOut = false;

      const timer = setTimeout(() => {
        timedOut = true;
        child.kill();
      }, this.timeoutMs);

      child.stdout.on("data", (d) => (stdout += d.toString()));
      child.stderr.on("data", (d) => (stderr += d.toString()));

      child.on("error", (err) => {
        clearTimeout(timer);
        resolve({
          stdout: "",
          stderr: `Failed to launch PowerShell (${this.command}): ${err.message}`,
          code: null,
          timedOut,
        });
      });

      child.on("close", (code) => {
        clearTimeout(timer);
        resolve({ stdout, stderr, code, timedOut });
      });
    });
  }
}
