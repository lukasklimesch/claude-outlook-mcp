import { describe, expect, test } from "bun:test";
import {
  choosePowerShellCommand,
  encodePayload,
  MAX_INLINE_PAYLOAD,
  PAYLOAD_ENV,
  PAYLOAD_FILE_ENV,
  payloadEnv,
  powerShellArgs,
} from "../src/util/ps-encode.js";

describe("encodePayload", () => {
  test("serializes to JSON and marks small payloads inline", () => {
    const e = encodePayload({ to: "a@b.com", limit: 5 });
    expect(JSON.parse(e.json)).toEqual({ to: "a@b.com", limit: 5 });
    expect(e.inline).toBe(true);
  });

  test("undefined payload becomes empty object json", () => {
    const e = encodePayload(undefined);
    expect(e.json).toBe("{}");
    expect(e.inline).toBe(true);
  });

  test("payloads over the limit are not inline", () => {
    const big = { body: "x".repeat(MAX_INLINE_PAYLOAD + 10) };
    expect(encodePayload(big).inline).toBe(false);
  });
});

describe("choosePowerShellCommand", () => {
  test("honors an explicit override first", () => {
    expect(
      choosePowerShellCommand({ override: "C:/ps/pwsh.exe", pwshAvailable: true }),
    ).toBe("C:/ps/pwsh.exe");
  });

  test("prefers pwsh when available", () => {
    expect(choosePowerShellCommand({ pwshAvailable: true })).toBe("pwsh.exe");
  });

  test("falls back to Windows PowerShell", () => {
    expect(choosePowerShellCommand({})).toBe("powershell.exe");
    expect(choosePowerShellCommand({ override: "   " })).toBe("powershell.exe");
  });
});

describe("powerShellArgs / payloadEnv", () => {
  test("non-interactive args reference the script path", () => {
    const args = powerShellArgs("C:/tmp/op.ps1");
    expect(args).toContain("-NoProfile");
    expect(args).toContain("-NonInteractive");
    expect(args[args.length - 1]).toBe("C:/tmp/op.ps1");
  });

  test("inline payload goes to the env var", () => {
    const env = payloadEnv(encodePayload({ a: 1 }));
    expect(env[PAYLOAD_ENV]).toBe('{"a":1}');
  });

  test("large payload is passed by file path", () => {
    const encoded = encodePayload({ body: "x".repeat(MAX_INLINE_PAYLOAD + 1) });
    const env = payloadEnv(encoded, "C:/tmp/payload.json");
    expect(env[PAYLOAD_FILE_ENV]).toBe("C:/tmp/payload.json");
    expect(env[PAYLOAD_ENV]).toBeUndefined();
  });

  test("large payload without a file path throws", () => {
    const encoded = encodePayload({ body: "x".repeat(MAX_INLINE_PAYLOAD + 1) });
    expect(() => payloadEnv(encoded)).toThrow();
  });
});
