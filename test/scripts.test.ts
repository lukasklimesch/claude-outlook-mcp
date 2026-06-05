import { describe, expect, test } from "bun:test";
import { buildScript, OUTLOOK_PREAMBLE } from "../src/ps/preamble.js";

describe("Outlook PowerShell preamble", () => {
  test("declares the COM folder constants with correct values", () => {
    expect(OUTLOOK_PREAMBLE).toContain("$olFolderInbox        = 6");
    expect(OUTLOOK_PREAMBLE).toContain("$olFolderCalendar     = 9");
    expect(OUTLOOK_PREAMBLE).toContain("$olFolderContacts     = 10");
    expect(OUTLOOK_PREAMBLE).toContain("$olMailItem           = 0");
    expect(OUTLOOK_PREAMBLE).toContain("$olAppointmentItem    = 1");
    expect(OUTLOOK_PREAMBLE).toContain("$olContactItem        = 2");
  });

  test("defines the JSON I/O contract helpers", () => {
    expect(OUTLOOK_PREAMBLE).toContain("function Get-McpPayload");
    expect(OUTLOOK_PREAMBLE).toContain("function Write-McpResult");
    expect(OUTLOOK_PREAMBLE).toContain("function Write-McpError");
    expect(OUTLOOK_PREAMBLE).toContain("ConvertTo-Json");
    expect(OUTLOOK_PREAMBLE).toContain("ConvertFrom-Json");
  });

  test("defines converter and folder helpers", () => {
    expect(OUTLOOK_PREAMBLE).toContain("function Resolve-MailFolder");
    expect(OUTLOOK_PREAMBLE).toContain("function Convert-MailSummary");
    expect(OUTLOOK_PREAMBLE).toContain("function Convert-MailDetail");
    expect(OUTLOOK_PREAMBLE).toContain("function Convert-Event");
    expect(OUTLOOK_PREAMBLE).toContain("function Convert-Contact");
  });

  test("contains no JS-style interpolation or PowerShell backticks", () => {
    // Backticks would terminate the JS template literal; ${ would be a JS
    // interpolation. Their absence proves the PS source is fully static.
    expect(OUTLOOK_PREAMBLE.includes("`")).toBe(false);
    expect(OUTLOOK_PREAMBLE.includes("${")).toBe(false);
  });
});

describe("buildScript", () => {
  test("wraps a body with preamble, payload load and try/catch", () => {
    const script = buildScript("Write-McpResult -Obj @{ ok = 1 }");
    expect(script).toContain(OUTLOOK_PREAMBLE);
    expect(script).toContain("$payload = Get-McpPayload");
    expect(script).toContain("$app = Get-OutlookApp");
    expect(script).toContain("$ns = $app.GetNamespace('MAPI')");
    expect(script).toContain("Write-McpResult -Obj @{ ok = 1 }");
    expect(script).toContain("} catch {");
    expect(script).toContain("Write-McpError $_.Exception.Message");
  });
});
