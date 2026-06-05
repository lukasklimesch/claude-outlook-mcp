#!/usr/bin/env bun
// ====================================================
// Outlook MCP Tool — Windows edition (entry point)
// ====================================================
//
// Drives the Microsoft Outlook desktop application on Windows through the
// Outlook COM object model (via PowerShell). All logic lives in ./src; this
// file just wires the production PowerShell runner to the MCP stdio transport.

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { PowerShellRunner } from "./src/powershell.js";
import { createServer, SERVER_NAME, SERVER_VERSION } from "./src/server.js";

async function main(): Promise<void> {
  console.error(`Starting ${SERVER_NAME} v${SERVER_VERSION}...`);

  if (process.platform !== "win32") {
    console.error(
      `[warning] Detected platform "${process.platform}". This server requires ` +
        "Windows with the Microsoft Outlook desktop app; tool calls will return " +
        "a platform error until run on Windows.",
    );
  }

  const runner = new PowerShellRunner();
  const server = createServer(runner);
  const transport = new StdioServerTransport();

  await server.connect(transport);
  console.error(`${SERVER_NAME} running on stdio.`);
}

main().catch((err) => {
  console.error("Fatal: failed to initialize Outlook MCP server:", err);
  process.exit(1);
});
