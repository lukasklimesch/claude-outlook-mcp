<#
.SYNOPSIS
  Install the Outlook MCP Tool (Windows edition) and print the Claude Desktop
  configuration snippet.

.DESCRIPTION
  Verifies Bun is installed, installs dependencies, runs the test suite, and
  emits a ready-to-paste claude_desktop_config.json entry pointing at this
  checkout. Run from the repository root in PowerShell.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File install.ps1
#>
[CmdletBinding()]
param(
  [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "Installing Outlook MCP Tool (Windows edition)..." -ForegroundColor Green

# --- Bun ---------------------------------------------------------------
$bun = Get-Command bun -ErrorAction SilentlyContinue
if (-not $bun) {
  Write-Host "Bun is not installed. Install it with:" -ForegroundColor Red
  Write-Host '  powershell -c "irm bun.sh/install.ps1 | iex"' -ForegroundColor Yellow
  exit 1
}
$bunPath = $bun.Source
Write-Host "Found Bun: $bunPath" -ForegroundColor DarkGray

# --- Outlook presence (best effort) ------------------------------------
try {
  $ol = New-Object -ComObject Outlook.Application
  Write-Host "Found Microsoft Outlook version $($ol.Version)" -ForegroundColor DarkGray
  [Runtime.InteropServices.Marshal]::ReleaseComObject($ol) | Out-Null
} catch {
  Write-Host "WARNING: Could not start Outlook via COM. Make sure the Outlook desktop app is installed and configured." -ForegroundColor Yellow
}

# --- Dependencies ------------------------------------------------------
Write-Host "Installing dependencies..." -ForegroundColor Green
& bun install
if ($LASTEXITCODE -ne 0) { throw "bun install failed." }

# --- Tests -------------------------------------------------------------
if (-not $SkipTests) {
  Write-Host "Type-checking and running tests..." -ForegroundColor Green
  & bun run check
  if ($LASTEXITCODE -ne 0) { throw "Verification (tsc + tests) failed." }
}

# --- Config snippet ----------------------------------------------------
$indexPath = (Join-Path $root 'index.ts')
$config = @"
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "$($bunPath -replace '\\','\\\\')",
      "args": ["run", "$($indexPath -replace '\\','\\\\')"]
    }
  }
}
"@

$configFile = Join-Path $env:APPDATA 'Claude\claude_desktop_config.json'

Write-Host ""
Write-Host "Installation complete!" -ForegroundColor Green
Write-Host "Add the following to your Claude Desktop config at:" -ForegroundColor Yellow
Write-Host "  $configFile" -ForegroundColor Yellow
Write-Host ""
Write-Host $config -ForegroundColor Green
Write-Host ""
Write-Host "Then restart Claude Desktop. Tip: run scripts\smoke-test.ps1 first to validate Outlook access." -ForegroundColor Yellow
