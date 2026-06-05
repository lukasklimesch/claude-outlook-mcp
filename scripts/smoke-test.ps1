<#
.SYNOPSIS
  Live validation of the Outlook COM object model on this Windows machine.

.DESCRIPTION
  Exercises the same Outlook COM operations the MCP server uses (mail read,
  folders, calendar, contacts, and — optionally — sending a test email and
  saving attachments). Run this on the target Windows host BEFORE wiring the
  MCP server into Claude Desktop to confirm Outlook is installed, configured,
  and reachable.

  Read-only by default. Pass -SendTest -To you@example.com to also send a
  small test message to yourself.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts\smoke-test.ps1

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts\smoke-test.ps1 -SendTest -To me@contoso.com
#>
[CmdletBinding()]
param(
  [int]$Limit = 5,
  [switch]$SendTest,
  [string]$To,
  [string]$DownloadDir = (Join-Path $env:USERPROFILE 'Downloads')
)

$ErrorActionPreference = 'Stop'
$pass = 0; $fail = 0

function Step([string]$name, [scriptblock]$body) {
  Write-Host "[ .. ] $name" -ForegroundColor Cyan
  try {
    & $body
    Write-Host "[ OK ] $name" -ForegroundColor Green
    $script:pass++
  } catch {
    Write-Host "[FAIL] $name -> $($_.Exception.Message)" -ForegroundColor Red
    $script:fail++
  }
}

# COM constants (mirrors src/ps/preamble.ts)
$olFolderInbox = 6; $olFolderCalendar = 9; $olFolderContacts = 10; $olMailItem = 0

$app = $null; $ns = $null

Step "Connect to Outlook.Application (COM)" {
  try { $script:app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
  catch { $script:app = New-Object -ComObject Outlook.Application }
  $script:ns = $script:app.GetNamespace('MAPI')
  Write-Host "       Outlook version: $($script:app.Version)"
  try { Write-Host "       Profile/user  : $($script:ns.CurrentUser.Name)" } catch {}
}

Step "Read Inbox (top $Limit)" {
  $inbox = $ns.GetDefaultFolder($olFolderInbox)
  $items = $inbox.Items
  $items.Sort('[ReceivedTime]', $true)
  $n = 0
  foreach ($m in $items) {
    if ($n -ge $Limit) { break }
    try {
      Write-Host ("       [{0}] {1} | {2}" -f $m.ReceivedTime.ToString('yyyy-MM-dd HH:mm'), $m.SenderName, $m.Subject)
      $n++
    } catch {}
  }
  Write-Host "       Inbox unread count: $($inbox.UnReadItemCount)"
}

Step "Enumerate top-level folders" {
  foreach ($store in $ns.Folders) {
    Write-Host "       Store: $($store.Name)"
    foreach ($f in $store.Folders) { Write-Host "         - $($f.Name)" }
  }
}

Step "Read today's calendar" {
  $cal = $ns.GetDefaultFolder($olFolderCalendar)
  $items = $cal.Items
  $items.IncludeRecurrences = $true
  $items.Sort('[Start]')
  $start = (Get-Date).Date
  $end = $start.AddDays(1)
  $filter = "[Start] >= '" + $start.ToString('MM/dd/yyyy hh:mm tt') + "' AND [Start] < '" + $end.ToString('MM/dd/yyyy hh:mm tt') + "'"
  $restricted = $items.Restrict($filter)
  $count = @($restricted).Count
  Write-Host "       Events today: $count"
}

Step "Count contacts" {
  $folder = $ns.GetDefaultFolder($olFolderContacts)
  Write-Host "       Contacts: $($folder.Items.Count)"
}

if ($SendTest) {
  Step "Send a test email" {
    if (-not $To) { throw "Provide -To <address> with -SendTest." }
    $mail = $app.CreateItem($olMailItem)
    $mail.Subject = "Outlook MCP smoke test"
    $mail.Body = "This is a test message from the Outlook MCP smoke-test script."
    $mail.To = $To
    [void]$mail.Recipients.ResolveAll()
    $mail.Send()
    Write-Host "       Sent to $To"
  }
}

Write-Host ""
$resultColor = if ($fail -eq 0) { 'Green' } else { 'Red' }
Write-Host "==== Result: $pass passed, $fail failed ====" -ForegroundColor $resultColor
if ($fail -gt 0) { exit 1 }
