// ====================================================
// PowerShell preamble: Outlook COM helpers & contracts
// ====================================================
//
// This preamble is prepended to every operation script. It establishes a
// strict contract:
//   * Input arrives ONLY through `$payload` (parsed from JSON out-of-band),
//     so no user value is ever interpolated into PowerShell source.
//   * Output is exactly one JSON object on stdout: `{ok:true,data:...}` or
//     `{ok:false,error:...,category:...}` via Write-McpResult / Write-McpError.
//   * All Outlook COM access goes through typed converter helpers so the
//     shapes returned to TypeScript are stable.

export const OUTLOOK_PREAMBLE = String.raw`
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# ---- Outlook COM enumeration constants --------------------------------
$olFolderDeletedItems = 3
$olFolderOutbox       = 4
$olFolderSentMail     = 5
$olFolderInbox        = 6
$olFolderCalendar     = 9
$olFolderContacts     = 10
$olFolderDrafts       = 16
$olFolderJunk         = 23
$olMailItem           = 0
$olAppointmentItem    = 1
$olContactItem        = 2
$olMeeting            = 1
$olFormatPlain        = 1
$olFormatHTML         = 2

# ---- Input / output contract ------------------------------------------
function Get-McpPayload {
  $raw = $null
  if ($env:OUTLOOK_MCP_PAYLOAD_FILE -and (Test-Path -LiteralPath $env:OUTLOOK_MCP_PAYLOAD_FILE)) {
    $raw = Get-Content -Raw -LiteralPath $env:OUTLOOK_MCP_PAYLOAD_FILE
  } elseif ($env:OUTLOOK_MCP_PAYLOAD) {
    $raw = $env:OUTLOOK_MCP_PAYLOAD
  }
  if (-not $raw) { $raw = '{}' }
  $parsed = $raw | ConvertFrom-Json
  if ($null -eq $parsed) { $parsed = [pscustomobject]@{} }
  return $parsed
}

function Write-McpResult {
  param($Obj)
  $out = [ordered]@{ ok = $true; data = $Obj }
  $out | ConvertTo-Json -Depth 12 -Compress
}

function Write-McpError {
  param([string]$Message, [string]$Category = 'outlook')
  $out = [ordered]@{ ok = $false; error = $Message; category = $Category }
  $out | ConvertTo-Json -Depth 6 -Compress
}

# ---- Outlook application / namespace ----------------------------------
function Get-OutlookApp {
  try {
    return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
  } catch {
    return New-Object -ComObject Outlook.Application
  }
}

# ---- Folder resolution -------------------------------------------------
function Find-FolderRecursive($root, $name) {
  foreach ($f in $root.Folders) {
    if ($f.Name -ieq $name) { return $f }
    $sub = Find-FolderRecursive $f $name
    if ($sub) { return $sub }
  }
  return $null
}

function Resolve-MailFolder($ns, $name) {
  if (-not $name -or $name -ieq 'Inbox') { return $ns.GetDefaultFolder($olFolderInbox) }
  switch -regex ($name) {
    '^(?i)sent(\s*items)?$'      { return $ns.GetDefaultFolder($olFolderSentMail) }
    '^(?i)drafts?$'              { return $ns.GetDefaultFolder($olFolderDrafts) }
    '^(?i)deleted(\s*items)?$'   { return $ns.GetDefaultFolder($olFolderDeletedItems) }
    '^(?i)outbox$'               { return $ns.GetDefaultFolder($olFolderOutbox) }
    '^(?i)junk(\s*e\-?mail)?$'   { return $ns.GetDefaultFolder($olFolderJunk) }
  }
  foreach ($store in $ns.Folders) {
    $found = Find-FolderRecursive $store $name
    if ($found) { return $found }
  }
  throw "Mail folder not found: $name"
}

# ---- Value converters --------------------------------------------------
function Format-OutlookDate($d) {
  try { return $d.ToString('yyyy-MM-ddTHH:mm:ss') } catch { return [string]$d }
}

function Get-SafeFileName($name) {
  $n = [string]$name
  if (-not $n) { return 'attachment' }
  foreach ($ch in [IO.Path]::GetInvalidFileNameChars()) { $n = $n.Replace($ch, '_') }
  $n = $n.Trim()
  if (-not $n) { return 'attachment' }
  return $n
}

function Get-SenderSmtp($mail) {
  try {
    if ($mail.SenderEmailType -eq 'EX') {
      $sender = $mail.Sender
      if ($sender) {
        $ex = $sender.GetExchangeUser()
        if ($ex) { return [string]$ex.PrimarySmtpAddress }
      }
    }
  } catch {}
  try { return [string]$mail.SenderEmailAddress } catch { return '' }
}

function Convert-MailSummary($m, [int]$previewLen) {
  # Folders routinely contain non-MailItem entries (MeetingItem, ReportItem
  # receipts, PostItem) that lack some MailItem members. Guard every field so
  # any item type yields a best-effort summary instead of being dropped.
  $preview = ''
  if ($previewLen -gt 0) {
    try {
      $b = [string]$m.Body
      if ($b.Length -gt $previewLen) { $b = $b.Substring(0, $previewLen) }
      $preview = ($b -replace '\s+', ' ').Trim()
    } catch {}
  }
  $entryId = '';     try { $entryId = [string]$m.EntryID } catch {}
  $subject = '';     try { $subject = [string]$m.Subject } catch {}
  $senderName = '';  try { $senderName = [string]$m.SenderName } catch {}
  $senderEmail = ''; try { $senderEmail = Get-SenderSmtp $m } catch {}
  $to = '';          try { $to = [string]$m.To } catch {}
  $cc = '';          try { $cc = [string]$m.CC } catch {}
  $received = '';    try { $received = Format-OutlookDate $m.ReceivedTime } catch {}
  $unread = $false;  try { $unread = [bool]$m.UnRead } catch {}
  $size = 0;         try { $size = [int]$m.Size } catch {}
  $attCount = 0;     try { $attCount = [int]$m.Attachments.Count } catch {}
  [ordered]@{
    entryId         = $entryId
    subject         = $subject
    senderName      = $senderName
    senderEmail     = $senderEmail
    to              = $to
    cc              = $cc
    received        = $received
    unread          = $unread
    hasAttachments  = ($attCount -gt 0)
    attachmentCount = $attCount
    size            = $size
    preview         = $preview
  }
}

function Convert-Attachment($a, [int]$index) {
  [ordered]@{
    index    = $index
    fileName = [string]$a.FileName
    size     = [int]$a.Size
    type     = [int]$a.Type
  }
}

function Convert-MailDetail($m) {
  $summary = Convert-MailSummary $m 0
  $body = ''
  try { $body = [string]$m.Body } catch {}
  $atts = @()
  try {
    for ($i = 1; $i -le $m.Attachments.Count; $i++) {
      $atts += (Convert-Attachment $m.Attachments.Item($i) $i)
    }
  } catch {}
  $summary['body'] = $body
  $summary['attachments'] = @($atts)
  return $summary
}

function Convert-Event($e, [bool]$includeBody) {
  $h = [ordered]@{
    entryId  = [string]$e.EntryID
    subject  = [string]$e.Subject
    start    = (Format-OutlookDate $e.Start)
    end      = (Format-OutlookDate $e.End)
    location = [string]$e.Location
    isAllDay = [bool]$e.AllDayEvent
  }
  try { $h['organizer'] = [string]$e.Organizer } catch {}
  if ($includeBody) { try { $h['body'] = [string]$e.Body } catch {} }
  return $h
}

function Convert-Contact($c) {
  [ordered]@{
    entryId  = [string]$c.EntryID
    fullName = [string]$c.FullName
    email    = [string]$c.Email1Address
    phone    = [string]$c.BusinessTelephoneNumber
    company  = [string]$c.CompanyName
    jobTitle = [string]$c.JobTitle
  }
}
`;

/**
 * Wrap an operation body with the preamble and a uniform try/catch that
 * supplies `$payload`, `$app`, and `$ns`. The body must terminate by calling
 * `Write-McpResult -Obj <data>`; any thrown error is reported as JSON.
 */
export function buildScript(body: string): string {
  return `${OUTLOOK_PREAMBLE}
try {
  $payload = Get-McpPayload
  $app = Get-OutlookApp
  $ns = $app.GetNamespace('MAPI')
${body}
} catch {
  Write-McpError $_.Exception.Message
}
`;
}
