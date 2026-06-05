// ====================================================
// Mail operations (Outlook COM via PowerShell)
// ====================================================

import type { MailDetail, MailFolderInfo, MailSummary, OutlookRunner } from "../types.js";
import { buildScript } from "../ps/preamble.js";
import { asArray, expectOk } from "./common.js";
import type { MailArgs } from "../util/validate.js";

const DEFAULT_LIMIT = 10;

// --- PowerShell bodies (static; data flows through $payload at runtime) ---

const BODY_UNREAD = String.raw`
$folder = Resolve-MailFolder $ns $payload.folder
$limit = if ($payload.limit) { [int]$payload.limit } else { 10 }
$items = $folder.Items
$items.Sort('[ReceivedTime]', $true)
$unread = $items.Restrict('[UnRead] = True')
$result = @(); $count = 0
foreach ($m in $unread) {
  try { $result += (Convert-MailSummary $m 240); $count++ } catch {}
  if ($count -ge $limit) { break }
}
Write-McpResult -Obj @($result)
`;

const BODY_READ = String.raw`
$folder = Resolve-MailFolder $ns $payload.folder
$limit = if ($payload.limit) { [int]$payload.limit } else { 10 }
$items = $folder.Items
$items.Sort('[ReceivedTime]', $true)
$result = @(); $count = 0
foreach ($m in $items) {
  try { $result += (Convert-MailSummary $m 240); $count++ } catch {}
  if ($count -ge $limit) { break }
}
Write-McpResult -Obj @($result)
`;

const BODY_SEARCH = String.raw`
$folder = Resolve-MailFolder $ns $payload.folder
$limit = if ($payload.limit) { [int]$payload.limit } else { 10 }
$term = ([string]$payload.searchTerm).ToLower()
$items = $folder.Items
$items.Sort('[ReceivedTime]', $true)
$result = @(); $count = 0
foreach ($m in $items) {
  try {
    $hay = ''
    try { $hay = ([string]$m.Subject) + ' ' + ([string]$m.SenderName) + ' ' + ([string]$m.Body) } catch { $hay = [string]$m.Subject }
    if ($hay.ToLower().Contains($term)) { $result += (Convert-MailSummary $m 240); $count++ }
  } catch {}
  if ($count -ge $limit) { break }
}
Write-McpResult -Obj @($result)
`;

const BODY_GET = String.raw`
$item = $ns.GetItemFromID($payload.entryId)
Write-McpResult -Obj (Convert-MailDetail $item)
`;

const BODY_SEND = String.raw`
$mail = $app.CreateItem($olMailItem)
$mail.Subject = [string]$payload.subject
if ($payload.isHtml) { $mail.HTMLBody = [string]$payload.body } else { $mail.Body = [string]$payload.body }
$mail.To = [string]$payload.to
if ($payload.cc) { $mail.CC = [string]$payload.cc }
if ($payload.bcc) { $mail.BCC = [string]$payload.bcc }
if ($payload.attachments) { foreach ($p in $payload.attachments) { [void]$mail.Attachments.Add([string]$p) } }
[void]$mail.Recipients.ResolveAll()
$mail.Send()
Write-McpResult -Obj @{ status = 'sent'; to = [string]$payload.to; subject = [string]$payload.subject }
`;

const BODY_DRAFT = String.raw`
$mail = $app.CreateItem($olMailItem)
$mail.Subject = [string]$payload.subject
if ($payload.isHtml) { $mail.HTMLBody = [string]$payload.body } else { $mail.Body = [string]$payload.body }
$mail.To = [string]$payload.to
if ($payload.cc) { $mail.CC = [string]$payload.cc }
if ($payload.bcc) { $mail.BCC = [string]$payload.bcc }
if ($payload.attachments) { foreach ($p in $payload.attachments) { [void]$mail.Attachments.Add([string]$p) } }
$mail.Save()
Write-McpResult -Obj @{ status = 'draft'; entryId = [string]$mail.EntryID; subject = [string]$mail.Subject }
`;

const BODY_REPLY = String.raw`
$item = $ns.GetItemFromID($payload.entryId)
$reply = if ($payload.replyAll) { $item.ReplyAll() } else { $item.Reply() }
$new = [string]$payload.body
if ($payload.isHtml) { $reply.HTMLBody = $new + $reply.HTMLBody } else { $reply.Body = $new + [Environment]::NewLine + [Environment]::NewLine + $reply.Body }
[void]$reply.Recipients.ResolveAll()
$reply.Send()
Write-McpResult -Obj @{ status = 'sent'; replyAll = [bool]$payload.replyAll }
`;

const BODY_FORWARD = String.raw`
$item = $ns.GetItemFromID($payload.entryId)
$fwd = $item.Forward()
$fwd.To = [string]$payload.to
if ($payload.cc) { $fwd.CC = [string]$payload.cc }
if ($payload.body) {
  $new = [string]$payload.body
  if ($payload.isHtml) { $fwd.HTMLBody = $new + $fwd.HTMLBody } else { $fwd.Body = $new + [Environment]::NewLine + [Environment]::NewLine + $fwd.Body }
}
[void]$fwd.Recipients.ResolveAll()
$fwd.Send()
Write-McpResult -Obj @{ status = 'forwarded'; to = [string]$payload.to }
`;

const BODY_MOVE = String.raw`
$item = $ns.GetItemFromID($payload.entryId)
$target = Resolve-MailFolder $ns $payload.targetFolder
[void]$item.Move($target)
Write-McpResult -Obj @{ status = 'moved'; targetFolder = [string]$payload.targetFolder }
`;

const BODY_DELETE = String.raw`
$item = $ns.GetItemFromID($payload.entryId)
$item.Delete()
Write-McpResult -Obj @{ status = 'deleted' }
`;

const BODY_MARK = String.raw`
$item = $ns.GetItemFromID($payload.entryId)
$item.UnRead = -not [bool]$payload.read
$item.Save()
Write-McpResult -Obj @{ status = 'updated'; read = [bool]$payload.read }
`;

const BODY_FLAG = String.raw`
$item = $ns.GetItemFromID($payload.entryId)
if ($payload.flag) {
  $item.FlagRequest = 'Follow up'
  try { $item.FlagStatus = 2 } catch {}
} else {
  try { $item.ClearTaskFlag() } catch { $item.FlagRequest = '' }
}
$item.Save()
Write-McpResult -Obj @{ status = 'updated'; flag = [bool]$payload.flag }
`;

const BODY_FOLDERS = String.raw`
$result = New-Object System.Collections.ArrayList
function Add-FolderInfo($f, $path) {
  $full = if ($path) { "$path\$($f.Name)" } else { [string]$f.Name }
  $unread = 0; $total = 0
  try { $unread = [int]$f.UnReadItemCount } catch {}
  try { $total = [int]$f.Items.Count } catch {}
  [void]$result.Add([ordered]@{ name = [string]$f.Name; path = $full; unread = $unread; total = $total })
  foreach ($c in $f.Folders) { Add-FolderInfo $c $full }
}
foreach ($store in $ns.Folders) { Add-FolderInfo $store '' }
Write-McpResult -Obj @($result.ToArray())
`;

// --- TypeScript entry points -------------------------------------------

export async function getUnread(runner: OutlookRunner, args: MailArgs): Promise<MailSummary[]> {
  const res = await runner.run(buildScript(BODY_UNREAD), {
    folder: args.folder,
    limit: args.limit ?? DEFAULT_LIMIT,
  });
  return asArray<MailSummary>(expectOk(res));
}

export async function readMail(runner: OutlookRunner, args: MailArgs): Promise<MailSummary[]> {
  const res = await runner.run(buildScript(BODY_READ), {
    folder: args.folder,
    limit: args.limit ?? DEFAULT_LIMIT,
  });
  return asArray<MailSummary>(expectOk(res));
}

export async function searchMail(runner: OutlookRunner, args: MailArgs): Promise<MailSummary[]> {
  const res = await runner.run(buildScript(BODY_SEARCH), {
    folder: args.folder,
    limit: args.limit ?? DEFAULT_LIMIT,
    searchTerm: args.searchTerm,
  });
  return asArray<MailSummary>(expectOk(res));
}

export async function getMail(runner: OutlookRunner, args: MailArgs): Promise<MailDetail> {
  const res = await runner.run<MailDetail>(buildScript(BODY_GET), {
    entryId: args.entryId,
  });
  return expectOk(res);
}

export async function sendMail(runner: OutlookRunner, args: MailArgs): Promise<unknown> {
  const res = await runner.run(buildScript(BODY_SEND), {
    to: args.to,
    cc: args.cc,
    bcc: args.bcc,
    subject: args.subject,
    body: args.body,
    isHtml: args.isHtml ?? false,
    attachments: args.attachments,
  });
  return expectOk(res);
}

export async function createDraft(runner: OutlookRunner, args: MailArgs): Promise<unknown> {
  const res = await runner.run(buildScript(BODY_DRAFT), {
    to: args.to,
    cc: args.cc,
    bcc: args.bcc,
    subject: args.subject,
    body: args.body,
    isHtml: args.isHtml ?? false,
    attachments: args.attachments,
  });
  return expectOk(res);
}

export async function replyMail(runner: OutlookRunner, args: MailArgs): Promise<unknown> {
  const res = await runner.run(buildScript(BODY_REPLY), {
    entryId: args.entryId,
    body: args.body,
    isHtml: args.isHtml ?? false,
    replyAll: args.replyAll ?? false,
  });
  return expectOk(res);
}

export async function forwardMail(runner: OutlookRunner, args: MailArgs): Promise<unknown> {
  const res = await runner.run(buildScript(BODY_FORWARD), {
    entryId: args.entryId,
    to: args.to,
    cc: args.cc,
    body: args.body,
    isHtml: args.isHtml ?? false,
  });
  return expectOk(res);
}

export async function moveMail(runner: OutlookRunner, args: MailArgs): Promise<unknown> {
  const res = await runner.run(buildScript(BODY_MOVE), {
    entryId: args.entryId,
    targetFolder: args.targetFolder,
  });
  return expectOk(res);
}

export async function deleteMail(runner: OutlookRunner, args: MailArgs): Promise<unknown> {
  const res = await runner.run(buildScript(BODY_DELETE), { entryId: args.entryId });
  return expectOk(res);
}

export async function markMail(runner: OutlookRunner, args: MailArgs): Promise<unknown> {
  const res = await runner.run(buildScript(BODY_MARK), {
    entryId: args.entryId,
    read: args.read,
  });
  return expectOk(res);
}

export async function flagMail(runner: OutlookRunner, args: MailArgs): Promise<unknown> {
  const res = await runner.run(buildScript(BODY_FLAG), {
    entryId: args.entryId,
    flag: args.flag,
  });
  return expectOk(res);
}

export async function listFolders(runner: OutlookRunner): Promise<MailFolderInfo[]> {
  const res = await runner.run(buildScript(BODY_FOLDERS), {});
  return asArray<MailFolderInfo>(expectOk(res));
}
