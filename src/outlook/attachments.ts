// ====================================================
// Attachment operations — list & download (Outlook -> disk)
// ====================================================
//
// "Download files" in the Outlook sense: pull attachment bytes out of a
// message and persist them to disk via Attachment.SaveAsFile. The target
// directory is resolved on the TypeScript side (default: the user's
// Downloads folder) so behaviour is predictable and testable.

import type { AttachmentInfo, OutlookRunner } from "../types.js";
import { buildScript } from "../ps/preamble.js";
import { asArray, expectOk } from "./common.js";
import type { AttachmentArgs } from "../util/validate.js";
import { defaultDownloadDir, normalizeWindowsPath } from "../util/paths.js";

const BODY_LIST = String.raw`
$item = $ns.GetItemFromID($payload.entryId)
$result = @()
for ($i = 1; $i -le $item.Attachments.Count; $i++) {
  $result += (Convert-Attachment $item.Attachments.Item($i) $i)
}
Write-McpResult -Obj @($result)
`;

const BODY_DOWNLOAD = String.raw`
$item = $ns.GetItemFromID($payload.entryId)
$dir = [string]$payload.saveDir
if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
$saved = @()
$total = $item.Attachments.Count
for ($i = 1; $i -le $total; $i++) {
  $att = $item.Attachments.Item($i)
  if (($null -ne $payload.index) -and ([int]$payload.index -ne $i)) { continue }
  if ($payload.fileName -and ([string]$att.FileName -ne [string]$payload.fileName)) { continue }
  $name = Get-SafeFileName $att.FileName
  $dest = Join-Path $dir $name
  $att.SaveAsFile($dest)
  $saved += $dest
}
Write-McpResult -Obj @($saved)
`;

export async function listAttachments(runner: OutlookRunner, args: AttachmentArgs): Promise<AttachmentInfo[]> {
  const res = await runner.run(buildScript(BODY_LIST), { entryId: args.entryId });
  return asArray<AttachmentInfo>(expectOk(res));
}

export async function downloadAttachments(runner: OutlookRunner, args: AttachmentArgs): Promise<string[]> {
  const saveDir = normalizeWindowsPath(args.saveDir ?? defaultDownloadDir());
  const res = await runner.run(buildScript(BODY_DOWNLOAD), {
    entryId: args.entryId,
    saveDir,
    index: args.index,
    fileName: args.fileName,
  });
  return asArray<string>(expectOk(res));
}
