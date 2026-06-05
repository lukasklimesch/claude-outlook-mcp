// ====================================================
// Contacts operations (Outlook COM via PowerShell)
// ====================================================

import type { ContactInfo, OutlookRunner } from "../types.js";
import { buildScript } from "../ps/preamble.js";
import { asArray, expectOk } from "./common.js";
import type { ContactArgs } from "../util/validate.js";

const DEFAULT_LIMIT = 25;

const BODY_LIST = String.raw`
$folder = $ns.GetDefaultFolder($olFolderContacts)
$limit = if ($payload.limit) { [int]$payload.limit } else { 25 }
$items = $folder.Items
$items.Sort('[FileAs]')
$result = @(); $count = 0
foreach ($c in $items) {
  try {
    $h = Convert-Contact $c
    if ($h.fullName) { $result += $h; $count++ }
  } catch {}
  if ($count -ge $limit) { break }
}
Write-McpResult -Obj @($result)
`;

const BODY_SEARCH = String.raw`
$folder = $ns.GetDefaultFolder($olFolderContacts)
$limit = if ($payload.limit) { [int]$payload.limit } else { 25 }
$term = ([string]$payload.searchTerm).ToLower()
$items = $folder.Items
$result = @(); $count = 0
foreach ($c in $items) {
  try {
    $h = Convert-Contact $c
    $hay = (([string]$h.fullName) + ' ' + ([string]$h.email) + ' ' + ([string]$h.company)).ToLower()
    if ($h.fullName -and $hay.Contains($term)) { $result += $h; $count++ }
  } catch {}
  if ($count -ge $limit) { break }
}
Write-McpResult -Obj @($result)
`;

const BODY_CREATE = String.raw`
$c = $app.CreateItem($olContactItem)
$c.FullName = [string]$payload.fullName
if ($payload.email) { $c.Email1Address = [string]$payload.email }
if ($payload.phone) { $c.BusinessTelephoneNumber = [string]$payload.phone }
if ($payload.company) { $c.CompanyName = [string]$payload.company }
if ($payload.jobTitle) { $c.JobTitle = [string]$payload.jobTitle }
$c.Save()
Write-McpResult -Obj (Convert-Contact $c)
`;

export async function listContacts(runner: OutlookRunner, args: ContactArgs): Promise<ContactInfo[]> {
  const res = await runner.run(buildScript(BODY_LIST), { limit: args.limit ?? DEFAULT_LIMIT });
  return asArray<ContactInfo>(expectOk(res));
}

export async function searchContacts(runner: OutlookRunner, args: ContactArgs): Promise<ContactInfo[]> {
  const res = await runner.run(buildScript(BODY_SEARCH), {
    searchTerm: args.searchTerm,
    limit: args.limit ?? DEFAULT_LIMIT,
  });
  return asArray<ContactInfo>(expectOk(res));
}

export async function createContact(runner: OutlookRunner, args: ContactArgs): Promise<ContactInfo> {
  const res = await runner.run<ContactInfo>(buildScript(BODY_CREATE), {
    fullName: args.fullName,
    email: args.email,
    phone: args.phone,
    company: args.company,
    jobTitle: args.jobTitle,
  });
  return expectOk(res);
}
