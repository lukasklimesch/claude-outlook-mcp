// ====================================================
// Windows path handling helpers
// ====================================================

/**
 * Normalize a Windows path: strip surrounding quotes/whitespace and convert
 * forward slashes to backslashes. Leaves UNC and drive-letter prefixes intact.
 */
export function normalizeWindowsPath(p: string): string {
  if (typeof p !== "string") return p as unknown as string;
  let out = p.trim();
  if (
    (out.startsWith('"') && out.endsWith('"')) ||
    (out.startsWith("'") && out.endsWith("'"))
  ) {
    out = out.slice(1, -1);
  }
  return out.replace(/\//g, "\\");
}

/**
 * Resolve the default directory for saving downloaded attachments. Honors an
 * explicit override, then `OUTLOOK_MCP_DOWNLOAD_DIR`, then the user's
 * Downloads folder, then the current working directory.
 */
export function defaultDownloadDir(
  env: Record<string, string | undefined> = process.env,
): string {
  const override = env.OUTLOOK_MCP_DOWNLOAD_DIR?.trim();
  if (override) return normalizeWindowsPath(override);
  const userProfile = env.USERPROFILE?.trim();
  if (userProfile) return normalizeWindowsPath(`${userProfile}\\Downloads`);
  return ".";
}
