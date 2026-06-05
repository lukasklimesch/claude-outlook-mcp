// ====================================================
// Date normalization for Outlook appointments
// ====================================================
//
// Outlook stores appointment times as local "wall clock" values. We therefore
// preserve the local components the user supplied rather than converting to
// UTC (which would shift e.g. "2pm" by the timezone offset). The normalized
// string is parsed identically on the Windows host via [datetime]::Parse with
// the invariant culture.

function pad(n: number): string {
  return String(n).padStart(2, "0");
}

/**
 * Validate an incoming date/time string and return a canonical
 * `YYYY-MM-DDTHH:mm:ss` local representation.
 *
 * @throws Error if the value cannot be parsed into a real date.
 */
export function normalizeLocalDateTime(input: string): string {
  if (typeof input !== "string" || input.trim() === "") {
    throw new Error("A date/time value is required.");
  }
  let s = input.trim();
  // A bare date ("2026-06-05") is parsed by JS as UTC midnight, which shifts
  // to the previous/next local day under a non-zero offset. Pin it to local
  // midnight so the calendar day is preserved. (Strings carrying a time or an
  // explicit Z/offset are intentionally converted to local wall-clock, which
  // is what Outlook's local appointment times expect.)
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) s = `${s}T00:00:00`;
  const d = new Date(s);
  if (Number.isNaN(d.getTime())) {
    throw new Error(
      `Invalid date/time: "${input}". Use ISO 8601, e.g. 2026-06-05T14:00:00`,
    );
  }
  return (
    `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` +
    `T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`
  );
}

/**
 * Format a Date as the US-style string Outlook's `Items.Restrict` filter
 * expects (`MM/dd/yyyy hh:mm tt`). This format is the widely-documented,
 * broadly-compatible input for COM date restriction.
 */
export function toRestrictFormat(d: Date): string {
  let hours = d.getHours();
  const ampm = hours >= 12 ? "PM" : "AM";
  hours = hours % 12;
  if (hours === 0) hours = 12;
  return (
    `${pad(d.getMonth() + 1)}/${pad(d.getDate())}/${d.getFullYear()} ` +
    `${pad(hours)}:${pad(d.getMinutes())} ${ampm}`
  );
}

/** Start of the local day for a given date. */
export function startOfDay(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
}

/** Add whole days to a date, returning a new Date. */
export function addDays(d: Date, days: number): Date {
  const copy = new Date(d.getTime());
  copy.setDate(copy.getDate() + days);
  return copy;
}
