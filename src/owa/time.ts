/**
 * Date/time handling.
 *
 * - Values with an explicit offset ("Z", "+01:00") are absolute and sent as UTC.
 * - Values without an offset are wall-clock times. If an IANA time zone is given
 *   they are converted to UTC here; otherwise they are sent as-is and the server
 *   interprets them in the request's TimeZoneContext (the mailbox time zone).
 * - All reads use a UTC TimeZoneContext, so returned times end in "Z".
 */

const OFFSET_SUFFIX = /(Z|[+-]\d{2}:?\d{2})$/i;
const NAIVE_PATTERN = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?$/;

export interface DateTimeValue {
  /** Value to send to the service. */
  value: string;
  /** True if the value is wall-clock time that the server must interpret in a time zone. */
  floating: boolean;
}

export function hasExplicitOffset(input: string): boolean {
  return /T|\s/.test(input.trim()) && OFFSET_SUFFIX.test(input.trim());
}

export function toUtcString(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, "Z");
}

interface WallClock {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
}

function parseWallClock(input: string): WallClock | undefined {
  const match = NAIVE_PATTERN.exec(input.trim());
  if (!match) return undefined;
  const [, y, mo, d, h = "00", mi = "00", s = "00"] = match;
  const wall = { year: +y!, month: +mo!, day: +d!, hour: +h, minute: +mi, second: +s };
  const check = new Date(Date.UTC(wall.year, wall.month - 1, wall.day, wall.hour, wall.minute, wall.second));
  if (
    check.getUTCFullYear() !== wall.year ||
    check.getUTCMonth() !== wall.month - 1 ||
    check.getUTCDate() !== wall.day ||
    wall.hour > 23 ||
    wall.minute > 59 ||
    wall.second > 59
  ) {
    return undefined;
  }
  return wall;
}

function formatWallClock(w: WallClock): string {
  const p = (n: number, len = 2) => String(n).padStart(len, "0");
  return `${p(w.year, 4)}-${p(w.month)}-${p(w.day)}T${p(w.hour)}:${p(w.minute)}:${p(w.second)}`;
}

export function assertTimeZone(timeZone: string): void {
  try {
    new Intl.DateTimeFormat("en-US", { timeZone });
  } catch {
    throw new Error(`Unknown IANA time zone "${timeZone}" (expected e.g. "Europe/London")`);
  }
}

/** Offset of `timeZone` from UTC at the given instant, in milliseconds. */
function zoneOffset(instant: number, timeZone: string): number {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone,
    hourCycle: "h23",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  }).formatToParts(new Date(instant));
  const get = (type: string) => Number(parts.find((part) => part.type === type)?.value);
  const asUtc = Date.UTC(get("year"), get("month") - 1, get("day"), get("hour"), get("minute"), get("second"));
  return asUtc - Math.floor(instant / 1000) * 1000;
}

/** Convert wall-clock time in an IANA zone to a UTC instant. */
export function wallClockToUtc(w: WallClock, timeZone: string): Date {
  const guess = Date.UTC(w.year, w.month - 1, w.day, w.hour, w.minute, w.second);
  let instant = guess - zoneOffset(guess, timeZone);
  // Second pass corrects for a DST transition between the guess and the result.
  instant = guess - zoneOffset(instant, timeZone);
  return new Date(instant);
}

/**
 * Normalize a user-supplied date/time for a write request.
 * @param timeZone optional IANA zone for values without an offset.
 */
export function parseDateTimeInput(input: string, field: string, timeZone?: string): DateTimeValue {
  const text = input.trim();
  if (hasExplicitOffset(text)) {
    const date = new Date(text);
    if (Number.isNaN(date.getTime())) throw new Error(`${field}: "${input}" is not a valid date/time`);
    return { value: toUtcString(date), floating: false };
  }
  const wall = parseWallClock(text);
  if (!wall) {
    throw new Error(`${field}: "${input}" is not a valid ISO 8601 date/time (e.g. 2026-11-02T14:30:00 or 2026-11-02T14:30:00Z)`);
  }
  if (timeZone) {
    assertTimeZone(timeZone);
    return { value: toUtcString(wallClockToUtc(wall, timeZone)), floating: false };
  }
  return { value: formatWallClock(wall), floating: true };
}

/** Parse a date/time for a read range; values without an offset are treated as UTC. */
export function parseRangeInput(input: string, field: string, timeZone?: string): Date {
  const parsed = parseDateTimeInput(input, field, timeZone);
  return new Date(parsed.floating ? `${parsed.value}Z` : parsed.value);
}

/** A calendar date (YYYY-MM-DD) for all-day events; time components are rejected unless midnight. */
export function parseDateOnly(input: string, field: string): string {
  const wall = parseWallClock(input.trim().replace(/Z$/i, ""));
  if (!wall || wall.hour || wall.minute || wall.second) {
    throw new Error(`${field}: all-day events need a date like 2026-11-02, got "${input}"`);
  }
  return formatWallClock(wall).slice(0, 10);
}

export function addDays(date: string, days: number): string {
  const d = new Date(`${date}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}
