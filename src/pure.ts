/**
 * Pure, testable logic — no Google Apps Script globals here.
 * Everything in this file can be unit-tested with plain Vitest.
 */
import { CONFIG, type TouchKind } from './config';

/** Case-insensitive, trim-normalized comparison form of a header label. */
export function norm(value: string): string {
  return String(value ?? '')
    .trim()
    .toLowerCase();
}

/** 0-based index of a header, or -1 when absent. */
export function findColumn(headers: string[], wanted: string): number {
  const target = norm(wanted);
  return headers.findIndex((h) => norm(h) === target);
}

/** 1900 date system: day 0 = 1899-12-30 local midnight. */
const SERIAL_EPOCH = new Date(1899, 11, 30);

/** Parse a cell value into a Date, or null when it isn't one. */
export function parseDateValue(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return new Date(value.getTime());
  if (typeof value === 'number' && Number.isFinite(value)) return fromNumber(value);
  if (typeof value === 'string') {
    const t = value.trim();
    if (!t) return null;
    // Numeric-only strings come from the Sheets date picker and are date
    // serials, not years — `new Date(t)` would read them as `+046282-01-01`.
    if (/^-?\d+(?:\.\d+)?$/.test(t)) return fromNumber(Number(t));
    // Date-only ISO strings (and ISO datetimes) parse as UTC; a date *cell*
    // means the local calendar date, so build a local-midnight Date instead.
    const iso = /^(\d{4})-(\d{2})-(\d{2})(?:[T ].*)?$/.exec(t);
    const [, year, month, day] = iso ?? [];
    if (year && month && day) return inBound(new Date(+year, +month - 1, +day));
    const d = new Date(t);
    if (!Number.isNaN(d.getTime())) return inBound(d);
  }
  return null;
}

/** A number is either a Sheets date serial or epoch milliseconds. */
function fromNumber(value: number): Date | null {
  const days = Math.floor(value);
  const frac = value - days;
  const serial = new Date(SERIAL_EPOCH.getTime());
  serial.setDate(serial.getDate() + days);
  serial.setTime(serial.getTime() + Math.round(frac * 86400000));
  if (inBound(serial)) return serial;
  // Serials for 1900–2100 are ≤ ~73400; large values are epoch milliseconds.
  if (Math.abs(value) >= 1e11) return inBound(new Date(value));
  return null;
}

/** Dates outside 1900–2100 mean the input wasn't a real date value. */
function inBound(d: Date): Date | null {
  const y = d.getFullYear();
  return y >= 1900 && y <= 2100 ? d : null;
}

export function addDays(date: Date, days: number): Date {
  const d = new Date(date.getTime());
  d.setDate(d.getDate() + days);
  return d;
}

export function atHour(date: Date, hour: number): Date {
  const d = new Date(date.getTime());
  d.setHours(hour, 0, 0, 0);
  return d;
}

export interface TouchDates {
  pre: Date;
  post: Date;
}

/** The coordinator's timing knobs, read from the Settings tab. */
export interface Settings {
  firstContactDays: number;
  preDaysBefore: number;
  postDaysAfter: number;
  reminderHour: number;
}

/** Header row for the Settings tab. */
export function settingsHeaders(): string[] {
  return [CONFIG.settingsColumns.label, CONFIG.settingsColumns.value];
}

/** The label/default pairs seeded into a fresh Settings tab. */
export function settingsRows(): [string, number][] {
  const s = CONFIG.settings;
  return [
    [s.firstContactDays.label, s.firstContactDays.default],
    [s.preDaysBefore.label, s.preDaysBefore.default],
    [s.postDaysAfter.label, s.postDaysAfter.default],
    [s.reminderHour.label, s.reminderHour.default],
  ];
}

/**
 * Read the four knobs from Settings label/value rows. Returns null and a
 * plain-English problem list when anything is missing or not a usable
 * number, so callers can either report it at Set Up or refuse to proceed.
 */
export function parseSettings(rows: unknown[][]): {
  settings: Settings | null;
  problems: string[];
} {
  const problems: string[] = [];
  const numberFor = (label: string, kind: 'days' | 'hour'): number | null => {
    const row = rows.find((r) => norm(String(r[0] ?? '')) === norm(label));
    if (!row) {
      problems.push(`${label} is missing`);
      return null;
    }
    const raw = row[1];
    if (raw === '' || raw === null || raw === undefined) {
      problems.push(`${label} is blank`);
      return null;
    }
    const value = typeof raw === 'number' ? raw : Number(String(raw).trim());
    if (!Number.isFinite(value)) {
      problems.push(`${label} must be a number (found "${String(raw)}")`);
      return null;
    }
    if (!Number.isInteger(value)) {
      problems.push(`${label} must be a whole number (found "${String(raw)}")`);
      return null;
    }
    if (kind === 'days' && value < 0) {
      problems.push(`${label} cannot be negative (found "${String(raw)}")`);
      return null;
    }
    if (kind === 'hour' && (value < 0 || value > 23)) {
      problems.push(`${label} must be an hour from 0 to 23 (found "${String(raw)}")`);
      return null;
    }
    return value;
  };

  const s = CONFIG.settings;
  const firstContactDays = numberFor(s.firstContactDays.label, 'days');
  const preDaysBefore = numberFor(s.preDaysBefore.label, 'days');
  const postDaysAfter = numberFor(s.postDaysAfter.label, 'days');
  const reminderHour = numberFor(s.reminderHour.label, 'hour');
  if (
    firstContactDays === null ||
    preDaysBefore === null ||
    postDaysAfter === null ||
    reminderHour === null
  ) {
    return { settings: null, problems };
  }
  return {
    settings: { firstContactDays, preDaysBefore, postDaysAfter, reminderHour },
    problems: [],
  };
}

/** Pre/post touch dates derived from a Deadline. */
export function deriveTouchDates(
  deadline: Date,
  preDays: number,
  postDays: number,
  hour: number,
): TouchDates {
  return {
    pre: atHour(addDays(deadline, -preDays), hour),
    post: atHour(addDays(deadline, postDays), hour),
  };
}

/** yyyy-MM-dd for display and event descriptions. */
export function fmtDate(date: Date): string {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

// ---- Event content ------------------------------------------------

export function firstContactTitle(name: string): string {
  return `First contact: ${name}`;
}

export function preTouchTitle(name: string): string {
  return `Pre-touch: ${name}`;
}

export function postTouchTitle(name: string, ask: string): string {
  return ask ? `Post-touch: ${name} — ${ask}` : `Post-touch: ${name}`;
}

export function touchTitle(kind: TouchKind, name: string, ask?: string): string {
  switch (kind) {
    case 'first-contact':
      return firstContactTitle(name);
    case 'pre':
      return preTouchTitle(name);
    case 'post':
      return postTouchTitle(name, ask ?? '');
  }
}

export interface EventDetails {
  name: string;
  email?: string;
  phone?: string;
  about?: string;
  ask?: string;
  deadline?: Date | null;
  kind: TouchKind;
}

/**
 * Name plus labeled email/phone lines. The labels matter: Google Calendar
 * linkifies them, so the coordinator gets tap-to-call/text from the event.
 */
function contactLines(d: Pick<EventDetails, 'name' | 'email' | 'phone'>): string[] {
  const lines: string[] = [d.name];
  const email = String(d.email ?? '').trim();
  const phone = String(d.phone ?? '').trim();
  if (email) lines.push(`Email: ${email}`);
  if (phone) lines.push(`Phone: ${phone}`);
  return lines;
}

export function eventDescription(d: EventDetails): string {
  const lines = contactLines(d);
  const about = String(d.about ?? '').trim();
  if (about) lines.push(`About: ${about}`);
  if (d.kind === 'pre' || d.kind === 'post') {
    if (d.ask) lines.push(`Ask: ${d.ask}`);
    if (d.deadline) lines.push(`Deadline: ${fmtDate(d.deadline)}`);
    lines.push(
      d.kind === 'pre'
        ? 'Reach out before the date — confirm, encourage, thank.'
        : 'Check in after the date — how did it go? Thank them, invite to community events.',
    );
  } else if (d.deadline) {
    lines.push(`Reach out by: ${fmtDate(d.deadline)}`);
    lines.push('First contact for a new volunteer.');
  } else {
    lines.push('First contact for a new volunteer.');
  }
  return lines.join('\n');
}

/** Which touch kind a Tracker column (0-based index) corresponds to, if any. */
export function touchKindForColumn(
  trackerHeaders: string[],
  colIndex0Based: number,
): TouchKind | null {
  const header = trackerHeaders[colIndex0Based];
  if (header === undefined) return null;
  if (norm(header) === norm(CONFIG.trackerColumns.preTouch)) return 'pre';
  if (norm(header) === norm(CONFIG.trackerColumns.postTouch)) return 'post';
  return null;
}

/** Calendar event id stored in a cell; '' when the cell is empty. */
export function eventIdFromCell(cell: string): string {
  return String(cell ?? '').trim();
}

// ---- Sheet layouts -------------------------------------------------

/** Ordered header row for the Tracker tab. */
export function trackerHeaders(): string[] {
  const c = CONFIG.trackerColumns;
  return [
    c.name,
    c.email,
    c.ask,
    c.deadline,
    c.preTouch,
    c.postTouch,
    c.preTouchEvent,
    c.postTouchEvent,
    c.notes,
    c.done,
  ];
}

/** Directory roster headers: the declared fields the script mirrors, then the intake fields. */
export function directoryHeaders(): string[] {
  const { name, email, phone } = CONFIG.requiredFields;
  const { title: about } = CONFIG.optionalFields.about;
  const { reachOutBy, firstContactEvent } = CONFIG.directoryColumns;
  return [name, email, phone, about, reachOutBy, firstContactEvent];
}
