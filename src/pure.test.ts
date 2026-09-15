import { describe, expect, it } from 'vitest';
import {
  addDays,
  atHour,
  deriveTouchDates,
  eventDescription,
  eventIdFromCell,
  findColumn,
  fmtDate,
  parseDateValue,
  parseSettings,
  postTouchTitle,
  settingsRows,
  touchKindForColumn,
} from './pure';

describe('findColumn', () => {
  it('matches headers case-insensitively and trimmed', () => {
    const headers = ['Timestamp', 'Name', 'Sign-up Date', 'Status'];
    expect(findColumn(headers, 'name')).toBe(1);
    expect(findColumn(headers, '  NAME ')).toBe(1);
    expect(findColumn(headers, 'sign-up date')).toBe(2);
  });

  it('returns -1 for a missing header', () => {
    expect(findColumn(['A', 'B'], 'C')).toBe(-1);
  });
});

describe('parseDateValue', () => {
  it('accepts Date objects', () => {
    const d = new Date(2026, 2, 15);
    expect(parseDateValue(d)?.getTime()).toBe(d.getTime());
  });

  it('accepts ISO strings', () => {
    expect(fmtDate(parsed('2026-03-15'))).toBe('2026-03-15');
  });

  it('accepts m/d/y strings', () => {
    expect(fmtDate(parsed('3/15/2026'))).toBe('2026-03-15');
  });

  it('accepts epoch milliseconds', () => {
    expect(parsed(1773604800000).getTime()).toBe(1773604800000);
  });

  it('reads a numeric string as a Sheets date serial', () => {
    expect(fmtDate(parsed('46282'))).toBe('2026-09-17');
  });

  it('reads a number as a Sheets date serial', () => {
    expect(fmtDate(parsed(46282))).toBe('2026-09-17');
  });

  it('keeps the time-of-day in a fractional serial', () => {
    const d = parsed('46282.375');
    expect(fmtDate(d)).toBe('2026-09-17');
    expect(d.getHours()).toBe(9);
  });

  it('rejects results outside 1900–2100', () => {
    expect(parseDateValue('46282')).not.toBeNull(); // sanity: the bug case is in range
    expect(parseDateValue('5/15/2500')).toBeNull();
    expect(parseDateValue(1e12 * 60)).toBeNull(); // epoch-ms far in the future
  });

  it('rejects garbage and empty strings', () => {
    expect(parseDateValue('not a date')).toBeNull();
    expect(parseDateValue('')).toBeNull();
    expect(parseDateValue(null)).toBeNull();
  });

  it('ignores time-of-day when parsing date strings', () => {
    expect(fmtDate(parsed('2026-03-15T14:30:00'))).toBe('2026-03-15');
  });
});

/** parseDateValue for values the test expects to parse; throws instead of returning null. */
function parsed(value: unknown): Date {
  const d = parseDateValue(value);
  if (!d) throw new Error(`expected ${String(value)} to parse as a Date`);
  return d;
}

describe('deriveTouchDates', () => {
  const deadline = new Date(2026, 2, 16, 18, 0, 0); // Mar 16, 6pm local

  it('derives pre = deadline - 2 days at 9am and post = deadline + 3 days at 9am', () => {
    const { pre, post } = deriveTouchDates(deadline, 2, 3, 9);
    expect(fmtDate(pre)).toBe('2026-03-14');
    expect(pre.getHours()).toBe(9);
    expect(fmtDate(post)).toBe('2026-03-19');
    expect(post.getHours()).toBe(9);
  });

  it('does not mutate the input deadline', () => {
    const copy = new Date(deadline.getTime());
    deriveTouchDates(deadline, 2, 3, 9);
    expect(deadline.getTime()).toBe(copy.getTime());
  });
});

describe('time helpers', () => {
  it('addDays crosses month boundaries', () => {
    expect(fmtDate(addDays(new Date(2026, 1, 27), 3))).toBe('2026-03-02');
  });

  it('atHour zeroes minutes/seconds', () => {
    const d = atHour(new Date(2026, 2, 15, 9, 42, 13), 9);
    expect(d.getHours()).toBe(9);
    expect(d.getMinutes()).toBe(0);
    expect(d.getSeconds()).toBe(0);
  });
});

describe('parseSettings', () => {
  const defaults = settingsRows();

  it('reads the four knobs from the default rows', () => {
    expect(parseSettings(defaults).settings).toEqual({
      firstContactDays: 1,
      preDaysBefore: 2,
      postDaysAfter: 3,
      reminderHour: 9,
    });
  });

  it('reports a missing row by label', () => {
    const { settings, problems } = parseSettings(defaults.slice(1));
    expect(settings).toBeNull();
    expect(problems.join()).toContain('First contact');
  });

  it('reports blank and non-numeric values', () => {
    expect(parseSettings(withValue(0, '')).problems.join()).toMatch(/blank/);
    expect(parseSettings(withValue(0, 'soon')).problems.join()).toMatch(/must be a number/);
  });

  it('rejects fractional and negative day counts', () => {
    expect(parseSettings(withValue(0, 1.5)).problems.join()).toMatch(/whole number/);
    expect(parseSettings(withValue(1, -2)).problems.join()).toMatch(/negative/);
  });

  it('rejects a reminder hour outside 0–23', () => {
    expect(parseSettings(withValue(3, 24)).problems.join()).toMatch(/0 to 23/);
  });

  /** The default rows with one value replaced. */
  function withValue(index: number, value: unknown): unknown[][] {
    return defaults.map((row, i) => (i === index ? [row[0], value] : row));
  }
});

describe('event content', () => {
  const deadline = new Date(2026, 2, 16);

  it('builds touch titles, embedding the ask for post-touch', () => {
    expect(postTouchTitle('Jane Doe', 'Canvassing (base)')).toBe(
      'Post-touch: Jane Doe — Canvassing (base)',
    );
    expect(postTouchTitle('Jane Doe', '')).toBe('Post-touch: Jane Doe');
  });

  it('builds a post-touch description with context and purpose', () => {
    const desc = eventDescription({
      name: 'Jane Doe',
      email: 'jane@example.org',
      ask: 'Canvassing (base)',
      deadline,
      kind: 'post',
    });
    expect(desc).toContain('jane@example.org');
    expect(desc).toContain('Ask: Canvassing (base)');
    expect(desc).toContain('Deadline: 2026-03-16');
    expect(desc).toContain('Check in after the date');
  });

  it('builds a first-contact description', () => {
    const desc = eventDescription({
      name: 'Jane Doe',
      email: 'jane@example.org',
      deadline,
      kind: 'first-contact',
    });
    expect(desc).toContain('Reach out by: 2026-03-16');
    expect(desc).toContain('First contact');
  });

  it('labels email and phone as copy-ready contact lines for every kind', () => {
    for (const kind of ['first-contact', 'pre', 'post'] as const) {
      const desc = eventDescription({
        name: 'Jane Doe',
        email: 'jane@example.org',
        phone: '555-0100',
        deadline,
        kind,
      });
      expect(desc).toContain('Email: jane@example.org');
      expect(desc).toContain('Phone: 555-0100');
    }
  });

  it('omits blank email and phone lines instead of rendering empty labels', () => {
    const desc = eventDescription({
      name: 'Jane Doe',
      email: '  ',
      phone: '',
      kind: 'first-contact',
    });
    expect(desc).not.toContain('Email:');
    expect(desc).not.toContain('Phone:');
  });

  it('carries the About note as its own line when present', () => {
    const desc = eventDescription({
      name: 'Jane Doe',
      email: 'jane@example.org',
      phone: '555-0100',
      about: 'Loves dogs',
      deadline,
      kind: 'pre',
    });
    expect(desc).toContain('About: Loves dogs');
  });
});

describe('event id cells', () => {
  it('round-trips the raw id', () => {
    expect(eventIdFromCell('abc123')).toBe('abc123');
  });

  it('is empty for a missing id and trims whitespace', () => {
    expect(eventIdFromCell('')).toBe('');
    expect(eventIdFromCell('   ')).toBe('');
    expect(eventIdFromCell(' abc123 ')).toBe('abc123');
  });
});

describe('touchKindForColumn', () => {
  const headers = [
    'Name',
    'Email',
    'Ask',
    'Deadline',
    'Pre-touch',
    'Post-touch',
    'Pre-touch Event ID',
    'Post-touch Event ID',
  ];

  it('identifies pre/post columns by header text', () => {
    expect(touchKindForColumn(headers, 4)).toBe('pre');
    expect(touchKindForColumn(headers, 5)).toBe('post');
    expect(touchKindForColumn(headers, 0)).toBeNull();
    expect(touchKindForColumn(headers, 9)).toBeNull();
  });
});
