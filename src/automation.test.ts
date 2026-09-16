import { beforeEach, describe, expect, it } from 'vitest';
import {
  columnWidthFor,
  createWorld,
  type MockCalendarEvent,
  MockValidationBuilder,
  type World,
} from '../tests/mocks/gas';
import type { EditEventLike, FormSubmitEventLike, Services, SheetLike } from './automation';
import {
  actionOptions,
  createActionFromDialog,
  findFormResponseSheet,
  handleEdit,
  handleFormSubmit,
  openActionDialog,
  runSetupAutomation,
  showSignUpForm,
  validateRequiredHeaders,
} from './automation';
import { CONFIG } from './config';
import { atHour, fmtDate, trackerHeaders } from './pure';
import { buildServices } from './services';

function headersOf(sheet: SheetLike): string[] {
  const [row = []] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  return row.map(String);
}

function colIndex(sheet: SheetLike, header: string): number {
  const i = headersOf(sheet).findIndex(
    (h) => h.trim().toLowerCase() === header.trim().toLowerCase(),
  );
  return i + 1; // 1-based, 0 when missing
}

/** Calendar event at `index`; throws instead of returning undefined when absent. */
function eventAt(w: World, index: number): MockCalendarEvent {
  const ev = w.calendar.events[index];
  if (!ev) throw new Error(`expected a calendar event at index ${index}`);
  return ev;
}

function edit(
  w: World,
  tab: string,
  row: number,
  header: string,
  value: unknown,
  oldValue: unknown = '',
  services: Services = w.services,
): void {
  const sheet = w.byName(tab);
  const col = colIndex(sheet, header);
  const e: EditEventLike = { source: w.ss, range: sheet.getRange(row, col), value, oldValue };
  handleEdit(services, e);
}

/**
 * Services as a trigger execution sees them: no editor session, so the UI and
 * `Spreadsheet.toast` both throw, exactly as Apps Script does live.
 */
function withoutUi(w: World): { services: Services; logs: string[] } {
  w.ss.toast = () => {
    throw new Error('You do not have permission to call toast');
  };
  const logs: string[] = [];
  const services = buildServices({
    getActiveSpreadsheet: () => w.ss,
    getUi: () => {
      throw new Error('Cannot call SpreadsheetApp.getUi() from this context');
    },
    calendars: w.calendars,
    getScript: () => w.script,
    newDataValidation: () => w.services.makeValidation(),
    forms: w.forms,
    showHtml: () => {
      throw new Error('Cannot call SpreadsheetApp.getUi() from this context');
    },
    flush: () => {},
    log: (message) => logs.push(message),
  });
  return { services, logs };
}

/** Simulate the Form writing a row, then run the submit handler. Returns the
 * response tab's grid as it stood *before* the handler ran. */
function submitRow(
  w: World,
  row: number,
  values: unknown[],
  services: Services = w.services,
): unknown[][] {
  const sheet = w.byName('Form Responses 1');
  // mirror the real behavior: onFormSubmit fires after the form wrote the row
  values.forEach((v, i) => {
    sheet.getRange(row, i + 1).setValue(v);
  });
  const before = structuredClone(sheet.grid);
  const e: FormSubmitEventLike = { source: w.ss, range: sheet.getRange(row, 1, 1, values.length) };
  handleFormSubmit(services, e);
  return before;
}

const deadline = new Date(2026, 2, 15); // Mar 15 2026
const pre = new Date(2026, 2, 13); // −2 days
const post = new Date(2026, 2, 18); // +3 days

const FORM_HEADERS = ['Timestamp', 'Name', 'Email', 'Phone', 'About'];
const DIRECTORY_HEADERS = ['Name', 'Email', 'Phone', 'About', 'Reach out by', 'Calendar Event ID'];

function settingsSpec(rows?: unknown[][]) {
  return {
    name: 'Settings',
    headers: ['Setting', 'Value'],
    rows: rows ?? [
      ['First contact (days after signup)', 1],
      ['Pre-touch (days before deadline)', 2],
      ['Post-touch (days after deadline)', 3],
      ['Reminder hour (24h)', 9],
    ],
  };
}

/** Set one Settings value by its label text. */
function setSetting(w: World, label: string, value: unknown): void {
  const settings = w.byName('Settings');
  const rows = settings.getRange(1, 1, settings.getLastRow(), 2).getValues();
  const index = rows.findIndex(([l]) => String(l).trim() === label);
  if (index < 0) throw new Error(`no Settings row labelled "${label}"`);
  settings.getRange(index + 1, 2).setValue(value);
}

/** Blank the Settings row that records the Campaign Calendar. */
function clearCampaignCalendar(w: World): void {
  setSetting(w, CONFIG.calendar.settingLabel, '');
}

function trackerWorld(): World {
  return createWorld([
    {
      name: 'Form Responses 1',
      headers: FORM_HEADERS,
      rows: [['2026-02-01 09:00:00', 'Jane', 'jane@example.org', '555-0100', 'Loves dogs']],
    },
    {
      name: 'Directory',
      headers: DIRECTORY_HEADERS,
      rows: [['Jane', 'jane@example.org', '555-0100', 'Loves dogs', '', '']],
    },
    {
      name: 'Menu of Asks',
      headers: ['Ask'],
      rows: [['Canvassing (base)'], ['Phone banking']],
    },
    settingsSpec(),
    {
      name: 'Tracker',
      headers: [
        'Name',
        'Email',
        'Ask',
        'Deadline',
        'Pre-touch',
        'Post-touch',
        'Pre-touch Event ID',
        'Post-touch Event ID',
        'Notes',
        'Done',
      ],
      rows: [['Jane', 'jane@example.org', 'Phone banking', '', '', '', '', '', '', '']],
    },
  ]);
}

describe('deadline lifecycle (onEdit → calendar events)', () => {
  let w: World;
  beforeEach(() => {
    w = trackerWorld();
  });

  it('derives pre/post into cells and creates exactly two events', () => {
    edit(w, 'Tracker', 2, 'Deadline', deadline);

    const tracker = w.byName('Tracker');
    expect(fmtDate(tracker.getRange(2, colIndex(tracker, 'Pre-touch')).getValue() as Date)).toBe(
      fmtDate(pre),
    );
    expect(fmtDate(tracker.getRange(2, colIndex(tracker, 'Post-touch')).getValue() as Date)).toBe(
      fmtDate(post),
    );

    expect(w.calendar.events).toHaveLength(2);
    const preEv = eventAt(w, 0);
    const postEv = eventAt(w, 1);
    expect(preEv.title).toBe('Pre-touch: Jane');
    expect(preEv.start.getTime()).toBe(atHour(pre, CONFIG.settings.reminderHour.default).getTime());
    expect(postEv.title).toBe('Post-touch: Jane — Phone banking');
    expect(postEv.start.getTime()).toBe(
      atHour(post, CONFIG.settings.reminderHour.default).getTime(),
    );
    expect(postEv.description).toContain('jane@example.org');
    expect(postEv.description).toContain('Ask: Phone banking');
  });

  it('records event ids so later edits move, not duplicate', () => {
    edit(w, 'Tracker', 2, 'Deadline', deadline);
    const tracker = w.byName('Tracker');
    const storedPreId = String(
      tracker.getRange(2, colIndex(tracker, 'Pre-touch Event ID')).getValue(),
    );
    const storedPostId = String(
      tracker.getRange(2, colIndex(tracker, 'Post-touch Event ID')).getValue(),
    );
    expect(storedPreId).toMatch(/^ev/);
    expect(storedPostId).toMatch(/^ev/);

    // Re-edit to a later deadline
    const later = new Date(2026, 3, 1);
    edit(w, 'Tracker', 2, 'Deadline', later);

    expect(w.calendar.events).toHaveLength(2);
    const preEv = eventAt(w, 0);
    const postEv = eventAt(w, 1);
    expect(fmtDate(preEv.start)).toBe('2026-03-30');
    expect(fmtDate(postEv.start)).toBe('2026-04-04');
    expect(String(tracker.getRange(2, colIndex(tracker, 'Pre-touch Event ID')).getValue())).toBe(
      storedPreId,
    );
    expect(String(tracker.getRange(2, colIndex(tracker, 'Post-touch Event ID')).getValue())).toBe(
      storedPostId,
    );
  });

  it("resolves the volunteer's contact details and About note from the Directory", () => {
    edit(w, 'Tracker', 2, 'Deadline', deadline);

    for (const description of [eventAt(w, 0).description, eventAt(w, 1).description]) {
      expect(description).toContain('Email: jane@example.org');
      expect(description).toContain('Phone: 555-0100');
      expect(description).toContain('About: Loves dogs');
    }
  });

  it('clearing the deadline removes both events and derived cells', () => {
    edit(w, 'Tracker', 2, 'Deadline', deadline);
    edit(w, 'Tracker', 2, 'Deadline', '');

    expect(w.calendar.events).toHaveLength(0);
    const tracker = w.byName('Tracker');
    expect(tracker.getRange(2, colIndex(tracker, 'Pre-touch')).getValue()).toBe('');
    expect(tracker.getRange(2, colIndex(tracker, 'Post-touch')).getValue()).toBe('');
    expect(String(tracker.getRange(2, colIndex(tracker, 'Pre-touch Event ID')).getValue())).toBe(
      '',
    );
  });

  it('refuses a deadline with no Ask — reverts the cell, no events', () => {
    const w2 = createWorld([
      { name: 'Directory', headers: ['Name', 'Email'], rows: [['Jane', 'jane@example.org']] },
      { name: 'Menu of Asks', headers: ['Ask'], rows: [['Phone banking']] },
      settingsSpec(),
      {
        name: 'Tracker',
        headers: [
          'Name',
          'Email',
          'Ask',
          'Deadline',
          'Pre-touch',
          'Post-touch',
          'Pre-touch Event ID',
          'Post-touch Event ID',
        ],
        rows: [['Jane', 'jane@example.org', '', '', '', '', '', '']],
      },
    ]);
    edit(w2, 'Tracker', 2, 'Deadline', deadline, '2026-01-01');

    expect(w2.calendar.events).toHaveLength(0);
    const tracker = w2.byName('Tracker');
    expect(String(tracker.getRange(2, colIndex(tracker, 'Deadline')).getValue())).toBe(
      '2026-01-01',
    );
    expect(w2.logs.some((m) => m.includes('Choose an Ask first'))).toBe(true);
  });

  it('keeps the row displaying dates as yyyy-mm-dd after a deadline edit', () => {
    edit(w, 'Tracker', 2, 'Deadline', deadline);

    const tracker = w.byName('Tracker');
    for (const h of ['Deadline', 'Pre-touch', 'Post-touch']) {
      expect(tracker.getNumberFormat(2, colIndex(tracker, h))).toBe('yyyy-mm-dd');
    }
  });

  it('rejects garbage date input by reverting instead of clearing', () => {
    edit(w, 'Tracker', 2, 'Deadline', 'not a date', '2026-02-01');
    expect(w.calendar.events).toHaveLength(0);
    const tracker = w.byName('Tracker');
    expect(String(tracker.getRange(2, colIndex(tracker, 'Deadline')).getValue())).toBe(
      '2026-02-01',
    );
  });

  it('still creates reminders, minus contact lines, when no Directory row matches', () => {
    const w = trackerWorld();
    // A Directory row with a different email — the tracker volunteer is unmatched.
    const directory = w.byName('Directory');
    directory.getRange(2, colIndex(directory, 'Email')).setValue('other@example.org');

    edit(w, 'Tracker', 2, 'Deadline', deadline);

    expect(w.calendar.events).toHaveLength(2);
    for (const description of [eventAt(w, 0).description, eventAt(w, 1).description]) {
      expect(description).toContain('Jane');
      expect(description).not.toContain('Phone:');
      expect(description).not.toContain('About:');
    }
  });

  it('omits a blank phone but keeps the About note', () => {
    const w = trackerWorld();
    const directory = w.byName('Directory');
    directory.getRange(2, colIndex(directory, 'Phone')).setValue('');

    edit(w, 'Tracker', 2, 'Deadline', deadline);

    const description = eventAt(w, 0).description;
    expect(description).not.toContain('Phone:');
    expect(description).toContain('About: Loves dogs');
  });
});

describe('manual pre/post touch edits', () => {
  it('relocates the pre-touch event when the pre-touch cell is edited directly', () => {
    const w = trackerWorld();
    edit(w, 'Tracker', 2, 'Deadline', deadline);
    expect(w.calendar.events).toHaveLength(2);

    const custom = new Date(2026, 2, 10);
    edit(w, 'Tracker', 2, 'Pre-touch', custom);

    expect(w.calendar.events).toHaveLength(2);
    const preEv = eventAt(w, 0);
    expect(preEv.title).toBe('Pre-touch: Jane');
    expect(preEv.start.getTime()).toBe(
      atHour(custom, CONFIG.settings.reminderHour.default).getTime(),
    );
  });

  it('carries the Directory contact details when a touch is relocated by hand', () => {
    const w = trackerWorld();
    edit(w, 'Tracker', 2, 'Deadline', deadline);

    edit(w, 'Tracker', 2, 'Pre-touch', new Date(2026, 2, 10));

    const description = eventAt(w, 0).description;
    expect(description).toContain('Email: jane@example.org');
    expect(description).toContain('Phone: 555-0100');
    expect(description).toContain('About: Loves dogs');
  });
});

describe('first contact (Directory "Reach out by")', () => {
  it('creates one event and moves it on re-edit, removes on clear', () => {
    const w = trackerWorld();
    const d1 = new Date(2026, 2, 17);
    edit(w, 'Directory', 2, 'Reach out by', d1);

    expect(w.calendar.events).toHaveLength(1);
    expect(eventAt(w, 0).title).toBe('First contact: Jane');
    expect(fmtDate(eventAt(w, 0).start)).toBe('2026-03-17');

    edit(w, 'Directory', 2, 'Reach out by', new Date(2026, 2, 18));
    expect(w.calendar.events).toHaveLength(1);
    expect(fmtDate(eventAt(w, 0).start)).toBe('2026-03-18');

    edit(w, 'Directory', 2, 'Reach out by', '');
    expect(w.calendar.events).toHaveLength(0);
  });

  it('reads a Sheets date-picker serial as a date and moves the event', () => {
    const w = trackerWorld();
    edit(w, 'Directory', 2, 'Reach out by', new Date(2026, 2, 17));
    expect(w.calendar.events).toHaveLength(1);

    // The Sheets date picker surfaces the cell as a serial like "46282".
    edit(w, 'Directory', 2, 'Reach out by', '46282');
    expect(w.calendar.events).toHaveLength(1);
    expect(fmtDate(eventAt(w, 0).start)).toBe('2026-09-17');
  });

  it('carries the Directory contact details and About note when a reach-out is moved', () => {
    const w = trackerWorld();

    edit(w, 'Directory', 2, 'Reach out by', new Date(2026, 2, 17));

    const description = eventAt(w, 0).description;
    expect(description).toContain('Email: jane@example.org');
    expect(description).toContain('Phone: 555-0100');
    expect(description).toContain('About: Loves dogs');
  });

  it('moves and clears the event intake already scheduled for a signup', () => {
    const w = trackerWorld();
    submitRow(w, 3, [
      '2026-02-02 10:00:00',
      'New Person',
      'new@example.org',
      '555-0199',
      'Excited',
    ]);
    expect(w.calendar.events).toHaveLength(1);
    const scheduled = eventAt(w, 0).id;

    edit(w, 'Directory', 3, 'Reach out by', new Date(2026, 2, 10));
    expect(w.calendar.events).toHaveLength(1);
    expect(eventAt(w, 0).id).toBe(scheduled);
    expect(fmtDate(eventAt(w, 0).start)).toBe('2026-03-10');

    edit(w, 'Directory', 3, 'Reach out by', '');
    expect(w.calendar.events).toHaveLength(0);
  });
});

describe('response tab naming', () => {
  it('renames the freshly linked response tab to "Sign-ups"', () => {
    const w = createWorld([]);
    w.ui.queuePrompt('https://forms.example/sample', 'ok');
    runSetupAutomation(w.services);
    expect(colIndex(w.byName('Sign-ups'), 'Timestamp')).toBe(1);
  });

  it('on re-run names the linked form URL and keeps it when cancelled', () => {
    const w = createWorld([]);
    w.ui.queuePrompt('https://forms.example/sample', 'ok');
    runSetupAutomation(w.services);
    expect(w.ui.prompts).toHaveLength(1);

    w.ui.queuePrompt('', 'cancel'); // decline the replace
    const promptsBefore = w.ui.prompts.length;
    runSetupAutomation(w.services);

    expect(w.ui.prompts).toHaveLength(promptsBefore + 1); // the single replace prompt
    expect(w.ui.prompts[promptsBefore]?.text).toContain('https://forms.example/sample');
    expect(w.ss.getFormUrl()).toBe('https://forms.example/sample');
    expect(colIndex(w.byName('Sign-ups'), 'Timestamp')).toBe(1);
    expect(w.ss.getSheetByName('Form Responses 1')).toBeNull();
  });

  it('validates the linked response tab, not a stale unlinked one', () => {
    const w = createWorld([
      // Old form's responses, unlinked since the form changed: no Phone column.
      { name: 'Form Responses 1', headers: ['Timestamp', 'Name', 'Email'], rows: [] },
    ]);
    w.forms.questions = ['Name', 'Email', 'Phone', 'About'];
    w.ui.queuePrompt('https://forms.example/sample', 'ok');
    runSetupAutomation(w.services);
    const last = w.ui.alerts[w.ui.alerts.length - 1];
    expect(last?.text).not.toContain('Missing required field');
  });

  it('on replace, archives the raw responses before unlinking', () => {
    const w = createWorld([]);
    w.forms.questions = ['Name', 'Email', 'Phone', 'About'];
    w.ui.queuePrompt('https://forms.example/sample', 'ok');
    runSetupAutomation(w.services);
    const responses = w.byName('Sign-ups');
    responses.appendRow([
      '2026-02-02 10:00:00',
      'Jane',
      'jane@example.org',
      '555-0100',
      'Loves dogs',
    ]);
    const before = responses.getRange(1, 1, 2, 5).getValues();

    // one replace prompt handles the archive decision and the URL in one step
    w.ui.queuePrompt('https://forms.example/new', 'ok');
    const promptsBefore = w.ui.prompts.length;
    runSetupAutomation(w.services);

    expect(w.ui.prompts).toHaveLength(promptsBefore + 1); // replace = a single prompt
    expect(w.ui.alerts.some((a) => a.buttons === w.ui.ButtonSet.OK_CANCEL)).toBe(false);
    const archive = w.byName('Sign-ups archive');
    expect(archive.getRange(1, 1, 2, 5).getValues()).toEqual(before);
    expect(w.ss.getSheetByName('Form Responses 1')).toBeNull(); // old tab unlinked
    expect(w.ss.getSheetByName('Form Responses 2')).toBeNull(); // fresh tab never keeps the ugly name
    expect(w.ss.getFormUrl()).toBe('https://forms.example/new');
    expect(colIndex(w.byName('Sign-ups'), 'Timestamp')).toBe(1);
  });
});

describe('form generation (Set Up creates the intake form)', () => {
  it('creates a linked intake form when the prompt is answered OK with no URL', () => {
    const w = createWorld([]);
    w.ui.queuePrompt('', 'ok');

    runSetupAutomation(w.services);

    expect(w.forms.created).not.toBeNull();
    expect(w.ss.getFormUrl()).toBe(w.forms.createdFormUrl);
    expect(w.byName('Sign-ups')).toBeTruthy();
  });

  it('takes its questions from CONFIG — titles and required flags — and passes validation', () => {
    const w = createWorld([]);
    w.ui.queuePrompt('', 'ok');

    runSetupAutomation(w.services);

    const created = w.forms.created;
    if (!created) throw new Error('no form was created');
    expect(created.title).toBe(CONFIG.intakeFormTitle);
    expect(created.questions).toEqual([
      { title: CONFIG.requiredFields.name, required: true, type: 'text', helpText: '' },
      { title: CONFIG.requiredFields.email, required: true, type: 'text', helpText: '' },
      { title: CONFIG.requiredFields.phone, required: true, type: 'text', helpText: '' },
      {
        title: CONFIG.optionalFields.about.title,
        required: false,
        type: 'paragraph',
        helpText: CONFIG.optionalFields.about.helpText,
      },
    ]);
    // The generated response tab carries exactly Timestamp + the questions,
    // so the Directory's single source of truth flows through untouched.
    expect(headersOf(w.byName('Sign-ups'))).toEqual([
      CONFIG.timestampField,
      ...created.questions.map((q) => q.title),
    ]);
    const responses = findFormResponseSheet(w.ss);
    if (!responses) throw new Error('no response sheet');
    expect(validateRequiredHeaders(responses)).toEqual([]);
  });

  it('names the generated form URL in the Set Up summary', () => {
    const w = createWorld([]);
    w.ui.queuePrompt('', 'ok');

    runSetupAutomation(w.services);

    const last = w.ui.alerts[w.ui.alerts.length - 1];
    const url = w.forms.createdFormUrl ?? '';
    expect(last?.text).toContain(url);
  });

  it('still links a pasted URL without creating a form, and explains the create-or-paste choice', () => {
    const w = createWorld([]);
    w.ui.queuePrompt('https://forms.example/sample', 'ok');

    runSetupAutomation(w.services);

    expect(w.forms.created).toBeNull();
    expect(w.ss.getFormUrl()).toBe('https://forms.example/sample');
    expect(w.byName('Sign-ups')).toBeTruthy();
    expect(w.ui.prompts[0]?.text).toContain('Drive');
  });

  it('does not create or link anything when the prompt is cancelled', () => {
    const w = createWorld([]);
    w.ui.queuePrompt('', 'cancel');

    runSetupAutomation(w.services);

    expect(w.forms.created).toBeNull();
    expect(w.ss.getFormUrl()).toBeNull();
    expect(w.ss.getSheets().some((s) => s.getFormUrl() !== null)).toBe(false);
  });

  it('on replace, archives the old responses and generates the new form', () => {
    const w = createWorld([]);
    w.forms.questions = ['Name', 'Email', 'Phone', 'About'];
    w.ui.queuePrompt('https://forms.example/sample', 'ok');
    runSetupAutomation(w.services);
    w.byName('Sign-ups').appendRow([
      '2026-02-02 10:00:00',
      'Jane',
      'jane@example.org',
      '555-0100',
      'Loves dogs',
    ]);

    // one replace prompt handles the archive decision and the URL in one step
    w.ui.queuePrompt('', 'ok');
    const promptsBefore = w.ui.prompts.length;
    runSetupAutomation(w.services);

    expect(w.ui.prompts).toHaveLength(promptsBefore + 1); // replace = a single prompt
    expect(w.ui.alerts.some((a) => a.buttons === w.ui.ButtonSet.OK_CANCEL)).toBe(false);
    expect(w.forms.created).not.toBeNull();
    expect(w.ss.getFormUrl()).toBe(w.forms.createdFormUrl);
    expect(w.byName('Sign-ups archive').getLastRow()).toBe(2); // old responses archived
    expect(colIndex(w.byName('Sign-ups'), 'Timestamp')).toBe(1);
    // the replace prompt must say the replacement can be generated, not only pasted
    expect(w.ui.prompts[promptsBefore]?.text).toContain('generate');
  });

  it('cancel on the replace prompt keeps the linked form and its responses untouched', () => {
    const w = createWorld([]);
    w.forms.questions = ['Name', 'Email', 'Phone', 'About'];
    w.ui.queuePrompt('https://forms.example/sample', 'ok');
    runSetupAutomation(w.services);
    const responses = w.byName('Sign-ups');
    responses.appendRow([
      '2026-02-02 10:00:00',
      'Jane',
      'jane@example.org',
      '555-0100',
      'Loves dogs',
    ]);
    const before = responses.getRange(1, 1, 2, 5).getValues();

    w.ui.queuePrompt('', 'cancel');
    const promptsBefore = w.ui.prompts.length;
    runSetupAutomation(w.services);

    expect(w.ui.prompts).toHaveLength(promptsBefore + 1);
    expect(w.ui.alerts.some((a) => a.buttons === w.ui.ButtonSet.OK_CANCEL)).toBe(false);
    expect(w.forms.created).toBeNull();
    expect(w.forms.unlinkedUrls).toEqual([]);
    expect(w.ss.getFormUrl()).toBe('https://forms.example/sample');
    expect(responses.getRange(1, 1, 2, 5).getValues()).toEqual(before);
    expect(w.ss.getSheetByName('Sign-ups archive')).toBeNull();
  });

  it('leaves a legacy unlinked response tab alone when generating a new form', () => {
    const w = createWorld([
      {
        name: 'Form Responses 1',
        headers: ['Timestamp', 'Name', 'Email', 'Phone', 'About'],
        rows: [['2026-02-02 10:00:00', 'Jane', 'jane@example.org', '555-0100', 'Loves dogs']],
      },
    ]);
    w.ui.queuePrompt('', 'ok'); // generate straight away — no already-linked dialog

    runSetupAutomation(w.services);

    expect(w.forms.created).not.toBeNull();
    // nothing to archive or delete: the stale tab was never form-linked
    expect(w.ss.getSheetByName('Sign-ups archive')).toBeNull();
    expect(w.ss.getSheetByName('Form Responses 1')?.getLastRow()).toBe(2);
    // the rename lands on the fresh linked tab, not the stale lookalike
    expect(w.byName('Sign-ups').getFormUrl()).not.toBeNull();
    expect(colIndex(w.byName('Sign-ups'), 'Timestamp')).toBe(1);
  });
});

describe('the link-form prompt', () => {
  it('asks the same question whether or not a form is already linked', () => {
    const fresh = createWorld([]);
    fresh.ui.queuePrompt('', 'cancel');
    runSetupAutomation(fresh.services);

    const linked = createWorld([]);
    linked.ui.queuePrompt('https://forms.example/sample', 'ok');
    runSetupAutomation(linked.services);
    linked.ui.queuePrompt('', 'cancel');
    runSetupAutomation(linked.services);

    const freshPrompt = fresh.ui.prompts[0];
    const replacePrompt = linked.ui.prompts[1];
    expect(freshPrompt?.title).toBe('Link form or generate new one?');
    expect(replacePrompt?.title).toBe(freshPrompt?.title);

    // the shared half of the dialog is identical in both cases
    const shared = 'Press OK with this box empty and a sign-up form will be generated';
    expect(freshPrompt?.text).toContain(shared);
    expect(replacePrompt?.text).toContain(shared);
    // only the replacing dialog names the form it would replace
    expect(replacePrompt?.text).toContain('https://forms.example/sample');
    expect(freshPrompt?.text).not.toContain('already linked');
  });

  it('lists the fields a sign-up form needs, and the optional About question', () => {
    const w = createWorld([]);
    w.ui.queuePrompt('', 'cancel');
    runSetupAutomation(w.services);

    const text = w.ui.prompts[0]?.text ?? '';
    expect(text).toContain(Object.values(CONFIG.requiredFields).join(', '));
    expect(text).toContain(CONFIG.optionalFields.about.title);
  });
});

describe('showSignUpForm', () => {
  it('does nothing when no form is linked', () => {
    const w = createWorld([]);

    showSignUpForm(w.services);

    expect(w.ui.dialogs).toHaveLength(0);
  });

  it('shows a linked form as a new-tab link', () => {
    const w = createWorld([]);
    w.ss.formUrl = 'https://forms.example/sample';

    showSignUpForm(w.services);

    expect(w.ui.dialogs).toHaveLength(1);
    expect(w.ui.dialogs[0]?.title).toBe('Sign-Up Form');
    expect(w.ui.dialogs[0]?.html).toContain('href="https://forms.example/sample"');
    expect(w.ui.dialogs[0]?.html).toContain('target="_blank"');
  });
});

describe('runSetupAutomation', () => {
  it('leaves Form Responses untouched and creates a separate Directory roster', () => {
    const w = createWorld([
      {
        name: 'Form Responses 1',
        headers: FORM_HEADERS,
        rows: [['2026-02-01 09:00:00', 'Jane', 'jane@example.org', '555-0100', 'Loves dogs']],
      },
    ]);
    w.ui.queuePrompt('https://forms.example/sample', 'ok');
    const form = w.byName('Form Responses 1');
    const before = { grid: structuredClone(form.grid), lastColumn: form.getLastColumn() };

    runSetupAutomation(w.services);

    // the script never writes to the Form's response tab
    expect(form.grid).toEqual(before.grid);
    expect(form.getLastColumn()).toBe(before.lastColumn);
    expect(form.getFrozenRows()).toBe(0);

    // Directory is a separate roster carrying the declared fields and intake fields
    expect(headersOf(w.byName('Directory'))).toEqual(DIRECTORY_HEADERS);
  });

  it('freezes the Directory header row', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    expect(w.byName('Directory').getFrozenRows()).toBe(1);
  });

  it('freezes the Tracker header row', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    expect(w.byName('Tracker').getFrozenRows()).toBe(1);
  });

  it('freezes the Menu of Asks header row', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    expect(w.byName('Menu of Asks').getFrozenRows()).toBe(1);
  });

  it('freezes the Settings header row', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    expect(w.byName('Settings').getFrozenRows()).toBe(1);
  });

  it('styles the Directory header row for scanning', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    const header = w.byName('Directory').getStyle(1, 1);
    expect(header.fontWeight).toBe('bold');
    expect(header.background).toBe(CONFIG.sheetFormat.headerBackground);
    expect(header.fontColor).toBe(CONFIG.sheetFormat.headerFontColor);
  });

  it('wraps the Directory and caps a runaway event-id column', () => {
    const longId = 'x'.repeat(80);
    const w = createWorld([
      { name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] },
      {
        name: 'Directory',
        headers: DIRECTORY_HEADERS,
        rows: [['Jane', 'jane@example.org', '555-0100', '', '', longId]],
      },
    ]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    const directory = w.byName('Directory');
    const eventIdCol = DIRECTORY_HEADERS.indexOf('Calendar Event ID') + 1;

    expect(directory.getStyle(2, 1).wrap).toBe(true);
    expect(directory.getColumnWidth(eventIdCol)).toBe(CONFIG.sheetFormat.maxColumnWidth);
    expect(directory.getColumnWidth(1)).toBeLessThan(CONFIG.sheetFormat.maxColumnWidth);
  });

  it('fits a Directory column to its longest value, not its wrapped fragment', () => {
    const email = 'jane@example.org';
    const w = createWorld([
      { name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] },
      {
        name: 'Directory',
        headers: DIRECTORY_HEADERS,
        rows: [['Jane', email, '555-0100', '', '', '']],
      },
    ]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    const emailCol = DIRECTORY_HEADERS.indexOf('Email') + 1;
    expect(w.byName('Directory').getColumnWidth(emailCol)).toBe(columnWidthFor(email));
  });

  it('mutes the Directory machine column', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    const directory = w.byName('Directory');
    const eventIdCol = DIRECTORY_HEADERS.indexOf('Calendar Event ID') + 1;
    expect(directory.getStyle(2, eventIdCol).fontColor).toBe(CONFIG.sheetFormat.mutedFontColor);
    expect(directory.getStyle(2, eventIdCol).fontSize).toBe(CONFIG.sheetFormat.mutedFontSize);
    expect(directory.getStyle(2, 1).fontColor).toBeUndefined();
  });

  it('centers the Directory date column', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    const directory = w.byName('Directory');
    const reachCol = DIRECTORY_HEADERS.indexOf('Reach out by') + 1;
    expect(directory.getStyle(2, reachCol).horizontalAlignment).toBe('center');
    expect(directory.getStyle(2, 1).horizontalAlignment).toBeUndefined();
  });

  it('styles the Tracker: muted event ids, centered dates and Done', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    const tracker = w.byName('Tracker');
    const col = (header: string) => trackerHeaders().indexOf(header) + 1;

    expect(tracker.getStyle(1, 1).fontWeight).toBe('bold');
    expect(tracker.getStyle(2, col('Pre-touch Event ID')).fontColor).toBe(
      CONFIG.sheetFormat.mutedFontColor,
    );
    expect(tracker.getStyle(2, col('Post-touch Event ID')).fontColor).toBe(
      CONFIG.sheetFormat.mutedFontColor,
    );
    expect(tracker.getStyle(2, col('Deadline')).horizontalAlignment).toBe('center');
    expect(tracker.getStyle(2, col('Done')).horizontalAlignment).toBe('center');
    expect(tracker.getStyle(2, col('Name')).fontColor).toBeUndefined();
  });

  it('styles the Menu of Asks and Settings', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    for (const name of ['Menu of Asks', 'Settings']) {
      const sheet = w.byName(name);
      expect(sheet.getStyle(1, 1).fontWeight).toBe('bold');
      expect(sheet.getStyle(2, 1).wrap).toBe(true);
    }
  });

  it('leaves the form response tab unstyled', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');
    const form = w.byName('Form Responses 1');
    const widthBefore = form.getColumnWidth(1);

    runSetupAutomation(w.services);

    expect(form.getStyle(1, 1)).toEqual({});
    expect(form.getColumnWidth(1)).toBe(widthBefore);
  });

  it('re-fits the sheets on a second Set Up', () => {
    const longId = 'x'.repeat(80);
    const w = createWorld([
      { name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] },
      {
        name: 'Directory',
        headers: DIRECTORY_HEADERS,
        rows: [['Jane', 'jane@example.org', '555-0100', '', '', longId]],
      },
    ]);
    w.ui.queuePrompt('cancel');
    runSetupAutomation(w.services);
    w.ui.queuePrompt('cancel');
    runSetupAutomation(w.services);

    const directory = w.byName('Directory');
    const eventIdCol = DIRECTORY_HEADERS.indexOf('Calendar Event ID') + 1;
    expect(directory.getColumnWidth(eventIdCol)).toBe(CONFIG.sheetFormat.maxColumnWidth);
    expect(directory.getStyle(1, 1).fontWeight).toBe('bold');
  });

  it('scaffolds tabs, seeds the menu, links the form, installs both triggers', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('https://forms.example/sample', 'ok');

    runSetupAutomation(w.services);

    // menu created and seeded
    const menu = w.byName('Menu of Asks');
    expect(menu.getLastRow() - 1).toBe(CONFIG.defaultAsks.length);

    // Directory and Tracker created with their headers; the old Pipeline name is gone
    expect(colIndex(w.byName('Directory'), 'Reach out by')).toBeGreaterThan(0);
    expect(colIndex(w.byName('Tracker'), 'Pre-touch Event ID')).toBeGreaterThan(0);
    expect(w.sheets.map((s) => s.getName())).not.toContain('Pipeline');
    expect(headersOf(w.byName('Tracker'))).toEqual(
      expect.arrayContaining(['Name', 'Email', 'Ask', 'Deadline', 'Notes']),
    );

    // form linked and both triggers installed
    expect(w.forms.linkedSpreadsheetId).toBe(w.ss.id);
    expect(w.ss.getFormUrl()).toBe('https://forms.example/sample');
    expect([...w.script.triggers].sort()).toEqual(['onEditInstalled', 'onFormSubmitInstalled']);
  });

  it('creates Settings with the four timing defaults', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    const settings = w.byName('Settings');
    expect(headersOf(settings)).toEqual(['Setting', 'Value']);
    expect(settings.getRange(2, 1, 4, 2).getValues()).toEqual([
      ['First contact (days after signup)', 1],
      ['Pre-touch (days before deadline)', 2],
      ['Post-touch (days after deadline)', 3],
      ['Reminder hour (24h)', 9],
    ]);
  });

  it('reports an invalid Settings value loudly at Set Up', () => {
    const w = createWorld([
      { name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] },
      settingsSpec([
        ['First contact (days after signup)', 'soon'],
        ['Pre-touch (days before deadline)', 2],
        ['Post-touch (days after deadline)', 3],
        ['Reminder hour (24h)', 9],
      ]),
    ]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    const summary = w.ui.alerts.find((a) => a.title === 'Volunteer Tools — Set Up')?.text ?? '';
    expect(summary).toContain('First contact');
    expect(summary.toLowerCase()).toContain('number');
  });

  it('is idempotent — second run installs no duplicate triggers and does not reseed', () => {
    const w = createWorld([
      { name: 'Form Responses 1', headers: ['Timestamp', 'Name', 'Email'], rows: [] },
    ]);
    w.ui.queuePrompt('https://forms.example/sample', 'ok');
    runSetupAutomation(w.services);
    const menu = w.byName('Menu of Asks');
    const askCountAfterFirst = menu.getLastRow() - 1;

    w.ui.queuePrompt('cancel');
    runSetupAutomation(w.services);

    expect([...w.script.triggers].sort()).toEqual(['onEditInstalled', 'onFormSubmitInstalled']);
    expect(w.byName('Menu of Asks').getLastRow() - 1).toBe(askCountAfterFirst);
    // the freeze survives a re-run at row 1
    expect(w.byName('Directory').getFrozenRows()).toBe(1);
    expect(w.byName('Tracker').getFrozenRows()).toBe(1);
    expect(w.byName('Menu of Asks').getFrozenRows()).toBe(1);
    expect(w.byName('Settings').getFrozenRows()).toBe(1);
    // columns not duplicated
    expect(headersOf(w.byName('Directory'))).toEqual(DIRECTORY_HEADERS);
  });

  it('reports missing required fields loudly in the summary', () => {
    const w = createWorld([
      { name: 'Form Responses 1', headers: ['Timestamp', 'Name'], rows: [] }, // no Email
    ]);
    w.ui.queuePrompt('https://forms.example/sample', 'ok');
    runSetupAutomation(w.services);

    const alertText = w.ui.alerts
      .map((a) => a.text)
      .find((t) => t.includes('Missing required field(s)'));
    expect(alertText).toContain('Email');
  });

  it('flags a form without a Phone column as missing a required field', () => {
    const w = createWorld([
      { name: 'Form Responses 1', headers: ['Timestamp', 'Name', 'Email'], rows: [] }, // no Phone
    ]);
    w.ui.queuePrompt('https://forms.example/sample', 'ok');
    runSetupAutomation(w.services);

    const alertText = w.ui.alerts
      .map((a) => a.text)
      .find((t) => t.includes('Missing required field(s)'));
    expect(alertText).toContain('Phone');
  });

  it('reports 0 actions on an empty Tracker, despite the date-format block', () => {
    const w = createWorld([
      { name: 'Form Responses 1', headers: ['Timestamp', 'Name', 'Email', 'Phone'], rows: [] },
    ]);
    w.ui.queuePrompt('https://forms.example/sample', 'ok');
    runSetupAutomation(w.services);

    const summary = w.ui.alerts.find((a) => a.title === 'Volunteer Tools — Set Up')?.text ?? '';
    expect(summary).toContain('Tracker: 0 actions so far');
  });

  it('success summary names every required field (stays in sync with CONFIG)', () => {
    const w = createWorld([
      { name: 'Form Responses 1', headers: ['Timestamp', 'Name', 'Email', 'Phone'], rows: [] },
    ]);
    w.forms.questions = ['Name', 'Email', 'Phone'];
    w.ui.queuePrompt('https://forms.example/sample', 'ok');
    runSetupAutomation(w.services);

    const alertText = w.ui.alerts
      .map((a) => a.text)
      .find((t) => t.includes('Required fields present'));
    expect(alertText).toBeDefined();
    Object.values(CONFIG.requiredFields).forEach((f) => {
      expect(alertText).toContain(f);
    });
  });
});

describe('form submissions', () => {
  it('mirrors the declared fields into a Directory row and never writes to Form Responses', () => {
    const w = trackerWorld();
    const form = w.byName('Form Responses 1');

    const before = submitRow(w, 3, [
      '2026-02-02 10:00:00',
      'New Person',
      'new@example.org',
      '555-0199',
      'Excited',
    ]);

    expect(form.grid).toEqual(before);
    const directory = w.byName('Directory');
    const row = directory.getLastRow();
    expect(String(directory.getRange(row, colIndex(directory, 'Name')).getValue())).toBe(
      'New Person',
    );
    expect(String(directory.getRange(row, colIndex(directory, 'Email')).getValue())).toBe(
      'new@example.org',
    );
    expect(String(directory.getRange(row, colIndex(directory, 'Phone')).getValue())).toBe(
      '555-0199',
    );
    expect(String(directory.getRange(row, colIndex(directory, 'About')).getValue())).toBe(
      'Excited',
    );
  });

  it('adds a second Directory row for a repeat email — no Returning machinery', () => {
    const w = trackerWorld();
    const before = w.byName('Directory').getLastRow();

    submitRow(w, 3, ['2026-02-02 10:00:00', 'Jane', 'jane@example.org', '555-0100', 'again']);

    expect(w.byName('Directory').getLastRow()).toBe(before + 1);
    for (const sheet of w.sheets) {
      expect(headersOf(sheet)).not.toContain('Entry');
    }
  });
});

describe('intake scheduling', () => {
  it('schedules one first-contact reminder at signup + the configured delay', () => {
    const w = trackerWorld();

    submitRow(w, 3, [
      '2026-02-02 10:00:00',
      'New Person',
      'new@example.org',
      '555-0199',
      'Excited',
    ]);

    const directory = w.byName('Directory');
    const row = directory.getLastRow();
    const reachOutBy = directory.getRange(row, colIndex(directory, 'Reach out by')).getValue();
    expect(reachOutBy).toBeInstanceOf(Date);
    expect(fmtDate(reachOutBy as Date)).toBe('2026-02-03'); // signup + 1 day
    expect(directory.getNumberFormat(row, colIndex(directory, 'Reach out by'))).toBe('yyyy-mm-dd');

    expect(w.calendar.events).toHaveLength(1);
    const event = eventAt(w, 0);
    expect(event.title).toBe('First contact: New Person');
    expect(fmtDate(event.start)).toBe('2026-02-03');
    expect(event.start.getHours()).toBe(9);
    expect(
      String(directory.getRange(row, colIndex(directory, 'Calendar Event ID')).getValue()),
    ).toMatch(/^ev/);
  });

  it('carries the submitted contact details and About note in the first-contact event', () => {
    const w = trackerWorld();

    submitRow(w, 3, [
      '2026-02-02 10:00:00',
      'New Person',
      'new@example.org',
      '555-0199',
      'Excited',
    ]);

    const description = eventAt(w, 0).description;
    expect(description).toContain('Email: new@example.org');
    expect(description).toContain('Phone: 555-0199');
    expect(description).toContain('About: Excited');
  });
});

describe('campaign calendar', () => {
  it('Set Up creates it once and records it in Settings', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');

    runSetupAutomation(w.services);

    expect(w.calendars.created).toHaveLength(1);
    const created = w.calendars.created[0];
    expect(created?.getName()).toBe(`${CONFIG.calendar.namePrefix} — ${w.ss.getName()}`);
    const settings = w.byName('Settings');
    const recorded = settings
      .getRange(1, 1, settings.getLastRow(), 2)
      .getValues()
      .find(([label]) => String(label).trim() === CONFIG.calendar.settingLabel);
    expect(String(recorded?.[1])).toBe(created?.getId());

    const summary = w.ui.alerts.find((a) => a.title === 'Volunteer Tools — Set Up')?.text ?? '';
    expect(summary).toContain(created?.getName() ?? '');
    expect(summary).toContain('share');
  });

  it('reuses the recorded calendar on a second Set Up', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');
    runSetupAutomation(w.services);
    const id = w.calendars.created[0]?.getId();

    w.ui.queuePrompt('cancel');
    runSetupAutomation(w.services);

    expect(w.calendars.created).toHaveLength(1);
    expect(w.calendars.created[0]?.getId()).toBe(id);
  });

  it('recovers by name when the recorded id is lost but the calendar remains', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');
    runSetupAutomation(w.services);
    const id = w.calendars.created[0]?.getId();

    clearCampaignCalendar(w);
    w.ui.queuePrompt('cancel');
    runSetupAutomation(w.services);

    expect(w.calendars.created).toHaveLength(1);
    const settings = w.byName('Settings');
    const recorded = settings
      .getRange(1, 1, settings.getLastRow(), 2)
      .getValues()
      .find(([label]) => String(label).trim() === CONFIG.calendar.settingLabel);
    expect(String(recorded?.[1])).toBe(id);
  });

  it('creates a fresh calendar when the recorded one no longer exists', () => {
    const w = createWorld([{ name: 'Form Responses 1', headers: FORM_HEADERS, rows: [] }]);
    w.ui.queuePrompt('cancel');
    runSetupAutomation(w.services);
    w.calendars.remove(w.calendars.created[0]?.getId() ?? '');

    w.ui.queuePrompt('cancel');
    runSetupAutomation(w.services);

    expect(w.calendars.created).toHaveLength(2);
  });

  it('schedules first contact on the campaign calendar, never the personal one', () => {
    const w = trackerWorld();

    submitRow(w, 3, ['2026-02-02 10:00:00', 'New Person', 'new@example.org', '555-0199', 'Hi']);

    expect(w.calendar.events).toHaveLength(1);
    expect(w.defaultCalendar.events).toHaveLength(0);
  });

  it('schedules pre/post on the campaign calendar, never the personal one', () => {
    const w = trackerWorld();

    edit(w, 'Tracker', 2, 'Deadline', deadline);

    expect(w.calendar.events).toHaveLength(2);
    expect(w.defaultCalendar.events).toHaveLength(0);
  });

  it('moves the reminders when a different editor edits the deadline', () => {
    const w = trackerWorld();
    const editorA = w.services;
    const editorB = withoutUi(w).services;

    edit(w, 'Tracker', 2, 'Deadline', deadline, '', editorA);
    const ids = w.calendar.events.map((e) => e.id);

    edit(w, 'Tracker', 2, 'Deadline', new Date(2026, 3, 1), '', editorB);

    expect(w.calendar.events.map((e) => e.id)).toEqual(ids);
    expect(fmtDate(eventAt(w, 0).start)).toBe('2026-03-30');
  });

  it('reports a missing calendar inline instead of writing to the personal one', () => {
    const w = trackerWorld();
    clearCampaignCalendar(w);

    const result = createActionFromDialog(w.services, {
      volunteerRow: 2,
      ask: 'Phone banking',
      deadline: fmtDate(deadline),
    });

    expect(result.ok).toBe(false);
    expect(result.message).toMatch(/calendar/i);
    expect(w.calendar.events).toHaveLength(0);
    expect(w.defaultCalendar.events).toHaveLength(0);
  });

  it('throws a clear intake error with no partial Directory row when the calendar is gone', () => {
    const w = trackerWorld();
    clearCampaignCalendar(w);
    const before = w.byName('Directory').getLastRow();

    expect(() =>
      submitRow(w, 3, ['2026-02-02 10:00:00', 'New', 'new@example.org', '555-3', 'x']),
    ).toThrow(/campaign calendar/i);
    expect(w.byName('Directory').getLastRow()).toBe(before);
  });
});

describe('Settings', () => {
  it('applies a changed delay to later signups only', () => {
    const w = trackerWorld();
    submitRow(w, 3, ['2026-02-02 10:00:00', 'First', 'first@example.org', '555-1', 'a']);
    const directory = w.byName('Directory');
    const firstRow = directory.getLastRow();
    const reachCol = colIndex(directory, 'Reach out by');

    setSetting(w, 'First contact (days after signup)', 5);

    submitRow(w, 4, ['2026-02-05 10:00:00', 'Second', 'second@example.org', '555-2', 'b']);
    const secondRow = directory.getLastRow();

    expect(fmtDate(directory.getRange(firstRow, reachCol).getValue() as Date)).toBe('2026-02-03');
    expect(fmtDate(directory.getRange(secondRow, reachCol).getValue() as Date)).toBe('2026-02-10');
    expect(w.calendar.events).toHaveLength(2);
    expect(fmtDate(eventAt(w, 1).start)).toBe('2026-02-10');
  });

  it('throws loudly instead of scheduling when a value is invalid', () => {
    const w = trackerWorld();
    setSetting(w, 'First contact (days after signup)', 'soon');

    expect(() =>
      submitRow(w, 3, ['2026-02-02 10:00:00', 'New', 'new@example.org', '555-3', 'c']),
    ).toThrow(/First contact/);
    expect(w.calendar.events).toHaveLength(0);
  });
});

describe('trigger context (no editor session)', () => {
  it('intake writes the Directory row and schedules first contact without a UI', () => {
    const w = trackerWorld();
    const { services } = withoutUi(w);

    submitRow(
      w,
      3,
      ['2026-02-02 10:00:00', 'New Person', 'new@example.org', '555-0199', 'Excited'],
      services,
    );

    const directory = w.byName('Directory');
    const row = directory.getLastRow();
    expect(String(directory.getRange(row, colIndex(directory, 'Name')).getValue())).toBe(
      'New Person',
    );
    expect(w.calendar.events).toHaveLength(1);
    expect(eventAt(w, 0).title).toBe('First contact: New Person');
  });

  it('moves a Directory first-contact event without a UI', () => {
    const w = trackerWorld();
    edit(w, 'Directory', 2, 'Reach out by', new Date(2026, 2, 17), '', withoutUi(w).services);

    expect(w.calendar.events).toHaveLength(1);
    expect(fmtDate(eventAt(w, 0).start)).toBe('2026-03-17');
  });

  it('creates both Tracker reminders from a deadline without a UI', () => {
    const w = trackerWorld();
    edit(w, 'Tracker', 2, 'Deadline', deadline, '', withoutUi(w).services);

    expect(w.calendar.events).toHaveLength(2);
    expect(fmtDate(eventAt(w, 0).start)).toBe(fmtDate(pre));
    expect(fmtDate(eventAt(w, 1).start)).toBe(fmtDate(post));
  });

  it('relocates a manual touch edit without a UI', () => {
    const w = trackerWorld();
    const { services } = withoutUi(w);
    edit(w, 'Tracker', 2, 'Deadline', deadline, '', services);
    edit(w, 'Tracker', 2, 'Pre-touch', new Date(2026, 2, 10), '', services);

    expect(w.calendar.events).toHaveLength(2);
    expect(fmtDate(eventAt(w, 0).start)).toBe('2026-03-10');
  });

  it('reverts an invalid date and logs it without a UI', () => {
    const w = trackerWorld();
    const { services, logs } = withoutUi(w);
    edit(w, 'Directory', 2, 'Reach out by', 'not a date', '2026-02-01', services);

    const directory = w.byName('Directory');
    expect(String(directory.getRange(2, colIndex(directory, 'Reach out by')).getValue())).toBe(
      '2026-02-01',
    );
    expect(logs.some((m) => m.includes("doesn't look like a date"))).toBe(true);
  });
});

describe('New Action dialog', () => {
  it('offers every Directory volunteer and every Menu of Asks entry', () => {
    const w = trackerWorld();

    expect(actionOptions(w.services)).toEqual({
      volunteers: [{ row: 2, name: 'Jane', email: 'jane@example.org' }],
      asks: ['Canvassing (base)', 'Phone banking'],
    });
  });

  it('opens the dialog pre-loaded with the options', () => {
    const w = trackerWorld();

    openActionDialog(w.services);

    expect(w.ui.dialogs).toHaveLength(1);
    const dialog = w.ui.dialogs[0];
    expect(dialog?.title).toBe('New Action');
    expect(dialog?.html).toContain('jane@example.org');
    expect(dialog?.html).toContain('Phone banking');
  });

  it('creates the Action and schedules both touches', () => {
    const w = trackerWorld();

    const result = createActionFromDialog(w.services, {
      volunteerRow: 2,
      ask: 'Phone banking',
      deadline: fmtDate(deadline),
    });

    expect(result.ok).toBe(true);
    const tracker = w.byName('Tracker');
    const row = 3; // row 2 already holds Jane's first Action
    expect(tracker.getRange(row, colIndex(tracker, 'Name')).getValue()).toBe('Jane');
    expect(tracker.getRange(row, colIndex(tracker, 'Email')).getValue()).toBe('jane@example.org');
    expect(tracker.getRange(row, colIndex(tracker, 'Ask')).getValue()).toBe('Phone banking');
    expect(fmtDate(tracker.getRange(row, colIndex(tracker, 'Pre-touch')).getValue() as Date)).toBe(
      fmtDate(pre),
    );
    expect(fmtDate(tracker.getRange(row, colIndex(tracker, 'Post-touch')).getValue() as Date)).toBe(
      fmtDate(post),
    );
    for (const h of ['Deadline', 'Pre-touch', 'Post-touch']) {
      expect(tracker.getNumberFormat(row, colIndex(tracker, h))).toBe('yyyy-mm-dd');
    }

    expect(w.calendar.events).toHaveLength(2);
    expect(eventAt(w, 0).title).toBe('Pre-touch: Jane');
    expect(eventAt(w, 1).title).toBe('Post-touch: Jane — Phone banking');
  });

  it('adds a brand-new Ask to the Menu of Asks', () => {
    const w = trackerWorld();

    const result = createActionFromDialog(w.services, {
      volunteerRow: 2,
      ask: 'Text banking',
      deadline: fmtDate(deadline),
    });

    expect(result.ok).toBe(true);
    const menu = w.byName('Menu of Asks');
    const asks = menu
      .getRange(2, 1, menu.getLastRow() - 1, 1)
      .getValues()
      .map(([ask]) => String(ask));
    expect(asks.filter((ask) => ask === 'Text banking')).toHaveLength(1);
  });

  it('does not duplicate an Ask that is already on the menu', () => {
    const w = trackerWorld();

    createActionFromDialog(w.services, {
      volunteerRow: 2,
      ask: 'Phone banking',
      deadline: fmtDate(deadline),
    });

    const menu = w.byName('Menu of Asks');
    const asks = menu
      .getRange(2, 1, menu.getLastRow() - 1, 1)
      .getValues()
      .map(([ask]) => String(ask));
    expect(asks.filter((ask) => ask === 'Phone banking')).toHaveLength(1);
  });

  it('accepts an Ask added after the Tracker dropdown was pinned', () => {
    const w = trackerWorld();
    const tracker = w.byName('Tracker');
    const menu = w.byName('Menu of Asks');
    const askCol = colIndex(tracker, 'Ask');
    // Set Up pinned the strict Ask dropdown to the asks that existed then; a
    // later ask (here, added by the dialog) must not be rejected on write.
    const pinned = new MockValidationBuilder()
      .requireValueInRange(menu.getRange(2, 1, menu.getLastRow() - 1, 1))
      .setAllowInvalid(false)
      .build();
    tracker.getRange(2, askCol, 1000, 1).setDataValidation(pinned);

    const result = createActionFromDialog(w.services, {
      volunteerRow: 2,
      ask: 'Text banking',
      deadline: fmtDate(deadline),
    });

    expect(result.ok).toBe(true);
    expect(tracker.getRange(3, askCol).getValue()).toBe('Text banking');
  });

  it('refuses to create an Action without a volunteer, Ask, or parseable Deadline', () => {
    const cases = [
      { volunteerRow: 1, ask: 'Phone banking', deadline: fmtDate(deadline) },
      { volunteerRow: 2, ask: '   ', deadline: fmtDate(deadline) },
      { volunteerRow: 2, ask: 'Phone banking', deadline: 'soon' },
    ];

    for (const input of cases) {
      const w = trackerWorld();
      const before = w.byName('Tracker').getLastRow();

      const result = createActionFromDialog(w.services, input);

      expect(result.ok).toBe(false);
      expect(result.message).not.toBe('');
      expect(w.byName('Tracker').getLastRow()).toBe(before);
      expect(w.calendar.events).toHaveLength(0);
    }
  });
});
