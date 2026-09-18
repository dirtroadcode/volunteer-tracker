/**
 * The "glue" of Volunteerist — Apps Script behaviors factored out of the
 * entry points so they can be unit-tested against a mock Google harness.
 *
 * Everything here takes its services (sheets, calendar, ui, script, forms)
 * as parameters instead of reading Google globals; `index.ts` wires the real
 * services in, and `tests/mocks/gas.ts` provides fakes.
 */
import { CONFIG, type TouchKind } from './config';
import { type ActionOptions, type ActionVolunteer, actionDialogHtml } from './dialog';
import {
  addDays,
  atHour,
  deriveTouchDates,
  directoryHeaders,
  eventDescription,
  eventIdFromCell,
  findColumn,
  fmtDate,
  norm,
  parseDateValue,
  parseSettings,
  type Settings,
  settingsHeaders,
  settingsRows,
  touchKindForColumn,
  touchTitle,
  trackerHeaders,
} from './pure';

// ---- structural views of the Google services (subset used) -------

export interface SheetLike {
  getName(): string;
  setName(name: string): void;
  getLastRow(): number;
  getLastColumn(): number;
  getRange(row: number, col: number, numRows?: number, numCols?: number): RangeLike;
  appendRow(values: unknown[]): void;
  getSheetId(): number;
  activate(): void;
  setFrozenRows(rows: number): void;
  autoResizeColumns(startColumn: number, numColumns: number): void;
  getColumnWidth(columnPosition: number): number;
  setColumnWidth(columnPosition: number, width: number): void;
  getParent(): SheetsLike;
  getActiveCell(): RangeLike;
  setActiveRange(range: RangeLike): void;
  /** URL of the form linked to this tab, or null for an ordinary sheet. */
  getFormUrl(): string | null;
}

export interface SheetsLike {
  getSheetByName(name: string): SheetLike | null;
  insertSheet(name: string): SheetLike;
  getSheets(): SheetLike[];
  getId(): string;
  getName(): string;
  getFormUrl(): string | null;
  getSheetById(id: number): SheetLike | null;
  getActiveSheet(): SheetLike;
  toast(message: string, title?: string, timeoutSeconds?: number): void;
  setActiveSheet(sheet: SheetLike): void;
  deleteSheet(sheet: SheetLike): void;
}

export interface RangeLike {
  getValue(): unknown;
  getValues(): unknown[][];
  setValue(value: unknown): RangeLike;
  setNumberFormat(format: string): RangeLike;
  setDataValidation(rule: unknown): RangeLike;
  setWrap(isWrapEnabled: boolean): RangeLike;
  setFontWeight(weight: 'normal' | 'bold' | null): RangeLike;
  setBackground(color: string | null): RangeLike;
  setFontColor(color: string | null): RangeLike;
  setFontSize(size: number): RangeLike;
  setHorizontalAlignment(alignment: 'left' | 'center' | 'normal' | 'right' | null): RangeLike;
  setVerticalAlignment(alignment: 'top' | 'middle' | 'bottom' | null): RangeLike;
  getRow(): number;
  getColumn(): number;
  getSheet(): SheetLike;
  activate(): void;
}

export interface CalendarEventLike {
  setTitle(title: string): void;
  setDescription(description: string): void;
  setTime(start: Date, end: Date): void;
  deleteEvent(): void;
}

export interface CalendarLike {
  getId(): string;
  createEvent(
    title: string,
    start: Date,
    end: Date,
    options?: { description?: string },
  ): { getId(): string };
  getEventById(id: string): CalendarEventLike | null;
}

/** The calendar app itself — creating and finding calendars, not just using one. */
export interface CalendarsLike {
  createCalendar(name: string, options?: { description?: string; timeZone?: string }): CalendarLike;
  getCalendarById(id: string): CalendarLike | null;
  getOwnedCalendarsByName(name: string): CalendarLike[];
}

export interface UiLike {
  prompt(
    title: string,
    promptText: string,
    buttons: unknown,
  ): { getSelectedButton(): unknown; getResponseText(): string };
  alert(title: string, text: string, buttons: unknown): unknown;
  ButtonSet: { OK: unknown; OK_CANCEL: unknown };
  Button: { OK: unknown };
}

export interface ScriptLike {
  getProjectTriggers(): { getHandlerFunction(): string }[];
  newTrigger(handler: string): {
    forSpreadsheet(ss: SheetsLike): {
      onFormSubmit(): { create(): void };
      onEdit(): { create(): void };
    };
  };
}

export interface FormItemLike {
  setTitle(title: string): FormItemLike;
  setRequired(required: boolean): FormItemLike;
  setHelpText(text: string): FormItemLike;
}

export interface FormLinkLike {
  addTextItem(): FormItemLike;
  addParagraphTextItem(): FormItemLike;
  setDestination(spreadsheetId: string): void;
  removeDestination(): void;
}

export interface FormsLike {
  openByUrl(url: string): FormLinkLike;
  create(title: string): FormLinkLike;
}

export interface ValidationBuilderLike {
  requireValueInRange(range: unknown, showInvalid?: boolean): ValidationBuilderLike;
  requireCheckbox(): ValidationBuilderLike;
  requireFormulaSatisfied(formula: string): ValidationBuilderLike;
  setAllowInvalid(allow: boolean): ValidationBuilderLike;
  setHelpText(text: string): ValidationBuilderLike;
  build(): unknown;
}

/** Everything the glue needs, injected together. */
export interface Services {
  ss: SheetsLike;
  ui: UiLike;
  /** The calendar app — creating and finding calendars, not just using one. */
  calendars: CalendarsLike;
  /** The Campaign Calendar every Touchpoint is written to. Resolved lazily from
   *  the recorded id, so Set Up can create it before anything reads it. */
  cal: CalendarLike;
  script: ScriptLike;
  forms: FormsLike;
  /** Flush pending spreadsheet writes (SpreadsheetApp.flush) — needed after
   * linking a form, before the response tab it creates can be found. */
  flush(): void;
  makeValidation(): ValidationBuilderLike;
  /** The execution log — the only channel a trigger execution has. */
  log(message: string): void;
  /** Render HTML in a modal dialog — menu handlers cannot open a browser tab,
   *  so a clickable link has to ride in a dialog. */
  showHtml(html: string, title: string): void;
}

export interface EditEventLike {
  source: SheetsLike;
  range: RangeLike;
  value: unknown;
  oldValue?: unknown;
}

export interface FormSubmitEventLike {
  source: SheetsLike;
  range: RangeLike;
}

// ----------------------------------------------------------- helpers

function headersOf(sheet: SheetLike): string[] {
  const last = sheet.getLastColumn();
  if (last < 1) return [];
  const [row = []] = sheet.getRange(1, 1, 1, last).getValues();
  return row.map((v) => String(v ?? ''));
}

/** 1-based column index of a header, or -1. */
function columnIndexOf(sheet: SheetLike, headerText: string): number {
  return findColumn(headersOf(sheet), headerText) + 1;
}

function cellValue(sheet: SheetLike, row: number, headerText: string): unknown {
  const col = columnIndexOf(sheet, headerText);
  if (col <= 0) return '';
  return sheet.getRange(row, col).getValue();
}

function setCellValue(sheet: SheetLike, row: number, headerText: string, value: unknown): void {
  const col = columnIndexOf(sheet, headerText);
  if (col <= 0) throw new Error(`Column "${headerText}" not found on tab "${sheet.getName()}"`);
  sheet.getRange(row, col).setValue(value);
}

/** Append any headers that are missing. Returns true if anything was added. */
function appendColumns(sheet: SheetLike, headerTexts: readonly string[]): boolean {
  const headers = headersOf(sheet);
  let added = false;
  for (const h of headerTexts) {
    if (findColumn(headers, h) === -1) {
      sheet.getRange(1, headers.length + 1).setValue(h);
      headers.push(h);
      added = true;
    }
  }
  return added;
}

function eventEnd(start: Date): Date {
  return new Date(start.getTime() + CONFIG.touch.durationMinutes * 60 * 1000);
}

/**
 * Real data extent of a sheet, read from one identity column. A sheet's last
 * row counts formatted-but-empty rows, so it cannot tell us where data ends —
 * this scans cell values instead and never trusts getLastRow() alone.
 */
function sheetExtent(
  sheet: SheetLike,
  identityHeader: string,
): { count: number; firstFreeRow: number } {
  const identityCol = columnIndexOf(sheet, identityHeader);
  if (identityCol <= 0) {
    throw new Error(`Column "${identityHeader}" not found on tab "${sheet.getName()}"`);
  }
  const last = sheet.getLastRow();
  const identities = last < 2 ? [] : sheet.getRange(2, identityCol, last - 1, 1).getValues();
  const isEmpty = (row: unknown[]) => String(row[0] ?? '').trim() === '';
  const firstEmpty = identities.findIndex(isEmpty);
  return {
    count: identities.filter((row) => !isEmpty(row)).length,
    firstFreeRow: firstEmpty === -1 ? Math.max(last, 1) + 1 : firstEmpty + 2,
  };
}

/** Format a row's date columns as yyyy-mm-dd, whoever writes them. */
function formatDateCells(sheet: SheetLike, row: number, headers: readonly string[]): void {
  for (const h of headers) {
    const c = columnIndexOf(sheet, h);
    if (c > 0) sheet.getRange(row, c).setNumberFormat('yyyy-mm-dd');
  }
}

const TRACKER_DATE_HEADERS = [
  CONFIG.trackerColumns.deadline,
  CONFIG.trackerColumns.preTouch,
  CONFIG.trackerColumns.postTouch,
] as const;

/** Update an existing event (by stored id) or create one. Returns the id. */
function createOrUpdateEvent(
  services: Services,
  opts: { existingId: string; title: string; description: string; start: Date; end: Date },
): string {
  const existingId = eventIdFromCell(opts.existingId);
  if (existingId) {
    try {
      const existing = services.cal.getEventById(existingId);
      if (existing) {
        existing.setTitle(opts.title);
        existing.setDescription(opts.description);
        existing.setTime(opts.start, opts.end);
        return existingId;
      }
    } catch {
      // stale/missing id — fall through and create fresh
    }
  }
  return services.cal
    .createEvent(opts.title, opts.start, opts.end, {
      description: opts.description,
    })
    .getId();
}

function deleteEventById(services: Services, rawId: string): void {
  const id = eventIdFromCell(rawId);
  if (!id) return;
  try {
    services.cal.getEventById(id)?.deleteEvent();
  } catch {
    // stale id — nothing to delete
  }
}

function toast(services: Services, message: string, seconds = 6): void {
  services.ss.toast(message, 'Volunteerist', seconds);
}

// ------------------------------------------------------ setup (FR3/FR5)

export function runSetupAutomation(services: Services): void {
  const { ui, ss } = services;

  const linked = linkForm(services);
  const directory = ensureDirectory(services);
  const settingsSheet = ensureSettings(services);
  const calendar = ensureCampaignCalendar(services);

  const menu = ensureMenuOfAsks(services);
  const tracker = ensureTracker(services, menu);

  ensureTrigger(services);

  const formResponses = findFormResponseSheet(services.ss);
  const missing = formResponses ? validateRequiredHeaders(formResponses) : [];
  const { settings, problems: settingsProblems } = parseSettings(sheetRows(settingsSheet));

  const lines: string[] = [
    `• Directory (roster): ${directory.getName()}`,
    `• Menu of Asks: ${menu.getLastRow() - 1} asks (edit freely — no code needed)`,
    `• Tracker: ${sheetExtent(tracker, CONFIG.trackerColumns.name).count} actions so far`,
    `• Calendar: "${calendar.name}" — share it from Google Calendar → Other calendars. A Collaborator who will add Actions also needs Editor access to this spreadsheet (Share); watchers need the calendar only.`,
    `• Triggers: ${
      hasTrigger(services, SUBMIT_HANDLER) && hasTrigger(services, EDIT_HANDLER)
        ? 'installed ✓'
        : 'INSTALL FAILED — re-run Set Up'
    }`,
  ];
  if (settings) {
    lines.push(
      `• Settings: first contact +${settings.firstContactDays}d, pre −${settings.preDaysBefore}d, post +${settings.postDaysAfter}d, reminder ${settings.reminderHour}:00`,
    );
  }
  if (linked) {
    lines.push(`• Form: responses now flow into this spreadsheet — share it: ${ss.getFormUrl()}`);
  } else {
    lines.push('• Form: not linked yet — paste the form URL when you re-run Set Up');
  }
  if (missing.length > 0) {
    lines.push('');
    lines.push(`⚠️  Missing required field(s): ${missing.join(', ')}`);
    lines.push(
      'The intake form must ask for each of these — question titles become the Directory column headers. Use the Sample Form, or rename the questions on your own form to match.',
    );
  } else if (linked) {
    lines.push('');
    lines.push(
      '✅ Required fields present (' +
        Object.values(CONFIG.requiredFields).join(', ') +
        '). Ready to go!',
    );
  }
  if (settingsProblems.length > 0) {
    lines.push('');
    lines.push(`⚠️  Settings need fixing: ${settingsProblems.join('; ')}`);
    lines.push(
      `Type a whole number in every ${CONFIG.tabNames.settings} Value cell, then re-run Set Up.`,
    );
  }

  ui.alert('Volunteer Tools — Set Up', lines.join('\n\n'), ui.ButtonSet.OK);
}

export function linkForm(services: Services): boolean {
  const { ss, ui, forms, flush } = services;
  const current = ss.getFormUrl();
  const required = Object.values(CONFIG.requiredFields).join(', ');
  const about = CONFIG.optionalFields.about.title;
  const paragraphs: string[] = [];
  if (current) {
    paragraphs.push(`A form is already linked: ${current}`);
    paragraphs.push(
      'OK replaces it: the old responses are archived into a "Sign-ups archive" tab and the old form is unlinked.',
    );
  }
  paragraphs.push(
    'Press OK with this box empty and a sign-up form will be generated for you. The new form lives in your Google Drive — share it with volunteers from there.',
  );
  paragraphs.push(
    `Or paste the URL of a form you already have (the Sample Form works too). A sign-up form needs these fields: ${required} — plus the optional ${about} question. After linking, it will send responses here.`,
  );
  paragraphs.push('Press Cancel to leave things as they are — you can re-run Set Up anytime.');
  const result = ui.prompt(
    'Link form or generate new one?',
    paragraphs.join('\n\n'),
    ui.ButtonSet.OK_CANCEL,
  );
  if (result.getSelectedButton() !== ui.Button.OK) {
    // Cancel: keep whatever is linked now (or nothing) exactly as it is.
    return current !== null;
  }
  const url = result.getResponseText().trim();
  if (current) replaceIntake(services, ss);
  if (url) {
    forms.openByUrl(url).setDestination(ss.getId());
  } else {
    createIntakeForm(services).setDestination(ss.getId());
  }
  flush();
  renameFreshResponseTab(ss);
  return true;
}

/** No linked form is a no-op: the menu item only exists while one is linked,
 *  so this guard covers a direct invocation. */
export function showSignUpForm(services: Services): void {
  const url = services.ss.getFormUrl();
  if (!url) return;
  services.showHtml(
    `<p>Volunteers sign up here:</p><p><a href="${url}" target="_blank" rel="noopener">${url}</a></p>`,
    'Sign-Up Form',
  );
}

/** Build an intake form whose questions are exactly the Directory's fields. */
function createIntakeForm(services: Services): FormLinkLike {
  const form = services.forms.create(CONFIG.intakeFormTitle);
  for (const title of Object.values(CONFIG.requiredFields)) {
    form.addTextItem().setTitle(title).setRequired(true);
  }
  form
    .addParagraphTextItem()
    .setTitle(CONFIG.optionalFields.about.title)
    .setHelpText(CONFIG.optionalFields.about.helpText)
    .setRequired(false);
  return form;
}

/** Google always names a fresh linked tab "Form Responses N"; give it a clean name.
 *  Only a form-linked tab qualifies — a stale legacy tab can carry the same name. */
function renameFreshResponseTab(ss: SheetsLike): void {
  const fresh = ss
    .getSheets()
    .find((s) => /^Form Responses \d+$/.test(s.getName()) && s.getFormUrl() !== null);
  if (fresh) fresh.setName(CONFIG.tabNames.responses);
}

/**
 * Start intake over: copy the linked response tab's values into an archive
 * tab, then delete the tab.
 */
function replaceIntake(services: Services, ss: SheetsLike): void {
  const responses = findFormResponseSheet(ss);
  if (!responses) return;
  const formUrl = responses.getFormUrl();
  const last = responses.getLastRow();
  const values =
    last > 0 ? responses.getRange(1, 1, last, responses.getLastColumn()).getValues() : [];
  if (formUrl) {
    // Google refuses to delete a form-linked tab — unlink the form first.
    services.forms.openByUrl(formUrl).removeDestination();
    services.flush();
  }
  ss.deleteSheet(responses);
  if (values.length === 0) return;
  let name: string = CONFIG.tabNames.archive;
  let n = 2;
  while (ss.getSheetByName(name)) name = `${CONFIG.tabNames.archive} ${n++}`;
  const archive = ss.insertSheet(name);
  for (const row of values) archive.appendRow(row);
}

function ensureHeaderRow(sheet: SheetLike): void {
  sheet.setFrozenRows(1);
}

/** How a generated tab is presented: which columns recede or center. */
interface SheetStyle {
  /** Machine bookkeeping that reads as reference, not content. */
  muted: readonly string[];
  /** Values that read better centered than left-aligned. */
  centered: readonly string[];
}

const SHEET_STYLES: Record<string, SheetStyle> = {
  [CONFIG.tabNames.directory]: {
    muted: [CONFIG.directoryColumns.firstContactEvent],
    centered: [CONFIG.directoryColumns.reachOutBy],
  },
  [CONFIG.tabNames.tracker]: {
    muted: [CONFIG.trackerColumns.preTouchEvent, CONFIG.trackerColumns.postTouchEvent],
    centered: [
      CONFIG.trackerColumns.deadline,
      CONFIG.trackerColumns.preTouch,
      CONFIG.trackerColumns.postTouch,
      CONFIG.trackerColumns.done,
    ],
  },
  [CONFIG.tabNames.menu]: { muted: [], centered: [] },
  [CONFIG.tabNames.settings]: { muted: [], centered: [] },
};

/** Size, wrap, and style a script-generated tab so a coordinator can scan it. */
function formatSheet(sheet: SheetLike): void {
  const lastColumn = sheet.getLastColumn();
  const style = SHEET_STYLES[sheet.getName()];
  if (lastColumn < 1 || !style) return;
  const headers = headersOf(sheet);
  wrapAndFit(sheet, lastColumn);
  muteColumns(sheet, headers, lastColumn, style.muted);
  centerColumns(sheet, headers, lastColumn, style.centered);
  sheet
    .getRange(1, 1, 1, lastColumn)
    .setFontWeight('bold')
    .setBackground(CONFIG.sheetFormat.headerBackground)
    .setFontColor(CONFIG.sheetFormat.headerFontColor)
    .setVerticalAlignment('middle');
}

/** Dim a tab's machine columns so they read as reference, not content. */
function muteColumns(
  sheet: SheetLike,
  headers: string[],
  lastColumn: number,
  names: readonly string[],
): void {
  for (const col of columnPositions(headers, names, lastColumn)) {
    sheet
      .getRange(2, col, CONFIG.sheetFormat.wrapRows - 1, 1)
      .setFontColor(CONFIG.sheetFormat.mutedFontColor)
      .setFontSize(CONFIG.sheetFormat.mutedFontSize);
  }
}

/** Center a tab's date and checkbox columns. */
function centerColumns(
  sheet: SheetLike,
  headers: string[],
  lastColumn: number,
  names: readonly string[],
): void {
  for (const col of columnPositions(headers, names, lastColumn)) {
    sheet.getRange(1, col, CONFIG.sheetFormat.wrapRows, 1).setHorizontalAlignment('center');
  }
}

/** 1-based positions of the named columns that exist within the tab's used columns. */
function columnPositions(
  headers: string[],
  names: readonly string[],
  lastColumn: number,
): number[] {
  return names
    .map((name) => findColumn(headers, name) + 1)
    .filter((col) => col >= 1 && col <= lastColumn);
}

/**
 * Size each column to its unwrapped content, cap runaway widths, then wrap.
 * Sheets autofits to the widest *rendered* line, so fitting while wrap is on
 * would size columns to the widest wrapped fragment instead of the value.
 */
function wrapAndFit(sheet: SheetLike, lastColumn: number): void {
  const range = sheet.getRange(1, 1, CONFIG.sheetFormat.wrapRows, lastColumn);
  range.setWrap(false);
  sheet.autoResizeColumns(1, lastColumn);
  for (let column = 1; column <= lastColumn; column++) {
    if (sheet.getColumnWidth(column) > CONFIG.sheetFormat.maxColumnWidth) {
      sheet.setColumnWidth(column, CONFIG.sheetFormat.maxColumnWidth);
    }
  }
  range.setWrap(true);
}

export function ensureDirectory(services: Services): SheetLike {
  const { ss } = services;
  const declared = directoryHeaders();
  const existing = ss.getSheetByName(CONFIG.tabNames.directory);
  if (existing) {
    appendColumns(existing, declared);
    ensureHeaderRow(existing);
    formatSheet(existing);
    return existing;
  }
  const directory = ss.insertSheet(CONFIG.tabNames.directory);
  directory.appendRow(declared);
  ensureHeaderRow(directory);
  formatSheet(directory);
  return directory;
}

/** Rows below a sheet's header, as raw values. */
function sheetRows(sheet: SheetLike): unknown[][] {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
}

/**
 * The Directory is the source of volunteer identity, so the Tracker looks up
 * the details it doesn't carry — phone and About note — by email. A missing
 * Directory, column, or matching row yields empties rather than an error: the
 * reminder must still be created.
 */
function directoryContactForEmail(
  services: Services,
  email: string,
): { phone: string; about: string } {
  const empty = { phone: '', about: '' };
  const wanted = norm(email);
  if (!wanted) return empty;
  const directory = services.ss.getSheetByName(CONFIG.tabNames.directory);
  if (!directory) return empty;
  const emailCol = columnIndexOf(directory, CONFIG.requiredFields.email);
  const phoneCol = columnIndexOf(directory, CONFIG.requiredFields.phone);
  const aboutCol = columnIndexOf(directory, CONFIG.optionalFields.about.title);
  if (emailCol <= 0 || phoneCol <= 0) return empty;
  for (const row of sheetRows(directory)) {
    if (norm(String(row[emailCol - 1] ?? '')) === wanted) {
      return {
        phone: String(row[phoneCol - 1] ?? '').trim(),
        about: aboutCol > 0 ? String(row[aboutCol - 1] ?? '').trim() : '',
      };
    }
  }
  return empty;
}

export function ensureSettings(services: Services): SheetLike {
  const { ss } = services;
  const existing = ss.getSheetByName(CONFIG.tabNames.settings);
  if (existing) {
    ensureHeaderRow(existing);
    formatSheet(existing);
    return existing;
  }
  const settings = ss.insertSheet(CONFIG.tabNames.settings);
  settings.appendRow(settingsHeaders());
  ensureHeaderRow(settings);
  settingsRows().forEach(([label, value], i) => {
    settings.getRange(i + 2, 1).setValue(label);
    settings.getRange(i + 2, 2).setValue(value);
  });
  formatSheet(settings);
  return settings;
}

/** The four timing knobs, read at use time. Throws on a missing or invalid tab. */
export function readSettings(services: Services): Settings {
  const sheet = services.ss.getSheetByName(CONFIG.tabNames.settings);
  if (!sheet) {
    throw new Error(`"${CONFIG.tabNames.settings}" tab is missing — run Set Up.`);
  }
  const { settings, problems } = parseSettings(sheetRows(sheet));
  if (!settings) {
    throw new Error(`Invalid ${CONFIG.tabNames.settings}: ${problems.join('; ')}`);
  }
  return settings;
}

// ------------------------------------------------- campaign calendar

/** The name Set Up gives this campaign's calendar. */
function campaignCalendarName(ss: SheetsLike): string {
  return `${CONFIG.calendar.namePrefix} — ${ss.getName()}`;
}

/** The recorded calendar id, or '' when Set Up has not created it yet. */
export function campaignCalendarId(services: Services): string {
  const sheet = services.ss.getSheetByName(CONFIG.tabNames.settings);
  return sheet ? settingValue(sheet, CONFIG.calendar.settingLabel) : '';
}

/**
 * The Campaign Calendar every Touchpoint is written to. Throws — never falls
 * back to a personal calendar — when Set Up has recorded no calendar or the
 * recorded id no longer resolves (deleted, or a share not yet accepted).
 */
export function campaignCalendar(services: Services): CalendarLike {
  const id = campaignCalendarId(services);
  const cal = id ? services.calendars.getCalendarById(id) : null;
  if (!cal) {
    throw new Error(
      'The campaign calendar is missing. Run Volunteer Tools → Set Up, or accept the calendar share first.',
    );
  }
  return cal;
}

/**
 * Create the Campaign Calendar once — reusing the recorded one, or an owned one
 * already bearing the campaign's name — and write its id to Settings so every
 * editor and trigger resolves the same calendar.
 */
export function ensureCampaignCalendar(services: Services): { id: string; name: string } {
  const settings = services.ss.getSheetByName(CONFIG.tabNames.settings);
  if (!settings) throw new Error(`"${CONFIG.tabNames.settings}" tab is missing — run Set Up.`);

  const name = campaignCalendarName(services.ss);
  const recorded = campaignCalendarId(services);
  const cal =
    (recorded ? services.calendars.getCalendarById(recorded) : null) ??
    services.calendars.getOwnedCalendarsByName(name)[0] ??
    services.calendars.createCalendar(name, { description: CONFIG.calendar.description });

  setSettingValue(settings, CONFIG.calendar.settingLabel, cal.getId());
  return { id: cal.getId(), name };
}

/** One Settings label's value, or ''. */
function settingValue(sheet: SheetLike, label: string): string {
  const row = sheetRows(sheet).find((r) => norm(String(r[0] ?? '')) === norm(label));
  return row ? String(row[1] ?? '').trim() : '';
}

/** Write a Settings label's value, adding the row when it is missing. */
function setSettingValue(sheet: SheetLike, label: string, value: string): void {
  const index = sheetRows(sheet).findIndex((r) => norm(String(r[0] ?? '')) === norm(label));
  const row = index >= 0 ? index + 1 : sheet.getLastRow() + 1;
  sheet.getRange(row, 1).setValue(label);
  sheet.getRange(row, 2).setValue(value);
}

/**
 * The tab a form currently writes responses to: the linked tab when there is
 * one, else the Timestamp-first sheet (legacy tabs that Google created).
 * Read-only to the script.
 */
export function findFormResponseSheet(ss: SheetsLike): SheetLike | null {
  const linked = ss.getSheets().find((s) => s.getFormUrl() !== null);
  if (linked) return linked;
  for (const sheet of ss.getSheets()) {
    if (findColumn(headersOf(sheet), CONFIG.timestampField) === 0) return sheet;
  }
  return null;
}

export function ensureMenuOfAsks(services: Services): SheetLike {
  const { ss, makeValidation } = services;
  let menu = ss.getSheetByName(CONFIG.tabNames.menu);
  if (!menu) {
    menu = ss.insertSheet(CONFIG.tabNames.menu);
    menu.appendRow(['Ask']);
  }
  if (menu.getLastRow() <= 1) {
    CONFIG.defaultAsks.forEach((ask, i) => {
      menu.getRange(i + 2, 1).setValue(ask);
    });
  }
  ensureHeaderRow(menu);
  // Warn (not block) on duplicate ask names — keeps the list scannable.
  const rule = makeValidation()
    .requireFormulaSatisfied('=COUNTIF(A:A, A2)<2')
    .setAllowInvalid(true)
    .setHelpText('Duplicate ask names are confusing — rename or delete one.')
    .build();
  menu.getRange(2, 1, 1000, 1).setDataValidation(rule);
  formatSheet(menu);
  return menu;
}

export function ensureTracker(services: Services, menu: SheetLike): SheetLike {
  const { ss, makeValidation } = services;
  let tracker = ss.getSheetByName(CONFIG.tabNames.tracker);
  if (!tracker) {
    tracker = ss.insertSheet(CONFIG.tabNames.tracker);
    tracker.appendRow(trackerHeaders());
  }
  ensureHeaderRow(tracker);
  formatSheet(tracker);
  syncAskValidation(services, menu, tracker);
  const doneCol = columnIndexOf(tracker, CONFIG.trackerColumns.done);
  if (doneCol > 0) {
    const rule = makeValidation().requireCheckbox().build();
    tracker.getRange(2, doneCol, 1000, 1).setDataValidation(rule);
  }
  return tracker;
}

const SUBMIT_HANDLER = 'onFormSubmitInstalled';
const EDIT_HANDLER = 'onEditInstalled';

export function hasTrigger(services: Services, handler: string): boolean {
  return services.script.getProjectTriggers().some((t) => t.getHandlerFunction() === handler);
}

/**
 * Installable triggers run as the account that created them, so the handlers'
 * calendar calls are authorized. A function named `onEdit` would instead be a
 * simple trigger, which cannot access CalendarApp — hence the `*Installed`
 * names and the explicit installation here.
 */
export function ensureTrigger(services: Services): void {
  const install = (handler: string, kind: 'onFormSubmit' | 'onEdit') => {
    if (hasTrigger(services, handler)) return;
    services.script.newTrigger(handler).forSpreadsheet(services.ss)[kind]().create();
  };
  install(SUBMIT_HANDLER, 'onFormSubmit');
  install(EDIT_HANDLER, 'onEdit');
}

export function validateRequiredHeaders(formResponses: SheetLike): string[] {
  const h = headersOf(formResponses);
  const missing: string[] = [];
  if (findColumn(h, CONFIG.requiredFields.name) === -1) missing.push(CONFIG.requiredFields.name);
  if (findColumn(h, CONFIG.requiredFields.email) === -1) missing.push(CONFIG.requiredFields.email);
  if (findColumn(h, CONFIG.requiredFields.phone) === -1) missing.push(CONFIG.requiredFields.phone);
  return missing;
}

// ------------------------------------------------------- form submit

/** Value of one Form Responses cell, located by header. */
function submissionValue(formHeaders: string[], values: unknown[], header: string): unknown {
  const i = findColumn(formHeaders, header);
  return i === -1 ? '' : (values[i] ?? '');
}

/** Map a Form Responses row onto the Directory's declared fields. */
function directoryRowFromSubmission(
  formHeaders: string[],
  values: unknown[],
): Record<string, unknown> {
  const { name, email, phone } = CONFIG.requiredFields;
  const { title: about } = CONFIG.optionalFields.about;
  const { reachOutBy, firstContactEvent } = CONFIG.directoryColumns;
  const cell = (header: string) => submissionValue(formHeaders, values, header);
  return {
    [name]: cell(name),
    [email]: cell(email),
    [phone]: cell(phone),
    [about]: cell(about),
    [reachOutBy]: '',
    [firstContactEvent]: '',
  };
}

/** Write a row whose values are keyed by header, so reordered columns stay safe. */
function writeRow(sheet: SheetLike, row: number, valuesByHeader: Record<string, unknown>): void {
  for (const header of Object.keys(valuesByHeader)) {
    setCellValue(sheet, row, header, valuesByHeader[header] ?? '');
  }
}

export function handleFormSubmit(services: Services, e: FormSubmitEventLike): void {
  const responseSheet = e.range.getSheet();
  const formHeaders = headersOf(responseSheet);
  const values = e.range.getValues()[0] ?? [];

  const directory = services.ss.getSheetByName(CONFIG.tabNames.directory);
  if (!directory) return;

  const submittedAt = parseDateValue(submissionValue(formHeaders, values, CONFIG.timestampField));
  if (!submittedAt) {
    throw new Error(
      `No parseable "${CONFIG.timestampField}" on a submission to "${responseSheet.getName()}"`,
    );
  }

  const settings = readSettings(services);
  const reachOutBy = addDays(submittedAt, settings.firstContactDays);
  const rowData = directoryRowFromSubmission(formHeaders, values);
  const name = String(rowData[CONFIG.requiredFields.name] ?? '').trim() || 'Volunteer';
  const email = String(rowData[CONFIG.requiredFields.email] ?? '');
  const phone = String(rowData[CONFIG.requiredFields.phone] ?? '');
  const about = String(rowData[CONFIG.optionalFields.about.title] ?? '');

  const start = atHour(reachOutBy, settings.reminderHour);
  const eventId = createOrUpdateEvent(services, {
    existingId: '',
    title: touchTitle('first-contact', name),
    description: eventDescription({
      name,
      email,
      phone,
      about,
      deadline: reachOutBy,
      kind: 'first-contact',
    }),
    start,
    end: eventEnd(start),
  });
  rowData[CONFIG.directoryColumns.reachOutBy] = reachOutBy;
  rowData[CONFIG.directoryColumns.firstContactEvent] = eventId;

  const { firstFreeRow } = sheetExtent(directory, CONFIG.requiredFields.name);
  writeRow(directory, firstFreeRow, rowData);
  formatDateCells(directory, firstFreeRow, [CONFIG.directoryColumns.reachOutBy]);
}

// ---------------------------------------------------------- row edits

export function handleEdit(services: Services, e: EditEventLike): void {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  if (row < 2) return;
  const col = e.range.getColumn(); // 1-based
  const headers = headersOf(sheet);

  if (sheet.getName() === CONFIG.tabNames.directory) {
    if (col === findColumn(headers, CONFIG.directoryColumns.reachOutBy) + 1) {
      handleFirstContactEdit(services, sheet, row, e);
      return;
    }
  }

  if (sheet.getName() === CONFIG.tabNames.tracker) {
    const deadlineCol = findColumn(headers, CONFIG.trackerColumns.deadline) + 1;
    if (col === deadlineCol && deadlineCol > 0) {
      handleDeadlineEdit(services, sheet, row, e);
      return;
    }
    const touchKind = touchKindForColumn(headers, col - 1);
    if (touchKind) {
      handleTouchEdit(services, sheet, row, touchKind, e);
    }
  }
}

/** First contact: a date in Directory "Reach out by" → one calendar event. */
export function handleFirstContactEdit(
  services: Services,
  sheet: SheetLike,
  row: number,
  e: EditEventLike,
): void {
  const idHeader = CONFIG.directoryColumns.firstContactEvent;
  const storedId = String(cellValue(sheet, row, idHeader) ?? '');
  const d = parseDateValue(e.value);
  if (d === null) {
    if (e.value === '' || e.value === null) {
      deleteEventById(services, storedId);
      setCellValue(sheet, row, idHeader, '');
      return;
    }
    revertAndLog(services, e, "That doesn't look like a date — use a date like 3/15/2026.");
    return;
  }
  const name =
    String(cellValue(sheet, row, CONFIG.requiredFields.name) ?? '').trim() || 'Volunteer';
  const email = String(cellValue(sheet, row, CONFIG.requiredFields.email) ?? '');
  const phone = String(cellValue(sheet, row, CONFIG.requiredFields.phone) ?? '');
  const about = String(cellValue(sheet, row, CONFIG.optionalFields.about.title) ?? '');
  const start = atHour(d, readSettings(services).reminderHour);
  const id = createOrUpdateEvent(services, {
    existingId: storedId,
    title: touchTitle('first-contact', name),
    description: eventDescription({
      name,
      email,
      phone,
      about,
      deadline: d,
      kind: 'first-contact',
    }),
    start,
    end: eventEnd(start),
  });
  setCellValue(sheet, row, idHeader, id);
}

/** Derive Pre/Post from a Deadline and create or move both calendar events. */
function scheduleTouches(services: Services, sheet: SheetLike, row: number, deadline: Date): void {
  const ask = String(cellValue(sheet, row, CONFIG.trackerColumns.ask) ?? '').trim();
  const settings = readSettings(services);
  const { pre, post } = deriveTouchDates(
    deadline,
    settings.preDaysBefore,
    settings.postDaysAfter,
    settings.reminderHour,
  );
  const name =
    String(cellValue(sheet, row, CONFIG.trackerColumns.name) ?? '').trim() || 'Volunteer';
  const email = String(cellValue(sheet, row, CONFIG.trackerColumns.email) ?? '');
  const { phone, about } = directoryContactForEmail(services, email);
  const details = { name, email, phone, about, ask, deadline };

  const storedPreId = String(cellValue(sheet, row, CONFIG.trackerColumns.preTouchEvent) ?? '');
  const preId = createOrUpdateEvent(services, {
    existingId: storedPreId,
    title: touchTitle('pre', name),
    description: eventDescription({ ...details, kind: 'pre' }),
    start: pre,
    end: eventEnd(pre),
  });

  const storedPostId = String(cellValue(sheet, row, CONFIG.trackerColumns.postTouchEvent) ?? '');
  const postId = createOrUpdateEvent(services, {
    existingId: storedPostId,
    title: touchTitle('post', name, ask),
    description: eventDescription({ ...details, kind: 'post' }),
    start: post,
    end: eventEnd(post),
  });

  setCellValue(sheet, row, CONFIG.trackerColumns.preTouch, pre);
  setCellValue(sheet, row, CONFIG.trackerColumns.postTouch, post);
  setCellValue(sheet, row, CONFIG.trackerColumns.preTouchEvent, preId);
  setCellValue(sheet, row, CONFIG.trackerColumns.postTouchEvent, postId);
  formatDateCells(sheet, row, TRACKER_DATE_HEADERS);
}

/** Deadline edit: an Ask is required first, then Pre/Post are scheduled. */
export function handleDeadlineEdit(
  services: Services,
  sheet: SheetLike,
  row: number,
  e: EditEventLike,
): void {
  const d = parseDateValue(e.value);
  if (d === null) {
    if (e.value === '' || e.value === null) {
      // Deadline cleared — remove both reminders and the derived cells.
      clearTouch(services, sheet, row, 'pre');
      clearTouch(services, sheet, row, 'post');
      setCellValue(sheet, row, CONFIG.trackerColumns.preTouch, '');
      setCellValue(sheet, row, CONFIG.trackerColumns.postTouch, '');
      return;
    }
    revertAndLog(services, e, "That doesn't look like a date — use a date like 3/15/2026.");
    return;
  }

  const ask = String(cellValue(sheet, row, CONFIG.trackerColumns.ask) ?? '').trim();
  if (!ask) {
    e.range.setValue(e.oldValue ?? '');
    services.log(
      'Choose an Ask first: pick one from the dropdown in the Ask column, then type the Deadline.',
    );
    return;
  }

  scheduleTouches(services, sheet, row, d);
}

/** Manual edit of a Pre/Post-touch date: relocate (or clear) that one event. */
export function handleTouchEdit(
  services: Services,
  sheet: SheetLike,
  row: number,
  kind: TouchKind,
  e: EditEventLike,
): void {
  const idHeader =
    kind === 'pre' ? CONFIG.trackerColumns.preTouchEvent : CONFIG.trackerColumns.postTouchEvent;
  const d = parseDateValue(e.value);
  if (d === null) {
    if (e.value === '' || e.value === null) {
      clearTouch(services, sheet, row, kind);
      return;
    }
    revertAndLog(services, e, "That doesn't look like a date — use a date like 3/15/2026.");
    return;
  }
  const name =
    String(cellValue(sheet, row, CONFIG.trackerColumns.name) ?? '').trim() || 'Volunteer';
  const email = String(cellValue(sheet, row, CONFIG.trackerColumns.email) ?? '');
  const ask = String(cellValue(sheet, row, CONFIG.trackerColumns.ask) ?? '');
  const { phone, about } = directoryContactForEmail(services, email);
  const deadline = parseDateValue(cellValue(sheet, row, CONFIG.trackerColumns.deadline));
  const storedId = String(cellValue(sheet, row, idHeader) ?? '');
  const start = atHour(d, readSettings(services).reminderHour);
  const id = createOrUpdateEvent(services, {
    existingId: storedId,
    title: touchTitle(kind, name, kind === 'post' ? ask : undefined),
    description: eventDescription({ name, email, phone, about, ask, deadline, kind }),
    start,
    end: eventEnd(start),
  });
  setCellValue(sheet, row, idHeader, id);
}

function clearTouch(services: Services, sheet: SheetLike, row: number, kind: TouchKind): void {
  if (kind === 'first-contact') return;
  const idHeader =
    kind === 'pre' ? CONFIG.trackerColumns.preTouchEvent : CONFIG.trackerColumns.postTouchEvent;
  const storedId = String(cellValue(sheet, row, idHeader) ?? '');
  if (storedId) {
    deleteEventById(services, storedId);
    setCellValue(sheet, row, idHeader, '');
  }
  setCellValue(
    sheet,
    row,
    kind === 'pre' ? CONFIG.trackerColumns.preTouch : CONFIG.trackerColumns.postTouch,
    '',
  );
}

function revertAndLog(services: Services, e: EditEventLike, message: string): void {
  e.range.setValue(e.oldValue ?? '');
  services.log(message);
}

// ---------------------------------------------------- actions

export interface ActionSubmission {
  /** 1-based Directory row chosen in the dialog. */
  volunteerRow: number;
  /** An existing Ask, or a brand-new one to add to the Menu of Asks. */
  ask: string;
  /** yyyy-MM-dd from the dialog's date input. */
  deadline: string;
}

/** Outcome the dialog displays; validation failures are the coordinator's to fix. */
export interface ActionResult {
  ok: boolean;
  message: string;
}

/** Volunteers and Asks the New Action dialog offers. */
export function actionOptions(services: Services): ActionOptions {
  const { ss } = services;

  const volunteers: ActionVolunteer[] = [];
  const directory = ss.getSheetByName(CONFIG.tabNames.directory);
  if (directory) {
    const nameCol = columnIndexOf(directory, CONFIG.requiredFields.name);
    const emailCol = columnIndexOf(directory, CONFIG.requiredFields.email);
    sheetRows(directory).forEach((row, i) => {
      const name = nameCol > 0 ? String(row[nameCol - 1] ?? '').trim() : '';
      const email = emailCol > 0 ? String(row[emailCol - 1] ?? '').trim() : '';
      if (name || email) volunteers.push({ row: i + 2, name, email });
    });
  }

  const asks: string[] = [];
  const menu = ss.getSheetByName(CONFIG.tabNames.menu);
  if (menu) {
    for (const row of sheetRows(menu)) {
      const ask = String(row[0] ?? '').trim();
      if (ask) asks.push(ask);
    }
  }

  return { volunteers, asks };
}

/** Open the New Action dialog, pre-loaded with volunteers and Asks. */
export function openActionDialog(services: Services): void {
  services.showHtml(actionDialogHtml(actionOptions(services)), 'New Action');
}

/** Add an Ask to the Menu of Asks so it joins the dropdown. No-op if present. */
function ensureAsk(services: Services, ask: string): void {
  const menu = services.ss.getSheetByName(CONFIG.tabNames.menu);
  if (!menu) return;
  const known = sheetRows(menu).some((row) => norm(String(row[0] ?? '')) === norm(ask));
  if (known) return;
  const { firstFreeRow } = sheetExtent(menu, CONFIG.trackerColumns.ask);
  menu.getRange(firstFreeRow, 1).setValue(ask);
}

/**
 * Point the Tracker's Ask dropdown at a generous span of the Menu of Asks.
 * The span reaches future Menu rows, and re-running it before a write means an
 * Ask added since Set Up (typed or via the dialog) is accepted rather than
 * rejected by the strict validation list.
 */
function syncAskValidation(services: Services, menu: SheetLike, tracker: SheetLike): void {
  const askCol = columnIndexOf(tracker, CONFIG.trackerColumns.ask);
  if (askCol <= 0) return;
  const range = menu.getRange(2, 1, 1000, 1);
  const rule = services
    .makeValidation()
    .requireValueInRange(range, true)
    .setAllowInvalid(false)
    .build();
  // Cover future Tracker rows too, so the dropdown reaches every Action row.
  tracker.getRange(2, askCol, 1000, 1).setDataValidation(rule);
}

/**
 * Create an Action from the dialog: write the Tracker row and schedule its
 * Pre/Post reminders. Returns rather than throws so the dialog can show the
 * problem inline.
 */
export function createActionFromDialog(services: Services, input: ActionSubmission): ActionResult {
  const directory = services.ss.getSheetByName(CONFIG.tabNames.directory);
  const tracker = services.ss.getSheetByName(CONFIG.tabNames.tracker);
  if (!directory || !tracker) {
    return { ok: false, message: 'Run Set Up first — a required tab is missing.' };
  }

  // The Campaign Calendar is resolved lazily and per read; check it up front so
  // a missing share reports inline rather than mid-write (or on a personal one).
  try {
    campaignCalendar(services);
  } catch (error) {
    return { ok: false, message: error instanceof Error ? error.message : String(error) };
  }

  const row = Math.trunc(Number(input.volunteerRow));
  const name =
    row >= 2 ? String(cellValue(directory, row, CONFIG.requiredFields.name) ?? '').trim() : '';
  const email =
    row >= 2 ? String(cellValue(directory, row, CONFIG.requiredFields.email) ?? '').trim() : '';
  if (!name && !email) {
    return { ok: false, message: 'Choose a volunteer from the list.' };
  }

  const ask = String(input.ask ?? '').trim();
  if (!ask) {
    return { ok: false, message: 'Choose an Ask or type a new one.' };
  }

  const deadline = parseDateValue(input.deadline);
  if (!deadline) {
    return { ok: false, message: 'Pick a Deadline date.' };
  }

  ensureAsk(services, ask);

  // The dialog reads Asks straight from the Menu, but the Tracker's Ask cell is
  // a strict dropdown: re-point it so an Ask added since Set Up is accepted.
  const menu = services.ss.getSheetByName(CONFIG.tabNames.menu);
  if (menu) syncAskValidation(services, menu, tracker);

  const headers = trackerHeaders();
  const rowData: (string | Date)[] = headers.map(() => '');
  rowData[headers.indexOf(CONFIG.trackerColumns.name)] = name;
  rowData[headers.indexOf(CONFIG.trackerColumns.email)] = email;
  rowData[headers.indexOf(CONFIG.trackerColumns.ask)] = ask;
  rowData[headers.indexOf(CONFIG.trackerColumns.deadline)] = deadline;
  const { firstFreeRow } = sheetExtent(tracker, CONFIG.trackerColumns.name);
  rowData.forEach((value, i) => {
    tracker.getRange(firstFreeRow, i + 1).setValue(value);
  });

  scheduleTouches(services, tracker, firstFreeRow, deadline);

  toast(services, `Action created for ${name} — ${ask}, deadline ${fmtDate(deadline)}.`, 8);
  return { ok: true, message: '' };
}
