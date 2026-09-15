/**
 * In-memory mock of the Google services Volunteerist's glue depends on,
 * shaped to the structural interfaces in src/automation.ts (SheetLike,
 * SheetsLike, RangeLike, CalendarLike, UiLike, ScriptLike, FormsLike,
 * ValidationBuilderLike). No global stubbing — services are injected.
 */
import type {
  CalendarEventLike,
  CalendarLike,
  FormItemLike,
  FormLinkLike,
  FormsLike,
  RangeLike,
  ScriptLike,
  Services,
  SheetLike,
  SheetsLike,
  UiLike,
  ValidationBuilderLike,
} from '../../src/automation';

export interface TabSpec {
  name: string;
  headers: string[];
  rows?: unknown[][];
}

/** Cell presentation recorded by the mock so tests can assert formatting. */
export interface CellStyle {
  wrap?: boolean;
  fontWeight?: string;
  background?: string;
  fontColor?: string;
  fontSize?: number;
  horizontalAlignment?: string;
  verticalAlignment?: string;
}

export class MockRange implements RangeLike {
  constructor(
    private sheet: MockSheet,
    private row: number,
    private col: number,
    private numRows = 1,
    private numCols = 1,
  ) {}

  getValue(): unknown {
    return this.sheet.raw(this.row, this.col);
  }

  getValues(): unknown[][] {
    const out: unknown[][] = [];
    for (let r = 0; r < this.numRows; r++) {
      const line: unknown[] = [];
      for (let c = 0; c < this.numCols; c++) {
        line.push(this.sheet.raw(this.row + r, this.col + c));
      }
      out.push(line);
    }
    return out;
  }

  setValue(value: unknown): this {
    this.sheet.assertValid(this.row, this.col, value);
    this.sheet.write(this.row, this.col, value);
    return this;
  }

  setNumberFormat(format: string): this {
    this.sheet.applyNumberFormat(this.row, this.col, this.numRows, format);
    return this;
  }

  setDataValidation(rule: unknown): this {
    this.sheet.applyValidation(this.row, this.col, this.numRows, this.numCols, rule);
    return this;
  }

  setWrap(isWrapEnabled: boolean): this {
    this.sheet.applyStyle(this.row, this.col, this.numRows, this.numCols, {
      wrap: isWrapEnabled,
    });
    return this;
  }

  setFontWeight(weight: 'normal' | 'bold' | null): this {
    this.sheet.applyStyle(this.row, this.col, this.numRows, this.numCols, {
      fontWeight: weight ?? undefined,
    });
    return this;
  }

  setBackground(color: string | null): this {
    this.sheet.applyStyle(this.row, this.col, this.numRows, this.numCols, {
      background: color ?? undefined,
    });
    return this;
  }

  setFontColor(color: string | null): this {
    this.sheet.applyStyle(this.row, this.col, this.numRows, this.numCols, {
      fontColor: color ?? undefined,
    });
    return this;
  }

  setFontSize(size: number): this {
    this.sheet.applyStyle(this.row, this.col, this.numRows, this.numCols, { fontSize: size });
    return this;
  }

  setHorizontalAlignment(alignment: 'left' | 'center' | 'normal' | 'right' | null): this {
    this.sheet.applyStyle(this.row, this.col, this.numRows, this.numCols, {
      horizontalAlignment: alignment ?? undefined,
    });
    return this;
  }

  setVerticalAlignment(alignment: 'top' | 'middle' | 'bottom' | null): this {
    this.sheet.applyStyle(this.row, this.col, this.numRows, this.numCols, {
      verticalAlignment: alignment ?? undefined,
    });
    return this;
  }

  getRow(): number {
    return this.row;
  }

  getColumn(): number {
    return this.col;
  }

  getSheet(): SheetLike {
    return this.sheet;
  }

  activate(): void {}
}

/** Mock stand-in for Sheets' rendered text metrics. */
const COLUMN_CHAR_WIDTH = 7;
const COLUMN_PADDING = 10;

/** Width the mock reports for a column sized to content `text`. */
export function columnWidthFor(text: string): number {
  return text.length * COLUMN_CHAR_WIDTH + COLUMN_PADDING;
}

/** Spreadsheet-style column letter for a 1-based column index. */
export function columnLetter(col: number): string {
  let n = col;
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/** A built data-validation rule — enough to enforce a strict list of values. */
export interface MockDataValidation {
  kind: 'range' | 'checkbox' | 'formula' | null;
  source: RangeLike | null;
  allowInvalid: boolean;
}

export class MockSheet implements SheetLike {
  readonly id: number;
  readonly grid: unknown[][] = [];
  name: string;
  /** URL of the form linked to this tab, set when a form is linked. */
  formUrl: string | null = null;
  private frozenRows = 0;
  private activeSelectionRow = 2;
  /** Number formats by "row:col", and the deepest row any format touches. */
  private readonly formats = new Map<string, string>();
  private formattedMaxRow = 0;
  /** Cell presentation by "row:col". */
  private readonly styles = new Map<string, CellStyle>();
  /** Data-validation rules by "row:col"; setValue enforces strict list rules. */
  private readonly validations = new Map<string, MockDataValidation>();
  /** Explicit column widths; unset columns fall back to the Sheets default. */
  private readonly columnWidths = new Map<number, number>();

  constructor(name: string, id: number, headers: string[], rows: unknown[][] = []) {
    this.name = name;
    this.id = id;
    // A blank sheet has no rows; only the header row counts as content.
    if (headers.length > 0) this.grid.push([...headers]);
    rows.forEach((r) => {
      this.grid.push([...r]);
    });
  }

  raw(row: number, col: number): unknown {
    return this.grid[row - 1]?.[col - 1] ?? '';
  }

  write(row: number, col: number, value: unknown): void {
    while (this.grid.length < row) this.grid.push([]);
    const target = this.grid[row - 1];
    if (!target) throw new Error(`grid row ${row} is missing`);
    while (target.length < col) target.push('');
    target[col - 1] = value;
  }

  /** A sheet's last row counts formatted-but-empty rows, not just written cells. */
  getLastRow(): number {
    return Math.max(this.grid.length, this.formattedMaxRow);
  }

  applyNumberFormat(row: number, col: number, numRows: number, format: string): void {
    for (let r = 0; r < numRows; r++) {
      this.formats.set(`${row + r}:${col}`, format);
      this.formattedMaxRow = Math.max(this.formattedMaxRow, row + r);
    }
  }

  /** Test-only: the number format currently applying to a cell. */
  getNumberFormat(row: number, col: number): string {
    return this.formats.get(`${row}:${col}`) ?? '';
  }

  applyStyle(row: number, col: number, numRows: number, numCols: number, patch: CellStyle): void {
    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        const key = `${row + r}:${col + c}`;
        this.styles.set(key, { ...this.styles.get(key), ...patch });
      }
    }
  }

  /** Test-only: the recorded presentation of a cell. */
  getStyle(row: number, col: number): CellStyle {
    return this.styles.get(`${row}:${col}`) ?? {};
  }

  /** Record a data-validation rule across a block of cells. */
  applyValidation(row: number, col: number, numRows: number, numCols: number, rule: unknown): void {
    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        this.validations.set(`${row + r}:${col + c}`, rule as MockDataValidation);
      }
    }
  }

  /**
   * Mirror Google: writing to a cell with a strict list validation throws
   * when the value is not among the source range's options.
   */
  assertValid(row: number, col: number, value: unknown): void {
    const rule = this.validations.get(`${row}:${col}`);
    if (!rule) return;
    if (rule.kind !== 'range' || rule.allowInvalid || !rule.source) return;
    const text = String(value ?? '').trim();
    if (text === '') return;
    const allowed = rule.source
      .getValues()
      .flat()
      .map((v) => String(v ?? '').trim());
    if (!allowed.includes(text)) {
      throw new Error(
        `The data you entered in cell ${columnLetter(col)}${row} violates the data validation rules set on this cell.`,
      );
    }
  }

  /** Test-only: the validation rule currently applying to a cell. */
  getValidation(row: number, col: number): MockDataValidation | undefined {
    return this.validations.get(`${row}:${col}`);
  }

  getColumnWidth(columnPosition: number): number {
    return this.columnWidths.get(columnPosition) ?? 100;
  }

  setColumnWidth(columnPosition: number, width: number): void {
    this.columnWidths.set(columnPosition, width);
  }

  /**
   * Sheets autofits to the widest *rendered* line, so a wrapped cell shows
   * only the fragment that fits its current width and a column sized this way
   * stays pinned to the wrapped text rather than the value.
   */
  autoResizeColumns(startColumn: number, numColumns: number): void {
    for (let c = 0; c < numColumns; c++) {
      const col = startColumn + c;
      const width = this.getColumnWidth(col);
      const visibleChars = Math.max(1, Math.floor((width - COLUMN_PADDING) / COLUMN_CHAR_WIDTH));
      let longest = '';
      this.grid.forEach((row, i) => {
        const text = String(row[col - 1] ?? '');
        const shown = this.getStyle(i + 1, col).wrap ? text.slice(0, visibleChars) : text;
        if (shown.length > longest.length) longest = shown;
      });
      this.setColumnWidth(col, columnWidthFor(longest));
    }
  }

  getLastColumn(): number {
    return Math.max(0, ...this.grid.map((r) => r.length));
  }

  getRange(row: number, col: number, numRows = 1, numCols = 1): RangeLike {
    return new MockRange(this, row, col, numRows, numCols);
  }

  /** Appends below the last row, so phantom formatted rows push it down too. */
  appendRow(values: unknown[]): void {
    const target = this.getLastRow() + 1;
    while (this.grid.length < target) this.grid.push([]);
    this.grid[target - 1] = [...values];
  }

  getName(): string {
    return this.name;
  }

  setName(name: string): void {
    this.name = name;
  }

  getFormUrl(): string | null {
    return this.formUrl;
  }

  getSheetId(): number {
    return this.id;
  }

  setFrozenRows(rows: number): void {
    this.frozenRows = rows;
  }

  /** Test-only: how many top rows the sheet has pinned. */
  getFrozenRows(): number {
    return this.frozenRows;
  }

  activate(): void {}

  getParent(): SheetsLike {
    return this.spreadsheet;
  }

  getActiveCell(): RangeLike {
    return new MockRange(this, this.activeSelectionRow, 1);
  }

  setActiveRange(range: RangeLike): void {
    this.activeSelectionRow = range.getRow();
  }

  /** Test-only: pick which row counts as the active selection. */
  selectRow(row: number): void {
    this.activeSelectionRow = row;
  }

  spreadsheet!: MockSpreadsheet;
}

export class MockCalendarEvent implements CalendarEventLike {
  constructor(
    readonly id: string,
    public title: string,
    public start: Date,
    public end: Date,
    public description: string,
    private owner: MockCalendar,
  ) {}

  getId(): string {
    return this.id;
  }
  setTitle(title: string): void {
    this.title = title;
  }

  setDescription(description: string): void {
    this.description = description;
  }

  setTime(start: Date, end: Date): void {
    this.start = start;
    this.end = end;
  }

  deleteEvent(): void {
    this.owner.remove(this.id);
  }
}

export class MockCalendar implements CalendarLike {
  readonly events: MockCalendarEvent[] = [];
  private byId = new Map<string, MockCalendarEvent>();
  private counter = 0;

  createEvent(
    title: string,
    start: Date,
    end: Date,
    options?: { description?: string },
  ): MockCalendarEvent {
    const ev = new MockCalendarEvent(
      `ev${++this.counter}`,
      title,
      start,
      end,
      options?.description ?? '',
      this,
    );
    this.events.push(ev);
    this.byId.set(ev.id, ev);
    return ev;
  }

  getEventById(id: string): CalendarEventLike | null {
    return this.byId.get(id) ?? null;
  }

  remove(id: string): void {
    this.byId.delete(id);
    const i = this.events.findIndex((e) => e.id === id);
    if (i >= 0) this.events.splice(i, 1);
  }
}

export class MockUi implements UiLike {
  readonly ButtonSet = { OK: 'ok', OK_CANCEL: 'ok-cancel' };
  readonly Button = { OK: 'ok', CANCEL: 'cancel' };
  readonly prompts: { title: string; text: string; buttons: unknown }[] = [];
  readonly alerts: { title: string; text: string; buttons: unknown }[] = [];
  readonly dialogs: { html: string; title: string }[] = [];
  private promptQueue: { text: string; button: string }[] = [];
  private alertQueue: string[] = [];

  /** Test-only: queue a response for the next prompt(). */
  queuePrompt(text: string, button: 'ok' | 'cancel' = 'ok'): void {
    this.promptQueue.push({ text, button: button === 'ok' ? this.Button.OK : this.Button.CANCEL });
  }

  /** Test-only: queue a button response for the next alert(). */
  queueAlert(button: 'ok' | 'cancel' = 'ok'): void {
    this.alertQueue.push(button === 'ok' ? this.Button.OK : this.Button.CANCEL);
  }

  prompt(title: string, text: string, buttons: unknown) {
    const next = this.promptQueue.shift() ?? { text: '', button: this.Button.CANCEL };
    this.prompts.push({ title, text, buttons });
    return { getSelectedButton: () => next.button, getResponseText: () => next.text };
  }

  alert(title: string, text: string, buttons: unknown): unknown {
    this.alerts.push({ title, text, buttons });
    return this.alertQueue.shift() ?? this.ButtonSet.OK;
  }

  showHtml(html: string, title: string): void {
    this.dialogs.push({ html, title });
  }
}

export class MockScript implements ScriptLike {
  readonly triggers: string[] = [];

  getProjectTriggers() {
    return this.triggers.map((handler) => ({ getHandlerFunction: () => handler }));
  }

  newTrigger(handler: string) {
    return {
      forSpreadsheet: () => ({
        onFormSubmit: () => ({
          create: () => {
            this.triggers.push(handler);
          },
        }),
        onEdit: () => ({
          create: () => {
            this.triggers.push(handler);
          },
        }),
      }),
    };
  }
}

export interface MockFormQuestion {
  title: string;
  required: boolean;
  type: 'text' | 'paragraph';
  helpText: string;
}

/** A form handle: item builders record questions; linking creates the response tab. */
class MockFormLink implements FormLinkLike {
  private readonly items: MockFormQuestion[] = [];

  constructor(
    private owner: MockForms,
    private url: string,
    /** Created forms record their built questions here; opened forms fall back to the owner's preset questions. */
    private sink: MockFormQuestion[] | null,
  ) {}

  addTextItem(): FormItemLike {
    return this.addItem('text');
  }

  addParagraphTextItem(): FormItemLike {
    return this.addItem('paragraph');
  }

  private addItem(type: MockFormQuestion['type']): FormItemLike {
    const q: MockFormQuestion = { title: '', required: false, type, helpText: '' };
    this.items.push(q);
    this.sink?.push(q);
    const item: FormItemLike = {
      setTitle: (title) => {
        q.title = title;
        return item;
      },
      setRequired: (required) => {
        q.required = required;
        return item;
      },
      setHelpText: (text) => {
        q.helpText = text;
        return item;
      },
    };
    return item;
  }

  setDestination(spreadsheetId: string): void {
    const titles = this.items.length > 0 ? this.items.map((q) => q.title) : this.owner.questions;
    this.owner.link(spreadsheetId, this.url, titles);
  }

  removeDestination(): void {
    this.owner.unlink(this.url);
  }
}

export class MockForms implements FormsLike {
  /** The spreadsheet the last linked form now writes to (test-visible). */
  linkedSpreadsheetId: string | null = null;
  /** Makes ss.getFormUrl() report the linked form. */
  linkedFormUrl: string | null = null;
  /** The linked form's question titles — the response tab's columns after Timestamp. */
  questions: string[] = [];
  /** The form built by Set Up via create() — null until one is generated. */
  created: { title: string; questions: MockFormQuestion[] } | null = null;
  /** URL assigned to the form create() made. */
  createdFormUrl: string | null = null;
  private spreadsheet!: MockSpreadsheet;
  private createdCount = 0;

  openByUrl(url: string): FormLinkLike {
    return new MockFormLink(this, url, null);
  }

  create(title: string): FormLinkLike {
    const url = `https://forms.example/created-${++this.createdCount}`;
    this.created = { title, questions: [] };
    this.createdFormUrl = url;
    return new MockFormLink(this, url, this.created.questions);
  }

  /** Visible to MockFormLink, which owns the linking flow. */
  link(spreadsheetId: string, url: string, questionTitles: string[]): void {
    this.linkedSpreadsheetId = spreadsheetId;
    this.linkedFormUrl = url;
    this.spreadsheet.formUrl = url;
    // Mirror Google: linking creates a response tab, named "Form
    // Responses N" (next free number), with a column per question.
    let n = 1;
    while (this.spreadsheet.getSheetByName(`Form Responses ${n}`)) n++;
    const tab = this.spreadsheet.insertSheet(`Form Responses ${n}`);
    tab.formUrl = url;
    tab.appendRow(['Timestamp', ...questionTitles]);
  }

  spreadsheetRef(ss: MockSpreadsheet): void {
    this.spreadsheet = ss;
  }

  /** Visible to MockFormLink; mirrors the real API: unlink the form, its response tab stays but loses the link. */
  readonly unlinkedUrls: string[] = [];

  unlink(url: string): void {
    this.unlinkedUrls.push(url);
    if (this.spreadsheet.formUrl === url) this.spreadsheet.formUrl = null;
    for (const sheet of this.spreadsheet.getSheets()) {
      if (sheet.getFormUrl() === url) sheet.formUrl = null;
    }
  }
}

export class MockValidationBuilder implements ValidationBuilderLike {
  readonly specs: string[] = [];
  private kind: MockDataValidation['kind'] = null;
  private source: RangeLike | null = null;
  private allowInvalid = false;

  requireValueInRange(range: unknown): this {
    this.specs.push('range');
    this.kind = 'range';
    this.source = range as RangeLike;
    return this;
  }

  requireCheckbox(): this {
    this.specs.push('checkbox');
    this.kind = 'checkbox';
    return this;
  }

  requireFormulaSatisfied(formula: string): this {
    this.specs.push(`formula:${formula}`);
    this.kind = 'formula';
    return this;
  }

  setAllowInvalid(allow: boolean): this {
    this.allowInvalid = allow;
    return this;
  }

  setHelpText(): this {
    return this;
  }

  build(): MockDataValidation {
    return { kind: this.kind, source: this.source, allowInvalid: this.allowInvalid };
  }
}

export class MockSpreadsheet implements SheetsLike {
  readonly id = 'ss-test-id';
  formUrl: string | null = null;
  readonly sheets: MockSheet[] = [];
  private activeIndex = 0;
  private sheetSeq = 0;

  constructor(tabs: TabSpec[] = []) {
    for (const t of tabs) this.addSheet(t);
  }

  private addSheet(tab: TabSpec): MockSheet {
    const s = new MockSheet(tab.name, ++this.sheetSeq, tab.headers, tab.rows ?? []);
    s.spreadsheet = this;
    this.sheets.push(s);
    return s;
  }

  getSheetByName(name: string): MockSheet | null {
    return this.sheets.find((s) => s.name === name) ?? null;
  }

  insertSheet(name: string): MockSheet {
    return this.addSheet({ name, headers: [] });
  }

  getSheets(): MockSheet[] {
    return [...this.sheets];
  }

  getId(): string {
    return this.id;
  }

  getFormUrl(): string | null {
    return this.formUrl;
  }

  getSheetById(id: number): MockSheet | null {
    return this.sheets.find((s) => s.id === id) ?? null;
  }

  getActiveSheet(): MockSheet {
    const sheet = this.sheets[this.activeIndex] ?? this.sheets[0];
    if (!sheet) throw new Error('mock workbook has no sheets');
    return sheet;
  }

  toast(): void {}

  setActiveSheet(sheet: MockSheet): void {
    this.activeIndex = this.sheets.indexOf(sheet);
  }

  flush(): void {
    // no-op — Apps Script batches are synchronous in the fake
  }

  deleteSheet(sheet: MockSheet): void {
    // Mirror Google: a tab with a linked form cannot be deleted — the form
    // must be unlinked (removeDestination) first.
    if (sheet.getFormUrl() !== null) {
      throw new Error(
        'You cannot delete a sheet with a linked form. Please unlink the form first.',
      );
    }
    const i = this.sheets.indexOf(sheet);
    if (i >= 0) this.sheets.splice(i, 1);
  }
}

export interface World {
  services: Services;
  ss: MockSpreadsheet;
  sheets: MockSheet[];
  calendar: MockCalendar;
  ui: MockUi;
  script: MockScript;
  forms: MockForms;
  logs: string[];
  byName(name: string): MockSheet;
}

/** Build a fresh world, optionally starting with the given tabs. */
export function createWorld(tabs: TabSpec[] = []): World {
  const ss = new MockSpreadsheet(tabs);
  const calendar = new MockCalendar();
  const ui = new MockUi();
  const script = new MockScript();
  const forms = new MockForms();
  forms.spreadsheetRef(ss);
  const logs: string[] = [];
  const services: Services = {
    ss,
    ui,
    cal: calendar,
    script,
    forms,
    flush: () => {},
    makeValidation: () => new MockValidationBuilder(),
    showHtml: (html, title) => ui.showHtml(html, title),
    log: (message) => logs.push(message),
  };
  return {
    services,
    ss,
    sheets: ss.sheets,
    calendar,
    ui,
    script,
    forms,
    logs,
    byName: (name: string) => {
      const s = ss.getSheetByName(name);
      if (!s) throw new Error(`tab "${name}" not found`);
      return s;
    },
  };
}
