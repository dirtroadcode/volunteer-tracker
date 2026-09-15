/**
 * Assembles the live Google services for `automation.ts`. Kept out of
 * `index.ts` so it can be unit-tested: trigger executions have no editor
 * session, so building services must never resolve the UI.
 */
import type {
  CalendarLike,
  FormsLike,
  Services,
  SheetsLike,
  UiLike,
  ValidationBuilderLike,
} from './automation';

/** The Google globals `buildServices` reads, passed in so tests can drive it. */
export interface GoogleGlobals {
  getActiveSpreadsheet(): SheetsLike;
  getUi(): UiLike;
  getDefaultCalendar(): CalendarLike;
  getScript(): Services['script'];
  newDataValidation(): ValidationBuilderLike;
  forms: FormsLike;
  showHtml(html: string, title: string): void;
  flush(): void;
  log(message: string): void;
}

export function buildServices(google: GoogleGlobals): Services {
  return {
    ss: google.getActiveSpreadsheet(),
    get ui() {
      return google.getUi();
    },
    cal: google.getDefaultCalendar(),
    script: google.getScript(),
    forms: google.forms,
    flush: () => google.flush(),
    makeValidation: () => google.newDataValidation(),
    showHtml: (html, title) => google.showHtml(html, title),
    log: (message) => google.log(message),
  };
}
