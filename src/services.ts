/**
 * Assembles the live Google services for `automation.ts`. Kept out of
 * `index.ts` so it can be unit-tested: trigger executions have no editor
 * session, so building services must never resolve the UI.
 */

import type {
  CalendarsLike,
  FormsLike,
  Services,
  SheetsLike,
  UiLike,
  ValidationBuilderLike,
} from './automation';
import { campaignCalendar } from './automation';

/** The Google globals `buildServices` reads, passed in so tests can drive it. */
export interface GoogleGlobals {
  getActiveSpreadsheet(): SheetsLike;
  getUi(): UiLike;
  calendars: CalendarsLike;
  getScript(): Services['script'];
  newDataValidation(): ValidationBuilderLike;
  forms: FormsLike;
  showHtml(html: string, title: string): void;
  flush(): void;
  log(message: string): void;
}

export function buildServices(google: GoogleGlobals): Services {
  const services: Services = {
    ss: google.getActiveSpreadsheet(),
    calendars: google.calendars,
    get ui() {
      return google.getUi();
    },
    // Lazy: Set Up creates the calendar after services are built, and a trigger
    // resolves it per execution.
    get cal() {
      return campaignCalendar(services);
    },
    script: google.getScript(),
    forms: google.forms,
    flush: () => google.flush(),
    makeValidation: () => google.newDataValidation(),
    showHtml: (html, title) => google.showHtml(html, title),
    log: (message) => google.log(message),
  };
  return services;
}
