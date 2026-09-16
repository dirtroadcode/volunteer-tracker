/**
 * Volunteerist — Apps Script entry points.
 *
 * Thin wiring: assembles the live Google services and delegates to the
 * injected, unit-tested logic in `automation.ts`. Coordinators interact via
 * the "Volunteer Tools" menu; the only code-touching moment ever is the
 * one-time Set Up.
 */
import type {
  ActionResult,
  ActionSubmission,
  FormItemLike,
  FormLinkLike,
  FormsLike,
  Services,
} from './automation';
import {
  createActionFromDialog,
  handleEdit,
  handleFormSubmit,
  openActionDialog,
  runSetupAutomation,
  showSignUpForm,
} from './automation';
import { buildServices } from './services';

// ---------------------------------------------------------------- menus

function onOpen(): void {
  const menu = SpreadsheetApp.getUi()
    .createMenu('Volunteer Tools')
    .addItem('New Action', 'newAction')
    .addSeparator();
  // A form has no tab to click, so the route to it only exists while one is linked.
  if (SpreadsheetApp.getActiveSpreadsheet().getFormUrl() !== null) {
    menu.addItem('Go to Sign-Up Form', 'goToSignUpForm');
  }
  menu.addItem('Set Up', 'setupAutomation').addToUi();
}

// ------------------------------------------------------------ wiring

function services(): Services {
  return buildServices({
    getActiveSpreadsheet: () => SpreadsheetApp.getActiveSpreadsheet(),
    getUi: () => SpreadsheetApp.getUi(),
    calendars: {
      createCalendar: (name, options) => CalendarApp.createCalendar(name, options ?? {}),
      getCalendarById: (id) => CalendarApp.getCalendarById(id),
      getOwnedCalendarsByName: (name) => CalendarApp.getOwnedCalendarsByName(name),
    },
    getScript: () => ScriptApp,
    newDataValidation: () => SpreadsheetApp.newDataValidation(),
    showHtml: (html, title) =>
      SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), title),
    flush: () => SpreadsheetApp.flush(),
    log: (message: string) => Logger.log(message),
    forms: formService(),
  });
}

/** The two things we ever do with forms: open one by URL, and build one from scratch. */
function formService(): FormsLike {
  const link = (form: GoogleAppsScript.Forms.Form): FormLinkLike => {
    const item = (
      i: GoogleAppsScript.Forms.TextItem | GoogleAppsScript.Forms.ParagraphTextItem,
    ) => {
      const handle: FormItemLike = {
        setTitle: (title) => {
          i.setTitle(title);
          return handle;
        },
        setRequired: (required) => {
          i.setRequired(required);
          return handle;
        },
        setHelpText: (text) => {
          i.setHelpText(text);
          return handle;
        },
      };
      return handle;
    };
    return {
      addTextItem: () => item(form.addTextItem()),
      addParagraphTextItem: () => item(form.addParagraphTextItem()),
      setDestination: (spreadsheetId: string) =>
        form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId),
      removeDestination: () => form.removeDestination(),
    };
  };
  return {
    openByUrl: (url: string) => link(FormApp.openByUrl(url)),
    create: (title: string) => link(FormApp.create(title)),
  };
}

// ------------------------------------------------------------ actions

function setupAutomation(): void {
  runSetupAutomation(services());
}

function newAction(): void {
  openActionDialog(services());
}

/** Called from the New Action dialog via `google.script.run`. */
function submitActionFromDialog(input: ActionSubmission): ActionResult {
  return createActionFromDialog(services(), input);
}

function goToSignUpForm(): void {
  showSignUpForm(services());
}

// ------------------------------------------------------------ triggers

function onFormSubmitInstalled(e: GoogleAppsScript.Events.SheetsOnFormSubmit): void {
  handleFormSubmit(services(), e as unknown as Parameters<typeof handleFormSubmit>[1]);
}

// Installed by Set Up as an installable trigger. Naming it `onEdit` would make
// it a simple trigger, which cannot access CalendarApp.
function onEditInstalled(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  handleEdit(services(), e as unknown as Parameters<typeof handleEdit>[1]);
}
