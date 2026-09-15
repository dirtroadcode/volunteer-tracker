# Volunteerist — PRD

## Problem Statement

Volunteer coordinators collect signups through a Google Form, then each new
volunteer needs (1) a first contact, (2) connection to a concrete activity
the campaign needs done, and (3) a follow-up that checks how it went and
shows appreciation. These touchpoints are tracked ad hoc — memory, notes, an
unstructured spreadsheet — so volunteers fall through the cracks between
signup and first contact, or after they've done a task once and are never
asked again. The coordinator needs a reliable system that requires **no
technical skill** to operate day-to-day: no installation, no hosting, no
code, ever.

## Solution

A **copyable Sample**: one Spreadsheet (with a bound script) plus one intake
Form, built once and distributed to coordinators, who make a copy for their
own campaign. Each copy keeps the Form's raw response log separate from a
script-written **Directory** roster, a **Menu of Asks** the coordinator edits
by typing, a **Tracker** of Actions (one volunteer↔Ask commitment per row),
and a **Settings** tab of four timing knobs. Intake writes a Directory row
and schedules the first contact automatically; the coordinator's deadline
entry derives the pre-ask and post-ask reminders, all landing on the
coordinator's **own Google Calendar**. The only setup is one menu action that links the form (the
Sample Form *or* the coordinator's own, validated against a three-field
contract), scaffolds the tabs, and installs the submit and edit triggers. The tool
reminds the coordinator and never messages volunteers; appreciation is
concrete reminders around each ask's deadline, which is what keeps a
one-time volunteer coming back.

## User Stories

1. As a coordinator, when someone submits my intake Form, I want their row
   to appear in my Directory automatically, so that I never retype signup
   data.
2. As a coordinator, I want each submission to add a Directory row and
   schedule a first contact, so that a signup never waits unheard-from and a
   returning volunteer is simply given another Action.
3. As a coordinator, I want first contact scheduled automatically when
   someone signs up and the date editable afterward, so that nobody waits
   unheard-from between signup and first contact.
4. As a coordinator, I want to build and edit my Menu of Asks by typing
   directly into a tab, so that maintaining it never requires code (my
   campaign has no developers).
5. As a coordinator, I want duplicate ask names flagged with a warning, so
   that my dropdown and follow-up context stay unambiguous.
6. As a coordinator, I want to create an Action for a Directory volunteer
   with one menu action that pre-fills their name and email, so that
   connecting people to work is two clicks.
7. As a coordinator, I want to pick an Ask from a dropdown rather than
   typing it, so that activity names can't typo-drift.
8. As a coordinator, I want to type a single Deadline for an Action and
   have the Pre-touch (2 days before) and Post-touch (3 days after) dates
   derived automatically, so that one date schedules the whole reminder
   loop.
9. As a coordinator, I want those reminders on my own Google Calendar, so
   that I can't forget them and nothing messages the volunteer.
10. As a coordinator, I want each calendar event to say who it's for and
    how to reach them (name, email, phone, About) as well as why (ask,
    deadline), so that I never have to open the spreadsheet to know what to
    do.
11. As a coordinator, I want the reminders to carry the appreciation intent
    ("see you there, thanks for signing up" before; "how did it go? thank
    you, come to our next thing" after), so that appreciation — the
    retention loop — actually happens.
12. As a coordinator, I want to edit a date and have the event **move**
    rather than duplicate, so that corrections never double-book my
    calendar.
13. As a coordinator, I want to clear a deadline and have its reminders and
    derived dates removed, so that cancelled plans leave no ghost events.
14. As a coordinator, I want to jot notes and tick Done on a finished
    Action, so that there's a record and I can hide wrapped-up work.
15. As a coordinator, when a Returning volunteer comes back for another
    ask, I want them to get a fresh Action with their history visible,
    so that repeat engagement is effortless.
16. As a coordinator, I want to use my **own** intake form as long as it
    asks for Name, Email, and Phone, with setup telling me loudly if fields
    are missing, so that I'm not locked into the Sample Form.
17. As a coordinator, I want all setup to happen through one menu action
    that runs once, so that I never configure anything again.
18. As a coordinator, I want the one-time "unverified app" consent screen
    explained and to only ever see it once, so that it doesn't scare me off.
19. As a coordinator, I want copying the Sample to bring everything along —
    script, tabs, configuration — so that starting my campaign is one File
    → Make a copy.
20. As a trainer, I want the Sample to ship with example Asks so that the
    training exercise ("make the menu yours") is concrete, not blank.
21. As a trainer, I want a demo-ready Sample (a few seeded volunteers and a
    Action) so that the live walkthrough is fast and visual.
22. As a toolkit maintainer, I want the non-Google logic unit-tested, so
    that a distributed product doesn't break silently for every copier.
23. As a coordinator, I want the tool to always only remind me, never
    message volunteers, so that I stay in control of what volunteers hear.
24. As a coordinator, I want four timing knobs (first-contact delay,
    pre/post-touch offsets, reminder hour) in a Settings tab, so that I can
    retune cadence without code, applied going forward only.

## Implementation Decisions

- **Copy-based distribution (ADR 0001).** The Sample Spreadsheet is the copy
  unit; the bound script and tabs travel with the copy. Installable triggers
  never survive a copy, so one menu action ("Set Up") re-installs
  them and links the form per copy. Calendar events use each copy's default
  calendar, giving every campaign isolated reminders for free. Trigger
  handlers run without an editor session, so they never touch the UI.
- **Form/Directory/Tracker split (ADR 0004, supersedes ADR 0002).** Form
  Responses is the Form's own tab, read-only to the script; Directory is the
  script-written roster mirroring only the declared fields and carrying the
  intake fields; Menu of Asks is the directly-editable registry feeding the
  Tracker's Ask dropdown; Tracker holds one Action per row (volunteer, Ask,
  Deadline, derived Pre/Post touch, notes, done); Settings holds the four
  timing knobs. Rejected mirroring every submitted column into the Directory
  (it re-imports Form-driven drift) and a separate Intake tab with no roster.
- **Script-owned tab presentation.** Set Up bolds, tints, and freezes the
  header row, dims machine columns, centers dates, and sizes every column to
  its unwrapped content (capped at `maxColumnWidth`) before turning wrapping
  on. Row banding is deliberately not used: alternating colors are a live rule
  bound to a fixed range that drifts from rows the script writes later.
- **Coordinator-only boundary (ADR 0003).** No MailApp/GmailApp anywhere;
  all three touchpoints are calendar events.
- **Header-based column lookup.** All reads/writes locate columns by header
  text via a single `CONFIG` map; reordering columns is safe; renaming
  without updating `CONFIG` merely makes setup re-add the missing column
  (never silently merges).
- **Required-fields contract.** `Name`, `Email`, and `Phone` (header text).
  Setup validates a (BYO) form's response headers against it and reports
  exactly what's missing; forms link via `FormApp.setDestination`.
- **No classification on the Form's log.** Every submission adds a Directory
  row; there is no New/Returning flag. A returning volunteer is found in the
  Directory and given another Action, and the Form's own tab is never written
  to.
- **Idempotent events.** The event id is stored raw in its column;
  re-editing a date moves the existing event, never creates a second;
  clearing a date deletes the event(s) and derived cells.
- **Trigger split.** Both the submit and edit handlers are installable
  triggers, installed by Set Up idempotently. The edit handler is deliberately
  not named `onEdit`: a simple trigger cannot access the calendar.
- **Settings-driven timing, applied forward only.** Intake schedules first
  contact at signup + the Settings delay; Deadline-derived Pre-touch =
  Deadline − the Settings offset and Post-touch = Deadline + the Settings
  offset, both at the Settings hour for 30 minutes. Values are read at use
  time, so a change affects later signups and deadlines but never rewrites
  existing rows or events; blank or non-numeric values are reported at Set Up
  and refuse to schedule. Derived cells can still be overwritten per-row.

## Testing Decisions

- **What makes a good test:** external behavior only — given inputs, assert
  observable outputs (a derived date, a Directory row, a scheduled event, a
  column index, a flag round-trip). Never assert implementation internals.
- **Module under test:** `pure` (header lookup, date parsing and derivation,
  event titles/descriptions, idempotency flags, Settings parsing) unit-tested
  in Vitest, no Google globals.
- **Glue tests** use an in-memory Google Sheets/Calendar harness
  (`tests/mocks/gas.ts`): the intake path (a submission writes one Directory
  row and never touches Form Responses, schedules one first contact), the
  first-contact edit lifecycle (move/clear), the onEdit deadline lifecycle
  (derive → create, edit → move, clear → remove, missing-ask → revert),
  Set Up idempotency, and Settings (defaults, forward-only application, loud
  validation). A live dry-run on the provisioned Sample still covers the
  Google-side behavior the harness can't.

## Out of Scope

- Volunteer-facing messaging of any kind (emails, texts); the tool only
  reminds the coordinator.
- Multi-coordinator/multi-calendar routing beyond each copy's own default
  calendar.
- Recurring "overdue" scan or digest.
- Reporting/dashboard beyond scanning the Tracker; Status dropdown; HTML
  sidebar.
- A dedicated web or mobile app; installing or hosting anything.

## Further Notes

- **Delivery:** trainer runs the volunteer recruitment training in two
  days; provisioning is on the critical path and requires clasp
  authentication as **jordan@dirtroadorganizing.org** (Sample owner).
- **Sample Form** (created at provisioning): Name, Email, and Phone
  (required contract), plus one optional open-ended question titled **About**
  (its description carries the friendly prompt) that lands in the Directory's
  About column. No consent field: the signup itself is the consent.
- **Demo seeding:** a few example Directory rows and one example Action so
  the live walkthrough is fast.
- **Plan B** if clasp auth fails: bind the script manually (Extensions →
  Apps Script → paste the bundle) — same end state.
- The Sample's links (spreadsheet + form) are placed in the README once
  provisioned.
