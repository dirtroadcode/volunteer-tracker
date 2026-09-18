# Volunteerist

A volunteer follow-up tracker for campaigns, distributed as a **copyable
Sample**: coordinators copy one Spreadsheet, run **Volunteer Tools → Set Up**
once, and their volunteer tracker starts working — with first contact,
pre-ask appreciation, and post-ask check-in reminders landing on *their own*
Google Calendar. No installs, no hosting, no code — ever.

Built for [Dirtroad Organizing](https://www.dirtroadorganizing.org/) volunteer
recruitment training.

## How it works

```
Volunteer fills the intake Form
        │
        ▼
Form Responses tab  (the Form's own log; the script only reads it)
        │  intake writes a Directory row and schedules the first contact
        ▼
Directory tab  (roster)  ──►  Campaign calendar: "First contact: Jane Doe"
        │  coordinator runs "New Action", picks volunteer · Ask · Deadline in the dialog
        ▼
Tracker tab: Name · Ask (dropdown) · Deadline · Pre-touch · Post-touch · Notes
        │  the dialog schedules both reminders on submit
        ▼
Campaign calendar: "Pre-touch: Jane" (Deadline − 2 days)  and
          "Post-touch: Jane — canvassing (base)" (Deadline + 3 days)
```

Five tabs, each with one job:

| Tab | Job | Who edits it |
|---|---|---|
| **Form Responses** | The Form's own response log — one row per submission, exactly as written | The Form only; the script **reads and never writes** it |
| **Directory** | The volunteer roster — who signed up, plus the intake fields (`Reach out by`, `Calendar Event ID`) | The script on submit; the coordinator changes a reach-out date |
| **Menu of Asks** | The campaign's list of concrete asks a volunteer can do | The coordinator (type to add/rename/delete — no code); the New Action dialog adds a new ask when you type one |
| **Tracker** | One row per **Action** — a volunteer committed to an Ask, with its deadline, derived touches, and notes | The coordinator; the script fills pre/post dates and event markers |
| **Settings** | Four timing knobs (first-contact delay, pre/post offsets, reminder hour), plus the Campaign Calendar Set Up records | Only the coordinator |

The split between Form Responses and Directory is deliberate: the coordinator's
working state never sits on a tab the Form manages, so a new Form question
cannot shift the roster and rows can be deleted freely. (See
`docs/adr/0004-form-responses-directory-tracker-split.md`.)

The coordinator-only boundary is deliberate too: the tool reminds *the
coordinator* and never messages volunteers. (See `docs/adr/0003-coordinator-only-boundary.md`.)

A campaign's reminders live on **one Campaign Calendar** that Set Up creates and
records in Settings — not on whoever happens to run New Action. That keeps the
event ids stored in the Tracker resolvable for every editor, so a second person
can't split the reminders onto their personal calendar.

Inviting a **Collaborator** to help run the campaign takes two grants, in two
Google surfaces: **Share** the Spreadsheet with them as *Editor* (a bound-script
menu item only runs for users who can edit the copy), and share the Campaign
Calendar from **Google Calendar → Other calendars** with *Make changes to
events* (New Action runs as the person clicking it, so their own execution
writes the reminders). Someone who only watches needs the calendar alone, as
*See all event details* — no Spreadsheet access. The Spreadsheet grant is broad:
a Collaborator can read every tab (the roster included) and open
**Extensions → Apps Script**. (See `docs/adr/0005-campaign-calendar.md`.)

## For coordinators

Setup lives at **[dirtroadcode.github.io/volunteer-tracker](https://dirtroadcode.github.io/volunteer-tracker/)** —
the landing page at the repo root (`index.html`). It is the whole quick-start:
copy the Sample, run **Volunteer Tools → Set Up** once (which creates the
sign-up Form for you), and the reminders take care of themselves. There is no
separate guide document.

**Sample link** (created during provisioning — fill in once deployed):
- Sample Spreadsheet: https://drive.google.com/open?id=1kkDbFm7887KsAoYB507DA3ClLXC75XMhGuornfhgoHA

## For developers

Requires a nix flake dev shell (`nix develop`, or direnv).

```sh
npm install
npm run build   # TS → dist/Code.js + appsscript.json
npm test        # vitest over the automation logic (intake, dates, settings, idempotency)
npm run check   # biome (lint + format + imports), tsc --noEmit, then the tests
npm run push    # build + clasp push to the bound Apps Script project
npm run login   # clasp Google auth (interactive, one-time per machine/account)
npm run login:status  # is clasp authenticated, and as whom? (stale token → "invalid_grant")
```

The gate is split across two git hooks (`simple-git-hooks`, installed by
`npm install`): `pre-commit` runs lint + typecheck, so a lint error or type
error blocks the commit, and `pre-push` runs the tests, so a failing test
blocks the push. `npm run check` runs the whole gate by hand. Nothing applies
fixes — run `npm run lint:fix` for formatting and safe fixes, then commit. Two
escape hatches bypass the hooks: `--no-verify` (`git commit --no-verify` or
`git push --no-verify`) skips the hook for that command only, and
`SKIP_SIMPLE_GIT_HOOKS=1` (honored by the hook script itself, so it still works
when the gate is what's broken) skips them all. Use either only when you know
why the gate is wrong. `npm install` (via `prepare`) is what arms the hooks, so
an existing clone re-runs it once to pick up the `pre-push` hook.

Biome comes from the flake dev shell (`biome` in `flake.nix`), not npm: its
prebuilt npm binary is dynamically linked and cannot run on NixOS. `npm run
lint` and `lint:fix` therefore need the dev shell on `PATH`, while the hooks
resolve that shell themselves (`nix develop . -c npm run check:commit` on
commit, `nix develop . -c npm run test` on push) so commits and pushes are
gated the same way from a bare shell, an editor, or an already-open dev
shell. Biome ignores whatever `.gitignore` ignores (`dist/`, `node_modules/`,
`.worktrees/`) plus the two tracked-but-machine-local files `.clasp.json` and
`.pi/settings.json`, which their own tools rewrite in their own format.

Deployment model: the repo is the source of truth; `clasp push` ships
`dist/` (see `.clasp.json`) to the Apps Script project bound to the Sample
Spreadsheet. Coordinators never redeploy — they copy the live Sample.

Separately, the coordinator landing page is served by GitHub Pages from branch
`main`, folder `/`, at https://dirtroadcode.github.io/volunteer-tracker/. The
page is static, so no Actions workflow or build step is involved. To move it to
a custom domain, add a `CNAME` file at the repo root and point a DNS record at
`dirtroadcode.github.io`.

### Layout

```
src/config.ts     CONFIG — every tunable (tab/column names, required fields, Settings defaults, asks)
src/pure.ts       pure, unit-tested logic (headers, dates, Settings parsing, event content)
src/dialog.ts     the New Action modal — a pure HTML builder
src/automation.ts Apps Script behaviors (Set Up, intake, edits, New Action), seam-injected
src/services.ts   wires the live Google globals into the seam (UI resolved lazily)
src/index.ts      Apps Script entry points (menu, triggers)
tests/mocks/      in-memory Google Sheets/Calendar harness
tests/landing-page.test.ts  contract tests for the landing page
index.html        the coordinator landing + setup page (GitHub Pages)
docs/adr/         decisions: copy distribution, Form/Directory split, coordinator-only boundary
CONTEXT.md        domain glossary (Coordinator, Campaign, Sample, Ask, Action, touches)
```

### Tuning

Tab and column names live in `src/config.ts` → `CONFIG`, along with the
required-fields contract (`Name`, `Email`, `Phone`), the declared Directory
fields (including the optional `About` note), the Settings defaults, and the
seeded Menu of Asks. Renaming a header on a live copy means updating `CONFIG`
*and* re-running Set Up to re-add it; renaming without updating hides the
column from the script (Set Up re-creates any missing ones, but never merges a
renamed one back).

Timing knobs are **not** edited in `CONFIG` on a live copy — they live in the
Settings tab, which Set Up seeds from the `CONFIG.settings` defaults.
