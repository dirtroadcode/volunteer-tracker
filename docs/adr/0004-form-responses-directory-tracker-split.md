# Form Responses, Directory, and Tracker split

**Status:** accepted — supersedes ADR 0002

The Sample separates the Form's raw output from the coordinator's working
state. **Form Responses** is the Form's own response tab: one row per
submission, exactly as the Form wrote it, which the script reads and never
writes to. **Directory** is the volunteer roster, maintained by the script —
one row per signup, mirroring the Required Fields and the About note, and
carrying the intake fields (Reach out by, First Contact Event). **Menu of
Asks** stays the directly-editable registry, **Tracker** holds Actions (one
volunteer's dated commitment to one Ask), and **Settings** holds the timing
knobs.

ADR 0002 kept operational columns on the Form-linked response tab — the very
thing its own rationale called untenable. A coordinator's typed state sat in a
tab the Form manages, where a new question adds a column and where rows cannot
be deleted without desyncing the Form. Splitting the raw log from the roster
moves hand-edited state onto a script-owned sheet and leaves the Form alone.

Two consequences are deliberate. First contact is scheduled automatically when
a signup arrives (signup + the Settings delay) rather than waiting for the
coordinator to type a date, because not dropping a volunteer is the product's
whole job. And the Directory mirrors only the declared fields, not every
column the Form submits, so a coordinator's own questions cannot shift the
roster's shape.

Rejected alternatives worth remembering: a separate **Intake** tab consumed by
onboarding (it left no roster to find a returning volunteer in, which
reassignment needs); **mirroring every submitted column** into the Directory
(re-imports the form-driven drift this decision removes); and **migration code**
for existing copies (none exist, so the Sample is restructured by hand).
