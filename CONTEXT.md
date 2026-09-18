# Volunteer Follow-Up Tracker

Coordinates volunteer recruitment follow-ups for a coordinator: raw Form
responses become a volunteer roster, and each volunteer's commitments become
Actions that put reminders on the Campaign Calendar. Distributed as a copyable
Sample, so each coordinator runs their own instance.

## Language

**Coordinator**:
The person who sets up a campaign copy, runs it, and owns its Campaign
Calendar. May invite Collaborators to help.
_Avoid_: owner, admin, user, operator

**Collaborator**:
Someone the Coordinator invites to help run a Campaign: they add Actions and
see the same reminders. They hold edit access to the copy and write access to
its Campaign Calendar.
_Avoid_: teammate, helper, member, viewer

**Campaign**:
One coordinator's volunteer recruitment effort, operated from their own copy
of the Sample. Fully independent of other campaigns: its own volunteer rows,
its own calendar events.
_Avoid_: program, roster

**Campaign Calendar**:
The one Google Calendar a Campaign's Touchpoints are written to, created by Set
Up and recorded so every run of the automation resolves the same calendar. The
Coordinator shares it with Collaborators from Google Calendar; its event
descriptions carry volunteer contact details, so the sharing scope is
deliberate.
_Avoid_: team calendar, shared calendar, feed

**Sample**:
The distributable original — intake Form, response Spreadsheet, and
automation — built once and copied by coordinators to start a Campaign.
_Avoid_: template, demo, boilerplate

**Intake Form**:
The Google Form volunteers fill out when they sign up. A Campaign uses the
Sample's Form or the coordinator's own; either must capture the Required
Fields for the tool to work.
_Avoid_: application form, signup sheet

**Required Fields**:
The volunteer information the tool must receive from the Intake Form, located
by header text. A form missing them is flagged at setup rather than failing
silently later. (Currently: Name, Email, and Phone.)

**About**:
The one optional free-text answer the Directory keeps about a volunteer, for
the coordinator's eyes at contact time; the tool never depends on it. The Form
asks for it and the Directory stores it under the same title.
_Avoid_: profile, bio, share note

**Form Responses**:
The Form's own response tab — one row per submission, exactly as the Form
wrote it. The script reads it to build the Directory and never writes to it,
so changing the Form cannot disturb the Campaign's working state.
_Avoid_: raw tab, raw responses

**Directory**:
The volunteer roster — the record of who signed up, where volunteer identity
comes from, and the tab the coordinator reaches out from. One row per signup,
written by the script from the submission (the Required Fields and the Share
note), and carrying the intake fields: Reach out by and its calendar event id.
_Avoid_: raw tab, responses sheet, intake sheet

**Intake**:
A new signup becoming a volunteer in the Directory — the script writes the
row and schedules the first contact. Automatic on submission; the coordinator
does not run it.
_Avoid_: signup processing

**Menu of Asks**:
The Campaign's registry of Asks — a directly-editable tab, the "menu of
volunteer asks" built during training. An Action's Ask cell is a dropdown fed
from it; no script access needed to maintain it.
_Avoid_: task list, registry

**Ask**:
A concrete thing a volunteer can do to help the Campaign, listed on the Menu
of Asks (e.g. "canvassing (base)"). Generic at the menu level; concrete when
it is the subject of an Action.
_Avoid_: task, job, assignment

**Action**:
One volunteer's dated commitment to one Ask — the unit of the Tracker. An
Action pairs a volunteer with an Ask from the Menu of Asks, a definitive
Deadline, the deadline-derived pre/post touches, and notes. The same
volunteer can have many Actions.
_Avoid_: connection, assignment, task, specific ask

**Tracker**:
The Campaign's list of Actions — one row per Action, carrying its Ask,
Deadline, pre/post touches, and notes. Where a coordinator gives a returning
volunteer something new to do.
_Avoid_: pipeline

**Settings**:
The four timing knobs a coordinator can retune — first-contact delay,
pre/post-touch offsets, and the reminder hour. Read at use time, so a change
applies to later signups and deadlines but never rewrites existing rows or
events.
_Avoid_: config, options

**Deadline**:
The definitive date of an Action — the day the volunteer does the thing.
Anchors the pre- and post-touch reminders.
_Avoid_: due date, event date

**Pre-touch**:
The appreciation/confirmation reminder scheduled before a Deadline ("see
you on the 15th; thanks for signing up"). A calendar reminder to the
Coordinator, never a message to the volunteer.
_Avoid_: n/a

**Post-touch**:
The reminder after a Deadline — "how did it go?" plus thanks. A calendar
reminder to the Coordinator, never a message to the volunteer.
_Avoid_: follow-up, check-in

**Appreciation**:
The purpose of the pre/post touches — recognizing a volunteer's help,
inviting them to community events, and building warm connection so they
engage again. What turns a one-time volunteer into a Returning volunteer.
_Avoid_: thanks, thank-you note

**Returning volunteer**:
A volunteer who comes back for another Action. They are found in the Directory
or a past Action and given another Action — never a second Form submission.
_Avoid_: repeat volunteer, re-signup

**Touchpoint**:
A coordinator–volunteer interaction that must not be forgotten: first
contact, pre-touch, or post-touch. All are calendar reminders to the
Coordinator — the tool never messages the volunteer. First contact's reminder
is scheduled as soon as a volunteer reaches the Directory.
