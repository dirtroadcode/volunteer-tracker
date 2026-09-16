# Campaign calendar

**Status:** accepted

Every campaign gets exactly one **Campaign Calendar** — a secondary Google
Calendar created by Set Up and recorded in the Settings tab — and every
Touchpoint is written to it rather than to the running user's default calendar.
This makes reminders one shared surface the Coordinator can hand to the campaign
team, and it keeps stored event ids valid no matter who runs New Action.

Using the default calendar broke the moment a second editor ran New Action: the
events landed on their calendar, their ids were stored in the shared Tracker,
and the Coordinator's later edit could not find them — so it created a duplicate
and left an orphan behind. One fixed calendar per campaign removes the whole
class of bug, because every reader and writer resolves the same calendar and the
stored ids always match.

Sharing stays a manual step in the Calendar UI. Programmatic sharing needs the
advanced Calendar service, the teammates' email addresses, and a Workspace
policy that permits it; a one-time click matches the existing "run Set Up once,
then share the Form" model and adds no dependency and no OAuth scope. We also
rejected an ICS feed: Google refreshes external subscriptions only every 8–24
hours, feeds are read-only, and the secret iCal address is not exposed to Apps
Script.

Consequence to keep in mind: Touchpoint event descriptions carry the volunteer's
name, email, phone, and About note, so anyone granted "See all event details"
can read the roster. Share with "See only free/busy" when the calendar is for
awareness rather than contact.
