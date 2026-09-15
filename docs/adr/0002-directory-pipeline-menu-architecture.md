# Directory, Menu of Asks, and Pipeline architecture

**Status:** superseded by ADR 0004

The Sample has three tabs: an append-mostly **Directory** of raw Form
responses (deduped by email, with a "Reach out by" first-contact column), a
directly-editable **Menu of Asks** registry that feeds the Pipeline's Ask
dropdown, and a **Pipeline** where each row is one Connection — a
volunteer↔Specific Ask pair with a Deadline, deadline-derived Pre/Post-touch
dates, and an Outcome. This deliberately departs from the PRD's single-tab
"append operational columns to the response sheet" model: coordinator-driven
Connection creation plus bring-your-own Forms made grafting our columns onto
unpredictable raw response sheets untenable, and per-Connection follow-up
(not per-volunteer status) matches how repeat volunteers actually work.
