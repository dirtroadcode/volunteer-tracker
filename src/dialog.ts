/**
 * The New Action dialog: an HTML form shown in an Apps Script modal so a
 * coordinator builds an Action in one place — volunteer, Ask, Deadline —
 * instead of selecting a Directory row and hand-editing the Tracker.
 *
 * The markup is a pure function so Vitest can assert it without a browser;
 * `automation.ts` gathers the options and handles the submit the form posts
 * back through `google.script.run`.
 */

export interface ActionVolunteer {
  /** 1-based Directory row — the server re-reads identity from it on submit. */
  row: number;
  name: string;
  email: string;
}

export interface ActionOptions {
  volunteers: ActionVolunteer[];
  asks: string[];
}

/** Sentinel `<select>` value that reveals the free-text new-Ask field. */
export const NEW_ASK = '__new__';

/** Escape text for an HTML text node or a double-quoted attribute value. */
export function escapeHtml(value: unknown): string {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** What a coordinator sees for a volunteer: name, then email for disambiguation. */
function volunteerLabel(volunteer: ActionVolunteer): string {
  return [volunteer.name, volunteer.email].filter(Boolean).join(' — ');
}

export function actionDialogHtml(options: ActionOptions): string {
  const volunteerOptions = options.volunteers
    .map((v) => `<option value="${v.row}">${escapeHtml(volunteerLabel(v))}</option>`)
    .join('');
  const askOptions = options.asks
    .map((ask) => `<option value="${escapeHtml(ask)}">${escapeHtml(ask)}</option>`)
    .join('');

  return `<!DOCTYPE html>
<html>
  <head>
    <base target="_top" />
    <style>
      :root { color-scheme: light; }
      body {
        margin: 0;
        padding: 16px;
        font-family: Roboto, Arial, sans-serif;
        font-size: 14px;
        color: #202124;
      }
      .field { display: block; margin-bottom: 14px; }
      .field > span { display: block; margin-bottom: 4px; font-weight: 500; }
      select, input[type='text'], input[type='date'] {
        box-sizing: border-box;
        width: 100%;
        padding: 8px;
        font: inherit;
        border: 1px solid #dadce0;
        border-radius: 4px;
        background: #fff;
      }
      select:focus, input:focus { outline: 2px solid #001f33; outline-offset: -1px; }
      .actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 20px; }
      button {
        padding: 8px 16px;
        font: inherit;
        font-weight: 500;
        border: 1px solid transparent;
        border-radius: 4px;
        cursor: pointer;
      }
      button#submit { background: #001f33; color: #fff; }
      button#submit:disabled { opacity: 0.6; cursor: default; }
      button.secondary { background: #fff; color: #001f33; border-color: #dadce0; }
      .error { min-height: 1em; margin: 0 0 4px; color: #b3261e; }
    </style>
  </head>
  <body>
    <form id="action-form">
      <label class="field">
        <span>Volunteer</span>
        <select id="volunteer" required>
          <option value="" disabled selected>Choose a volunteer…</option>
          ${volunteerOptions}
        </select>
      </label>
      <label class="field">
        <span>Ask</span>
        <select id="ask" required>
          <option value="" disabled selected>Choose an ask…</option>
          ${askOptions}
          <option value="${NEW_ASK}">Add a new ask…</option>
        </select>
      </label>
      <label class="field" id="new-ask-field" hidden>
        <span>New ask</span>
        <input id="new-ask" type="text" placeholder="e.g. Canvassing (base)" />
      </label>
      <label class="field">
        <span>Deadline</span>
        <input id="deadline" type="date" required />
      </label>
      <p id="error" class="error" role="alert"></p>
      <div class="actions">
        <button type="button" class="secondary" onclick="google.script.host.close()">Cancel</button>
        <button type="submit" id="submit">Create action</button>
      </div>
    </form>
    <script>
      (function () {
        var form = document.getElementById('action-form');
        var askSelect = document.getElementById('ask');
        var newAskField = document.getElementById('new-ask-field');
        var newAskInput = document.getElementById('new-ask');
        var error = document.getElementById('error');
        var submit = document.getElementById('submit');

        askSelect.addEventListener('change', function () {
          var adding = askSelect.value === '${NEW_ASK}';
          newAskField.hidden = !adding;
          newAskInput.required = adding;
          if (adding) newAskInput.focus();
        });

        form.addEventListener('submit', function (event) {
          event.preventDefault();
          error.textContent = '';
          var payload = {
            volunteerRow: Number(document.getElementById('volunteer').value),
            ask: askSelect.value === '${NEW_ASK}' ? newAskInput.value.trim() : askSelect.value,
            deadline: document.getElementById('deadline').value
          };
          submit.disabled = true;
          google.script.run
            .withSuccessHandler(function (result) {
              if (result && result.ok) {
                google.script.host.close();
                return;
              }
              error.textContent = (result && result.message) || 'Could not create the action.';
              submit.disabled = false;
            })
            .withFailureHandler(function (err) {
              error.textContent = (err && err.message) || 'Could not create the action.';
              submit.disabled = false;
            })
            .submitActionFromDialog(payload);
        });
      })();
    </script>
  </body>
</html>`;
}
