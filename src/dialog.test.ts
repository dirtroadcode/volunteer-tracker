import { describe, expect, it } from 'vitest';
import { actionDialogHtml, escapeHtml, NEW_ASK } from './dialog';

describe('actionDialogHtml', () => {
  it('builds a volunteer select, an ask select, and a date input', () => {
    const html = actionDialogHtml({
      volunteers: [{ row: 2, name: 'Jane', email: 'jane@example.org' }],
      asks: ['Phone banking'],
    });

    expect(html).toContain('<select id="volunteer"');
    expect(html).toContain('<select id="ask"');
    expect(html).toContain('<input id="deadline" type="date"');
    expect(html).toContain('<option value="2">Jane — jane@example.org</option>');
    expect(html).toContain('<option value="Phone banking">Phone banking</option>');
    expect(html).toContain(`<option value="${NEW_ASK}">Add a new ask…</option>`);
    expect(html).toContain('submitActionFromDialog');
  });

  it('labels a volunteer with whichever of name/email is present', () => {
    const html = actionDialogHtml({
      volunteers: [
        { row: 2, name: 'Jane', email: '' },
        { row: 3, name: '', email: 'sam@example.org' },
      ],
      asks: [],
    });

    expect(html).toContain('<option value="2">Jane</option>');
    expect(html).toContain('<option value="3">sam@example.org</option>');
  });

  it('escapes volunteer and ask text that contains markup', () => {
    const html = actionDialogHtml({
      volunteers: [{ row: 2, name: '<b>Jane</b>', email: 'jane@example.org' }],
      asks: ['<script>'],
    });

    expect(html).not.toContain('<b>Jane</b>');
    expect(html).toContain('&lt;b&gt;Jane&lt;/b&gt;');
    expect(html).not.toContain('<option value="<script>">');
    expect(html).toContain('<option value="&lt;script&gt;">&lt;script&gt;</option>');
  });
});

describe('escapeHtml', () => {
  it('escapes the five HTML-significant characters', () => {
    expect(escapeHtml(`<>&"'`)).toBe('&lt;&gt;&amp;&quot;&#39;');
  });
});
