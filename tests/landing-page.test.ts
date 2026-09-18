import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseHTML } from 'linkedom';
import { describe, expect, it } from 'vitest';

const pagePath = join(import.meta.dirname, '..', 'index.html');

function loadPage() {
  return parseHTML(readFileSync(pagePath, 'utf-8')).document;
}

describe('Volunteerist landing page', () => {
  it('identifies itself as the Volunteerist quick-start', () => {
    const document = loadPage();

    expect(document.documentElement.getAttribute('lang')).toBe('en');
    expect(document.querySelector('meta[name="viewport"]')).not.toBeNull();
    expect(document.querySelectorAll('h1')).toHaveLength(1);
    expect(document.querySelector('h1')?.textContent?.trim().length).toBeGreaterThan(0);
  });

  it('links to the Sample tracker with a one-click copy', () => {
    const document = loadPage();
    const copyLink = [...document.querySelectorAll('a')].find((a) =>
      a.getAttribute('href')?.includes('1kkDbFm7887KsAoYB507DA3ClLXC75XMhGuornfhgoHA/copy'),
    );

    expect(copyLink).toBeDefined();
    expect(copyLink?.textContent?.trim().length).toBeGreaterThan(0);
  });

  it('links the open source claim to the source repository', () => {
    const document = loadPage();
    const sourceLink = [...document.querySelectorAll('a')].find(
      (a) => a.getAttribute('href') === 'https://github.com/dirtroadcode/volunteer-tracker',
    );

    expect(sourceLink).toBeDefined();
    expect(sourceLink?.textContent?.trim()).toBe('open source');
  });

  it('tells the coordinator how to make a copy', () => {
    const text = loadPage().body.textContent ?? '';

    expect(text).toContain('Make a copy');
  });

  it('explains the one-time Set Up and where the form comes from', () => {
    const text = loadPage().body.textContent ?? '';

    expect(text).toContain('Volunteer Tools');
    expect(text).toContain('Set Up');
    expect(text).toMatch(/create a sign-up form/i);
  });

  it('reassures the coordinator about the one-time Google permission screen', () => {
    const text = loadPage().body.textContent ?? '';

    expect(text).toMatch(/permission/i);
    expect(text).toMatch(/once/i);
    expect(text).toMatch(/hasn[’']t verified/i);
    expect(text).toMatch(/your own/i);
    expect(text).toMatch(/Advanced/i);
    expect(text).toMatch(/unsafe/i);
  });

  it('explains the everyday loop: signups land in the Directory', () => {
    const text = loadPage().body.textContent ?? '';

    expect(text).toContain('Directory');
    expect(text).toContain('New Action');
  });

  it('covers what the coordinator maintains after setup', () => {
    const text = loadPage().body.textContent ?? '';

    expect(text).toContain('Menu of Asks');
    expect(text).toContain('Settings');
  });

  it('tells the coordinator to share the campaign calendar with the team', () => {
    const text = loadPage().body.textContent ?? '';

    expect(text).toMatch(/campaign calendar/i);
    expect(text).toMatch(/share/i);
    expect(text).toContain('Other calendars');
    expect(text).toContain('Editor');
  });

  it('routes form setup through Set Up instead of a dead Sample Form link', () => {
    const document = loadPage();
    const hrefs = [...document.querySelectorAll('a')].map((a) => a.getAttribute('href') ?? '');

    expect(hrefs.some((href) => href.includes('forms/d'))).toBe(false);
    expect(document.body.textContent ?? '').not.toMatch(/SETUP\.md/);
  });

  it('keeps every link usable', () => {
    const document = loadPage();

    for (const anchor of document.querySelectorAll('a')) {
      expect((anchor.textContent ?? '').trim().length).toBeGreaterThan(0);
      if (anchor.getAttribute('target') === '_blank') {
        expect(anchor.getAttribute('rel')).toContain('noopener');
      }
    }
  });

  it('is self-contained and carries the brand', () => {
    const document = loadPage();
    const stylesheets = [...document.querySelectorAll('link[rel="stylesheet"]')];
    const inline = document.querySelector('style')?.textContent ?? '';

    expect(document.querySelectorAll('script')).toHaveLength(0);
    expect(stylesheets).toHaveLength(1);
    expect(stylesheets[0]?.getAttribute('href')).toContain('fonts.googleapis.com');
    expect(inline).toContain('--navy: #001f33');
    expect(inline).toContain('--gold: #fedd00');
    expect(inline).toContain('--ink: #001726');
    expect(inline).toContain('Lora');
  });
});
