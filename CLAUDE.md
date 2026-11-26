# Done With To Do: notes for Claude

A free, open-source web page that backs up Microsoft To Do lists (zip: CSV, printable HTML,
Obsidian Tasks Markdown, .ics, JSON, attachments) and moves them to Google Tasks or, through
an Apple Shortcut, to Apple Reminders. Live at https://lifeboatlabs.github.io/done-with-todo/
(GitHub Pages, `main` branch). Part of Lifeboat Labs (https://lifeboatlabs.github.io).

## Files
- `index.html`, `styles.css`: the page. Its CSP only allows connections to Microsoft and Google.
- `app.js`: UI and flow. `core.js`: data model and every output format (shared by tests).
- `microsoft.js`: PKCE sign-in and Microsoft Graph. Sends GET requests only (a test checks).
- `google.js`: Google Identity Services token client and the Google Tasks transfer.
- `zip.js`: dependency-free zip writer. `config.js`: client IDs and switches (see below).
- `privacy.html`/`privacy.js`: must match the permissions the code asks for.

## Rules
- No build step, no dependencies, plain browser JavaScript.
- When JS or CSS changes, bump `?v=N` in `index.html` and `privacy.html` and `core.VERSION`
  together. GitHub Pages lets browsers cache files for ten minutes.
- Write for non-technical people: plain words, say what to do next, no jargon in the UI.
- Never send anything to a server of ours; there isn't one.

## Switches in config.js
- `googleForEveryone`: false until Google approves the app; testers open the site with `?testers`.
- `remindersShortcutUrl`: the iCloud link of the "Done With To Do" Apple Shortcut; empty means
  only testers see the Reminders option. The shortcut reads JSON from the clipboard
  (`core.toRemindersData`) and is started with `shortcuts://run-shortcut?name=Done%20With%20To%20Do`.
- Microsoft only shares attached files on personal accounts with `Tasks.ReadWrite`; the page asks
  for it only when the person clicks "Include attached files" (or opens the site with `?files`).

## Tests
`node tests/unit.test.js` (no dependencies) and `python3 tests/e2e.py` (Playwright, Chrome).
Run both before pushing.

## Commits
Author commits as `LifeboatLabs <361480394+LifeboatLabs@users.noreply.github.com>`.
