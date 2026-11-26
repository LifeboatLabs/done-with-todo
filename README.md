# Done With To Do

**Microsoft To Do shuts down on June 30, 2027. Take your lists with you.**

Done With To Do is a free web page that copies your Microsoft To Do lists into Google Tasks, or saves them as a full backup you can open anywhere. It runs entirely in your browser: there's no server, no account to create and no tracking.

**Use it here:** https://lifeboatlabs.github.io/done-with-todo/

## What it does

1. **Sign in with Microsoft.** It asks for read-only access, so it can't change or delete anything in To Do.
2. **Choose your lists.** You see every list and task it found.
3. **Save or move them.**
    - **Full backup (.zip):** a spreadsheet (CSV), a printable page, Markdown files for Obsidian (Tasks plugin format), calendar files (.ics) for Thunderbird, Nextcloud, Tasks.org and others, a JSON file with everything, and the files attached to your tasks.
    - **Google Tasks:** each list becomes a Google Tasks list. Steps become subtasks, and notes, due dates and completed tasks come too.
    - **Apple Reminders:** coming soon.

## What moves where

| To Do has | Backup | Google Tasks |
| --- | --- | --- |
| Lists and tasks | Yes | Yes |
| Steps | Yes | Yes, as subtasks |
| Notes | Yes | Yes |
| Due dates | Yes | Yes |
| Completed tasks | Yes | Optional |
| Repeats | Yes | Written into the task's notes, because Google Tasks can't import repeats |
| Reminders | Yes | Written into the task's notes |
| Importance, categories, links | Yes | Written into the task's notes |
| Attached files | Yes, in the zip | Names written into the notes |
| My Day, list groups | No, because Microsoft doesn't share these with other apps | No |

## Privacy

- Your tasks go straight from Microsoft to your browser and, if you choose, straight to Google.
- Sign-in tokens stay in the page's memory and disappear when you close the tab.
- The page's Content Security Policy only lets it connect to Microsoft and Google.
- No cookies, analytics or ads. Read the full [privacy page](privacy.html).

## Support the project

If this saved you an afternoon of retyping, **star this repository** so other people can find it.

## Run your own copy

The site is plain HTML, CSS and JavaScript, with no build step and no dependencies.

1. Fork this repository and turn on GitHub Pages (Settings, Pages, deploy from the `main` branch).
2. In Microsoft Entra, register an app with "Accounts in any organizational directory and personal Microsoft accounts". Add your site address as a **Single-page application** redirect URI and add the delegated permission `Tasks.Read`.
3. In Google Cloud, enable the Google Tasks API, configure the Google Auth Platform, and create a **Web application** OAuth client with your site's origin as an authorized JavaScript origin.
4. Put both client IDs in `config.js`. Until the Microsoft ID is set, the page shows a yellow setup box with the exact values to register.

## Tests

```
node tests/unit.test.js
```

The tests cover the backup formats, the zip writer, Microsoft sign-in (PKCE) and paging, and the Google Tasks transfer (ordering, subtasks, rate limits and resuming after a pause), using simulated Microsoft and Google servers.

## Licence

MIT. The Fira Sans font is under the SIL Open Font License (see `FONT-LICENSE.txt`).

Done With To Do isn't made by, or affiliated with, Microsoft or Google. Microsoft To Do is a trademark of Microsoft.
