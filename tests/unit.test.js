// Unit + mocked end-to-end tests for Done With To Do. Run: node tests/unit.test.js
'use strict';
const assert = require('assert');
const path = require('path');
const fs = require('fs');
const { execFileSync } = require('child_process');

const SITE = process.env.DWTD_SITE || path.join(__dirname, '..');
const core = require(path.join(SITE, 'core.js'));
const zip = require(path.join(SITE, 'zip.js'));
const MicrosoftClient = require(path.join(SITE, 'microsoft.js'));
const GoogleClient = require(path.join(SITE, 'google.js'));

let passed = 0, failed = 0;
const queue = [];
function test(name, fn) { queue.push([name, fn]); }

// ------------------------------------------------------------------ dates
test('due dates: midnight UTC keeps its day in any time zone', () => {
  assert.strictEqual(core.graphDate({ dateTime: '2026-11-30T00:00:00.0000000', timeZone: 'UTC' }), '2026-11-30');
});
test('due dates: local midnight stored as UTC converts back to local day', () => {
  // Only meaningful when TZ is set; we compute the expectation from the same rule.
  const d = new Date(Date.UTC(2026, 10, 29, 23, 0, 0));
  const expected = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  assert.strictEqual(core.graphDate({ dateTime: '2026-11-29T23:00:00.0000000', timeZone: 'UTC' }), expected);
});
test('due dates: named zone keeps wall-clock day', () => {
  assert.strictEqual(core.graphDate({ dateTime: '2026-03-01T00:00:00', timeZone: 'W. Europe Standard Time' }), '2026-03-01');
});
test('dates: empty / year 0001 treated as none', () => {
  assert.strictEqual(core.graphDate(null), null);
  assert.strictEqual(core.graphDate({ dateTime: '0001-01-01T00:00:00', timeZone: 'UTC' }), null);
  assert.strictEqual(core.graphISO('2020-07-24T19:15:25.8836143Z'), '2020-07-24T19:15:25.000Z');
});

// ------------------------------------------------------------------ repeats
test('recurrence text and RRULE', () => {
  const cases = [
    [{ type: 'daily', interval: 1 }, 'every day', 'FREQ=DAILY'],
    [{ type: 'daily', interval: 3 }, 'every 3 days', 'FREQ=DAILY;INTERVAL=3'],
    [{ type: 'weekly', interval: 1, daysOfWeek: ['thursday', 'monday'] }, 'every week on Monday, Thursday', 'FREQ=WEEKLY;BYDAY=MO,TH'],
    [{ type: 'weekly', interval: 1, daysOfWeek: ['friday', 'monday', 'tuesday', 'wednesday', 'thursday'] }, 'every weekday', 'FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR'],
    [{ type: 'weekly', interval: 2, daysOfWeek: ['saturday'] }, 'every 2 weeks on Saturday', 'FREQ=WEEKLY;INTERVAL=2;BYDAY=SA'],
    [{ type: 'absoluteMonthly', interval: 1, dayOfMonth: 15 }, 'every month on the 15th', 'FREQ=MONTHLY;BYMONTHDAY=15'],
    [{ type: 'absoluteMonthly', interval: 1, dayOfMonth: 22 }, 'every month on the 22nd', 'FREQ=MONTHLY;BYMONTHDAY=22'],
    [{ type: 'relativeMonthly', interval: 1, daysOfWeek: ['monday'], index: 'first' }, 'every month on the first Monday', 'FREQ=MONTHLY;BYDAY=1MO'],
    [{ type: 'relativeMonthly', interval: 1, daysOfWeek: ['friday'], index: 'last' }, 'every month on the last Friday', 'FREQ=MONTHLY;BYDAY=-1FR'],
    [{ type: 'absoluteYearly', interval: 1, month: 3, dayOfMonth: 3 }, 'every year on March 3rd', 'FREQ=YEARLY;BYMONTH=3;BYMONTHDAY=3'],
    [{ type: 'relativeYearly', interval: 1, month: 11, daysOfWeek: ['thursday'], index: 'fourth' }, 'every year on the fourth Thursday of November', 'FREQ=YEARLY;BYMONTH=11;BYDAY=4TH']
  ];
  for (const [pattern, text, rrule] of cases) {
    const rec = { pattern, range: { type: 'noEnd', startDate: '2026-01-01' } };
    assert.strictEqual(core.describeRecurrence(rec), text);
    assert.strictEqual(core.recurrenceToRRule(rec), rrule);
  }
  const until = { pattern: { type: 'daily', interval: 1 }, range: { type: 'endDate', endDate: '2027-06-30' } };
  assert.strictEqual(core.describeRecurrence(until), 'every day until 2027-06-30');
  assert.strictEqual(core.recurrenceToRRule(until), 'FREQ=DAILY;UNTIL=20270630');
  const count = { pattern: { type: 'weekly', interval: 1, daysOfWeek: ['monday'] }, range: { type: 'numbered', numberOfOccurrences: 10 } };
  assert.strictEqual(core.recurrenceToRRule(count), 'FREQ=WEEKLY;BYDAY=MO;COUNT=10');
  assert.strictEqual(core.describeRecurrence(null), null);
});

// ------------------------------------------------------------------ html notes
test('htmlToText', () => {
  assert.strictEqual(core.htmlToText('<p>Hello&nbsp;<b>world</b></p><p>Line&amp;2<br>3</p>'), 'Hello world\nLine&2\n3');
  assert.strictEqual(core.htmlToText('<ul><li>a</li><li>b</li></ul>'), '• a\n• b');
  assert.strictEqual(core.htmlToText('x &#128512; &#x1F600; &bogus; &#99999999;'), 'x 😀 😀 &bogus; &#99999999;');
  assert.strictEqual(core.htmlToText('<style>.a{}</style>ok<script>alert(1)</script>'), 'ok');
});

// ------------------------------------------------------------------ model from sample
const sample = core.sampleData(new Date(2026, 10, 25, 12, 0, 0));
const model = core.normalize(sample);
test('normalize sample', () => {
  assert.strictEqual(model.lists.length, 5);
  const s = core.summarize(model);
  assert.strictEqual(s.tasks, 20);
  assert.ok(s.done >= 2 && s.steps >= 10 && s.repeating >= 5 && s.attachments === 2 && s.reminders === 2);
  const passport = model.lists[0].tasks[0];
  assert.strictEqual(passport.title, 'Renew passport');
  assert.strictEqual(passport.important, true);
  assert.strictEqual(passport.steps.length, 3);
  assert.strictEqual(passport.steps[0].done, true);
  assert.ok(/^\d{4}-\d{2}-\d{2}$/.test(passport.due));
  assert.ok(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(passport.reminder));
});

// ------------------------------------------------------------------ CSV
function parseCSV(text) {
  const rows = []; let row = [], cell = '', i = 0, q = false;
  if (text.charCodeAt(0) === 0xfeff) i = 1;
  for (; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += c; continue; }
    if (c === '"') q = true; else if (c === ',') { row.push(cell); cell = ''; }
    else if (c === '\r') { /* skip */ } else if (c === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
    else cell += c;
  }
  if (cell || row.length) { row.push(cell); rows.push(row); }
  return rows;
}
test('CSV round-trips and is formula-safe', () => {
  const csv = core.toCSV(model);
  assert.strictEqual(csv.charCodeAt(0), 0xfeff);
  const rows = parseCSV(csv);
  assert.strictEqual(rows.length, 1 + 20);
  rows.forEach(r => assert.strictEqual(r.length, core.CSV_HEADER.length));
  const passport = rows.find(r => r[1] === 'Renew passport');
  assert.ok(passport[8].includes('"fast track"'));
  assert.ok(passport[9].startsWith('[x] Get photos taken\n[ ] Fill in the form'));
  const formula = rows.find(r => r[1].includes('2+2'));
  assert.strictEqual(formula[1], "'=2+2 is not a formula");
  assert.strictEqual(core.csvCell('-milk'), '-milk');
  assert.strictEqual(core.csvCell('- milk'), '- milk');
  assert.strictEqual(core.csvCell("-2+3+cmd|' /C calc'!A0"), "'-2+3+cmd|' /C calc'!A0");
  assert.strictEqual(core.csvCell('@SUM(A1)'), "'@SUM(A1)");
});

// ------------------------------------------------------------------ printable HTML
test('printable HTML escapes everything', () => {
  const evil = core.normalize({
    rawLists: [{ id: 'l1', displayName: '<img src=x onerror=alert(1)>' }],
    rawTasksByList: { l1: [{ id: 't1', title: '<script>alert(1)</script>', status: 'notStarted', body: { content: '"><svg onload=alert(1)>', contentType: 'text' },
      linkedResources: [{ webUrl: 'javascript:alert(1)' }, { webUrl: 'https://example.com/?a="b"' }] }] }
  });
  const html = core.toPrintableHTML(evil);
  assert.ok(!html.includes('<script>alert'));
  assert.ok(!html.includes('<img src=x'));
  assert.ok(!html.includes('<svg onload'));
  assert.ok(!html.includes('href="javascript:'));
  assert.ok(html.includes('href="https://example.com/?a=&quot;b&quot;"'));
  const real = core.toPrintableHTML(model);
  assert.ok(real.includes('Renew passport') && real.includes('Completed'));
});

// ------------------------------------------------------------------ Markdown
test('Markdown files', () => {
  const files = core.toMarkdownFiles(model);
  assert.strictEqual(files.length, 5);
  const tasks = files[0].content;
  assert.ok(tasks.startsWith('# Tasks\n'));
  assert.ok(/- \[ \] Renew passport ⏫ 📅 \d{4}-\d{2}-\d{2}/.test(tasks), tasks);
  assert.ok(tasks.includes('    - [x] Get photos taken'));
  assert.ok(tasks.includes('    > Need: old passport'));
  assert.ok(tasks.includes('🔁 every month on the 15th'));
  assert.ok(/- \[x\] Return library books ✅ \d{4}-\d{2}-\d{2}/.test(tasks));
  const names = files.map(f => f.name);
  assert.deepStrictEqual(names, ['Tasks.md', 'Groceries 🛒.md', 'Home & garden.md', "Mum's 60th.md", 'Flagged email.md']);
});

// ------------------------------------------------------------------ iCalendar
test('ICS files are valid-looking and folded', () => {
  const files = core.toICSFiles(model);
  assert.strictEqual(files.length, 5);
  for (const f of files) {
    const lines = f.content.split('\r\n');
    assert.strictEqual(lines[0], 'BEGIN:VCALENDAR');
    assert.strictEqual(lines[lines.length - 2], 'END:VCALENDAR');
    assert.ok(!/[^\r]\n/.test(f.content), 'bare LF found');
    for (const line of lines) assert.ok(Buffer.byteLength(line, 'utf8') <= 75, 'line too long: ' + line);
    const begins = (f.content.match(/^BEGIN:VTODO/gm) || []).length, ends = (f.content.match(/^END:VTODO/gm) || []).length;
    assert.strictEqual(begins, ends);
  }
  const home = files[2].content.replace(/\r\n /g, '');
  assert.ok(home.includes('RRULE:FREQ=WEEKLY;INTERVAL=2;BYDAY=SA'));
  assert.ok(home.includes('RELATED-TO;RELTYPE=PARENT:dwtd-'));
  assert.ok(home.includes('STATUS:IN-PROCESS'));
  const tasks = files[0].content.replace(/\r\n /g, '');
  assert.ok(tasks.includes('DESCRIPTION:Need: old passport\\, 2 photos\\, card for the fee.\\nCheck the "fast track" option.'));
  assert.ok(tasks.includes('BEGIN:VALARM') && tasks.includes('TRIGGER;VALUE=DATE-TIME:'));
  assert.ok(tasks.includes('PRIORITY:1'));
  // folding never splits a multi-byte character
  const long = core.icsFold('SUMMARY:' + 'é'.repeat(100) + '😀'.repeat(30));
  long.split('\r\n').forEach(l => { assert.ok(Buffer.byteLength(l, 'utf8') <= 75); assert.ok(!l.includes('�')); });
  assert.strictEqual(long.replace(/\r\n /g, ''), 'SUMMARY:' + 'é'.repeat(100) + '😀'.repeat(30));
});

// ------------------------------------------------------------------ Google bodies
test('Google task bodies', () => {
  const passport = model.lists[0].tasks[0];
  const b = core.googleTaskBody(passport);
  assert.strictEqual(b.title, 'Renew passport');
  assert.strictEqual(b.status, 'needsAction');
  assert.ok(/^\d{4}-\d{2}-\d{2}T00:00:00\.000Z$/.test(b.due));
  assert.ok(b.notes.includes('Need: old passport') && b.notes.includes('— From Microsoft To Do —') && b.notes.includes('Reminder was set for') && b.notes.includes('Marked as important'));
  const done = model.lists[0].tasks[4];
  assert.strictEqual(core.googleTaskBody(done).status, 'completed');
  const huge = Object.assign({}, passport, { notes: 'x'.repeat(20000), title: 'y'.repeat(5000) });
  const hb = core.googleTaskBody(huge);
  assert.ok(hb.notes.length <= 8192, 'notes ' + hb.notes.length);
  assert.ok(hb.notes.includes('shortened'));
  assert.ok(hb.notes.includes('Marked as important'), 'footer kept');
  assert.ok(hb.title.length <= 1024);
  const plain = core.googleTaskBody({ title: '', done: false, status: 'notstarted', notes: '', categories: [], links: [], attachments: [], steps: [] });
  assert.deepStrictEqual(plain, { title: '(untitled)', status: 'needsAction' });
  assert.strictEqual(core.uniqueListTitle('Groceries', ['groceries', 'groceries (from to do)']), 'Groceries (from To Do 2)');
  assert.strictEqual(core.uniqueListTitle('Home', ['my tasks']), 'Home');
});

test('safe file names', () => {
  const used = new Set();
  assert.strictEqual(core.safeFileName('a/b:c*?', used, '.md'), 'a-b-c-.md');
  // file systems on Windows and Mac ignore case, so these clash too
  assert.strictEqual(core.safeFileName('A-B-C-', used, '.md'), 'A-B-C- (2).md');
  assert.strictEqual(core.safeFileName('a/b:c*?', used, '.md'), 'a-b-c- (3).md');
  assert.strictEqual(core.safeFileName('CON', used), '_CON');
  assert.strictEqual(core.safeFileName('...', used), 'Untitled');
  assert.ok(Array.from(core.safeFileName('😀'.repeat(200), used)).length <= 80);
});

// ------------------------------------------------------------------ zip
test('zip opens with Python zipfile and unzip', () => {
  const files = core.backupFiles(model, sample);
  core.attachmentPlan(model).forEach(p => files.push({ path: p.path, data: core.sampleAttachmentBytes(p.attachment) }));
  const parts = zip.makeZip(files, new Date(2026, 10, 25, 12, 0, 0));
  const buf = Buffer.concat(parts.map(p => Buffer.from(p)));
  const out = path.join(require('os').tmpdir(), 'dwtd-test-backup.zip');
  fs.writeFileSync(out, buf);
  const py = execFileSync('python3', ['-I', '-c', `
import zipfile, sys, json
z = zipfile.ZipFile(sys.argv[1])
assert z.testzip() is None
names = z.namelist()
data = json.loads(z.read([n for n in names if n.endswith('.json')][0]).decode('utf-8'))
print(len(names), data['summary']['tasks'], data['format'])
print('\\n'.join(sorted(names)))
`, out]).toString();
  const [first] = py.split('\n');
  assert.strictEqual(first, `${files.length} 20 done-with-todo-backup`);
  assert.ok(py.includes('Attachments/Home & garden/Fix the dripping tap/tap-model.pdf'), py);
  execFileSync('unzip', ['-tq', out]);
  assert.throws(() => zip.makeZip([{ path: 'a.txt', data: '1' }, { path: 'A.txt', data: '2' }]), /Duplicate/);
});

// ------------------------------------------------------------------ PKCE
test('PKCE challenge = base64url(SHA-256(verifier)), checked against Node crypto', async () => {
  const nodeCrypto = require('crypto');
  for (let i = 0; i < 20; i++) {
    const verifier = nodeCrypto.randomBytes(48).toString('base64url');
    const expected = nodeCrypto.createHash('sha256').update(verifier).digest('base64url');
    assert.strictEqual(await MicrosoftClient._test.pkceChallenge(verifier), expected);
  }
});

// ------------------------------------------------------------------ fake fetch helpers
function response(status, body, headers) {
  return {
    ok: status >= 200 && status < 300, status,
    headers: { get: k => (headers || {})[k] || (headers || {})[k.toLowerCase()] || null },
    json: async () => { if (body === undefined) throw new Error('no body'); return JSON.parse(JSON.stringify(body)); },
    arrayBuffer: async () => (body instanceof Uint8Array ? body.buffer : new TextEncoder().encode(String(body)).buffer)
  };
}
function memoryStorage() { const m = new Map(); return { getItem: k => (m.has(k) ? m.get(k) : null), setItem: (k, v) => m.set(k, String(v)), removeItem: k => m.delete(k) }; }

// ------------------------------------------------------------------ Microsoft sign-in flow
test('Microsoft sign-in: URL, state check, token exchange, error mapping', async () => {
  const storage = memoryStorage();
  let tokenBody = null;
  const ms = new MicrosoftClient({
    clientId: '11111111-2222-3333-4444-555555555555', redirectUri: 'https://sam.github.io/done-with-todo/', storage,
    fetch: async (url, init) => {
      assert.strictEqual(url, 'https://login.microsoftonline.com/common/oauth2/v2.0/token');
      tokenBody = new URLSearchParams(init.body);
      return response(200, { access_token: 'AT', expires_in: 3600 });
    }
  });
  const url = new URL(await ms.signInUrl());
  assert.strictEqual(url.origin + url.pathname, 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize');
  assert.strictEqual(url.searchParams.get('response_mode'), 'fragment');
  assert.strictEqual(url.searchParams.get('code_challenge_method'), 'S256');
  assert.strictEqual(url.searchParams.get('scope'), 'openid profile User.Read Tasks.Read');
  assert.strictEqual(url.searchParams.get('redirect_uri'), 'https://sam.github.io/done-with-todo/');
  const state = url.searchParams.get('state');
  assert.strictEqual(await ms.completeSignIn(''), null);
  assert.strictEqual(await ms.completeSignIn('#code=abc&state=' + state), true);
  assert.strictEqual(ms.token, 'AT');
  assert.strictEqual(tokenBody.get('code'), 'abc');
  assert.strictEqual(tokenBody.get('grant_type'), 'authorization_code');
  assert.ok(tokenBody.get('code_verifier').length >= 43);
  assert.strictEqual(storage.getItem(MicrosoftClient.STORAGE_KEY), null, 'one-time secret removed');
  // wrong state
  await ms.signInUrl();
  await assert.rejects(ms.completeSignIn('#code=abc&state=wrong'), e => e.code === 'ms_state');
  // error coming back from Microsoft
  await assert.rejects(ms.completeSignIn('#error=invalid_request&error_description=' + encodeURIComponent('AADSTS50011: The redirect URI does not match')),
    e => e.code === 'ms_signin' && e.message.includes('https://sam.github.io/done-with-todo/'));
  // token endpoint unreachable (e.g. registered as Web -> CORS blocked)
  const blocked = new MicrosoftClient({ clientId: 'x', redirectUri: 'https://a/', storage, fetch: async () => { throw new TypeError('Failed to fetch'); } });
  const u2 = new URL(await blocked.signInUrl());
  await assert.rejects(blocked.completeSignIn('#code=c&state=' + u2.searchParams.get('state')), e => e.code === 'ms_token_blocked' && /Single-page application/.test(e.message));
  // token endpoint returns AADSTS9002326
  const web = new MicrosoftClient({ clientId: 'x', redirectUri: 'https://a/', storage, fetch: async () => response(400, { error: 'invalid_request', error_description: 'AADSTS9002326: Cross-origin token redemption is permitted only for the Single-Page Application client-type.' }) });
  const u3 = new URL(await web.signInUrl());
  await assert.rejects(web.completeSignIn('#code=c&state=' + u3.searchParams.get('state')), e => /listed under "Web"/.test(e.message));
});

// ------------------------------------------------------------------ Microsoft Graph reading
test('Graph: paging, $expand fallback, throttling, attachments', async () => {
  const calls = [];
  let throttled = false;
  const sleeps = [];
  const tasksA = Array.from({ length: 150 }, (_, i) => ({ id: 'a' + i, title: 'Task ' + i, status: 'notStarted', checklistItems: [{ id: 's' + i, displayName: 'Step', isChecked: false }] }));
  const graph = async (url) => {
    calls.push(url);
    const u = new URL(url);
    const p = decodeURIComponent(u.pathname);
    if (p === '/v1.0/me') return response(200, { displayName: 'Sam', mail: null, userPrincipalName: 'sam@outlook.com' });
    if (p === '/v1.0/me/todo/lists') {
      if (!throttled) { throttled = true; return response(429, { error: { code: 'TooManyRequests' } }, { 'Retry-After': '2' }); }
      return u.searchParams.get('$skiptoken')
        ? response(200, { value: [{ id: 'B/+=', displayName: 'Old list' }] })
        : response(200, { value: [{ id: 'A', displayName: 'Groceries' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/todo/lists?$skiptoken=X' });
    }
    if (p === '/v1.0/me/todo/lists/A/tasks') {
      const skip = +(u.searchParams.get('$skip') || 0);
      const page = tasksA.slice(skip, skip + 100);
      const body = { value: page };
      if (skip + 100 < tasksA.length) body['@odata.nextLink'] = 'https://graph.microsoft.com/v1.0/me/todo/lists/A/tasks?$expand=checklistItems,linkedResources&$top=100&$skip=' + (skip + 100);
      return response(200, body);
    }
    if (p === '/v1.0/me/todo/lists/B/+=/tasks') {
      // This list rejects $expand entirely, and ignores $top
      if (u.search.includes('expand') || u.search.includes('top')) return response(400, { error: { code: 'BadRequest', message: 'nope' } });
      return response(200, { value: [{ id: 'b1', title: 'Old', status: 'completed', hasAttachments: true }] });
    }
    if (p === '/v1.0/me/todo/lists/B/+=/tasks/b1/checklistItems') return response(200, { value: [{ id: 'c1', displayName: 'Fetched step', isChecked: true }] });
    if (p === '/v1.0/me/todo/lists/B/+=/tasks/b1/attachments') {
      if (u.searchParams.get('$select')) return response(400, { error: { code: 'BadRequest' } });
      return response(200, { value: [{ id: 'f1', name: 'scan.pdf', size: 10, contentType: 'application/pdf', contentBytes: 'AAAA' }] });
    }
    if (p === '/v1.0/me/todo/lists/B/+=/tasks/b1/attachments/f1/$value') return response(500, { error: {} });
    if (p === '/v1.0/me/todo/lists/B/+=/tasks/b1/attachments/f1') return response(200, { id: 'f1', contentBytes: Buffer.from('hello').toString('base64') });
    return response(404, { error: { code: 'NotFound', message: p } });
  };
  const ms = new MicrosoftClient({ clientId: 'x', redirectUri: 'https://a/', storage: memoryStorage(), fetch: graph, sleep: async ms => { sleeps.push(ms); } });
  ms.token = 'AT'; ms.tokenExpires = Date.now() + 3600e3;
  const progress = [];
  const raw = await ms.readEverything(p => progress.push(p));
  assert.deepStrictEqual(raw.account, { name: 'Sam', email: 'sam@outlook.com' });
  assert.strictEqual(raw.rawLists.length, 2);
  assert.strictEqual(raw.rawTasksByList.A.length, 150);
  assert.strictEqual(raw.rawTasksByList['B/+='][0].checklistItems[0].displayName, 'Fetched step');
  assert.strictEqual(raw.rawTasksByList['B/+='][0].attachments[0].name, 'scan.pdf');
  assert.ok(!('contentBytes' in raw.rawTasksByList['B/+='][0].attachments[0]), 'no attachment bytes kept in memory');
  assert.deepStrictEqual(sleeps.slice(0, 1), [2000], 'Retry-After respected');
  assert.ok(calls.some(c => c.includes('/me/todo/lists/B%2F%2B%3D/tasks')), 'list ids are URL-encoded');
  assert.ok(progress.some(p => p.phase === 'tasks' && p.listName === 'Old list'));
  const bytes = await ms.downloadAttachment('B/+=', 'b1', 'f1');
  assert.strictEqual(Buffer.from(bytes).toString(), 'hello');
  const m2 = core.normalize(raw);
  assert.strictEqual(m2.lists[1].tasks[0].steps[0].done, true);
  // 401 surfaces as ms_auth
  const expired = new MicrosoftClient({ clientId: 'x', redirectUri: 'https://a/', storage: memoryStorage(), fetch: async () => response(401, { error: { code: 'InvalidAuthenticationToken' } }), sleep: async () => {} });
  expired.token = 'old';
  await assert.rejects(expired.readEverything(), e => e.code === 'ms_auth');
  // refuses to send the token anywhere but Graph
  await assert.rejects(ms.get('https://evil.example/steal'), e => e.code === 'graph');
});

test('Graph: attachments Microsoft refuses don\'t block the backup; beta is tried; a real expiry still stops', async () => {
  const calls = [];
  const graph = async (url) => {
    calls.push(url);
    const p = decodeURIComponent(new URL(url).pathname);
    if (p === '/v1.0/me') return response(200, { displayName: 'Lifeboat Labs', mail: 'lifeboatlabs@outlook.com' });
    if (p === '/v1.0/me/todo/lists') return response(200, { value: [{ id: 'L1', displayName: 'Tasks' }] });
    if (p === '/v1.0/me/todo/lists/L1/tasks') return response(200, { value: [
      { id: 'T1', title: 'eat', status: 'notStarted', hasAttachments: true, checklistItems: [{ id: 'c1', displayName: 'breakfast' }] },
      { id: 'T2', title: 'photo', status: 'notStarted', hasAttachments: true, checklistItems: [] }
    ] });
    // T1: every way of listing its files is refused, with a fresh sign-in
    if (/^\/(v1\.0|beta)\/me\/todo\/lists\/L1\/tasks\/T1\/attachments$/.test(p)) return response(401, { error: { code: 'UnknownError', message: '' } });
    // T2: v1.0 refuses; beta lists and downloads it
    if (p === '/v1.0/me/todo/lists/L1/tasks/T2/attachments' || p.startsWith('/v1.0/me/todo/lists/L1/tasks/T2/attachments/')) return response(401, { error: { code: 'UnknownError' } });
    if (p === '/beta/me/todo/lists/L1/tasks/T2/attachments') return response(200, { value: [{ id: 'F2', name: 'shot.png', size: 5, contentType: 'image/png' }] });
    if (p === '/beta/me/todo/lists/L1/tasks/T2/attachments/F2/$value') return response(200, new Uint8Array([1, 2, 3]));
    return response(404, { error: { code: 'NotFound', message: p } });
  };
  const ms = new MicrosoftClient({ clientId: 'x', redirectUri: 'https://a/', storage: memoryStorage(), fetch: graph, sleep: async () => {} });
  ms.token = 'AT'; ms.tokenExpires = Date.now() + 3600e3;
  const raw = await ms.readEverything();
  const [t1, t2] = raw.rawTasksByList.L1;
  assert.strictEqual(t1.attachmentsUnavailable, true);
  assert.deepStrictEqual(t1.attachments, []);
  assert.strictEqual(t2.attachments[0].name, 'shot.png');
  assert.ok(!t2.attachmentsUnavailable);
  assert.ok(calls.some(c => c.startsWith('https://graph.microsoft.com/beta/')), 'beta tried');
  assert.strictEqual(raw.attachmentProblems.length, 1);
  assert.strictEqual(raw.attachmentProblems[0].task, 'eat');
  const said = raw.attachmentProblems[0].details;
  assert.ok(said.includes('GET /v1.0/me/todo/lists/L1/tasks/T1/attachments?$select=id,name,size,contentType → 401\nUnknownError'), said);
  assert.ok(said.includes('GET /beta/me/todo/lists/L1/tasks/T1/attachments → 401'), said);
  assert.ok(!/: $/m.test(said), 'no dangling colon when Microsoft sends no message');
  assert.deepStrictEqual(Array.from(await ms.downloadAttachment('L1', 'T2', 'F2')), [1, 2, 3]);

  // What the person sees
  const m = core.normalize(raw);
  const eat = m.lists[0].tasks[0];
  assert.strictEqual(eat.attachmentsUnavailable, true);
  assert.strictEqual(eat.steps[0].title, 'breakfast', 'the rest of the task is intact');
  assert.strictEqual(core.summarize(m).attachmentsUnavailable, 1);
  assert.ok(/wouldn't share[\s\S]*- Tasks \/ eat/.test(core.backupReadme(m, [])), 'README lists the task');
  assert.ok(core.toCSV(m).includes('Microsoft wouldn’t share'));
  assert.ok(core.toPrintableHTML(m).includes('Microsoft wouldn’t share'));
  assert.ok(core.googleTaskBody(eat).notes.includes('Save them from To Do'));

  // A 401 after the sign-in's expiry time is still an expired sign-in, and stops the read
  const late = new MicrosoftClient({ clientId: 'x', redirectUri: 'https://a/', storage: memoryStorage(), fetch: graph, sleep: async () => {} });
  late.token = 'AT'; late.tokenExpires = Date.now() + 30e3; // inside the one-minute safety margin
  await assert.rejects(late.readEverything(), e => e.code === 'ms_auth' && !e.fresh && /expired/.test(e.message));

  // A 401 on the lists themselves with a fresh sign-in says so honestly, and names the request
  const refuse = new MicrosoftClient({ clientId: 'x', redirectUri: 'https://a/', storage: memoryStorage(), sleep: async () => {},
    fetch: async (url) => new URL(url).pathname === '/v1.0/me' ? response(200, {}) : response(401, { error: { code: 'UnknownError' } }) });
  refuse.token = 'AT'; refuse.tokenExpires = Date.now() + 3600e3;
  await assert.rejects(refuse.readEverything(), e => e.code === 'ms_auth' && e.fresh && !/expired/.test(e.message) &&
    e.details.startsWith('GET /v1.0/me/todo/lists → 401') && MicrosoftClient.signInExpired(e) === false);
});

test('Opting in to attached files asks for Tasks.ReadWrite, but the site still only sends GETs to Graph', async () => {
  const storage = memoryStorage();
  let tokenBody = null;
  const methods = [];
  const ms = new MicrosoftClient({
    clientId: 'x', redirectUri: 'https://a/', storage, scopes: MicrosoftClient.FILE_SCOPES, sleep: async () => {},
    fetch: async (url, init) => {
      if (url.startsWith('https://login.microsoftonline.com/')) { tokenBody = new URLSearchParams(init.body); return response(200, { access_token: 'AT', expires_in: 3600 }); }
      methods.push((init && init.method) || 'GET');
      const p = decodeURIComponent(new URL(url).pathname);
      if (p === '/v1.0/me') return response(200, { displayName: 'Sam' });
      if (p === '/v1.0/me/todo/lists') return response(200, { value: [{ id: 'L', displayName: 'Tasks' }] });
      if (p === '/v1.0/me/todo/lists/L/tasks') return response(200, { value: [{ id: 'T', title: 'eat', hasAttachments: true, checklistItems: [] }] });
      if (p === '/v1.0/me/todo/lists/L/tasks/T/attachments') return response(200, { value: [{ id: 'F', name: 'shot.png', size: 3, contentType: 'image/png' }] });
      if (p === '/v1.0/me/todo/lists/L/tasks/T/attachments/F/$value') return response(200, new Uint8Array([7, 8, 9]));
      return response(404, {});
    }
  });
  const url = new URL(await ms.signInUrl());
  assert.strictEqual(url.searchParams.get('scope'), 'openid profile User.Read Tasks.ReadWrite');
  await ms.completeSignIn('#code=c&state=' + url.searchParams.get('state'));
  assert.strictEqual(tokenBody.get('scope'), 'openid profile User.Read Tasks.ReadWrite');
  const raw = await ms.readEverything();
  assert.strictEqual(raw.rawTasksByList.L[0].attachments[0].name, 'shot.png');
  assert.deepStrictEqual(Array.from(await ms.downloadAttachment('L', 'T', 'F')), [7, 8, 9]);
  assert.ok(methods.length >= 5 && methods.every(m => m === 'GET'), 'only reads: ' + methods.join(','));
  // and the default stays read-only
  const plain = new MicrosoftClient({ clientId: 'x', redirectUri: 'https://a/', storage: memoryStorage() });
  assert.strictEqual(new URL(await plain.signInUrl()).searchParams.get('scope'), 'openid profile User.Read Tasks.Read');
});

test('Apple Reminders data: lists, dates, open steps as subtasks, everything else in the notes', () => {
  const m = core.normalize(core.sampleData());
  const all = core.toRemindersData(m, {});
  assert.strictEqual(all.app, 'Done With To Do');
  assert.ok(all.lists.length >= 4);
  all.lists.forEach(l => l.reminders.forEach(r => {
    assert.strictEqual(typeof r.title, 'string');
    assert.ok(r.date === '' || /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(r.date), r.date);
    assert.ok(Array.isArray(r.subtasks));
  }));
  // Only open tasks unless asked; finished steps go into the notes
  const openCount = m.lists.reduce((n, l) => n + l.tasks.filter(t => !t.done).length, 0);
  assert.strictEqual(core.remindersCount(all).reminders, openCount);
  const withDone = core.toRemindersData(m, { includeCompleted: true });
  assert.strictEqual(core.remindersCount(withDone).reminders, m.lists.reduce((n, l) => n + l.tasks.length, 0));
  assert.ok(withDone.lists.some(l => l.reminders.some(r => /Completed in To Do/.test(r.notes))));
  // A task with a reminder uses the reminder time; one with only a due date gets 09:00 that day
  const passport = all.lists.flatMap(l => l.reminders).find(r => r.title === 'Renew passport');
  assert.strictEqual(passport.date, m.lists.flatMap(l => l.tasks).find(t => t.title === 'Renew passport').reminder);
  assert.ok(passport.notes.includes('Due date in To Do: '));
  assert.ok(passport.notes.includes('Steps already done: Get photos taken'));
  assert.deepStrictEqual(passport.subtasks, ['Fill in the form', 'Book appointment']);
  const dueOnly = m.lists.flatMap(l => l.tasks).find(t => t.due && !t.reminder && !t.done);
  const dueOnlyR = all.lists.flatMap(l => l.reminders).find(r => r.title === dueOnly.title);
  assert.strictEqual(dueOnlyR.date, dueOnly.due + ' 09:00');
  const repeating = all.lists.flatMap(l => l.reminders).find(r => /Repeats /.test(r.notes));
  assert.ok(/set the repeat again in Reminders/.test(repeating.notes));
  // Picking lists, and dropping lists with nothing to send
  const one = core.toRemindersData(m, { listIds: [m.lists[1].id] });
  assert.deepStrictEqual(one.lists.map(l => l.name), [m.lists[1].name]);
  const empty = core.normalize({ rawLists: [{ id: 'x', displayName: 'Done stuff' }], rawTasksByList: { x: [{ id: 't', title: 'old', status: 'completed' }] } });
  assert.strictEqual(core.toRemindersData(empty, {}).lists.length, 0);
  // Google's notes keep their own wording
  assert.ok(/set the repeat again in Google Tasks/.test(core.googleTaskBody(m.lists.flatMap(l => l.tasks).find(t => t.recurrence)).notes));
});

test('Graph addresses in technical details are short', () => {
  const s = MicrosoftClient._test.shortPath('https://graph.microsoft.com/v1.0/me/todo/lists/AAMkADAwATM3ZmYAZS0yZjQ2LTk0ZWEtMDAyAi0wMAoALgAAA%3D%3D/tasks?$expand=checklistItems,linkedResources&$top=100');
  assert.strictEqual(s, '/v1.0/me/todo/lists/…/tasks?$expand=checklistItems,linkedResources&$top=100');
});

// ------------------------------------------------------------------ Google Tasks transfer
function fakeGoogle(options) {
  options = options || {};
  const lists = [{ id: 'gl-0', title: 'My Tasks', tasks: [] }, { id: 'gl-x', title: 'Groceries', tasks: [] }];
  let n = 0, requests = 0, failOnce = new Set(options.failOnce || []);
  const log = [];
  async function fetch(url, init) {
    requests++;
    const u = new URL(url);
    const method = init.method;
    const body = init.body ? JSON.parse(init.body) : null;
    log.push(method + ' ' + u.pathname + u.search);
    if (options.authFailAt && requests === options.authFailAt) return response(401, { error: { code: 401, status: 'UNAUTHENTICATED' } });
    if (options.rateLimitEvery && requests % options.rateLimitEvery === 0 && !failOnce.has('rl' + requests)) {
      failOnce.add('rl' + requests);
      return response(403, { error: { code: 403, errors: [{ reason: 'userRateLimitExceeded' }], message: 'slow down' } });
    }
    if (u.pathname === '/tasks/v1/users/@me/lists' && method === 'GET') return response(200, { items: lists.map(l => ({ id: l.id, title: l.title })) });
    if (u.pathname === '/tasks/v1/users/@me/lists' && method === 'POST') { const l = { id: 'gl-' + (++n), title: body.title, tasks: [] }; lists.push(l); return response(200, { id: l.id, title: l.title }); }
    const m = /^\/tasks\/v1\/lists\/([^/]+)\/tasks$/.exec(u.pathname);
    if (m && method === 'POST') {
      const l = lists.find(x => x.id === decodeURIComponent(m[1]));
      if (!l) return response(404, { error: { message: 'no list' } });
      if (options.rejectTitle && body.title === options.rejectTitle) return response(400, { error: { code: 400, message: 'Invalid value' } });
      const t = { id: 'gt-' + (++n), title: body.title, notes: body.notes, due: body.due, status: body.status, parent: u.searchParams.get('parent') || null };
      const prev = u.searchParams.get('previous');
      const siblings = l.tasks.filter(x => x.parent === t.parent);
      let index;
      if (prev) { index = l.tasks.indexOf(l.tasks.find(x => x.id === prev)) + 1; } else { index = siblings.length ? l.tasks.indexOf(siblings[0]) : l.tasks.length; }
      l.tasks.splice(index, 0, t);
      return response(200, t);
    }
    return response(404, { error: { message: 'unknown ' + u.pathname } });
  }
  return { fetch, lists, log, get requests() { return requests; } };
}

test('Google transfer: order, subtasks, name clashes, completed filter', async () => {
  const fake = fakeGoogle();
  const g = new GoogleClient({ clientId: 'x', fetch: fake.fetch, sleep: async () => {}, minGap: 0, token: 'GT' });
  const opts = { listIds: ['demo-list-tasks', 'demo-list-groceries', 'demo-list-home'], includeCompleted: false };
  const progress = [];
  const result = await g.transfer(model, opts, p => progress.push(p), {});
  const byTitle = t => fake.lists.find(l => l.title === t);
  const tasksList = byTitle('Tasks');
  const groceries = byTitle('Groceries 🛒');
  assert.ok(tasksList && groceries && byTitle('Home & garden'));
  const top = tasksList.tasks.filter(t => !t.parent).map(t => t.title);
  assert.deepStrictEqual(top, ['Renew passport', 'Call the dentist', 'Pay credit card', 'Back up laptop'], 'order kept, completed skipped');
  const passport = tasksList.tasks.find(t => t.title === 'Renew passport');
  const steps = tasksList.tasks.filter(t => t.parent === passport.id).map(t => [t.title, t.status]);
  assert.deepStrictEqual(steps, [['Get photos taken', 'completed'], ['Fill in the form', 'needsAction'], ['Book appointment', 'needsAction']]);
  assert.ok(passport.notes.includes('Marked as important'));
  assert.strictEqual(result.lists, 3);
  assert.strictEqual(result.failures.length, 0);
  assert.ok(result.repeating >= 3);
  const last = progress[progress.length - 1];
  assert.strictEqual(last.done, last.total);
  assert.strictEqual(last.total, GoogleClient.workFor(model, opts));
});

test('Google transfer: clash renames list; failures recorded; retries on rate limit', async () => {
  const clash = core.normalize({ rawLists: [{ id: 'L', displayName: 'Groceries' }], rawTasksByList: { L: [
    { id: 't1', title: 'Apples', status: 'notStarted' }, { id: 't2', title: 'BAD', status: 'notStarted', checklistItems: [{ id: 'c', displayName: 'never' }] }, { id: 't3', title: 'Pears', status: 'notStarted' }] } });
  const fake = fakeGoogle({ rejectTitle: 'BAD', rateLimitEvery: 3 });
  const sleeps = [];
  const g = new GoogleClient({ clientId: 'x', fetch: fake.fetch, sleep: async ms => sleeps.push(ms), minGap: 0, token: 'GT' });
  const r = await g.transfer(clash, { listIds: ['L'], includeCompleted: true }, null, {});
  const list = fake.lists.find(l => l.title === 'Groceries (from To Do)');
  assert.ok(list, 'renamed to avoid clash');
  assert.deepStrictEqual(list.tasks.map(t => t.title), ['Apples', 'Pears']);
  assert.strictEqual(r.failures.length, 1);
  assert.strictEqual(r.failures[0].task, 'BAD');
  assert.deepStrictEqual(r.renamed, [{ from: 'Groceries', to: 'Groceries (from To Do)' }]);
  assert.ok(sleeps.length >= 1, 'backed off on rate limit');
});

test('Google transfer: pauses on expired sign-in and resumes without duplicates', async () => {
  const fake = fakeGoogle({ authFailAt: 9 });
  const g = new GoogleClient({ clientId: 'x', fetch: fake.fetch, sleep: async () => {}, minGap: 0, token: 'GT' });
  const opts = { listIds: ['demo-list-tasks', 'demo-list-party'], includeCompleted: true };
  const state = {};
  await assert.rejects(g.transfer(model, opts, null, state), e => e.code === 'google_auth');
  assert.ok(!state.finished);
  const before = fake.lists.reduce((n, l) => n + l.tasks.length, 0);
  const r = await g.transfer(model, opts, null, state);
  assert.ok(state.finished);
  const tasksList = fake.lists.find(l => l.title === 'Tasks');
  const party = fake.lists.find(l => l.title === "Mum's 60th");
  const expectedTasks = model.lists[0].tasks.length + model.lists[3].tasks.length;
  const expectedSteps = model.lists[0].tasks.concat(model.lists[3].tasks).reduce((n, t) => n + t.steps.length, 0);
  const created = tasksList.tasks.length + party.tasks.length;
  assert.strictEqual(created, expectedTasks + expectedSteps, 'nothing duplicated or missing');
  assert.strictEqual(fake.lists.filter(l => l.title === 'Tasks').length, 1, 'list not created twice');
  assert.ok(before > 0 && before < created);
  assert.strictEqual(r.tasks, expectedTasks);
  assert.strictEqual(r.steps, expectedSteps);
  assert.deepStrictEqual(tasksList.tasks.filter(t => !t.parent).map(t => t.title), model.lists[0].tasks.map(t => t.title));
});

test('Google transfer: API switched off is a pause, not a crash', async () => {
  const g = new GoogleClient({ clientId: 'x', sleep: async () => {}, minGap: 0, token: 'GT',
    fetch: async () => response(403, { error: { code: 403, status: 'PERMISSION_DENIED', message: 'Google Tasks API has not been used in project 123 before or it is disabled.', details: [{ reason: 'SERVICE_DISABLED' }] } }) });
  await assert.rejects(g.transfer(model, { listIds: ['demo-list-tasks'], includeCompleted: false }, null, {}), e => e.code === 'google_api_disabled');
});

test('Google pacing waits between requests', async () => {
  let clock = 0; const sleeps = [];
  const g = new GoogleClient({ clientId: 'x', token: 'GT', minGap: 100, now: () => clock, sleep: async ms => { sleeps.push(ms); clock += ms; },
    fetch: async () => response(200, { items: [] }) });
  await g.call('GET', '/users/@me/lists'); await g.call('GET', '/users/@me/lists'); await g.call('GET', '/users/@me/lists');
  assert.deepStrictEqual(sleeps, [100, 100]);
});

// ------------------------------------------------------------------ run
(async () => {
  for (const [name, fn] of queue) {
    try { await fn(); passed++; console.log('  ok   ' + name); }
    catch (e) { failed++; console.log('  FAIL ' + name + '\n       ' + (e && e.stack || e).toString().split('\n').slice(0, 6).join('\n       ')); }
  }
  console.log(`\n${passed} passed, ${failed} failed (TZ=${process.env.TZ || 'system'})`);
  process.exit(failed ? 1 : 0);
})();
