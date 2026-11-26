"""
Browser tests for Done With To Do, using real Chrome with Microsoft and Google faked.

Run from the repository:  python3 tests/e2e.py
Needs Python Playwright (pip install playwright). Uses Playwright's Chromium, or set CHROME
to a Chrome binary. The site is copied to a temporary folder with an empty config.js, so the
first tests can check the setup screen; later tests swap in a test config.
"""
import json, os, re, shutil, subprocess, sys, tempfile, time, zipfile, io, threading
from urllib.parse import urlparse, parse_qs, unquote
from playwright.sync_api import sync_playwright

PLACEHOLDER_CONFIG = """window.DWTD_CONFIG = {
  microsoftClientId: "PASTE-YOUR-MICROSOFT-CLIENT-ID-HERE",
  googleClientId: "PASTE-YOUR-GOOGLE-CLIENT-ID-HERE",
  googleForEveryone: false,
  donateUrl: "",
  ownerName: "",
  contactEmail: "",
  githubUrl: "",
  microsoftAuthority: "common"
};
"""
SOURCE = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SITE = os.path.join(tempfile.mkdtemp(prefix='dwtd-e2e-'), 'site')
shutil.copytree(SOURCE, SITE, ignore=shutil.ignore_patterns('.git', 'tests', 'node_modules'))
with open(os.path.join(SITE, 'config.js'), 'w') as f:
    f.write(PLACEHOLDER_CONFIG)
CHROME = os.environ.get('CHROME') or ('/usr/bin/google-chrome' if os.path.exists('/usr/bin/google-chrome') else None)
SHOTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'shots')
os.makedirs(SHOTS, exist_ok=True)
PORT = 8765
BASE = f'http://localhost:{PORT}/'
CONFIG_URL = re.compile(r'^' + re.escape(BASE) + r'config\.js(\?.*)?$')

server = subprocess.Popen([sys.executable, '-m', 'http.server', str(PORT), '--bind', '127.0.0.1', '--directory', SITE],
                          stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(0.8)

results = []
def check(name, cond, extra=''):
    results.append((name, bool(cond), extra))
    print(('  ok   ' if cond else '  FAIL ') + name + (('  -- ' + str(extra)) if (extra and not cond) else ''))

CORS = {'Access-Control-Allow-Origin': BASE.rstrip('/'), 'Access-Control-Allow-Headers': 'Authorization, Content-Type, Accept',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS', 'Access-Control-Max-Age': '600'}

TEST_CONFIG = """window.DWTD_CONFIG = {
  microsoftClientId: "11111111-2222-3333-4444-555555555555",
  googleClientId: "123456789012-abcdefghijklmnop.apps.googleusercontent.com",
  donateUrl: "https://ko-fi.com/example", ownerName: "Sam", contactEmail: "sam@example.com",
  githubUrl: "https://github.com/example/done-with-todo", microsoftAuthority: "common" };"""

FAKE_GIS = """window.google = { accounts: { oauth2: {
  initTokenClient: function (cfg) { return { requestAccessToken: function () {
    window.__gisCalls = (window.__gisCalls || 0) + 1;
    setTimeout(function () { cfg.callback({ access_token: 'GTOKEN', expires_in: 3599, scope: cfg.scope }); }, 30); } }; },
  hasGrantedAllScopes: function () { return true; } } } };"""

# ---------------------------------------------------------------- fake Microsoft Graph data
def dt(s): return {'dateTime': s + 'T00:00:00.0000000', 'timeZone': 'UTC'}
LISTS = [
    {'id': 'L1==', 'displayName': 'Work stuff', 'isOwner': True, 'isShared': False, 'wellknownListName': 'defaultList'},
    {'id': 'L2/x', 'displayName': 'Shopping', 'isOwner': True, 'isShared': True, 'wellknownListName': 'none'},
    {'id': 'L3', 'displayName': 'Flagged email', 'isOwner': True, 'isShared': False, 'wellknownListName': 'flaggedEmails'},
]
TASKS = {
    'L1==': [
        {'id': 'T1', 'title': 'Quarterly report <b>draft</b>', 'status': 'notStarted', 'importance': 'high', 'dueDateTime': dt('2026-12-04'),
         'body': {'content': '<p>Use the <i>new</i> template</p>', 'contentType': 'html'}, 'hasAttachments': True,
         'recurrence': {'pattern': {'type': 'weekly', 'interval': 1, 'daysOfWeek': ['friday']}, 'range': {'type': 'noEnd'}},
         'checklistItems': [{'id': 'c1', 'displayName': 'Collect numbers', 'isChecked': True}, {'id': 'c2', 'displayName': 'Write summary', 'isChecked': False}],
         'linkedResources': []},
        {'id': 'T2', 'title': 'Book flights', 'status': 'completed', 'completedDateTime': dt('2026-11-20'), 'checklistItems': [], 'linkedResources': []},
        {'id': 'T3', 'title': 'Email Priya', 'status': 'inProgress', 'isReminderOn': True,
         'reminderDateTime': {'dateTime': '2026-11-27T09:30:00.0000000', 'timeZone': 'UTC'}, 'checklistItems': [], 'linkedResources': []},
    ],
    'L2/x': [{'id': 'S%d' % i, 'title': 'Item %d' % i, 'status': 'notStarted', 'checklistItems': [], 'linkedResources': []} for i in range(1, 121)],
    'L3': [{'id': 'F1', 'title': 'Re: invoice', 'status': 'notStarted', 'checklistItems': [], 'linkedResources': [{'webUrl': 'https://outlook.live.com/x', 'displayName': 'invoice'}]}],
}
ATTACHMENT_BYTES = b'%PDF-1.4 fake pdf bytes'
graph_log = []
token_log = []

def graph_handler(route, request):
    if request.method == 'OPTIONS':
        return route.fulfill(status=204, headers=CORS)
    u = urlparse(request.url)
    path = unquote(u.path)
    q = parse_qs(u.query)
    graph_log.append(request.method + ' ' + request.url)
    auth = request.headers.get('authorization', '')
    if auth not in ('Bearer MS-ACCESS-TOKEN', 'Bearer MS-RW-TOKEN'):
        return route.fulfill(status=401, headers=CORS, json={'error': {'code': 'InvalidAuthenticationToken', 'message': 'bad token'}})
    if path == '/v1.0/me':
        return route.fulfill(status=200, headers=CORS, json={'displayName': 'Sam Smith', 'mail': None, 'userPrincipalName': 'sam@outlook.com'})
    if path == '/v1.0/me/todo/lists':
        return route.fulfill(status=200, headers=CORS, json={'value': LISTS})
    m = re.match(r'^/v1\.0/me/todo/lists/(.+)/tasks$', path)
    if m:
        lid = m.group(1)
        skip = int(q.get('$skip', ['0'])[0])
        items = TASKS.get(lid, [])
        page = items[skip:skip + 100]
        body = {'value': page}
        if skip + 100 < len(items):
            body['@odata.nextLink'] = f'https://graph.microsoft.com/v1.0/me/todo/lists/{lid}/tasks?$expand=checklistItems,linkedResources&$top=100&$skip={skip + 100}'
        return route.fulfill(status=200, headers=CORS, json=body)
    m = re.match(r'^/v1\.0/me/todo/lists/(.+)/tasks/(.+)/attachments$', path)
    if m:
        return route.fulfill(status=200, headers=CORS, json={'value': [{'id': 'A1', 'name': 'template.pdf', 'size': len(ATTACHMENT_BYTES), 'contentType': 'application/pdf'}]})
    m = re.match(r'^/v1\.0/me/todo/lists/(.+)/tasks/(.+)/attachments/A1/\$value$', path)
    if m:
        return route.fulfill(status=200, headers=dict(CORS, **{'Content-Type': 'application/pdf'}), body=ATTACHMENT_BYTES)
    return route.fulfill(status=404, headers=CORS, json={'error': {'code': 'NotFound', 'message': path}})

def login_handler(route, request):
    u = urlparse(request.url)
    if request.method == 'OPTIONS':
        return route.fulfill(status=204, headers=CORS)
    if u.path.endswith('/authorize'):
        q = parse_qs(u.query)
        assert q['code_challenge_method'] == ['S256']
        target = q['redirect_uri'][0] + '#code=AUTHCODE&state=' + q['state'][0] + '&session_state=x'
        return route.fulfill(status=302, headers={'Location': target})
    if u.path.endswith('/token'):
        form = parse_qs(request.post_data or '')
        ok = form.get('code') == ['AUTHCODE'] and len(form.get('code_verifier', [''])[0]) >= 43 and form.get('grant_type') == ['authorization_code']
        if not ok:
            return route.fulfill(status=400, headers=CORS, json={'error': 'invalid_grant', 'error_description': 'AADSTS54005: bad'})
        wide = 'Tasks.ReadWrite' in form.get('scope', [''])[0]
        token_log.append(form.get('scope', [''])[0])
        return route.fulfill(status=200, headers=CORS, json={'token_type': 'Bearer', 'access_token': 'MS-RW-TOKEN' if wide else 'MS-ACCESS-TOKEN', 'expires_in': 3600,
                                                            'scope': 'Tasks.ReadWrite User.Read' if wide else 'Tasks.Read User.Read'})
    return route.fulfill(status=404, body='no')

# ---------------------------------------------------------------- fake Google Tasks
G = {'lists': [{'id': 'g-default', 'title': 'My Tasks', 'tasks': []}, {'id': 'g-old', 'title': 'Shopping', 'tasks': []}], 'n': 0, 'calls': 0}
def tasks_handler(route, request):
    if request.method == 'OPTIONS':
        return route.fulfill(status=204, headers=CORS)
    G['calls'] += 1
    if request.headers.get('authorization') != 'Bearer GTOKEN':
        return route.fulfill(status=401, headers=CORS, json={'error': {'code': 401, 'status': 'UNAUTHENTICATED'}})
    u = urlparse(request.url)
    q = parse_qs(u.query)
    body = json.loads(request.post_data) if request.post_data else None
    if G['calls'] == 7:  # one rate-limit hiccup
        return route.fulfill(status=429, headers=CORS, json={'error': {'code': 429, 'status': 'RESOURCE_EXHAUSTED'}})
    if u.path == '/tasks/v1/users/@me/lists' and request.method == 'GET':
        return route.fulfill(status=200, headers=CORS, json={'items': [{'id': l['id'], 'title': l['title']} for l in G['lists']]})
    if u.path == '/tasks/v1/users/@me/lists' and request.method == 'POST':
        G['n'] += 1
        l = {'id': 'gl-%d' % G['n'], 'title': body['title'], 'tasks': []}
        G['lists'].append(l)
        return route.fulfill(status=200, headers=CORS, json={'id': l['id'], 'title': l['title']})
    m = re.match(r'^/tasks/v1/lists/([^/]+)/tasks$', u.path)
    if m and request.method == 'POST':
        l = next(x for x in G['lists'] if x['id'] == unquote(m.group(1)))
        G['n'] += 1
        t = dict(body, id='gt-%d' % G['n'], parent=q.get('parent', [None])[0])
        prev = q.get('previous', [None])[0]
        if prev:
            idx = [x['id'] for x in l['tasks']].index(prev) + 1
        else:
            sib = [x for x in l['tasks'] if x['parent'] == t['parent']]
            idx = l['tasks'].index(sib[0]) if sib else len(l['tasks'])
        l['tasks'].insert(idx, t)
        return route.fulfill(status=200, headers=CORS, json=t)
    return route.fulfill(status=404, headers=CORS, json={'error': {'message': 'unknown'}})

def watch(page, bag):
    page.on('console', lambda m: bag.append(('console.' + m.type, m.text)) if m.type in ('error', 'warning') else None)
    page.on('pageerror', lambda e: bag.append(('pageerror', str(e))))

def verify_zip(data, expect_attachment=None):
    z = zipfile.ZipFile(io.BytesIO(data))
    assert z.testzip() is None
    names = z.namelist()
    js = json.loads(z.read([n for n in names if n.endswith('.json')][0]))
    att = [n for n in names if '/Attachments/' in n]
    if expect_attachment:
        assert any(n.endswith(expect_attachment) for n in att), att
    return names, js, z

with sync_playwright() as p:
    browser = p.chromium.launch(executable_path=CHROME, headless=True, args=['--no-sandbox'])

    # ------------------------------------------------ A. fresh site (placeholders): setup mode + demo
    ctx = browser.new_context(accept_downloads=True, viewport={'width': 1280, 'height': 900})
    page = ctx.new_page(); bag = []; watch(page, bag)
    page.goto(BASE)
    page.wait_for_selector('#view-start:not([hidden])')
    check('A setup box shows on an unconfigured site', page.is_visible('#setup'))
    check('A setup box shows the exact site address', page.text_content('#setup-address') == BASE, page.text_content('#setup-address'))
    check('A setup box shows the origin', page.text_content('#setup-origin') == BASE.rstrip('/'))
    check('A sign-in disabled until configured', page.is_disabled('#signin-ms'))
    check('A Google shows Coming soon when not configured', page.text_content('#dest-google-state') == 'Coming soon')
    page.screenshot(path=os.path.join(SHOTS, 'a-start-setup.png'), full_page=True)
    page.click('#try-demo')
    page.wait_for_selector('#view-overview:not([hidden])')
    check('A demo overview summary', 'You have 5 lists with 20 tasks' in page.text_content('#ov-summary'), page.text_content('#ov-summary'))
    rows = page.query_selector_all('#ov-lists tr')
    check('A five list rows', len(rows) == 5)
    check('A flagged-email list unchecked by default', not page.is_checked('#list-4') and page.is_checked('#list-0'))
    check('A Google button hidden when not configured', not page.is_visible('#google-move'))
    page.screenshot(path=os.path.join(SHOTS, 'a-overview-demo.png'), full_page=True)
    with page.expect_download() as dl:
        page.click('#backup-zip')
    path = dl.value.path()
    names, js, z = verify_zip(open(path, 'rb').read(), 'tap-model.pdf')
    check('A backup zip valid with 20 tasks and attachments', js['summary']['tasks'] == 20 and len(names) >= 16, len(names))
    check('A backup filename', re.match(r'Done With To Do backup \d{4}-\d{2}-\d{2}\.zip', dl.value.suggested_filename), dl.value.suggested_filename)
    page.wait_for_selector('#backup-status:has-text("Saved")')
    check('A thanks card hidden on unconfigured site without links', page.is_hidden('#backup-thanks') or 'github' in (page.inner_html('#backup-thanks') or ''))
    with page.expect_popup() as pop:
        page.click('#backup-print')
    pp = pop.value; pp.wait_for_load_state()
    check('A printable page opens', 'My Microsoft To Do lists' in pp.text_content('h1'))
    pp.screenshot(path=os.path.join(SHOTS, 'a-printable.png'), full_page=False)
    pp.close()
    with page.expect_download() as dl2:
        page.click('#backup-csv')
    csv = open(dl2.value.path(), 'rb').read().decode('utf-8-sig')
    check('A CSV downloads with header', csv.startswith('List,Task,Done'), csv[:40])
    check('A no console errors (demo)', not [b for b in bag if b[0] in ('pageerror', 'console.error')], bag)
    # mobile layout
    m = browser.new_context(viewport={'width': 390, 'height': 844}, device_scale_factor=2)
    mp = m.new_page(); mp.goto(BASE); mp.wait_for_selector('#view-start:not([hidden])')
    mp.screenshot(path=os.path.join(SHOTS, 'a-start-mobile.png'), full_page=True)
    mp.goto(BASE + '?demo'); mp.wait_for_selector('#view-overview:not([hidden])')
    mp.screenshot(path=os.path.join(SHOTS, 'a-overview-mobile.png'), full_page=True)
    m.close(); ctx.close()

    # ------------------------------------------------ B. configured site: real sign-in flow against fakes
    ctx = browser.new_context(accept_downloads=True, viewport={'width': 1280, 'height': 900})
    ctx.route(CONFIG_URL, lambda r, q: r.fulfill(status=200, content_type='application/javascript', body=TEST_CONFIG))
    ctx.route('https://login.microsoftonline.com/**', login_handler)
    ctx.route('https://graph.microsoft.com/**', graph_handler)
    ctx.route('https://accounts.google.com/gsi/client', lambda r, q: r.fulfill(status=200, content_type='application/javascript', body=FAKE_GIS))
    ctx.route('https://tasks.googleapis.com/**', tasks_handler)
    page = ctx.new_page(); bag = []; watch(page, bag)
    page.goto(BASE)
    page.wait_for_selector('#view-start:not([hidden])')
    check('B setup box hidden once configured', page.is_hidden('#setup'))
    check('B header shows source and tip links', page.is_visible('#header-source') and page.is_visible('#header-donate'))
    page.screenshot(path=os.path.join(SHOTS, 'b-start-configured.png'), full_page=False)
    page.click('#signin-ms')
    page.wait_for_selector('#view-overview:not([hidden])', timeout=20000)
    check('B URL cleaned after sign-in', '#' not in page.url and 'code=' not in page.url, page.url)
    check('B signed-in name shown', 'Sam Smith' in page.text_content('#ov-account'), page.text_content('#ov-account'))
    check('B all 124 tasks read across pages', 'with 124 tasks' in page.text_content('#ov-summary'), page.text_content('#ov-summary'))
    check('B list ids URL-encoded for Graph', any('/lists/L2%2Fx/tasks' in c for c in graph_log), [c for c in graph_log if 'L2' in c][:2])
    check('B Google button counts lists', page.text_content('#google-move') == 'Move 2 lists to Google Tasks', page.text_content('#google-move'))
    page.screenshot(path=os.path.join(SHOTS, 'b-overview.png'), full_page=True)
    # backup with a real attachment download
    with page.expect_download() as dl:
        page.click('#backup-zip')
    names, js, z = verify_zip(open(dl.value.path(), 'rb').read(), 'template.pdf')
    att = [n for n in names if n.endswith('template.pdf')][0]
    check('B attachment bytes in zip', z.read(att) == ATTACHMENT_BYTES)
    check('B raw Graph data kept in JSON', js['raw']['lists'][0]['id'] == 'L1==')
    t1 = js['lists'][0]['tasks'][0]
    check('B HTML notes converted to text', t1['notes'] == 'Use the new template', t1['notes'])
    # Google transfer
    page.click('#google-move')
    page.wait_for_selector('#g-done:not([hidden])', timeout=60000)
    page.wait_for_timeout(900)
    page.screenshot(path=os.path.join(SHOTS, 'b-google-done.png'), full_page=True)
    work = next(l for l in G['lists'] if l['title'] == 'Work stuff')
    shop = next((l for l in G['lists'] if l['title'] == 'Shopping (from To Do)'), None)
    check('B list name clash handled', shop is not None, [l['title'] for l in G['lists']])
    top = [t['title'] for t in work['tasks'] if not t['parent']]
    check('B completed task skipped, order kept', top == ['Quarterly report <b>draft</b>', 'Email Priya'], top)
    q = next(t for t in work['tasks'] if t['title'].startswith('Quarterly'))
    subs = [(t['title'], t['status']) for t in work['tasks'] if t['parent'] == q['id']]
    check('B steps became ordered subtasks', subs == [('Collect numbers', 'completed'), ('Write summary', 'needsAction')], subs)
    check('B due date sent', q.get('due') == '2026-12-04T00:00:00.000Z', q.get('due'))
    check('B repeat written into notes', 'Repeats every week on Friday' in q.get('notes', ''), q.get('notes'))
    check('B 120 shopping items in order', shop and [t['title'] for t in shop['tasks']][:3] == ['Item 1', 'Item 2', 'Item 3'] and len(shop['tasks']) == 120)
    check('B flagged emails not moved', not any(l['title'].startswith('Flagged') for l in G['lists']))
    summary = page.text_content('#g-done-summary')
    check('B done summary', 'Copied 122 tasks and 2 steps into 2 new lists' in summary, summary)
    notes = page.text_content('#g-done-notes')
    check('B done notes mention repeat, rename and account hint', '1 task repeats' in notes and 'Shopping (from To Do)' in notes and 'different Google account' in notes, notes)
    check('B thanks card with star + tip', page.is_visible('#g-thanks a:has-text("Star it on GitHub")') and page.is_visible('#g-thanks a:has-text("Leave a tip")'))
    page.click('#g-back')
    page.wait_for_selector('#view-overview:not([hidden])')
    check('B moved lists tagged and unticked', page.is_visible('text=Moved to Google Tasks') and not page.is_checked('#list-0'))
    check('B no console errors (configured flow)', not [b for b in bag if b[0] == 'pageerror' or (b[0] == 'console.error' and '429' not in b[1])], bag)
    ctx.close()

    # ------------------------------------------------ C. sign-in error from Microsoft is explained
    ctx = browser.new_context()
    ctx.route(CONFIG_URL, lambda r, q: r.fulfill(status=200, content_type='application/javascript', body=TEST_CONFIG))
    page = ctx.new_page()
    page.goto(BASE + '#error=invalid_request&error_description=' + 'AADSTS50011%3A+The+redirect+URI+does+not+match')
    page.wait_for_selector('#view-error:not([hidden])')
    check('C redirect mismatch explained with exact address', BASE in page.text_content('#err-message'), page.text_content('#err-message'))
    page.screenshot(path=os.path.join(SHOTS, 'c-error.png'), full_page=False)
    ctx.close()

    # ------------------------------------------------ D. privacy page fills owner details
    ctx = browser.new_context()
    ctx.route(CONFIG_URL, lambda r, q: r.fulfill(status=200, content_type='application/javascript', body=TEST_CONFIG))
    page = ctx.new_page(); bag = []; watch(page, bag)
    page.goto(BASE + 'privacy.html')
    check('D privacy owner and email', 'run by Sam (sam@example.com)' in page.text_content('main'), page.text_content('main')[:300])
    check('D privacy no console errors', not [b for b in bag if b[0] in ('pageerror', 'console.error')], bag)
    page.screenshot(path=os.path.join(SHOTS, 'd-privacy.png'), full_page=False)
    ctx.close()
    # ------------------------------------------------ E. Google hidden from the public until approved
    GATED = TEST_CONFIG.replace('microsoftAuthority: "common"', 'microsoftAuthority: "common", googleForEveryone: false')
    ctx = browser.new_context()
    ctx.route(CONFIG_URL, lambda r, q: r.fulfill(status=200, content_type='application/javascript', body=GATED))
    page = ctx.new_page()
    page.goto(BASE + '?demo'); page.wait_for_selector('#view-overview:not([hidden])')
    check('E Google hidden from the public while googleForEveryone is false', not page.is_visible('#google-move') and 'Coming soon' in page.text_content('#google-status'))
    page.goto(BASE + '?demo&testers'); page.wait_for_selector('#view-overview:not([hidden])')
    check('E testers see Google', page.is_visible('#google-move'))
    page.goto(BASE + '?demo'); page.wait_for_selector('#view-overview:not([hidden])')
    check('E tester flag remembered in the same tab', page.is_visible('#google-move'))
    ctx.close()
    # a blocker that swaps Google's script for an empty one gets a plain explanation, not a JS error
    ctx = browser.new_context()
    ctx.route(CONFIG_URL, lambda r, q: r.fulfill(status=200, content_type='application/javascript', body=TEST_CONFIG))
    ctx.route('https://accounts.google.com/gsi/client', lambda r, q: r.fulfill(status=200, content_type='application/javascript', body='/* blocked */'))
    page = ctx.new_page()
    page.goto(BASE + '?demo'); page.wait_for_selector('#view-overview:not([hidden])')
    page.wait_for_function("document.getElementById('google-status').textContent.includes('Google sign-in')", timeout=5000)
    gs = page.text_content('#google-status')
    check('E emptied Google script explained plainly', 'blocking accounts.google.com' in gs and 'undefined' not in gs, gs)
    ctx.close()

    # ------------------------------------------------ F. Microsoft refuses attachments (401 with a fresh sign-in)
    def refuse_attachments(route, request):
        if request.method != 'OPTIONS' and request.headers.get('authorization') == 'Bearer MS-ACCESS-TOKEN' and re.search(r'/me/todo/lists/.+/tasks/.+/attachments', unquote(urlparse(request.url).path)):
            graph_log.append('REFUSED ' + request.url)
            return route.fulfill(status=401, headers=CORS, json={'error': {'code': 'UnknownError', 'message': ''}})
        return graph_handler(route, request)
    ctx = browser.new_context(accept_downloads=True, viewport={'width': 1280, 'height': 900})
    ctx.route(CONFIG_URL, lambda r, q: r.fulfill(status=200, content_type='application/javascript', body=TEST_CONFIG))
    ctx.route('https://login.microsoftonline.com/**', login_handler)
    ctx.route('https://graph.microsoft.com/**', refuse_attachments)
    ctx.route('https://accounts.google.com/gsi/client', lambda r, q: r.fulfill(status=200, content_type='application/javascript', body=FAKE_GIS))
    page = ctx.new_page(); bag = []; watch(page, bag)
    page.goto(BASE)
    page.wait_for_selector('#view-start:not([hidden])')
    page.click('#signin-ms')
    page.wait_for_selector('#view-overview:not([hidden]), #view-error:not([hidden])', timeout=20000)
    check('F lists still load when Microsoft refuses attachments', page.is_visible('#view-overview'), page.text_content('#err-message') if page.is_visible('#view-error') else '')
    check('F beta address tried for attachments', any('graph.microsoft.com/beta/' in c for c in graph_log))
    warn = page.text_content('#ov-att-text') or ''
    check('F warning names the problem and the fix', page.is_visible('#ov-att-warning') and 'on 1 task' in warn and 'permission to change tasks' in warn and 'only reads' in warn, warn)
    page.click('#ov-att-warning summary')
    det = page.text_content('#ov-att-details') or ''
    check('F technical details name the request', 'Work stuff / Quarterly report' in det and '/attachments' in det and '401' in det, det)
    page.screenshot(path=os.path.join(SHOTS, 'f-attachments-refused.png'), full_page=True)
    check('F attachment option hidden (nothing to download)', page.is_hidden('#opt-attachments-row'))
    with page.expect_download() as dl:
        page.click('#backup-zip')
    names, js, z = verify_zip(open(dl.value.path(), 'rb').read())
    readme = z.read([n for n in names if n.endswith('README.txt')][0]).decode('utf-8')
    check('F README lists the task to save by hand', "wouldn't share" in readme and 'Work stuff / Quarterly report <b>draft</b>' in readme, readme[-400:])
    check('F backup still has every task', js['summary']['tasks'] == 124, js['summary']['tasks'])
    page.wait_for_selector('#backup-status:has-text("Saved")')
    check('F no page errors', not [b for b in bag if b[0] == 'pageerror' or (b[0] == 'console.error' and '401' not in b[1])], bag)
    # The button signs in again with the permission Microsoft wants, and the files come through
    check('F offers to include attached files', page.is_visible('#ov-att-allow') and 'permission to change tasks' in (page.text_content('#ov-att-text') or ''))
    page.click('#ov-att-allow')
    page.wait_for_selector('#view-overview:not([hidden])', timeout=20000)
    page.wait_for_function("document.getElementById('ov-att-warning').hidden", timeout=10000)
    check('F second sign-in asked for Tasks.ReadWrite', token_log and 'Tasks.ReadWrite' in token_log[-1], token_log)
    check('F attached file offered after allowing', page.is_visible('#opt-attachments-row') and '1 attached file' in page.text_content('#opt-attachments-label'), page.text_content('#opt-attachments-label'))
    with page.expect_download() as dl2:
        page.click('#backup-zip')
    names2, js2, z2 = verify_zip(open(dl2.value.path(), 'rb').read(), 'template.pdf')
    check('F attachment bytes in zip after allowing', z2.read([n for n in names2 if n.endswith('template.pdf')][0]) == ATTACHMENT_BYTES)
    graph_writes = [c for c in graph_log if not c.startswith(('GET ', 'OPTIONS ', 'REFUSED '))]
    check('F the site only ever reads from Microsoft', not graph_writes, graph_writes[:3])
    page.screenshot(path=os.path.join(SHOTS, 'f-attachments-allowed.png'), full_page=False)
    ctx.close()
    # ------------------------------------------------ G. On an iPhone
    IPHONE = 'Mozilla/5.0 (iPhone; CPU iPhone OS 26_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/26.1 Mobile/15E148 Safari/604.1'
    ctx = browser.new_context(user_agent=IPHONE, viewport={'width': 390, 'height': 844})
    ctx.route(CONFIG_URL, lambda r, q: r.fulfill(status=200, content_type='application/javascript', body=TEST_CONFIG))
    ctx.route('https://accounts.google.com/gsi/client', lambda r, q: r.fulfill(status=200, content_type='application/javascript', body=FAKE_GIS))
    page = ctx.new_page(); bag = []; watch(page, bag)
    page.goto(BASE); page.wait_for_selector('#view-start:not([hidden])')
    check('G start page promises nothing for Apple Reminders', 'Apple Reminders' not in page.text_content('body'))
    page.goto(BASE + '?demo'); page.wait_for_selector('#view-overview:not([hidden])')
    check('G phone sees the backup and Google Tasks buttons', page.is_visible('#backup-zip') and page.is_visible('#google-move'))
    check('G no page errors on a phone', not [b for b in bag if b[0] in ('pageerror', 'console.error')], bag)
    page.screenshot(path=os.path.join(SHOTS, 'g-iphone.png'), full_page=True)
    ctx.close()
    browser.close()

server.terminate()
shutil.rmtree(os.path.dirname(SITE), ignore_errors=True)
failed = [r for r in results if not r[1]]
print(f'\n{len(results) - len(failed)} passed, {len(failed)} failed')
sys.exit(1 if failed else 0)
