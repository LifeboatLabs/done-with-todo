/*
 * Done With To Do: the page.
 * Wires the buttons to the sign-in, reading, backup and Google Tasks code.
 * All text from your tasks is added with textContent, never as HTML.
 * MIT licence.
 */
(function () {
  'use strict';

  var cfg = window.DWTD_CONFIG || {};
  var D = window.DWTD || {};
  var core = D.core, Zip = D.zip, MicrosoftClient = D.MicrosoftClient, GoogleClient = D.GoogleClient;

  var GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  var GOOGLE_ID = /^[0-9]+-[0-9a-z_]+\.apps\.googleusercontent\.com$/i;
  var msId = String(cfg.microsoftClientId || '').trim();
  var gId = String(cfg.googleClientId || '').trim();
  var msOk = GUID.test(msId);
  // Until Google approves the app, only testers (who open the site with ?testers) see Google Tasks.
  var testers = false;
  try {
    if (/[?&]testers\b/.test(location.search)) sessionStorage.setItem('dwtd.testers', '1');
    testers = sessionStorage.getItem('dwtd.testers') === '1';
  } catch (e) { testers = /[?&]testers\b/.test(location.search); }
  var gIdOk = GOOGLE_ID.test(gId);
  var gOk = gIdOk && (cfg.googleForEveryone !== false || testers);

  var state = {
    ms: null, google: null, raw: null, model: null, demo: false,
    selected: new Set(), moved: new Set(),
    gOpts: null, gState: null, gError: null, transferring: false
  };

  // ------------------------------------------------------------ small helpers

  function $(id) { return document.getElementById(id); }
  function el(tag, attrs, children) {
    var node = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      if (k === 'text') node.textContent = attrs[k];
      else if (k === 'class') node.className = attrs[k];
      else node.setAttribute(k, attrs[k]);
    });
    (children || []).forEach(function (c) { node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); });
    return node;
  }
  function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); }
  var plural = core.plural;

  function formatBytes(n) {
    if (n < 1024) return n + ' bytes';
    if (n < 1024 * 1024) return Math.round(n / 1024) + ' KB';
    return (n / 1024 / 1024).toFixed(1) + ' MB';
  }

  /** This page's address, ending in "/". It's what Microsoft sends people back to. */
  function siteAddress() {
    var p = location.pathname;
    if (!/\/$/.test(p)) p = /\.[a-z0-9]+$/i.test(p) ? p.replace(/[^/]*$/, '') : p + '/';
    return location.origin + p;
  }

  function sourceUrl() {
    if (/^https:\/\/github\.com\/[^/]+\/[^/]+/.test(cfg.githubUrl || '')) return cfg.githubUrl;
    var host = location.hostname;
    if (/\.github\.io$/i.test(host)) {
      var user = host.split('.')[0];
      var repo = location.pathname.split('/').filter(Boolean)[0];
      if (repo && /\.html?$/i.test(repo)) repo = '';
      return 'https://github.com/' + user + '/' + (repo || host);
    }
    return '';
  }
  function donateUrl() { return /^https:\/\//.test(cfg.donateUrl || '') ? cfg.donateUrl : ''; }

  function show(view) {
    ['start', 'loading', 'overview', 'google', 'error'].forEach(function (v) { $('view-' + v).hidden = v !== view; });
    window.scrollTo(0, 0);
    var heading = $('view-' + view).querySelector('h1:not([hidden])');
    if (heading && document.activeElement !== heading) heading.focus({ preventScroll: true });
  }

  function setBusy(btn, busy, label) {
    if (busy) {
      btn.dataset.label = btn.textContent;
      btn.textContent = label || btn.textContent;
      btn.setAttribute('aria-busy', 'true');
      btn.disabled = true;
    } else {
      if (btn.dataset.label) btn.textContent = btn.dataset.label;
      btn.removeAttribute('aria-busy');
      btn.disabled = false;
    }
  }

  function saveBlob(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = el('a', { href: url, download: name });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 60000);
  }

  function copyText(id, button) {
    var text = $(id).textContent;
    var done = function () { var old = button.textContent; button.textContent = 'Copied'; setTimeout(function () { button.textContent = old; }, 1500); };
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(done, function () { selectText(id); });
    else selectText(id);
  }
  function selectText(id) {
    var range = document.createRange(); range.selectNodeContents($(id));
    var sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
  }

  // ------------------------------------------------------------ setup mode (for the site owner)

  function idState(value, good, otherKind) {
    if (!value || /^PASTE-/i.test(value)) return 'missing';
    if (good.test(value)) return 'ok';
    if (otherKind.test(value)) return 'swapped';
    return 'bad';
  }

  function renderSetup() {
    var ms = idState(msId, GUID, GOOGLE_ID), g = idState(gId, GOOGLE_ID, GUID);
    var asFile = location.protocol === 'file:';
    var wanted = ms !== 'ok' || asFile || /[?&]setup\b/.test(location.search);
    $('setup').hidden = !wanted;
    if (!wanted) return;
    $('setup-address').textContent = siteAddress();
    $('setup-origin').textContent = location.origin;
    var msText = {
      ok: 'Microsoft client ID: set. Sign-in is ready to test.',
      missing: 'Microsoft client ID: not set yet. Paste it into config.js (setup guide, Steps 2 and 4).',
      swapped: 'Microsoft client ID: this looks like the Google ID. Swap the two values in config.js.',
      bad: 'Microsoft client ID: this doesn’t look right. It should look like 1a2b3c4d-5e6f-7a8b-9c0d-1e2f3a4b5c6d.'
    }[ms];
    var gText = {
      ok: 'Google client ID: set. Moving to Google Tasks is ready to test.',
      missing: 'Google client ID: not set yet. Until it is, visitors see "Coming soon" for Google Tasks (setup guide, Step 3).',
      swapped: 'Google client ID: this looks like the Microsoft ID. Swap the two values in config.js.',
      bad: 'Google client ID: this doesn’t look right. It should end in .apps.googleusercontent.com.'
    }[g];
    $('setup-ms-status').textContent = msText;
    if (g === 'ok' && cfg.googleForEveryone === false) {
      gText += ' Only testers see it for now (open the site with ?testers). Change googleForEveryone to true once Google approves the app.';
    }
    $('setup-google-status').textContent = gText;
    $('setup-donate-status').textContent = donateUrl()
      ? 'Tip button: set.'
      : 'Tip button: not set (optional). Add your Ko-fi or GitHub Sponsors link as donateUrl in config.js.';
    $('setup-file-warning').hidden = !asFile;
  }

  function renderLinks() {
    var src = sourceUrl(), tip = donateUrl();
    [['header-source', src], ['footer-source', src], ['header-donate', tip], ['footer-donate', tip]].forEach(function (pair) {
      var a = $(pair[0]);
      if (pair[1]) { a.href = pair[1]; a.hidden = false; a.target = '_blank'; a.rel = 'noopener'; }
    });
    if (src) {
      $('trust-source').href = src;
      document.querySelectorAll('.source-sentence').forEach(function (n) { n.hidden = false; });
    }
    if (tip) {
      document.querySelectorAll('.donate-sentence').forEach(function (n) { n.hidden = false; });
      document.querySelectorAll('.donate-link').forEach(function (a) { a.href = tip; a.target = '_blank'; a.rel = 'noopener'; });
    }
    if (!gOk) {
      $('dest-google-state').textContent = 'Coming soon';
      $('dest-google-state').closest('.direction').classList.add('direction-unavailable');
    }
  }

  // ------------------------------------------------------------ start

  function renderStart() {
    var btn = $('signin-ms');
    if (!msOk) {
      btn.disabled = true;
      $('signin-note').textContent = 'Sign-in isn’t set up on this copy of the site yet. You can still try it with sample data.';
    }
    show('start');
  }

  async function signInWithMicrosoft() {
    var btn = $('signin-ms');
    if (!state.ms) return;
    setBusy(btn, true, 'Opening Microsoft…');
    try { location.assign(await state.ms.signInUrl()); }
    catch (e) { setBusy(btn, false); showError(e, 'Couldn’t start signing in'); }
  }

  // ------------------------------------------------------------ reading from Microsoft

  function setLoading(title, detail, fraction) {
    $('loading-title').textContent = title;
    $('loading-detail').textContent = detail;
    var bar = $('loading-bar');
    if (fraction == null) bar.removeAttribute('value');
    else { bar.max = 1; bar.value = Math.max(0.02, Math.min(1, fraction)); }
  }

  async function readFromMicrosoft() {
    show('loading');
    setLoading('Reading your lists', 'Connecting to Microsoft.', null);
    try {
      var raw = await state.ms.readEverything(function (p) {
        if (p.phase === 'lists') setLoading('Reading your lists', 'Getting the names of your lists.', null);
        if (p.phase === 'tasks') {
          setLoading('Reading your lists', 'List ' + (p.index + 1) + ' of ' + p.count + ': “' + p.listName + '”. ' +
            plural(p.tasks, 'task') + ' so far.', p.index / Math.max(1, p.count));
        }
      });
      state.raw = raw;
      state.demo = false;
      state.model = core.normalize(raw);
      showOverview();
    } catch (e) {
      showError(e, 'Couldn’t read your lists');
    }
  }

  function startDemo() {
    state.raw = core.sampleData();
    state.demo = true;
    state.model = core.normalize(state.raw);
    showOverview();
  }

  // ------------------------------------------------------------ overview

  function selectedIds() {
    return state.model.lists.filter(function (l) { return state.selected.has(l.id); }).map(function (l) { return l.id; });
  }

  function showOverview() {
    var m = state.model, s = core.summarize(m);
    if (state.demo) $('ov-account').textContent = 'This is sample data, so you can see how it works. Nothing here is real.';
    else if (m.account && (m.account.name || m.account.email)) {
      $('ov-account').textContent = 'Signed in as ' + (m.account.name || m.account.email) + (m.account.name && m.account.email ? ' (' + m.account.email + ')' : '') + '.';
    } else $('ov-account').textContent = 'Signed in with Microsoft.';

    var bits = [plural(s.tasks, 'task') + ' (' + s.open + ' open, ' + s.done + ' done)'];
    if (s.steps) bits.push(plural(s.steps, 'step'));
    if (s.attachments) bits.push(plural(s.attachments, 'attachment'));
    $('ov-summary').textContent = 'You have ' + plural(s.lists, 'list') + ' with ' +
      (bits.length > 1 ? bits.slice(0, -1).join(', ') + ' and ' + bits[bits.length - 1] : bits[0]) + '.';
    $('ov-empty').hidden = s.tasks > 0;

    // The box is new: a page saved by the browser before it existed won't have it.
    var heldBack = s.attachmentsUnavailable, attBox = $('ov-att-warning');
    if (attBox) attBox.hidden = !heldBack;
    if (heldBack && attBox) {
      $('ov-att-text').textContent = 'Microsoft wouldn’t share the attached files on ' + plural(heldBack, 'task') +
        ', so they can’t go in the backup. Everything else is fine. Open ' + (heldBack === 1 ? 'that task' : 'those tasks') +
        ' in To Do and save the files by hand. The backup’s README lists ' + (heldBack === 1 ? 'it' : 'them') + '.';
      $('ov-att-details').textContent = ((state.raw && state.raw.attachmentProblems) || []).map(function (p) {
        return p.list + ' / ' + p.task + '\n' + p.details;
      }).join('\n\n');
    }

    if (!state.selected.size) {
      state.selected = new Set(m.lists.filter(function (l) { return l.wellknown !== 'flaggedEmails' && !state.moved.has(l.id); }).map(function (l) { return l.id; }));
    }
    renderLists();

    var hasFiles = s.attachments > 0;
    $('opt-attachments-row').hidden = !hasFiles;
    $('backup-att-phrase').hidden = !hasFiles;
    if (hasFiles) $('opt-attachments-label').textContent = 'Include ' + plural(s.attachments, 'attached file') + ' (' + formatBytes(s.attachmentBytes) + ')';

    var gBtn = $('google-move');
    if (!gOk) {
      $('google-title').closest('.dest').classList.add('dest-unavailable');
      $('google-status').textContent = 'Coming soon to this site. Save a backup now and you’ll be ready.';
    } else if (!state.google) {
      state.google = new GoogleClient({ clientId: gId });
      state.google.prepare().catch(function (e) { $('google-status').textContent = e.message; $('google-status').classList.add('status-error'); });
    }
    if (gOk && state.demo) {
      $('google-status').textContent = 'This is sample data. Moving it adds the sample lists to your Google Tasks, which is a handy way to test.';
    }
    gBtn.hidden = !gOk;
    $('opt-completed').closest('.check').hidden = !gOk;
    updateCounts();
    show('overview');
  }

  function renderLists() {
    var tbody = $('ov-lists');
    clear(tbody);
    state.model.lists.forEach(function (l, i) {
      var open = l.tasks.filter(function (t) { return !t.done; }).length;
      var cb = el('input', { type: 'checkbox', id: 'list-' + i });
      cb.checked = state.selected.has(l.id);
      cb.addEventListener('change', function () {
        if (cb.checked) state.selected.add(l.id); else state.selected.delete(l.id);
        updateCounts();
      });
      var label = el('label', { for: 'list-' + i }, [el('span', { text: l.name })]);
      if (l.isShared) label.appendChild(el('span', { class: 'tag', text: 'Shared' }));
      if (l.wellknown === 'flaggedEmails') label.appendChild(el('span', { class: 'tag', text: 'Emails you flagged in Outlook' }));
      if (state.moved.has(l.id)) label.appendChild(el('span', { class: 'tag tag-moved', text: 'Moved to Google Tasks' }));
      tbody.appendChild(el('tr', null, [
        el('td', { class: 'cell-check' }, [cb]),
        el('td', null, [label]),
        el('td', { class: 'num', text: String(open) }),
        el('td', { class: 'num', text: String(l.tasks.length - open) })
      ]));
    });
  }

  function updateCounts() {
    var ids = selectedIds(), btn = $('google-move');
    if (!gOk) return;
    var work = GoogleClient.workFor(state.model, { listIds: ids, includeCompleted: $('opt-completed').checked });
    btn.disabled = !ids.length || !work;
    btn.textContent = !ids.length ? 'Choose at least one list' : 'Move ' + plural(ids.length, 'list') + ' to Google Tasks';
    var minutes = Math.ceil(work * 0.3 / 60);
    $('google-estimate').textContent = !ids.length ? '' : !work
      ? 'There’s nothing to move in the lists you chose.'
      : 'That’s ' + plural(work, 'item') + ' (tasks and steps), about ' + (minutes <= 1 ? 'a minute' : plural(minutes, 'minute')) + '.';
  }

  function setAll(on) {
    state.selected = new Set(on ? state.model.lists.map(function (l) { return l.id; }) : []);
    renderLists();
    updateCounts();
  }

  // ------------------------------------------------------------ backups

  function backupFolder() { return 'Done With To Do backup ' + core.exportDate(state.model.exportedAt); }

  async function downloadBackup() {
    var btn = $('backup-zip'), status = $('backup-status');
    status.classList.remove('status-error');
    setBusy(btn, true, 'Preparing your backup…');
    try {
      var m = state.model, folder = backupFolder(), missing = [], extra = [], anyExpired = false;
      if (!$('opt-attachments-row').hidden && $('opt-attachments').checked) {
        var plan = core.attachmentPlan(m, folder);
        for (var i = 0; i < plan.length; i++) {
          var p = plan[i];
          status.textContent = 'Downloading attached file ' + (i + 1) + ' of ' + plan.length + '…';
          try {
            var bytes = state.demo ? core.sampleAttachmentBytes(p.attachment) : await state.ms.downloadAttachment(p.list.id, p.task.id, p.attachment.id);
            extra.push({ path: p.path, data: bytes });
          } catch (e) {
            var expired = MicrosoftClient.signInExpired(e);
            if (expired) anyExpired = true;
            missing.push({ list: p.list.name, task: p.task.title, name: p.attachment.name, reason: expired ? 'Microsoft sign-in expired' : 'Microsoft didn’t send it' });
            if (expired) {
              for (var j = i + 1; j < plan.length; j++) missing.push({ list: plan[j].list.name, task: plan[j].task.title, name: plan[j].attachment.name, reason: 'Microsoft sign-in expired' });
              break;
            }
          }
        }
      }
      status.textContent = 'Making the zip file…';
      var files = core.backupFiles(m, state.raw, folder, missing).concat(extra);
      var blob = new Blob(Zip.makeZip(files), { type: 'application/zip' });
      saveBlob(blob, folder + '.zip');
      status.textContent = 'Saved “' + folder + '.zip” (' + formatBytes(blob.size) + ') to your downloads.' +
        (missing.length ? ' ' + plural(missing.length, 'attached file') + ' couldn’t be downloaded. The README in the zip lists them' +
          (anyExpired ? '; sign in again to retry.' : '. Save them from To Do by hand.') : '');
      showThanks('backup-thanks');
    } catch (e) {
      status.textContent = 'Couldn’t make the backup: ' + e.message;
      status.classList.add('status-error');
    } finally {
      setBusy(btn, false);
    }
  }

  function openPrintable() {
    var blob = new Blob([core.toPrintableHTML(state.model)], { type: 'text/html' });
    var url = URL.createObjectURL(blob);
    var w = window.open(url, '_blank');
    if (!w) saveBlob(blob, 'To Do lists ' + core.exportDate(state.model.exportedAt) + '.html');
    setTimeout(function () { URL.revokeObjectURL(url); }, 120000);
  }

  function downloadCSV() {
    saveBlob(new Blob([core.toCSV(state.model)], { type: 'text/csv;charset=utf-8' }), 'To Do lists ' + core.exportDate(state.model.exportedAt) + '.csv');
  }

  // ------------------------------------------------------------ Google Tasks

  // Called straight from the click, so the browser allows Google's sign-in window.
  function moveToGoogle() {
    var ids = selectedIds();
    if (!state.google || !ids.length) return;
    var status = $('google-status');
    status.classList.remove('status-error');
    status.textContent = 'Waiting for Google sign-in…';
    state.google.signIn().then(function () {
      status.textContent = '';
      state.gOpts = { listIds: ids, includeCompleted: $('opt-completed').checked };
      state.gState = {};
      runTransfer();
    }, function (e) {
      status.textContent = e.message;
      status.classList.add('status-error');
    });
  }

  async function runTransfer() {
    $('g-running').hidden = false; $('g-paused').hidden = true; $('g-done').hidden = true;
    show('google');
    state.transferring = true;
    var started = Date.now(), startedAt = state.gState.done || 0;
    try {
      var result = await state.google.transfer(state.model, state.gOpts, function (p) {
        $('g-title').textContent = 'Copying “' + p.listName + '” (' + (p.listIndex + 1) + ' of ' + p.listCount + ')';
        var bar = $('g-bar'); bar.max = Math.max(1, p.total); bar.value = p.done;
        var doneNow = p.done - startedAt, secs = (Date.now() - started) / 1000;
        var left = doneNow > 8 && secs > 3 ? Math.ceil((p.total - p.done) / (doneNow / secs) / 60) : 0;
        $('g-detail').textContent = p.done + ' of ' + plural(p.total, 'item') + ' copied' +
          (left ? ', about ' + (left <= 1 ? 'a minute' : plural(left, 'minute')) + ' to go' : '') + '.';
      }, state.gState);
      state.transferring = false;
      showGoogleDone(result);
    } catch (e) {
      state.transferring = false;
      state.gError = e;
      $('g-running').hidden = true;
      $('g-paused').hidden = false;
      $('g-paused-msg').textContent = e.message;
      $('g-paused').querySelector('h1').focus({ preventScroll: true });
    }
  }

  function continueTransfer() {
    var e = state.gError;
    if (e && (e.code === 'google_auth' || e.code === 'google_scope')) {
      state.google.signIn().then(runTransfer, function (err) { $('g-paused-msg').textContent = err.message; });
    } else runTransfer();
  }

  function showGoogleDone(r) {
    $('g-running').hidden = true; $('g-paused').hidden = true; $('g-done').hidden = false;
    state.gOpts.listIds.forEach(function (id) { state.moved.add(id); state.selected.delete(id); });
    $('g-done-summary').textContent = 'Copied ' + plural(r.tasks, 'task') + (r.steps ? ' and ' + plural(r.steps, 'step') : '') +
      ' into ' + plural(r.lists, 'new list') + '.' + (state.demo ? ' These were the sample lists, so you can delete them in Google Tasks.' : '');
    var notes = $('g-done-notes');
    clear(notes);
    function note(text) { notes.appendChild(el('li', { text: text })); }
    note('Not seeing them? Google Tasks may be showing a different Google account. Click your picture at the top right of Google Tasks and switch to the account you just used.');
    if (r.repeating) note((r.repeating === 1 ? '1 task repeats' : r.repeating + ' tasks repeat') + '. Google Tasks can’t copy repeats, so each one says how it repeated in its notes. Open it in Google Tasks and set the repeat again.');
    if (r.reminders) note((r.reminders === 1 ? '1 task had a reminder' : r.reminders + ' tasks had reminders') + '. The reminder time is in the task’s notes.');
    if (r.important) note(r.important === 1 ? '1 task was marked important. That’s noted in the task.' : r.important + ' tasks were marked important. That’s noted in each one.');
    if (r.attachments) note('Attached files can’t go into Google Tasks. Their names are in the notes, and the files are in the backup.');
    r.renamed.forEach(function (x) { note('You already had a list called “' + x.from + '”, so the new one is called “' + x.to + '”.'); });
    note('Nothing in Microsoft To Do has changed. When you’re happy, you can delete your old lists there yourself.');
    var fails = $('g-failures');
    fails.hidden = !r.failures.length;
    if (r.failures.length) {
      $('g-failures-summary').textContent = plural(r.failures.length, 'item') + ' couldn’t be copied';
      var list = $('g-failures-list'); clear(list);
      r.failures.forEach(function (f) { list.appendChild(el('li', { text: f.list + ': ' + f.task + ' (' + f.error + ')' })); });
    }
    showThanks('g-thanks');
    $('g-done-title').focus({ preventScroll: true });
  }

  // ------------------------------------------------------------ thanks (stars and tips)

  function showThanks(id) {
    var box = $(id), src = sourceUrl(), tip = donateUrl();
    clear(box);
    if (!src && !tip) { box.hidden = true; return; }
    box.appendChild(el('p', { class: 'thanks-title', text: 'Did that help?' }));
    box.appendChild(el('p', { text: 'Done With To Do is free and open source. A star on GitHub helps other people find it' + (tip ? ', and a tip keeps it going.' : '.') }));
    var actions = el('p', { class: 'actions' });
    if (src) actions.appendChild(el('a', { class: 'btn btn-go', href: src, target: '_blank', rel: 'noopener', text: 'Star it on GitHub' }));
    if (tip) actions.appendChild(el('a', { class: 'btn btn-quiet', href: tip, target: '_blank', rel: 'noopener', text: 'Leave a tip' }));
    box.appendChild(actions);
    box.hidden = false;
  }

  // ------------------------------------------------------------ errors

  function showError(e, title) {
    $('err-title').textContent = title || 'Something went wrong';
    $('err-message').textContent = (e && e.message) || String(e);
    var details = e && e.details;
    $('err-details-box').hidden = !details;
    $('err-details').textContent = details ? details + '\nSite version ' + core.VERSION : '';
    show('error');
  }

  function startOver() {
    try { sessionStorage.removeItem(MicrosoftClient.STORAGE_KEY); } catch (e) { /* storage may be unavailable */ }
    location.replace(siteAddress());
  }

  // ------------------------------------------------------------ start-up

  function wire() {
    $('signin-ms').addEventListener('click', signInWithMicrosoft);
    $('try-demo').addEventListener('click', startDemo);
    $('start-over').addEventListener('click', startOver);
    $('ov-select-all').addEventListener('click', function () { setAll(true); });
    $('ov-select-none').addEventListener('click', function () { setAll(false); });
    $('opt-completed').addEventListener('change', updateCounts);
    $('backup-zip').addEventListener('click', downloadBackup);
    $('backup-print').addEventListener('click', openPrintable);
    $('backup-csv').addEventListener('click', downloadCSV);
    $('google-move').addEventListener('click', moveToGoogle);
    $('g-continue').addEventListener('click', continueTransfer);
    $('g-stop').addEventListener('click', function () { showOverview(); });
    $('g-back').addEventListener('click', function () { showOverview(); });
    $('err-retry').addEventListener('click', startOver);
    document.querySelectorAll('[data-copy]').forEach(function (b) { b.addEventListener('click', function () { copyText(b.getAttribute('data-copy'), b); }); });
    window.addEventListener('beforeunload', function (e) {
      if (state.transferring) { e.preventDefault(); e.returnValue = ''; }
    });
  }

  async function init() {
    if (!core || !Zip || !MicrosoftClient || !GoogleClient) {
      document.body.insertBefore(el('p', { class: 'notice notice-stop', text: 'Part of this site didn’t load. Refresh the page; if that doesn’t help, check that every file was uploaded.' }), document.body.firstChild);
      return;
    }
    wire();
    renderSetup();
    renderLinks();
    if (msOk) state.ms = new MicrosoftClient({ clientId: msId, authority: cfg.microsoftAuthority || 'common', redirectUri: siteAddress() });
    if (/[?&]demo\b/.test(location.search)) return startDemo();
    if (state.ms && /(^#|&)(code|error)=/.test(location.hash)) {
      show('loading');
      setLoading('Finishing sign-in', 'Checking with Microsoft.', null);
      try { await state.ms.completeSignIn(location.hash); }
      catch (e) {
        history.replaceState(null, '', location.pathname + location.search);
        return showError(e, 'Sign-in didn’t finish');
      }
      history.replaceState(null, '', location.pathname + location.search);
      return readFromMicrosoft();
    }
    renderStart();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
