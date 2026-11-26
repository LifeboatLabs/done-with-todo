/*
 * Done With To Do — moving lists into Google Tasks.
 *
 * Uses Google's own sign-in library (Google Identity Services) to get permission for
 * Google Tasks only, then creates one Google Tasks list per To Do list.
 * Steps become subtasks. Work is paced and retried, and can resume after a pause.
 * MIT licence.
 */
(function (root, factory) {
  var mod = factory(root);
  root.DWTD = root.DWTD || {};
  root.DWTD.GoogleClient = mod;
  if (typeof module === 'object' && module.exports) module.exports = mod;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (root) {
  'use strict';

  var API = 'https://tasks.googleapis.com/tasks/v1';
  var SCOPE = 'https://www.googleapis.com/auth/tasks';
  var FATAL = ['google_auth', 'google_api_disabled', 'google_scope', 'google_quota', 'network'];

  function core() { return root.DWTD.core; }
  function err(code, message, details) { return core().DWTDError(code, message, details); }
  async function safeJson(res) { try { return await res.json(); } catch (e) { return null; } }

  function reasonOf(data) {
    var e = data && data.error;
    if (!e) return '';
    var r = (e.errors && e.errors[0] && e.errors[0].reason) || '';
    var details = (e.details || []).map(function (d) { return d && d.reason; }).filter(Boolean).join(' ');
    return [r, details, e.status || '', e.message || ''].join(' ');
  }

  function apiError(status, data) {
    var reason = reasonOf(data), message = (data && data.error && data.error.message) || '';
    var x;
    if (status === 401) x = err('google_auth', 'Your Google sign-in has expired. Click Continue to sign in again and carry on where it stopped.', reason);
    else if (status === 403 && /accessNotConfigured|SERVICE_DISABLED|has not been used|is disabled/i.test(reason)) {
      x = err('google_api_disabled', 'The Google Tasks API isn’t switched on for this site yet. Turn it on in Google Cloud (setup guide, Step 3), wait a minute, then click Continue.', reason);
    } else if (status === 403 && /insufficientPermissions|ACCESS_TOKEN_SCOPE_INSUFFICIENT|insufficient/i.test(reason)) {
      x = err('google_scope', 'Google Tasks access wasn’t granted. Click Continue, sign in again and tick the box that lets Done With To Do create tasks.', reason);
    } else if (/dailyLimitExceeded|quotaExceeded/i.test(reason)) {
      x = err('google_quota', 'Google’s daily limit for this site has been reached. Click Continue tomorrow and it will carry on where it stopped.', reason);
    } else x = err('google_api', 'Google returned an error (' + status + ')' + (message ? ': ' + message : '.'), reason);
    x.status = status;
    return x;
  }

  function authError(type, description) {
    var t = String(type || '');
    if (t === 'popup_closed') return err('google_signin', 'The Google window was closed before you finished signing in.');
    if (t === 'popup_failed_to_open') return err('google_signin', 'Your browser blocked the Google sign-in window. Allow pop-ups for this site, then try again.');
    if (t === 'access_denied') return err('google_signin', 'Google sign-in was cancelled. If you saw "Access blocked", this Google account isn’t a test user yet (setup guide, Step 3.5).');
    return err('google_signin', 'Google sign-in didn’t work' + (t ? ' (' + t + ')' : '') + (description ? ': ' + description : '.'));
  }

  function GoogleClient(opts) {
    this.clientId = opts.clientId;
    this.fetch = opts.fetch || (root.fetch && root.fetch.bind(root));
    this.sleep = opts.sleep || function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
    this.now = opts.now || function () { return Date.now(); };
    this.minGap = opts.minGap == null ? 110 : opts.minGap; // milliseconds between requests
    this.token = opts.token || null;
    this._last = null; // when the previous request started
  }

  /** Load Google's sign-in library early, so a click can open the sign-in window straight away. */
  GoogleClient.prototype.prepare = function () {
    var self = this;
    if (this._ready) return this._ready;
    this._ready = new Promise(function (resolve, reject) {
      if (root.google && root.google.accounts && root.google.accounts.oauth2) return resolve();
      var s = root.document.createElement('script');
      s.src = 'https://accounts.google.com/gsi/client';
      s.async = true;
      s.onload = function () { resolve(); };
      s.onerror = function () {
        self._ready = null;
        reject(err('google_load', 'Couldn’t load Google sign-in. A browser extension or network filter may be blocking accounts.google.com.'));
      };
      root.document.head.appendChild(s);
    }).then(function () {
      self._tokenClient = root.google.accounts.oauth2.initTokenClient({
        client_id: self.clientId,
        scope: SCOPE,
        callback: function (resp) {
          var p = self._pending; self._pending = null;
          if (!p) return;
          if (resp && resp.error) return p.reject(authError(resp.error, resp.error_description));
          if (!root.google.accounts.oauth2.hasGrantedAllScopes(resp, SCOPE)) {
            return p.reject(err('google_scope', 'Google Tasks access wasn’t granted. Try again and tick the box that lets Done With To Do create tasks.'));
          }
          self.token = resp.access_token;
          p.resolve();
        },
        error_callback: function (e) {
          var p = self._pending; self._pending = null;
          if (p) p.reject(authError(e && e.type));
        }
      });
    });
    return this._ready;
  };

  /** Opens Google's sign-in window. Call it directly from a click, after prepare() has finished. */
  GoogleClient.prototype.signIn = function () {
    var self = this;
    if (!this._tokenClient) return Promise.reject(err('google_load', 'Google sign-in is still loading. Wait a moment and click again.'));
    return new Promise(function (resolve, reject) {
      self._pending = { resolve: resolve, reject: reject };
      self._tokenClient.requestAccessToken(self.token ? { prompt: '' } : {});
    });
  };

  /** One call to the Google Tasks API, paced and retried when Google asks us to slow down. */
  GoogleClient.prototype.call = async function (method, path, body) {
    for (var attempt = 0; ; attempt++) {
      var wait = this._last === null ? 0 : this._last + this.minGap - this.now();
      if (wait > 0) await this.sleep(wait);
      this._last = this.now();
      var res;
      try {
        res = await this.fetch(API + path, {
          method: method,
          headers: { Authorization: 'Bearer ' + this.token, 'Content-Type': 'application/json' },
          body: body ? JSON.stringify(body) : undefined
        });
      } catch (e) {
        if (attempt < 4) { await this.sleep(1000 * Math.pow(2, attempt)); continue; }
        throw err('network', 'Couldn’t reach Google. Check your connection, then click Continue to carry on where it stopped.');
      }
      if (res.ok) return res.status === 204 ? null : safeJson(res);
      var data = await safeJson(res), reason = reasonOf(data);
      var slowDown = res.status === 429 || res.status >= 500 ||
        (res.status === 403 && /rateLimitExceeded|userRateLimitExceeded|RATE_LIMIT|RESOURCE_EXHAUSTED/i.test(reason) && !/dailyLimit/i.test(reason));
      if (slowDown && attempt < 6) { await this.sleep(Math.min(32000, 1000 * Math.pow(2, attempt)) + Math.floor(Math.random() * 400)); continue; }
      throw apiError(res.status, data);
    }
  };

  GoogleClient.prototype.existingListTitles = async function () {
    var titles = [], pageToken = '', pages = 0;
    do {
      var page = await this.call('GET', '/users/@me/lists?maxResults=100' + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : ''));
      ((page && page.items) || []).forEach(function (l) { titles.push(String(l.title || '').trim().toLowerCase()); });
      pageToken = page && page.nextPageToken;
    } while (pageToken && ++pages < 100);
    return titles;
  };

  function isFatal(e) { return e && FATAL.indexOf(e.code) !== -1; }

  /** How many items (tasks + steps) a transfer will create, for the progress bar. */
  GoogleClient.workFor = function (model, opts) {
    var n = 0;
    model.lists.forEach(function (l) {
      if (opts.listIds.indexOf(l.id) === -1) return;
      l.tasks.forEach(function (t) { if (opts.includeCompleted || !t.done) n += 1 + t.steps.length; });
    });
    return n;
  };

  /**
   * Copy the chosen lists into Google Tasks.
   *   opts:       { listIds: [...], includeCompleted: boolean }
   *   onProgress: ({ done, total, listName, listIndex, listCount })
   *   state:      pass the same object again to resume after a pause; nothing is created twice.
   * Throws on problems that need the person (sign-in expired, API off, offline); per-task
   * problems are recorded in state.failures and the transfer carries on.
   */
  GoogleClient.prototype.transfer = async function (model, opts, onProgress, state) {
    var C = core(), report = onProgress || function () {};
    state = state || {};
    state.lists = state.lists || {};
    state.failures = state.failures || [];
    state.done = state.done || 0;
    state.total = GoogleClient.workFor(model, opts);
    var lists = model.lists.filter(function (l) { return opts.listIds.indexOf(l.id) !== -1; });
    if (!state.existing) state.existing = await this.existingListTitles();

    for (var li = 0; li < lists.length; li++) {
      var l = lists[li];
      var ls = state.lists[l.id] || (state.lists[l.id] = { tasks: {}, steps: {}, lastTop: null });
      var tick = function () { report({ done: state.done, total: state.total, listName: l.name, listIndex: li, listCount: lists.length }); };
      tick();
      if (!ls.googleListId) {
        var title = C.uniqueListTitle(l.name, state.existing);
        var created = await this.call('POST', '/users/@me/lists', { title: title });
        ls.googleListId = created.id;
        ls.title = title;
        state.existing.push(title.toLowerCase());
      }
      var listPath = '/lists/' + encodeURIComponent(ls.googleListId) + '/tasks';
      var tasks = l.tasks.filter(function (t) { return opts.includeCompleted || !t.done; });

      for (var ti = 0; ti < tasks.length; ti++) {
        var t = tasks[ti];
        if (!ls.tasks[t.id]) {
          try {
            var q = ls.lastTop ? '?previous=' + encodeURIComponent(ls.lastTop) : '';
            var g = await this.call('POST', listPath + q, C.googleTaskBody(t));
            ls.tasks[t.id] = g.id;
            ls.lastTop = g.id;
          } catch (e) {
            if (isFatal(e)) throw e;
            ls.tasks[t.id] = 'failed';
            state.failures.push({ list: l.name, task: t.title || '(untitled)', error: e.message });
          }
          state.done++;
          tick();
        }
        var parent = ls.tasks[t.id];
        var st = ls.steps[t.id] || (ls.steps[t.id] = { n: 0, last: null });
        if (parent === 'failed') {
          if (st.n < t.steps.length) { state.done += t.steps.length - st.n; st.n = t.steps.length; tick(); }
          continue;
        }
        for (var si = st.n; si < t.steps.length; si++) {
          var s = t.steps[si];
          try {
            var sq = '?parent=' + encodeURIComponent(parent) + (st.last ? '&previous=' + encodeURIComponent(st.last) : '');
            var gs = await this.call('POST', listPath + sq, C.googleStepBody(s));
            st.last = gs.id;
          } catch (e) {
            if (isFatal(e)) throw e;
            state.failures.push({ list: l.name, task: (t.title || '(untitled)') + ' → ' + (s.title || '(untitled)'), error: e.message });
          }
          st.n = si + 1;
          state.done++;
          tick();
        }
      }
      ls.complete = true;
    }
    state.finished = true;
    return GoogleClient.summary(model, opts, state);
  };

  /** What happened, for the "done" screen. */
  GoogleClient.summary = function (model, opts, state) {
    var r = { lists: 0, tasks: 0, steps: 0, repeating: 0, reminders: 0, important: 0, attachments: 0, renamed: [], failures: state.failures || [] };
    model.lists.forEach(function (l) {
      var ls = state.lists && state.lists[l.id];
      if (!ls || !ls.googleListId) return;
      r.lists++;
      if (ls.title && ls.title !== l.name) r.renamed.push({ from: l.name, to: ls.title });
      l.tasks.forEach(function (t) {
        var id = ls.tasks[t.id];
        if (!id || id === 'failed') return;
        r.tasks++;
        if (t.recurrence) r.repeating++;
        if (t.reminder) r.reminders++;
        if (t.important) r.important++;
        r.attachments += t.attachments.length;
        var st = ls.steps[t.id];
        if (st) r.steps += st.n;
      });
    });
    r.steps = Math.max(0, r.steps - r.failures.filter(function (f) { return f.task.indexOf(' → ') !== -1; }).length);
    return r;
  };

  GoogleClient.SCOPE = SCOPE;
  return GoogleClient;
});
