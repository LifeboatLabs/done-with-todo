/*
 * Done With To Do — Microsoft sign-in and reading To Do through Microsoft Graph.
 *
 * Sign-in uses OAuth 2.0 with PKCE, the standard for apps that run only in the browser.
 * The only permissions requested are Tasks.Read (read your lists and tasks) and
 * User.Read (show your name). Tokens live in memory and vanish when the tab closes.
 * MIT licence.
 */
(function (root, factory) {
  var mod = factory(root);
  root.DWTD = root.DWTD || {};
  root.DWTD.MicrosoftClient = mod;
  if (typeof module === 'object' && module.exports) module.exports = mod;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (root) {
  'use strict';

  var GRAPH = 'https://graph.microsoft.com/v1.0';
  var SCOPES = 'openid profile User.Read Tasks.Read';
  var KEY = 'dwtd.microsoft.signin';

  function core() { return root.DWTD.core; }
  function err(code, message, details) { return core().DWTDError(code, message, details); }

  function b64url(bytes) {
    var s = '';
    for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }
  function randomString(byteCount) {
    var b = new Uint8Array(byteCount);
    root.crypto.getRandomValues(b);
    return b64url(b);
  }
  async function pkceChallenge(verifier) {
    var digest = await root.crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
    return b64url(new Uint8Array(digest));
  }
  function enc(id) { return encodeURIComponent(id); }
  function base64ToBytes(b64) {
    var bin = atob(String(b64).replace(/-/g, '+').replace(/_/g, '/')), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  async function safeJson(res) { try { return await res.json(); } catch (e) { return null; } }

  async function eachLimit(items, limit, fn) {
    var i = 0;
    async function worker() { while (i < items.length) { var item = items[i++]; await fn(item); } }
    var workers = [];
    for (var w = 0; w < Math.min(limit, items.length); w++) workers.push(worker());
    await Promise.all(workers);
  }

  /** Turns Microsoft's sign-in error codes into instructions a person can act on. */
  function friendlyAuthError(code, description, redirectUri) {
    var d = String(description || '');
    var aad = (/AADSTS(\d+)/.exec(d) || [])[1];
    var where = 'Microsoft Entra → App registrations → Done With To Do';
    var map = {
      '50011': 'Microsoft doesn’t recognise this page’s address. In ' + where + ' → Authentication, the "Single-page application" redirect URI must be exactly: ' + redirectUri,
      '9002326': 'In ' + where + ' → Authentication, this address is listed under "Web". Remove it there and add it again under "Single-page application".',
      '7000218': 'In ' + where + ' → Authentication, this address is listed under "Web". Remove it there and add it again under "Single-page application".',
      '700016': 'Microsoft can’t find an app with this client ID. Check microsoftClientId in config.js, and that the app allows personal Microsoft accounts.',
      '50020': 'This app is set to accept accounts from one organisation only. In ' + where + ' → Authentication, set "Supported account types" to include personal Microsoft accounts.',
      '50194': 'This app is set to accept accounts from one organisation only. In ' + where + ' → Authentication, set "Supported account types" to include personal Microsoft accounts.',
      '90072': 'This account can’t use the app yet. In ' + where + ' → Authentication, set "Supported account types" to include personal Microsoft accounts.',
      '65001': 'Permission wasn’t given, so nothing was read. Try again and choose Accept.',
      '65004': 'You chose not to give permission, so nothing was read.',
      '90094': 'Your organisation needs an administrator to approve this app. Try a personal Microsoft account, or ask your IT team.',
      '90008': 'Your organisation needs an administrator to approve this app. Try a personal Microsoft account, or ask your IT team.',
      '54005': 'That sign-in link was already used. Please sign in again.',
      '501481': 'The sign-in check didn’t match. Please sign in again.'
    };
    if (aad && map[aad]) return err('ms_signin', map[aad], d);
    if (code === 'access_denied') return err('ms_signin', 'Sign-in was cancelled, so nothing was read.', d);
    var firstLine = d.split(/\r?\n/)[0].replace(/^AADSTS\d+:\s*/, '');
    return err('ms_signin', 'Microsoft sign-in didn’t work' + (firstLine ? ': ' + firstLine : '.'), d);
  }

  function graphError(status, body) {
    var e = (body && body.error) || {};
    var detail = (e.code ? e.code + ': ' : '') + (e.message || '');
    var x;
    if (status === 401) x = err('ms_auth', 'Your Microsoft sign-in has expired (they last about an hour). Please sign in again.', detail);
    else if (status === 403) x = err('ms_forbidden', 'Microsoft didn’t allow access to your tasks. If this is a work or school account, your organisation may block it. Try a personal account.', detail);
    else if (status === 404) x = err('ms_notfound', 'Microsoft couldn’t find To Do data for this account. If you use To Do with another Microsoft account, start over and pick that one.', detail);
    else x = err('graph', 'Microsoft returned an error (' + status + ')' + (e.message ? ': ' + e.message : '.'), detail);
    x.status = status;
    return x;
  }

  function MicrosoftClient(opts) {
    this.clientId = opts.clientId;
    this.authority = opts.authority || 'common';
    this.redirectUri = opts.redirectUri;
    this.token = null;
    this.tokenExpires = 0;
    this.fetch = opts.fetch || root.fetch.bind(root);
    this.sleep = opts.sleep || function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
    this.storage = opts.storage || root.sessionStorage;
  }

  MicrosoftClient.prototype.loginBase = function () {
    return 'https://login.microsoftonline.com/' + encodeURIComponent(this.authority) + '/oauth2/v2.0/';
  };

  /** Step 1: build the Microsoft sign-in address (and remember the one-time secrets for step 2). */
  MicrosoftClient.prototype.signInUrl = async function () {
    var verifier = randomString(48), state = randomString(16);
    this.storage.setItem(KEY, JSON.stringify({ v: verifier, s: state, t: Date.now() }));
    var q = new URLSearchParams({
      client_id: this.clientId, response_type: 'code', redirect_uri: this.redirectUri,
      response_mode: 'fragment', scope: SCOPES, state: state,
      code_challenge: await pkceChallenge(verifier), code_challenge_method: 'S256', prompt: 'select_account'
    });
    return this.loginBase() + 'authorize?' + q.toString();
  };

  /**
   * Step 2: after Microsoft sends the person back here, swap the one-time code for an access token.
   * Returns null when this page load isn't a return from Microsoft.
   */
  MicrosoftClient.prototype.completeSignIn = async function (hash) {
    var params = new URLSearchParams(String(hash || '').replace(/^#/, ''));
    if (!params.has('code') && !params.has('error')) return null;
    var saved = null;
    try { saved = JSON.parse(this.storage.getItem(KEY) || 'null'); } catch (e) { saved = null; }
    this.storage.removeItem(KEY);
    if (params.has('error')) throw friendlyAuthError(params.get('error'), params.get('error_description'), this.redirectUri);
    if (!saved || saved.s !== params.get('state')) {
      throw err('ms_state', 'The sign-in was interrupted or started in another tab. Please sign in again.');
    }
    var body = new URLSearchParams({
      client_id: this.clientId, grant_type: 'authorization_code', code: params.get('code'),
      redirect_uri: this.redirectUri, code_verifier: saved.v, scope: SCOPES
    });
    var res, data;
    try {
      res = await this.fetch(this.loginBase() + 'token', {
        method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: body.toString()
      });
    } catch (e) {
      throw err('ms_token_blocked', 'Couldn’t finish signing in with Microsoft. Check your internet connection. If it keeps happening, check that this site’s address is listed under "Single-page application" (not "Web") in Microsoft Entra → App registrations → Done With To Do → Authentication.');
    }
    data = await safeJson(res);
    if (!res.ok || !data || !data.access_token) {
      throw friendlyAuthError(data && data.error, data && data.error_description, this.redirectUri);
    }
    this.token = data.access_token;
    this.tokenExpires = Date.now() + (+data.expires_in || 3600) * 1000;
    return true;
  };

  MicrosoftClient.prototype.tokenValid = function () { return !!this.token && Date.now() < this.tokenExpires - 60000; };

  /** GET from Microsoft Graph, waiting and retrying when Microsoft asks us to slow down. */
  MicrosoftClient.prototype.get = async function (url, as) {
    var full = /^https:\/\//.test(url) ? url : GRAPH + url;
    if (full.indexOf('https://graph.microsoft.com/') !== 0) throw err('graph', 'Unexpected address from Microsoft.');
    for (var attempt = 0; ; attempt++) {
      var res;
      try {
        res = await this.fetch(full, { headers: { Authorization: 'Bearer ' + this.token, Accept: as === 'bytes' ? '*/*' : 'application/json' } });
      } catch (e) {
        if (attempt < 3) { await this.sleep(1000 * Math.pow(2, attempt)); continue; }
        throw err('network', 'Couldn’t reach Microsoft. Check your internet connection and try again.');
      }
      if ((res.status === 429 || res.status === 502 || res.status === 503 || res.status === 504) && attempt < 6) {
        var after = parseFloat(res.headers && res.headers.get && res.headers.get('Retry-After'));
        await this.sleep(isFinite(after) ? Math.min(after, 60) * 1000 : 1000 * Math.pow(2, attempt));
        continue;
      }
      if (!res.ok) throw graphError(res.status, await safeJson(res));
      if (as === 'bytes') return new Uint8Array(await res.arrayBuffer());
      return res.json();
    }
  };

  /** GET every page of a Graph collection. */
  MicrosoftClient.prototype.getAll = async function (url) {
    var out = [], next = url, pages = 0;
    while (next) {
      var page = await this.get(next);
      out = out.concat((page && page.value) || []);
      next = page && page['@odata.nextLink'];
      if (++pages > 5000) break;
    }
    return out;
  };

  /** All tasks in one list, with their steps (and links when Microsoft allows it). */
  MicrosoftClient.prototype.getTasks = async function (listId) {
    var self = this, base = '/me/todo/lists/' + enc(listId) + '/tasks';
    var tries = ['?$expand=checklistItems,linkedResources&$top=100', '?$expand=checklistItems&$top=100', '?$top=100', ''];
    var tasks = null, used = '';
    for (var k = 0; k < tries.length; k++) {
      try { tasks = await this.getAll(base + tries[k]); used = tries[k]; break; }
      catch (e) { if (e.status === 400 && k < tries.length - 1) continue; throw e; }
    }
    var gotSteps = used.indexOf('checklistItems') !== -1 &&
      (tasks.length === 0 || tasks.some(function (t) { return Array.isArray(t.checklistItems); }));
    if (!gotSteps) {
      await eachLimit(tasks, 4, async function (t) {
        try { t.checklistItems = await self.getAll(base + '/' + enc(t.id) + '/checklistItems'); }
        catch (e) { if (e.code === 'ms_auth') throw e; t.checklistItems = []; }
      });
    }
    return tasks;
  };

  /**
   * Read everything: your name, every list, every task, every step, and the names of attachments.
   * onProgress({ phase, index, count, listName, tasks })
   */
  MicrosoftClient.prototype.readEverything = async function (onProgress) {
    var self = this, report = onProgress || function () {};
    var account = null;
    try {
      var me = await this.get('/me?$select=displayName,mail,userPrincipalName');
      account = { name: me.displayName || '', email: me.mail || me.userPrincipalName || '' };
    } catch (e) { if (e.code === 'ms_auth') throw e; }
    report({ phase: 'lists' });
    var lists = await this.getAll('/me/todo/lists');
    var tasksByList = {}, total = 0;
    for (var i = 0; i < lists.length; i++) {
      var l = lists[i];
      report({ phase: 'tasks', index: i, count: lists.length, listName: l.displayName || 'Untitled list', tasks: total });
      var tasks = await this.getTasks(l.id);
      var withFiles = tasks.filter(function (t) { return t.hasAttachments; });
      await eachLimit(withFiles, 4, async function (t) {
        var path = '/me/todo/lists/' + enc(l.id) + '/tasks/' + enc(t.id) + '/attachments';
        try { t.attachments = await self.getAll(path + '?$select=id,name,size,contentType'); }
        catch (e) {
          if (e.code === 'ms_auth') throw e;
          try { t.attachments = (await self.getAll(path)).map(function (a) { return { id: a.id, name: a.name, size: a.size, contentType: a.contentType }; }); }
          catch (e2) { if (e2.code === 'ms_auth') throw e2; t.attachments = []; }
        }
      });
      tasksByList[l.id] = tasks;
      total += tasks.length;
    }
    report({ phase: 'done', count: lists.length, tasks: total });
    return { account: account, rawLists: lists, rawTasksByList: tasksByList, exportedAt: new Date().toISOString() };
  };

  /** Download one attachment's bytes. */
  MicrosoftClient.prototype.downloadAttachment = async function (listId, taskId, attachmentId) {
    var path = '/me/todo/lists/' + enc(listId) + '/tasks/' + enc(taskId) + '/attachments/' + enc(attachmentId);
    try { return await this.get(path + '/$value', 'bytes'); }
    catch (e) {
      if (e.code === 'ms_auth') throw e;
      var a = await this.get(path);
      if (a && a.contentBytes) return base64ToBytes(a.contentBytes);
      throw e;
    }
  };

  MicrosoftClient.SCOPES = SCOPES;
  MicrosoftClient.STORAGE_KEY = KEY;
  MicrosoftClient._test = { b64url: b64url, pkceChallenge: pkceChallenge, friendlyAuthError: friendlyAuthError };
  return MicrosoftClient;
});
