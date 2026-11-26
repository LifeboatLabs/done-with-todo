/*
 * Done With To Do — core logic.
 *
 * Turns Microsoft To Do data (as Microsoft Graph returns it) into:
 *   - a tidy model the page can show,
 *   - backup files (spreadsheet, printable page, Markdown, calendar files, JSON),
 *   - request bodies for Google Tasks.
 *
 * Nothing in this file touches the network or the page, so it can be tested on its own.
 * MIT licence. https://github.com/ (see README)
 */
(function (root, factory) {
  var core = factory(root);
  root.DWTD = root.DWTD || {};
  root.DWTD.core = core;
  if (typeof module === 'object' && module.exports) module.exports = core;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (root) {
  'use strict';

  var core = {};
  // Shown in technical details, so a report says which version of the site the person had.
  // Bump it together with the ?v= numbers in index.html and privacy.html.
  core.VERSION = '3';

  // ---------------------------------------------------------------- errors

  /** An error with a short code and a message written for people, not programmers. */
  function DWTDError(code, message, details) {
    var e = new Error(message);
    e.name = 'DWTDError';
    e.code = code;
    if (details) e.details = details;
    return e;
  }
  core.DWTDError = DWTDError;

  // ---------------------------------------------------------------- small helpers

  function lower(s) { return String(s == null ? '' : s).toLowerCase(); }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  /** Cut a string to at most n UTF-16 units without splitting an emoji, adding an ellipsis. */
  function clip(s, n) {
    s = String(s == null ? '' : s);
    if (s.length <= n) return s;
    var cut = s.slice(0, Math.max(0, n - 1));
    if (/[\ud800-\udbff]$/.test(cut)) cut = cut.slice(0, -1);
    return cut + '\u2026';
  }
  core.clip = clip;

  var DT_RE = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?/;
  function parseParts(s) {
    var m = DT_RE.exec(String(s || ''));
    if (!m) return null;
    return { y: +m[1], mo: +m[2], d: +m[3], h: +(m[4] || 0), mi: +(m[5] || 0), s: +(m[6] || 0) };
  }

  function isUtcZone(tz) {
    if (!tz) return true;
    var t = lower(tz).trim();
    return t === 'utc' || t === 'etc/utc' || t === 'gmt' || t === 'etc/gmt' || t === 'z' ||
      t === 'coordinated universal time' || t === 'tzone://microsoft/utc';
  }

  /**
   * Graph dates come either as plain ISO strings ending in "Z" (createdDateTime) or as
   * { dateTime, timeZone } objects (dueDateTime, reminderDateTime...). Returns a Date or null.
   */
  function instant(dt) {
    if (!dt) return null;
    var str = typeof dt === 'string' ? dt : dt.dateTime;
    var zone = typeof dt === 'string' ? 'UTC' : dt.timeZone;
    var p = parseParts(str);
    if (!p) return null;
    if (typeof dt === 'string' && !/z$/i.test(str)) {
      var off = /T.*([+-])(\d{2}):?(\d{2})$/.exec(str);
      if (off) {
        var sign = off[1] === '-' ? -1 : 1;
        return new Date(Date.UTC(p.y, p.mo - 1, p.d, p.h, p.mi, p.s) - sign * (+off[2] * 60 + +off[3]) * 60000);
      }
    }
    if (isUtcZone(zone)) return new Date(Date.UTC(p.y, p.mo - 1, p.d, p.h, p.mi, p.s));
    // A named time zone: treat the wall-clock time as this computer's local time.
    return new Date(p.y, p.mo - 1, p.d, p.h, p.mi, p.s);
  }

  function localYMD(d) { return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
  function localYMDHM(d) { return localYMD(d) + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes()); }

  /**
   * A whole-day date (due date, start date) as "YYYY-MM-DD".
   * To Do stores these either as midnight UTC or as local midnight converted to UTC.
   * Midnight UTC keeps its day; anything else is converted to this computer's local day.
   */
  core.graphDate = function (dt) {
    if (!dt) return null;
    var str = typeof dt === 'string' ? dt : dt.dateTime;
    var zone = typeof dt === 'string' ? 'UTC' : dt.timeZone;
    var p = parseParts(str);
    if (!p || p.y < 1900) return null;
    if (!isUtcZone(zone) || (p.h === 0 && p.mi === 0 && p.s === 0)) return p.y + '-' + pad2(p.mo) + '-' + pad2(p.d);
    return localYMD(instant(dt));
  };

  /** A moment in time (reminder) as local "YYYY-MM-DD HH:MM". */
  core.graphLocalDateTime = function (dt) {
    var i = instant(dt);
    return i && !isNaN(i) && i.getFullYear() >= 1900 ? localYMDHM(i) : null;
  };

  /** A moment in time as an ISO string in UTC. */
  core.graphISO = function (dt) {
    var i = instant(dt);
    return i && !isNaN(i) && i.getUTCFullYear() >= 1900 ? i.toISOString() : null;
  };

  function isoToLocal(iso) {
    if (!iso) return '';
    var d = new Date(iso);
    return isNaN(d) ? '' : localYMDHM(d);
  }

  // ---------------------------------------------------------------- notes (HTML -> text)

  var ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };
  core.htmlToText = function (html) {
    if (!html) return '';
    var s = String(html);
    s = s.replace(/<(script|style|head|title)\b[\s\S]*?<\/\1\s*>/gi, '');
    s = s.replace(/<br\s*\/?>/gi, '\n');
    s = s.replace(/<li\b[^>]*>/gi, '\n\u2022 ');
    s = s.replace(/<\/(p|div|h[1-6]|tr|table|ul|ol|blockquote)\s*>/gi, '\n');
    s = s.replace(/<[^>]*>/g, '');
    s = s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, function (m, e) {
      if (e.charAt(0) === '#') {
        var code = e.charAt(1).toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        return code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : m;
      }
      var k = e.toLowerCase();
      return Object.prototype.hasOwnProperty.call(ENTITIES, k) ? ENTITIES[k] : m;
    });
    s = s.replace(/\r\n?/g, '\n').replace(/[ \t\u00a0]+\n/g, '\n').replace(/\n{3,}/g, '\n\n');
    return s.trim();
  };

  // ---------------------------------------------------------------- repeats

  var DAY_ORDER = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
  var DAY_NAME = { monday: 'Monday', tuesday: 'Tuesday', wednesday: 'Wednesday', thursday: 'Thursday', friday: 'Friday', saturday: 'Saturday', sunday: 'Sunday' };
  var DAY_ICS = { monday: 'MO', tuesday: 'TU', wednesday: 'WE', thursday: 'TH', friday: 'FR', saturday: 'SA', sunday: 'SU' };
  var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  var INDEX_WORD = { first: 'first', second: 'second', third: 'third', fourth: 'fourth', last: 'last' };
  var INDEX_NUM = { first: 1, second: 2, third: 3, fourth: 4, last: -1 };

  function ordinal(n) {
    var s = ['th', 'st', 'nd', 'rd'], v = n % 100;
    return n + (s[(v - 20) % 10] || s[v] || s[0]);
  }
  function sortDays(days) {
    return (days || []).map(lower).filter(function (d) { return DAY_ICS[d]; })
      .filter(function (d, i, a) { return a.indexOf(d) === i; })
      .sort(function (a, b) { return DAY_ORDER.indexOf(a) - DAY_ORDER.indexOf(b); });
  }
  function every(n, one, many) { return n === 1 ? 'every ' + one : 'every ' + n + ' ' + many; }
  function monthName(m) { return MONTHS[Math.min(12, Math.max(1, +m || 1)) - 1]; }
  function validRangeDate(s) { return /^\d{4}-\d{2}-\d{2}/.test(s || '') && +String(s).slice(0, 4) >= 1900; }

  /** "every week on Monday, Thursday" — readable, and close to the Obsidian Tasks syntax. */
  core.describeRecurrence = function (rec) {
    if (!rec || !rec.pattern || !rec.pattern.type) return null;
    var p = rec.pattern, n = Math.max(1, Math.floor(+p.interval || 1)), t = lower(p.type);
    var days = sortDays(p.daysOfWeek), names = days.map(function (d) { return DAY_NAME[d]; });
    var idx = INDEX_WORD[lower(p.index)] || 'first';
    var text;
    if (t === 'daily') text = every(n, 'day', 'days');
    else if (t === 'weekly') {
      if (n === 1 && days.join() === 'monday,tuesday,wednesday,thursday,friday') text = 'every weekday';
      else text = every(n, 'week', 'weeks') + (names.length ? ' on ' + names.join(', ') : '');
    } else if (t === 'absolutemonthly') text = every(n, 'month', 'months') + (p.dayOfMonth ? ' on the ' + ordinal(+p.dayOfMonth) : '');
    else if (t === 'relativemonthly') text = every(n, 'month', 'months') + ' on the ' + idx + ' ' + (names.join(', ') || 'day');
    else if (t === 'absoluteyearly') text = every(n, 'year', 'years') + ' on ' + monthName(p.month) + ' ' + ordinal(+p.dayOfMonth || 1);
    else if (t === 'relativeyearly') text = every(n, 'year', 'years') + ' on the ' + idx + ' ' + (names.join(', ') || 'day') + ' of ' + monthName(p.month);
    else return 'repeats (' + p.type + ')';
    var r = rec.range || {}, rt = lower(r.type);
    if (rt === 'enddate' && validRangeDate(r.endDate)) text += ' until ' + String(r.endDate).slice(0, 10);
    else if (rt === 'numbered' && +r.numberOfOccurrences > 0) text += ', ' + (+r.numberOfOccurrences) + ' times';
    return text;
  };

  function relativeByDay(days, idx) {
    if (!days.length) days = ['MO'];
    return days.length === 1 ? ['BYDAY=' + idx + days[0]] : ['BYDAY=' + days.join(','), 'BYSETPOS=' + idx];
  }

  /** The same repeat as an iCalendar RRULE value, e.g. "FREQ=WEEKLY;BYDAY=MO,TH". */
  core.recurrenceToRRule = function (rec) {
    if (!rec || !rec.pattern || !rec.pattern.type) return null;
    var p = rec.pattern, n = Math.max(1, Math.floor(+p.interval || 1)), t = lower(p.type);
    var days = sortDays(p.daysOfWeek).map(function (d) { return DAY_ICS[d]; });
    var idx = INDEX_NUM[lower(p.index)] || 1;
    var parts;
    if (t === 'daily') parts = ['FREQ=DAILY'];
    else if (t === 'weekly') { parts = ['FREQ=WEEKLY']; if (days.length) parts.push('BYDAY=' + days.join(',')); }
    else if (t === 'absolutemonthly') { parts = ['FREQ=MONTHLY']; if (+p.dayOfMonth) parts.push('BYMONTHDAY=' + (+p.dayOfMonth)); }
    else if (t === 'relativemonthly') parts = ['FREQ=MONTHLY'].concat(relativeByDay(days, idx));
    else if (t === 'absoluteyearly') parts = ['FREQ=YEARLY', 'BYMONTH=' + (+p.month || 1), 'BYMONTHDAY=' + (+p.dayOfMonth || 1)];
    else if (t === 'relativeyearly') parts = ['FREQ=YEARLY', 'BYMONTH=' + (+p.month || 1)].concat(relativeByDay(days, idx));
    else return null;
    if (n > 1) parts.splice(1, 0, 'INTERVAL=' + n);
    var r = rec.range || {}, rt = lower(r.type);
    if (rt === 'enddate' && validRangeDate(r.endDate)) parts.push('UNTIL=' + String(r.endDate).slice(0, 10).replace(/-/g, ''));
    else if (rt === 'numbered' && +r.numberOfOccurrences > 0) parts.push('COUNT=' + (+r.numberOfOccurrences));
    return parts.join(';');
  };

  // ---------------------------------------------------------------- the model

  var STATUS_LABEL = {
    notstarted: 'Not started', inprogress: 'In progress', completed: 'Completed',
    waitingonothers: 'Waiting on someone else', deferred: 'Deferred'
  };

  function normalizeTask(t) {
    var status = lower(t.status) || 'notstarted';
    var body = t.body || {};
    var notes = lower(body.contentType) === 'html'
      ? core.htmlToText(body.content)
      : String(body.content || '').replace(/\r\n?/g, '\n').trim();
    var rec = t.recurrence && t.recurrence.pattern && t.recurrence.pattern.type ? t.recurrence : null;
    var reminderOn = !!t.isReminderOn && !!t.reminderDateTime;
    return {
      id: t.id,
      title: String(t.title || '').replace(/\s+/g, ' ').trim(),
      status: status,
      statusLabel: STATUS_LABEL[status] || String(t.status || ''),
      done: status === 'completed',
      importance: lower(t.importance) || 'normal',
      important: lower(t.importance) === 'high',
      notes: notes,
      due: core.graphDate(t.dueDateTime),
      start: core.graphDate(t.startDateTime),
      reminder: reminderOn ? core.graphLocalDateTime(t.reminderDateTime) : null,
      reminderISO: reminderOn ? core.graphISO(t.reminderDateTime) : null,
      completedAt: core.graphISO(t.completedDateTime),
      completedDate: core.graphDate(t.completedDateTime),
      createdAt: core.graphISO(t.createdDateTime),
      modifiedAt: core.graphISO(t.lastModifiedDateTime),
      categories: (t.categories || []).map(String),
      recurrence: rec ? { text: core.describeRecurrence(rec), rrule: core.recurrenceToRRule(rec), raw: rec } : null,
      steps: (t.checklistItems || []).map(function (c) {
        return {
          id: c.id, title: String(c.displayName || '').replace(/\s+/g, ' ').trim(), done: !!c.isChecked,
          createdAt: core.graphISO(c.createdDateTime), checkedAt: core.graphISO(c.checkedDateTime)
        };
      }),
      links: (t.linkedResources || []).filter(function (r) { return r && r.webUrl; }).map(function (r) {
        return { url: String(r.webUrl), app: r.applicationName || '', name: r.displayName || '' };
      }),
      attachments: (t.attachments || []).map(function (a) {
        return { id: a.id, name: a.name || 'attachment', size: +a.size || 0, contentType: a.contentType || 'application/octet-stream' };
      }),
      hasAttachments: !!t.hasAttachments || !!(t.attachments && t.attachments.length),
      // The task has files, but Microsoft wouldn't list them, so they can't be in the backup.
      attachmentsUnavailable: !!t.attachmentsUnavailable
    };
  }

  var FILES_UNAVAILABLE = 'files Microsoft wouldn’t share, so save them from To Do';

  /** What to show for a task's attached files: their names, plus a warning when Microsoft held them back. */
  function attachmentNames(t) {
    var names = t.attachments.map(function (a) { return a.name; });
    if (t.attachmentsUnavailable) names.push('(' + FILES_UNAVAILABLE + ')');
    return names;
  }

  /**
   * input: { account, rawLists, rawTasksByList, exportedAt? }
   *   rawLists       — the "value" array from GET /me/todo/lists
   *   rawTasksByList — { listId: [todoTask (+ checklistItems, linkedResources, attachments)] }
   */
  core.normalize = function (input) {
    input = input || {};
    var tasksByList = input.rawTasksByList || {};
    return {
      exportedAt: input.exportedAt || new Date().toISOString(),
      account: input.account || null,
      lists: (input.rawLists || []).map(function (l) {
        var name = String(l.displayName || '').trim() || 'Untitled list';
        return {
          id: l.id, name: name,
          isShared: !!l.isShared, isOwner: l.isOwner !== false,
          wellknown: l.wellknownListName || 'none',
          tasks: (tasksByList[l.id] || []).map(normalizeTask)
        };
      })
    };
  };

  /** Counts for the page. selected: optional Set of list ids to count. */
  core.summarize = function (model, selected) {
    var s = { lists: 0, tasks: 0, open: 0, done: 0, steps: 0, repeating: 0, reminders: 0, important: 0, attachments: 0, attachmentBytes: 0, attachmentsUnavailable: 0 };
    (model.lists || []).forEach(function (l) {
      if (selected && !selected.has(l.id)) return;
      s.lists++;
      l.tasks.forEach(function (t) {
        s.tasks++;
        if (t.done) s.done++; else s.open++;
        s.steps += t.steps.length;
        if (t.recurrence) s.repeating++;
        if (t.reminder) s.reminders++;
        if (t.important) s.important++;
        s.attachments += t.attachments.length;
        t.attachments.forEach(function (a) { s.attachmentBytes += a.size; });
        if (t.attachmentsUnavailable) s.attachmentsUnavailable++;
      });
    });
    return s;
  };

  // ---------------------------------------------------------------- file names

  /** A name that works as a file name on Windows, Mac and Linux, unique within `used`. */
  core.safeFileName = function (name, used, ext) {
    ext = ext || '';
    var base = String(name == null ? '' : name)
      .replace(/[\u0000-\u001f\u007f<>:"\/\\|?*]+/g, '-')
      .replace(/\s+/g, ' ').trim()
      .replace(/^[.\s]+|[.\s]+$/g, '');
    if (!base) base = 'Untitled';
    if (/^(con|prn|aux|nul|com\d|lpt\d)$/i.test(base)) base = '_' + base;
    var chars = Array.from(base);
    if (chars.length > 80) base = chars.slice(0, 80).join('').trim();
    var candidate = base + ext, i = 2;
    while (used && used.has(candidate.toLowerCase())) candidate = base + ' (' + (i++) + ')' + ext;
    if (used) used.add(candidate.toLowerCase());
    return candidate;
  };

  function ymd(iso) { var d = iso ? new Date(iso) : new Date(); return isNaN(d) ? localYMD(new Date()) : localYMD(d); }
  core.exportDate = ymd;

  // ---------------------------------------------------------------- spreadsheet (CSV)

  function csvCell(v) {
    var s = v == null ? '' : String(v);
    // Stop spreadsheet apps from treating text like "=SUM(...)" as a formula.
    if (/^[=+@\t\r]/.test(s) || (/^-\S/.test(s) && /[=(|!]/.test(s))) s = "'" + s;
    if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
    return s;
  }
  core.csvCell = csvCell;

  core.CSV_HEADER = ['List', 'Task', 'Done', 'Due date', 'Reminder', 'Repeats', 'Important', 'Status', 'Notes', 'Steps',
    'Start date', 'Categories', 'Links', 'Attachments', 'Created', 'Completed on'];

  core.toCSV = function (model) {
    var rows = [core.CSV_HEADER];
    model.lists.forEach(function (l) {
      l.tasks.forEach(function (t) {
        rows.push([
          l.name, t.title, t.done ? 'Yes' : 'No', t.due || '', t.reminder || '',
          t.recurrence ? t.recurrence.text : '', t.important ? 'Yes' : '', t.statusLabel, t.notes,
          t.steps.map(function (s) { return (s.done ? '[x] ' : '[ ] ') + s.title; }).join('\n'),
          t.start || '', t.categories.join(', '),
          t.links.map(function (x) { return x.url; }).join('\n'),
          attachmentNames(t).join('\n'),
          isoToLocal(t.createdAt), t.done ? (t.completedDate || '') : ''
        ]);
      });
    });
    // The byte-order mark makes Excel read accents and emoji correctly.
    return '\ufeff' + rows.map(function (r) { return r.map(csvCell).join(','); }).join('\r\n') + '\r\n';
  };

  // ---------------------------------------------------------------- printable page (HTML)

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  core.escapeHTML = esc;

  function prettyDay(ymdStr) {
    var p = parseParts(ymdStr);
    if (!p) return ymdStr || '';
    try {
      return new Intl.DateTimeFormat(undefined, { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC' })
        .format(new Date(Date.UTC(p.y, p.mo - 1, p.d)));
    } catch (e) { return ymdStr; }
  }
  function prettyMoment(local) {
    var p = parseParts(local);
    if (!p) return local || '';
    try {
      return new Intl.DateTimeFormat(undefined, { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })
        .format(new Date(p.y, p.mo - 1, p.d, p.h, p.mi));
    } catch (e) { return local; }
  }
  core.prettyDay = prettyDay;
  core.prettyMoment = prettyMoment;

  function plural(n, one, many) { return n + ' ' + (n === 1 ? one : (many || one + 's')); }
  core.plural = plural;

  function printableTask(t) {
    var meta = [];
    if (t.due) meta.push('Due ' + prettyDay(t.due));
    if (t.start) meta.push('Starts ' + prettyDay(t.start));
    if (t.recurrence) meta.push('Repeats ' + t.recurrence.text);
    if (t.reminder) meta.push('Reminder ' + prettyMoment(t.reminder));
    if (t.important) meta.push('Important');
    if (!t.done && t.status !== 'notstarted') meta.push(t.statusLabel);
    if (t.done && t.completedDate) meta.push('Completed ' + prettyDay(t.completedDate));
    if (t.categories.length) meta.push(t.categories.join(', '));
    var h = '<li class="task' + (t.done ? ' done' : '') + '"><span class="box">' + (t.done ? '\u2611' : '\u2610') + '</span>' +
      '<span class="title">' + esc(t.title || '(untitled)') + '</span>';
    if (meta.length) h += '<span class="meta">' + esc(meta.join(' \u00b7 ')) + '</span>';
    if (t.steps.length) {
      h += '<ul class="steps">' + t.steps.map(function (s) {
        return '<li' + (s.done ? ' class="done"' : '') + '>' + (s.done ? '\u2611 ' : '\u2610 ') + esc(s.title || '(untitled)') + '</li>';
      }).join('') + '</ul>';
    }
    if (t.notes) h += '<div class="notes">' + esc(t.notes) + '</div>';
    if (t.links.length) {
      h += '<span class="meta">' + t.links.map(function (x) {
        return /^https?:\/\//i.test(x.url) ? '<a href="' + esc(x.url) + '">' + esc(x.name || x.url) + '</a>' : esc(x.url);
      }).join(' \u00b7 ') + '</span>';
    }
    var files = attachmentNames(t);
    if (files.length) h += '<span class="meta">Attachments: ' + esc(files.join(', ')) + '</span>';
    return h + '</li>';
  }

  core.toPrintableHTML = function (model) {
    var s = core.summarize(model);
    var when = prettyDay(ymd(model.exportedAt));
    var out = [];
    out.push('<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">',
      '<meta name="viewport" content="width=device-width, initial-scale=1">',
      '<title>My Microsoft To Do lists \u2014 backup from ' + esc(when) + '</title><style>',
      'body{font:15px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:#1f1d1a;max-width:760px;margin:32px auto;padding:0 20px}',
      'h1{font-size:26px;margin:0 0 4px}.sub{color:#6b665e;margin:0 0 16px}',
      '.hint{background:#f3efe7;padding:10px 12px;border-radius:8px;font-size:13px}',
      'nav ol{columns:2;padding-left:20px;font-size:14px}nav a{color:inherit}',
      'section{margin-top:30px}h2{font-size:19px;border-bottom:2px solid #1f1d1a;padding-bottom:4px;margin-bottom:6px}',
      'h2 small{font-weight:400;color:#6b665e;font-size:13px;margin-left:8px}',
      'h3{font-size:12px;text-transform:uppercase;letter-spacing:.06em;color:#6b665e;margin:18px 0 4px}',
      'ul.tasks{list-style:none;padding:0;margin:0}li.task{padding:7px 0;border-bottom:1px solid #e7e1d6;break-inside:avoid}',
      '.box{display:inline-block;width:1.3em}.done>.title{text-decoration:line-through;color:#6b665e}',
      '.meta{display:block;font-size:12.5px;color:#6b665e;margin-left:1.3em}.meta a{color:#6b665e}',
      '.notes{white-space:pre-wrap;font-size:13.5px;margin:4px 0 0 1.3em;color:#3b3731}',
      'ul.steps{list-style:none;margin:4px 0 0 1.3em;padding:0;font-size:13.5px}ul.steps li.done{color:#6b665e}',
      'footer{margin:40px 0 20px;color:#6b665e;font-size:12px}',
      '@media print{body{margin:0;max-width:none}.hint,nav{display:none}a{text-decoration:none}}',
      '</style></head><body>');
    out.push('<h1>My Microsoft To Do lists</h1>');
    out.push('<p class="sub">Backup made ' + esc(when) + ' \u00b7 ' + esc(plural(s.lists, 'list')) + ' \u00b7 ' +
      esc(plural(s.tasks, 'task')) + ' (' + s.open + ' open, ' + s.done + ' done)</p>');
    out.push('<p class="hint">To print this or save it as a PDF, press Ctrl+P (or \u2318P on a Mac).</p>');
    if (model.lists.length > 1) {
      out.push('<nav><ol>' + model.lists.map(function (l, i) {
        return '<li><a href="#list-' + (i + 1) + '">' + esc(l.name) + '</a> (' + l.tasks.length + ')</li>';
      }).join('') + '</ol></nav>');
    }
    model.lists.forEach(function (l, i) {
      var open = l.tasks.filter(function (t) { return !t.done; });
      var done = l.tasks.filter(function (t) { return t.done; });
      out.push('<section id="list-' + (i + 1) + '"><h2>' + esc(l.name) + '<small>' + open.length + ' open \u00b7 ' + done.length + ' done' +
        (l.isShared ? ' \u00b7 shared' : '') + '</small></h2>');
      if (!l.tasks.length) out.push('<p class="sub">No tasks.</p>');
      if (open.length) out.push('<ul class="tasks">' + open.map(printableTask).join('') + '</ul>');
      if (done.length) out.push('<h3>Completed</h3><ul class="tasks">' + done.map(printableTask).join('') + '</ul>');
      out.push('</section>');
    });
    out.push('<footer>Made with Done With To Do \u2014 a free, open-source tool. Not affiliated with Microsoft.</footer></body></html>');
    return out.join('\n');
  };

  // ---------------------------------------------------------------- Markdown (Obsidian Tasks style)

  function mdInline(s) { return String(s || '').replace(/\s*\n\s*/g, ' ').replace(/([\\`\[\]])/g, '\\$1'); }

  function markdownTask(t) {
    var head = ['- [' + (t.done ? 'x' : ' ') + '] ' + (mdInline(t.title) || '(untitled)')];
    if (t.important) head.push('\u23eb');
    if (t.recurrence) head.push('\ud83d\udd01 ' + t.recurrence.text);
    if (t.start) head.push('\ud83d\udeeb ' + t.start);
    if (t.due) head.push('\ud83d\udcc5 ' + t.due);
    if (t.done && t.completedDate) head.push('\u2705 ' + t.completedDate);
    var lines = [head.join(' ')];
    t.steps.forEach(function (s) { lines.push('    - [' + (s.done ? 'x' : ' ') + '] ' + (mdInline(s.title) || '(untitled)')); });
    if (t.reminder) lines.push('    > Reminder: ' + t.reminder);
    if (t.notes) t.notes.split('\n').forEach(function (line) { lines.push('    > ' + line); });
    t.links.forEach(function (x) { lines.push('    > Link: ' + x.url); });
    attachmentNames(t).forEach(function (name) { lines.push('    > Attachment: ' + name); });
    return lines.join('\n');
  }

  /** One Markdown file per list: [{ name, content }]. */
  core.toMarkdownFiles = function (model) {
    var used = new Set();
    return model.lists.map(function (l) {
      var lines = ['# ' + String(l.name).replace(/\s*\n\s*/g, ' '), '', '_Exported from Microsoft To Do on ' + ymd(model.exportedAt) + '_', ''];
      l.tasks.forEach(function (t) { lines.push(markdownTask(t)); });
      if (!l.tasks.length) lines.push('_No tasks._');
      return { name: core.safeFileName(l.name, used, '.md'), content: lines.join('\n') + '\n' };
    });
  };

  // ---------------------------------------------------------------- calendar files (iCalendar VTODO)

  function hash(str) {
    // Two FNV-1a passes with different seeds -> 16 hex characters. Only used to make stable IDs.
    function fnv(seed) {
      var h = seed >>> 0;
      for (var i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; }
      return ('0000000' + h.toString(16)).slice(-8);
    }
    return fnv(2166136261) + fnv(374761393);
  }
  core.hash = hash;

  function icsText(s) {
    return String(s == null ? '' : s).replace(/\r\n?/g, '\n')
      .replace(/[\u0000-\u0008\u000b-\u001f\u007f]/g, '')
      .replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\n/g, '\\n');
  }
  function icsUTC(d) {
    d = d instanceof Date ? d : new Date(d);
    return d.getUTCFullYear() + pad2(d.getUTCMonth() + 1) + pad2(d.getUTCDate()) + 'T' +
      pad2(d.getUTCHours()) + pad2(d.getUTCMinutes()) + pad2(d.getUTCSeconds()) + 'Z';
  }
  function utf8Len(cp) { return cp < 0x80 ? 1 : cp < 0x800 ? 2 : cp < 0x10000 ? 3 : 4; }
  /** Lines longer than 75 bytes are folded onto continuation lines, as RFC 5545 requires. */
  function fold(line) {
    var out = '', bytes = 0;
    for (var ch of line) {
      var b = utf8Len(ch.codePointAt(0));
      if (bytes + b > 75) { out += '\r\n '; bytes = 1; }
      out += ch; bytes += b;
    }
    return out;
  }
  core.icsFold = fold;

  /** One .ics file per list: [{ name, content }]. Steps become sub-tasks (RELATED-TO). */
  core.toICSFiles = function (model) {
    var used = new Set(), stamp = icsUTC(new Date(model.exportedAt || Date.now()));
    return model.lists.map(function (l) {
      var lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Done With To Do//EN', 'CALSCALE:GREGORIAN',
        'X-WR-CALNAME:' + icsText(l.name)];
      l.tasks.forEach(function (t) {
        var uid = 'dwtd-' + hash(l.id + '|' + t.id) + '@done-with-todo';
        lines.push('BEGIN:VTODO', 'UID:' + uid, 'DTSTAMP:' + stamp);
        if (t.createdAt) lines.push('CREATED:' + icsUTC(t.createdAt));
        if (t.modifiedAt) lines.push('LAST-MODIFIED:' + icsUTC(t.modifiedAt));
        lines.push('SUMMARY:' + icsText(t.title || '(untitled)'));
        var extra = [];
        attachmentNames(t).forEach(function (name) { extra.push('Attachment: ' + name); });
        var desc = [t.notes].concat(extra).filter(Boolean).join('\n');
        if (desc) lines.push('DESCRIPTION:' + icsText(desc));
        var start = t.start || (t.recurrence && t.recurrence.rrule ? t.due : null);
        if (start) lines.push('DTSTART;VALUE=DATE:' + start.replace(/-/g, ''));
        if (t.due) lines.push('DUE;VALUE=DATE:' + t.due.replace(/-/g, ''));
        lines.push('STATUS:' + (t.done ? 'COMPLETED' : t.status === 'inprogress' ? 'IN-PROCESS' : 'NEEDS-ACTION'));
        if (t.done) {
          lines.push('PERCENT-COMPLETE:100');
          if (t.completedAt) lines.push('COMPLETED:' + icsUTC(t.completedAt));
        }
        if (t.importance === 'high') lines.push('PRIORITY:1');
        else if (t.importance === 'low') lines.push('PRIORITY:9');
        if (t.recurrence && t.recurrence.rrule && start) lines.push('RRULE:' + t.recurrence.rrule);
        if (t.categories.length) lines.push('CATEGORIES:' + t.categories.map(icsText).join(','));
        var link = t.links.filter(function (x) { return /^https?:\/\//i.test(x.url); })[0];
        if (link) lines.push('URL:' + link.url.replace(/[\u0000-\u001f\u007f\s]/g, ''));
        if (t.reminderISO) {
          lines.push('BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:' + icsText(t.title || 'Reminder'),
            'TRIGGER;VALUE=DATE-TIME:' + icsUTC(t.reminderISO), 'END:VALARM');
        }
        lines.push('END:VTODO');
        t.steps.forEach(function (s, i) {
          lines.push('BEGIN:VTODO', 'UID:dwtd-' + hash(l.id + '|' + t.id + '|' + (s.id || i)) + '@done-with-todo',
            'DTSTAMP:' + stamp, 'SUMMARY:' + icsText(s.title || '(untitled)'),
            'STATUS:' + (s.done ? 'COMPLETED' : 'NEEDS-ACTION'));
          if (s.done) {
            lines.push('PERCENT-COMPLETE:100');
            if (s.checkedAt) lines.push('COMPLETED:' + icsUTC(s.checkedAt));
          }
          lines.push('RELATED-TO;RELTYPE=PARENT:' + uid, 'END:VTODO');
        });
      });
      lines.push('END:VCALENDAR');
      return { name: core.safeFileName(l.name, used, '.ics'), content: lines.map(fold).join('\r\n') + '\r\n' };
    });
  };

  // ---------------------------------------------------------------- JSON (everything)

  core.toJSON = function (model, raw) {
    return JSON.stringify({
      format: 'done-with-todo-backup', version: 1,
      exportedAt: model.exportedAt, account: model.account,
      summary: core.summarize(model), lists: model.lists,
      raw: raw ? { lists: raw.rawLists || [], tasksByList: raw.rawTasksByList || {} } : null
    }, null, 2);
  };

  // ---------------------------------------------------------------- Google Tasks

  core.GOOGLE_TITLE_MAX = 1024;
  core.GOOGLE_NOTES_MAX = 8192;

  /** Google Tasks has no repeats, reminders, importance or files, so they're written into the notes. */
  core.googleNotes = function (t) {
    var extra = [];
    if (t.recurrence) extra.push('Repeats ' + t.recurrence.text + ' (set the repeat again in Google Tasks)');
    if (t.reminder) extra.push('Reminder was set for ' + t.reminder);
    if (t.important) extra.push('Marked as important');
    if (!t.done && t.status !== 'notstarted' && t.statusLabel) extra.push('Status: ' + t.statusLabel);
    if (t.start) extra.push('Start date: ' + t.start);
    if (t.categories.length) extra.push('Categories: ' + t.categories.join(', '));
    t.links.forEach(function (x) { extra.push('Link: ' + x.url); });
    if (t.attachments.length) {
      extra.push('Attachments not copied (they are in your backup): ' + t.attachments.map(function (a) { return a.name; }).join(', '));
    }
    if (t.attachmentsUnavailable) extra.push('This task has attached files that Microsoft wouldn’t share. Save them from To Do.');
    var footer = extra.length ? '\u2014 From Microsoft To Do \u2014\n' + extra.join('\n') : '';
    var notes = t.notes || '';
    var sep = notes && footer ? '\n\n' : '';
    var max = core.GOOGLE_NOTES_MAX;
    if ((notes + sep + footer).length > max) {
      var tail = '\n\u2026(shortened \u2014 the full note is in your backup)';
      notes = clip(notes, Math.max(0, max - footer.length - sep.length - tail.length - 1)) + tail;
      var all = notes + sep + footer;
      return all.length > max ? clip(all, max) : all;
    }
    return notes + sep + footer;
  };

  /** The request body for one task in the Google Tasks API. */
  core.googleTaskBody = function (t) {
    var body = { title: clip(t.title || '(untitled)', core.GOOGLE_TITLE_MAX), status: t.done ? 'completed' : 'needsAction' };
    var notes = core.googleNotes(t);
    if (notes) body.notes = notes;
    if (t.due) body.due = t.due + 'T00:00:00.000Z';
    if (t.done && t.completedAt) body.completed = t.completedAt;
    return body;
  };

  core.googleStepBody = function (s) {
    return { title: clip(s.title || '(untitled)', core.GOOGLE_TITLE_MAX), status: s.done ? 'completed' : 'needsAction' };
  };

  /** A list name that doesn't clash with the person's existing Google Tasks lists. */
  core.uniqueListTitle = function (name, existingLower) {
    var base = clip(String(name || '').trim() || 'Untitled list', 900), title = base, n = 1;
    while (existingLower.indexOf(title.toLowerCase()) !== -1) {
      title = base + (n === 1 ? ' (from To Do)' : ' (from To Do ' + n + ')');
      n++;
    }
    return title;
  };

  // ---------------------------------------------------------------- backup bundle

  core.backupReadme = function (model, missing) {
    var s = core.summarize(model);
    var lines = [
      'Your Microsoft To Do backup',
      '===========================',
      '',
      'Made on ' + ymd(model.exportedAt) + ' with Done With To Do.',
      s.lists + ' lists, ' + s.tasks + ' tasks (' + s.open + ' open, ' + s.done + ' done), ' + s.steps + ' steps.',
      '',
      'What is in this folder',
      '----------------------',
      '- To Do backup.html  Open in any browser to read or print your lists (Ctrl+P / Cmd+P to save as PDF).',
      '- To Do backup.csv   Opens in Excel, Numbers or Google Sheets. One row per task.',
      '- Markdown (Obsidian)/  One file per list, in the Obsidian Tasks format.',
      '- Calendar files (.ics)/  One file per list. Many task apps can import these',
      '  (Thunderbird, Nextcloud Tasks, Tasks.org, Fastmail and others).',
      '- To Do backup.json  Everything, including the original data from Microsoft,',
      '  so it can be imported into other apps later.',
      '- Attachments/  Files attached to your tasks, in a folder per list and task.',
      ''
    ];
    if (missing && missing.length) {
      lines.push('Attachments that could not be downloaded', '----------------------------------------');
      missing.forEach(function (m) { lines.push('- ' + m.list + ' / ' + m.task + ' / ' + m.name + (m.reason ? ' (' + m.reason + ')' : '')); });
      lines.push('');
    }
    var heldBack = [];
    model.lists.forEach(function (l) {
      l.tasks.forEach(function (t) { if (t.attachmentsUnavailable) heldBack.push('- ' + l.name + ' / ' + (t.title || 'Untitled task')); });
    });
    if (heldBack.length) {
      lines.push('Tasks with attached files Microsoft wouldn\'t share', '-------------------------------------------------',
        'Open these tasks in Microsoft To Do and save their files by hand before it closes:');
      lines.push.apply(lines, heldBack);
      lines.push('');
    }
    lines.push('Done With To Do runs entirely in your browser. Your tasks were never sent to its makers.', '');
    return lines.join('\r\n');
  };

  /**
   * All text files for the backup zip. Attachments are added separately (they need downloading).
   * Returns [{ path, data }], paths inside a single top folder.
   */
  core.backupFiles = function (model, raw, folder, missing) {
    folder = folder || ('Done With To Do backup ' + ymd(model.exportedAt));
    var files = [
      { path: folder + '/README.txt', data: core.backupReadme(model, missing) },
      { path: folder + '/To Do backup.html', data: core.toPrintableHTML(model) },
      { path: folder + '/To Do backup.csv', data: core.toCSV(model) },
      { path: folder + '/To Do backup.json', data: core.toJSON(model, raw) }
    ];
    core.toMarkdownFiles(model).forEach(function (f) { files.push({ path: folder + '/Markdown (Obsidian)/' + f.name, data: f.content }); });
    core.toICSFiles(model).forEach(function (f) { files.push({ path: folder + '/Calendar files (.ics)/' + f.name, data: f.content }); });
    return files;
  };

  /** Where each attachment goes in the zip: [{ list, task, attachment, path }]. */
  core.attachmentPlan = function (model, folder) {
    folder = folder || ('Done With To Do backup ' + ymd(model.exportedAt));
    var plan = [], listNames = new Set();
    model.lists.forEach(function (l) {
      var listDir = core.safeFileName(l.name, listNames), taskNames = new Set();
      l.tasks.forEach(function (t) {
        if (!t.attachments.length) return;
        var taskDir = core.safeFileName(t.title || 'Untitled task', taskNames), fileNames = new Set();
        t.attachments.forEach(function (a) {
          var dot = a.name.lastIndexOf('.');
          var ext = dot > 0 && a.name.length - dot <= 10 ? a.name.slice(dot) : '';
          var base = ext ? a.name.slice(0, dot) : a.name;
          var file = core.safeFileName(base, fileNames, ext.replace(/[^\w.]/g, ''));
          plan.push({ list: l, task: t, attachment: a, path: folder + '/Attachments/' + listDir + '/' + taskDir + '/' + file });
        });
      });
    });
    return plan;
  };

  // ---------------------------------------------------------------- sample data (for "Try it with sample data")

  /** Data shaped exactly like Microsoft Graph's, so the demo runs the same code as the real thing. */
  core.sampleData = function (now) {
    now = now || new Date();
    function dayUTC(offset) {
      var d = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate() + offset));
      return { dateTime: d.toISOString().slice(0, 10) + 'T00:00:00.0000000', timeZone: 'UTC' };
    }
    function momentUTC(offset, hour) {
      var d = new Date(now.getFullYear(), now.getMonth(), now.getDate() + offset, hour, 0, 0);
      return { dateTime: d.toISOString().slice(0, 19) + '.0000000', timeZone: 'UTC' };
    }
    function iso(offset) { return new Date(now.getTime() + offset * 86400000).toISOString(); }
    var id = 0;
    function task(fields) {
      id++;
      var t = {
        id: 'demo-task-' + id, title: fields.title, status: fields.done ? 'completed' : (fields.status || 'notStarted'),
        importance: fields.important ? 'high' : 'normal', isReminderOn: !!fields.reminder,
        createdDateTime: iso(-30 - id), lastModifiedDateTime: iso(-id),
        body: { content: fields.notes || '', contentType: 'text' },
        categories: fields.categories || [], hasAttachments: !!(fields.attachments && fields.attachments.length),
        checklistItems: (fields.steps || []).map(function (s, i) {
          var done = s.charAt(0) === '+';
          return { id: 'demo-step-' + id + '-' + i, displayName: done ? s.slice(1) : s, isChecked: done, createdDateTime: iso(-20) };
        }),
        linkedResources: fields.link ? [{ id: 'demo-link-' + id, webUrl: fields.link, applicationName: 'Outlook', displayName: fields.linkName || 'Email' }] : [],
        attachments: (fields.attachments || []).map(function (name, i) {
          return { id: 'demo-att-' + id + '-' + i, name: name, size: 1200 + i * 300, contentType: name.slice(-4) === '.pdf' ? 'application/pdf' : 'text/plain' };
        })
      };
      if (fields.due != null) t.dueDateTime = dayUTC(fields.due);
      if (fields.reminder) t.reminderDateTime = momentUTC(fields.reminder[0], fields.reminder[1]);
      if (fields.done) t.completedDateTime = dayUTC(fields.doneAgo || -3);
      if (fields.repeat) t.recurrence = fields.repeat;
      return t;
    }
    var weekly = function (days) { return { pattern: { type: 'weekly', interval: 1, daysOfWeek: days, firstDayOfWeek: 'sunday' }, range: { type: 'noEnd', startDate: dayUTC(0).dateTime.slice(0, 10) } }; };
    var lists = [
      { id: 'demo-list-tasks', displayName: 'Tasks', isOwner: true, isShared: false, wellknownListName: 'defaultList' },
      { id: 'demo-list-groceries', displayName: 'Groceries \ud83d\uded2', isOwner: true, isShared: true, wellknownListName: 'none' },
      { id: 'demo-list-home', displayName: 'Home & garden', isOwner: true, isShared: false, wellknownListName: 'none' },
      { id: 'demo-list-party', displayName: "Mum's 60th", isOwner: true, isShared: false, wellknownListName: 'none' },
      { id: 'demo-list-flagged', displayName: 'Flagged email', isOwner: true, isShared: false, wellknownListName: 'flaggedEmails' }
    ];
    var tasks = {
      'demo-list-tasks': [
        task({ title: 'Renew passport', due: 12, important: true, reminder: [10, 9], notes: 'Need: old passport, 2 photos, card for the fee.\nCheck the "fast track" option.', steps: ['+Get photos taken', 'Fill in the form', 'Book appointment'] }),
        task({ title: 'Call the dentist', due: 2, reminder: [2, 10] }),
        task({ title: 'Pay credit card', due: 5, repeat: { pattern: { type: 'absoluteMonthly', interval: 1, dayOfMonth: 15 }, range: { type: 'noEnd' } }, categories: ['Bills'] }),
        task({ title: 'Back up laptop', repeat: weekly(['sunday']), due: 4 }),
        task({ title: 'Return library books', done: true, doneAgo: -6 })
      ],
      'demo-list-groceries': [
        task({ title: 'Milk', repeat: weekly(['saturday']) }),
        task({ title: 'Eggs (free range)' }),
        task({ title: 'Coffee beans, the dark roast', notes: 'The one from the market, not the supermarket' }),
        task({ title: 'Bread', done: true, doneAgo: -1 }),
        task({ title: '=2+2 is not a formula', notes: 'A test of spreadsheet safety; this text should show as typed.' })
      ],
      'demo-list-home': [
        task({ title: 'Fix the dripping tap', status: 'inProgress', steps: ['+Buy washers', 'Turn off water', 'Replace washer'], attachments: ['tap-model.pdf'] }),
        task({ title: 'Mow the lawn', repeat: { pattern: { type: 'weekly', interval: 2, daysOfWeek: ['saturday'] }, range: { type: 'noEnd' } }, due: 3 }),
        task({ title: 'Clean gutters', due: 20, notes: 'Ladder is in the shed; borrow the long one from next door if needed.' }),
        task({ title: 'Take bins out', repeat: { pattern: { type: 'weekly', interval: 1, daysOfWeek: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'] }, range: { type: 'noEnd' } } }),
        task({ title: 'Boiler service', repeat: { pattern: { type: 'absoluteYearly', interval: 1, month: 10, dayOfMonth: 1 }, range: { type: 'noEnd' } }, due: 40, important: true })
      ],
      'demo-list-party': [
        task({ title: 'Book the restaurant', due: 7, important: true, steps: ['+Shortlist 3 places', '+Check accessibility', 'Book for 14 people'] }),
        task({ title: 'Send invitations', due: 9, attachments: ['guest-list.txt'] }),
        task({ title: 'Order the cake', due: 25, notes: 'Lemon drizzle; "Happy 60th, Mum!" on top' }),
        task({ title: 'Make a photo slideshow', steps: ['Ask everyone for photos', 'Pick music'] })
      ],
      'demo-list-flagged': [
        task({ title: 'Re: Quote for the kitchen', link: 'https://outlook.live.com/mail/0/inbox', linkName: 'Quote for the kitchen' })
      ]
    };
    return {
      account: { name: 'Sample account', email: 'sample@example.com' },
      rawLists: lists, rawTasksByList: tasks, exportedAt: now.toISOString(), isDemo: true
    };
  };

  /** Bytes for a sample attachment (the demo can't download real ones). */
  core.sampleAttachmentBytes = function (attachment) {
    var text = 'This is a sample attachment called "' + attachment.name + '".\r\nIn a real backup, your actual file is saved here.\r\n';
    return new TextEncoder().encode(text);
  };

  return core;
});
