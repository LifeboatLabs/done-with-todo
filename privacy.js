/* Done With To Do: fills in the owner's name and contact on the privacy page from config.js. */
(function () {
  'use strict';
  var cfg = window.DWTD_CONFIG || {};
  function each(sel, fn) { Array.prototype.forEach.call(document.querySelectorAll(sel), fn); }
  var name = String(cfg.ownerName || '').trim();
  var email = String(cfg.contactEmail || '').trim();
  if (name) each('.owner-name', function (n) { n.textContent = name; });
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    each('.owner-contact', function (n) { n.hidden = false; });
    each('.owner-email', function (a) { a.textContent = email; a.href = 'mailto:' + email; });
    each('.contact-sentence', function (n) { n.textContent = 'Email ' + email + ', or open an issue on the project’s GitHub page.'; });
  }
  var src = /^https:\/\/github\.com\//.test(cfg.githubUrl || '') ? cfg.githubUrl : '';
  if (!src && /\.github\.io$/i.test(location.hostname)) {
    var repo = location.pathname.split('/').filter(Boolean)[0];
    if (repo && /\.html?$/i.test(repo)) repo = '';
    src = 'https://github.com/' + location.hostname.split('.')[0] + '/' + (repo || location.hostname);
  }
  if (src) {
    each('.source-sentence', function (n) { n.hidden = false; });
    each('.source-link', function (a) { a.href = src + '/commits/main/privacy.html'; });
  }
})();
