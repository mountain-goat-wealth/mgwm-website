// ── Shared form submission ──
// Every form POSTs to the Google Apps Script web app (scripts/mgwm-forms.gs),
// which saves the lead, sends alerts and adds the subscriber to MailerLite.
// No API keys belong in this file or anywhere in the browser.
(function () {
  // Apps Script web app URL (Deploy → Manage deployments → Web app URL, ends in /exec)
  var FORMS_ENDPOINT = 'https://script.google.com/macros/s/AKfycbx0GmHI6NY0Jweud17aOjfqVaNlJzkCn-mSgDdxtfG_0ca7IDWM-5chRuwTWkr2u4cURw/exec';
  var TIMEOUT_MS = 10000;

  // Resolves true only when the backend replies {ok:true}; never rejects.
  window.submitForm = function (formName, data) {
    var payload = Object.assign({ page_url: window.location.href }, data, { form: formName });
    var controller = typeof AbortController === 'function' ? new AbortController() : null;
    var timer = setTimeout(function () { if (controller) controller.abort(); }, TIMEOUT_MS);

    var request = fetch(FORMS_ENDPOINT, {
      method: 'POST',
      // text/plain keeps this a "simple" request, so there's no CORS preflight
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      signal: controller ? controller.signal : undefined
    })
      .then(function (res) { return res.ok ? res.json() : null; })
      .then(function (body) { return !!(body && body.ok === true); });

    var timeout = new Promise(function (resolve) {
      setTimeout(function () { resolve(false); }, TIMEOUT_MS);
    });

    return Promise.race([request, timeout])
      .catch(function () { return false; })
      .then(function (ok) { clearTimeout(timer); return ok; });
  };
})();
