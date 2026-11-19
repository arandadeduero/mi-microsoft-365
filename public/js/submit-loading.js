/**
 * Prevent double POST submissions across all forms and fetch buttons:
 * disables the submit button(s) and shows a Bootstrap spinner while the
 * request is in flight, and restores them if the browser returns via bfcache.
 */
(function () {
  const SUBMIT_SELECTOR = 'button[type="submit"], button:not([type]), input[type="submit"]';

  function addSpinner(button) {
    if (button.tagName !== 'BUTTON') return;
    if (button.querySelector('[data-submit-spinner]')) return;
    const spinner = document.createElement('span');
    spinner.className = 'spinner-border spinner-border-sm me-2';
    spinner.setAttribute('role', 'status');
    spinner.setAttribute('aria-hidden', 'true');
    spinner.dataset.submitSpinner = 'true';
    button.prepend(spinner);
  }

  function removeSpinner(button) {
    if (button.tagName !== 'BUTTON') return;
    const spinner = button.querySelector('[data-submit-spinner]');
    if (spinner) spinner.remove();
  }

  function setButtonLoading(button, loading) {
    if (!button) return;
    if (loading) {
      button.disabled = true;
      button.setAttribute('aria-busy', 'true');
      addSpinner(button);
    } else {
      button.disabled = false;
      button.removeAttribute('aria-busy');
      removeSpinner(button);
    }
  }

  // Shared helper for fetch-based POST actions (profile.js, groups.js)
  window.setButtonLoading = setButtonLoading;

  // Listen on window so form-level validation (reset-password.js) and
  // document-level confirmations (confirm.js) run first and can preventDefault().
  window.addEventListener('submit', (event) => {
    const form = event.target;
    if (!(form instanceof HTMLFormElement)) return;
    if (event.defaultPrevented) return;
    if ((form.getAttribute('method') || 'GET').toUpperCase() !== 'POST') return;

    if (form.dataset.submitting === 'true') {
      event.preventDefault();
      return;
    }
    form.dataset.submitting = 'true';

    const submitter = event.submitter instanceof HTMLElement ? event.submitter : form.querySelector(SUBMIT_SELECTOR);

    // Keep submitter name/value in the POST payload when disabling it
    if (submitter && submitter.name) {
      const hidden = document.createElement('input');
      hidden.type = 'hidden';
      hidden.name = submitter.name;
      hidden.value = submitter.value || '';
      hidden.dataset.submitterMirror = 'true';
      form.appendChild(hidden);
    }

    const buttons = Array.from(form.querySelectorAll(SUBMIT_SELECTOR));
    if (submitter && !buttons.includes(submitter)) buttons.push(submitter);

    buttons.forEach((button) => {
      if (button === submitter) {
        setButtonLoading(button, true);
      } else {
        button.disabled = true;
      }
    });
  });

  // Reset forms if the user navigates back from browser bfcache
  window.addEventListener('pageshow', () => {
    document.querySelectorAll('form[data-submitting="true"]').forEach((form) => {
      delete form.dataset.submitting;
      form.querySelectorAll('input[data-submitter-mirror="true"]').forEach((el) => el.remove());
      form.querySelectorAll(SUBMIT_SELECTOR).forEach((button) => setButtonLoading(button, false));
    });
  });
})();
