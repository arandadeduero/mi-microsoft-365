/**
 * Client-side helpers for Administración → Usuarios pages:
 * - [data-copy-target]: copy text/value from a target element to the clipboard.
 * - [data-check-all]: toggle a group of checkboxes matching a CSS selector.
 * - [data-filter-list]: live case- and accent-insensitive filtering of checkbox lists.
 * - [data-filter-type-toggle]: switch between department and job-title selects.
 */
(function () {
  function copyText(text) {
    if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
      return navigator.clipboard.writeText(text).catch(() => fallbackCopy(text));
    }
    return fallbackCopy(text);
  }

  function fallbackCopy(text) {
    const tmp = document.createElement('textarea');
    tmp.value = text;
    tmp.setAttribute('readonly', '');
    tmp.style.position = 'fixed';
    tmp.style.left = '-9999px';
    document.body.appendChild(tmp);
    tmp.select();
    try {
      document.execCommand('copy');
    } finally {
      tmp.remove();
    }
    return Promise.resolve();
  }

  document.querySelectorAll('[data-copy-target]').forEach((button) => {
    button.addEventListener('click', () => {
      const selector = button.dataset.copyTarget;
      const target = selector ? document.querySelector(selector) : null;
      if (!target) return;

      const text = 'value' in target ? String(target.value) : (target.textContent || '');
      if (!text) return;

      if (!button.dataset.originalLabel) {
        button.dataset.originalLabel = button.textContent.trim();
      }

      copyText(text).then(() => {
        button.textContent = 'Copiado';
        window.clearTimeout(Number(button.dataset.copyTimer || 0));
        const timer = window.setTimeout(() => {
          button.textContent = button.dataset.originalLabel;
        }, 1800);
        button.dataset.copyTimer = String(timer);
      });
    });
  });

  document.querySelectorAll('[data-check-all]').forEach((master) => {
    const selector = master.dataset.checkAll;
    if (!selector) return;

    const getItems = () => Array.from(document.querySelectorAll(selector));

    master.addEventListener('change', () => {
      const checked = master.checked;
      getItems().forEach((box) => {
        box.checked = checked;
      });
    });

    getItems().forEach((box) => {
      box.addEventListener('change', () => {
        const items = getItems();
        const checkedCount = items.filter((item) => item.checked).length;
        master.checked = items.length > 0 && checkedCount === items.length;
        master.indeterminate = checkedCount > 0 && checkedCount < items.length;
      });
    });
  });

  function normalizeText(value) {
    return String(value || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  }

  document.querySelectorAll('[data-filter-list]').forEach((input) => {
    const selector = input.dataset.filterList;
    const container = selector ? document.querySelector(selector) : null;
    if (!container) return;

    input.addEventListener('input', () => {
      const query = normalizeText(input.value.trim());
      const items = container.querySelectorAll('[data-filter-item], .form-check');
      items.forEach((item) => {
        const matches = !query || normalizeText(item.textContent).includes(query);
        item.classList.toggle('d-none', !matches);
      });
    });
  });

  document.querySelectorAll('[data-filter-type-toggle]').forEach((select) => {
    const form = select.closest('form');
    if (!form) return;

    const syncVisibility = () => {
      const activeType = select.value;
      form.querySelectorAll('[data-filter-type-group]').forEach((group) => {
        const isActive = group.dataset.filterTypeGroup === activeType;
        group.classList.toggle('d-none', !isActive);
        group.querySelectorAll('select, input').forEach((field) => {
          field.disabled = !isActive;
        });
      });
    };

    select.addEventListener('change', syncVisibility);
    syncVisibility();
  });

  document.querySelectorAll('[data-enable-target]').forEach((checkbox) => {
    const selector = checkbox.dataset.enableTarget;
    const target = selector ? document.querySelector(selector) : null;
    if (!target) return;

    const syncDisabled = () => {
      target.disabled = !checkbox.checked;
    };

    checkbox.addEventListener('change', syncDisabled);
    syncDisabled();

    const modal = checkbox.closest('.modal');
    if (modal) {
      modal.addEventListener('hidden.bs.modal', () => {
        checkbox.checked = false;
        syncDisabled();
      });
    }
  });
})();
