/**
 * Light / dark theme switcher.
 *
 * Loaded synchronously in <head> (external file, CSP-safe) so data-bs-theme is
 * set before first paint with no flash. Defaults to the OS preference until the
 * user toggles it explicitly, which is stored in localStorage.
 */
(function () {
  const STORAGE_KEY = 'aranda-theme';
  const media = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null;

  function storedTheme() {
    try {
      const value = window.localStorage.getItem(STORAGE_KEY);
      return value === 'dark' || value === 'light' ? value : null;
    } catch (error) {
      return null;
    }
  }

  function resolveTheme() {
    return storedTheme() || (media && media.matches ? 'dark' : 'light');
  }

  function syncButtons(theme) {
    const isDark = theme === 'dark';
    const label = isDark ? 'Cambiar a modo claro' : 'Cambiar a modo oscuro';
    document.querySelectorAll('[data-theme-toggle]').forEach((button) => {
      button.setAttribute('aria-pressed', String(isDark));
      button.setAttribute('aria-label', label);
      button.setAttribute('title', label);
      const text = button.querySelector('[data-theme-label]');
      if (text) text.textContent = isDark ? 'Modo claro' : 'Modo oscuro';
    });
  }

  function applyTheme(theme) {
    document.documentElement.setAttribute('data-bs-theme', theme);
    syncButtons(theme);
  }

  // Apply before first paint
  applyTheme(resolveTheme());

  document.addEventListener('DOMContentLoaded', () => {
    syncButtons(document.documentElement.getAttribute('data-bs-theme') || resolveTheme());

    document.addEventListener('click', (event) => {
      const button = event.target.closest('[data-theme-toggle]');
      if (!button) return;
      const current = document.documentElement.getAttribute('data-bs-theme') || resolveTheme();
      const next = current === 'dark' ? 'light' : 'dark';
      try {
        window.localStorage.setItem(STORAGE_KEY, next);
      } catch (error) {
        // Storage unavailable (private mode); still switch for this view
      }
      applyTheme(next);
    });
  });

  if (media && media.addEventListener) {
    media.addEventListener('change', (event) => {
      if (!storedTheme()) {
        applyTheme(event.matches ? 'dark' : 'light');
      }
    });
  }
})();
