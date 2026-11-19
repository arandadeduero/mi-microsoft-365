// /firma: the concejalía list reloads the page, and the two buttons copy the signature
(function () {
  document.getElementById('concejalia').addEventListener('change', function () {
    document.getElementById('concejalia-form').requestSubmit();
  });

  // The HTML to copy sits in a hidden textarea; the preview shows the same signature
  var source = document.getElementById('signature-html');
  var preview = document.getElementById('signature-preview');
  if (!source || !preview) return;

  var copiers = {
    html: function () {
      return navigator.clipboard.writeText(source.value);
    },
    rich: function () {
      return navigator.clipboard.write([new ClipboardItem({
        'text/html': new Blob([source.value], { type: 'text/html' }),
        'text/plain': new Blob([preview.innerText], { type: 'text/plain' }),
      })]);
    },
  };

  document.querySelectorAll('[data-copy]').forEach(function (button) {
    button.addEventListener('click', async function () {
      if (button.dataset.label) return; // already showing the result
      var label = button.textContent;
      var ok = true;
      try {
        await copiers[button.dataset.copy]();
      } catch (error) {
        ok = false;
      }
      button.dataset.label = label;
      button.textContent = ok ? '✓ ¡Copiado!' : '✗ No se pudo copiar';
      setTimeout(function () {
        button.textContent = label;
        delete button.dataset.label;
      }, 2000);
    });
  });
})();
