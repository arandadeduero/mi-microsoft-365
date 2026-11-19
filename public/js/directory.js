/**
 * Directory page: filter the table as the user types.
 */
(function () {
  const search = document.getElementById('directorySearch');
  const rows = Array.from(document.querySelectorAll('#directoryTable tbody tr'));
  const empty = document.getElementById('directoryEmpty');
  if (!search || !rows.length) return;

  // Ignore case and accents, so "garcia" finds "García"
  const normalize = (text) => text.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const texts = rows.map((row) => normalize(row.textContent));

  let frame = null;
  function filterRows() {
    frame = null;
    const words = normalize(search.value).split(/\s+/).filter(Boolean);
    let visible = 0;
    for (let i = 0; i < rows.length; i += 1) {
      const text = texts[i];
      let matches = true;
      for (let j = 0; j < words.length; j += 1) {
        if (!text.includes(words[j])) {
          matches = false;
          break;
        }
      }
      rows[i].classList.toggle('d-none', !matches);
      if (matches) visible += 1;
    }
    empty.classList.toggle('d-none', visible > 0);
  }

  search.addEventListener('input', function () {
    if (frame !== null) cancelAnimationFrame(frame);
    frame = requestAnimationFrame(filterRows);
  });
})();
