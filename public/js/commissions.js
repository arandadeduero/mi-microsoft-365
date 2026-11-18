/**
 * Picks the usual room of the kind of meeting chosen, and warns, without
 * stopping anything, when the chosen start is more than a week away in either
 * direction: usually a slip of the month.
 */
const commission = document.getElementById('commission');
const room = document.getElementById('room');

if (commission && room) {
  commission.addEventListener('change', () => {
    const usual = commission.selectedOptions[0].dataset.room;
    if (usual) room.value = usual;
  });
}

const start = document.getElementById('start');
const warning = document.getElementById('start-warning');

if (start && warning) {
  start.addEventListener('input', () => {
    const days = (new Date(start.value) - Date.now()) / 86400000;
    warning.textContent = days < -7 ? 'Ojo: esa fecha es de hace más de 7 días.'
      : days > 7 ? 'Ojo: faltan más de 7 días para esa fecha.'
        : '';
  });
}
