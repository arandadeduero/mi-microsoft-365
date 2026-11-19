/**
 * Groups page: pick groups and ask to join them.
 */
(function () {
  const checkboxes = Array.from(document.querySelectorAll('.group-checkbox'));
  const sendButton = document.getElementById('generateBtn');
  const dialog = document.getElementById('reportDialog');
  if (!sendButton || !dialog) return;

  const countEl = document.getElementById('selectedCount');
  const modal = () => bootstrap.Modal.getOrCreateInstance(dialog);
  const selected = () => checkboxes.filter((cb) => cb.checked).map((cb) => ({ id: cb.dataset.groupId, name: cb.dataset.groupName }));

  function updateSelection() {
    let count = 0;
    for (let i = 0; i < checkboxes.length; i += 1) {
      if (checkboxes[i].checked) count += 1;
    }
    countEl.textContent = count;
    sendButton.disabled = count === 0;
  }

  function openDialog() {
    const groups = selected();
    if (groups.length === 0) return;
    const list = document.getElementById('selectedGroupsList');
    const fragment = document.createDocumentFragment();
    groups.forEach((group) => {
      const item = document.createElement('li');
      item.textContent = group.name;
      fragment.appendChild(item);
    });
    list.replaceChildren(fragment);
    modal().show();
  }

  const submitBtn = document.getElementById('submitGroupRequestBtn');

  async function submitRequest() {
    if (submitBtn.disabled) return;
    const groups = selected();
    if (groups.length === 0) return;
    window.setButtonLoading?.(submitBtn, true);
    try {
      const response = await fetch('/api/group-access-request', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-CSRF-Token': csrfToken() },
        body: JSON.stringify({ selectedGroups: groups.map((group) => group.id) }),
      });
      const data = await response.json();
      if (data.success) {
        window.alert('Solicitud de acceso enviada correctamente');
        modal().hide();
      } else {
        window.alert('Error al enviar la solicitud: ' + (data.error || 'Error desconocido'));
      }
    } catch (error) {
      console.error('Error sending request:', error);
      window.alert('No se pudo enviar la solicitud. Inténtalo de nuevo.');
    } finally {
      window.setButtonLoading?.(submitBtn, false);
    }
  }

  document.addEventListener('change', (event) => {
    if (event.target.classList?.contains('group-checkbox')) updateSelection();
  });
  sendButton.addEventListener('click', openDialog);
  submitBtn.addEventListener('click', submitRequest);
  document.getElementById('closeDialogBtn').addEventListener('click', () => modal().hide());
  updateSelection();
})();
