/**
 * The requests any person can make from /peticiones. Same shape as
 * config/teamRequests.js; all text is here so the form can change without
 * touching the code. `prefill` names the profile field that fills a field in.
 */
module.exports = {
  soporte: {
    title: 'Soporte informático',
    event: 'request.support',
    intro: 'Cuéntanos qué te pasa con el ordenador, el correo, una impresora o cualquier programa. Con estos datos Informática no tiene que volver a preguntarte lo básico.',
    fields: [
      { name: 'category', label: 'Qué tipo de problema es', type: 'select', required: true, options: [
        'Ordenador o portátil', 'Impresora o escáner', 'Correo (Outlook)', 'Teams, OneDrive o SharePoint',
        'Cuenta, contraseña o acceso', 'Programa o aplicación', 'Teléfono o centralita', 'Otra cosa',
      ] },
      { name: 'impact', label: 'Cuánto te afecta', type: 'select', required: true, options: [
        'No puedo trabajar', 'Puedo trabajar, pero con dificultad', 'Es una molestia, no es urgente',
      ] },
      { name: 'subject', label: 'Asunto', placeholder: 'en pocas palabras', required: true },
      { name: 'description', label: 'Qué ocurre', type: 'textarea', placeholder: 'qué intentabas hacer, qué mensaje sale y desde cuándo pasa', required: true },
      { name: 'device', label: 'Equipo afectado', placeholder: 'nombre del equipo o número de la etiqueta de inventario', required: false },
      { name: 'location', label: 'Dónde estás', placeholder: 'edificio y despacho', required: false, prefill: 'officeLocation' },
      { name: 'phone', label: 'Teléfono para contactarte', required: false, prefill: 'phone' },
    ],
  },
};
