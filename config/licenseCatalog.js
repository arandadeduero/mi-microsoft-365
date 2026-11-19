/**
 * License Catalog
 *
 * Microsoft names licenses and the services inside them with codes. This
 * file says what they are in plain Spanish. A license or service that is not
 * listed is shown by its code (licenses) or left out (services), so the page
 * stays correct when Microsoft adds new ones.
 */

// License (SKU) codes and the name Microsoft sells them under
const SKU_NAMES = {
  // Microsoft 365 & Office 365 Suites
  O365_BUSINESS_ESSENTIALS: 'Microsoft 365 Empresa Básico',
  SMB_BUSINESS_ESSENTIALS: 'Microsoft 365 Empresa Básico',
  O365_BUSINESS_PREMIUM: 'Microsoft 365 Empresa Estándar',
  SMB_BUSINESS_PREMIUM: 'Microsoft 365 Empresa Estándar',
  SPB: 'Microsoft 365 Empresa Premium',
  O365_BUSINESS: 'Aplicaciones de Microsoft 365 para negocios',
  SMB_BUSINESS: 'Aplicaciones de Microsoft 365 para negocios',
  OFFICESUBSCRIPTION: 'Aplicaciones de Microsoft 365 para empresas',
  STANDARDPACK: 'Office 365 E1',
  STANDARDWOFFPACK: 'Office 365 E2',
  ENTERPRISEPACK: 'Office 365 E3',
  ENTERPRISEPACKLRG: 'Office 365 E3',
  ENTERPRISEPREMIUM: 'Office 365 E5',
  ENTERPRISEPREMIUM_NOPSTNCONF: 'Office 365 E5 (sin audioconferencia)',
  DESKLESSPACK: 'Office 365 F3',
  SPE_E3: 'Microsoft 365 E3',
  SPE_E5: 'Microsoft 365 E5',
  SPE_F1: 'Microsoft 365 F3',
  SPE_F3: 'Microsoft 365 F3',
  M365_F1: 'Microsoft 365 F1',
  M365_F1_COMM: 'Microsoft 365 F1',
  Microsoft_365_Copilot: 'Microsoft 365 Copilot',
  MICROSOFT_365_COPILOT: 'Microsoft 365 Copilot',
  MICROSOFT_365_COPILOT_FOR_BUSINESS: 'Microsoft 365 Copilot',
  Microsoft_365_Copilot_for_Business: 'Microsoft 365 Copilot',

  // Exchange Online
  EXCHANGESTANDARD: 'Exchange Online (Plan 1)',
  EXCHANGEENTERPRISE: 'Exchange Online (Plan 2)',
  EXCHANGEDESKLESS: 'Exchange Online Kiosk',
  EXCHANGEARCHIVE: 'Archivo en línea de Exchange',
  EXCHANGEARCHIVE_ADDON: 'Archivo en línea de Exchange',
  EXCHANGETELCO: 'Exchange Online POP',
  SMB_APPS: 'Microsoft Business Center',

  // Copilot Studio & Power Platform (Power Automate, Power BI, Power Apps, Power Pages)
  CCIBOTS_PRIVPREV_VIRAL: 'Microsoft Copilot Studio (prueba gratuita)',
  VIRTUAL_AGENT_USL: 'Microsoft Copilot Studio (licencia de usuario)',
  POWER_VIRTUAL_AGENTS: 'Microsoft Copilot Studio',
  FLOW_FREE: 'Power Automate (gratuito)',
  FLOW_PER_USER: 'Power Automate por usuario',
  FLOW_P1: 'Power Automate (Plan 1)',
  FLOW_P2: 'Power Automate (Plan 2)',
  POWERAUTOMATE_ATTENDED_RPA: 'Power Automate Premium',
  POWERAUTOMATE_UNATTENDED_RPA: 'Power Automate Process',
  POWER_BI_STANDARD: 'Power BI (gratuito)',
  POWER_BI_PRO: 'Power BI Pro',
  POWER_BI_PREMIUM_PER_USER: 'Power BI Premium por usuario',
  PBI_PREMIUM_PER_USER: 'Power BI Premium por usuario',
  POWERAPPS_VIRAL: 'Power Apps (prueba gratuita)',
  POWERAPPS_DEV: 'Power Apps para desarrolladores',
  POWERAPPS_PER_USER: 'Power Apps Premium',
  POWERAPPS_PER_APP: 'Power Apps por aplicación',
  POWERPAGES_VIRAL: 'Power Pages (prueba gratuita)',
  Power_Pages_vTrial_for_Makers: 'Power Pages (prueba gratuita)',
  POWER_PAGES_VTRIAL_FOR_MAKERS: 'Power Pages (prueba gratuita)',

  // Microsoft Teams, Telefonía y Salas
  TEAMS_EXPLORATORY: 'Teams Exploratory',
  TEAMS_FREE: 'Microsoft Teams (gratuito)',
  TEAMS_ESSENTIALS: 'Microsoft Teams Essentials',
  TEAMS_PREMIUM: 'Microsoft Teams Premium',
  MCOEV: 'Teléfono de Microsoft Teams Estándar',
  PHONESYSTEM_VIRTUALUSER: 'Cuenta de recursos del sistema telefónico de Teams',
  MCOMEETADV: 'Audioconferencia de Microsoft Teams',
  MCOIMP: 'Skype Empresarial Online (Plan 1)',
  MCOSTANDARD: 'Skype Empresarial Online (Plan 2)',
  MCOPSTN1: 'Plan de llamadas nacionales de Teams',
  MCOPSTN2: 'Plan de llamadas nacionales e internacionales de Teams',
  MCOPSTNC: 'Créditos de comunicación',
  MEETING_ROOM: 'Salas de Microsoft Teams Estándar',
  TEAMS_ROOMS_BASIC: 'Salas de Microsoft Teams Básico',
  Microsoft_Teams_Rooms_Basic: 'Salas de Microsoft Teams Básico',
  MICROSOFT_TEAMS_ROOMS_BASIC: 'Salas de Microsoft Teams Básico',
  TEAMS_ROOMS_PRO: 'Salas de Microsoft Teams Pro',
  Microsoft_Teams_Rooms_Pro: 'Salas de Microsoft Teams Pro',
  MICROSOFT_TEAMS_ROOMS_PRO: 'Salas de Microsoft Teams Pro',

  // SharePoint, OneDrive, Stream, Clipchamp, Visio y Project
  SHAREPOINTSTANDARD: 'SharePoint Online (Plan 1)',
  SHAREPOINTENTERPRISE: 'SharePoint Online (Plan 2)',
  ONEDRIVESTANDARD: 'OneDrive para la Empresa (Plan 1)',
  WACONEDRIVESTANDARD: 'OneDrive para la Empresa (Plan 1)',
  WACONEDRIVEENTERPRISE: 'OneDrive para la Empresa (Plan 2)',
  STREAM: 'Microsoft Stream',
  CLIPCHAMP: 'Microsoft Clipchamp',
  FORMS_PRO: 'Dynamics 365 Customer Voice (Forms Pro)',
  VISIOONLINE_PLAN1: 'Visio (Plan 1)',
  VISIOCLIENT: 'Visio (Plan 2)',
  PROJECT_P1: 'Planner y Project (Plan 1)',
  PROJECTLITE: 'Project Online Essentials',
  PROJECTESSENTIALS: 'Project Online Essentials',
  PROJECTPROFESSIONAL: 'Planner y Project (Plan 3)',
  PROJECTPREMIUM: 'Planner y Project (Plan 5)',
  PROJECTONLINE_PLAN_1: 'Project Online Premium',

  // Entra ID, Seguridad, Intune y Windows
  AAD_BASIC: 'Microsoft Entra ID Básico',
  AAD_PREMIUM: 'Microsoft Entra ID P1',
  AAD_PREMIUM_P2: 'Microsoft Entra ID P2',
  EMS: 'Enterprise Mobility + Security E3',
  EMSPREMIUM: 'Enterprise Mobility + Security E5',
  INTUNE_A: 'Microsoft Intune (Plan 1)',
  INTUNE_A_D: 'Microsoft Intune para dispositivos',
  ATP_ENTERPRISE: 'Microsoft Defender para Office 365 (Plan 1)',
  THREAT_INTELLIGENCE: 'Microsoft Defender para Office 365 (Plan 2)',
  WIN_DEF_ATP: 'Microsoft Defender para punto de conexión',
  DEFENDER_ENDPOINT_P1: 'Microsoft Defender para punto de conexión P1',
  MDATP_XPLAT: 'Microsoft Defender para punto de conexión P2',
  RIGHTSMANAGEMENT: 'Azure Information Protection Premium P1',
  RIGHTSMANAGEMENT_ADHOC: 'Rights Management Adhoc',
  WINDOWS_STORE: 'Tienda Windows para Empresas',
  WIN10_PRO_ENT_SUB: 'Windows 10/11 Enterprise E3',
  WIN10_VDA_E5: 'Windows 10/11 Enterprise E5',
};

// What a user can use, and the service plans that grant it. A plan matches
// when its name starts with the text given, since Microsoft adds suffixes
// for editions and regions. `detail` overrides the description for that plan.
const APPS = [
  {
    name: 'Correo y calendario (Outlook)',
    description: 'Tu buzón, calendario y contactos',
    url: 'https://outlook.office.com/mail/',
    plans: [
      { name: 'EXCHANGE_S_ENTERPRISE', detail: 'Tu buzón, calendario y contactos. Buzón de 100 GB' },
      { name: 'EXCHANGE_S_STANDARD', detail: 'Tu buzón, calendario y contactos. Buzón de 50 GB' },
      { name: 'EXCHANGE_S_DESKLESS', detail: 'Tu buzón, calendario y contactos. Buzón de 2 GB, solo desde el navegador o el móvil' },
    ],
  },
  {
    name: 'Teams',
    description: 'Chat, reuniones y trabajo en equipo',
    url: 'https://teams.microsoft.com',
    plans: [{ name: 'TEAMS1' }],
  },
  {
    name: 'OneDrive y SharePoint',
    description: 'Tus archivos en la nube y los de tus equipos',
    url: 'https://www.office.com/launch/onedrive',
    plans: [{ name: 'SHAREPOINTENTERPRISE' }, { name: 'SHAREPOINTSTANDARD' }, { name: 'SHAREPOINTDESKLESS' }],
  },
  {
    name: 'Word, Excel y PowerPoint en el navegador',
    description: 'Crear y editar documentos sin instalar nada',
    url: 'https://www.office.com',
    plans: [{ name: 'SHAREPOINTWAC' }],
  },
  {
    name: 'Word, Excel y PowerPoint para instalar',
    description: 'Las aplicaciones de escritorio, en tu ordenador',
    url: 'https://portal.office.com/account#installs',
    plans: [{ name: 'OFFICE_BUSINESS' }, { name: 'OFFICESUBSCRIPTION' }],
  },
  {
    name: 'Planner',
    description: 'Repartir y seguir las tareas de un equipo',
    url: 'https://tasks.office.com',
    plans: [{ name: 'PROJECTWORKMANAGEMENT' }],
  },
  {
    name: 'To Do',
    description: 'Tu lista de tareas personal',
    url: 'https://to-do.office.com',
    plans: [{ name: 'BPOS_S_TODO' }],
  },
  {
    name: 'Forms',
    description: 'Formularios y encuestas',
    url: 'https://forms.office.com',
    plans: [{ name: 'FORMS_PLAN' }],
  },
  {
    name: 'Bookings',
    description: 'Agenda de citas',
    url: 'https://outlook.office.com/bookings',
    plans: [{ name: 'MICROSOFTBOOKINGS' }],
  },
  {
    name: 'Power Automate',
    description: 'Automatizar tareas repetitivas',
    url: 'https://make.powerautomate.com',
    plans: [{ name: 'FLOW_O365' }],
  },
];

// Case-insensitive lookup so mixed-case Graph SKU codes (e.g. Microsoft_Teams_Rooms_Basic) always match
const SKU_NAMES_UPPER = Object.fromEntries(
  Object.entries(SKU_NAMES).map(([k, v]) => [k.toUpperCase(), v])
);
const SKU_NAMES_PROXY = new Proxy(SKU_NAMES, {
  get(target, prop, receiver) {
    if (typeof prop === 'string') {
      if (prop in target) return Reflect.get(target, prop, receiver);
      const upper = prop.toUpperCase();
      if (upper in SKU_NAMES_UPPER) return SKU_NAMES_UPPER[upper];
    }
    return Reflect.get(target, prop, receiver);
  },
});

module.exports = { SKU_NAMES: SKU_NAMES_PROXY, APPS };
