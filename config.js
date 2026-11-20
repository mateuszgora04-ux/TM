// Konfiguracja TM – uzupełnij, aby włączyć prawdziwe logowanie i integracje.
// Bez kluczy aplikacja działa w trybie DEMO (dane przykładowe, linki "deep link" do Outlook/Teams/Gmail).
window.TM_CONFIG = {
  // Google Cloud Console → APIs & Services → Credentials → OAuth client ID (Web application)
  // Authorized JavaScript origins: http://localhost:8765 oraz adres produkcyjny (https://...)
  GOOGLE_CLIENT_ID: "",
  GOOGLE_SCOPES: "https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/gmail.send",

  // Microsoft Entra ID (Azure AD) → App registrations → SPA, Redirect URI = adres aplikacji
  MS_CLIENT_ID: "",
  MS_TENANT: "common",           // lub ID tenanta InfoConsulting
  MS_SCOPES: ["User.Read", "Calendars.ReadWrite", "Mail.Send"],

  // Telefon (wersja na GitHub Pages): synchronizacja przez plik w OneDrive (Microsoft Graph).
  // Entra → Rejestracje aplikacji → Twoja aplikacja: Client ID, Tenant ID. Można też podać na telefonie (link z laptopa).
  CLOUD_CLIENT_ID: "",
  CLOUD_TENANT: "",
  CLOUD_PATH: "_CLAUDE/TASK MANAGMENT/TM-TaskManagement/TM-TaskManagement/Dane",

  HOURS_JOB_PER_DAY: 8,          // capacity firmowe na dzień roboczy
  HOURS_PRIV_WORKDAY: 2,         // czas prywatny w dzień roboczy
  HOURS_PRIV_WEEKEND: 5,         // czas prywatny w weekend / święto / urlop
  TIMEZONE: "Europe/Warsaw"
};
