/* TM · Task Management – prototyp (vanilla JS, bez zależności)
 * Moduły: auth, dane (localStorage per konto), kontekst ALL/PRIV/JOB, zadania, podprojekty, szybki kreator,
 * kalendarz (dzień/tydzień/miesiąc/rok), wykresy, Gantt, capacity (urlopy, święta PL, weekendy),
 * integracje Outlook/Teams (Microsoft Graph) i Gmail/Google Calendar (Google API) + tryb demo z deep linkami.
 */
'use strict';
const CFG = Object.assign({HOURS_JOB_PER_DAY: 8, HOURS_PRIV_WORKDAY: 2, HOURS_PRIV_WEEKEND: 5, MS_TIMEZONE: 'Central European Standard Time'}, window.TM_CONFIG || {});
const $ = (s, r = document) => r.querySelector(s), $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]));
const uid = () => Math.random().toString(36).slice(2, 10);
const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/ł/g, 'l');
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
const r1 = v => Math.round(v * 10) / 10;

/* ================= ikony (jednolity zestaw liniowy, monochromatyczny) ================= */
const ICONS = {
  schedule: '<rect x="3.5" y="5" width="17" height="15.5" rx="2.5"/><path d="M3.5 10h17M8 3v4M16 3v4"/><path d="M7 13.5h5M10 16.5h7"/>',
  sidebar: '<rect x="3.5" y="4.5" width="17" height="15" rx="2.5"/><path d="M9.5 4.5v15"/>',
  home: '<rect x="3.5" y="3.5" width="17" height="17" rx="2.5"/><path d="M3.5 9.5h17M9.5 9.5v11"/>',
  layers: '<path d="M12 3.5l8.5 4.5L12 12.5 3.5 8z"/><path d="M3.5 12.2L12 16.7l8.5-4.5"/><path d="M3.5 16.2L12 20.7l8.5-4.5"/>',
  briefcase: '<rect x="3" y="7" width="18" height="13" rx="2"/><path d="M9 7V5.5A1.5 1.5 0 0 1 10.5 4h3A1.5 1.5 0 0 1 15 5.5V7M3 12.5h18"/>',
  activity: '<rect x="3.5" y="5" width="17" height="15.5" rx="2.5"/><path d="M3.5 10h17"/><path d="M7.5 13.5h4.5v3.5H7.5z"/>',
  library: '<path d="M5 4v16M10 4v16M14.5 4.5l5 15"/>',
  more: '<circle cx="5.5" cy="12" r="1.3"/><circle cx="12" cy="12" r="1.3"/><circle cx="18.5" cy="12" r="1.3"/>',
  'arrow-right': '<path d="M5 12h14M13 6l6 6-6 6"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  plane: '<path d="M10.5 21l1.5-1v-5.5l8 3v-2l-8-5V4.5a1.5 1.5 0 0 0-3 0V10.5l-8 5v2l8-3V20l1.5 1"/>',
  bolt: '<path d="M13 2.5L4.5 13.5H11L10 21.5l8.5-11H12z"/>',
  gear: '<path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/>',
  monitor: '<rect x="3" y="4" width="18" height="12.5" rx="2"/><path d="M8.5 20h7M12 16.5V20"/>',
  plug: '<path d="M9 3v5M15 3v5M6.5 8h11v3.5a5.5 5.5 0 0 1-11 0zM12 17v4"/>',
  sliders: '<path d="M4 6h10M18 6h2M4 12h4M12 12h8M4 18h12M20 18h0"/><circle cx="16" cy="6" r="2"/><circle cx="10" cy="12" r="2"/><circle cx="18" cy="18" r="2"/>',
  bell: '<path d="M6 16V11a6 6 0 0 1 12 0v5l1.5 2h-15z"/><path d="M10 20.5a2.2 2.2 0 0 0 4 0"/>',
  wallet: '<rect x="3" y="6" width="18" height="13" rx="2.5"/><path d="M3 9.5h18M16.5 14h1.5"/><path d="M6 6l9-2.5 1 2.5"/>',
  trash: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 12a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-12M9 7V4.5A.5.5 0 0 1 9.5 4h5a.5.5 0 0 1 .5.5V7"/>',
  folder: '<rect x="3.5" y="7" width="17" height="13" rx="2.5"/><path d="M8.5 7V5.5a2 2 0 0 1 2-2h3a2 2 0 0 1 2 2V7M3.5 12.5h17"/>',
  check: '<path d="M5 12.5l4.5 4.5L19 7.5"/>',
  'check-square': '<rect x="3.5" y="3.5" width="17" height="17" rx="2.5"/><path d="M8 12.3l2.7 2.7L16 9.5"/>',
  target: '<circle cx="12" cy="12" r="8.5"/><circle cx="12" cy="12" r="4.5"/><circle cx="12" cy="12" r=".6" fill="currentColor"/>',
  calendar: '<rect x="3.5" y="5" width="17" height="15.5" rx="2.5"/><path d="M3.5 10h17M8 3v4M16 3v4"/>',
  user: '<circle cx="12" cy="8.5" r="3.75"/><path d="M4.5 20c.9-3.6 3.9-5.8 7.5-5.8s6.6 2.2 7.5 5.8"/>',
  building: '<path d="M4.5 20.5V5A1.5 1.5 0 0 1 6 3.5h7A1.5 1.5 0 0 1 14.5 5v15.5M14.5 9.5H18a1.5 1.5 0 0 1 1.5 1.5v9.5M3 20.5h18M8 7.5h3M8 11h3M8 14.5h3"/>',
  list: '<path d="M9 6h11M9 12h11M9 18h11"/><circle cx="4.75" cy="6" r=".9" fill="currentColor"/><circle cx="4.75" cy="12" r=".9" fill="currentColor"/><circle cx="4.75" cy="18" r=".9" fill="currentColor"/>',
  pencil: '<path d="M4 20h4L19.2 8.8a2 2 0 0 0 0-2.8l-1.2-1.2a2 2 0 0 0-2.8 0L4 16v4zM13.5 6.5l4 4"/>',
  'chev-left': '<path d="M14.5 6l-6 6 6 6"/>', 'chev-right': '<path d="M9.5 6l6 6-6 6"/>',
  'chev-down': '<path d="M6 9.5l6 6 6-6"/>', 'chev-up': '<path d="M6 14.5l6-6 6 6"/>',
  undo: '<path d="M9 14.5L4 9.5l5-5"/><path d="M4 9.5h10.5a5.5 5.5 0 0 1 0 11H11"/>',
  gantt: '<path d="M4 4v16h16"/><path d="M8 8h6M10 12h8M8 16h4" stroke-width="2.4"/>',
  chart: '<path d="M4 4v16h16"/><path d="M7.5 15l3.5-4 3 2.5 5-6.5"/>',
  gauge: '<path d="M4.2 17.5a8.5 8.5 0 1 1 15.6 0"/><path d="M12 13.5l3.5-4"/><circle cx="12" cy="14" r="1.2"/>',
  mail: '<rect x="3" y="5" width="18" height="14" rx="2.5"/><path d="M3.5 7l8.5 6 8.5-6"/>',
  alert: '<path d="M10.3 4.3L2.9 17.5A2 2 0 0 0 4.6 20.5h14.8a2 2 0 0 0 1.7-3L13.7 4.3a2 2 0 0 0-3.4 0z"/><path d="M12 9.5v4M12 17h.01"/>',
  sun: '<circle cx="12" cy="12" r="3.75"/><path d="M12 2.5v2M12 19.5v2M4.6 4.6l1.4 1.4M18 18l1.4 1.4M2.5 12h2M19.5 12h2M4.6 19.4L6 18M18 6l1.4-1.4"/>',
  video: '<path d="M4 7a2.5 2.5 0 0 1 2.5-2.5h11A2.5 2.5 0 0 1 20 7v7a2.5 2.5 0 0 1-2.5 2.5H11l-4 3.5v-3.5h-.5A2.5 2.5 0 0 1 4 14z"/><path d="M8.5 9.5h7M8.5 12.5h4"/>',
  pin: '<path d="M12 21s-6.5-5.8-6.5-11a6.5 6.5 0 0 1 13 0c0 5.2-6.5 11-6.5 11z"/><circle cx="12" cy="10" r="2.3"/>',
  users: '<circle cx="9" cy="8.5" r="3.5"/><path d="M3.5 20a5.5 5.5 0 0 1 11 0"/><path d="M15.5 5.2a3.5 3.5 0 0 1 0 6.6M17.5 14.6a5.5 5.5 0 0 1 3 4.9"/>',
  search: '<circle cx="11" cy="11" r="6.5"/><path d="M20 20l-4.2-4.2"/>',
  columns: '<rect x="3.5" y="4" width="17" height="16" rx="2.5"/><path d="M9.2 4v16M14.8 4v16"/>',
  x: '<path d="M6.5 6.5l11 11M17.5 6.5l-11 11"/>',
  flag: '<path d="M5.5 21V4M5.5 4.5h11l-2.2 4 2.2 4h-11"/>',
  pie: '<path d="M12 3.5a8.5 8.5 0 1 0 8.5 8.5H12z"/><path d="M15 3.9A8.5 8.5 0 0 1 20.1 9H15z"/>',
  settings: '<path d="M4 7h9M17 7h3M4 17h3M11 17h9"/><circle cx="15" cy="7" r="2"/><circle cx="9" cy="17" r="2"/>',
  repeat: '<path d="M17 2.5l3 3-3 3"/><path d="M4 11.5v-1a5 5 0 0 1 5-5h11"/><path d="M7 21.5l-3-3 3-3"/><path d="M20 12.5v1a5 5 0 0 1-5 5H4"/>',
  refresh: '<path d="M19.5 12a7.5 7.5 0 1 1-2.2-5.3"/><path d="M19.5 4.5v4h-4"/>',
  zap: '<path d="M13 3L5 13.5h6.5L10.5 21 19 10.5h-6.5z"/>',
  external: '<path d="M14 4.5h5.5V10M19.5 4.5L11 13M17.5 14v4.5a1.5 1.5 0 0 1-1.5 1.5H6a1.5 1.5 0 0 1-1.5-1.5V8.5A1.5 1.5 0 0 1 6 7h4.5"/>',
  download: '<path d="M12 4v11M7.5 10.5L12 15l4.5-4.5M4.5 19.5h15"/>',
  upload: '<path d="M12 15V4M7.5 8.5L12 4l4.5 4.5M4.5 19.5h15"/>',
  rotate: '<path d="M4.5 4.5v5h5"/><path d="M5 14a7.5 7.5 0 1 0 1.6-7.4L4.5 9.5"/>',
  timer: '<circle cx="12" cy="13.5" r="7"/><path d="M12 10v3.5l2.2 1.6M9.5 3.5h5M18.2 6.8l1.3-1.3"/>',
  play: '<path d="M8 5.5v13l10.5-6.5z"/>',
  pause: '<path d="M9 5.5v13M15 5.5v13" stroke-width="2.2"/>',
  switch: '<path d="M7 7.5h12.5L16.5 4.5M17 16.5H4.5l3 3"/>',
  template: '<rect x="3.5" y="3.5" width="17" height="6.5" rx="1.8"/><rect x="3.5" y="13" width="7.5" height="7.5" rx="1.8"/><rect x="13" y="13" width="7.5" height="7.5" rx="1.8"/>',
  note: '<path d="M14 3.5H7A2.5 2.5 0 0 0 4.5 6v12A2.5 2.5 0 0 0 7 20.5h10a2.5 2.5 0 0 0 2.5-2.5V9z"/><path d="M14 3.5V9h5.5M8.5 13h7M8.5 16.5h5"/>',
  eye: '<path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z"/><circle cx="12" cy="12" r="3"/>',
  'eye-off': '<path d="M10.6 5.6A10 10 0 0 1 12 5.5c6 0 9.5 6.5 9.5 6.5a17 17 0 0 1-2.6 3.4M6.6 6.6C3.9 8.3 2.5 12 2.5 12s3.5 6.5 9.5 6.5c1.7 0 3.2-.5 4.5-1.2M9.9 9.9a3 3 0 0 0 4.2 4.2M3.5 3.5l17 17"/>',
  lock: '<rect x="5" y="10.5" width="14" height="10" rx="2.5"/><path d="M8.5 10.5V7.5a3.5 3.5 0 0 1 7 0v3"/>',
  clock: '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/>',
  dashboard: '<rect x="3.5" y="3.5" width="7" height="9" rx="1.8"/><rect x="13.5" y="3.5" width="7" height="5" rx="1.8"/><rect x="13.5" y="11.5" width="7" height="9" rx="1.8"/><rect x="3.5" y="15.5" width="7" height="5" rx="1.8"/>'
};
const ic = (n, cls = '') => `<svg class="ic ${cls}" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">${ICONS[n] || ''}</svg>`;
/* zamiana znaczników ⟨i:nazwa⟩ na ikony SVG – stosowana przy każdym wstawianiu HTML */
const H = html => String(html).replace(/⟨i:([\w-]+)⟩/g, (_, n) => ic(n));

/* ================= daty ================= */
const DAY = 864e5;
const pad = n => String(n).padStart(2, '0');
const iso = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const P = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
const today = () => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; };
const TODAY = () => iso(today());
const add = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
const addM = (d, n) => new Date(d.getFullYear(), d.getMonth() + n, 1);
const sow = d => { const x = new Date(d); x.setHours(0, 0, 0, 0); return add(x, -((x.getDay() + 6) % 7)); };
const diffD = (a, b) => Math.round((P(b) - P(a)) / DAY); // b - a (stringi ISO)
const MON = ['Styczeń', 'Luty', 'Marzec', 'Kwiecień', 'Maj', 'Czerwiec', 'Lipiec', 'Sierpień', 'Wrzesień', 'Październik', 'Listopad', 'Grudzień'];
const MONS = ['sty', 'lut', 'mar', 'kwi', 'maj', 'cze', 'lip', 'sie', 'wrz', 'paź', 'lis', 'gru'];
const DOW = ['Pn', 'Wt', 'Śr', 'Cz', 'Pt', 'Sb', 'Nd'];
const DOWL = ['Niedziela', 'Poniedziałek', 'Wtorek', 'Środa', 'Czwartek', 'Piątek', 'Sobota'];
const fmt = s => { if (!s) return '—'; const d = P(s); return `${d.getDate()} ${MONS[d.getMonth()]}${d.getFullYear() !== today().getFullYear() ? ' ' + d.getFullYear() : ''}`; };
const relDay = s => { const n = diffD(TODAY(), s); return n === 0 ? 'dziś' : n === 1 ? 'jutro' : n === -1 ? 'wczoraj' : n < 0 ? `${-n} dni temu` : n < 7 ? DOWL[P(s).getDay()].toLowerCase() : fmt(s); };
const hm = s => s.slice(11, 16);
const toMin = s => +s.slice(11, 13) * 60 + +s.slice(14, 16);

/* Święta w Polsce (z Wielkanocą wg algorytmu Meeusa) */
const _hol = {};
function holidays(y) {
  if (_hol[y]) return _hol[y];
  const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25),
    g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4,
    l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451),
    mo = Math.floor((h + l - 7 * m + 114) / 31), da = ((h + l - 7 * m + 114) % 31) + 1;
  const E = new Date(y, mo - 1, da), D = (mm, dd) => new Date(y, mm - 1, dd), r = {}, s = (dt, n) => r[iso(dt)] = n;
  s(D(1, 1), 'Nowy Rok'); s(D(1, 6), 'Trzech Króli'); s(E, 'Wielkanoc'); s(add(E, 1), 'Poniedziałek Wielkanocny');
  s(D(5, 1), 'Święto Pracy'); s(D(5, 3), 'Święto Konstytucji 3 Maja'); s(add(E, 49), 'Zielone Świątki'); s(add(E, 60), 'Boże Ciało');
  s(D(8, 15), 'Wniebowzięcie NMP'); s(D(11, 1), 'Wszystkich Świętych'); s(D(11, 11), 'Narodowe Święto Niepodległości');
  if (y >= 2025) s(D(12, 24), 'Wigilia'); s(D(12, 25), 'Boże Narodzenie'); s(D(12, 26), 'Drugi dzień Świąt');
  return _hol[y] = r;
}
const holName = d => holidays(d.getFullYear())[iso(d)];
const isWE = d => d.getDay() === 0 || d.getDay() === 6;
const vacOf = d => { const s = iso(d); return (S.data?.vacations || []).find(v => s >= v.from && s <= v.to); };
const dayType = d => { if (holName(d)) return 'hol'; const v = vacOf(d); return v && !v.half ? 'vac' : isWE(d) ? 'we' : 'work'; };
const halfVac = d => { const v = vacOf(d); return !!(v && v.half && !isWE(d) && !holName(d)); };
const isWork = d => dayType(d) === 'work';

/* ================= stan + storage ================= */
const LS = {
  get(k, d) { try { return JSON.parse(localStorage.getItem(k)) ?? d; } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { } },
  del(k) { try { localStorage.removeItem(k); } catch { } }
};
const S = {
  user: null, ctx: 'all', view: 'dashboard', data: null, privacy: true,
  cal: {mode: 'week', date: today()},
  tasks: {mode: 'list', q: '', status: 'open', goal: '', proj: '', selMode: false, sel: new Set()},
  goals: {proj: '', group: true},
  proj: {status: 'active', open: null, q: '', pq: '', sort: 'status', dir: 1, tstat: 'open', mper: 'next'},
  mtgs: {q: '', period: 'cal', from: '', to: '', proj: '', src: '', sort: 'asc', sel: new Set()},
  gantt: {days: 42, start: add(sow(today()), -7), group: 'project'},
  cap: {weeks: 12},
  integ: {tab: 'meetings', fbDate: TODAY(), people: [], prefill: null},
  conn: {ms: false, google: false}
};
const save = () => { if (tabInactive) return; S.data.savedAt = Date.now(); stampChanges(S.data); LS.set('tm:data:' + S.user.email, S.data); diskSchedule(); cloudSchedule(); };
/* ---------- dane w pliku: <folder TM>\Dane\TM-<e-mail>.json (OneDrive) przez serwer start-TM; przeglądarka = kopia robocza ---------- */
/* błąd połączenia z serwerem TM (start-TM.bat nie działa) → zrozumiały komunikat */
const errMsg = e => (e instanceof TypeError && /fetch|network/i.test(e.message)) ? 'Serwer TM nie działa – uruchom start-TM.bat' : String((e && e.message) || e);
const DISK = {on: /^(localhost|127\.0\.0\.1)$/.test(location.hostname), ready: false, ok: null, err: '', last: 0, path: '', t: null, busy: false, again: false};
const diskUrl = (ep, q = '') => `${ep}?u=${encodeURIComponent(S.user.email)}${q}`;
const diskHdr = {'X-TM': '1'};
/* jedno aktywne okno TM: nowe okno przejmuje pracę, starsze przechodzi w tryb uśpienia (bez zapisów), żeby nie nadpisywać danych */
let tabInactive = false;
const TAB_ID = Math.random().toString(36).slice(2);
const TABCH = (() => { try { return new BroadcastChannel('tm-app'); } catch { return null; } })();
if (TABCH) TABCH.onmessage = e => { if (e.data && e.data.t === 'claim' && e.data.id !== TAB_ID) tabSleep(); };
const tabClaim = () => { if (TABCH) TABCH.postMessage({t: 'claim', id: TAB_ID}); };
function tabSleep() {
  if (tabInactive) return; tabInactive = true; DISK.ready = false; clearTimeout(DISK.t);
  document.body.insertAdjacentHTML('beforeend', `<div class="tab-sleep"><div class="card"><h3>TM jest otwarty w innym oknie</h3><p class="muted">To okno zostało wstrzymane, żeby dwa okna nie nadpisywały sobie danych. Pracuj w nowszym oknie albo przejmij pracę tutaj.</p><button class="btn primary" onclick="location.reload()">Używaj TM w tym oknie</button></div></div>`);
}
function diskSchedule() { if (!DISK.on || !DISK.ready) return; clearTimeout(DISK.t); DISK.t = setTimeout(diskWrite, 600); }
async function diskWrite() {
  if (!S.user || !DISK.ready) return;
  if (DISK.busy) { DISK.again = true; return; }
  DISK.busy = true;
  try {
    if (DISK.rev) await diskPull().catch(() => { });             // najpierw zmiany z telefonu (inny _rev w pliku)
    S.data._rev = uid(); LS.set('tm:data:' + S.user.email, S.data);
    const r = await fetch(diskUrl('/__data'), {method: 'POST', headers: {...diskHdr, 'Content-Type': 'text/plain;charset=utf-8'}, body: JSON.stringify(S.data)});
    if (!r.ok) throw new Error(r.status === 404 || r.status === 405 ? 'Działa stara wersja serwera TM – uruchom ponownie start-TM.bat' : (await r.text()).slice(0, 200));
    const j = await r.json(); DISK.ok = true; DISK.err = ''; DISK.last = Date.now(); DISK.path = j.path || DISK.path; DISK.rev = S.data._rev; cloudBadge();
  } catch (e) {
    const first = DISK.ok !== false; DISK.ok = false; DISK.err = errMsg(e);
    if (first) toast('Nie zapisano danych do pliku w OneDrive: ' + DISK.err + ' (dane są bezpieczne w przeglądarce, ponowię zapis)');
    clearTimeout(DISK.t); DISK.t = setTimeout(diskWrite, 15000);
  } finally {
    DISK.busy = false; if (DISK.again) { DISK.again = false; diskSchedule(); }
    if (S.view === 'profile' && S.profTab === 'data' && $('#diskStatus')) $('#diskStatus').innerHTML = H(diskStatusHtml());
  }
}
/* przy logowaniu: wybór nowszej wersji (plik w OneDrive albo przeglądarka) */
async function diskLoad(prevLocal) {
  if (!DISK.on) return;
  DISK.ready = false;
  try {
    const r = await fetch(diskUrl('/__data'), {headers: diskHdr, cache: 'no-store'});
    if (r.status === 204) { DISK.ready = true; await diskWrite(); if (DISK.ok) toast('Dane TM zapisane w pliku w folderze aplikacji (OneDrive)'); return; }
    if (!r.ok) { DISK.ok = false; DISK.err = r.status === 404 ? 'Działa stara wersja serwera TM – uruchom ponownie start-TM.bat' : (await r.text()).slice(0, 200); return; }
    const fileData = JSON.parse(await r.text()); DISK.path = decodeURIComponent(r.headers.get('X-TM-Path') || '');
    const fs = +fileData.savedAt || 0, ls = prevLocal ? (+prevLocal.savedAt || 0) : -1;
    DISK.ready = true; DISK.ok = true; DISK.last = fs; DISK.rev = fileData._rev;
    if (prevLocal && prevLocal._rev && fileData._rev && prevLocal._rev !== fileData._rev) {
      dataApply(dataMerge(prevLocal, fileData)); render();
      if (JSON.stringify(S.data, noM) !== JSON.stringify(fileData, noM)) await diskWrite();
    } else if (!prevLocal || fs > ls) {
      S.data = migrate(fileData); applySettings(); refreshDemoMeetings(); LS.set('tm:data:' + S.user.email, S.data); snapInit(S.data); render();
      if (prevLocal) toast('Wczytano nowszą wersję danych z pliku w OneDrive');
    } else if (ls > fs) await diskWrite();
  } catch (e) { DISK.ok = false; DISK.err = 'Serwer TM niedostępny – dane tylko w przeglądarce'; }
}
function diskStatusHtml() {
  if (CLOUD.on) return cloudStatusHtml();
  if (!DISK.on) return `<p class="muted" style="margin:0">TM nie działa przez start-TM.bat – dane są tylko w przeglądarce. Uruchom TM przez <b>start-TM.bat</b>, aby zapisywać je w OneDrive.</p>`;
  const st = DISK.ok === false ? `<span class="tag" style="background:#FEE2E2;color:#991B1B">błąd zapisu</span> <span class="muted">${esc(DISK.err)}</span>` : DISK.ok ? `<span class="tag ok-tag">zapisane</span> <span class="muted">ostatni zapis: ${ago(DISK.last)}</span>` : '<span class="muted">sprawdzam…</span>';
  return `<div class="row" style="gap:8px;flex-wrap:wrap">${st}</div>${DISK.path ? `<div class="muted" style="font-size:12px;margin-top:6px;overflow-wrap:anywhere">${esc(DISK.path)}</div>` : ''}`;
}
async function diskBackupsLoad() {
  const box = $('#diskBackups'); if (!box || !DISK.on) return;
  try {
    const r = await fetch(diskUrl('/__backups'), {headers: diskHdr, cache: 'no-store'}); if (!r.ok) throw 0;
    let l = await r.json(); if (!Array.isArray(l)) l = l ? [l] : [];
    const lab = n => { const m = n.match(/-(\d{4}-\d{2}-\d{2})(?:-(\d{2})(\d{2})(\d{2}))?\.json$/); return m ? fmt(m[1]) + (m[2] ? ` ${m[2]}:${m[3]} (przed przywróceniem)` : ' (stan z początku dnia)') : n; };
    box.innerHTML = H(l.length ? `<table class="tbl"><tbody>${l.slice(0, 15).map(b => `<tr><td>⟨i:calendar⟩ ${lab(b.name)}</td><td class="num muted">${Math.max(1, Math.round(b.size / 1024))} KB</td><td class="num"><button class="btn sm" data-action="disk-restore" data-f="${esc(b.name)}">⟨i:undo⟩ Przywróć</button></td></tr>`).join('')}</tbody></table>${l.length > 15 ? `<p class="hint">…i ${l.length - 15} starszych w folderze Kopie</p>` : ''}` : '<p class="muted" style="margin:0">Brak kopii – pierwsza powstanie przy pierwszej zmianie jutro (kopia stanu z początku każdego dnia).</p>');
  } catch { box.innerHTML = '<p class="muted" style="margin:0">Lista kopii dostępna po uruchomieniu nowej wersji serwera (start-TM.bat).</p>'; }
}
/* ================= SYNCHRONIZACJA LAPTOP ⇄ TELEFON =================
   Jeden plik danych w firmowym OneDrive (Dane\TM-<login>.json):
   • laptop zapisuje go przez serwer start-TM (jak dotąd), OneDrive synchronizuje go do chmury,
   • telefon (wersja na GitHub Pages) czyta i zapisuje ten sam plik przez Microsoft Graph.
   Każdy rekord (zadanie, spotkanie, wpis czasu…) ma znacznik zmiany _m, usunięcia trafiają do _del,
   pozostałe pola (ustawienia, stoper…) mają znaczniki w _k. Dzięki temu zmiany z dwóch urządzeń się scalają, a nie nadpisują. */
const SNAP = new Map(); let SNAP_OK = false;
const noM = (k, v) => k === '_m' ? undefined : v;
const isIdArr = a => Array.isArray(a) && a.every(x => x && typeof x === 'object' && x.id != null);
const syncKey = k => !k.startsWith('_') && k !== 'savedAt';
function snapInit(d) {
  SNAP.clear(); SNAP_OK = !!d; if (!d) return;
  for (const [k, v] of Object.entries(d)) { if (!syncKey(k)) continue;
    if (isIdArr(v)) v.forEach(x => SNAP.set(k + '\u0001' + x.id, JSON.stringify(x, noM))); else SNAP.set('\u0002' + k, JSON.stringify(v ?? null)); }
}
function stampChanges(d) {
  if (!SNAP_OK) return snapInit(d);
  const now = Date.now(), seen = new Set(), del = d._del ||= {}, kk = d._k ||= {};
  for (const [k, v] of Object.entries(d)) { if (!syncKey(k)) continue;
    if (isIdArr(v)) { for (const x of v) { const key = k + '\u0001' + x.id, s = JSON.stringify(x, noM); seen.add(key); if (SNAP.get(key) !== s) { x._m = now; SNAP.set(key, s); } } }
    else { const key = '\u0002' + k, s = JSON.stringify(v ?? null); seen.add(key); if (SNAP.get(key) !== s) { kk[k] = now; SNAP.set(key, s); } } }
  for (const key of [...SNAP.keys()]) if (!seen.has(key)) { if (key[0] !== '\u0002') del[key.split('\u0001')[1]] = now; SNAP.delete(key); }
}
/* scalanie dwóch wersji danych (kolejność argumentów bez znaczenia) */
function dataMerge(L, R) {
  if (!L) return R; if (!R) return L;
  const lr = (+R.savedAt || 0) > (+L.savedAt || 0), N = lr ? R : L, O = lr ? L : R, out = {...O, ...N};
  const mx = (a, b) => { const o = {...(a || {})}; for (const [k, v] of Object.entries(b || {})) if (!(o[k] >= v)) o[k] = v; return o; };
  const del = mx(O._del, N._del), kk = mx(O._k, N._k), lim = Date.now() - 120 * 864e5;
  for (const k of new Set([...Object.keys(L), ...Object.keys(R)])) {
    if (!syncKey(k)) continue; const a = N[k], b = O[k];
    if ((a === undefined || isIdArr(a)) && (b === undefined || isIdArr(b))) {
      const m = new Map(); (a || []).forEach(x => m.set(x.id, x));
      (b || []).forEach(x => { const y = m.get(x.id); if (!y || (+x._m || 0) > (+y._m || 0)) m.set(x.id, x); });
      out[k] = [...m.values()].filter(x => !(del[x.id] >= (+x._m || 0)));
    } else out[k] = ((O._k || {})[k] || 0) > ((N._k || {})[k] || 0) ? b : (a === undefined ? b : a);
  }
  out._del = Object.fromEntries(Object.entries(del).filter(([, v]) => v > lim)); out._k = kk;
  out.savedAt = Math.max(+L.savedAt || 0, +R.savedAt || 0); out._rev = N._rev;
  return out;
}
/* podmiana danych po scaleniu (bez ponownego oznaczania rekordów) */
function dataApply(d, why) {
  const before = JSON.stringify(S.data, noM);
  S.data = migrate(d); applySettings(); LS.set('tm:data:' + S.user.email, S.data); snapInit(S.data);
  const changed = before !== JSON.stringify(S.data, noM);
  if (changed && S.user && !$('#app').classList.contains('hidden')) { render(); if (why) toast(why); }
  return changed;
}

/* ---------- laptop: plik przez serwer start-TM + kopie konfliktów OneDrive ---------- */
async function diskPull({conflicts = false} = {}) {
  if (!DISK.on || !DISK.ready || !S.user || tabInactive) return false;
  let changed = false;
  const r = await fetch(diskUrl('/__data'), {headers: diskHdr, cache: 'no-store'});
  if (r.ok && r.status !== 204) { const f = JSON.parse(await r.text());
    if (f._rev && f._rev !== DISK.rev) { DISK.rev = f._rev; changed = dataApply(dataMerge(S.data, f), 'Wczytano zmiany z innego urządzenia (telefon)') || changed;
      if (JSON.stringify(S.data, noM) !== JSON.stringify(f, noM)) DISK.dirty = true; } }
  if (conflicts) { try {
    const c = await fetch(diskUrl('/__data-conflicts'), {headers: diskHdr, cache: 'no-store'});
    if (c.ok) { let l = await c.json(); if (!Array.isArray(l)) l = l ? [l] : []; const done = [];
      l.forEach(x => { try { const d = JSON.parse(x.text); if (d && typeof d === 'object') { dataApply(dataMerge(S.data, d)); done.push(x.name); } } catch { } });
      if (done.length) { DISK.dirty = true; changed = true; render(); toast('Scalono zmiany z kopii konfliktu OneDrive'); await fetch(diskUrl('/__data-conflicts-done'), {method: 'POST', headers: {...diskHdr, 'Content-Type': 'application/json'}, body: JSON.stringify({names: done})}); } }
  } catch { } }
  if (DISK.dirty) { DISK.dirty = false; diskSchedule(); }
  return changed;
}
setInterval(() => { if (document.visibilityState === 'visible' && !DISK.busy) diskPull({conflicts: true}).catch(() => { }); }, 20000);

/* ---------- telefon: Microsoft Graph (OneDrive) ---------- */
const CLOUD_DEF_PATH = '_CLAUDE/TASK MANAGMENT/TM-TaskManagement/TM-TaskManagement/Dane';
const CLOUD = {on: !/^(localhost|127\.0\.0\.1)$/.test(location.hostname) && location.protocol === 'https:', app: null, acct: null, cid: '', etag: '', ready: false, busy: false, again: false, t: null, ok: null, err: '', last: 0, scopes: ['User.Read', 'Files.ReadWrite']};
const cloudBase = () => location.origin + location.pathname.replace(/[^/]*$/, '');
const cloudCfg = () => { const l = LS.get('tm:cloud', {}) || {}; return {clientId: String(l.clientId || CFG.CLOUD_CLIENT_ID || '').trim(), tenant: String(l.tenant || CFG.CLOUD_TENANT || '').trim() || 'organizations', path: String(l.path || CFG.CLOUD_PATH || CLOUD_DEF_PATH).trim().replace(/\\/g, '/').replace(/^\/+|\/+$/g, '')}; };
const cloudFile = () => S.user?.file || ('TM-' + (S.user?.email || '') + '.json');
const odPath = (name) => '/me/drive/root:/' + [...cloudCfg().path.split('/').filter(Boolean), ...(name ? [name] : [])].map(encodeURIComponent).join('/');
/* link konfiguracyjny z laptopa: …/index.html#tmcloud=<base64 JSON> */
function cloudSetupFromHash() {
  const m = location.hash.match(/tmcloud=([^&]+)/); if (!m) return false;
  try { const j = JSON.parse(decodeURIComponent(escape(atob(decodeURIComponent(m[1]))))); if (j.clientId) { LS.set('tm:cloud', {clientId: j.clientId, tenant: j.tenant || '', path: j.path || ''}); history.replaceState(null, '', location.pathname); return true; } } catch { }
  return false;
}
Object.assign(CLOUD, {
  async init() {
    const c = cloudCfg(); if (!c.clientId) throw new Error('Brak identyfikatora aplikacji (Client ID)');
    if (this.app && this.cid === c.clientId + c.tenant) return;
    await loadScript('https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js');
    this.app = new msal.PublicClientApplication({auth: {clientId: c.clientId, authority: 'https://login.microsoftonline.com/' + c.tenant, redirectUri: cloudBase(), navigateToLoginRequestUrl: false}, cache: {cacheLocation: 'localStorage'}});
    if (this.app.initialize) await this.app.initialize();
    this.cid = c.clientId + c.tenant;
    const r = await this.app.handleRedirectPromise().catch(e => { this.err = e.errorMessage || e.message; return null; });
    this.acct = r?.account || this.app.getAllAccounts()[0] || null;
  },
  login() { return this.app.loginRedirect({scopes: this.scopes, prompt: 'select_account'}); },
  async token() {
    try { return (await this.app.acquireTokenSilent({scopes: this.scopes, account: this.acct})).accessToken; }
    catch (e) { const er = new Error('Sesja Microsoft wygasła – zaloguj się ponownie'); er.relogin = true; throw er; }
  },
  async g(path, opt = {}) {
    const r = await fetch('https://graph.microsoft.com/v1.0' + path, {...opt, headers: {...(opt.headers || {}), Authorization: 'Bearer ' + await this.token()}});
    if (r.status === 412) { const e = new Error('Plik zmieniony na innym urządzeniu'); e.conflict = true; throw e; }
    if (r.status === 404) { const e = new Error('Nie znaleziono pliku/folderu w OneDrive: ' + cloudCfg().path); e.notFound = true; throw e; }
    if (!r.ok) { const t = await r.text(); throw new Error(r.status === 401 ? 'Logowanie wygasło – zaloguj się ponownie' : r.status === 403 ? 'Brak uprawnienia Files.ReadWrite (zgoda w rejestracji aplikacji Entra)' : `OneDrive ${r.status}: ${t.slice(0, 160)}`); }
    return r.status === 204 ? null : r.json();
  },
  async list() { const j = await this.g(odPath() + ':/children?$select=name,size,lastModifiedDateTime,file&$top=200'); return (j.value || []).filter(x => x.file && /^TM-.+\.json$/i.test(x.name)).sort((a, b) => b.lastModifiedDateTime.localeCompare(a.lastModifiedDateTime)); },
  async meta() { return this.g(odPath(cloudFile())); },
  async read() { const m = await this.meta(); const r = await fetch(m['@microsoft.graph.downloadUrl'], {cache: 'no-store'}); if (!r.ok) throw new Error('Nie udało się pobrać pliku danych (' + r.status + ')'); return {etag: m.eTag, data: JSON.parse(await r.text())}; },
  async write(d, etag) { const j = await this.g(odPath(cloudFile()) + ':/content', {method: 'PUT', headers: {'Content-Type': 'application/json', ...(etag ? {'If-Match': etag} : {'If-None-Match': '*'})}, body: JSON.stringify(d)}); return j.eTag; }
});
function cloudSchedule() { if (!CLOUD.on || !CLOUD.ready || CLOUD.t) return; CLOUD.t = setTimeout(() => { CLOUD.t = null; cloudWrite(); }, 1500); }
async function cloudWrite(retry = 0) {
  if (!CLOUD.ready || !S.user) return;
  if (CLOUD.busy) { CLOUD.again = true; return; }
  CLOUD.busy = true;
  try {
    S.data._rev = uid(); LS.set('tm:data:' + S.user.email, S.data);
    CLOUD.etag = await CLOUD.write(S.data, CLOUD.etag); CLOUD.ok = true; CLOUD.err = ''; CLOUD.last = Date.now();
  } catch (e) {
    if (e.conflict && retry < 3) { try { const f = await CLOUD.read(); CLOUD.etag = f.etag; dataApply(dataMerge(S.data, f.data), 'Scalono zmiany z laptopa'); CLOUD.busy = false; return cloudWrite(retry + 1); } catch (e2) { e = e2; } }
    const first = CLOUD.ok !== false; CLOUD.ok = false; CLOUD.err = navigator.onLine === false ? 'Brak internetu – zmiany są na telefonie, wyślę je po połączeniu' : errMsg(e);
    if (first && navigator.onLine !== false) e.relogin ? toastAction(CLOUD.err, 'Zaloguj', () => CLOUD.login()) : toast('Synchronizacja: ' + CLOUD.err);
    clearTimeout(CLOUD.t); CLOUD.t = setTimeout(() => { CLOUD.t = null; cloudWrite(); }, 20000);
  } finally {
    CLOUD.busy = false; if (CLOUD.again) { CLOUD.again = false; cloudSchedule(); }
    cloudBadge();
  }
}
async function cloudPull({silent = true} = {}) {
  if (!CLOUD.ready || CLOUD.busy || !S.user || tabInactive) return;
  try {
    const m = await CLOUD.meta(); if (m.eTag === CLOUD.etag) { CLOUD.ok = true; CLOUD.err = ''; CLOUD.last = Date.now(); cloudBadge(); if (!silent) toast('Dane są aktualne'); return; }
    const f = await CLOUD.read(); CLOUD.etag = f.etag; CLOUD.ok = true; CLOUD.err = ''; CLOUD.last = Date.now();
    dataApply(dataMerge(S.data, f.data), silent ? '' : 'Wczytano zmiany z laptopa');
    if (JSON.stringify(S.data, noM) !== JSON.stringify(f.data, noM)) cloudSchedule();
    if (!silent) toast('Zsynchronizowano');
  } catch (e) { CLOUD.ok = false; CLOUD.err = errMsg(e); if (!silent) e.relogin ? toastAction(CLOUD.err, 'Zaloguj', () => CLOUD.login()) : toast('Synchronizacja: ' + CLOUD.err); }
  cloudBadge();
}
async function cloudLoad(prevLocal) {
  if (!CLOUD.on) return; CLOUD.ready = false;
  try {
    if (!CLOUD.acct) await CLOUD.init();
    let f = null; try { f = await CLOUD.read(); } catch (e) { if (!e.notFound) throw e; }
    if (!f) { CLOUD.ready = true; CLOUD.etag = ''; await cloudWrite(); toast('Utworzono plik danych w OneDrive: ' + cloudFile()); return; }
    CLOUD.etag = f.etag; CLOUD.ready = true; CLOUD.ok = true; CLOUD.last = Date.now();
    const merged = prevLocal ? dataMerge(prevLocal, f.data) : f.data;
    dataApply(merged); render();
    if (JSON.stringify(merged, noM) !== JSON.stringify(f.data, noM)) cloudSchedule();
  } catch (e) { CLOUD.ok = false; CLOUD.err = errMsg(e); CLOUD.ready = false;
    toastAction('Synchronizacja z OneDrive: ' + CLOUD.err, e.relogin ? 'Zaloguj' : 'Ponów', () => e.relogin ? CLOUD.login() : cloudLoad(LS.get('tm:data:' + S.user.email, null))); }
  cloudBadge();
}
function cloudBadge() {
  const b = $('#syncDot'); if (!b) return; const on = CLOUD.on ? CLOUD.ready && CLOUD.ok !== false : DISK.ok !== false;
  b.className = 'sync-dot ' + (on ? 'ok' : 'bad'); b.title = CLOUD.on ? (CLOUD.ok === false ? 'Synchronizacja: ' + CLOUD.err : 'Zsynchronizowano z OneDrive' + (CLOUD.last ? ' · ' + ago(CLOUD.last) : '')) : (DISK.ok === false ? DISK.err : 'Zapisano w OneDrive (plik)');
}
setInterval(() => { if (CLOUD.on && document.visibilityState === 'visible') cloudPull(); }, 30000);
document.addEventListener('visibilitychange', () => { if (document.visibilityState !== 'visible') return; if (CLOUD.on) cloudPull(); else diskPull({conflicts: true}).catch(() => { }); });
window.addEventListener('online', () => { if (CLOUD.on && CLOUD.ready) cloudWrite(); });

/* ekran logowania w wersji na telefon (GitHub Pages) */
function cloudLoginScreen(msg) {
  const c = cloudCfg(), box = $('#gbtn'), hint = $('#loginHint');
  $('#loginForm').hidden = true; $('#login .or').hidden = true;
  if (!c.clientId) {
    box.innerHTML = H(`<form id="cloudCfgF" class="grid" style="gap:10px;text-align:left;width:100%">
      <p class="muted" style="margin:0;font-size:13px">Najprościej: na laptopie otwórz Profil → Dane → „Telefon” i wyślij sobie link konfiguracyjny. Albo wpisz dane ręcznie:</p>
      <label class="f">Identyfikator aplikacji (Client ID)<input class="input" name="clientId" required placeholder="00000000-0000-0000-0000-000000000000"></label>
      <label class="f">Identyfikator katalogu (Tenant ID)<input class="input" name="tenant" placeholder="ID dzierżawy InfoConsulting"></label>
      <label class="f">Folder danych w OneDrive<input class="input" name="path" value="${esc(c.path)}"></label>
      <button class="btn primary" style="justify-content:center;padding:11px">Dalej</button></form>`);
    $('#cloudCfgF').onsubmit = e => { e.preventDefault(); const f = e.target; LS.set('tm:cloud', {clientId: f.clientId.value.trim(), tenant: f.tenant.value.trim(), path: f.path.value.trim()}); cloudBoot(null); };
    hint.innerHTML = 'Wersja z synchronizacją: dane z pliku w Twoim firmowym OneDrive (ten sam co na laptopie).';
    return;
  }
  box.innerHTML = H(`<button class="btn primary" type="button" id="msLogin" style="width:300px;justify-content:center;padding:11px;gap:10px"><svg width="16" height="16" viewBox="0 0 21 21"><path fill="#F25022" d="M0 0h10v10H0z"/><path fill="#7FBA00" d="M11 0h10v10H11z"/><path fill="#00A4EF" d="M0 11h10v10H0z"/><path fill="#FFB900" d="M11 11h10v10H11z"/></svg>Zaloguj kontem Microsoft</button>
    <button class="btn ghost sm" type="button" id="cloudReset" style="margin-top:10px">Zmień konfigurację</button>`);
  $('#msLogin').onclick = async () => { try { await CLOUD.init(); await CLOUD.login(); } catch (e) { hint.textContent = errMsg(e); } };
  $('#cloudReset').onclick = () => { LS.del('tm:cloud'); CLOUD.app = null; cloudLoginScreen(); };
  hint.innerHTML = msg ? esc(msg) : 'Zaloguj się kontem InfoConsulting – TM odczyta i zapisze plik danych w Twoim OneDrive (ten sam co na laptopie).';
}
/* po zalogowaniu: wybór pliku danych w folderze Dane */
async function cloudPick() {
  const box = $('#gbtn'), hint = $('#loginHint'); $('#app').classList.add('hidden'); $('#login').classList.remove('hidden'); $('#loginForm').hidden = true; $('#login .or').hidden = true;
  box.innerHTML = H('<p class="muted">Szukam pliku danych w OneDrive…</p>');
  let files;
  try { files = await CLOUD.list(); }
  catch (e) { return cloudLoginScreen(e.notFound ? `Nie znaleziono folderu „${cloudCfg().path}” w OneDrive konta ${CLOUD.acct?.username}. Sprawdź folder (Zmień konfigurację).` : errMsg(e)); }
  const go = name => { const key = name.replace(/^TM-/i, '').replace(/\.json$/i, ''); login({email: key, name: CLOUD.acct?.name || key, provider: 'ms', file: name, ms: CLOUD.acct?.username}); };
  if (files.length === 1) return go(files[0].name);
  const mine = 'TM-' + String(CLOUD.acct?.username || 'uzytkownik').toLowerCase() + '.json';
  box.innerHTML = H(`<div style="text-align:left;width:100%"><p style="margin:0 0 8px">Wybierz plik danych TM:</p>${files.map(f => `<button class="btn" type="button" data-f="${esc(f.name)}" style="width:100%;justify-content:space-between;margin-bottom:6px"><span>${esc(f.name)}</span><span class="muted num">${Math.round(f.size / 1024)} KB</span></button>`).join('')}
    ${files.length ? '' : `<p class="muted" style="font-size:13px">W folderze nie ma pliku TM. Możesz utworzyć nowy.</p><button class="btn primary" type="button" data-f="${esc(mine)}" style="width:100%;justify-content:center">Utwórz ${esc(mine)}</button>`}</div>`);
  box.querySelectorAll('[data-f]').forEach(b => b.onclick = () => go(b.dataset.f));
  hint.textContent = 'Zalogowano jako ' + (CLOUD.acct?.username || '');
}
async function cloudBoot(u) {
  cloudSetupFromHash();
  if (!cloudCfg().clientId) return cloudLoginScreen();
  try { await CLOUD.init(); }
  catch (e) { if (u && u.provider === 'ms') { login(u); CLOUD.ok = false; CLOUD.err = 'Offline – pracujesz na kopii z telefonu'; cloudBadge(); return; } return cloudLoginScreen(errMsg(e)); }
  if (!CLOUD.acct) return u && u.provider === 'ms' ? (login(u), toastAction('Zaloguj się kontem Microsoft, aby synchronizować dane', 'Zaloguj', () => CLOUD.login())) : cloudLoginScreen(CLOUD.err);
  if (u && u.provider === 'ms' && u.file) return login(u);
  return cloudPick();
}
/* laptop: karta „Telefon” – link konfiguracyjny i adres przekierowania */
function phoneCard() {
  const st = SET(), url = String(st.phoneUrl || '').trim().replace(/\/?(index\.html)?$/, '/'), ok = /^https:\/\/.+/.test(String(st.phoneUrl || ''));
  const cfg = {clientId: String(st.phoneClient || '').trim(), tenant: String(st.phoneTenant || '').trim(), path: String(st.phonePath || CLOUD_DEF_PATH).trim()};
  const link = ok && cfg.clientId ? `${url}index.html#tmcloud=${encodeURIComponent(btoa(unescape(encodeURIComponent(JSON.stringify(cfg)))))}` : '';
  return `<form data-form="settings" class="card"><h3>⟨i:monitor⟩ Telefon – TM z synchronizacją</h3>
    <p class="muted" style="margin-top:0">Wersja TM na GitHub Pages otwiera na telefonie <b>ten sam plik danych</b> z OneDrive (logowanie kontem Microsoft). Zmiany z laptopa i telefonu scalają się automatycznie.</p>
    <label class="f">Adres aplikacji (GitHub Pages)<input class="input" name="phoneUrl" value="${esc(st.phoneUrl || '')}" placeholder="https://twoj-login.github.io/tm/"></label>
    <div class="grid g2"><label class="f">Identyfikator aplikacji (Client ID)<input class="input" name="phoneClient" value="${esc(cfg.clientId)}"></label><label class="f">Identyfikator katalogu (Tenant ID)<input class="input" name="phoneTenant" value="${esc(cfg.tenant)}"></label></div>
    <label class="f">Folder danych w OneDrive (od głównego folderu OneDrive)<input class="input" name="phonePath" value="${esc(cfg.path)}"></label>
    ${ok ? `<p class="hint" style="margin:8px 0 0">Adres przekierowania do wpisania w Entra (platforma „Aplikacja jednostronicowa”): <code>${esc(url)}</code></p>` : ''}
    <div class="row" style="justify-content:flex-end;gap:6px;margin-top:12px;flex-wrap:wrap">${link ? `<button type="button" class="btn" data-action="phone-copy" data-l="${esc(link)}">⟨i:external⟩ Kopiuj link</button><button type="button" class="btn" data-action="phone-mail" data-l="${esc(link)}">⟨i:mail⟩ Wyślij link na telefon</button>` : ''}<button class="btn primary">Zapisz</button></div></form>`;
}
function cloudStatusHtml() {
  return `<div class="row" style="gap:8px;flex-wrap:wrap">${CLOUD.ok === false ? `<span class="tag" style="background:#3a1d22;color:#ff9aa8">błąd</span> <span class="muted">${esc(CLOUD.err)}</span>` : CLOUD.ready ? `<span class="tag ok-tag">zsynchronizowane</span> <span class="muted">${CLOUD.last ? ago(CLOUD.last) : ''}</span>` : '<span class="muted">łączę z OneDrive…</span>'}</div>
    <div class="muted" style="font-size:12px;margin-top:6px;overflow-wrap:anywhere">OneDrive (${esc(CLOUD.acct?.username || S.user.ms || '')}): ${esc(cloudCfg().path)}/${esc(cloudFile())}</div>
    <div class="row" style="margin-top:10px;gap:6px"><button class="btn sm" data-action="cloud-sync">⟨i:refresh⟩ Synchronizuj teraz</button><button class="btn sm" data-action="cloud-relogin">Zaloguj ponownie</button></div>`;
}

const inCtx = x => S.ctx === 'all' || x.ctx === S.ctx;
const ctxLabel = () => ({all: 'wszystko: prywatne + firmowe', priv: 'tylko prywatne', job: 'tylko firmowe'}[S.ctx]);
const ctxTag = c => `<span class="tag ${c}">${c === 'job' ? 'JOB' : 'PRIV'}</span>`;

const PEOPLE = [
  {name: 'Anna Nowak', email: 'anna.nowak@infoconsulting.pl'},
  {name: 'Piotr Zieliński', email: 'piotr.zielinski@infoconsulting.pl'},
  {name: 'Katarzyna Wiśniewska', email: 'katarzyna.wisniewska@infoconsulting.pl'},
  {name: 'Tomasz Lewandowski', email: 'tomasz.lewandowski@infoconsulting.pl'},
  {name: 'Jan Kowalczyk (klient)', email: 'j.kowalczyk@klient.pl'}
];

/* ================= dane przykładowe ================= */
function seed() {
  const t = today(), o = n => iso(add(t, n));
  const goals = [
    ['Wdrożenie IFS Cloud – klient Alfa', 'job', -45, 60, 'Go-live modułów Finanse i Logistyka'],
    ['Standard dokumentacji HLS/DLS', 'job', -20, 30, 'Szablony + 5 przykładowych specyfikacji'],
    ['Certyfikacja IFS Cloud Technical', 'job', -10, 120, 'Szkolenia i egzamin'],
    ['Półmaraton < 2h', 'priv', -60, 150, 'Plan treningowy 4x w tygodniu'],
    ['Remont łazienki', 'priv', -15, 75, 'Ekipa, materiały, odbiór'],
    ['Hiszpański A2', 'priv', -90, 200, 'Lekcje + aplikacja codziennie']
  ].map(([title, ctx, s, d, desc]) => ({id: uid(), title, ctx, start: o(s), due: o(d), desc, progress: 0}));
  const G = i => goals[i].id;
  // [tytuł, ctx, podprojekt, start, termin, status, prio, est(h)]
  const T = [
    ['Analiza wymagań – Finanse', 'job', 0, -40, -30, 'done', 1, 24], ['Warsztaty z kluczowymi użytkownikami', 'job', 0, -30, -22, 'done', 1, 16],
    ['Konfiguracja planu kont', 'job', 0, -21, -12, 'done', 2, 20], ['Migracja danych – próba 1', 'job', 0, -12, -2, 'doing', 1, 24],
    ['Testy UAT – Logistyka', 'job', 0, 1, 14, 'todo', 1, 32], ['Szkolenie użytkowników końcowych', 'job', 0, 20, 35, 'todo', 2, 24],
    ['Cutover plan + go-live checklist', 'job', 0, 36, 58, 'todo', 1, 16],
    ['Szablon HLS v2', 'job', 1, -18, -10, 'done', 2, 8], ['Szablon DLS v2', 'job', 1, -9, 0, 'doing', 2, 8],
    ['Przykładowe HLS ×5', 'job', 1, 1, 20, 'todo', 3, 20], ['Review z zespołem', 'job', 1, 21, 28, 'todo', 2, 4],
    ['Kurs: IFS Developer Studio', 'job', 2, -8, 10, 'doing', 3, 16], ['Kurs: Projections & APIs', 'job', 2, 12, 40, 'todo', 3, 16],
    ['Egzamin próbny', 'job', 2, 80, 95, 'todo', 2, 4],
    ['Raport miesięczny dla PM', 'job', null, -1, -1, 'todo', 1, 2], ['Rozliczenie delegacji', 'job', null, 0, 0, 'todo', 2, 1],
    ['Plan treningowy – blok 1', 'priv', 3, -55, -25, 'done', 2, 10], ['Plan treningowy – blok 2', 'priv', 3, -24, 6, 'doing', 2, 12],
    ['Bieg kontrolny 15 km', 'priv', 3, 7, 7, 'todo', 1, 2], ['Zapisy na półmaraton', 'priv', 3, 0, 3, 'todo', 1, 0.5],
    ['Wybór płytek i armatury', 'priv', 4, -14, -5, 'done', 2, 6], ['Umowa z ekipą remontową', 'priv', 4, -4, 2, 'doing', 1, 3],
    ['Zakup materiałów', 'priv', 4, 3, 12, 'todo', 2, 5], ['Odbiór prac', 'priv', 4, 60, 70, 'todo', 2, 2],
    ['Lekcje hiszpańskiego – listopad', 'priv', 5, -16, 13, 'doing', 3, 8], ['Test poziomujący A1', 'priv', 5, -40, -35, 'done', 3, 2],
    ['Przegląd auta', 'priv', null, 4, 4, 'todo', 2, 2], ['Kupić prezenty świąteczne', 'priv', null, 10, 30, 'todo', 3, 4]
  ];
  const tasks = T.map(([title, ctx, g, s, d, status, prio, est], i) => ({
    id: uid(), title, ctx, goalId: g == null ? null : G(g), projectId: null, start: o(s), due: o(d), status, prio, est, notes: '',
    createdAt: o(s - 3 - (i % 4)), doneAt: status === 'done' ? o(d + ((i % 3) - 1)) : null
  }));
  // projekty (PRIV/JOB) – podprojekty i zadania mogą być do nich przypięte
  const projects = [
    ['Alfa – wdrożenie IFS Cloud', 'job', 'active', -50, 75, 'Wdrożenie Finanse + Logistyka u klienta Alfa', 'Klient Alfa'],
    ['Kompetencje IFS i standardy IC', 'job', 'active', -25, 130, 'Standardy dokumentacji HLS/DLS i certyfikacja', 'InfoConsulting'],
    ['Dom – remont łazienki', 'priv', 'active', -20, 80, 'Remont, materiały, ekipa', ''],
    ['Zdrowie i sport', 'priv', 'active', -60, 160, 'Bieganie, basen, regeneracja', ''],
    ['Rozwój osobisty', 'priv', 'planned', -90, 210, 'Języki i kursy', '']
  ].map(([title, ctx, status, s, d, desc, client]) => ({id: uid(), title, ctx, status, start: o(s), due: o(d), desc, client}));
  [0, 1, 1, 3, 2, 4].forEach((p, i) => goals[i].projectId = projects[p].id);
  tasks.forEach(x => { if (x.title === 'Raport miesięczny dla PM') x.projectId = projects[0].id; });
  const y = t.getFullYear();
  const vacations = [
    {id: uid(), from: `${y}-12-28`, to: `${y + 1}-01-02`, note: 'Urlop świąteczny'},
    {id: uid(), from: `${y + 1}-02-16`, to: `${y + 1}-02-20`, note: 'Ferie – narty'}
  ];
  return {v: 2, notes: [], timeEntries: [], templates: seedTemplates(projects, goals), projects, goals, tasks, vacations, meetings: [], demoDate: '', settings: {leaveDays: 26}};
}

/* migracja danych z wersji bez projektów (zachowuje dane użytkownika) */
function migrate(d) {
  d.notes ||= [];
  d.timeEntries ||= [];
  if (!d.projTemplates) d.projTemplates = seedProjTemplates();
  if (!d.tripTemplates) d.tripTemplates = seedTripTemplates();
  if (!d.templates) d.templates = seedTemplates(d.projects, d.goals);
  if (!d.projects) {
    const sd = seed(), gmap = {}, tmap = {};
    sd.goals.forEach(g => gmap[g.title] = g.projectId);
    sd.tasks.forEach(t => t.projectId && (tmap[t.title] = t.projectId));
    const isDemo = d.goals.some(g => gmap[g.title]);
    d.projects = isDemo ? sd.projects : [];
    d.goals.forEach(g => g.projectId = isDemo ? (gmap[g.title] || null) : null);
    d.tasks.forEach(t => t.projectId = isDemo ? (tmap[t.title] || null) : null);
    d.v = 2;
  }
  return d;
}
function demoMeetings() {
  const out = [], t0 = sow(today()), team = PEOPLE.slice(0, 3).map(p => p.email);
  const pj = q => ((S.data && S.data.projects) || []).find(p => p.title.includes(q))?.id || null;
  const PJ = {'Daily': pj('Alfa'), 'Warsztat': pj('Alfa'), 'Przegląd HLS': pj('Kompetencje'), 'Steering': pj('Alfa'), 'Konsultacja DLS': pj('Kompetencje'), 'Trening': pj('Zdrowie'), 'Basen': pj('Zdrowie')};
  const hidden = new Set((S.data && S.data.hiddenDemo) || []);
  const M = (d, s, e, title, source, ctx, att = [], loc = '') => hidden.has(`${title}|${d}T${s}`) || out.push({
    id: 'demo-' + uid(), title, start: `${d}T${s}`, end: `${d}T${e}`, source, ctx, attendees: att, location: loc, demo: true,
    projectId: Object.entries(PJ).find(([k]) => title.startsWith(k))?.[1] || null,
    link: source === 'teams' ? 'https://teams.microsoft.com/' : source === 'google' ? '' : ''
  });
  for (let i = -7; i < 28; i++) {
    const d = add(t0, i), ds = iso(d);
    if (holName(d)) continue;
    if (!isWE(d) && !vacOf(d)) {
      M(ds, '09:00', '09:15', 'Daily stand-up – projekt Alfa', 'teams', 'job', team);
      const k = (d.getDate() + d.getMonth()) % 5;
      if (k === 0) M(ds, '11:00', '12:30', 'Warsztat analityczny – moduł Finanse', 'teams', 'job', [...team, 'j.kowalczyk@klient.pl']);
      if (k === 1) M(ds, '13:00', '14:30', 'Przegląd HLS z klientem', 'outlook', 'job', ['j.kowalczyk@klient.pl'], 'Sala Kraków');
      if (k === 2) M(ds, '10:00', '10:45', '1:1 z kierownikiem', 'teams', 'job', [PEOPLE[3].email]);
      if (k === 3) M(ds, '14:00', '15:00', 'Steering Committee', 'outlook', 'job', [PEOPLE[3].email, 'j.kowalczyk@klient.pl'], 'Biuro klienta');
      if (k === 4) M(ds, '15:30', '16:15', 'Konsultacja DLS – integracje', 'teams', 'job', [PEOPLE[1].email]);
    }
    if (d.getDay() === 2 || d.getDay() === 4) M(ds, '18:30', '19:30', 'Trening biegowy', 'google', 'priv');
    if (d.getDay() === 6) M(ds, '10:00', '11:30', 'Basen z rodziną', 'google', 'priv');
  }
  M(iso(add(today(), 2)), '17:00', '17:45', 'Dentysta', 'google', 'priv', [], 'ul. Długa 5');
  return out;
}
function refreshDemoMeetings() {
  if (!(S.data.settings && S.data.settings.demoOn)) { if (S.data.meetings.some(m => m.demo)) S.data.meetings = S.data.meetings.filter(m => !m.demo); return; } // spotkania przykładowe wyłączone (domyślnie)
  if (S.data.demoDate === TODAY() && S.data.demoVer === 3) return;
  S.data.demoVer = 3;
  S.data.meetings = S.data.meetings.filter(m => !m.demo).concat(demoMeetings().filter(m => !(S.data.integrations?.outlook?.real && (m.source === 'outlook' || m.source === 'teams'))));
  S.data.demoDate = TODAY(); save();
}

/* ================= logika domenowa ================= */
const tasksC = () => S.data.tasks.filter(inCtx);
const goalsC = () => S.data.goals.filter(inCtx);
const meetingsC = () => S.data.meetings.filter(inCtx);
const goalById = id => S.data.goals.find(g => g.id === id);
const projById = id => id && (S.data.projects || []).find(p => p.id === id);
const projectsC = () => (S.data.projects || []).filter(inCtx);
/* projekt zadania: z podprojektu (jeśli podprojekt jest w projekcie), inaczej przypisany bezpośrednio */
const projOf = t => { const g = t.goalId && goalById(t.goalId); return (g && g.projectId) || t.projectId || null; };
const projTasks = p => S.data.tasks.filter(t => projOf(t) === p.id);
const projGoals = p => S.data.goals.filter(g => g.projectId === p.id);
const PSTAT = {planned: ['Planowany', 'warn'], active: ['Aktywny', 'ok'], onhold: ['Wstrzymany', 'bad'], done: ['Zakończony', 'muted']};
function projectProgress(p, at) {
  const ts = projTasks(p);
  if (!ts.length) { const gs = projGoals(p); return gs.length ? Math.round(gs.reduce((a, g) => a + goalProgress(g, at), 0) / gs.length) : (p.status === 'done' ? 100 : 0); }
  const w = t => t.est || 1, tot = ts.reduce((a, t) => a + w(t), 0);
  return Math.round(100 * ts.filter(t => t.status === 'done' && (!at || (t.doneAt && t.doneAt <= at))).reduce((a, t) => a + w(t), 0) / tot);
}
const projTag = id => { const p = projById(id); return p ? `<span class="tag src" title="Projekt: ${esc(p.title)}">⟨i:folder⟩ ${esc(p.title.length > 24 ? p.title.slice(0, 23) + '…' : p.title)}</span>` : ''; };
const projOptions = (sel, ctx) => `<option value="">— bez projektu —</option>` + (S.data.projects || []).filter(p => !ctx || p.ctx === ctx || p.id === sel)
  .map(p => `<option value="${p.id}" ${sel === p.id ? 'selected' : ''}>[${p.ctx.toUpperCase()}] ${esc(p.title)}${p.status === 'done' ? ' (zakończony)' : ''}</option>`).join('');
const isLate = t => t.status !== 'done' && t.due && t.due < TODAY();
function goalProgress(g, at) {
  const ts = S.data.tasks.filter(t => t.goalId === g.id);
  if (!ts.length) return g.progress || 0;
  const w = t => t.est || 1, tot = ts.reduce((a, t) => a + w(t), 0);
  const done = ts.filter(t => t.status === 'done' && (!at || (t.doneAt && t.doneAt <= at))).reduce((a, t) => a + w(t), 0);
  return Math.round(100 * done / tot);
}
/* plan wg czasu: 0% przed startem podprojektu, liniowo do 100% w dniu terminu */
function goalExpected(g) {
  if (!g.due) return 0;
  const s = g.start || projById(g.projectId)?.start || g.due, td = TODAY();
  if (td < s) return 0;
  const len = diffD(s, g.due);
  if (!(len > 0)) return td >= g.due ? 100 : 0;
  return Math.round(100 * clamp(diffD(s, td) / len, 0, 1));
}
/* status podprojektu: osiągnięty / po terminie / zaplanowany (jeszcze nie wystartował) / opóźniony / ryzyko / w planie */
function goalStatus(g, p = goalProgress(g), e = goalExpected(g)) {
  if (p >= 100) return ['osiągnięty', 'ok'];
  if (g.due && g.due < TODAY()) return ['po terminie', 'bad'];
  if (g.start && g.start > TODAY()) return ['zaplanowany', 'muted'];
  return p + 10 < e ? ['opóźniony', 'bad'] : p >= e ? ['w planie', 'ok'] : ['ryzyko', 'warn'];
}
/* kolejność podprojektów: w realizacji / opóźnione najpierw, dalej rosnąco wg terminu; osiągnięte na końcu */
const goalSort = (a, b) => ((goalProgress(a) >= 100) - (goalProgress(b) >= 100)) || (a.due || '9999').localeCompare(b.due || '9999') || (a.start || '').localeCompare(b.start || '') || a.title.localeCompare(b.title, 'pl');
/* rozkład estymacji zadania na dni (job: dni robocze, priv: wszystkie dni) */
/* obciążenie zadania na dni: pozostała praca (estymacja − zaraportowany czas; bez estymacji = domyślna z ustawień),
   JOB tylko w dni robocze, PRIV we wszystkie dni; zadania zaległe obciążają najbliższe dni od dziś */
const defEst = () => { const v = SET().defEst; return v == null || v === '' ? 1 : +v; };
const taskRemain = t => { const est = t.est > 0 ? +t.est : defEst(); return Math.max(0, est - taskSpent(t.id) / 60); };
function taskLoad(t) {
  const res = {}; if (t.status === 'done') return res;
  const remain = taskRemain(t); if (!remain) return res;
  const okDay = d => t.ctx === 'priv' || isWork(d), td = TODAY();
  const days = [];
  if (t.due && t.due < td) { // zaległe: na najbliższe dni, po dziennym limicie kontekstu
    const per = Math.max(1, capacityOfDay(today(), t.ctx) || (t.ctx === 'priv' ? CFG.HOURS_PRIV_WORKDAY : CFG.HOURS_JOB_PER_DAY));
    for (let d = today(), i = 0; days.length < Math.ceil(remain / per) && i < 60; d = add(d, 1), i++) if (okDay(d)) days.push(iso(d));
  } else {
    const s = t.start && t.start > td ? t.start : td, e = t.due && t.due >= s ? t.due : s;
    for (let d = P(s); iso(d) <= e; d = add(d, 1)) if (okDay(d)) days.push(iso(d));
    if (!days.length) days.push(e);
  }
  days.forEach(d => res[d] = remain / days.length);
  return res;
}
/* spotkania dnia bez podwójnego liczenia nakładających się terminów (h) */
function meetingHoursByDay(list) {
  const by = {}; list.forEach(m => (by[m.start.slice(0, 10)] ||= []).push([toMin(m.start), Math.max(toMin(m.start), toMin(m.end))]));
  const out = {}; for (const [d, iv] of Object.entries(by)) { iv.sort((a, b) => a[0] - b[0]); let tot = 0, cs = -1, ce = -1;
    iv.forEach(([a, b]) => { if (a > ce) { if (ce > cs) tot += ce - cs; cs = a; ce = b; } else ce = Math.max(ce, b); }); if (ce > cs) tot += ce - cs; out[d] = tot / 60; }
  return out;
}
function capacityOfDay(d, ctx, raw) { // raw = bez delegacji (np. target realizacji 8h)
  const ty = dayType(d), tr = !raw && S.data && tripAt(iso(d)); let h = 0;
  if (ctx !== 'priv') h += (ty === 'work' ? CFG.HOURS_JOB_PER_DAY * (halfVac(d) ? .5 : 1) : 0) * (tr ? tripCap(tr, 'job') / 100 : 1);
  if (ctx !== 'job') h += (ty === 'work' ? CFG.HOURS_PRIV_WORKDAY : CFG.HOURS_PRIV_WEEKEND) * (tr ? tripCap(tr, 'priv') / 100 : 1);
  return h;
}
/* obciążenie zadaniami (bieżący kontekst): dzień → h, w tym część z zadań zaległych */
function capLoads() {
  const load = {}, late = {}, td = TODAY();
  tasksC().forEach(t => { const isLateT = t.status !== 'done' && t.due && t.due < td; for (const [d, h] of Object.entries(taskLoad(t))) { load[d] = (load[d] || 0) + h; if (isLateT) late[d] = (late[d] || 0) + h; } });
  return {load, late};
}
/* obciążenie konkretnego dnia (do ostrzeżenia przy wyborze terminu) */
function dayLoadInfo(ds, ctx, exceptId) {
  const d = P(ds), cap = capacityOfDay(d, ctx);
  let tasks = 0; S.data.tasks.filter(t => t.ctx === ctx && t.id !== exceptId).forEach(t => tasks += taskLoad(t)[ds] || 0);
  const mtg = meetingHoursByDay(S.data.meetings.filter(m => m.ctx === ctx && m.start.slice(0, 10) === ds))[ds] || 0;
  return {cap, tasks, mtg, tot: tasks + mtg, type: dayType(d), trip: tripAt(ds)};
}
function capacityRange(from, n) { // from: Date, n dni
  const {load, late} = capLoads(), mt = meetingHoursByDay(meetingsC());
  let cap = 0, tasks = 0, mtg = 0, lateH = 0, work = 0, hol = [], vac = 0, we = 0;
  for (let i = 0; i < n; i++) {
    const d = add(from, i), s = iso(d), ty = dayType(d);
    cap += capacityOfDay(d, S.ctx); tasks += load[s] || 0; lateH += late[s] || 0; mtg += mt[s] || 0;
    if (ty === 'work') work += halfVac(d) ? .5 : 1; if (ty === 'we') we++;
    if (ty === 'hol') hol.push(holName(d)); if ((ty === 'vac' && !isWE(d)) || halfVac(d)) vac += halfVac(d) ? .5 : 1;
  }
  return {cap, tasks, mtg, late: lateH, work, hol, vac, we, util: cap ? (tasks + mtg) / cap : 0};
}
function vacationWorkdays(v, year) {
  let n = 0; for (let d = P(v.from); iso(d) <= v.to; d = add(d, 1)) if ((!year || d.getFullYear() === year) && !isWE(d) && !holName(d)) n += v.half ? .5 : 1;
  return n;
}

/* ================= szybki kreator – parser ================= */
const DOWMAP = {pn: 1, pon: 1, poniedzialek: 1, wt: 2, wtorek: 2, sr: 3, sroda: 3, cz: 4, czw: 4, czwartek: 4, pt: 5, piatek: 5, sb: 6, sob: 6, sobota: 6, nd: 0, ndz: 0, niedziela: 0};
function parseQuick(txt) {
  const r = {ctx: null, prio: 2, est: null, dates: [], goalId: null, projectId: null, title: ''};
  const rest = [], t = today();
  for (const w of txt.trim().split(/\s+/).filter(Boolean)) {
    const lw = norm(w); let m;
    if (/^#(priv|prywatne|p|dom|home)$/.test(lw)) r.ctx = 'priv';
    else if (/^#(job|praca|j|work|firma)$/.test(lw)) r.ctx = 'job';
    else if ((m = lw.match(/^!([1-3])$/))) r.prio = +m[1];
    else if ((m = lw.match(/^(\d+(?:[.,]\d+)?)(h|godz|m|min)$/))) { const v = parseFloat(m[1].replace(',', '.')); r.est = m[2][0] === 'm' ? r1(v / 60) : v; }
    else if (lw === 'dzis' || lw === 'today') r.dates.push(t);
    else if (lw === 'jutro' || lw === 'tomorrow') r.dates.push(add(t, 1));
    else if (lw === 'pojutrze') r.dates.push(add(t, 2));
    else if (DOWMAP[lw] !== undefined) { let n = (DOWMAP[lw] - t.getDay() + 7) % 7 || 7; r.dates.push(add(t, n)); }
    else if ((m = lw.match(/^(\d{1,2})[./](\d{1,2})(?:[./](\d{2,4}))?$/))) {
      let y = m[3] ? (+m[3] < 100 ? 2000 + +m[3] : +m[3]) : t.getFullYear(); let d = new Date(y, +m[2] - 1, +m[1]);
      if (!m[3] && d < t) d = new Date(y + 1, +m[2] - 1, +m[1]); r.dates.push(d);
    }
    else if ((m = lw.match(/^\+(\d+)(d|t|w|tyg)?$/))) r.dates.push(add(t, +m[1] * (m[2] && m[2] !== 'd' ? 7 : 1)));
    else if (w.length > 1 && w[0] === '^') {
      const q = norm(w.slice(1)); const p = (S.data.projects || []).find(p => norm(p.title).includes(q));
      if (p) { r.projectId = p.id; r.ctx = p.ctx; } else rest.push(w);
    }
    else if (w.length > 1 && w[0] === '@') {
      const q = norm(w.slice(1)); const g = S.data.goals.find(g => norm(g.title).includes(q));
      if (g) { r.goalId = g.id; r.ctx = r.ctx || g.ctx; } else rest.push(w);
    }
    else rest.push(w);
  }
  r.title = rest.join(' ');
  if (r.goalId && goalById(r.goalId).projectId) r.projectId = goalById(r.goalId).projectId;
  if (!r.ctx) { const h = new Date().getHours(); r.ctx = S.ctx !== 'all' ? S.ctx : (!isWE(t) && h >= 7 && h < 17 ? 'job' : 'priv'); }
  const due = r.dates.length ? r.dates[r.dates.length - 1] : t;
  let start = r.dates.length > 1 ? r.dates[0] : due;
  if (r.dates.length < 2 && r.est > CFG.HOURS_JOB_PER_DAY) { let need = Math.ceil(r.est / CFG.HOURS_JOB_PER_DAY) - 1; while (need > 0) { start = add(start, -1); if (isWork(start)) need--; } if (start < t) start = t; }
  r.due = iso(due); r.start = iso(start);
  return r;
}

/* ================= UI: modal, toast ================= */
function modal(html, onMount) {
  $('#modalRoot').innerHTML = H(`<div class="modal-bg" data-action="modal-bg"><div class="modal" role="dialog">${html}</div></div>`);
  trkGuard($('#modalRoot')); onMount && onMount($('#modalRoot .modal'));
}
const closeModal = () => $('#modalRoot').innerHTML = '';
let toastT;
function toastAction(msg, label, fn, ms = 9000) {
  $('.toast')?.remove(); const el = document.createElement('div'); el.className = 'toast';
  el.innerHTML = H(`<span>${esc(msg)}</span> <button class="toast-btn">${esc(label)}</button>`);
  el.querySelector('button').onclick = () => { el.remove(); fn(); }; document.body.append(el);
  clearTimeout(toastT); toastT = setTimeout(() => el.remove(), ms);
}
function confirmDlg(title, html, okLabel, onOk) {
  modal(`<h3>${title}</h3><div style="margin-bottom:16px">${html}</div><div class="row" style="justify-content:flex-end"><button class="btn" data-action="modal-close">Anuluj</button><button class="btn primary danger-bg" id="dlgOk">${okLabel}</button></div>`,
    m => { const b = $('#dlgOk', m); b.focus(); b.onclick = () => { closeModal(); onOk(); }; });
}
/* usuwanie zadań z możliwością cofnięcia */
function deleteTasks(ids, what) {
  ids = new Set(ids); if (!ids.size) return;
  const removed = S.data.tasks.map((t, i) => [i, t]).filter(([, t]) => ids.has(t.id));
  S.data.tasks = S.data.tasks.filter(t => !ids.has(t.id));
  ids.forEach(id => S.tasks.sel.delete(id));
  save(); render();
  toastAction(`Usunięto ${what || (removed.length === 1 ? (pvCal() ? 'zadanie' : '„' + removed[0][1].title + '”') : removed.length + ' zadań')}`, '⟨i:undo⟩ Cofnij', () => {
    removed.forEach(([i, t]) => S.data.tasks.splice(Math.min(i, S.data.tasks.length), 0, t)); save(); render(); toast(`Przywrócono ${removed.length} ${removed.length === 1 ? 'zadanie' : 'zadań'}`);
  });
}
const plural = n => n === 1 ? 'zadanie' : (n % 10 >= 2 && n % 10 <= 4 && (n % 100 < 10 || n % 100 >= 20)) ? 'zadania' : 'zadań';
/* zadania widoczne w widoku Zadania (te same filtry co lista) */
function visibleTasks() {
  const F = S.tasks; let ts = tasksC();
  if (F.q) ts = ts.filter(t => norm(t.title + ' ' + (t.ic || '') + ' ' + pName(personById(t.assigneeId)) + ' ' + pName(personById(t.requesterId))).includes(norm(F.q)));
  if (F.goal) ts = ts.filter(t => (F.goal === '-' ? !t.goalId : t.goalId === F.goal));
  if (F.proj) ts = ts.filter(t => (F.proj === '-' ? !projOf(t) : projOf(t) === F.proj));
  if (F.mode === 'list' && F.status === 'open') ts = ts.filter(t => t.status !== 'done');
  if (F.mode === 'list' && F.status === 'done') ts = ts.filter(t => t.status === 'done');
  return ts;
}
function toast(msg) { let el = $('.toast'); if (!el) { el = document.createElement('div'); el.className = 'toast'; document.body.append(el); } el.textContent = msg; clearTimeout(toastT); toastT = setTimeout(() => el.remove(), 2600); }

/* ================= Logowanie ================= */
function showLogin() {
  $('#app').classList.add('hidden'); $('#login').classList.remove('hidden');
  if (CLOUD.on) return cloudLoginScreen();
  const hint = $('#loginHint');
  if (CFG.GOOGLE_CLIENT_ID) {
    loadScript('https://accounts.google.com/gsi/client').then(() => {
      google.accounts.id.initialize({client_id: CFG.GOOGLE_CLIENT_ID, callback: r => {
        const p = JSON.parse(decodeURIComponent(escape(atob(r.credential.split('.')[1].replace(/-/g, '+').replace(/_/g, '/')))));
        login({email: p.email, name: p.name, picture: p.picture, provider: 'google'});
      }});
      google.accounts.id.renderButton($('#gbtn'), {theme: 'outline', size: 'large', text: 'signin_with', shape: 'pill', locale: 'pl', width: 300});
    }).catch(() => hint.textContent = 'Nie udało się załadować logowania Google (offline?). Użyj e-maila.');
    hint.textContent = 'Konto osobiste oparte o Google (Gmail). Dane przechowywane per konto.';
  } else {
    $('#gbtn').innerHTML = ''; $('#login .or').hidden = true;
    hint.innerHTML = 'Laptop: dane zapisują się w pliku w Twoim OneDrive (przez start-TM.bat). Telefon: wersja na GitHub Pages z logowaniem kontem Microsoft – ten sam plik.';
  }
}
function login(u) {
  const prevLocal = LS.get('tm:data:' + u.email.trim().toLowerCase(), null);
  DISK.ready = false;
  S.user = {email: u.email.trim().toLowerCase(), name: u.name || u.email.split('@')[0], picture: u.picture || '', provider: u.provider || 'email', ...(u.file ? {file: u.file, ms: u.ms || ''} : {})};
  LS.set('tm:user', S.user);
  S.data = migrate(LS.get('tm:data:' + S.user.email, null) || seed()); snapInit(S.data);
  refreshDemoMeetings(); save();
  S.ctx = LS.get('tm:ctx', 'all');
  S.privacy = true; // domyślnie dane ukryte – pokazanie to świadoma decyzja użytkownika (nie jest zapamiętywane)
  $('#login').classList.add('hidden'); $('#app').classList.remove('hidden');
  const v = location.hash.slice(1); if (VIEWS[v]) S.view = v;
  render();
  if (new URLSearchParams(location.search).get('quick')) quickAdd();
  applySettings();
  if (CLOUD.on) cloudLoad(prevLocal); else diskLoad(prevLocal);
  setTimeout(autoImportTick, 1500); setTimeout(() => { if (mailCfg().mode !== 'manual') mailTmSync(); }, 4000); setTimeout(notifTick, 2500);
}
function logout() { LS.del('tm:user'); S.user = null; closeModal(); try { google.accounts.id.disableAutoSelect(); } catch { } showLogin(); }

/* ================= Render główny ================= */
const VIEWS = {
  dashboard: ['⟨i:home⟩', 'Start', vDashboard],
  projects: ['⟨i:folder⟩', 'Projekty', vProjects, null, 'Lista projektów'], project: ['⟨i:folder⟩', 'Projekt', vProject, 'projects'], goals: ['⟨i:layers⟩', 'Podprojekty', vGoals, 'projects'],
  tasks: ['⟨i:check-square⟩', 'Zadania', vTasks],
  meetings: ['⟨i:video⟩', 'Spotkania', vMeetings, null, 'Lista spotkań'], integ: ['⟨i:mail⟩', 'Poczta i import', vInteg, 'meetings'],
  calendar: ['⟨i:calendar⟩', 'Kalendarz', vCalendar],
  people: ['⟨i:users⟩', 'Osoby', vPeople, null, 'Lista osób'],
  avail: ['⟨i:activity⟩', 'Dostępność', vAvail, null, 'Przegląd'], capacity: ['⟨i:gauge⟩', 'Capacity', vCapacity, 'avail'], plan: ['⟨i:list⟩', 'Plan (terminy)', vPlan, 'avail'], sched: ['⟨i:schedule⟩', 'Harmonogram', vSched, 'avail'],
  track: ['⟨i:timer⟩', 'Realizacja', vTrack, 'avail'], timelog: ['⟨i:list⟩', 'Ewidencja czasu', vTimeLog, 'avail'], timereport: ['⟨i:chart⟩', 'Raport realizacji', vTimeReport, 'avail'],
  gantt: ['⟨i:gantt⟩', 'Gantt', vGantt, null, null, true], notifs: ['⟨i:bell⟩', 'Powiadomienia', vNotifs, null, null, true],
  templates: ['⟨i:template⟩', 'Szablony zadań', vTemplates, null, null, true], profile: ['⟨i:user⟩', 'Profil i ustawienia', vProfile, null, null, true]
};
/* render() – zob. moduł UI 2026 */
const afterRender = {};

/* ---------- komponenty ---------- */
function taskRow(t, opts = {}) {
  const g = t.goalId && goalById(t.goalId);
  const hide = S.privacy, sel = opts.sel && S.tasks.selMode, P_ = hide ? pv : (x => x);
  return `<div class="task ${t.status === 'done' ? 'done' : ''} ${sel && S.tasks.sel.has(t.id) ? 'selected' : ''}">
    ${sel ? `<input type="checkbox" class="selchk" data-action="sel-task" data-id="${t.id}" ${S.tasks.sel.has(t.id) ? 'checked' : ''} title="Zaznacz">` : ''}
    <span class="chk" data-action="toggle" data-id="${t.id}">${t.status === 'done' ? '⟨i:check⟩' : ''}</span>
    <span class="dot ${t.ctx}"></span>
    <span class="t" data-action="edit-task" data-id="${t.id}" title="${hide ? 'Zadanie (dane ukryte)' : esc(t.title)}">${P_(esc(t.title))}${t.status === 'doing' ? P_(' <span class="tag src">w toku</span>') : ''}</span>
    <span class="meta">${!opts.noTrip && t.tripId && tripById(t.tripId) ? `<span class="tag trip" data-action="trip-open" data-id="${t.tripId}" title="Przygotowanie do delegacji: ${esc(tripById(t.tripId).place)}">⟨i:plane⟩</span>` : ''}${personTags(t)}${t.ic ? icTag(t) : ''}${t.ctx === 'priv' && (t.costAct || t.costEst) ? P_(`<span class="tag src" title="Koszt rzeczywisty / szacowany">⟨i:wallet⟩ ${t.costAct ? zl(t.costAct) : '—'}${t.costEst ? ' / ' + zl(t.costEst) : ''}</span>`) : ''}${!opts.noProj ? P_(projTag(projOf(t))) : ''}${g && !opts.noGoal ? `<span class="tag src">⟨i:layers⟩ ${P_(esc(g.title.slice(0, 22)))}</span>` : ''}
      ${t.est ? `<span class="tag src">${t.est}h</span>` : ''}${P_(`<span class="tag p${t.prio}">P${t.prio}</span>`)}
      <span class="due ${isLate(t) ? 'late' : ''}">${t.due ? relDay(t.due) : ''}</span>${taskSpent(t.id) ? `<span class="tag src" title="Czas realizacji">⟨i:timer⟩ ${fmtH(taskSpent(t.id))}</span>` : ''}${noteBadge('task', t.id)}${t.status !== 'done' ? (timerOn()?.taskId === t.id ? '<span class="playbtn on" title="Realizacja w toku"><span class="tc-dot"></span></span>' : `<button class="playbtn" data-action="timer-start-task" data-id="${t.id}" title="Start realizacji (stoper)">⟨i:play⟩</button>`) : ''}
    ${cmtBadge('task', t.id)}<button class="del" data-action="del-task-row" data-id="${t.id}" title="Usuń zadanie">⟨i:trash⟩</button></span></div>${(S.cx ||= new Set()).has(cmtKey('task', t.id)) ? `<div class="cmt-sub">${cmtBlock('task', t.id)}</div>` : ''}`;
}
/* ---------- Ustawienia (profil) – sekcje dokładane przez kolejne moduły ---------- */
const SET = () => (S.data.settings ||= {});
const SETTINGS_SECTIONS = [];
const CFG0 = {...CFG};
/* ustawienia z profilu nadpisują wartości domyślne z config.js */
function applySettings() {
  const st = SET(), num = (k, d) => st[k] == null || st[k] === '' || isNaN(+st[k]) ? d : +st[k];
  CFG.HOURS_JOB_PER_DAY = num('hJob', CFG0.HOURS_JOB_PER_DAY); CFG.HOURS_PRIV_WORKDAY = num('hPrivWork', CFG0.HOURS_PRIV_WORKDAY); CFG.HOURS_PRIV_WEEKEND = num('hPrivWe', CFG0.HOURS_PRIV_WEEKEND);
  CFG.WORK_FROM = st.workFrom || '08:00'; CFG.WORK_TO = st.workTo || '16:00';
}
SETTINGS_SECTIONS.push(() => `<div class="card"><h3>⟨i:gauge⟩ Capacity i czas pracy</h3>
  <div class="grid g2"><label class="f">Godziny pracy od<input class="input" type="time" name="workFrom" value="${esc(SET().workFrom || '08:00')}"></label><label class="f">do<input class="input" type="time" name="workTo" value="${esc(SET().workTo || '16:00')}"></label>
  <label class="f">JOB – godzin w dzień roboczy<input class="input" type="number" min="0" max="16" step="0.5" name="hJob" value="${CFG.HOURS_JOB_PER_DAY}"></label>
  <label class="f">Domyślna estymacja zadania bez estymacji (h)<input class="input" type="number" min="0" max="16" step="0.25" name="defEst" value="${defEst()}"></label>
  <label class="f">PRIV – godzin w dzień roboczy (wieczór)<input class="input" type="number" min="0" max="16" step="0.5" name="hPrivWork" value="${CFG.HOURS_PRIV_WORKDAY}"></label>
  <label class="f">PRIV – godzin w weekend / święto / urlop<input class="input" type="number" min="0" max="16" step="0.5" name="hPrivWe" value="${CFG.HOURS_PRIV_WEEKEND}"></label></div>
  <p class="hint" style="margin:6px 0 0">Target dnia w „Realizacja tydzień” = godziny JOB. Godziny pracy wyznaczają czas JOB w kalendarzu i ciche godziny powiadomień.</p></div>`);
function vSettings(sec = 'general') {
  return `<form data-form="settings" class="grid settings-grid">${SETTINGS_SECTIONS.filter(f => (f.sec || 'general') === sec).map(f => f()).join('')}
    <div class="row" style="grid-column:1/-1;justify-content:flex-end"><button class="btn primary">Zapisz ustawienia</button></div></form>`;
}
SETTINGS_SECTIONS.push(() => `<div class="card"><h3>⟨i:check-square⟩ Zadania JOB · IC case</h3>
  <label class="f">Wzór linku do IC case (opcjonalnie)<input class="input" name="icUrl" value="${esc(SET().icUrl || '')}" placeholder="np. https://…/case/{id}"></label>
  <p class="hint" style="margin:6px 0 0"><code>{id}</code> = sam numer (12345), <code>{case}</code> = pełny identyfikator (IC-12345). Puste = etykieta bez linku.</p></div>`);
/* ================= USTAWIENIA APLIKACJI (zębatka w lewym górnym rogu) ================= */
function mailCfgCard() {
  const c = mailCfg(), st = SET(), mt = mailTmCfg();
  return `<form data-form="settings" class="card"><h3>⟨i:mail⟩ Zadania z maili (kategoria w Outlooku)</h3>
    <p class="muted" style="margin-top:0;font-size:13px">Mail oznaczony w klasycznym Outlooku kategorią poniżej staje się zadaniem JOB; kolejne maile z tego samego wątku dopisują się jako komentarze. Po przetworzeniu kategoria zmienia się na „${esc(c.cat)} ✓”.</p>
    <label class="row" style="gap:8px;font-size:14px;margin-bottom:8px"><input type="checkbox" name="mailOn" ${c.on ? 'checked' : ''}> <b>Twórz zadania z maili</b></label>
    <div class="grid g2"><label class="f">Kategoria (tag) maila<input class="input" name="mailCat" value="${esc(c.cat)}" placeholder="TM"></label>
    <label class="f">Kiedy sprawdzać maile<select class="input" name="mailMode"><option value="import" ${c.mode === 'import' ? 'selected' : ''}>razem z importem spotkań</option><option value="interval" ${c.mode === 'interval' ? 'selected' : ''}>co X minut</option><option value="manual" ${c.mode === 'manual' ? 'selected' : ''}>tylko ręcznie („Importuj teraz”)</option></select></label>
    <label class="f">Co ile minut (tryb „co X minut”)<input class="input" type="number" min="5" max="240" name="mailEvery" value="${c.every}"></label>
    <label class="f">Pomijaj maile starsze niż (dni, 0 = bez limitu)<input class="input" type="number" min="0" max="365" name="mailMaxAge" value="${c.maxAge}"></label>
    <label class="f">Usuwaj zadania z maili starszych niż (dni, 0 = nie usuwaj)<input class="input" type="number" min="0" max="365" name="mailPurge" value="${c.purge}"></label>
    <label class="f">Usuwaj zaimportowane spotkania starsze niż (dni, 0 = nie usuwaj)<input class="input" type="number" min="0" max="730" name="mtgPurge" value="${c.mtgPurge}"></label></div>
    <p class="hint" style="margin:6px 0 0">Nie są usuwane zadania z zaraportowanym czasem, komentarzem lub notatką ani spotkania z komentarzem lub notatką. Ostatnie sprawdzenie maili: ${mt.lastOk ? ago(mt.lastOk) : '—'}${mt.err ? ` · <span class="bbad">${esc(mt.err)}</span>` : ''}.</p>
    <div class="row" style="justify-content:flex-end;margin-top:8px"><button type="button" class="btn" data-action="mail-now">⟨i:refresh⟩ Sprawdź maile teraz</button><button class="btn primary">Zapisz</button></div></form>`;
}
function vSettingsScreen() {
  const T = S.setTab || 'general', strip = h => h.replace(/^<div class="seg caltabs"[\s\S]*?<\/div>/, '');
  const tabs = [['general', '⟨i:sliders⟩ Ogólne'], ['tpl', `⟨i:template⟩ Szablony zadań (${tplAll().length})`], ['ptpl', `⟨i:folder⟩ Szablony projektów (${ptplAll().length})`], ['dtpl', `⟨i:plane⟩ Szablony delegacji (${dtplAll().length})`], ['integ', '⟨i:plug⟩ Integracje'], ['notif', '⟨i:bell⟩ Powiadomienia']];
  const head = `<div style="margin-bottom:14px"><div class="seg settabs">${tabs.map(([k, l]) => `<button data-action="set" data-k="setTab" data-v="${k}" class="${T === k ? 'on' : ''}">${l}</button>`).join('')}</div></div>`;
  if (T === 'tpl') return head + strip(vTemplates());
  if (T === 'ptpl') return head + strip(vProjTemplates());
  if (T === 'dtpl') return head + vTripTemplates();
  if (T === 'notif') return head + vSettings('notif');
  if (T === 'integ') {
    const I = S.setInteg || 'outlook';
    return head + `<div style="margin-bottom:14px"><div class="seg"><button data-action="set" data-k="setInteg" data-v="outlook" class="${I === 'outlook' ? 'on' : ''}">${srcBadge('outlook')} Outlook</button><button data-action="set" data-k="setInteg" data-v="teams" class="${I === 'teams' ? 'on' : ''}">${srcBadge('teams')} Teams</button><button data-action="set" data-k="setInteg" data-v="google" class="${I === 'google' ? 'on' : ''}">${srcBadge('google')} Google (planowane)</button></div></div>` +
      (I === 'teams' ? vTeamsCfg() : I === 'google' ? `<div class="card"><h3>Google Calendar / Gmail</h3><p class="muted" style="margin:0">Integracja z prywatnym kalendarzem Google – kolejny etap.</p></div>` :
        `<h3 class="group-h">Spotkania – kalendarz Outlook</h3>${vOutlookCfg()}<h3 class="group-h" style="margin-top:18px">Zadania z maili</h3><div class="grid settings-grid">${mailCfgCard()}${pplCfgCard()}</div>`);
  }
  return head + vSettings('general');
}
afterRender.settings = () => { bindImportForm(); afterRender.templates?.(); };
/* IC case (zadania JOB): etykieta, opcjonalnie link wg wzoru z profilu (Ustawienia) */
const icNorm = v => { v = String(v || '').trim().toUpperCase().replace(/\s+/g, ''); return /^\d+$/.test(v) ? 'IC-' + v : v; };
const icFind = txt => { const m = String(txt || '').match(/\bIC[-_ ]?(\d{3,})\b/i); return m ? 'IC-' + m[1] : ''; };
function icTag(t) {
  if (!t.ic) return '';
  const url = (S.data.settings && S.data.settings.icUrl || '').trim(), lbl = `<span class="tag ic" title="IC case">${esc(t.ic)}</span>`;
  return url ? `<a href="${esc(url.replace('{id}', encodeURIComponent(t.ic.replace(/^IC-/i, ''))).replace('{case}', encodeURIComponent(t.ic)))}" target="_blank" rel="noopener" class="ic-link" title="Otwórz ${esc(t.ic)}">${lbl}</a>` : lbl;
}
const prioSort = (a, b) => (a.due || '9') .localeCompare(b.due || '9') || a.prio - b.prio;
function ring(p, color, size = 56) {
  const r = size / 2 - 5, c = 2 * Math.PI * r;
  return `<svg class="ring" width="${size}" height="${size}"><circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke="var(--line)" stroke-width="6"/>
  <circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke="${color}" stroke-width="6" stroke-linecap="round" stroke-dasharray="${c * p / 100} ${c}" transform="rotate(-90 ${size / 2} ${size / 2})"/>
  <text x="50%" y="54%" text-anchor="middle" dominant-baseline="middle" class="lbl-strong">${p}%</text></svg>`;
}
const ctxColor = c => c === 'priv' ? 'var(--priv)' : 'var(--job)';
const SRC = {outlook: ['Outlook', '#0F6CBD'], teams: ['Teams', '#5B5FC7'], google: ['Google Calendar', '#EA4335'], local: ['Tylko TM', '#64748B']};
const SRCB = {outlook: 'O', teams: 'T', google: 'G', local: 'TM'};
/* odznaka źródła: G = Google Calendar, O = Outlook, T = Teams (Outlook), TM = bez platformy */
const srcBadge = s => `<span class="srcb" style="background:${SRC[s]?.[1] || '#64748B'}" title="Źródło: ${SRC[s]?.[0] || s}">${SRCB[s] || 'TM'}</span>`;
const CTXI = {job: '⟨i:building⟩', priv: '⟨i:user⟩'};
const ctxIcon = c => `<span class="ctxi" title="${c === 'job' ? 'JOB – firmowe' : 'PRIV – prywatne'}">${CTXI[c] || ''}</span>`;
const PCOL = ['#4F46E5', '#0EA5E9', '#16A34A', '#DB2777', '#F59E0B', '#8B5CF6', '#14B8A6', '#E5484D', '#0F766E', '#A16207'];
const projColor = p => p ? (p.color || PCOL[Math.max(0, (S.data.projects || []).indexOf(p)) % PCOL.length]) : null;
const NOPROJ = '#94A3B8';
const mtgColor = m => projColor(projById(m.projectId)) || NOPROJ;
const addMin = (s, n) => { const d = new Date(s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10), +s.slice(11, 13), +s.slice(14, 16) + n); return `${iso(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const mtgKey = m => `${m.title}|${m.start}`;
/* tryb prywatności spotkań: dane spotkań (temat, miejsce, projekt, notatki) są rozmyte; siatka, czas i ikony zostają */
const pv = html => `<span class="pv">${html}</span>`;
/* w kalendarzu tryb prywatności obejmuje też zadania */
const pvCal = () => S.privacy && (S.view === 'calendar' || S.view === 'meetings' || S.view === 'tasks');
const eyeBtn = (sm = '') => `<button class="btn ${sm} eyebtn ${S.privacy ? 'on' : ''}" data-action="privacy" title="${S.privacy ? 'Dane spotkań i zadań ukryte – kliknij, aby pokazać' : 'Ukryj (rozmyj) dane spotkań i zadań'}">${S.privacy ? '⟨i:eye-off⟩' : '⟨i:eye⟩'}${sm ? '' : `<span>${S.privacy ? 'Ukryte' : 'Widoczne'}</span>`}</button>`;
/* spotkania Outlook/Teams są zarządzane w Outlooku: w TM nie można ich usuwać ani zmieniać terminu */
const isOutlook = m => m.source === 'outlook' || m.source === 'teams';
const canDelMtg = m => !isOutlook(m);
const outlookDayUrl = m => { const d = P(m.start.slice(0, 10)); return `https://outlook.office.com/calendar/view/day/${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`; };
const delMtgBtn = (m, cls = 'del') => canDelMtg(m) ? `<button class="${cls}" data-action="del-mtg" data-id="${m.id}" title="Usuń spotkanie">⟨i:trash⟩</button>` : `<button class="${cls} lock" data-action="ol-open" data-id="${m.id}" title="Spotkanie z Outlooka – usuwasz je w Outlooku (kliknij, aby otworzyć w aplikacji Outlook)">⟨i:lock⟩</button>`;
/* Outlook w aplikacji na komputerze (klasyczny Outlook przez serwer TM – start-TM.bat); bez serwera lub Outlooka: wersja w przeglądarce na żądanie */
/* linki Teams otwierane w aplikacji Microsoft Teams (protokół msteams:), nie w przeglądarce */
const teamsApp = u => u && /^https:\/\/teams\.(microsoft|live)\.com/i.test(u) ? u.replace(/^https:\/\/teams\.(microsoft|live)\.com/i, 'msteams:') : u;
const joinLink = (m, label, cls = 'btn sm') => { const u = teamsApp(m.link); return /^msteams:/.test(u) ? `<button type="button" class="${cls}" data-action="teams-open" data-url="${esc(u)}" title="Otwórz w aplikacji Microsoft Teams">${label}</button>` : `<a class="${cls}" href="${esc(u)}" target="_blank" rel="noopener">${label}</a>`; };
const openUrl = u => /^(msteams|mailto):/i.test(u) ? (location.href = u) : window.open(u, '_blank', 'noopener');
/* uruchomienie aplikacji Teams przez serwer TM (Windows otwiera Teams bez przeglądarki); bez serwera – protokół msteams: */
async function teamsLaunch(url, okMsg, extra = {}) {
  if (hasLocalSrv()) {
    try {
      const r = await fetch('/__outlook', {method: 'POST', headers: {'Content-Type': 'application/json', 'X-TM': '1'}, body: JSON.stringify({action: 'teams', url, ...extra})});
      if (r.ok) { const j = await r.json(); toast(j.via === 'outlook' ? 'Teams niedostępny – otworzono spotkanie w Outlook (dodaj „Spotkanie w aplikacji Teams”)' : okMsg || 'Otworzono Microsoft Teams'); return; }
      if (r.status === 404) toast('Uruchom ponownie start-TM.bat, aby otwierać Teams bez przeglądarki');
    } catch { }
  }
  openUrl(url); if (okMsg) toast(okMsg);
}
const hasLocalSrv = () => /^(localhost|127\.0\.0\.1)$/.test(location.hostname);
async function outlookApp(payload, webUrl, okMsg) {
  if (hasLocalSrv()) {
    try {
      const r = await fetch('/__outlook', {method: 'POST', headers: {'Content-Type': 'application/json', 'X-TM': '1'}, body: JSON.stringify(payload)});
      if (r.ok) { const j = await r.json(); if (okMsg) toast(j.via === 'ics' ? okMsg.replace('Outlook', 'kalendarz (plik .ics)') : j.via === 'mailto' ? okMsg.replace('Outlook', 'domyślny program pocztowy') : okMsg); return true; }
      const t = r.status === 404 ? 'Działa stara wersja serwera TM – uruchom ponownie start-TM.bat, aby otwierać aplikację Outlook.' : (await r.text()).slice(0, 160);
      if (webUrl) toastAction(t, 'Otwórz w przeglądarce', () => window.open(webUrl, '_blank', 'noopener'), 12000); else toast(t);
      return false;
    } catch { /* serwer niedostępny */ }
  }
  if (webUrl) toastAction('Aplikacja Outlook jest dostępna po uruchomieniu TM przez start-TM.bat.', 'Otwórz w przeglądarce', () => window.open(webUrl, '_blank', 'noopener'), 12000);
  return false;
}
const outlookOpenMtg = m => outlookApp({action: 'open', subject: m.title, start: m.start}, outlookDayUrl(m), 'Otworzono spotkanie w Outlook');
const srcTag = s => `<span class="tag src" style="color:${SRC[s]?.[1]}">${SRC[s]?.[0] || s}</span>`;

/* ---------- wykresy SVG ---------- */
const niceMax = v => { if (v <= 0) return 1; const p = 10 ** Math.floor(Math.log10(v)); return [1, 2, 2.5, 5, 10].find(m => m * p >= v) * p; };
function legend(items) { return `<div class="legend">${items.map(s => `<span><i style="background:${s.color}"></i>${esc(s.name)}</span>`).join('')}</div>`; }
function lineChart({labels, series, h = 230, max, suffix = ''}) {
  const W = 640, pl = 40, pr = 12, pt = 12, pb = 26, n = labels.length;
  const mx = max ?? niceMax(Math.max(1, ...series.flatMap(s => s.values.filter(v => v != null))));
  const x = i => pl + (n <= 1 ? 0 : i * (W - pl - pr) / (n - 1)), y = v => pt + (h - pt - pb) * (1 - v / mx);
  let s = `<svg viewBox="0 0 ${W} ${h}" width="100%" role="img">`;
  const nt = [4, 5, 3].find(n => Number.isInteger(mx / n)) || 4; for (let k = 0; k <= nt; k++) { const v = mx * k / nt; s += `<line x1="${pl}" x2="${W - pr}" y1="${y(v)}" y2="${y(v)}" stroke="var(--line)"/><text x="${pl - 6}" y="${y(v) + 4}" text-anchor="end">${r1(v)}${suffix}</text>`; }
  labels.forEach((l, i) => { if (n > 14 && i % 2) return; s += `<text x="${x(i)}" y="${h - 6}" text-anchor="middle">${esc(l)}</text>`; });
  series.forEach(se => {
    const pts = se.values.map((v, i) => v == null ? null : `${x(i)},${y(v)}`).filter(Boolean).join(' ');
    s += `<polyline points="${pts}" fill="none" stroke="${se.color}" stroke-width="2.5" ${se.dash ? 'stroke-dasharray="6 5"' : ''} stroke-linejoin="round" stroke-linecap="round"/>`;
    if (!se.dash) se.values.forEach((v, i) => { if (v != null) s += `<circle cx="${x(i)}" cy="${y(v)}" r="3.5" fill="${se.color}"><title>${esc(se.name)} · ${esc(labels[i])}: ${r1(v)}${suffix}</title></circle>`; });
  });
  if (series.todayIdx != null) s += `<line x1="${x(series.todayIdx)}" x2="${x(series.todayIdx)}" y1="${pt}" y2="${h - pb}" stroke="var(--bad)" stroke-dasharray="3 3"/>`;
  return s + '</svg>' + legend(series);
}
function barChart({labels, series, h = 220, suffix = ''}) {
  const W = 640, pl = 36, pr = 8, pt = 12, pb = 26, n = labels.length, k = series.length;
  const mx = niceMax(Math.max(1, ...series.flatMap(s => s.values)));
  const bw = (W - pl - pr) / n, y = v => pt + (h - pt - pb) * (1 - v / mx);
  let s = `<svg viewBox="0 0 ${W} ${h}" width="100%">`;
  const nt = [4, 5, 3, 2].find(n => Number.isInteger(mx / n)) || 4; for (let j = 0; j <= nt; j++) { const v = mx * j / nt; s += `<line x1="${pl}" x2="${W - pr}" y1="${y(v)}" y2="${y(v)}" stroke="var(--line)"/><text x="${pl - 6}" y="${y(v) + 4}" text-anchor="end">${r1(v)}</text>`; }
  labels.forEach((l, i) => {
    const gx = pl + i * bw, inner = bw * .7 / k;
    series.forEach((se, j) => { const v = se.values[i]; const bx = gx + bw * .15 + j * inner;
      s += `<rect x="${bx}" y="${y(v)}" width="${inner - 2}" height="${Math.max(0, y(0) - y(v))}" rx="3" fill="${se.color}"><title>${esc(se.name)} · ${esc(l)}: ${r1(v)}${suffix}</title></rect>`; });
    s += `<text x="${gx + bw / 2}" y="${h - 6}" text-anchor="middle">${esc(l)}</text>`;
  });
  return s + '</svg>' + legend(series);
}
function donut(parts, center) {
  const tot = parts.reduce((a, p) => a + p.value, 0) || 1, R = 70, C = 2 * Math.PI * R; let off = 0;
  let s = `<svg viewBox="0 0 200 200" width="180" height="180"><g transform="rotate(-90 100 100)">`;
  parts.forEach(p => { const L = C * p.value / tot; s += `<circle cx="100" cy="100" r="${R}" fill="none" stroke="${p.color}" stroke-width="26" stroke-dasharray="${L} ${C - L}" stroke-dashoffset="${-off}"><title>${esc(p.name)}: ${p.value}</title></circle>`; off += L; });
  return s + `</g><text x="100" y="96" text-anchor="middle" class="lbl-strong" style="font-size:24px">${center ?? tot}</text><text x="100" y="118" text-anchor="middle">zadań</text></svg>`;
}

/* ================= WIDOK: Pulpit ================= */
function meetingRow(m) {
  return `<div class="mtg" style="border-left:4px solid ${mtgColor(m)};padding-left:10px"><div class="tm">${hm(m.start)}–${hm(m.end)}</div><div style="flex:1;min-width:0"><div class="tt" data-action="edit-mtg" data-id="${m.id}" style="cursor:pointer">${ctxIcon(m.ctx)} ${srcBadge(m.source)} ${pv(esc(m.title))}</div>
    <div class="row" style="gap:6px;margin-top:4px">${srcTag(m.source)} ${ctxTag(m.ctx)} ${pv(projTag(m.projectId))} ${m.location ? `<span class="muted">⟨i:pin⟩ ${pv(esc(m.location))}</span>` : ''} ${m.attendees?.length ? `<span class="muted">⟨i:users⟩ ${m.attendees.length}</span>` : ''}</div></div>
    <div class="row">${m.link ? joinLink(m, 'Dołącz') : ''}<button class="btn sm" data-action="mtg-task" data-id="${m.id}" title="Utwórz zadanie follow-up">⟨i:plus⟩ To-do</button>${noteBadge('mtg', m.id)}<button class="btn sm ghost" data-action="edit-mtg" data-id="${m.id}" title="Edytuj">⟨i:pencil⟩</button>${delMtgBtn(m)}</div></div>`;
}
function capBar(label, c) {
  const u = Math.round(c.util * 100), sc = Math.max(c.cap, c.tasks + c.mtg, 1);
  const cls = u > 100 ? 'bad' : u > 85 ? 'warn' : 'ok';
  return `<div class="capbar"><div class="w">${label}</div><div class="b" title="Capacity ${r1(c.cap)}h, spotkania ${r1(c.mtg)}h, zadania ${r1(c.tasks)}h">
    <i class="m" style="width:${100 * c.mtg / sc}%"></i><i class="t" style="left:${100 * c.mtg / sc}%;width:${100 * c.tasks / sc}%"></i><span class="lim" style="left:${100 * c.cap / sc}%"></span></div>
    <div class="p ${cls}">${u}%</div></div>`;
}

/* ================= WIDOK: Zadania ================= */
function vTasks() {
  const F = S.tasks; let ts = tasksC();
  if (F.q) ts = ts.filter(t => norm(t.title + ' ' + (t.ic || '') + ' ' + pName(personById(t.assigneeId)) + ' ' + pName(personById(t.requesterId))).includes(norm(F.q)));
  if (F.goal) ts = ts.filter(t => (F.goal === '-' ? !t.goalId : t.goalId === F.goal));
  if (F.proj) ts = ts.filter(t => (F.proj === '-' ? !projOf(t) : projOf(t) === F.proj));
  const tt = taskTabs();
  const bar = `<div class="row" style="margin-bottom:12px">
    <input class="input" style="max-width:280px" placeholder="Szukaj: nazwa, IC case…" data-bind="tasks.q" value="${esc(F.q)}">
    <select class="input" style="max-width:240px" data-bind="tasks.proj"><option value="">Wszystkie projekty</option><option value="-" ${F.proj === '-' ? 'selected' : ''}>Bez projektu</option>
      ${projectsC().map(p => `<option value="${p.id}" ${F.proj === p.id ? 'selected' : ''}>${esc(p.title)}</option>`).join('')}</select>
    <select class="input" style="max-width:240px" data-bind="tasks.goal"><option value="">Wszystkie podprojekty</option><option value="-" ${F.goal === '-' ? 'selected' : ''}>Bez podprojektu</option>
      ${goalsC().map(g => `<option value="${g.id}" ${F.goal === g.id ? 'selected' : ''}>${esc(g.title)}</option>`).join('')}</select>
    ${F.mode === 'list' ? `<div class="seg">${[['open', 'Otwarte'], ['all', 'Wszystkie'], ['done', 'Zrobione']].map(([k, l]) => `<button data-action="set" data-k="tasks.status" data-v="${k}" class="${F.status === k ? 'on' : ''}">${l}</button>`).join('')}</div>` : ''}
    <div class="spacer"></div>
    <div class="seg"><button data-action="set" data-k="tasks.mode" data-v="list" class="${F.mode === 'list' ? 'on' : ''}">⟨i:list⟩ Lista</button><button data-action="set" data-k="tasks.mode" data-v="kanban" class="${F.mode === 'kanban' ? 'on' : ''}">⟨i:columns⟩ Kanban</button></div>
    ${eyeBtn()}<button class="btn" data-action="gantt-open" data-group="project" ${F.proj && F.proj !== '-' ? `data-only="${F.proj}"` : ''} title="Harmonogram zadań (Gantt)">⟨i:gantt⟩ Gantt</button><button class="btn primary" data-action="edit-task" ${F.proj && F.proj !== '-' ? `data-project="${F.proj}"` : ''}>⟨i:plus⟩ Nowe zadanie</button></div>`;
  const vis = visibleTasks(), visDone = ts.filter(t => t.status === 'done'), nSel = [...F.sel].filter(id => S.data.tasks.some(t => t.id === id)).length;
  const actions = `<div class="row bulkbar ${F.selMode ? 'on' : ''}" style="margin-bottom:12px">
    <button class="btn ${F.selMode ? 'primary' : ''}" data-action="sel-mode">⟨i:check-square⟩ ${F.selMode ? 'Zakończ zaznaczanie' : 'Zaznacz wiele'}</button>
    ${F.selMode ? `<button class="btn sm" data-action="sel-all">Zaznacz widoczne (${vis.length})</button><button class="btn sm" data-action="sel-none" ${nSel ? '' : 'disabled'}>Odznacz</button>
      <span class="muted"><b>${nSel}</b> zaznaczonych</span>
      <button class="btn sm" data-action="sel-done" ${nSel ? '' : 'disabled'}>⟨i:check⟩ Oznacz jako zrobione</button>
      <button class="btn sm danger" data-action="sel-delete" ${nSel ? '' : 'disabled'}>⟨i:trash⟩ Usuń zaznaczone (${nSel})</button>` : ''}
    <div class="spacer"></div>
    <button class="btn sm danger" data-action="del-done" ${visDone.length ? '' : 'disabled'} title="Usuwa zrobione zadania pasujące do bieżących filtrów i kontekstu">⟨i:trash⟩ Usuń zrobione (${visDone.length})</button></div>`;
  if (F.mode === 'kanban') {
    const cols = [['todo', 'Do zrobienia'], ['doing', 'W toku'], ['done', 'Zrobione']];
    return tt + bar + actions + `<div class="kanban">${cols.map(([k, l]) => { const items = ts.filter(t => t.status === k).sort(prioSort); return `
      <div class="kcol" data-status="${k}"><h4>${l}<span class="muted">${items.length}${F.selMode && items.length ? ` <button class="btn sm ghost" data-action="sel-group" data-ids="${items.map(t => t.id).join(',')}">zaznacz</button>` : ''}${k === 'done' && items.length ? ` <button class="btn sm ghost danger" data-action="del-done" title="Usuń wszystkie zrobione z tej kolumny">⟨i:trash⟩ Wyczyść</button>` : ''}</span></h4>
      ${items.map(t => `<div class="kcard ${t.ctx} ${F.selMode && F.sel.has(t.id) ? 'selected' : ''}" draggable="true" data-id="${t.id}" data-action="edit-task"><div class="t" style="display:flex;gap:6px">${F.selMode ? `<input type="checkbox" class="selchk" data-action="sel-task" data-id="${t.id}" ${F.sel.has(t.id) ? 'checked' : ''}>` : ''}<span style="flex:1">${S.privacy ? pv(esc(t.title)) : esc(t.title)}</span><button class="del" data-action="del-task-row" data-id="${t.id}" title="Usuń zadanie">⟨i:trash⟩</button></div>
        <div class="row" style="gap:4px">${(S.privacy ? pv : (x => x))(`${ctxTag(t.ctx)}${projTag(projOf(t))}<span class="tag p${t.prio}">P${t.prio}</span>`)}${personTags(t)}${t.ic ? icTag(t) : ''}${t.est ? `<span class="tag src">${t.est}h</span>` : ''}<span class="due ${isLate(t) ? 'late' : ''}" style="margin-left:auto;font-size:12px">${t.due ? relDay(t.due) : ''}</span>${noteBadge('task', t.id)}</div></div>`).join('')}</div>`; }).join('')}</div>`;
  }
  const td = TODAY(), tm = iso(add(today(), 1)), we = iso(add(sow(today()), 6));
  const open = ts.filter(t => t.status !== 'done').sort(prioSort), done = ts.filter(t => t.status === 'done').sort((a, b) => (b.doneAt || '').localeCompare(a.doneAt || ''));
  const groups = [];
  if (F.status !== 'done') groups.push(
    ['⟨i:alert⟩ Zaległe', open.filter(t => t.due && t.due < td)], ['Dziś', open.filter(t => t.due === td)], ['Jutro', open.filter(t => t.due === tm)],
    ['W tym tygodniu', open.filter(t => t.due > tm && t.due <= we)], ['Później', open.filter(t => t.due > we)], ['Bez terminu', open.filter(t => !t.due)]);
  if (F.status !== 'open') groups.push(['⟨i:check⟩ Zrobione', done]);
  const body = groups.filter(g => g[1].length).map(([h, l]) => `<div class="group-h row" style="gap:8px">${h} · ${l.length}
      ${F.selMode ? `<button class="btn sm ghost" data-action="sel-group" data-ids="${l.map(t => t.id).join(',')}">zaznacz grupę</button>` : ''}
      ${h.startsWith('⟨i:check⟩') ? `<button class="btn sm ghost danger" data-action="del-done">⟨i:trash⟩ usuń wszystkie zrobione</button>` : ''}</div>${l.map(t => taskRow(t, {sel: true})).join('')}`).join('');
  return tt + bar + actions + `<div class="card">${body || '<div class="empty">Brak zadań</div>'}</div>`;
}
afterRender.tasks = () => {
  $$('.kcard').forEach(c => c.addEventListener('dragstart', e => e.dataTransfer.setData('text', c.dataset.id)));
  $$('.kcol').forEach(col => {
    col.addEventListener('dragover', e => { e.preventDefault(); col.classList.add('over'); });
    col.addEventListener('dragleave', () => col.classList.remove('over'));
    col.addEventListener('drop', e => { e.preventDefault(); const t = S.data.tasks.find(x => x.id === e.dataTransfer.getData('text')); if (t) setStatus(t, col.dataset.status); });
  });
};
function setStatus(t, st) { t.status = st; t.doneAt = st === 'done' ? TODAY() : null; save(); render(); }

/* ================= WIDOK: Podprojekty ================= */
function goalCard(g) {
  const p = goalProgress(g), e = goalExpected(g), ts = S.data.tasks.filter(t => t.goalId === g.id), open = ts.filter(t => t.status !== 'done');
  const st = goalStatus(g, p, e); if (st[0] === 'osiągnięty') st[0] = '⟨i:check⟩ osiągnięty';
  return `<div class="card"><div class="goal">${ring(p, ctxColor(g.ctx))}<div style="flex:1;min-width:0">
    <div class="gt" data-action="edit-goal" data-id="${g.id}" style="cursor:pointer">${esc(g.title)}</div>
    <div class="row" style="gap:6px;margin-top:4px">${ctxTag(g.ctx)}${g.projectId ? `<span data-action="open-project" data-id="${g.projectId}" style="cursor:pointer">${projTag(g.projectId)}</span>` : ''}<span class="muted">${fmt(g.start)} → ${fmt(g.due)}</span><span class="p ${st[1]}" style="font-weight:700;font-size:12px">● ${st[0]}</span></div>
    <div class="muted" style="font-size:12px;margin-top:4px">${esc(g.desc || '')}</div></div></div>
    <div class="muted" style="font-size:12px;margin:10px 0 2px">Realizacja ${p}% · plan wg czasu ${e}% · zadania ${ts.length - open.length}/${ts.length}</div>
    ${open.sort(prioSort).slice(0, 4).map(t => taskRow(t, {noGoal: true, noProj: true})).join('')}
    ${ts.length - open.length ? `<details class="done-box"><summary class="muted" style="font-size:12px;cursor:pointer;margin-top:6px">⟨i:check⟩ Zrobione (${ts.length - open.length}) – pokaż / usuń</summary>${ts.filter(t => t.status === 'done').map(t => taskRow(t, {noGoal: true, noProj: true})).join('')}
      <button class="btn sm ghost danger" data-action="del-done-ids" data-ids="${ts.filter(t => t.status === 'done').map(t => t.id).join(',')}">⟨i:trash⟩ usuń zrobione z tego podprojektu</button></details>` : ''}
    <div class="row" style="margin-top:8px"><button class="btn sm" data-action="edit-task" data-goal="${g.id}">⟨i:plus⟩ Zadanie do podprojektu</button></div></div>`;
}
function vGoals() {
  const F = S.goals; let gs = goalsC().sort(goalSort);
  if (F.proj) gs = gs.filter(g => (F.proj === '-' ? !g.projectId : g.projectId === F.proj));
  const bar = `<div class="row" style="margin-bottom:12px">
    <select class="input" style="max-width:260px" data-bind="goals.proj"><option value="">Wszystkie projekty</option><option value="-" ${F.proj === '-' ? 'selected' : ''}>Bez projektu</option>
      ${projectsC().map(p => `<option value="${p.id}" ${F.proj === p.id ? 'selected' : ''}>${esc(p.title)}</option>`).join('')}</select>
    <div class="seg"><button data-action="set" data-k="goals.group" data-v="1" class="${F.group ? 'on' : ''}">⟨i:folder⟩ Wg projektów</button><button data-action="set" data-k="goals.group" data-v="0" class="${!F.group ? 'on' : ''}">⟨i:columns⟩ Siatka</button></div>
    <div class="spacer"></div><button class="btn" data-action="gantt-open" data-group="goal" ${F.proj && F.proj !== '-' ? `data-only="${F.proj}"` : ''} title="Harmonogram podprojektów (Gantt)">⟨i:gantt⟩ Gantt</button><button class="btn primary" data-action="edit-goal" ${F.proj && F.proj !== '-' ? `data-project="${F.proj}"` : ''}>⟨i:plus⟩ Nowy podprojekt</button></div>`;
  if (!gs.length) return bar + '<div class="card empty">Brak podprojektów w tym kontekście</div>';
  if (!F.group) return bar + `<div class="grid g2">${gs.map(goalCard).join('')}</div>`;
  const groups = [...projectsC().map(p => [p, gs.filter(g => g.projectId === p.id)]), [null, gs.filter(g => !projById(g.projectId))]].filter(x => x[1].length);
  return bar + groups.map(([p, l]) => `<div class="group-h" style="margin-top:18px;display:flex;align-items:center;gap:8px">${p ? `<span class="dot ${p.ctx}"></span><span data-action="open-project" data-id="${p.id}" style="cursor:pointer">⟨i:folder⟩ ${esc(p.title)}</span> · ${projectProgress(p)}%` : 'Podprojekty bez projektu'} · ${l.length}</div>
    <div class="grid g2">${l.map(goalCard).join('')}</div>`).join('');
}

/* ================= WIDOK: Projekty ================= */
/* ================= BUDŻETY: PRIV – złotówki (podprojekt: budżet; zadania: koszt szacowany / rzeczywisty), JOB – godziny (budżet h vs realizacja z ewidencji) ================= */
const zl = v => Math.round(+v || 0).toLocaleString('pl-PL') + ' zł';
const hh = v => (Math.round((+v || 0) * 10) / 10).toLocaleString('pl-PL') + ' h';
const sumBy = (l, f) => l.reduce((a, x) => a + (+f(x) || 0), 0);
function goalMoney(g) { const ts = S.data.tasks.filter(t => t.goalId === g.id); return {budget: +g.budget || 0, est: sumBy(ts, t => t.costEst), act: sumBy(ts, t => t.costAct)}; }
function projMoney(p) { const ts = projTasks(p); return {budget: sumBy(projGoals(p), g => g.budget), est: sumBy(ts, t => t.costEst), act: sumBy(ts, t => t.costAct)}; }
function goalHours(g) { const ts = S.data.tasks.filter(t => t.goalId === g.id); return {budget: +g.budgetH || 0, est: sumBy(ts, t => t.est), act: sumBy(ts, t => taskSpent(t.id)) / 60}; }
function projHours(p) {
  const ts = projTasks(p), ids = new Set(ts.map(t => t.id));
  const act = entriesWithLive().filter(e => ids.has(e.taskId) || (!e.taskId && e.projectId === p.id)).reduce((a, e) => a + entMin(e), 0) / 60;
  return {budget: +p.budgetH || sumBy(projGoals(p), g => g.budgetH), est: sumBy(ts, t => t.est), act};
}
/* poziom: none (brak budżetu) / ok / warn (≥90% albo szacunek ponad budżet) / bad (>100%) */
function budgetLevel(b) { if (!b.budget) return 'none'; const r = b.act / b.budget; return r > 1 ? 'bad' : r >= .9 || b.est > b.budget ? 'warn' : 'ok'; }
function budgetCell(b, money, compact) {
  const f = money ? zl : hh, lv = budgetLevel(b);
  if (!b.budget) return b.act || b.est ? `<span class="muted" style="font-size:12px">${money ? 'koszt' : 'real.'} ${f(b.act)}${b.est ? ` · ${money ? 'szac.' : 'plan'} ${f(b.est)}` : ''}<br>bez budżetu</span>` : '<span class="muted">—</span>';
  const pa = Math.min(100, Math.round(100 * b.act / b.budget)), pe = Math.min(100, Math.round(100 * b.est / b.budget));
  return `<div class="bcell ${lv}" title="Budżet ${f(b.budget)} · ${money ? 'szacunek' : 'plan (estymacje)'} ${f(b.est)} · ${money ? 'koszt rzeczywisty' : 'realizacja'} ${f(b.act)}">
    <div class="bar" style="position:relative;height:7px"><i style="width:${pa}%"></i>${b.est ? `<b style="left:${pe}%"></b>` : ''}</div>
    <div class="bl"><b>${f(b.act)}</b> / ${f(b.budget)}${compact ? '' : ` <span class="muted">· ${money ? 'szac.' : 'plan'} ${f(b.est)}</span>`}${lv === 'bad' ? ' <span class="bbad">przekroczony</span>' : lv === 'warn' && b.est > b.budget && b.act <= b.budget * .9 ? ' <span class="bwarn">szacunek ponad budżet</span>' : ''}</div></div>`;
}
function budgetCard(p) {
  const money = p.ctx === 'priv', b = money ? projMoney(p) : projHours(p), f = money ? zl : hh, lv = budgetLevel(b);
  const left = b.budget - b.act;
  return `<div class="card budget-card ${lv}" style="margin-bottom:16px"><div class="row" style="gap:10px;align-items:center"><h3 style="margin:0">⟨i:wallet⟩ ${money ? 'Budżet i koszty' : 'Budżet godzin'}</h3>
    <span class="muted" style="font-size:12px">${money ? 'budżet = suma budżetów podprojektów · szacunek i koszt = suma z zadań' : `budżet ${p.budgetH ? 'projektu' : '= suma budżetów podprojektów'} · plan = estymacje zadań · realizacja = ewidencja czasu`}</span>
    <div class="spacer"></div><button class="btn sm ghost" data-action="edit-project" data-id="${p.id}">⟨i:pencil⟩ ${money ? '' : 'Budżet h'}</button></div>
    <div class="grid g4 kpis" style="margin-top:10px">
      <div class="kpi"><div class="v">${b.budget ? f(b.budget) : '—'}</div><div class="l">Budżet</div></div>
      <div class="kpi"><div class="v">${f(b.est)}</div><div class="l">${money ? 'Szacunek (zadania)' : 'Plan (estymacje)'}</div></div>
      <div class="kpi"><div class="v ${lv === 'bad' ? 'bbad' : ''}">${f(b.act)}</div><div class="l">${money ? 'Koszt rzeczywisty' : 'Realizacja'}</div></div>
      <div class="kpi"><div class="v ${b.budget && left < 0 ? 'bbad' : ''}">${b.budget ? f(left) : '—'}</div><div class="l">${b.budget && left < 0 ? 'Przekroczenie' : 'Pozostało'}</div></div></div>
    ${(() => { const over = projGoals(p).filter(g => budgetLevel(money ? goalMoney(g) : goalHours(g)) === 'bad'); return over.length ? `<p class="bbad" style="margin:8px 0 0;font-size:13px">⟨i:alert⟩ Przekroczony budżet podprojektu: ${over.map(g => esc(g.title)).join(', ')}</p>` : ''; })()}
    ${b.budget ? budgetCell(b, money) : `<p class="hint" style="margin:8px 0 0">${money ? 'Ustaw budżet w podprojektach projektu (edycja podprojektu → Budżet zł).' : 'Ustaw budżet godzin w edycji projektu albo w podprojektach.'}</p>`}</div>`;
}
/* pola formularza widoczne tylko dla PRIV / JOB (data-only="priv|job") */
function ctxFields(m) {
  const f = $('form', m); if (!f || !f.ctx) return;
  const sync = () => $$('[data-only]', m).forEach(el => { el.hidden = el.dataset.only !== f.ctx.value; });
  f.addEventListener('change', () => setTimeout(sync)); sync();
}
function projStats(p) {
  const ts = projTasks(p), open = ts.filter(t => t.status !== 'done'), late = open.filter(isLate);
  const rest = open.reduce((a, t) => a + (t.est || 0), 0), pr = projectProgress(p), ex = p.start && p.due ? (() => { const len = diffD(p.start, p.due); return len > 0 ? Math.round(100 * clamp(diffD(p.start, TODAY()) / len, 0, 1)) : 100; })() : 0;
  return {ts, open, late, rest, pr, ex, gs: projGoals(p)};
}
function projCard(p, sel) {
  const x = projStats(p), P_ = S.privacy ? pv : y => y, nx = S.data.tasks.filter(t => projOf(t) === p.id && t.status !== 'done').sort((a, b) => a.prio - b.prio || prioSort(a, b))[0];
  return `<div class="ncard ${sel === p.id ? 'sel' : ''}" data-action="${sel !== undefined ? 'home-proj' : 'open-project'}" data-id="${p.id}" title="${sel !== undefined ? 'Pokaż zadania projektu poniżej' : 'Otwórz projekt'}"><div class="nc-thumb">${art(p.id, projColor(p) || ctxColor(p.ctx), '', 'wide')}
      <span class="nc-badge ${p.ctx}">${p.ctx === 'priv' ? 'PRIV' : 'JOB'}</span>${p.status === 'planned' ? '<span class="nc-flag">Planowany</span>' : x.late.length ? `<span class="nc-flag bad">${x.late.length} po terminie</span>` : ''}
      <div class="nc-title">${P_(esc(p.title))}</div><span class="nc-prog"><i style="width:${x.pr}%"></i></span></div>
    <div class="nc-info"><div class="nc-meta"><b class="ok">${x.pr}%</b><span>${x.ts.length - x.open.length}/${x.ts.length} zadań</span>${p.due ? `<span>do ${fmt(p.due)}</span>` : ''}</div>
      <div class="nc-sub">${P_(esc(p.client || (p.ctx === 'priv' ? 'Projekt prywatny' : 'Projekt firmowy')))}</div>
      <div class="nc-acts"><button class="rbtn light sm" data-action="proj-play" data-id="${p.id}" title="${nx ? 'Realizuj: ' + esc(nx.title) : 'Start realizacji w projekcie'}">${PLAY_F} Realizuj</button>${sel !== undefined ? `<button class="rbtn sm ghost" data-action="open-project" data-id="${p.id}" title="Otwórz projekt">⟨i:folder⟩</button>` : ''}<span class="nc-next">${nx ? 'Dalej: ' + P_(esc(nx.title)) : 'Brak otwartych zadań'}</span></div></div>${sel === p.id ? '<span class="nc-sel">Zadania poniżej ⟨i:chev-down⟩</span>' : ''}</div>`;
}
function projTabs() {
  return ''; return `<div class="seg caltabs" style="margin-bottom:12px"><button data-action="nav" data-view="projects" class="${S.view === 'projects' ? 'on' : ''}">⟨i:list⟩ Lista projektów</button><button data-action="nav" data-view="project" class="${S.view === 'project' ? 'on' : ''}">⟨i:folder⟩ Projekt${projById(S.proj.open) ? ': ' + esc(projById(S.proj.open).title) : ''}</button></div>`;
}
const projMeetings = p => S.data.meetings.filter(m => m.projectId === p.id);
const progCell = (pr, ex, color) => { const d = pr - ex; return `<div class="pcell"><div class="bar" style="position:relative;height:7px"><i style="width:${pr}%;background:${color}"></i><b style="position:absolute;left:${ex}%;top:-3px;bottom:-3px;width:2px;background:var(--text)"></b></div><span><b>${pr}%</b> <span class="p ${d >= 0 ? 'ok' : d > -10 ? 'warn' : 'bad'}">${d >= 0 ? '+' : ''}${d} pp</span></span></div>`; };
const sortTh = (k, label, cls = '') => `<th class="sortable ${cls}" data-action="psort" data-k="${k}">${label}${S.proj.sort === k ? (S.proj.dir > 0 ? ' ⟨i:chev-up⟩' : ' ⟨i:chev-down⟩') : ''}</th>`;

/* ================= WIDOK: Projekty › Lista projektów (tabela) ================= */
function vProjects() {
  const P0 = S.proj, q = norm(P0.q);
  let ps = projectsC();
  if (P0.status === 'active') ps = ps.filter(p => p.status !== 'done');
  if (P0.status === 'done') ps = ps.filter(p => p.status === 'done');
  if (q) ps = ps.filter(p => norm([p.title, p.client, p.desc].join(' ')).includes(q));
  const rows = ps.map(p => { const x = projStats(p), ms = projMeetings(p), up = ms.filter(m => m.start.slice(0, 10) >= TODAY()); return {p, x, up, ms}; });
  const order = {active: 0, planned: 1, onhold: 2, done: 3};
  const key = {status: r => (order[r.p.status] ?? 9) * 1e4 + (r.p.due ? diffD(TODAY(), r.p.due) : 9999), title: r => norm(r.p.title), ctx: r => r.p.ctx, due: r => r.p.due || '9999', progress: r => r.x.pr, delta: r => r.x.pr - r.x.ex,
    goals: r => r.x.gs.length, tasks: r => r.x.ts.length, late: r => r.x.late.length, mtgs: r => r.up.length, rest: r => r.x.rest}[P0.sort] || (r => 0);
  rows.sort((a, b) => { const A = key(a), B = key(b); return (A > B ? 1 : A < B ? -1 : 0) * P0.dir; });
  const tot = rows.reduce((a, r) => ({g: a.g + r.x.gs.length, t: a.t + r.x.ts.length, d: a.d + r.x.ts.length - r.x.open.length, l: a.l + r.x.late.length, m: a.m + r.up.length, h: a.h + r.x.rest}), {g: 0, t: 0, d: 0, l: 0, m: 0, h: 0});
  return projTabs() + `<div class="row" style="margin-bottom:12px">
    <input class="input" style="max-width:300px" placeholder="Szukaj projektu, klienta…" data-bind="proj.q" value="${esc(P0.q)}">
    <div class="seg">${[['active', 'W realizacji'], ['all', 'Wszystkie'], ['done', 'Zakończone']].map(([k, l]) => `<button data-action="set" data-k="proj.status" data-v="${k}" class="${P0.status === k ? 'on' : ''}">${l}</button>`).join('')}</div>
    <div class="spacer"></div><button class="btn" data-action="ptpl-use" title="Nowy projekt z gotowym zestawem podprojektów i zadań">⟨i:template⟩ Z szablonu</button><button class="btn primary" data-action="edit-project">⟨i:plus⟩ Nowy projekt</button></div>
  <div class="card" style="padding:0;overflow:auto"><table class="tbl ptbl"><thead><tr><th style="width:8px"></th>
    ${sortTh('title', 'Projekt')}${sortTh('ctx', 'Kontekst')}${sortTh('status', 'Status')}${sortTh('due', 'Termin')}${sortTh('progress', 'Realizacja vs plan')}
    ${sortTh('goals', 'Podprojekty', 'num')}${sortTh('tasks', 'Zadania', 'num')}${sortTh('late', 'Po terminie', 'num')}${sortTh('mtgs', 'Spotkania', 'num')}${sortTh('rest', 'Pozostało', 'num')}<th>Budżet</th><th></th></tr></thead><tbody>
    ${rows.map(({p, x, up}) => { const [sl, sc] = PSTAT[p.status] || PSTAT.active; return `<tr class="prow" data-action="open-project" data-id="${p.id}" title="Otwórz projekt">
      <td><span class="mcol" style="background:${projColor(p)}"></span></td>
      <td class="pt"><b>${esc(p.title)}</b>${p.client ? `<div class="muted" style="font-size:11px">${esc(p.client)}</div>` : ''}</td>
      <td>${ctxIcon(p.ctx)}${p.ctx === 'job' ? 'JOB' : 'PRIV'}</td>
      <td><span class="p ${sc}" style="font-weight:600">● ${sl}</span></td>
      <td class="nowrap">${fmt(p.due)}<div class="muted" style="font-size:11px">od ${fmt(p.start)}</div></td>
      <td style="min-width:170px">${progCell(x.pr, x.ex, projColor(p))}</td>
      <td class="num">${x.gs.length}</td><td class="num">${x.ts.length - x.open.length}/${x.ts.length}</td>
      <td class="num" style="color:${x.late.length ? 'var(--bad)' : 'inherit'};font-weight:${x.late.length ? 700 : 400}">${x.late.length}</td>
      <td class="num">${up.length}</td><td class="num">${r1(x.rest)}h</td><td style="min-width:150px">${budgetCell(p.ctx === 'priv' ? projMoney(p) : projHours(p), p.ctx === 'priv', true)}</td>
      <td class="nowrap acts"><button class="btn sm ghost" data-action="edit-project" data-id="${p.id}" title="Edytuj projekt">⟨i:pencil⟩</button><button class="btn sm ghost" data-action="open-project" data-id="${p.id}" title="Otwórz">⟨i:chev-right⟩</button></td></tr>`; }).join('')
      || `<tr><td colspan="12" class="empty">Brak projektów dla wybranych filtrów. Utwórz projekt przyciskiem „Nowy projekt”.</td></tr>`}</tbody>
    ${rows.length ? `<tfoot><tr><td></td><td><b>${rows.length} projektów</b></td><td colspan="4"></td><td class="num">${tot.g}</td><td class="num">${tot.d}/${tot.t}</td><td class="num">${tot.l}</td><td class="num">${tot.m}</td><td class="num">${r1(tot.h)}h</td><td></td></tr></tfoot>` : ''}</table></div>
  <div class="legend"><span>Kliknij wiersz, aby otworzyć ekran projektu</span><span>· Spotkania = nadchodzące</span><span>· pionowa kreska na pasku = plan wg upływu czasu</span></div>`;
}

/* ================= WIDOK: Projekty › Projekt ================= */
function vProject() {
  const P0 = S.proj, p = projById(P0.open);
  const picker = `<div class="row picker" style="margin-bottom:14px"><span class="muted" style="font-weight:600">Projekt</span>
    <select class="input" style="max-width:360px" data-bind="proj.open"><option value="">— wybierz projekt —</option>
      ${['job', 'priv'].map(c => { const l = (S.data.projects || []).filter(x => x.ctx === c && (S.ctx === 'all' || S.ctx === c || x.id === P0.open)); return l.length ? `<optgroup label="${c === 'job' ? 'JOB – firmowe' : 'PRIV – prywatne'}">${l.map(x => `<option value="${x.id}" ${x.id === P0.open ? 'selected' : ''}>${esc(x.title)}${x.status === 'done' ? ' (zakończony)' : ''}</option>`).join('')}</optgroup>` : ''; }).join('')}</select>
    <button class="btn" data-action="edit-project">⟨i:plus⟩ Nowy projekt</button><button class="btn" data-action="ptpl-use">⟨i:template⟩ Z szablonu</button>
    ${p ? `<div class="spacer"></div><button class="btn sm" data-action="proj-to-tpl" data-id="${p.id}" title="Zapisz podprojekty i zadania tego projektu jako szablon">⟨i:template⟩ Zapisz jako szablon</button><button class="btn sm" data-action="proj-gantt" data-id="${p.id}">⟨i:gantt⟩ Gantt</button><button class="btn sm" data-action="edit-project" data-id="${p.id}">⟨i:pencil⟩ Edytuj</button>` : ''}</div>`;
  if (!p) {
    const q = norm(P0.pq), list = projectsC().filter(x => !q || norm([x.title, x.client].join(' ')).includes(q));
    return projTabs() + picker + `<div class="card"><h3>⟨i:search⟩ Znajdź projekt</h3>
      <input class="input" style="max-width:420px;margin-bottom:10px" placeholder="Nazwa projektu lub klient…" data-bind="proj.pq" value="${esc(P0.pq)}">
      ${list.map(x => { const st = projStats(x); return `<div class="pick" data-action="open-project" data-id="${x.id}"><span class="mcol" style="background:${projColor(x)};height:22px"></span><b>${esc(x.title)}</b><span class="muted">${x.client ? esc(x.client) + ' · ' : ''}${ctxIcon(x.ctx)}${x.ctx.toUpperCase()} · ${PSTAT[x.status]?.[0]}</span><span class="spacer"></span><span class="muted">${st.pr}% · termin ${fmt(x.due)}</span>⟨i:chev-right⟩</div>`; }).join('') || '<div class="empty">Nie znaleziono projektu</div>'}
      <div class="row" style="margin-top:12px"><button class="btn primary" data-action="edit-project">⟨i:plus⟩ Utwórz nowy projekt</button><button class="btn" data-action="ptpl-use">⟨i:template⟩ Nowy projekt z szablonu</button></div></div>`;
  }
  const x = projStats(p), [sl, sc] = PSTAT[p.status] || PSTAT.active, col = projColor(p), ms = projMeetings(p);
  const outCtx = S.ctx !== 'all' && S.ctx !== p.ctx;
  // podprojekty
  x.gs = [...x.gs].sort(goalSort);
  const goalRows = x.gs.map(g => { const gt = S.data.tasks.filter(t => t.goalId === g.id), go = gt.filter(t => t.status !== 'done'), gp = goalProgress(g), ge = goalExpected(g);
    const st = goalStatus(g, gp, ge);
    return `<tr><td class="pt" data-action="edit-goal" data-id="${g.id}"><b>${esc(g.title)}</b>${g.desc ? `<div class="muted" style="font-size:11px">${esc(g.desc)}</div>` : ''}</td>
      <td class="nowrap">${fmt(g.due)}</td><td class="nowrap"><span class="p ${st[1]}" style="font-weight:600">● ${st[0]}</span></td><td style="min-width:170px">${progCell(gp, ge, col)}</td><td style="min-width:190px">${budgetCell(p.ctx === 'priv' ? goalMoney(g) : goalHours(g), p.ctx === 'priv', true)}</td>
      <td class="num">${gt.length - go.length}/${gt.length}</td><td class="num" style="color:${go.filter(isLate).length ? 'var(--bad)' : 'inherit'}">${go.filter(isLate).length}</td><td class="num">${r1(go.reduce((a, t) => a + (t.est || 0), 0))}h</td>
      <td class="nowrap acts"><button class="btn sm ghost" data-action="edit-task" data-goal="${g.id}" title="Dodaj zadanie do podprojektu">⟨i:plus⟩</button><button class="btn sm ghost" data-action="edit-goal" data-id="${g.id}" title="Edytuj podprojekt">⟨i:pencil⟩</button></td></tr>`; }).join('');
  // zadania wg podprojektów
  const tf = t => P0.tstat === 'all' || (P0.tstat === 'open' ? t.status !== 'done' : t.status === 'done');
  const grp = [...x.gs.map(g => [g, x.ts.filter(t => t.goalId === g.id)]), [null, x.ts.filter(t => !t.goalId || !goalById(t.goalId) || goalById(t.goalId).projectId !== p.id)]];
  const taskBlocks = grp.map(([g, l]) => { const v = l.filter(tf).sort(P0.tstat === 'done' ? (a, b) => (b.doneAt || '').localeCompare(a.doneAt || '') : prioSort); if (!v.length && !g) return '';
    return `<div class="group-h row" style="gap:8px;margin-top:14px">${g ? `⟨i:layers⟩ ${esc(g.title)} <span class="muted">· ${goalProgress(g)}% · ${l.filter(t => t.status === 'done').length}/${l.length}</span>` : '⟨i:folder⟩ Zadania projektu bez podprojektu'}
      <button class="btn sm ghost" data-action="edit-task" ${g ? `data-goal="${g.id}"` : `data-project="${p.id}"`}>⟨i:plus⟩ zadanie</button></div>
      ${v.map(t => taskRow(t, {noProj: true, noGoal: true})).join('') || '<div class="muted" style="padding:6px 8px;font-size:12px">Brak zadań w tym filtrze</div>'}`; }).join('');
  const doneIds = x.ts.filter(t => t.status === 'done').map(t => t.id);
  // spotkania
  const mf = m => P0.mper === 'all' || (P0.mper === 'next' ? m.end.slice(0, 10) >= TODAY() : m.end.slice(0, 10) < TODAY());
  const mv = ms.filter(mf).sort((a, b) => P0.mper === 'past' ? b.start.localeCompare(a.start) : a.start.localeCompare(b.start));
  const mHours = r1(mv.reduce((a, m) => a + (toMin(m.end) - toMin(m.start)) / 60, 0));
  const mRows = mv.map(m => `<tr class="mrow"><td class="nowrap"><b>${fmt(m.start.slice(0, 10))}</b> <span class="muted">${DOW[(P(m.start.slice(0, 10)).getDay() + 6) % 7]}</span></td><td class="nowrap">${hm(m.start)}–${hm(m.end)}</td>
    <td class="mt" data-action="edit-mtg" data-id="${m.id}">${ctxIcon(m.ctx)}${pv(esc(m.title))}${m.location ? `<div class="muted" style="font-size:11px">⟨i:pin⟩ ${pv(esc(m.location))}</div>` : ''}</td>
    <td class="nowrap">${srcBadge(m.source)} ${SRC[m.source]?.[0]}</td><td class="num">${(m.attendees || []).length || '—'}</td>
    <td class="nowrap acts">${cmtBadge('mtg', m.id)}${noteBadge('mtg', m.id)}<button class="btn sm ghost" data-action="show-in-cal" data-id="${m.id}" title="Pokaż w kalendarzu">⟨i:calendar⟩</button><button class="btn sm ghost" data-action="edit-mtg" data-id="${m.id}" title="Edytuj">⟨i:pencil⟩</button>${delMtgBtn(m, 'del btn sm ghost')}</td></tr>${S.cx?.has(cmtKey('mtg', m.id)) ? `<tr class="subrow"><td colspan="6">${cmtBlock('mtg', m.id)}</td></tr>` : ''}`).join('');
  return projTabs() + picker + (outCtx ? `<div class="hint lockinfo" style="margin-bottom:12px">Ten projekt należy do kontekstu <b>${p.ctx.toUpperCase()}</b>, a przełącznik jest ustawiony na <b>${S.ctx.toUpperCase()}</b>. Projekt wyświetlam w całości.</div>` : '') + `
   <div class="card phead" style="margin-bottom:16px;border-top:4px solid ${col}"><div class="goal">${ring(x.pr, col, 72)}<div style="flex:1;min-width:0">
     <div class="gt" style="font-size:19px">${esc(p.title)}</div>
     <div class="row" style="gap:8px;margin-top:6px">${ctxTag(p.ctx)}<span class="p ${sc}" style="font-weight:700;font-size:12px">● ${sl}</span><span class="muted">${fmt(p.start)} → ${fmt(p.due)}</span>${p.client ? `<span class="muted">· ${esc(p.client)}</span>` : ''}</div>
     <div class="muted" style="margin-top:6px">${esc(p.desc || '')}</div></div></div>
     <div class="bar" style="position:relative;margin:14px 0 4px;height:10px"><i style="width:${x.pr}%;background:${col}"></i><b title="Plan wg czasu" style="position:absolute;left:${x.ex}%;top:-3px;bottom:-3px;width:3px;background:var(--text)"></b></div>
     <div class="muted" style="font-size:12px">Realizacja ${x.pr}% · plan wg czasu ${x.ex}% · pozostała praca ${r1(x.rest)}h</div></div>
   <div class="grid g4 kpis" style="margin-bottom:16px">
     <div class="card kpi"><div class="v">${x.gs.length}</div><div class="l">Podprojekty</div></div>
     <div class="card kpi"><div class="v">${x.ts.length - x.open.length}/${x.ts.length}</div><div class="l">Zadania zrobione</div></div>
     <div class="card kpi"><div class="v" style="color:${x.late.length ? 'var(--bad)' : 'inherit'}">${x.late.length}</div><div class="l">Zadania po terminie</div></div>
     <div class="card kpi"><div class="v">${ms.filter(m => m.start.slice(0, 10) >= TODAY()).length}</div><div class="l">Nadchodzące spotkania</div></div></div>
   ${budgetCard(p)}
   <div class="psec-nav row"><a href="#" data-action="psec" data-t="psGoals">Podprojekty</a><a href="#" data-action="psec" data-t="psTasks">Zadania</a><a href="#" data-action="psec" data-t="psMtgs">Spotkania</a><a href="#" data-action="psec" data-t="psTrips">Delegacje (${tripsAll().filter(t => t.projectId === p.id).length})</a><a href="#" data-action="psec" data-t="psNotes">Notatki (${notesAll().filter(n => noteProj(n) === p.id).length})</a></div>
   <div class="card" id="psGoals" style="margin-bottom:16px;padding:0"><div class="row card-h"><h3 style="margin:0">⟨i:layers⟩ Zestawienie podprojektów projektu</h3><div class="spacer"></div><button class="btn sm" data-action="edit-goal" data-project="${p.id}">⟨i:plus⟩ Podprojekt</button></div>
     <table class="tbl ptbl"><thead><tr><th>Podprojekt</th><th>Termin</th><th>Status</th><th>Realizacja vs plan</th><th>${p.ctx === 'priv' ? 'Budżet (koszt / budżet)' : 'Budżet godzin'}</th><th class="num">Zadania</th><th class="num">Po terminie</th><th class="num">Pozostało</th><th></th></tr></thead>
     <tbody>${goalRows || `<tr><td colspan="9" class="empty">Brak podprojektów – dodaj podprojekt do projektu</td></tr>`}</tbody></table></div>
   <div class="card" id="psTasks" style="margin-bottom:16px"><div class="row" style="margin-bottom:4px"><h3 style="margin:0">⟨i:check-square⟩ Zadania wg podprojektów</h3><div class="spacer"></div>
     <div class="seg">${[['open', 'Otwarte'], ['all', 'Wszystkie'], ['done', 'Zrobione']].map(([k, l]) => `<button data-action="set" data-k="proj.tstat" data-v="${k}" class="${P0.tstat === k ? 'on' : ''}">${l}</button>`).join('')}</div>
     ${doneIds.length ? `<button class="btn sm ghost danger" data-action="del-done-ids" data-ids="${doneIds.join(',')}">⟨i:trash⟩ Usuń zrobione (${doneIds.length})</button>` : ''}
     <button class="btn sm" data-action="edit-task" data-project="${p.id}">⟨i:plus⟩ Zadanie</button></div>${taskBlocks || '<div class="empty">Brak zadań</div>'}</div>
   <div class="card" id="psMtgs" style="padding:0"><div class="row card-h"><h3 style="margin:0">⟨i:video⟩ Spotkania powiązane z projektem</h3>${eyeBtn('sm')}<span class="muted">${mv.length} spotk. · ${mHours}h</span><div class="spacer"></div>
     <div class="seg">${[['next', 'Nadchodzące'], ['past', 'Minione'], ['all', 'Wszystkie']].map(([k, l]) => `<button data-action="set" data-k="proj.mper" data-v="${k}" class="${P0.mper === k ? 'on' : ''}">${l}</button>`).join('')}</div>
     <button class="btn sm" data-action="new-mtg" data-d="${TODAY()}" data-project="${p.id}">⟨i:plus⟩ Spotkanie</button></div>
     <table class="tbl mtbl"><thead><tr><th>Data</th><th>Godzina</th><th>Temat</th><th>Platforma</th><th class="num">Osoby</th><th></th></tr></thead>
     <tbody>${mRows || `<tr><td colspan="6" class="empty">Brak spotkań${P0.mper === 'next' ? ' nadchodzących' : ''} przypisanych do projektu. Przypisz projekt w oknie spotkania lub masowo na Liście spotkań.</td></tr>`}</tbody></table>
     ${(() => { const sg = mtgSuggest(p); return sg.length ? `<div class="mtg-sug"><b>⟨i:search⟩ Spotkania bez projektu, które pasują do „${esc(p.title)}”</b> <span class="muted">(nazwa, klient lub IC case w temacie)</span>
       ${sg.map(m => `<div class="row sug-row"><span class="nowrap"><b>${fmt(m.start.slice(0, 10))}</b> ${hm(m.start)}</span>${srcBadge(m.source)}<span style="flex:1;min-width:0" class="ell">${pv(esc(m.title))}</span><button class="btn sm" data-action="mtg-assign" data-id="${m.id}" data-project="${p.id}">⟨i:plus⟩ Przypisz</button></div>`).join('')}</div>` : ''; })()}</div>` + projTripsSection(p) + projNotesSection(p);
}

/* ================= WIDOK: Kalendarz ================= */
function calItems(ds) {
  return {
    tasks: tasksC().filter(t => t.due === ds).sort((a, b) => a.prio - b.prio),
    mtgs: meetingsC().filter(m => m.start.slice(0, 10) === ds).sort((a, b) => a.start.localeCompare(b.start))
  };
}
function vCalendar() {
  const C = S.cal, d = C.date;
  const period = C.mode === 'day' ? `${DOWL[d.getDay()]}, ${d.getDate()} ${MONS[d.getMonth()]} ${d.getFullYear()}`
    : C.mode === 'week' ? (() => { const a = sow(d), b = add(a, 6); return `${a.getDate()} ${MONS[a.getMonth()]} – ${b.getDate()} ${MONS[b.getMonth()]} ${b.getFullYear()}`; })()
      : C.mode === 'month' ? `${MON[d.getMonth()]} ${d.getFullYear()}` : `${d.getFullYear()}`;
  const tb = `<div class="cal-toolbar"><div class="seg">${[['day', 'Dzień'], ['week', 'Tydzień'], ['month', 'Miesiąc'], ['year', 'Rok']].map(([k, l]) => `<button data-action="set" data-k="cal.mode" data-v="${k}" class="${C.mode === k ? 'on' : ''}">${l}</button>`).join('')}</div>
    <button class="btn sm" data-action="cal-nav" data-n="-1">⟨i:chev-left⟩</button><button class="btn sm" data-action="cal-today">Dziś</button><button class="btn sm" data-action="cal-nav" data-n="1">⟨i:chev-right⟩</button>
    <span class="period">${period}</span><div class="spacer"></div><button class="btn primary" data-action="new-mtg" data-d="${C.mode === 'day' ? iso(d) : TODAY()}">⟨i:plus⟩ Spotkanie</button></div>`.replace('<div class="spacer"></div>', `<div class="spacer"></div><button class="btn" data-action="cal-to-list" title="Te same spotkania jako tabela">⟨i:list⟩ Lista spotkań z tego zakresu (${calRangeMeetings().length})</button>`);
  const pl = projectsC().filter(p => p.status !== 'done');
  const leg = `<div class="legend"><span><i style="background:var(--job)"></i>zadanie JOB</span><span><i style="background:var(--priv)"></i>zadanie PRIV</span><span><i style="background:var(--hol)"></i>święto</span><span><i style="background:var(--vac)"></i>urlop</span><span><i style="background:var(--we)"></i>weekend</span></div>
   <div class="legend"><b>Spotkania:</b> ${pl.map(p => `<span><i style="background:${projColor(p)}"></i>${esc(p.title)}</span>`).join('')}<span><i style="background:${NOPROJ}"></i>bez projektu</span></div>
   <div class="legend"><span>⟨i:building⟩ JOB</span><span>⟨i:user⟩ PRIV</span><span>${srcBadge('google')} Google Calendar</span><span>${srcBadge('outlook')} Outlook</span><span>${srcBadge('teams')} Teams</span><span>${srcBadge('local')} tylko TM (bez platformy)</span><span><i class="te-leg"></i>realizacja (stoper / wpis ręczny)</span><span class="muted">· kliknij w pustą godzinę, aby dodać spotkanie · kliknij spotkanie, aby edytować</span></div>`;
  return calTabs() + tb + ({day: calDay, week: calWeek, month: calMonth, year: calYear}[C.mode])() + leg;
}
function chipT(t) { return `<span class="chip ${t.ctx} ${t.status === 'done' ? 'done' : ''}" data-action="edit-task" data-id="${t.id}" title="${S.privacy ? `Zadanie · ${t.ctx === 'job' ? 'JOB' : 'PRIV'} · P${t.prio}${t.est ? ' · ' + t.est + 'h' : ''} (dane ukryte)` : esc(t.title)}">${t.status === 'done' ? '⟨i:check⟩ ' : ''}${pv(esc(t.title))}</span>`; }
function mtgTitle(m) { const p = projById(m.projectId); if (S.privacy) return `Spotkanie · ${hm(m.start)}–${hm(m.end)} · ${m.ctx === 'job' ? 'JOB' : 'PRIV'} · ${SRC[m.source]?.[0] || m.source} (dane ukryte)`; return `${m.title} · ${hm(m.start)}–${hm(m.end)} · ${m.ctx === 'job' ? 'JOB' : 'PRIV'} · ${SRC[m.source]?.[0] || m.source}${p ? ' · projekt: ' + p.title : ''}`; }
function chipM(m) { const c = mtgColor(m); return `<span class="chip mtg" data-action="edit-mtg" data-id="${m.id}" style="background:${c}22;border-left-color:${c};color:var(--text)" title="${esc(mtgTitle(m))}">${CTXI[m.ctx]}${srcBadge(m.source)} ${hm(m.start)} ${pv(esc(m.title))}</span>`; }
function calMonth() {
  const d = S.cal.date, first = new Date(d.getFullYear(), d.getMonth(), 1), s = sow(first), td = TODAY();
  let h = `<div class="month">${DOW.map(x => `<div class="dh">${x}</div>`).join('')}`;
  for (let i = 0; i < 42; i++) {
    const x = add(s, i), ds = iso(x), ty = dayType(x), it = calItems(ds), rt = dayTotals(ds).total, all = [...(rt ? [`<span class="chip te" data-action="cal-goto" data-d="${ds}" title="Realizacja (stoper) – kliknij, aby zobaczyć dzień">⟨i:timer⟩ ${fmtH(rt)} realizacji</span>`] : []), ...it.mtgs.map(chipM), ...it.tasks.map(chipT)];
    const hn = holName(x) || (vacOf(x) ? '⟨i:sun⟩ ' + (vacOf(x).note || 'Urlop') : ''), trp = tripOf(x);
    h += `<div class="d ${ty !== 'work' ? ty : ''} ${x.getMonth() !== d.getMonth() ? 'out' : ''} ${ds === td ? 'today' : ''}" data-action="cal-goto" data-d="${ds}">
      <div class="n"><b>${x.getDate()}</b><button class="addm" data-action="new-mtg" data-d="${ds}" title="Dodaj spotkanie">⟨i:plus⟩</button></div>${hn ? `<div class="hn">${esc(hn)}</div>` : ''}${trp ? `<div class="hn trip" data-action="trip-open" data-id="${trp.id}" title="Delegacja – kliknij, aby zobaczyć szczegóły">⟨i:plane⟩ ${pv(esc(trp.place))}</div>` : ''}${all.slice(0, 4).join('')}${all.length > 4 ? `<span class="muted" style="font-size:11px">+${all.length - 4} więcej</span>` : ''}</div>`;
  }
  return h + '</div>';
}
function timeGrid(days) {
  const H0 = 7, H1 = 21, PX = 40, td = TODAY(), now = new Date();
  const cols = `56px repeat(${days.length},minmax(0,1fr))`;
  let h = `<div class="week" style="grid-template-columns:${cols}"><div class="colh"></div>`;
  days.forEach(x => { const ds = iso(x); h += `<div class="colh ${ds === td ? 'today' : ''}" data-action="cal-goto" data-d="${ds}" style="cursor:pointer;${dayType(x) !== 'work' ? `background:var(--${dayType(x)})` : ''}">${DOW[(x.getDay() + 6) % 7]} ${x.getDate()}.${pad(x.getMonth() + 1)}</div>`; });
  h += `<div class="hours" style="border-bottom:1px solid var(--line);padding-top:8px">zad.</div>`;
  days.forEach(x => { const ds = iso(x), it = calItems(ds), hn = holName(x) || (vacOf(x) ? '⟨i:sun⟩ Urlop' : ''); h += `<div class="allday">${hn ? `<div class="hn">${esc(hn)}</div>` : ''}${it.tasks.map(chipT).join('')}</div>`; });
  h += `<div>${Array.from({length: H1 - H0}, (_, i) => `<div class="hline hours">${H0 + i}:00</div>`).join('')}</div>`;
  days.forEach(x => {
    const ds = iso(x), it = calItems(ds);
    h += `<div class="daycol" data-action="cal-slot" data-d="${ds}" title="Kliknij, aby dodać spotkanie" style="${dayType(x) !== 'work' ? `background:var(--${dayType(x)})` : ''}">${Array.from({length: H1 - H0}, () => '<div class="hline"></div>').join('')}`;
    // rozkładanie kolizji: spotkania + wpisy realizacji (stoper)
    const ents = entriesWithLive().filter(e => e.start.slice(0, 10) === ds);
    const items = [...it.mtgs.map(m => ({k: 'm', o: m, s: m.start, e: m.end})), ...ents.map(e => ({k: 'e', o: e, s: e.start, e: e.end}))].sort((p, q) => p.s.localeCompare(q.s));
    const lanes = []; items.forEach(x => { let l = lanes.findIndex(e => e <= x.s); if (l < 0) { l = lanes.length; lanes.push(''); } lanes[l] = x.e; x.lane = l; });
    const nl = Math.max(1, lanes.length);
    items.forEach(x => { const top = (toMin(x.s) / 60 - H0) * PX, ht = Math.max(16, (toMin(x.e) - toMin(x.s)) / 60 * PX - 2);
      if (top + ht < 0 || top > (H1 - H0) * PX) return;
      const pos = `top:${Math.max(0, top)}px;height:${ht}px;left:${3 + x.lane * (100 / nl)}%;width:calc(${100 / nl}% - 6px);right:auto;`;
      if (x.k === 'e') { const e = x.o, c = entColor(e);
        h += `<div class="ev te ${e.live ? 'live' : ''}" ${e.live ? 'data-action="nav" data-view="track"' : `data-action="entry-edit" data-id="${e.id}"`} style="${pos}border-left-color:${c};--tc:${c}" title="Realizacja ${hm(e.start)}–${e.live ? 'teraz' : hm(e.end)} · ${fmtH(entMin(e))}${S.privacy ? '' : ' · ' + esc(e.title)}">
          <div class="evh">⟨i:timer⟩<b>${hm(e.start)}</b><span class="evt">${fmtH(entMin(e))}${e.live ? ' · trwa' : ''}</span></div>${ht < 30 ? '' : pv(esc(e.title))}</div>`; return; }
      const m = x.o, c = mtgColor(m);
      h += `<div class="ev" data-action="edit-mtg" data-id="${m.id}" style="${pos}background:${c}26;border-left-color:${c};color:var(--text)" title="${esc(mtgTitle(m))}">
        <div class="evh">${ctxIcon(m.ctx)}${srcBadge(m.source)}<b>${hm(m.start)}</b>${ht < 30 ? `<span class="evt">${pv(esc(m.title))}</span>` : ''}${notesOf('mtg', m.id).length ? `<span class="evn" title="Notatki: ${notesOf('mtg', m.id).length}">⟨i:note⟩</span>` : ''}${delMtgBtn(m)}</div>${ht < 30 ? '' : pv(esc(m.title))}</div>`; });
    if (ds === td) { const mm = now.getHours() * 60 + now.getMinutes(); if (mm >= H0 * 60 && mm <= H1 * 60) h += `<div class="nowline" style="top:${(mm / 60 - H0) * PX}px"></div>`; }
    h += '</div>';
  });
  return h + '</div>';
}
function calWeek() { const s = sow(S.cal.date); return timeGrid(Array.from({length: 7}, (_, i) => add(s, i))); }
function calDay() {
  const ds = iso(S.cal.date), it = calItems(ds), open = tasksC().filter(t => t.status !== 'done' && t.start <= ds && t.due >= ds && t.due !== ds);
  return `<div class="grid" style="grid-template-columns:minmax(0,1.4fr) minmax(0,1fr)">${timeGrid([S.cal.date])}
    <div class="card"><h3>Zadania z terminem na ten dzień</h3>${it.tasks.map(t => taskRow(t, {pv: true})).join('') || '<div class="empty">Brak</div>'}
    <h3 style="margin-top:16px">W realizacji (trwające)</h3>${open.map(t => taskRow(t, {pv: true})).join('') || '<div class="empty">Brak</div>'}
    <button class="btn sm" style="margin-top:10px" data-action="edit-task" data-due="${ds}">⟨i:plus⟩ Zadanie na ten dzień</button></div></div>`;
}
function calYear() {
  const y = S.cal.date.getFullYear(), td = TODAY(), cnt = {};
  tasksC().forEach(t => t.due && (cnt[t.due] = (cnt[t.due] || 0) + 1));
  meetingsC().forEach(m => { const d = m.start.slice(0, 10); cnt[d] = (cnt[d] || 0) + 1; });
  let h = '<div class="year">';
  for (let m = 0; m < 12; m++) {
    const f = new Date(y, m, 1), s = sow(f);
    h += `<div class="card mini"><h4 data-action="cal-goto" data-d="${iso(f)}" data-mode="month">${MON[m]}</h4><div class="g">${DOW.map(x => `<div class="c muted" style="font-weight:700">${x[0]}</div>`).join('')}`;
    for (let i = 0; i < 42; i++) {
      const x = add(s, i); if (x.getMonth() !== m) { h += '<div></div>'; continue; }
      const ds = iso(x), ty = dayType(x), c = cnt[ds] || 0, heat = c >= 4 ? 'h3' : c >= 2 ? 'h2' : c >= 1 ? 'h1' : '';
      h += `<div class="c ${ty !== 'work' ? ty : ''} ${heat} ${ds === td ? 'today' : ''}" data-action="cal-goto" data-d="${ds}" title="${ds}${holName(x) ? ' · ' + holName(x) : ''}${c ? ` · ${c} pozycji` : ''}">${x.getDate()}</div>`;
    }
    h += '</div></div>';
  }
  return h + '</div><div class="legend"><span><i style="background:#c7d2fe"></i>1</span><span><i style="background:#818cf8"></i>2–3</span><span><i style="background:#4338ca"></i>4+ zadań/spotkań</span></div>';
}

/* ================= WIDOK: Gantt ================= */
function vGantt() {
  const G = S.gantt, CW = 28, days = Array.from({length: G.days}, (_, i) => add(G.start, i)), s0 = iso(G.start), e0 = iso(days[days.length - 1]), td = TODAY();
  const tb = `<div class="cal-toolbar"><div class="seg">${[[28, '4 tyg.'], [42, '6 tyg.'], [91, '3 mies.'], [182, '6 mies.']].map(([n, l]) => `<button data-action="set" data-k="gantt.days" data-v="${n}" class="${G.days === n ? 'on' : ''}">${l}</button>`).join('')}</div>
    <button class="btn sm" data-action="gantt-shift" data-n="-7">⟨i:chev-left⟩ tydz.</button><button class="btn sm" data-action="gantt-today">Dziś</button><button class="btn sm" data-action="gantt-shift" data-n="7">tydz. ⟨i:chev-right⟩</button>
    <div class="seg"><button data-action="set" data-k="gantt.group" data-v="project" class="${G.group === 'project' ? 'on' : ''}">⟨i:folder⟩ Projekty</button><button data-action="set" data-k="gantt.group" data-v="goal" class="${G.group === 'goal' ? 'on' : ''}">⟨i:layers⟩ Podprojekty</button></div>
    ${G.only && projById(G.only) ? `<span class="tag src" style="padding:6px 10px">⟨i:folder⟩ ${esc(projById(G.only).title)} <a href="#" data-action="set" data-k="gantt.only" data-v="">⟨i:x⟩</a></span>` : ''}
    <span class="muted">Przeciągnij pasek, aby przesunąć zadanie · kliknij, aby edytować</span></div>`;
  // nagłówki miesięcy
  let mh = '', cur = null, span = 0;
  days.forEach((d, i) => { const k = d.getMonth(); if (k !== cur) { if (cur != null) mh += `<th colspan="${span}" class="dcell" style="text-align:left;padding-left:4px;font-weight:700">${MON[cur]}</th>`; cur = k; span = 0; } span++; if (i === days.length - 1) mh += `<th colspan="${span}" class="dcell" style="text-align:left;padding-left:4px;font-weight:700">${MON[cur]}</th>`; });
  const cls = d => { const ty = dayType(d); return (iso(d) === td ? 'today ' : '') + (ty !== 'work' ? ty : ''); };
  const bg = `<div class="bg">${days.map(d => `<i class="${cls(d)}"></i>`).join('')}</div>`;
  const bar = (s, e, inner, extra = '', css = '') => { const a = Math.max(0, diffD(s0, s)), b = Math.min(G.days - 1, diffD(s0, e)); if (b < 0 || a > G.days - 1) return ''; return `<div ${extra} style="${css}left:${a * CW + 1}px;width:${(b - a + 1) * CW - 2}px">${inner}</div>`; };
  let rows = '';
  let groups;
  if (G.group === 'project') {
    groups = projectsC().filter(p => !G.only || p.id === G.only).map(p => [Object.assign({isProj: true}, p), projTasks(p)]);
    const loose = tasksC().filter(t => !projById(projOf(t)));
    if (loose.length && !G.only) groups.push([{id: '', title: 'Bez projektu', ctx: S.ctx === 'priv' ? 'priv' : 'job'}, loose]);
  } else {
    groups = goalsC().map(g => [g, S.data.tasks.filter(t => t.goalId === g.id)]);
    const loose = tasksC().filter(t => !t.goalId || !goalById(t.goalId) || !inCtx(goalById(t.goalId)));
    if (loose.length) groups.push([{id: '', title: 'Zadania bez podprojektu', ctx: S.ctx === 'priv' ? 'priv' : 'job'}, loose]);
  }
  groups.forEach(([g, ts]) => {
    const p = !g.id ? null : g.isProj ? projectProgress(g) : goalProgress(g);
    rows += `<tr class="grp"><td class="lbl" ${g.id ? `data-action="${g.isProj ? 'open-project' : 'edit-goal'}" data-id="${g.id}" style="cursor:pointer"` : ''}><span class="dot ${g.ctx}"></span> ${g.isProj ? '⟨i:folder⟩ ' : '⟨i:layers⟩ '}${esc(g.title)}${p != null ? ` · ${p}%` : ''}</td><td colspan="${G.days}"><div class="track" style="height:26px">${bg}
      ${g.id ? bar(g.start, g.due, `<div class="pf" style="width:${p}%"></div>`, `class="gbar ${g.ctx}"`, 'opacity:.35;top:5px;height:16px;cursor:default;') : ''}</div></td></tr>`;
    ts.filter(t => t.start && t.due && t.due >= s0 && t.start <= e0).sort((a, b) => a.start.localeCompare(b.start)).forEach(t => {
      const pr = t.status === 'done' ? 100 : t.status === 'doing' ? 50 : 0;
      rows += `<tr><td class="lbl" data-action="edit-task" data-id="${t.id}" style="cursor:pointer;padding-left:24px" title="${esc(t.title)}${g.isProj && t.goalId && goalById(t.goalId) ? ' · podprojekt: ' + esc(goalById(t.goalId).title) : ''}">${t.status === 'done' ? '⟨i:check⟩ ' : ''}${esc(t.title)}</td><td colspan="${G.days}"><div class="track">${bg}
        ${bar(t.start, t.due, `<div class="pf" style="width:${pr}%"></div><span style="position:relative">${esc(t.title)}${t.est ? ' · ' + t.est + 'h' : ''}</span>`, `class="gbar ${t.ctx} ${t.status === 'done' ? 'done' : ''} ${isLate(t) ? 'late' : ''}" data-gid="${t.id}" title="${esc(t.title)}: ${fmt(t.start)} → ${fmt(t.due)}"`)}</div></td></tr>`;
    });
  });
  return tb + `<div class="gantt"><table><thead><tr><th class="lbl"></th>${mh}</tr><tr><th class="lbl">${G.group === 'project' ? 'Projekt' : 'Podprojekt'} / zadanie</th>${days.map(d => `<th class="dcell ${cls(d)}">${d.getDate()}<br>${DOW[(d.getDay() + 6) % 7][0]}</th>`).join('')}</tr></thead>
    <tbody>${rows || `<tr><td class="empty" colspan="${G.days + 1}">Brak zadań w zakresie</td></tr>`}</tbody></table></div>
    <div class="legend"><span><i style="background:var(--job)"></i>JOB</span><span><i style="background:var(--priv)"></i>PRIV</span><span><i style="outline:2px solid var(--bad)"></i>po terminie</span><span><i style="background:var(--hol)"></i>święto</span><span><i style="background:var(--vac)"></i>urlop</span><span><i style="background:var(--we)"></i>weekend</span></div>`;
}
afterRender.gantt = () => {
  $$('.gbar[data-gid]').forEach(el => {
    let x0 = null, moved = false;
    el.addEventListener('pointerdown', e => { x0 = e.clientX; moved = false; el.setPointerCapture(e.pointerId); el.style.cursor = 'grabbing'; });
    el.addEventListener('pointermove', e => { if (x0 == null) return; const dx = e.clientX - x0; if (Math.abs(dx) > 4) moved = true; el.style.transform = `translateX(${dx}px)`; });
    el.addEventListener('pointerup', e => {
      if (x0 == null) return; const n = Math.round((e.clientX - x0) / 28); x0 = null; el.style.transform = '';
      const t = S.data.tasks.find(x => x.id === el.dataset.gid);
      if (!moved) return editTask(t);
      if (n && t) { t.start = iso(add(P(t.start), n)); t.due = iso(add(P(t.due), n)); save(); render(); toast(`Przesunięto „${t.title}” o ${n} dni`); }
    });
  });
};

/* ================= WIDOK: Capacity ================= */
/* wykres dzienny: spotkania + zadania (+ zaległe) vs capacity dnia */
function capDayChart(n) {
  const {load, late} = capLoads(), mt = meetingHoursByDay(meetingsC());
  const days = [...Array(n)].map((_, i) => { const d = add(today(), i), s = iso(d); return {d, s, cap: capacityOfDay(d, S.ctx), m: mt[s] || 0, t: (load[s] || 0) - (late[s] || 0), l: late[s] || 0, ty: dayType(d)}; });
  const maxV = Math.max(10, ...days.map(x => Math.max(x.cap, x.m + x.t + x.l))) * 1.12;
  const W = 900, Ht = 230, pl = 32, pr = 8, pt = 18, pb = 36, cw = (W - pl - pr) / n, y = v => pt + (Ht - pt - pb) * (1 - v / maxV);
  let svg = `<svg viewBox="0 0 ${W} ${Ht}" width="100%" class="realchart weekchart capday">`;
  for (let h = 0; h <= maxV; h += maxV > 20 ? 4 : 2) svg += `<line x1="${pl}" x2="${W - pr}" y1="${y(h)}" y2="${y(h)}" stroke="var(--line)"/><text x="${pl - 5}" y="${y(h) + 4}" text-anchor="end">${h}h</text>`;
  days.forEach((x, i) => {
    const x0 = pl + i * cw, bx = x0 + cw * .18, bw = cw * .64, tot = x.m + x.t + x.l, over = tot > x.cap + .01;
    if (x.ty !== 'work') svg += `<rect x="${x0 + 1}" y="${pt}" width="${cw - 2}" height="${Ht - pt - pb}" fill="var(--line)" opacity=".3" rx="4"/>`;
    if (x.cap) svg += `<rect x="${bx - 3}" y="${y(x.cap)}" width="${bw + 6}" height="${y(0) - y(x.cap)}" rx="4" class="wk-target"/><line x1="${bx - 5}" x2="${bx + bw + 5}" y1="${y(x.cap)}" y2="${y(x.cap)}" class="wk-tline"/>`;
    let acc = 0; [[x.m, '#5E8BFF', 'spotkania'], [x.t, '#8E929B', 'zadania'], [x.l, '#E5484D', 'zadania zaległe']].forEach(([v, c, l]) => { if (v <= 0) return; svg += `<rect x="${bx}" y="${y(acc + v)}" width="${bw}" height="${Math.max(0, y(acc) - y(acc + v) - 1)}" rx="2" fill="${c}"><title>${fmt(x.s)} · ${l}: ${r1(v)}h</title></rect>`; acc += v; });
    if (tot > .05) svg += `<text x="${bx + bw / 2}" y="${y(Math.max(tot, x.cap)) - 5}" text-anchor="middle" class="wk-val ${over ? 'wk-low' : 'wk-ok'}" style="font-size:${n > 20 ? 9 : 11}px">${r1(tot)}${x.cap ? '/' + r1(x.cap) : ''}</text>`;
    svg += `<text x="${x0 + cw / 2}" y="${Ht - 20}" text-anchor="middle" class="${x.s === TODAY() ? 'wk-today' : ''}" style="font-size:${n > 20 ? 9 : 11}px">${['Nd', 'Pn', 'Wt', 'Śr', 'Cz', 'Pt', 'Sb'][x.d.getDay()]}</text><text x="${x0 + cw / 2}" y="${Ht - 7}" text-anchor="middle" style="font-size:${n > 20 ? 8 : 10}px">${x.d.getDate()}.${pad(x.d.getMonth() + 1)}</text>`;
  });
  const overDays = days.filter(x => x.m + x.t + x.l > x.cap + .01 && (x.cap || x.m + x.t + x.l > 0));
  return {svg: svg + '</svg>', overDays};
}
function vCapacity() {
  const t0 = sow(today()), y = today().getFullYear(), m0 = new Date(y, today().getMonth(), 1), mDays = new Date(y, today().getMonth() + 1, 0).getDate();
  const cm = capacityRange(m0, mDays), nm = new Date(y, today().getMonth() + 1, 1), cnm = capacityRange(nm, new Date(nm.getFullYear(), nm.getMonth() + 1, 0).getDate());
  const used = S.data.vacations.reduce((a, v) => a + vacationWorkdays(v, y), 0), limit = S.data.settings.leaveDays || 26;
  const weeks = Array.from({length: S.cap.weeks}, (_, i) => { const s = add(t0, i * 7); return [s, capacityRange(s, 7)]; });
  const hol = []; for (let d = today(), i = 0; i < 365; i++, d = add(d, 1)) if (holName(d)) hol.push([iso(d), holName(d), isWE(d)]);
  const strip = Array.from({length: 90}, (_, i) => { const d = add(today(), i), ty = dayType(d), tr = tripAt(iso(d)); return `<i class="${ty}${tr ? ' trip' : ''}" title="${iso(d)} · ${{work: 'roboczy', we: 'weekend', hol: holName(d), vac: 'urlop'}[ty]}${tr ? ' · delegacja: ' + esc(tr.place) : ''}"></i>`; }).join('');
  return `<div class="grid g4 kpis" style="margin-bottom:16px">
    <div class="card kpi"><div class="v">${cm.work}</div><div class="l">Dni robocze – ${MON[m0.getMonth()].toLowerCase()} (bez świąt i urlopu)</div></div>
    <div class="card kpi"><div class="v">${r1(cm.cap)}h</div><div class="l">Capacity w miesiącu (${S.ctx.toUpperCase()})</div></div>
    <div class="card kpi"><div class="v" style="color:${cm.util > 1 ? 'var(--bad)' : 'var(--ok)'}">${Math.round(cm.util * 100)}%</div><div class="l">Obciążenie: spotkania ${r1(cm.mtg)}h + zadania ${r1(cm.tasks)}h${cm.late ? ` (w tym zaległe ${r1(cm.late)}h)` : ''}</div></div>
    <div class="card kpi"><div class="v">${limit - used}</div><div class="l">Pozostały urlop ${y} (z ${limit} dni)</div></div></div>
  <div class="grid g2">
   <div class="card" style="grid-column:1/-1">${(() => { const n = S.cap.days || 14, c = capDayChart(n), open = tasksC().filter(t => t.status !== 'done'), noEst = open.filter(t => !(t.est > 0)), lateT = open.filter(isLate);
     return `<div class="row" style="gap:10px"><h3 style="margin:0">⟨i:chart⟩ Obciążenie dzień po dniu – kolejne ${n} dni (${S.ctx.toUpperCase()})</h3><div class="spacer"></div>
       <div class="seg">${[14, 28].map(k => `<button data-action="set" data-k="cap.days" data-v="${k}" class="${n === k ? 'on' : ''}">${k} dni</button>`).join('')}</div></div>
       ${c.svg}
       <div class="row wk-legend"><span><b style="color:#5E8BFF">■</b> spotkania (nakładające się liczone raz)</span><span><b style="color:#8E929B">■</b> zadania – pozostała praca</span><span><b style="color:#E5484D">■</b> zadania zaległe</span><span><i class="wk-sw"></i> capacity dnia</span>
         <span class="spacer"></span>${c.overDays.length ? `<span class="wk-low" style="font-weight:600">Przeładowane dni: ${c.overDays.length}</span>` : '<span class="wk-ok" style="font-weight:600">Brak przeładowanych dni</span>'}</div>
       ${noEst.length || lateT.length ? `<div class="cap-warn">${noEst.length ? `<span>⟨i:alert⟩ <b>${noEst.length}</b> ${noEst.length === 1 ? 'zadanie' : 'zadań'} bez estymacji – liczę po ${defEst()}h <button class="btn sm ghost" data-action="cap-list" data-k="noest">pokaż</button></span>` : ''}${lateT.length ? `<span>⟨i:alert⟩ <b>${lateT.length}</b> zaległych (${r1(lateT.reduce((a, t) => a + taskRemain(t), 0))}h) – obciążają najbliższe dni <button class="btn sm ghost" data-action="cap-list" data-k="late">pokaż</button></span>` : ''}</div>` : ''}`; })()}</div>
   <div class="card"><h3>⟨i:gauge⟩ Capacity vs obciążenie – kolejne ${S.cap.weeks} tygodni</h3>
     ${weeks.map(([s, c]) => capBar(`${s.getDate()}.${pad(s.getMonth() + 1)}–${add(s, 6).getDate()}.${pad(add(s, 6).getMonth() + 1)}`, c)).join('')}
     <div class="legend"><span><i style="background:#5E8BFF"></i>spotkania</span><span><i style="background:#8E929B"></i>zadania (estymacja rozłożona na dni)</span><span><i style="background:var(--text)"></i>limit capacity</span></div>
     <table class="tbl" style="margin-top:12px"><tr><th>Tydzień</th><th class="num">Robocze</th><th class="num">Capacity</th><th class="num">Spotk.</th><th class="num">Zadania</th><th class="num">Wolne</th><th>Uwagi</th></tr>
     ${weeks.map(([s, c]) => `<tr><td>${fmt(iso(s))}</td><td class="num">${c.work}</td><td class="num">${r1(c.cap)}h</td><td class="num">${r1(c.mtg)}h</td><td class="num">${r1(c.tasks)}h</td><td class="num" style="color:${c.cap - c.tasks - c.mtg < 0 ? 'var(--bad)' : 'inherit'}">${r1(c.cap - c.tasks - c.mtg)}h</td><td class="muted">${[...c.hol, c.vac ? `urlop ${c.vac}d` : ''].filter(Boolean).join(', ')}</td></tr>`).join('')}</table></div>
   <div class="grid" style="align-content:start">
    <div class="card"><h3>⟨i:sun⟩ Plan urlopów</h3>
      <form data-form="vac" class="row" style="align-items:flex-end"><label class="f">Od<input class="input" type="date" name="from" required></label><label class="f">Do<input class="input" type="date" name="to" required></label><label class="f" style="flex:1">Opis<input class="input" name="note" placeholder="np. Majówka"></label><label class="row" style="gap:6px;font-size:13px;padding-bottom:8px" title="Tylko dla urlopu jednodniowego"><input type="checkbox" name="half"> pół dnia</label><button class="btn primary">Dodaj</button></form>
      <table class="tbl" style="margin-top:10px">${S.data.vacations.slice().sort((a, b) => a.from.localeCompare(b.from)).map(v => `<tr><td>${fmt(v.from)} → ${fmt(v.to)}</td><td>${esc(v.note || '')}${v.half ? ' <span class="tag src">pół dnia</span>' : ''}</td><td class="num">${vacationWorkdays(v)} dni rob.</td><td class="num"><button class="btn sm ghost danger" data-action="vac-del" data-id="${v.id}">⟨i:x⟩</button></td></tr>`).join('') || '<tr><td class="empty">Brak zaplanowanych urlopów</td></tr>'}</table>
      <div class="row" style="margin-top:10px"><label class="f">Limit dni urlopu / rok<input class="input" type="number" min="0" max="60" style="width:100px" value="${limit}" data-bind="settings.leaveDays"></label>
        <span class="muted">Wykorzystane/zaplanowane w ${y}: <b>${used}</b> dni</span></div></div>
    <div class="card"><h3>⟨i:calendar⟩ Najbliższe 90 dni</h3><div class="strip">${strip}</div>
      <div class="legend"><span><i style="background:#c7d2fe"></i>roboczy</span><span><i style="background:var(--we)"></i>weekend</span><span><i style="background:#fcd34d"></i>święto</span><span><i style="background:#86efac"></i>urlop</span></div>
      <p class="muted" style="margin:10px 0 0">Następny miesiąc (${MON[nm.getMonth()].toLowerCase()}): ${cnm.work} dni roboczych, capacity ${r1(cnm.cap)}h, obciążenie ${Math.round(cnm.util * 100)}%.</p></div>
    <div class="card"><h3>⟨i:flag⟩ Święta (najbliższe 12 mies.)</h3><table class="tbl">${hol.map(([d, n, we]) => `<tr><td>${fmt(d)}</td><td>${DOWL[P(d).getDay()]}</td><td>${esc(n)}${we ? ' <span class="muted">(weekend)</span>' : ''}</td></tr>`).join('')}</table></div>
    <div class="card"><h3>⟨i:settings⟩ Założenia capacity</h3><p class="muted" style="margin:0">JOB: ${CFG.HOURS_JOB_PER_DAY}h/dzień roboczy. PRIV: ${CFG.HOURS_PRIV_WORKDAY}h w dzień roboczy, ${CFG.HOURS_PRIV_WEEKEND}h w weekend/święto/urlop. Zmiana: profil → <a href="#" data-action="open-profile" data-tab="settings">Ustawienia</a>. Zadania bez estymacji liczone po ${defEst()}h, z odjęciem zaraportowanego czasu. W trybie ALL wartości się sumują.</p></div>
   </div></div>`;
}

/* ================= WIDOK: Integracje ================= */
function vInteg() {
  const I = S.integ, ms = S.conn.ms, gg = S.conn.google;
  const card = (k, logo, color, name, sub, on, act) => `<div class="card src-card"><div class="logo" style="background:${color}">${logo}</div><div style="flex:1"><b>${name}</b><div class="muted" style="font-size:12px">${sub}</div>
    <div class="status ${on ? 'on' : 'demo'}">${on ? '● Połączono' : '● Tryb demo'}</div></div><button class="btn sm" data-action="${act}">${on ? 'Odśwież' : 'Połącz'}</button></div>`;
  const tabs = [['meetings', '⟨i:video⟩ Spotkania'], ...(ms ? [['avail', '⟨i:users⟩ Dostępność']] : []), ['new', '⟨i:plus⟩ Nowe spotkanie'], ['mail', '⟨i:mail⟩ Napisz e-mail']];
  if (I.tab === 'avail' && !ms) I.tab = 'meetings';
  let body = '';
  if (I.tab === 'meetings') {
    const ms7 = meetingsC().filter(m => m.start.slice(0, 10) >= TODAY() && m.start.slice(0, 10) <= iso(add(today(), 7))).sort((a, b) => a.start.localeCompare(b.start));
    const by = {}; ms7.forEach(m => (by[m.start.slice(0, 10)] ||= []).push(m));
    body = `<div class="row" style="justify-content:flex-end;margin-bottom:6px;gap:8px">${importBtn('sm')}${eyeBtn('sm')}</div>` + Object.entries(by).map(([d, l]) => `<div class="group-h">${DOWL[P(d).getDay()]}, ${fmt(d)} · ${l.length} spotk. · ${r1(l.reduce((a, m) => a + (toMin(m.end) - toMin(m.start)) / 60, 0))}h</div>${l.map(meetingRow).join('')}`).join('') || '<div class="empty">Brak spotkań w najbliższych 7 dniach</div>';
  } else if (I.tab === 'avail') {
    const sel = I.people.length ? I.people : PEOPLE.slice(0, 3).map(p => p.email);
    I.people = sel;
    body = `<div class="row" style="margin-bottom:10px"><label class="f">Dzień<input class="input" type="date" value="${I.fbDate}" data-bind="integ.fbDate"></label>
      <div class="row" style="gap:6px">${PEOPLE.map(p => `<label class="tag src" style="cursor:pointer;padding:6px 10px"><input type="checkbox" data-action="fb-person" value="${p.email}" ${sel.includes(p.email) ? 'checked' : ''}> ${esc(p.name)}</label>`).join('')}</div>
      <button class="btn sm" data-action="fb-load">⟨i:refresh⟩ Pobierz dostępność</button></div><div id="fbGrid">${fbGrid(I.fb)}</div>
      <div class="legend"><span><i style="background:#fca5a5"></i>zajęty</span><span><i style="background:#fde68a"></i>wstępnie</span><span><i style="background:#c4b5fd"></i>poza biurem</span><span><i style="background:#bbf7d0"></i>wszyscy wolni – kliknij, aby zaplanować</span></div>`;
  } else if (I.tab === 'new') {
    const p = I.prefill || {};
    body = `<form data-form="meeting" class="grid g2">
      <label class="f" style="grid-column:1/-1">Temat<input class="input" name="title" required value="${esc(p.title || '')}" placeholder="np. Przegląd DLS – integracje"></label>
      <label class="f" style="grid-column:1/-1">Uczestnicy (e-maile, przecinek)<input class="input" name="att" value="${esc((p.att || []).join(', '))}" list="peopleList" placeholder="anna.nowak@infoconsulting.pl, …"></label>
      <datalist id="peopleList">${PEOPLE.map(x => `<option value="${x.email}">${esc(x.name)}</option>`).join('')}</datalist>
      <label class="f">Data<input class="input" type="date" name="date" required value="${p.date || TODAY()}"></label>
      <div class="row"><label class="f">Start<input class="input" type="time" name="time" required value="${p.time || '10:00'}"></label><label class="f">Czas<select class="input" name="dur">${[15, 30, 45, 60, 90, 120].map(n => `<option value="${n}" ${n === (p.dur || 30) ? 'selected' : ''}>${n} min</option>`).join('')}</select></label></div>
      <label class="f">Kanał<select class="input" name="via"><option value="teams">Microsoft Teams (Outlook)</option><option value="outlook">Outlook – spotkanie stacjonarne</option><option value="google">Google Calendar (Meet) – prywatne</option></select></label>
      <label class="f">Lokalizacja<input class="input" name="loc" placeholder="opcjonalnie"></label>
      <label class="f" style="grid-column:1/-1">Projekt (kolor w kalendarzu)<select class="input" name="projectId">${projOptions(null)}</select></label>
      <label class="f" style="grid-column:1/-1">Agenda<textarea class="input" name="body" rows="4" placeholder="Agenda spotkania…">${esc(p.body || '')}</textarea></label>
      <div class="row" style="grid-column:1/-1"><button class="btn primary">⟨i:calendar⟩ Utwórz spotkanie</button><span class="muted" style="font-size:12px">${ms || gg ? 'Spotkanie zostanie utworzone przez API.' : 'Tryb demo: otworzy się kreator w Teams/Outlook/Google z wypełnionymi polami, a spotkanie zapisze się w TM.'}</span></div></form>`;
  } else {
    body = `<form data-form="mail" class="grid">
      <div class="row"><span class="muted">Szablony:</span><button type="button" class="btn sm" data-action="mail-tpl" data-t="status">Status zadań na dziś</button><button type="button" class="btn sm" data-action="mail-tpl" data-t="week">Plan tygodnia</button><button type="button" class="btn sm" data-action="mail-tpl" data-t="late">Zaległe zadania</button></div>
      <label class="f">Do<input class="input" name="to" required list="peopleList2" placeholder="adres@domena.pl"></label><datalist id="peopleList2">${PEOPLE.map(x => `<option value="${x.email}">${esc(x.name)}</option>`).join('')}</datalist>
      <label class="f">DW (CC)<input class="input" name="cc"></label>
      <label class="f">Temat<input class="input" name="subject" required></label>
      <label class="f">Treść<textarea class="input" name="body" rows="10"></textarea></label>
      <div class="row"><label class="f">Wyślij przez<select class="input" name="via"><option value="outlook">Outlook (firmowe)</option><option value="gmail">Gmail (prywatne)</option><option value="mailto">Domyślny klient poczty</option></select></label>
      <button class="btn primary" style="align-self:flex-end">⟨i:mail⟩ ${ms || gg ? 'Wyślij' : 'Otwórz w poczcie'}</button></div></form>`;
  }
  const anyCfg = CFG.MS_CLIENT_ID || CFG.GOOGLE_CLIENT_ID;
  return `${anyCfg ? `<div class="grid g3" style="margin-bottom:16px">
    ${card('outlook', 'O', '#0F6CBD', 'Outlook / Microsoft 365', 'Kalendarz, poczta, dostępność (Graph) · JOB', ms, 'connect-ms')}
    ${card('teams', 'T', '#5B5FC7', 'Microsoft Teams', 'Spotkania online · JOB', ms, 'connect-ms')}
    ${card('gmail', 'G', '#EA4335', 'Gmail / Google Calendar', 'Poczta i kalendarz prywatny · PRIV', gg, 'connect-google')}</div>` : ''}
    <div class="seg" style="margin-bottom:12px">${tabs.map(([k, l]) => `<button data-action="set" data-k="integ.tab" data-v="${k}" class="${I.tab === k ? 'on' : ''}">${l}</button>`).join('')}</div>
    <div class="card">${body}</div>`;
}
const SLOTS = Array.from({length: 20}, (_, i) => `${pad(8 + Math.floor(i / 2))}:${i % 2 ? '30' : '00'}`);
function hash(s) { let h = 2166136261; for (const c of s) h = Math.imul(h ^ c.charCodeAt(0), 16777619); return h >>> 0; }
function demoAvail(email, ds) {
  const d = P(ds); if (isWE(d) || holName(d)) return Array(20).fill(3);
  let h = hash(email + ds); const a = [];
  for (let i = 0; i < 20; i++) { h = (Math.imul(h, 1103515245) + 12345) >>> 0; const r = (h >>> 8) % 100; a.push(i === 2 && email.includes('infoconsulting') ? 2 : r < 20 ? 2 : r < 26 ? 1 : 0); }
  for (let i = 1; i < 19; i++) if (a[i - 1] === 2 && a[i + 1] === 2) a[i] = 2; // bloki
  return a;
}
function myAvail(ds) {
  const a = Array(20).fill(0); if (!isWork(P(ds))) return Array(20).fill(3);
  S.data.meetings.filter(m => m.ctx === 'job' && m.start.slice(0, 10) === ds).forEach(m => SLOTS.forEach((s, i) => { const t = 480 + i * 30; if (t < toMin(m.end) && t + 30 > toMin(m.start)) a[i] = 2; }));
  return a;
}
function fbGrid(fb) {
  const I = S.integ, ds = I.fbDate;
  fb = fb && fb.date === ds ? fb.rows : [{name: 'Ja', a: myAvail(ds)}, ...I.people.map(e => ({name: PEOPLE.find(p => p.email === e)?.name || e, a: demoAvail(e, ds)}))];
  const cls = v => ['', 'tent', 'busy', 'oof', 'tent'][v] || '';
  let h = `<div class="fb" style="grid-template-columns:160px repeat(20,minmax(22px,1fr))"><div class="h"></div>${SLOTS.map((s, i) => `<div class="h">${i % 2 ? '' : s}</div>`).join('')}`;
  fb.forEach(r => { h += `<div class="n">${esc(r.name)}</div>${r.a.map(v => `<div class="s ${cls(v)}"></div>`).join('')}`; });
  h += `<div class="n">⟨i:check⟩ Wspólnie wolne</div>${SLOTS.map((s, i) => fb.every(r => !r.a[i]) ? `<div class="s free-all" data-action="fb-slot" data-t="${s}" title="Zaplanuj ${s}"></div>` : '<div class="s"></div>').join('')}</div>`;
  return h;
}

/* ================= API: Microsoft Graph / Google ================= */
function loadScript(src) { return new Promise((res, rej) => { if ($(`script[src="${src}"]`)) return res(); const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = rej; document.head.append(s); }); }
const MS = {
  app: null, acct: null,
  async init() {
    if (!CFG.MS_CLIENT_ID) return false;
    await loadScript('https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js');
    this.app = new msal.PublicClientApplication({auth: {clientId: CFG.MS_CLIENT_ID, authority: 'https://login.microsoftonline.com/' + (CFG.MS_TENANT || 'common'), redirectUri: location.origin + location.pathname}, cache: {cacheLocation: 'localStorage'}});
    if (this.app.initialize) await this.app.initialize();
    this.acct = this.app.getAllAccounts()[0] || null; return true;
  },
  async login() { if (!this.app && !(await this.init())) return false; const r = await this.app.loginPopup({scopes: CFG.MS_SCOPES, loginHint: S.user.email}); this.acct = r.account; return true; },
  async token() { try { return (await this.app.acquireTokenSilent({scopes: CFG.MS_SCOPES, account: this.acct})).accessToken; } catch { return (await this.app.acquireTokenPopup({scopes: CFG.MS_SCOPES})).accessToken; } },
  async g(path, opt = {}) {
    const r = await fetch('https://graph.microsoft.com/v1.0' + path, {...opt, headers: {Authorization: 'Bearer ' + await this.token(), 'Content-Type': 'application/json', Prefer: `outlook.timezone="${CFG.MS_TIMEZONE}"`}});
    if (!r.ok) throw new Error('Graph ' + r.status + ' ' + await r.text()); return r.status === 202 || r.status === 204 ? null : r.json();
  },
  async events(from, to) {
    const j = await this.g(`/me/calendarView?startDateTime=${from}T00:00:00&endDateTime=${to}T23:59:59&$top=250&$select=subject,start,end,attendees,isOnlineMeeting,onlineMeeting,location`);
    return j.value.map(e => ({id: 'ms-' + e.id, title: e.subject, start: e.start.dateTime.slice(0, 16), end: e.end.dateTime.slice(0, 16), source: e.isOnlineMeeting ? 'teams' : 'outlook', ctx: 'job',
      attendees: (e.attendees || []).map(a => a.emailAddress.address), location: e.location?.displayName || '', link: e.onlineMeeting?.joinUrl || '', ext: true}));
  },
  async schedule(emails, ds) {
    const j = await this.g('/me/calendar/getSchedule', {method: 'POST', body: JSON.stringify({schedules: emails, startTime: {dateTime: ds + 'T08:00:00', timeZone: CFG.MS_TIMEZONE}, endTime: {dateTime: ds + 'T18:00:00', timeZone: CFG.MS_TIMEZONE}, availabilityViewInterval: 30})});
    return j.value.map(v => ({name: PEOPLE.find(p => p.email === v.scheduleId)?.name || v.scheduleId, a: [...(v.availabilityView || '').padEnd(20, '0')].slice(0, 20).map(Number)}));
  },
  createEvent(m) { return this.g('/me/events', {method: 'POST', body: JSON.stringify({subject: m.title, body: {contentType: 'Text', content: m.body || ''}, start: {dateTime: m.start + ':00', timeZone: CFG.MS_TIMEZONE}, end: {dateTime: m.end + ':00', timeZone: CFG.MS_TIMEZONE}, location: {displayName: m.location || ''}, attendees: m.attendees.map(a => ({emailAddress: {address: a}, type: 'required'})), isOnlineMeeting: m.source === 'teams', onlineMeetingProvider: m.source === 'teams' ? 'teamsForBusiness' : undefined})}); },
  sendMail(x) { const rc = s => s.split(/[,;]/).map(a => a.trim()).filter(Boolean).map(a => ({emailAddress: {address: a}})); return this.g('/me/sendMail', {method: 'POST', body: JSON.stringify({message: {subject: x.subject, body: {contentType: 'Text', content: x.body}, toRecipients: rc(x.to), ccRecipients: rc(x.cc || '')}, saveToSentItems: true})}); }
};
const GG = {
  token: null, exp: 0,
  async ensure() {
    if (!CFG.GOOGLE_CLIENT_ID) return false;
    if (this.token && Date.now() < this.exp) return true;
    await loadScript('https://accounts.google.com/gsi/client');
    return new Promise((res, rej) => google.accounts.oauth2.initTokenClient({client_id: CFG.GOOGLE_CLIENT_ID, scope: CFG.GOOGLE_SCOPES, login_hint: S.user.email,
      callback: r => { if (r.error) return rej(r); this.token = r.access_token; this.exp = Date.now() + (r.expires_in - 60) * 1000; res(true); }}).requestAccessToken());
  },
  async api(url, opt = {}) { await this.ensure(); const r = await fetch(url, {...opt, headers: {Authorization: 'Bearer ' + this.token, 'Content-Type': 'application/json'}}); if (!r.ok) throw new Error('Google ' + r.status); return r.json(); },
  async events(from, to) {
    const j = await this.api(`https://www.googleapis.com/calendar/v3/calendars/primary/events?singleEvents=true&orderBy=startTime&maxResults=250&timeMin=${new Date(from).toISOString()}&timeMax=${add(P(to), 1).toISOString()}`);
    const loc = s => { const d = new Date(s); return `${iso(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`; };
    return (j.items || []).filter(e => e.start?.dateTime).map(e => ({id: 'g-' + e.id, title: e.summary || '(bez tytułu)', start: loc(e.start.dateTime), end: loc(e.end.dateTime), source: 'google', ctx: 'priv', attendees: (e.attendees || []).map(a => a.email), location: e.location || '', link: e.hangoutLink || e.htmlLink || '', ext: true}));
  },
  createEvent(m) {
    const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
    return this.api('https://www.googleapis.com/calendar/v3/calendars/primary/events?conferenceDataVersion=1&sendUpdates=all', {method: 'POST', body: JSON.stringify({summary: m.title, description: m.body, location: m.location, start: {dateTime: m.start + ':00', timeZone: tz}, end: {dateTime: m.end + ':00', timeZone: tz}, attendees: m.attendees.map(email => ({email})), conferenceData: {createRequest: {requestId: uid(), conferenceSolutionKey: {type: 'hangoutsMeet'}}}})});
  },
  sendMail(x) {
    const b64 = s => btoa(unescape(encodeURIComponent(s)));
    const raw = [`To: ${x.to}`, x.cc ? `Cc: ${x.cc}` : '', `Subject: =?UTF-8?B?${b64(x.subject)}?=`, 'MIME-Version: 1.0', 'Content-Type: text/plain; charset=UTF-8', 'Content-Transfer-Encoding: base64', '', b64(x.body)].filter((l, i) => l || i > 3).join('\r\n');
    return this.api('https://gmail.googleapis.com/gmail/v1/users/me/messages/send', {method: 'POST', body: JSON.stringify({raw: b64(raw).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')})});
  }
};
async function syncExternal() {
  const from = iso(add(sow(today()), -7)), to = iso(add(today(), 35)); let ext = [];
  try { if (S.conn.ms) ext = ext.concat(await MS.events(from, to)); } catch (e) { toast('Outlook: ' + e.message); }
  try { if (S.conn.google) ext = ext.concat(await GG.events(from, to)); } catch (e) { toast('Google: ' + e.message); }
  S.data.meetings = S.data.meetings.filter(m => !m.ext && !(m.demo && ((S.conn.ms && m.ctx === 'job') || (S.conn.google && m.ctx === 'priv'))))
    .concat(ext); save(); render();
}

/* ================= Modale: zadanie, podprojekt, szybki kreator, konto ================= */
/* ================= SZABLONY ZADAŃ ================= */
const tplAll = () => (S.data.templates ||= []);
const tplById = id => tplAll().find(t => t.id === id);
const tplCats = () => [...new Set(tplAll().filter(inCtx).map(t => t.category || 'Bez kategorii'))].sort((a, b) => a.localeCompare(b, 'pl'));
const isoWeek = d => { const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())); const n = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - n); const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1)); return Math.ceil(((t - y0) / 864e5 + 1) / 7); };
/* zmienne w tytule szablonu: {data} {miesiąc} {tydzień} {rok} – liczone od terminu zadania */
function tplTitle(tp, dueIso) {
  const d = P(dueIso || TODAY()), mon = `${MON[d.getMonth()].toLowerCase()} ${d.getFullYear()}`;
  return (tp.title || tp.name).replace(/\{data\}/gi, `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()}`).replace(/\{miesi[aą]c\}/gi, mon)
    .replace(/\{tydzie[nń]\}/gi, String(isoWeek(d))).replace(/\{rok\}/gi, String(d.getFullYear()));
}
const DUE_RULES = [['', 'bez domyślnego (dziś)'], ['d1', 'jutro'], ['wd2', 'za 2 dni robocze'], ['wd5', 'za 5 dni roboczych'], ['eow', 'koniec tygodnia (piątek)'], ['eom', 'ostatni dzień roboczy miesiąca'], ['d7', 'za tydzień'], ['d14', 'za 2 tygodnie']];
function dueFromRule(rule) {
  const t = today(); if (!rule) return t;
  if (rule === 'eow') { let d = add(sow(t), 4); if (d < t) d = add(d, 7); while (!isWork(d) && d > t) d = add(d, -1); return d; }
  if (rule === 'eom') { let d = new Date(t.getFullYear(), t.getMonth() + 1, 0); while (!isWork(d) && d > t) d = add(d, -1); return d; }
  let m = rule.match(/^wd(\d+)$/); if (m) { let d = t, n = +m[1]; while (n > 0) { d = add(d, 1); if (isWork(d)) n--; } return d; }
  m = rule.match(/^d(\d+)$/); if (m) return add(t, +m[1]);
  return t;
}
function startFor(dueIso, est) { let start = P(dueIso); if (est > CFG.HOURS_JOB_PER_DAY) { let need = Math.ceil(est / CFG.HOURS_JOB_PER_DAY) - 1; while (need > 0) { start = add(start, -1); if (isWork(start)) need--; } if (start < today()) start = today(); } return iso(start); }
/* szybki kreator z szablonem: szablon daje dane bazowe, wpisane skróty (#job !1 2h ^projekt @podprojekt daty) je nadpisują */
function quickResolve(txt) {
  const r = parseQuick(txt), tp = S.qaTpl && tplById(S.qaTpl); if (!tp) return r;
  const nt = norm(txt), explicitCtx = /(^|\s)#(priv|prywatne|p|dom|home|job|praca|j|work|firma)(\s|$)/.test(nt);
  if (!explicitCtx) r.ctx = tp.ctx || r.ctx;
  if (!/(^|\s)![1-3](\s|$)/.test(txt)) r.prio = tp.prio || r.prio;
  if (r.est == null && tp.est) r.est = tp.est;
  if (!r.projectId && projById(tp.projectId)) { r.projectId = tp.projectId; if (!explicitCtx) r.ctx = projById(tp.projectId).ctx; }
  if (!r.goalId && goalById(tp.goalId)) { r.goalId = tp.goalId; r.projectId = goalById(tp.goalId).projectId || r.projectId; }
  if (!r.dates.length && tp.dueRule) r.due = iso(dueFromRule(tp.dueRule));
  if (r.dates.length < 2) r.start = startFor(r.due, r.est || 0);
  const base = tplTitle(tp, r.due); r.title = r.title ? `${base} – ${r.title}` : base; r.notes = tp.notes || ''; r.tpl = tp;
  return r;
}
function tplUsed(tp) { if (!tp) return; tp.uses = (tp.uses || 0) + 1; tp.lastUsed = new Date().toISOString(); }
function tplChips(q) {
  const qq = norm(q || ''), list = tplAll().filter(inCtx);
  if (qq) { const f = list.filter(t => norm([t.name, t.category, t.title].join(' ')).includes(qq)).sort((a, b) => (b.uses || 0) - (a.uses || 0)); return f.length ? f.slice(0, 12).map((t, i) => tplChip(t, i === 0)).join('') : '<span class="muted" style="font-size:12px">Brak szablonu – Enter doda zadanie bez szablonu</span>'; }
  const cat = S.qaCat || '__top';
  const l = cat === '__top' ? [...list].sort((a, b) => (b.lastUsed || '').localeCompare(a.lastUsed || '') || (b.uses || 0) - (a.uses || 0)).slice(0, 10) : list.filter(t => (t.category || 'Bez kategorii') === cat).sort((a, b) => a.name.localeCompare(b.name, 'pl'));
  return l.map(t => tplChip(t)).join('') || '<span class="muted" style="font-size:12px">Brak szablonów w tej kategorii</span>';
}
const tplChip = (t, first) => `<button type="button" class="tplchip ${t.ctx} ${S.qaTpl === t.id ? 'on' : ''} ${first ? 'first' : ''}" data-action="qa-tpl" data-id="${t.id}" title="${esc(tplTitle(t))}${t.est ? ' · ' + t.est + 'h' : ''}">${t.ctx === 'job' ? '⟨i:building⟩' : '⟨i:user⟩'}${esc(t.name)}${t.est ? `<span>${t.est}h</span>` : ''}</button>`;
function quickAdd(preTpl) {
  S.qaTpl = preTpl || null; S.qaCat = S.qaCat || '__top';
  modal(`<h3>⟨i:zap⟩ Szybkie zadanie</h3>
    <div id="qaTplSel"></div>
    <input class="input qa-input" id="qa" autocomplete="off">
    <div class="qa-preview" id="qaPrev"></div>
    <div class="qa-tpls"><div class="row qa-cats" id="qaCats"></div><div class="tplchips" id="qaChips"></div></div>
    <div class="hint">Wybierz szablon (lub wpisz <code>/nazwa</code> i Enter), potem podaj tylko termin i czas, np. <code>pt 2h</code>. Skróty: <code>#job</code> <code>#priv</code> · <code>!1</code>–<code>!3</code> · <code>2h</code> <code>30m</code> ·
    <code>dziś</code> <code>jutro</code> <code>pt</code> <code>24.11</code> <code>+3d</code> · projekt <code>^słowo</code> · podprojekt <code>@słowo</code>. <b>Shift+Enter</b> – zapisz i dodaj kolejne. Backspace w pustym polu usuwa szablon.</div>
    <div class="row" style="margin-top:14px;justify-content:space-between"><button class="btn ghost" data-action="qa-manage">⟨i:template⟩ Zarządzaj szablonami</button><div class="row"><button class="btn" data-action="qa-more">Więcej pól…</button><button class="btn primary" data-action="qa-save">Dodaj</button></div></div>`, m => {
    const inp = $('#qa', m); inp.focus(); qaRefresh();
    inp.addEventListener('input', qaRefresh);
    inp.addEventListener('keydown', e => {
      const v = inp.value;
      if ((e.key === 'Enter' || e.key === 'Tab') && v.trim().startsWith('/')) { e.preventDefault(); const f = $('#qaChips .tplchip.first'); if (f) { S.qaTpl = f.dataset.id; inp.value = ''; qaRefresh(); } else if (e.key === 'Enter') { inp.value = v.replace(/^\s*\/\S*\s*/, ''); qaRefresh(); } return; }
      if (e.key === 'Backspace' && !v && S.qaTpl) { S.qaTpl = null; qaRefresh(); return; }
      if (e.key === 'Enter') { e.preventDefault(); qaSave(e.shiftKey); }
    });
  });
}
function qaRefresh() {
  const inp = $('#qa'); if (!inp) return; const v = inp.value, slash = v.trim().startsWith('/'), tp = tplById(S.qaTpl);
  $('#qaTplSel').innerHTML = H(tp ? `<div class="qa-tplsel">⟨i:template⟩ Szablon: <b>${esc(tp.name)}</b> <span class="muted">· ${esc(tp.category || '')}</span><button type="button" class="btn sm ghost" data-action="qa-tpl-clear" title="Usuń szablon">⟨i:x⟩</button></div>` : '');
  inp.placeholder = tp ? 'Termin i czas, np. „pt 2h” (opcjonalnie dopisek do tytułu)' : 'np. Wysłać ofertę jutro !1 2h ^Alfa  ·  /raport – wybór szablonu';
  const cats = tplCats();
  $('#qaCats').innerHTML = H(slash ? `<span class="muted" style="font-size:12px">⟨i:search⟩ Szablony pasujące do „${esc(v.trim().slice(1))}” – Enter wybiera pierwszy</span>` : [['__top', 'Ostatnio używane'], ...cats.map(c => [c, c])].map(([k, l]) => `<button type="button" class="qcat ${S.qaCat === k ? 'on' : ''}" data-action="qa-cat" data-c="${esc(k)}">${esc(l)}</button>`).join(''));
  $('#qaChips').innerHTML = H(tplChips(slash ? v.trim().slice(1) : ''));
  if (slash) { $('#qaPrev').innerHTML = ''; return; }
  const show = v.trim() || tp;
  const r = show ? quickResolve(v) : null, g = r && r.goalId && goalById(r.goalId), pj = r && projById(r.projectId);
  $('#qaPrev').innerHTML = H(r ? `${ctxTag(r.ctx)}<span class="tag p${r.prio}">P${r.prio}</span><span class="tag src">⟨i:calendar⟩ ${r.start !== r.due ? fmt(r.start) + ' → ' : ''}${fmt(r.due)}</span>${r.est ? `<span class="tag src">⟨i:clock⟩ ${r.est}h</span>` : `<span class="tag src muted">⟨i:clock⟩ czas?</span>`}${pj ? `<span class="tag src">⟨i:folder⟩ ${esc(pj.title)}</span>` : ''}${g ? `<span class="tag src">⟨i:layers⟩ ${esc(g.title)}</span>` : ''}<b style="margin-left:4px">${esc(r.title)}</b>` : '');
}
function qaSave(again) {
  const inp = $('#qa'); if (!inp) return; const tp = tplById(S.qaTpl);
  if (!inp.value.trim() && !tp) return;
  const r = quickResolve(inp.value); if (!r.title) return toast('Podaj nazwę zadania');
  S.data.tasks.push({id: uid(), title: r.title, ctx: r.ctx, goalId: r.goalId, projectId: r.projectId, start: r.start, due: r.due, status: 'todo', prio: r.prio, est: r.est || 0, notes: r.notes || '', createdAt: TODAY(), doneAt: null, tplId: tp ? tp.id : undefined});
  tplUsed(tp); save(); toast(`Dodano: ${r.title} (${r.ctx.toUpperCase()}, ${fmt(r.due)}${r.est ? ', ' + r.est + 'h' : ''})`);
  if (again) { inp.value = ''; render(); qaRefresh(); inp.focus(); } else { closeModal(); render(); }
}
/* edytor szablonu */
function editTemplate(tp, preset = {}) {
  const isNew = !tp; tp = tp || Object.assign({name: '', category: S.tplCat && S.tplCat !== '__all' ? S.tplCat : '', ctx: S.ctx === 'all' ? 'job' : S.ctx, projectId: null, goalId: null, title: '', prio: 2, est: '', dueRule: '', notes: ''}, preset);
  const cats = [...new Set(tplAll().map(t => t.category).filter(Boolean))];
  modal(`<h3>⟨i:template⟩ ${isNew ? 'Nowy szablon zadania' : 'Edycja szablonu'}</h3><form data-form="tpl" data-id="${isNew ? '' : tp.id}" class="grid g2">
    <label class="f">Nazwa szablonu<input class="input" name="name" required value="${esc(tp.name)}" placeholder="np. Raport miesięczny"></label>
    <label class="f">Kategoria<input class="input" name="category" list="tplCatList" value="${esc(tp.category || '')}" placeholder="np. Raportowanie"><datalist id="tplCatList">${cats.map(c => `<option value="${esc(c)}">`).join('')}</datalist></label>
    <label class="f" style="grid-column:1/-1">Tytuł tworzonego zadania<input class="input" name="title" value="${esc(tp.title || '')}" placeholder="np. Raport dla PM – {miesiąc}"><span class="hint" style="font-weight:400">Zmienne: <code>{data}</code> <code>{miesiąc}</code> <code>{tydzień}</code> <code>{rok}</code> – liczone od terminu zadania. Puste = nazwa szablonu.</span></label>
    <label class="f">Kontekst<select class="input" name="ctx"><option value="job" ${tp.ctx === 'job' ? 'selected' : ''}>JOB – firmowe</option><option value="priv" ${tp.ctx === 'priv' ? 'selected' : ''}>PRIV – prywatne</option></select></label>
    <label class="f">Projekt<select class="input" name="projectId">${projOptions(tp.projectId)}</select></label>
    <label class="f" style="grid-column:1/-1">Podprojekt<select class="input" name="goalId"><option value="">— bez podprojektu —</option>${S.data.goals.map(g => `<option value="${g.id}" data-proj="${g.projectId || ''}" data-ctx="${g.ctx}" ${tp.goalId === g.id ? 'selected' : ''}>[${g.ctx.toUpperCase()}] ${esc(g.title)}</option>`).join('')}</select></label>
    <div class="row"><label class="f">Priorytet<select class="input" name="prio">${[1, 2, 3].map(n => `<option ${+tp.prio === n ? 'selected' : ''}>${n}</option>`).join('')}</select></label>
      <label class="f">Czas (h) – opcjonalnie<input class="input" type="number" step="0.25" min="0" name="est" value="${tp.est || ''}" placeholder="—" style="width:130px"></label></div>
    <label class="f">Domyślny termin<select class="input" name="dueRule">${DUE_RULES.map(([k, l]) => `<option value="${k}" ${tp.dueRule === k ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
    <label class="f" style="grid-column:1/-1">Opis zadania (wstawiany do zadania)<textarea class="input" name="notes" rows="4" placeholder="np. lista kroków / checklista">${esc(tp.notes || '')}</textarea></label>
    <div class="hint" style="grid-column:1/-1;margin:0">Termin i czas możesz zawsze zmienić przy tworzeniu zadania – szablon podpowiada wartości domyślne.</div>
    <div class="row" style="grid-column:1/-1;justify-content:space-between">${isNew ? '<span></span>' : `<div class="row"><button type="button" class="btn danger" data-action="tpl-del" data-id="${tp.id}">⟨i:trash⟩ Usuń</button><button type="button" class="btn" data-action="tpl-dup" data-id="${tp.id}">Duplikuj</button></div>`}
    <div class="row"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">Zapisz szablon</button></div></div></form>`, m => { $('[name=name]', m).focus(); linkForm(m, true); });
}
/* ekran Zadania › Szablony zadań */
function taskTabs() {
  return ''; // szablony zadań przeniesione do Ustawień (zębatka)
  return `<div class="seg caltabs" style="margin-bottom:12px"><button data-action="nav" data-view="tasks" class="${S.view === 'tasks' ? 'on' : ''}">⟨i:list⟩ Lista zadań</button><button data-action="nav" data-view="templates" class="${S.view === 'templates' ? 'on' : ''}">⟨i:template⟩ Szablony zadań (${tplAll().filter(inCtx).length})</button></div>`;
}
function vTemplates() {
  const cat = S.tplCat || '__all', q = norm(S.tplQ || ''), sort = S.tplSort || 'uses';
  let l = tplAll().filter(inCtx);
  const cats = tplCats(), cnt = c => l.filter(t => (t.category || 'Bez kategorii') === c).length, total = l.length;
  if (cat !== '__all') l = l.filter(t => (t.category || 'Bez kategorii') === cat);
  if (q) l = l.filter(t => norm([t.name, t.category, t.title, t.notes].join(' ')).includes(q));
  l.sort(sort === 'name' ? (a, b) => a.name.localeCompare(b.name, 'pl') : (a, b) => (b.uses || 0) - (a.uses || 0) || a.name.localeCompare(b.name, 'pl'));
  const groups = cat === '__all' ? cats.map(c => [c, l.filter(t => (t.category || 'Bez kategorii') === c)]).filter(g => g[1].length) : [[cat, l]];
  const card = t => { const pj = projById(t.projectId), g = goalById(t.goalId); return `<div class="card tplcard ${t.ctx}">
    <div class="row" style="gap:8px"><b class="tpl-name" data-action="tpl-edit" data-id="${t.id}">${esc(t.name)}</b><div class="spacer"></div>${ctxTag(t.ctx)}<span class="tag p${t.prio}">P${t.prio}</span>${t.est ? `<span class="tag src">⟨i:clock⟩ ${t.est}h</span>` : ''}</div>
    <div class="muted tpl-title">⟨i:check-square⟩ ${esc(tplTitle(t))}</div>
    <div class="row" style="gap:6px;margin-top:6px">${pj ? projTag(pj.id) : ''}${g ? `<span class="tag src">⟨i:layers⟩ ${esc(g.title.slice(0, 26))}</span>` : ''}${t.dueRule ? `<span class="tag src">⟨i:calendar⟩ ${DUE_RULES.find(x => x[0] === t.dueRule)?.[1]}</span>` : ''}${t.notes ? '<span class="tag src">opis</span>' : ''}</div>
    <div class="row" style="margin-top:10px"><span class="muted" style="font-size:11px">użyto ${t.uses || 0}×${t.lastUsed ? ' · ostatnio ' + fmt(t.lastUsed.slice(0, 10)) : ''}</span><div class="spacer"></div>
      <button class="btn sm ghost" data-action="tpl-edit" data-id="${t.id}" title="Edytuj">⟨i:pencil⟩</button><button class="btn sm ghost" data-action="tpl-dup" data-id="${t.id}" title="Duplikuj">⟨i:template⟩</button><button class="btn sm primary" data-action="tpl-use" data-id="${t.id}">⟨i:zap⟩ Utwórz zadanie</button></div></div>`; };
  return taskTabs() + `<div class="tpl-layout"><aside class="card tpl-cats">
      <button class="qcat ${cat === '__all' ? 'on' : ''}" data-action="set" data-k="tplCat" data-v="__all">Wszystkie <span>${total}</span></button>
      ${cats.map(c => `<button class="qcat ${cat === c ? 'on' : ''}" data-action="set" data-k="tplCat" data-v="${esc(c)}">${esc(c)} <span>${cnt(c)}</span></button>`).join('')}
      <div class="hint" style="margin-top:10px">Kategorię nadajesz w edycji szablonu – wpisz nową nazwę, aby utworzyć kategorię.</div></aside>
    <div><div class="row" style="margin-bottom:12px"><input class="input" style="max-width:300px" placeholder="Szukaj szablonu…" data-bind="tplQ" value="${esc(S.tplQ || '')}">
      <div class="seg"><button data-action="set" data-k="tplSort" data-v="uses" class="${sort === 'uses' ? 'on' : ''}">Najczęściej używane</button><button data-action="set" data-k="tplSort" data-v="name" class="${sort === 'name' ? 'on' : ''}">A–Z</button></div>
      <div class="spacer"></div><button class="btn primary" data-action="tpl-new">⟨i:plus⟩ Nowy szablon</button></div>
      ${groups.map(([c, gl]) => `<div class="group-h">${esc(c)} · ${gl.length}</div><div class="grid g2">${gl.map(card).join('')}</div>`).join('') || '<div class="card empty">Brak szablonów. Utwórz szablon albo zapisz istniejące zadanie jako szablon (przycisk w oknie zadania).</div>'}
    </div></div>`;
}
function seedTemplates(projects, goals) {
  const pj = q => (projects || []).find(p => p.title.includes(q))?.id || null, gl = q => (goals || []).find(g => g.title.includes(q))?.id || null;
  return [
    ['Raport miesięczny dla PM', 'Raportowanie', 'job', 'Raport miesięczny dla PM – {miesiąc}', 1, 2, 'eom', pj('Alfa'), null, 'Postęp prac, ryzyka, plan na kolejny miesiąc.'],
    ['Status tygodniowy', 'Raportowanie', 'job', 'Status tygodniowy – tydz. {tydzień}', 2, 0.5, 'eow', null, null, ''],
    ['Przygotowanie HLS', 'Dokumentacja IFS', 'job', 'HLS', 2, 8, 'wd5', pj('Kompetencje'), gl('HLS'), 'Szablon HLS: cel, zakres, wymagania, założenia, ryzyka.'],
    ['Przygotowanie DLS', 'Dokumentacja IFS', 'job', 'DLS', 2, 8, 'wd5', pj('Kompetencje'), gl('HLS'), 'Szablon DLS: projekt rozwiązania, obiekty IFS, testy.'],
    ['Review dokumentacji', 'Dokumentacja IFS', 'job', 'Review dokumentacji', 2, 2, 'wd2', null, null, ''],
    ['Follow-up po spotkaniu', 'Spotkania', 'job', 'Follow-up', 2, 0.5, 'd1', null, null, 'Wysłać podsumowanie i listę ustaleń uczestnikom.'],
    ['Przygotowanie warsztatu', 'Spotkania', 'job', 'Przygotowanie warsztatu', 1, 3, 'wd2', pj('Alfa'), null, 'Agenda, materiały, sala/Teams, lista uczestników.'],
    ['Rozliczenie delegacji', 'Administracja', 'job', 'Rozliczenie delegacji – {miesiąc}', 2, 1, 'eom', null, null, ''],
    ['Wniosek urlopowy', 'Administracja', 'job', 'Wniosek urlopowy', 3, 0.25, '', null, null, ''],
    ['Opłaty i rachunki', 'Dom', 'priv', 'Opłacić rachunki – {miesiąc}', 1, 0.5, 'd7', null, null, 'Prąd, gaz, internet, czynsz.'],
    ['Zakupy', 'Dom', 'priv', 'Zakupy', 3, 1, '', null, null, ''],
    ['Trening biegowy', 'Zdrowie i sport', 'priv', 'Trening biegowy', 2, 1, '', pj('Zdrowie'), gl('Półmaraton'), ''],
    ['Wizyta lekarska', 'Zdrowie i sport', 'priv', 'Wizyta lekarska', 1, 1.5, '', pj('Zdrowie'), null, 'Zabrać wyniki badań.']
  ].map(([name, category, ctx, title, prio, est, dueRule, projectId, goalId, notes]) => ({id: uid(), name, category, ctx, title, prio, est, dueRule, projectId, goalId, notes, uses: 0}));
}

function editTask(t, preset = {}) {
  const isNew = !t; t = t || {title: '', ctx: S.ctx === 'all' ? 'job' : S.ctx, goalId: preset.goal || null, projectId: preset.project || (preset.goal && goalById(preset.goal)?.projectId) || null, start: preset.due || TODAY(), due: preset.due || TODAY(), status: 'todo', prio: 2, est: 0, notes: '', ...preset.task};
  if (preset.goal) t.ctx = goalById(preset.goal)?.ctx || t.ctx;
  if (preset.project) t.ctx = projById(preset.project)?.ctx || t.ctx;
  const curProj = projOf(t) || t.projectId;
  modal(`<h3>${isNew ? 'Nowe zadanie' : 'Edycja zadania'}</h3>${!isNew && pvCal() ? `<p class="hint lockinfo" style="margin:0 0 12px">⟨i:eye-off⟩ Dane zadania są ukryte. Kliknij oko, aby je odsłonić. ${eyeBtn('sm')}</p>` : ''}<form data-form="task" ${!isNew && pvCal() ? 'data-pv="1"' : ''} data-id="${isNew ? '' : t.id}" data-tpl="${isNew && t.tplId ? t.tplId : ''}" class="grid g2">
    ${isNew && tplAll().length ? `<label class="f tplpick" style="grid-column:1/-1">⟨i:template⟩ Z szablonu<select class="input" id="tplPick"><option value="">— bez szablonu —</option>${tplCats().map(c => `<optgroup label="${esc(c)}">${tplAll().filter(x => inCtx(x) && (x.category || 'Bez kategorii') === c).map(x => `<option value="${x.id}" ${t.tplId === x.id ? 'selected' : ''}>${esc(x.name)}${x.est ? ' · ' + x.est + 'h' : ''}</option>`).join('')}</optgroup>`).join('')}</select></label>` : ''}
    <label class="f" style="grid-column:1/-1">Nazwa<input class="input" name="title" required value="${esc(t.title)}"></label>
    <label class="f">Kontekst<select class="input" name="ctx"><option value="job" ${t.ctx === 'job' ? 'selected' : ''}>JOB – firmowe</option><option value="priv" ${t.ctx === 'priv' ? 'selected' : ''}>PRIV – prywatne</option></select></label>
    <label class="f">Projekt<select class="input" name="projectId">${projOptions(curProj)}</select></label>
    <label class="f" style="grid-column:1/-1">Podprojekt<select class="input" name="goalId"><option value="">— bez podprojektu —</option>${S.data.goals.map(g => `<option value="${g.id}" data-proj="${g.projectId || ''}" data-ctx="${g.ctx}" ${t.goalId === g.id ? 'selected' : ''}>[${g.ctx.toUpperCase()}] ${esc(g.title)}${projById(g.projectId) ? ' · projekt: ' + esc(projById(g.projectId).title) : ''}</option>`).join('')}</select></label>
    <label class="f">Start<input class="input" type="date" name="start" value="${t.start || ''}"></label>
    <label class="f">Termin<input class="input" type="date" name="due" value="${t.due || ''}"></label>
    <div class="dueload" id="dueLoad" style="grid-column:1/-1" hidden></div>
    <label class="f">Status<select class="input" name="status">${[['todo', 'Do zrobienia'], ['doing', 'W toku'], ['done', 'Zrobione']].map(([k, l]) => `<option value="${k}" ${t.status === k ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
    <div class="row"><label class="f">Priorytet<select class="input" name="prio">${[1, 2, 3].map(n => `<option ${t.prio === n ? 'selected' : ''}>${n}</option>`).join('')}</select></label>
    <label class="f">Estymacja (h)<input class="input" type="number" step="0.5" min="0" name="est" value="${t.est || 0}" style="width:110px"></label>
    <label class="f ic-field" data-only="job">IC case<input class="input" name="ic" value="${esc(t.ic || '')}" placeholder="np. IC-12345" style="width:140px"></label>
    <label class="f" data-only="priv">Koszt szacowany (zł)<input class="input" type="number" min="0" step="10" name="costEst" value="${t.costEst || ''}" style="width:140px"></label>
    <label class="f" data-only="priv">Koszt rzeczywisty (zł)<input class="input" type="number" min="0" step="10" name="costAct" value="${t.costAct || ''}" style="width:140px"></label></div>
    <label class="f">→ Odbiorca (komu zlecam)<select class="input ctx-filter" name="assigneeId">${personOptions(t.assigneeId)}</select></label>
    <label class="f">← Zlecający (od kogo)<select class="input ctx-filter" name="requesterId">${personOptions(t.requesterId)}</select></label>
    <label class="f" style="grid-column:1/-1">Opis<textarea class="input" name="notes" rows="3">${esc(t.notes || '')}</textarea></label>
    ${isNew ? '' : `<div class="trk-line" style="grid-column:1/-1">⟨i:timer⟩ Realizacja: <b>${fmtH(taskSpent(t.id))}</b>${t.est ? ` z ${t.est}h` : ''} · ${entriesAll().filter(e => e.taskId === t.id).length} wpisów<div class="spacer"></div>${timerOn()?.taskId === t.id ? '<span class="tag src"><span class="tc-dot"></span> w toku</span>' : `<button type="button" class="btn sm" data-action="timer-start-task" data-id="${t.id}">⟨i:play⟩ Start</button>`}<button type="button" class="btn sm ghost" data-action="entry-new" data-task="${t.id}">⟨i:plus⟩ Wpis ręczny</button></div>`}
    ${!isNew ? (() => { const bl = schAll().filter(b => b.taskId === t.id && b.d >= TODAY()).sort((a, b) => (a.d + a.s).localeCompare(b.d + b.s)), tot = bl.reduce((a, b) => a + schDur(b), 0); return `<div class="trk-line sch-line" style="grid-column:1/-1">⟨i:schedule⟩ Harmonogram: ${bl.length ? `<b>${fmtH(tot)}</b> w ${plN(bl.length, 'bloku', 'blokach', 'blokach')} · najbliższy ${fmt(bl[0].d)} ${bl[0].s}–${bl[0].e}` : '<span class="muted">brak bloków</span>'}<div class="spacer"></div><button type="button" class="btn sm ghost" data-action="sch-add" data-task="${t.id}">⟨i:plus⟩ Blok</button><button type="button" class="btn sm ghost" data-action="sch-open">Otwórz harmonogram</button></div>`; })() : ''}
    ${!isNew && t.teamsUrl ? `<div class="row" style="grid-column:1/-1"><button type="button" class="btn sm" data-action="teams-open" data-id="${t.id}">⟨i:video⟩ Otwórz wiadomość w Teams</button><span class="muted" style="font-size:12px">zadanie utworzone z czatu Teams</span></div>` : ''}
    ${!isNew && t.mailId ? `<div class="row" style="grid-column:1/-1"><button type="button" class="btn sm" data-action="mail-open" data-id="${t.id}">⟨i:mail⟩ Otwórz mail w Outlooku</button><span class="muted" style="font-size:12px">zadanie utworzone z maila (kategoria TM)</span></div>` : ''}
    ${recurLine(t, 'task', isNew)}
    ${isNew ? '' : cmtSection('task', t.id)}
    ${isNew ? '' : notesSection('task', t.id)}
    <div class="row" style="grid-column:1/-1;justify-content:space-between"><div class="row">${isNew ? '' : `<button type="button" class="btn danger" data-action="del-task" data-id="${t.id}">Usuń</button>`}<button type="button" class="btn ghost" data-action="task-to-tpl" title="Zapisz dane tego zadania jako szablon">⟨i:template⟩ Zapisz jako szablon</button></div>
    <div class="row"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">Zapisz</button></div></div></form>`, m => { $('[name=title]', m).focus(); linkForm(m, true);
    ctxFields(m);
    { const f0 = $('form[data-form=task]', m), box = $('#dueLoad', m), DN = ['Nd', 'Pn', 'Wt', 'Śr', 'Cz', 'Pt', 'Sb'];
      const upd = () => { const due = f0.due.value; if (!due || f0.status.value === 'done' || due < TODAY()) { box.hidden = true; return; }
        const tmp = {id: t.id || '__new', ctx: f0.ctx.value, status: 'todo', est: +f0.est.value || 0, start: f0.start.value || due, due}, own = taskLoad(tmp)[due] || 0, x = dayLoadInfo(due, tmp.ctx, t.id), tot = x.tot + own, over = tot > x.cap + .01;
        box.hidden = false; box.className = 'dueload ' + (over ? 'bad' : tot > x.cap * .85 ? 'warn' : 'ok');
        box.innerHTML = H(`⟨i:gauge⟩ ${DN[P(due).getDay()]} ${fmt(due)} (${tmp.ctx.toUpperCase()}): <b>${r1(tot)} h / ${r1(x.cap)} h</b> – spotkania ${r1(x.mtg)}h, inne zadania ${r1(x.tasks)}h, to zadanie ${r1(own)}h${x.type !== 'work' && tmp.ctx === 'job' ? ` · <b>${x.type === 'hol' ? 'święto' : x.type === 'vac' ? 'urlop' : 'weekend'}</b>` : ''}${x.trip && f0.dataset.trip !== x.trip.id ? ` · <b>⟨i:plane⟩ delegacja: ${esc(x.trip.place)}</b>${tmp.ctx === 'priv' ? ' – jesteś poza domem' : ''}` : ''}${over ? ' · <b>dzień przeładowany</b>' : ''}`);
        if (x.trip && tmp.ctx === 'priv' && f0.dataset.trip !== x.trip.id) box.className = 'dueload bad'; };
      ['change', 'input'].forEach(ev => f0.addEventListener(ev, e => { if (['due', 'start', 'est', 'ctx', 'status'].includes(e.target.name)) upd(); })); upd(); }
    const tp = $('#tplPick', m); if (tp) tp.addEventListener('change', () => { const x = tplById(tp.value); if (!x) return; const f = $('form[data-form=task]', m);
      const due = x.dueRule ? iso(dueFromRule(x.dueRule)) : f.due.value || TODAY();
      f.title.value = tplTitle(x, due); f.ctx.value = x.ctx; f.projectId.value = projById(x.projectId) ? x.projectId : ''; f.goalId.value = goalById(x.goalId) ? x.goalId : '';
      if (goalById(x.goalId)?.projectId) f.projectId.value = goalById(x.goalId).projectId;
      f.prio.value = x.prio; f.est.value = x.est || 0; if (x.notes) f.notes.value = x.notes; f.due.value = due; f.start.value = startFor(due, +x.est || 0); f.dataset.tpl = x.id;
      toast(`Wypełniono z szablonu „${x.name}” – uzupełnij termin i czas`); f.due.focus(); });
  });
}
/* kaskada w formularzach: kontekst (PRIV/JOB) → projekty → podprojekty → zadania
   wybór kontekstu zawęża projekty, wybór projektu zawęża podprojekty, wybór podprojektu (lub projektu) zawęża zadania;
   wybór elementu niższego poziomu ustawia wyższe (podprojekt → projekt → kontekst) */
function linkForm(m) {
  const ps = $('[name=projectId]', m), cs = $('[name=ctx]', m), gs = $('[name=goalId]', m), ts = $('[name=taskId]', m);
  const pOk = o => { const p = projById(o.value); return !cs || !p || p.ctx === cs.value; };
  const gOk = o => (!cs || o.dataset.ctx === cs.value) && (!ps || !ps.value || o.dataset.proj === ps.value);
  const tOk = o => (!cs || !o.dataset.ctx || o.dataset.ctx === cs.value) && (!ps || !ps.value || o.dataset.proj === ps.value) && (!gs || !gs.value || o.dataset.goal === gs.value);
  const filt = (sel, ok, first) => {
    if (!sel) return;
    [...sel.options].forEach(o => { if (!o.value) return; const v = ok(o) || (first && o.selected); o.hidden = !v; o.disabled = !v; });
    if (sel.value && sel.selectedOptions[0]?.disabled) sel.value = '';
    [...sel.querySelectorAll('optgroup')].forEach(g => g.hidden = ![...g.children].some(o => !o.hidden));
  };
  const apply = first => { filt(ps, pOk, first); filt(gs, gOk, first); filt(ts, tOk, first); $$('select.ctx-filter', m).forEach(sel => filt(sel, o => !cs || o.dataset.ctx === cs.value, first)); };
  cs && cs.addEventListener('change', () => apply());
  ps && ps.addEventListener('change', () => { const p = projById(ps.value); if (p && cs) cs.value = p.ctx; apply(); });
  gs && gs.addEventListener('change', () => { const o = gs.selectedOptions[0]; if (o && o.value) { if (ps && o.dataset.proj) ps.value = o.dataset.proj; if (cs) cs.value = o.dataset.ctx; } apply(); });
  ts && ts.addEventListener('change', () => { const o = ts.selectedOptions[0]; if (o && o.value) { if (ps && o.dataset.proj) ps.value = o.dataset.proj; if (cs && o.dataset.ctx) cs.value = o.dataset.ctx; } apply(); });
  apply(true);
}
function editProject(p) {
  const isNew = !p; p = p || {title: '', ctx: S.ctx === 'all' ? 'job' : S.ctx, status: 'active', start: TODAY(), due: iso(add(today(), 90)), desc: '', client: ''};
  modal(`<h3>${isNew ? 'Nowy projekt' : 'Edycja projektu'}</h3><form data-form="project" data-id="${isNew ? '' : p.id}" class="grid g2">
    <label class="f" style="grid-column:1/-1">Nazwa projektu<input class="input" name="title" required value="${esc(p.title)}" placeholder="np. Beta – migracja do IFS Cloud"></label>
    <label class="f">Kategoria<select class="input" name="ctx"><option value="job" ${p.ctx === 'job' ? 'selected' : ''}>JOB – firmowy</option><option value="priv" ${p.ctx === 'priv' ? 'selected' : ''}>PRIV – prywatny</option></select></label>
    <label class="f">Status<select class="input" name="status">${Object.entries(PSTAT).map(([k, [l]]) => `<option value="${k}" ${p.status === k ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
    <label class="f" style="grid-column:1/-1">Kolor projektu (spotkania w kalendarzu)<span class="row">${PCOL.map(c => `<label class="swatch" style="background:${c}"><input type="radio" name="color" value="${c}" ${projColor(p.id ? p : null) === c || (!p.id && c === PCOL[(S.data.projects || []).length % PCOL.length]) ? 'checked' : ''}></label>`).join('')}</span></label>
    <label class="f">Start<input class="input" type="date" name="start" value="${p.start || ''}"></label><label class="f">Termin zakończenia<input class="input" type="date" name="due" value="${p.due || ''}"></label>
    <label class="f">Klient / interesariusz<input class="input" name="client" value="${esc(p.client || '')}" placeholder="opcjonalnie"></label>
    <label class="f" data-only="job">Budżet godzin projektu (h)<input class="input" type="number" min="0" step="0.5" name="budgetH" value="${p.budgetH || ''}" placeholder="puste = suma z podprojektów"></label>
    <p class="hint" data-only="priv" style="margin:22px 0 0">Budżet PRIV ustawiasz w podprojektach (zł), koszty – w zadaniach.</p>
    <label class="f" style="grid-column:1/-1">Opis / zakres<textarea class="input" name="desc" rows="3">${esc(p.desc || '')}</textarea></label>
    ${isNew ? '' : `<p class="hint" style="grid-column:1/-1;margin:0">Zmiana kategorii PRIV/JOB przeniesie też ${projGoals(p).length} podprojekty i ${projTasks(p).length} zadania tego projektu.</p>`}
    <div class="row" style="grid-column:1/-1;justify-content:space-between">${isNew ? '<span></span>' : `<button type="button" class="btn danger" data-action="del-project" data-id="${p.id}">Usuń projekt</button>`}
    <div class="row"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">Zapisz</button></div></div></form>`, m => { $('[name=title]', m).focus(); ctxFields(m); });
}
function editGoal(g, preset = {}) {
  const isNew = !g; g = g || {title: '', ctx: projById(preset.project)?.ctx || (S.ctx === 'all' ? 'job' : S.ctx), projectId: preset.project || null, start: TODAY(), due: iso(add(today(), 90)), desc: '', progress: 0};
  const hasT = !isNew && S.data.tasks.some(t => t.goalId === g.id);
  modal(`<h3>${isNew ? 'Nowy podprojekt' : 'Edycja podprojektu'}</h3><form data-form="goal" data-id="${isNew ? '' : g.id}" class="grid g2">
    <label class="f" style="grid-column:1/-1">Nazwa podprojektu<input class="input" name="title" required value="${esc(g.title)}"></label>
    <label class="f">Kontekst<select class="input" name="ctx"><option value="job" ${g.ctx === 'job' ? 'selected' : ''}>JOB – firmowy</option><option value="priv" ${g.ctx === 'priv' ? 'selected' : ''}>PRIV – prywatny</option></select></label>
    <label class="f">Projekt<select class="input" name="projectId">${projOptions(g.projectId)}</select></label>
    <label class="f">Postęp ręczny (%)${hasT ? ' – liczony z zadań' : ''}<input class="input" type="number" min="0" max="100" name="progress" value="${hasT ? goalProgress(g) : g.progress || 0}" ${hasT ? 'disabled' : ''}></label>
    <label class="f">Start<input class="input" type="date" name="start" value="${g.start}"></label><label class="f">Termin<input class="input" type="date" name="due" value="${g.due}"></label>
    <label class="f" data-only="priv">Budżet podprojektu (zł)<input class="input" type="number" min="0" step="100" name="budget" value="${g.budget || ''}" placeholder="np. 10000"></label>
    <label class="f" data-only="job">Budżet godzin podprojektu (h)<input class="input" type="number" min="0" step="0.5" name="budgetH" value="${g.budgetH || ''}" placeholder="opcjonalnie"></label>
    ${isNew ? '' : `<div style="grid-column:1/-1" data-only="${g.ctx}">${budgetCell(g.ctx === 'priv' ? goalMoney(g) : goalHours(g), g.ctx === 'priv')}</div>`}
    <label class="f" style="grid-column:1/-1">Opis / miernik sukcesu<textarea class="input" name="desc" rows="3">${esc(g.desc || '')}</textarea></label>
    <div class="row" style="grid-column:1/-1;justify-content:space-between">${isNew ? '<span></span>' : `<button type="button" class="btn danger" data-action="del-goal" data-id="${g.id}">Usuń podprojekt</button>`}
    <div class="row"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">Zapisz</button></div></div></form>`, m => { $('[name=title]', m).focus(); linkForm(m, false); ctxFields(m); });
}
function editMeeting(m, preset = {}) {
  const isNew = !m;
  if (!isNew) S.mtgs.lastOpened = m.id;
  if (isNew) { const st = `${preset.date || TODAY()}T${preset.time || '10:00'}`; const pj = projById(preset.project);
    m = {title: '', start: st, end: addMin(st, 60), ctx: pj ? pj.ctx : (S.ctx === 'all' ? 'job' : S.ctx), projectId: pj ? pj.id : null, source: 'local', attendees: [], location: '', body: ''}; }
  const src = [['local', 'Tylko w TM (bez platformy)'], ['outlook', 'Outlook – spotkanie'], ['teams', 'Teams – spotkanie online (Outlook)'], ['google', 'Google Calendar']];
  const conn = s => s === 'google' ? S.conn.google : s === 'local' ? false : S.conn.ms;
  const lock = !isNew && isOutlook(m), ro = lock ? 'readonly' : '';
  modal(`<h3>${isNew ? 'Nowe spotkanie' : 'Spotkanie'} ${!isNew ? srcBadge(m.source) : ''}</h3>${!isNew && S.privacy ? `<p class="hint lockinfo" style="margin:0 0 12px">⟨i:eye-off⟩ Dane spotkania są ukryte. Kliknij oko, aby je odsłonić. ${eyeBtn('sm')}</p>` : ''}<form data-form="mtg" ${!isNew ? 'data-pv="1"' : ''}" data-id="${isNew ? '' : m.id}" class="grid g2">
    <label class="f" style="grid-column:1/-1">Temat<input class="input" name="title" required value="${esc(m.title)}" placeholder="np. Przegląd postępu projektu" ${ro}></label>
    <label class="f">Data<input class="input" type="date" name="date" required value="${m.start.slice(0, 10)}" ${ro}></label>
    <div class="row"><label class="f">Od<input class="input" type="time" name="from" required value="${hm(m.start)}" ${ro}></label><label class="f">Do<input class="input" type="time" name="to" required value="${hm(m.end)}" ${ro}></label></div>
    <label class="f">Kontekst<select class="input" name="ctx"><option value="job" ${m.ctx === 'job' ? 'selected' : ''}>JOB – firmowe</option><option value="priv" ${m.ctx === 'priv' ? 'selected' : ''}>PRIV – prywatne</option></select></label>
    <label class="f">Projekt (kolor spotkania)<select class="input" name="projectId">${projOptions(m.projectId)}</select></label>
    <label class="f" style="grid-column:1/-1">Platforma<select class="input" name="source" ${!isNew && (m.ext || lock) ? 'disabled' : ''}>${src.map(([k, l]) => `<option value="${k}" ${m.source === k ? 'selected' : ''}>${l}${k !== 'local' && conn(k) ? ' – połączono' : ''}</option>`).join('')}</select></label>
    <label class="f" style="grid-column:1/-1">Uczestnicy (e-maile, przecinek)<input class="input" name="att" list="peopleListM" value="${esc((m.attendees || []).join(', '))}" ${ro}></label>
    <datalist id="peopleListM">${PEOPLE.map(x => `<option value="${x.email}">${esc(x.name)}</option>`).join('')}</datalist>
    <label class="f">Lokalizacja<input class="input" name="loc" value="${esc(m.location || '')}" ${ro}></label>
    <label class="f" style="align-self:end">${isNew ? `<span class="row" style="gap:6px;font-weight:500"><input type="checkbox" name="invite" checked> Otwórz kreator platformy (zaproszenia)</span>` : (m.link ? joinLink(m, 'Dołącz do spotkania') : '')}</label>
    <label class="f" style="grid-column:1/-1">Agenda<textarea class="input" name="body" rows="3">${esc(m.body || '')}</textarea></label>
    ${lock ? '' : recurLine(m, 'mtg', isNew)}
    ${isNew ? '' : cmtSection('mtg', m.id)}
    ${isNew ? '' : notesSection('mtg', m.id)}
    ${lock ? `<p class="hint lockinfo" style="grid-column:1/-1;margin:0">⟨i:lock⟩ Spotkanie z <b>Outlooka</b>${m.imported ? ' (zaimportowane)' : ''}: temat, termin, uczestników i usuwanie zmieniasz w Outlooku. W TM możesz przypisać projekt (kolor), kontekst i notatki. <a href="#" data-action="ol-open" data-id="${m.id}">Otwórz w aplikacji Outlook ⟨i:external⟩</a></p>` : ''}
    ${!isNew && !lock && (m.demo || m.ext) ? `<p class="hint" style="grid-column:1/-1;margin:0">${m.demo ? 'Spotkanie przykładowe (tryb demo).' : `Spotkanie pobrane z ${SRC[m.source]?.[0]}. Usunięcie ${conn(m.source) ? 'odwoła je też w ' + SRC[m.source]?.[0] : 'działa tylko w TM'}.`}</p>` : ''}
    <div class="row" style="grid-column:1/-1;justify-content:space-between">${isNew ? '<span></span>' : `<div class="row">${lock ? '' : `<button type="button" class="btn danger" data-action="del-mtg" data-id="${m.id}">⟨i:trash⟩ Usuń spotkanie</button>`}<button type="button" class="btn" data-action="mtg-task" data-id="${m.id}">⟨i:plus⟩ To-do</button>${mtgTrkBtn(m, 'btn', true)}${S.view === 'meetings' ? `<button type="button" class="btn" data-action="show-in-cal" data-id="${m.id}">⟨i:calendar⟩ Pokaż w kalendarzu</button>` : `<button type="button" class="btn" data-action="show-in-list" data-id="${m.id}">⟨i:list⟩ Pokaż na liście</button>`}</div>`}
    <div class="row"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">Zapisz</button></div></div></form>`, mm => {
    $('[name=title]', mm).focus(); linkForm(mm, false);
    const f = $('[name=from]', mm), t = $('[name=to]', mm); let dur = toMin('T' + m.end.slice(11)) - toMin('T' + m.start.slice(11));
    const hmToMin = v => +v.slice(0, 2) * 60 + +v.slice(3, 5), minToHm = n => `${pad(Math.floor(clamp(n, 0, 1439) / 60))}:${pad(clamp(n, 0, 1439) % 60)}`;
    f.addEventListener('change', () => { t.value = minToHm(hmToMin(f.value) + Math.max(15, dur)); });
    t.addEventListener('change', () => { dur = hmToMin(t.value) - hmToMin(f.value); });
    const sSel = $('[name=source]', mm), cSel = $('[name=ctx]', mm);
    sSel.addEventListener('change', () => { if (sSel.value === 'google' && !$('[name=projectId]', mm).value) cSel.value = 'priv'; if ((sSel.value === 'outlook' || sSel.value === 'teams') && !$('[name=projectId]', mm).value) cSel.value = 'job'; });
  });
}
function meetingUrl(m, body) {
  const q = o => Object.entries(o).filter(([, v]) => v).map(([k, v]) => k + '=' + encodeURIComponent(v)).join('&');
  const st = new Date(m.start), en = new Date(m.end), gfmt = d => d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, ''), att = (m.attendees || []).join(',');
  if (m.source === 'teams') return 'msteams:/l/meeting/new?' + q({subject: m.title, attendees: att, startTime: st.toISOString(), endTime: en.toISOString(), content: body});
  if (m.source === 'outlook') return 'https://outlook.office.com/calendar/deeplink/compose?' + q({subject: m.title, to: att, startdt: st.toISOString(), enddt: en.toISOString(), body, location: m.location});
  if (m.source === 'google') return 'https://calendar.google.com/calendar/render?' + q({action: 'TEMPLATE', text: m.title, dates: gfmt(st) + '/' + gfmt(en), add: att, details: body, location: m.location});
  return null;
}
async function deleteMeeting(id) {
  const i = S.data.meetings.findIndex(x => x.id === id); if (i < 0) return; const m = S.data.meetings[i];
  if (!canDelMtg(m)) return toast('Spotkania z Outlooka usuwasz w Outlooku');
  if (m.ext) {
    const api = m.source === 'google' ? S.conn.google : S.conn.ms;
    if (api) { try { if (m.source === 'google') await GG.api(`https://www.googleapis.com/calendar/v3/calendars/primary/events/${m.id.slice(2)}`, {method: 'DELETE'}).catch(e => { if (!/Unexpected|JSON/.test(e.message)) throw e; }); else await MS.g(`/me/events/${m.id.slice(3)}`, {method: 'DELETE'}); }
      catch (e) { return toast('Nie udało się usunąć w ' + SRC[m.source][0] + ': ' + e.message); } }
  }
  S.data.meetings.splice(i, 1);
  if (m.demo) { S.data.hiddenDemo = [...new Set([...(S.data.hiddenDemo || []), mtgKey(m)])]; }
  save(); closeModal(); render();
  if (m.ext && (m.source === 'google' ? S.conn.google : S.conn.ms)) return toast(`Odwołano spotkanie w ${SRC[m.source][0]}`);
  toastAction(S.privacy ? 'Usunięto spotkanie' : `Usunięto spotkanie „${m.title}”`, '⟨i:undo⟩ Cofnij', () => {
    S.data.meetings.splice(Math.min(i, S.data.meetings.length), 0, m);
    if (m.demo) S.data.hiddenDemo = (S.data.hiddenDemo || []).filter(k => k !== mtgKey(m));
    save(); render(); toast('Przywrócono spotkanie');
  });
}
/* masowe usuwanie spotkań (pomija Outlook/Teams) z możliwością cofnięcia */
async function deleteMeetings(ids) {
  const set = new Set(ids), all = S.data.meetings.map((m, i) => [i, m]).filter(([, m]) => set.has(m.id));
  const del = all.filter(([, m]) => canDelMtg(m)), skipped = all.length - del.length;
  for (const [, m] of del) if (m.ext && m.source === 'google' && S.conn.google) {
    try { await GG.api(`https://www.googleapis.com/calendar/v3/calendars/primary/events/${m.id.slice(2)}`, {method: 'DELETE'}).catch(e => { if (!/Unexpected|JSON/.test(e.message)) throw e; }); } catch (e) { toast('Google: ' + e.message); }
  }
  const delIds = new Set(del.map(([, m]) => m.id));
  S.data.meetings = S.data.meetings.filter(m => !delIds.has(m.id));
  const demoKeys = del.filter(([, m]) => m.demo).map(([, m]) => mtgKey(m));
  if (demoKeys.length) S.data.hiddenDemo = [...new Set([...(S.data.hiddenDemo || []), ...demoKeys])];
  delIds.forEach(id => S.mtgs.sel.delete(id));
  save(); render();
  if (!del.length) return toast(`Nic nie usunięto – ${skipped} spotk. z Outlooka usuwasz w Outlooku`);
  toastAction(`Usunięto ${del.length} spotk.${skipped ? ` · pominięto ${skipped} z Outlooka` : ''}`, '⟨i:undo⟩ Cofnij', () => {
    del.forEach(([i, m]) => S.data.meetings.splice(Math.min(i, S.data.meetings.length), 0, m));
    S.data.hiddenDemo = (S.data.hiddenDemo || []).filter(k => !demoKeys.includes(k)); save(); render(); toast(`Przywrócono ${del.length} spotk.`);
  });
}

/* ================= WIDOK: Kalendarz › Lista spotkań ================= */
function calTabs() {
  return `<div class="row" style="margin-bottom:12px;gap:10px">${eyeBtn()}<div class="spacer"></div>${S.view === 'calendar' ? '<button class="btn" data-action="nav" data-view="meetings">⟨i:list⟩ Lista spotkań</button>' : '<button class="btn" data-action="nav" data-view="calendar">⟨i:calendar⟩ Kalendarz</button>'}${importBtn()}</div>`;
}
function calRange() {
  const C = S.cal, d = C.date;
  if (C.mode === 'day') return [iso(d), iso(d)];
  if (C.mode === 'week') return [iso(sow(d)), iso(add(sow(d), 6))];
  if (C.mode === 'month') return [iso(new Date(d.getFullYear(), d.getMonth(), 1)), iso(new Date(d.getFullYear(), d.getMonth() + 1, 0))];
  return [`${d.getFullYear()}-01-01`, `${d.getFullYear()}-12-31`];
}
const calRangeMeetings = () => { const [a, b] = calRange(); return meetingsC().filter(m => m.start.slice(0, 10) >= a && m.start.slice(0, 10) <= b); };
function mtgPeriod() {
  const F = S.mtgs, t = TODAY();
  switch (F.period) {
    case 'cal': return calRange();
    case 'today': return [t, t];
    case 'week': return [iso(sow(today())), iso(add(sow(today()), 6))];
    case 'month': { const d = today(); return [iso(new Date(d.getFullYear(), d.getMonth(), 1)), iso(new Date(d.getFullYear(), d.getMonth() + 1, 0))]; }
    case 'next30': return [t, iso(add(today(), 30))];
    case 'past': return ['0000-01-01', iso(add(today(), -1))];
    case 'range': return [F.from || '0000-01-01', F.to || '9999-12-31'];
    default: return ['0000-01-01', '9999-12-31'];
  }
}
function filteredMeetings() {
  const F = S.mtgs, [a, b] = mtgPeriod(), q = norm(F.q);
  let ms = meetingsC().filter(m => m.start.slice(0, 10) >= a && m.start.slice(0, 10) <= b);
  if (F.proj) ms = ms.filter(m => F.proj === '-' ? !projById(m.projectId) : m.projectId === F.proj);
  if (F.src) ms = ms.filter(m => F.src === 'outlook' ? isOutlook(m) : m.source === F.src);
  if (q) ms = ms.filter(m => norm([m.title, m.location, (m.attendees || []).join(' '), projById(m.projectId)?.title, m.body].join(' ')).includes(q));
  return ms.sort((x, y) => F.sort === 'asc' ? x.start.localeCompare(y.start) : y.start.localeCompare(x.start));
}
function vMeetings() {
  const F = S.mtgs, ms = filteredMeetings(), [a, b] = mtgPeriod();
  const sel = [...F.sel].filter(id => ms.some(m => m.id === id)), selM = ms.filter(m => F.sel.has(m.id)), selDel = selM.filter(canDelMtg).length;
  const hours = r1(ms.reduce((s, m) => s + (toMin(m.end) - toMin(m.start)) / 60, 0));
  const C = S.cal, calLbl = {day: 'dzień', week: 'tydzień', month: 'miesiąc', year: 'rok'}[C.mode];
  const per = [['cal', `Jak w kalendarzu (${calLbl})`], ['today', 'Dziś'], ['week', 'Ten tydzień'], ['month', 'Ten miesiąc'], ['next30', 'Najbliższe 30 dni'], ['past', 'Minione'], ['all', 'Wszystkie'], ['range', 'Zakres dat…']];
  const allChecked = ms.length && ms.every(m => F.sel.has(m.id));
  let rows = '', lastDay = '';
  ms.forEach(m => {
    const d = m.start.slice(0, 10), p = projById(m.projectId), c = mtgColor(m), dur = toMin(m.end) - toMin(m.start);
    if (d !== lastDay) { lastDay = d; const dayMs = ms.filter(x => x.start.slice(0, 10) === d);
      rows += `<tr class="dayrow ${d === TODAY() ? 'today' : ''}"><td colspan="9"><span data-action="show-day" data-d="${d}" class="crumb" title="Pokaż ten dzień w kalendarzu">⟨i:calendar⟩ ${DOWL[P(d).getDay()]}, ${fmt(d)}</span> <span class="muted">· ${dayMs.length} spotk. · ${r1(dayMs.reduce((s, x) => s + (toMin(x.end) - toMin(x.start)) / 60, 0))}h</span></td></tr>`; }
    rows += `<tr class="mrow ${F.sel.has(m.id) ? 'selected' : ''} ${S.mtgs.lastOpened === m.id ? 'flash' : ''}" data-mid="${m.id}">
      <td><input type="checkbox" class="selchk" data-action="msel" data-id="${m.id}" ${F.sel.has(m.id) ? 'checked' : ''}></td>
      <td class="nowrap"><b>${hm(m.start)}</b>–${hm(m.end)}<div class="muted" style="font-size:11px">${dur} min</div></td>
      <td><span class="mcol" style="background:${c}"></span></td>
      <td class="mt" data-action="edit-mtg" data-id="${m.id}">${ctxIcon(m.ctx)} ${pv(esc(m.title))}${m.location ? `<div class="muted" style="font-size:11px">⟨i:pin⟩ ${pv(esc(m.location))}</div>` : ''}</td>
      <td>${p ? `<span class="row" style="gap:6px;flex-wrap:nowrap"><span class="pcol" style="background:${projColor(p)}"></span>${pv(esc(p.title))}</span>` : '<span class="muted">—</span>'}</td>
      <td class="nowrap">${srcBadge(m.source)} ${SRC[m.source]?.[0] || m.source}</td>
      <td class="num">${(m.attendees || []).length || '<span class="muted">—</span>'}</td>
      <td class="nowrap acts">${cmtBadge('mtg', m.id)}${noteBadge('mtg', m.id)}<button class="btn sm ghost" data-action="show-in-cal" data-id="${m.id}" title="Pokaż w kalendarzu">⟨i:calendar⟩</button><button class="btn sm ghost" data-action="edit-mtg" data-id="${m.id}" title="Edytuj">⟨i:pencil⟩</button>${delMtgBtn(m, 'del btn sm ghost')}</td></tr>`;
  });
  const bySrc = ['local', 'google', 'outlook', 'teams'].map(k => [k, ms.filter(m => m.source === k).length]).filter(x => x[1]);
  return calTabs() + `<div class="row" style="margin-bottom:10px">
    <input class="input" style="max-width:260px" placeholder="Szukaj: temat, osoba, miejsce, projekt…" data-bind="mtgs.q" value="${esc(F.q)}">
    <select class="input" style="max-width:220px" data-bind="mtgs.period">${per.map(([k, l]) => `<option value="${k}" ${F.period === k ? 'selected' : ''}>${l}</option>`).join('')}</select>
    ${F.period === 'range' ? `<input class="input" type="date" style="max-width:160px" data-bind="mtgs.from" value="${F.from}"><input class="input" type="date" style="max-width:160px" data-bind="mtgs.to" value="${F.to}">` : ''}
    <select class="input" style="max-width:220px" data-bind="mtgs.proj"><option value="">Wszystkie projekty</option><option value="-" ${F.proj === '-' ? 'selected' : ''}>Bez projektu</option>${projectsC().map(p => `<option value="${p.id}" ${F.proj === p.id ? 'selected' : ''}>${esc(p.title)}</option>`).join('')}</select>
    <select class="input" style="max-width:200px" data-bind="mtgs.src"><option value="">Wszystkie platformy</option><option value="local" ${F.src === 'local' ? 'selected' : ''}>Tylko TM</option><option value="google" ${F.src === 'google' ? 'selected' : ''}>Google Calendar</option><option value="outlook" ${F.src === 'outlook' ? 'selected' : ''}>Outlook + Teams</option></select>
    <div class="spacer"></div><button class="btn primary" data-action="new-mtg" data-d="${a > TODAY() && a < '9999' ? a : TODAY()}">⟨i:plus⟩ Spotkanie</button></div>
    <div class="row muted" style="margin-bottom:10px;font-size:12px">${F.period === 'cal' ? `Zakres zsynchronizowany z widokiem kalendarza: <b>${fmt(a)}${a !== b ? ' – ' + fmt(b) : ''}</b> <button class="btn sm ghost" data-action="cal-nav-list" data-n="-1">⟨i:chev-left⟩</button><button class="btn sm ghost" data-action="cal-nav-list" data-n="1">⟨i:chev-right⟩</button> ·` : ''}
      <b>${ms.length}</b> spotkań · ${hours}h ${bySrc.map(([k, n]) => `· ${srcBadge(k)}${n}`).join(' ')}</div>
    <div class="row bulkbar ${sel.length ? 'on' : ''}" style="margin-bottom:10px">
      <label class="row" style="gap:6px;cursor:pointer"><input type="checkbox" class="selchk" data-action="msel-all" ${allChecked ? 'checked' : ''}> Zaznacz wszystkie (${ms.length})</label>
      <span class="muted"><b>${sel.length}</b> zaznaczonych</span>
      ${sel.length ? `<button class="btn sm" data-action="msel-none">Odznacz</button>
        <select class="input" style="max-width:220px;padding:5px 8px" id="bulkProj"><option value="">Przypisz projekt…</option><option value="-">— bez projektu —</option>${(S.data.projects || []).map(p => `<option value="${p.id}">${esc(p.title)}</option>`).join('')}</select>
        <button class="btn sm danger" data-action="msel-delete" ${selDel ? '' : 'disabled'}>⟨i:trash⟩ Usuń zaznaczone (${selDel})</button>
        ${selM.length - selDel ? `<span class="muted" title="Spotkania z Outlooka usuwasz w Outlooku">⟨i:lock⟩ ${selM.length - selDel} z Outlooka – bez usuwania</span>` : ''}` : ''}
    </div>
    <div class="card" style="padding:0;overflow:auto"><table class="tbl mtbl"><thead><tr><th style="width:34px"></th><th class="sortable" data-action="msort">Godzina ${F.sort === 'asc' ? '⟨i:chev-up⟩' : '⟨i:chev-down⟩'}</th><th style="width:14px"></th><th>Temat</th><th>Projekt</th><th>Platforma</th><th class="num">Osoby</th><th></th></tr></thead>
    <tbody>${rows || `<tr><td colspan="9" class="empty">Brak spotkań dla wybranych filtrów</td></tr>`}</tbody></table></div>
    <div class="legend"><span>⟨i:building⟩ JOB</span><span>⟨i:user⟩ PRIV</span><span>${srcBadge('google')} Google Calendar</span><span>${srcBadge('outlook')} Outlook</span><span>${srcBadge('teams')} Teams</span><span>${srcBadge('local')} tylko TM</span><span>⟨i:lock⟩ usuwanie tylko w Outlooku</span></div>`;
}
afterRender.meetings = () => { const r = $('tr.flash'); if (r) { r.scrollIntoView({block: 'center'}); setTimeout(() => { S.mtgs.lastOpened = null; r.classList.remove('flash'); }, 1600); } };
afterRender.calendar = () => { const id = S.cal.flash; if (!id) return; S.cal.flash = null; const el = $(`[data-action="edit-mtg"][data-id="${id}"]`); if (el) { el.scrollIntoView({block: 'center'}); el.classList.add('flash'); setTimeout(() => el.classList.remove('flash'), 2200); } };
/* ================= IMPORT KALENDARZY ZEWNĘTRZNYCH (Outlook teraz, Google docelowo) ================= */
/* Parser iCalendar (.ics) – obsługa: serie RRULE (DAILY/WEEKLY/MONTHLY/YEARLY, INTERVAL, COUNT, UNTIL, BYDAY, BYMONTHDAY),
   EXDATE, wyjątki RECURRENCE-ID, STATUS:CANCELLED, strefy TZID (IANA i nazwy Windows), UTC (Z), linki Teams. */
const WINTZ = {'Central European Standard Time': 'Europe/Warsaw', 'Central Europe Standard Time': 'Europe/Budapest', 'W. Europe Standard Time': 'Europe/Berlin',
  'Romance Standard Time': 'Europe/Paris', 'GMT Standard Time': 'Europe/London', 'Greenwich Standard Time': 'Atlantic/Reykjavik', 'UTC': 'UTC', 'Coordinated Universal Time': 'UTC',
  'E. Europe Standard Time': 'Europe/Chisinau', 'FLE Standard Time': 'Europe/Kiev', 'Russian Standard Time': 'Europe/Moscow', 'Eastern Standard Time': 'America/New_York',
  'Central Standard Time': 'America/Chicago', 'Mountain Standard Time': 'America/Denver', 'Pacific Standard Time': 'America/Los_Angeles', 'India Standard Time': 'Asia/Kolkata',
  'China Standard Time': 'Asia/Shanghai', 'Tokyo Standard Time': 'Asia/Tokyo', 'AUS Eastern Standard Time': 'Australia/Sydney'};
function zonedDate(y, mo, d, h, mi, s, tz) {
  const guess = Date.UTC(y, mo, d, h, mi, s);
  try {
    const f = new Intl.DateTimeFormat('en-US', {timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit'});
    const p = Object.fromEntries(f.formatToParts(new Date(guess)).map(x => [x.type, x.value]));
    const asUTC = Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour % 24, +p.minute, +p.second);
    return new Date(guess - (asUTC - guess));
  } catch { return new Date(y, mo, d, h, mi, s); }
}
function icsDate(val, pm = {}) {
  if (!val) return null;
  if (pm.VALUE === 'DATE' || /^\d{8}$/.test(val)) return {d: new Date(+val.slice(0, 4), +val.slice(4, 6) - 1, +val.slice(6, 8)), allDay: true};
  const m = val.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})?(Z)?$/); if (!m) return null;
  const a = [+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)];
  if (m[7]) return {d: new Date(Date.UTC(...a))};
  const tzid = (pm.TZID || '').replace(/^"|"$/g, '');
  const tz = WINTZ[tzid] || (/^[A-Za-z]+\/[A-Za-z_\-\/]+$/.test(tzid) ? tzid : /^(UTC|GMT|Etc\/UTC)$/i.test(tzid) ? 'UTC' : null);
  return {d: tz ? zonedDate(...a, tz) : new Date(...a)};
}
const icsText = v => (v || '').replace(/\\n/gi, '\n').replace(/\\,/g, ',').replace(/\;/g, ';').replace(/\\\\/g, '\\').trim();
function icsParse(text) {
  const lines = String(text).replace(/\r\n[ \t]/g, '').replace(/\n[ \t]/g, '').split(/\r?\n/);
  const evs = []; let cur = null, depth = 0;
  for (const ln of lines) {
    if (ln === 'BEGIN:VEVENT') { cur = {p: {}, m: {}}; depth = 0; continue; }
    if (ln === 'END:VEVENT') { if (cur) evs.push(cur); cur = null; continue; }
    if (!cur) continue;
    if (ln.startsWith('BEGIN:')) { depth++; continue; } if (ln.startsWith('END:')) { depth--; continue; } if (depth > 0) continue;
    let i = -1, q = false; for (let k = 0; k < ln.length; k++) { const c = ln[k]; if (c === '"') q = !q; else if (c === ':' && !q) { i = k; break; } }
    if (i < 0) continue;
    const [name, ...params] = ln.slice(0, i).split(';'), pm = {};
    params.forEach(x => { const j = x.indexOf('='); if (j > 0) pm[x.slice(0, j).toUpperCase()] = x.slice(j + 1).replace(/^"|"$/g, ''); });
    const N = name.toUpperCase(), v = {val: ln.slice(i + 1), pm};
    if (N === 'ATTENDEE' || N === 'EXDATE') (cur.m[N] ||= []).push(v); else cur.p[N] = v;
  }
  return evs;
}
function icsExpand(start, end, rrule, exset, from, to) {
  const dur = end - start, out = [], r = Object.fromEntries(rrule.split(';').filter(Boolean).map(x => x.split('=')));
  const iv = Math.max(1, +(r.INTERVAL || 1)), count = r.COUNT ? +r.COUNT : Infinity, until = r.UNTIL ? icsDate(r.UNTIL)?.d : null;
  const WD = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];
  const byday = r.BYDAY ? r.BYDAY.split(',').map(x => ({wd: WD.indexOf(x.slice(-2)), n: parseInt(x, 10) || 0})) : null;
  const bymd = r.BYMONTHDAY ? r.BYMONTHDAY.split(',').map(Number) : null;
  const hh = start.getHours(), mm = start.getMinutes(), at = (y, mo, d) => new Date(y, mo, d, hh, mm);
  let n = 0, guard = 0; const lim = until && until < to ? until : to;
  const take = d => { if (d < start) return true; if (until && d > until) return false; if (n >= count) return false; n++; if (!exset.has(d.getTime()) && d <= to && d.getTime() + dur >= from.getTime()) out.push(d); return true; };
  if (r.FREQ === 'DAILY') { for (let d = new Date(start); d <= lim && guard++ < 5000; d = at(d.getFullYear(), d.getMonth(), d.getDate() + iv)) if (!take(d)) break; }
  else if (r.FREQ === 'WEEKLY') {
    const days = (byday ? byday.map(b => b.wd) : [start.getDay()]).sort((a, b) => ((a + 6) % 7) - ((b + 6) % 7));
    let wk = new Date(start.getFullYear(), start.getMonth(), start.getDate() - ((start.getDay() + 6) % 7));
    outer: for (; wk <= lim && guard++ < 3000; wk = new Date(wk.getFullYear(), wk.getMonth(), wk.getDate() + 7 * iv))
      for (const wd of days) { const d = at(wk.getFullYear(), wk.getMonth(), wk.getDate() + ((wd + 6) % 7)); if (d > lim) break outer; if (!take(d)) break outer; }
  } else if (r.FREQ === 'MONTHLY') {
    for (let k = 0; guard++ < 1200; k += iv) {
      const y = start.getFullYear(), mo = start.getMonth() + k, first = new Date(y, mo, 1); if (first > lim) break;
      let cands = [];
      if (byday && byday.some(b => b.n)) byday.forEach(b => { const dim = new Date(first.getFullYear(), first.getMonth() + 1, 0).getDate(); const all = []; for (let d = 1; d <= dim; d++) if (new Date(first.getFullYear(), first.getMonth(), d).getDay() === b.wd) all.push(d); const dd = b.n > 0 ? all[b.n - 1] : all[all.length + b.n]; if (dd) cands.push(dd); });
      else cands = bymd || [start.getDate()];
      let stop = false; for (const dd of cands.sort((a, b) => a - b)) { const d = at(first.getFullYear(), first.getMonth(), dd); if (d.getMonth() !== first.getMonth()) continue; if (d > lim || !take(d)) { stop = true; break; } } if (stop) break;
    }
  } else if (r.FREQ === 'YEARLY') { for (let k = 0; guard++ < 200; k += iv) { const d = at(start.getFullYear() + k, start.getMonth(), start.getDate()); if (d > lim || !take(d)) break; } }
  else if (start <= to && end >= from) out.push(start);
  return out;
}
/* ICS → spotkania TM w zakresie dat; zwraca też statystyki pominiętych */
function icsToMeetings(text, from, to, srcKey = 'outlook') {
  const evs = icsParse(text), overrides = {}, res = [], st = {allDay: 0, cancelled: 0, events: evs.length};
  evs.forEach(e => { if (e.p['RECURRENCE-ID']) { const rid = icsDate(e.p['RECURRENCE-ID'].val, e.p['RECURRENCE-ID'].pm); if (rid) (overrides[e.p.UID?.val] ||= new Set()).add(rid.d.getTime()); } });
  const L = d => `${iso(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  for (const e of evs) {
    if ((e.p.STATUS?.val || '').toUpperCase() === 'CANCELLED') { st.cancelled++; continue; }
    const s = icsDate(e.p.DTSTART?.val, e.p.DTSTART?.pm); if (!s) continue;
    if (s.allDay) { st.allDay++; continue; }
    let en = icsDate(e.p.DTEND?.val, e.p.DTEND?.pm)?.d;
    if (!en) { const du = (e.p.DURATION?.val || '').match(/P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?/); en = new Date(s.d.getTime() + (du ? ((+du[1] || 0) * 1440 + (+du[2] || 0) * 60 + (+du[3] || 0)) * 60000 : 3600000)); }
    const u = e.p.UID?.val || ('noid-' + Math.random()), ex = new Set([...(e.m.EXDATE || []).flatMap(x => x.val.split(',').map(v => icsDate(v, x.pm)?.d?.getTime())), ...(e.p['RECURRENCE-ID'] ? [] : [...(overrides[u] || [])])].filter(Boolean));
    const starts = e.p.RRULE && !e.p['RECURRENCE-ID'] ? icsExpand(s.d, en, e.p.RRULE.val, ex, from, to) : (s.d <= to && en >= from ? [s.d] : []);
    const desc = icsText(e.p.DESCRIPTION?.val), teams = e.p['X-MICROSOFT-SKYPETEAMSMEETINGURL']?.val || (desc.match(/https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s>"]+/) || [])[0] || '';
    const att = (e.m.ATTENDEE || []).map(a => a.val.replace(/^mailto:/i, '')).filter(x => x.includes('@'));
    for (const d of starts) {
      const end2 = new Date(d.getTime() + (en - s.d));
      res.push({id: `${srcKey === 'outlook' ? 'ol' : 'gc'}-${hash(u + '|' + d.getTime()).toString(36)}`, title: icsText(e.p.SUMMARY?.val) || '(bez tytułu)', start: L(d), end: iso(end2) === iso(d) ? L(end2) : `${iso(d)}T23:59`,
        source: srcKey === 'outlook' ? (teams ? 'teams' : 'outlook') : 'google', ctx: srcKey === 'outlook' ? 'job' : 'priv', projectId: null, attendees: att, location: icsText(e.p.LOCATION?.val),
        body: desc.slice(0, 2000), link: teams, imported: srcKey});
    }
  }
  return {meetings: res, stats: st};
}
/* konfiguracja importu (per konto) */
function impCfg(k = 'outlook') {
  S.data.integrations ||= {};
  const c = S.data.integrations[k] ||= {method: 'ics', icsUrl: '', auto: true, every: 60, back: 7, fwd: 62, last: 0, lastOk: 0, log: [], real: false};
  c.every ||= 60; if (!c.fwdV2) { c.fwd = 62; c.fwdV2 = true; } return c;
}
const impRange = c => [add(today(), -c.back), add(today(), c.fwd + 1)];
/* domyślny horyzont importu spotkań: 2 miesiące do przodu (62 dni) */
const IMP_FWD_DEF = 62;
/* scalanie: zastępuje poprzednio zaimportowane spotkania w zakresie, zachowuje przypisany projekt/kontekst/notatki */
function applyImport(list, srcKey, from, to) {
  const a = iso(from), b = iso(to), inR = m => m.start.slice(0, 10) >= a && m.start.slice(0, 10) <= b;
  const prev = S.data.meetings.filter(m => m.imported === srcKey), pmap = Object.fromEntries(prev.map(m => [m.id, m]));
  const keepFields = m => { const o = pmap[m.id]; if (o) { m.projectId = o.projectId; m.ctx = o.ctx; if (o.notesTM) m.notesTM = o.notesTM; } return m; };
  const newIds = new Set(list.map(m => m.id));
  const st = {added: list.filter(m => !pmap[m.id]).length, changed: list.filter(m => pmap[m.id] && (pmap[m.id].start !== m.start || pmap[m.id].end !== m.end || pmap[m.id].title !== m.title)).length,
    removed: prev.filter(m => inR(m) && !newIds.has(m.id)).length, total: list.length};
  S.data.meetings = S.data.meetings.filter(m => !(m.imported === srcKey && (inR(m) || m.start.slice(0, 10) > b))).concat(list.map(keepFields).filter(m => m.start.slice(0, 10) <= b)); // poza horyzontem – usuwane
  const mp = mailCfg().mtgPurge;
  if (mp > 0) { const lim = iso(add(today(), -mp)); S.data.meetings = S.data.meetings.filter(m => !(m.imported === srcKey && m.end.slice(0, 10) < lim && !cmtAll().some(c => c.refId === m.id) && !notesAll().some(n => n.refId === m.id))); }
  if (srcKey === 'outlook') { impCfg().real = true; S.data.meetings = S.data.meetings.filter(m => !(m.demo && isOutlook(m))); }
  return st;
}
let importing = false;
async function runImport({silent = false, srcKey = 'outlook', fileText = null} = {}) {
  const c = impCfg(srcKey); if (importing) return; importing = true; if (!silent && ['calendar', 'meetings', 'integ', 'profile'].includes(S.view)) render();
  const [from, to] = impRange(c), t0 = Date.now(); let entry;
  try {
    let list, extra = {};
    if (fileText != null) { const r = icsToMeetings(fileText, from, to, srcKey); list = r.meetings; extra = r.stats; }
    else if (c.method === 'graph') {
      if (!CFG.MS_CLIENT_ID) throw new Error('Brak MS_CLIENT_ID w config.js – użyj linku ICS albo poproś IT o rejestrację aplikacji');
      if (!MS.app) await MS.init(); if (!MS.acct) { if (silent) throw new Error('Wymagane logowanie do Microsoft 365 – uruchom import ręcznie'); await MS.login(); }
      S.conn.ms = true; list = (await MS.events(iso(from), iso(to))).map(m => Object.assign(m, {id: 'ol-' + m.id.slice(3), imported: srcKey, ext: true}));
    } else {
      if (!c.icsUrl) throw new Error('Nie podano linku ICS do kalendarza Outlook');
      if (!/^https?:$/.test(location.protocol)) throw new Error('Import z linku wymaga uruchomienia przez start-TM.bat');
      const r = await fetch('/__ics?url=' + encodeURIComponent(c.icsUrl), {cache: 'no-store'});
      const txt = await r.text();
      if (!r.ok) throw new Error(r.status === 404 ? 'Działa stara wersja serwera TM – uruchom ponownie start-TM.bat (sam zatrzyma starą wersję) albo zamknij okno „TM - serwer” i uruchom go ponownie' : `Serwer zwrócił ${r.status}: ${txt.slice(0, 160)}`);
      if (!/BEGIN:VCALENDAR/.test(txt)) throw new Error('Link nie zwrócił kalendarza ICS (sprawdź, czy to link .ics z „Publikuj kalendarz”)');
      const p = icsToMeetings(txt, from, to, srcKey); list = p.meetings; extra = p.stats;
    }
    const st = applyImport(list, srcKey, from, to);
    c.last = c.lastOk = Date.now();
    entry = {t: c.last, ok: true, via: fileText != null ? 'plik .ics' : c.method === 'graph' ? 'Microsoft 365' : 'link ICS', ...st, ...extra, ms: Date.now() - t0};
    save(); render();
    if (!silent) toast(`Outlook: ${st.total} spotkań (nowe ${st.added}, zmienione ${st.changed}, usunięte ${st.removed})`);
  } catch (e) {
    c.last = Date.now(); entry = {t: c.last, ok: false, via: fileText != null ? 'plik .ics' : c.method === 'graph' ? 'Microsoft 365' : 'link ICS', err: errMsg(e)};
    save(); if (!silent) toast('Import Outlook: ' + entry.err);
  } finally {
    importing = false; c.log = [entry, ...(c.log || [])].slice(0, 8); save();
    if (['calendar', 'meetings', 'integ', 'profile'].includes(S.view)) render();
    if (mailCfg().mode === 'import') mailTmSync({silent});
  }
}
/* ---------- maile z kategorią „TM” (klasyczny Outlook) → zadania JOB; uruchamiane razem z importem (harmonogram i „Importuj teraz”) ---------- */
let mailTmBusy = false;
/* ustawienia „zadania z maili” (Ustawienia → Integracje → Outlook) */
const mailCfg = () => { const st = SET(); return {on: st.mailOn !== false, cat: (st.mailCat || 'TM').trim() || 'TM', mode: st.mailMode || 'import', every: +st.mailEvery || 15, maxAge: st.mailMaxAge == null ? 30 : +st.mailMaxAge, purge: st.mailPurge == null ? 30 : +st.mailPurge, mtgPurge: +st.mtgPurge || 0}; };
const mailSubj = s => { let x = String(s || '').trim(); for (let i = 0; i < 5; i++) x = x.replace(/^(RE|FW|FWD|PD|ODP|AW|WG|TR)\s*:\s*/i, '').trim(); return x; };
const mailAtOf = t => t.mailAt || (String(t.notes || '').match(/Otrzymano: (\d{4}-\d{2}-\d{2})/) || [])[1] || t.createdAt || '';
/* porządki: zadania z maili starszych niż miesiąc są usuwane (zostają te z zaraportowanym czasem albo komentarzami);
   duplikaty z jednego wątku (ten sam temat) łączone w jedno zadanie – pozostałe maile jako komentarze */
function mailTmPurge() {
  const pd = mailCfg().purge, cutoff = pd > 0 ? iso(add(today(), -pd)) : '0000', keep = t => entriesAll().some(e => e.taskId === t.id) || cmtAll().some(c => c.refId === t.id && !c.mailId) || notesAll().some(n => n.refId === t.id);
  let n = 0;
  const mt = S.data.tasks.filter(t => t.source === 'outlook-mail' || t.mailId);
  const delIds = new Set();
  mt.forEach(t => { const at = mailAtOf(t); if (at && at < cutoff && t.status !== 'done' && !keep(t)) delIds.add(t.id); });
  const by = {}; mt.filter(t => !delIds.has(t.id) && t.status !== 'done').forEach(t => (by[norm(mailSubj(t.title))] ||= []).push(t));
  Object.values(by).filter(g => g.length > 1).forEach(g => { g.sort((a, b) => mailAtOf(a).localeCompare(mailAtOf(b)) || String(a.notes).localeCompare(String(b.notes))); const host = g[0];
    g.slice(1).forEach(t => { if (keep(t)) return; cmtAll().push({id: uid(), kind: 'task', refId: host.id, text: String(t.notes || '').slice(0, 800), date: mailAtOf(t) || TODAY(), reaction: reactList()[0], due: '', done: false, createdAt: Date.now(), mailId: t.mailId}); delIds.add(t.id); }); });
  if (delIds.size) { S.data.tasks = S.data.tasks.filter(t => !delIds.has(t.id)); n = delIds.size; }
  return n;
}
const mailTmCfg = () => (S.data.integrations ||= {}).mailtm ||= {last: 0, lastOk: 0, added: 0, err: ''};
const nextWorkday = () => { let d = add(today(), 1); while (!isWork(d)) d = add(d, 1); return iso(d); };
async function mailTmSync({silent = true} = {}) {
  if (!hasLocalSrv() || mailTmBusy || !S.user || tabInactive) return;
  mailTmBusy = true; const c = mailTmCfg();
  try {
    const mc = mailCfg(); if (!mc.on) { mailTmBusy = false; return; }
    const r = await fetch('/__mail-tm?cat=' + encodeURIComponent(mc.cat), {headers: {'X-TM': '1'}, cache: 'no-store'});
    if (r.status === 404) throw new Error('Uruchom ponownie start-TM.bat (nowa wersja serwera)');
    if (!r.ok) throw new Error((await r.text()).slice(0, 200));
    let list = await r.json(); if (!Array.isArray(list)) list = list ? [list] : [];
    const created = [], ids = [], cutoff = mc.maxAge > 0 ? iso(add(today(), -mc.maxAge)) : '0000';
    let old = 0, threaded = 0;
    // jeden wątek (konwersacja) = jedno zadanie; kolejne maile wątku trafiają jako komentarze do zadania
    list = list.filter(m => m && m.id).sort((a, b) => String(a.received || '').localeCompare(String(b.received || '')));
    for (const m of list) {
      ids.push(m.id);
      if (S.data.tasks.some(t => t.mailId === m.id) || cmtAll().some(x => x.mailId === m.id)) continue;
      const rec = String(m.received || '').slice(0, 10);
      if (rec && rec < cutoff) { old++; continue; }                                  // mail starszy niż miesiąc – bez zadania
      const subj = mailSubj(m.subject), conv = m.conv || '';
      const body = String(m.body || '').replace(/\r\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
      const host = S.data.tasks.find(t => t.mailId && ((conv && t.mailConv === conv) || norm(mailSubj(t.title)) === norm(subj)) && t.status !== 'done');
      if (host) {
        cmtAll().push({id: uid(), kind: 'task', refId: host.id, text: `Mail od ${m.from || '?'}${m.received ? ' (' + m.received.replace('T', ' ') + ')' : ''}:\n${body.slice(0, 600)}`, date: rec || TODAY(), reaction: reactList()[0], due: '', done: false, createdAt: Date.now(), mailId: m.id});
        if (conv && !host.mailConv) host.mailConv = conv;
        threaded++; continue;
      }
      const due = nextWorkday();
      const t = {id: uid(), title: subj || '(mail bez tematu)', ctx: 'job', goalId: null, projectId: null, start: due, due, status: 'todo', prio: 2, est: 0,
        notes: `Mail od: ${m.from || ''}${m.fromEmail ? ' <' + m.fromEmail + '>' : ''}${m.received ? '\nOtrzymano: ' + m.received.replace('T', ' ') : ''}${m.folder ? '\nFolder: ' + m.folder : ''}\n\n${body}`.trim(),
        createdAt: TODAY(), doneAt: null, mailId: m.id, mailConv: conv, mailAt: rec, source: 'outlook-mail'};
      const ic = icFind(m.subject) || icFind(body.slice(0, 400)); if (ic) t.ic = ic;
      const snd = personByEmail(m.fromEmail) || (m.fromEmail && pplCfg().mail ? personUpsert({name: m.from, email: m.fromEmail, source: 'mail TM'}).p : null); if (snd) t.requesterId = snd.id;
      S.data.tasks.push(t); created.push(t);
      notifAdd({type: 'new', key: 'new:' + t.id, title: 'Nowe zadanie z maila: ' + t.title, body: (m.from ? 'od ' + m.from : '') + ' · termin ' + fmt(t.due), ref: {kind: 'task', id: t.id}});
    }
    setTimeout(() => mailPeopleSync({silent: true}), 1500);
    const purged = mailTmPurge();
    if (threaded || purged) save();
    if (threaded && !created.length) toast(`Outlook: ${threaded} ${threaded === 1 ? 'nowy mail dopisany' : 'nowe maile dopisane'} jako komentarz do istniejącego zadania`);
    if (old && !silent) toast(`Pominięto ${old} ${old === 1 ? 'mail starszy' : 'maile starsze'} niż miesiąc`);
    c.last = c.lastOk = Date.now(); c.err = ''; c.added = created.length;
    if (created.length) { notifFlush(); save(); render(); toastAction(`Outlook: ${created.length === 1 ? 'nowe zadanie' : 'nowe zadania'} z maili TM (${created.length}) – JOB, termin ${fmt(created[0].due)}${threaded ? ` · ${threaded} jako komentarze` : ''}`, 'Pokaż', () => { S.view = 'tasks'; S.tasks.status = 'open'; S.tasks.q = ''; history.replaceState(null, '', '#tasks'); render(); }, 12000); }
    else if (!silent) toast('Outlook: brak nowych maili z kategorią TM');
    if (ids.length) await fetch('/__mail-tm-done?cat=' + encodeURIComponent(mc.cat), {method: 'POST', headers: {'X-TM': '1', 'Content-Type': 'application/json'}, body: JSON.stringify({ids})});
  } catch (e) {
    c.last = Date.now(); c.err = errMsg(e);
    if (!silent) toast('Maile TM: ' + c.err);
  } finally { mailTmBusy = false; save(); }
}
const mailTmBadge = () => { if (!hasLocalSrv()) return ''; const c = mailTmCfg();
  return `<span class="impst ${c.err ? 'err' : ''}" title="${c.err ? esc(c.err) : 'Maile z kategorią TM w Outlooku zamieniane są w zadania JOB przy każdym imporcie'}">⟨i:mail⟩ maile TM: ${c.err ? 'błąd' : c.lastOk ? ago(c.lastOk) + (c.added ? ` · +${c.added}` : '') : 'jeszcze nie sprawdzano'}</span>`; };
/* harmonogram: co N minut wg ustawień w profilu (domyślnie 60), sprawdzane co minutę, gdy aplikacja jest otwarta */
function autoImportTick() {
  if (!S.user || !S.data || tabInactive) return; const c = impCfg();
  if (!hasLocalSrv() && c.method !== 'graph') return;           // telefon: spotkania przychodzą z laptopa (wspólny plik)
  const ready = c.auto && (c.method === 'graph' ? !!CFG.MS_CLIENT_ID : !!c.icsUrl);
  if (ready && !importing && Date.now() - (c.last || 0) >= c.every * 60e3) runImport({silent: true});
}
setInterval(autoImportTick, 60e3);
setInterval(() => notifTick(), 60e3);
/* maile TM w osobnym interwale (tryb „co X minut”) */
let mailLast = 0;
setInterval(() => { if (!S.user || !S.data) return; const mc = mailCfg(); if (mc.on && mc.mode === 'interval' && Date.now() - mailLast >= mc.every * 60e3) { mailLast = Date.now(); mailTmSync(); } }, 60e3);
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') autoImportTick(); });
const ago = t => { if (!t) return 'nigdy'; const m = Math.round((Date.now() - t) / 60000); return m < 1 ? 'przed chwilą' : m < 60 ? `${m} min temu` : m < 1440 ? `${Math.floor(m / 60)} h temu` : new Date(t).toLocaleString('pl-PL', {dateStyle: 'short', timeStyle: 'short'}); };
const impReady = (c = impCfg()) => c.method === 'graph' ? !!CFG.MS_CLIENT_ID : !!c.icsUrl;
const nextImport = c => c.auto && impReady(c) ? new Date((c.last || Date.now()) + c.every * 60e3) : null;
function importStatus() {
  const c = impCfg(), l = (c.log || [])[0];
  if (!impReady(c)) return `<span class="impst muted link" data-action="open-profile" data-tab="calendars" title="Ustaw link ICS Outlooka w profilu">${srcBadge('outlook')} Outlook: nie skonfigurowano – ustaw w profilu</span>`;
  const next = nextImport(c);
  return `<span class="impst link ${l && !l.ok ? 'err' : ''}" data-action="open-profile" data-tab="calendars" title="${l && !l.ok ? esc(l.err) : 'Ustawienia importu w profilu · ostatni udany: ' + ago(c.lastOk)}">${srcBadge('outlook')} ${importing ? 'importuję…' : l && !l.ok ? 'błąd importu – szczegóły w profilu' : 'import ' + ago(c.lastOk)}${next && !importing ? ` · następny ${pad(next.getHours())}:${pad(next.getMinutes())}` : ''}${!c.auto ? ' · harmonogram wyłączony' : ''}</span>`;
}
/* na ekranach kalendarza: tylko „Importuj teraz” + status (ustawienia są w profilu użytkownika) */
const importBtn = (sm = '') => `<button class="btn ${sm}" data-action="import-now" ${importing ? 'disabled' : ''} title="${impReady() ? 'Pobierz teraz spotkania z kalendarza Outlook i zadania z maili z kategorią TM' : 'Najpierw ustaw link ICS Outlooka w profilu użytkownika'}">⟨i:refresh⟩ ${importing ? 'Importuję…' : 'Importuj teraz'}</button>${importStatus()}${mailTmBadge()}`;

/* ================= NOTATKI (powiązane z zadaniem lub spotkaniem; zapis do .txt) ================= */
const notesAll = () => (S.data.notes ||= []);
const notesOf = (kind, id) => notesAll().filter(n => n.kind === kind && n.refId === id).sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
const noteObj = n => n.kind === 'task' ? S.data.tasks.find(t => t.id === n.refId) : S.data.meetings.find(m => m.id === n.refId);
const objTitle = (kind, o) => !o ? '' : kind === 'task' ? o.title : `${o.title} (${fmt(o.start.slice(0, 10))} ${hm(o.start)})`;
const noteProj = n => { const o = noteObj(n); return o ? (n.kind === 'task' ? projOf(o) : o.projectId) || null : n.projSnap || null; };
const noteBadge = (kind, id) => { const k = notesOf(kind, id).length; return `<button class="notebadge ${k ? 'has' : ''}" data-action="notes-of" data-kind="${kind}" data-id="${id}" title="${k ? `Notatki: ${k} – przeglądaj` : 'Dodaj notatkę'}">⟨i:note⟩${k ? `<span>${k}</span>` : ''}</button>`; };
const dt = s => s ? new Date(s).toLocaleString('pl-PL', {dateStyle: 'short', timeStyle: 'short'}) : '—';
const safeName = s => String(s || '').replace(/[\\/:*?"<>|\x00-\x1f]/g, '-').replace(/\s+/g, ' ').trim().replace(/[. ]+$/, '').slice(0, 80) || 'notatka';
function noteRel(n) {
  const p = projById(noteProj(n));
  return `Notatki/${safeName(p ? p.title : 'Bez projektu')}/${(n.createdAt || new Date().toISOString()).slice(0, 10)} - ${safeName(n.title)} [${n.id.slice(0, 6)}].txt`;
}
function noteTxt(n) {
  const o = noteObj(n), p = projById(noteProj(n)), L = [];
  L.push(`NOTATKA: ${n.title}`, '');
  L.push(`Utworzono:       ${dt(n.createdAt)}`, `Ostatnia zmiana: ${dt(n.updatedAt)}`);
  if (n.kind === 'task') L.push(`Powiązanie:      Zadanie „${o ? o.title : n.refTitle || '(usunięte)'}”${o ? ` · termin ${o.due || '—'} · status ${({todo: 'do zrobienia', doing: 'w toku', done: 'zrobione'})[o.status] || o.status}` : ''}`);
  else L.push(`Powiązanie:      Spotkanie „${o ? o.title : n.refTitle || '(usunięte)'}”${o ? ` · ${o.start.slice(0, 10)} ${hm(o.start)}–${hm(o.end)} · ${SRC[o.source]?.[0] || o.source}` : ''}`);
  if (n.kind === 'mtg' && o && o.attendees?.length) L.push(`Uczestnicy:      ${o.attendees.join(', ')}`);
  L.push(`Projekt:         ${p ? `${p.title} (${p.ctx.toUpperCase()})` : '—'}`, `Autor:           ${S.user.name} <${S.user.email}>`);
  L.push('-'.repeat(64), '', n.body || '', '');
  return L.join('\r\n');
}
function downloadTxt(n) { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob(['\ufeff' + noteTxt(n)], {type: 'text/plain;charset=utf-8'})); a.download = noteRel(n).split('/').pop(); a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 2000); }
/* zapis pliku .txt przez serwer lokalny (folder Notatki/<projekt>/ obok aplikacji); bez serwera – pobranie pliku */
async function writeNoteFile(n, oldRel) {
  const rel = noteRel(n);
  if (/^https?:$/.test(location.protocol)) {
    try {
      const r = await fetch('/__note', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({rel, oldRel: oldRel && oldRel !== rel ? oldRel : '', content: noteTxt(n)})});
      if (r.ok) { const j = await r.json(); n.file = rel; n.fileAbs = j.path; n.fileOk = true; return 'disk'; }
      if (r.status !== 404) { const t = await r.text(); toast('Zapis pliku: ' + t.slice(0, 140)); }
    } catch (e) { /* serwer niedostępny */ }
  }
  n.file = rel; n.fileOk = false; downloadTxt(n); return 'download';
}
async function deleteNoteFile(rel) { if (!rel || !/^https?:$/.test(location.protocol)) return; try { await fetch('/__note-del', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({rel})}); } catch { } }
function retTo(r) {
  if (!r) return () => { closeModal(); render(); };
  const [a, b, c] = r.split(':');
  if (a === 'task') return () => { render(); const t = S.data.tasks.find(x => x.id === b); t ? editTask(t) : closeModal(); };
  if (a === 'mtg') return () => { render(); const m = S.data.meetings.find(x => x.id === b); m ? editMeeting(m) : closeModal(); };
  if (a === 'list') return () => { render(); notesListModal(b, c); };
  return () => { closeModal(); render(); };
}
function linkOptions(sel, projectId) {
  const inP = o => !projectId || o === projectId;
  const ts = S.data.tasks.filter(t => inP(projOf(t)) && (inCtx(t) || `task:${t.id}` === sel)).sort((a, b) => (a.status === 'done') - (b.status === 'done') || (a.due || '').localeCompare(b.due || ''));
  const lo = iso(add(today(), -60)), hi = iso(add(today(), 60));
  const ms = S.data.meetings.filter(m => inP(m.projectId) && ((inCtx(m) && m.start.slice(0, 10) >= lo && m.start.slice(0, 10) <= hi) || `mtg:${m.id}` === sel)).sort((a, b) => b.start.localeCompare(a.start));
  return `<optgroup label="Zadania">${ts.map(t => `<option value="task:${t.id}" ${sel === `task:${t.id}` ? 'selected' : ''}>${esc(t.title)}${t.status === 'done' ? ' (zrobione)' : t.due ? ' · ' + fmt(t.due) : ''}</option>`).join('')}</optgroup>
    <optgroup label="Spotkania (±60 dni)">${ms.map(m => `<option value="mtg:${m.id}" ${sel === `mtg:${m.id}` ? 'selected' : ''}>${fmt(m.start.slice(0, 10))} ${hm(m.start)} · ${esc(m.title)}</option>`).join('')}</optgroup>`;
}
function noteEditor(n, preset = {}) {
  const isNew = !n;
  if (isNew) { const o = preset.kind && (preset.kind === 'task' ? S.data.tasks.find(t => t.id === preset.id) : S.data.meetings.find(m => m.id === preset.id));
    n = {id: uid(), kind: preset.kind || '', refId: preset.id || '', title: o ? `${o.title} – notatka ${fmt(TODAY())}` : `Notatka ${fmt(TODAY())}`, body: preset.kind === 'mtg' && o ? `Uczestnicy: ${(o.attendees || []).join(', ') || '—'}\n\nUstalenia:\n- \n\nZadania do wykonania:\n- \n` : ''}; }
  const sel = n.kind ? `${n.kind}:${n.refId}` : '', o = noteObj(n);
  modal(`<div class="note-ed"><h3>⟨i:note⟩ ${isNew ? 'Nowa notatka' : 'Notatka'}</h3>
    <form data-form="note" data-id="${isNew ? '' : n.id}" data-ret="${esc(preset.ret || '')}" class="grid" style="gap:12px">
      <label class="f">Tytuł<input class="input" name="title" required value="${esc(n.title)}" maxlength="120"></label>
      <label class="f">Powiązanie (zadanie lub spotkanie)<select class="input" name="link" required><option value="">— wybierz —</option>${linkOptions(sel, preset.project)}${sel && !o ? `<option value="${sel}" selected>(usunięte) ${esc(n.refTitle || '')}</option>` : ''}</select></label>
      <label class="f">Treść<textarea class="input note-body" name="body" rows="14" placeholder="Treść notatki…">${esc(n.body || '')}</textarea></label>
      <div class="hint note-meta">${isNew ? 'Po zapisaniu notatka zostanie zapisana także jako plik <b>.txt</b> w folderze <code>Notatki\\&lt;projekt&gt;\\</code> obok aplikacji.' : `Utworzono ${dt(n.createdAt)} · zmieniono ${dt(n.updatedAt)}<br>Plik: <code>${esc((n.file || noteRel(n)).replace(/\//g, '\\'))}</code> ${n.fileOk ? '<span class="p ok">● zapisany na dysku</span>' : '<span class="p warn">● pobrany (bez zapisu na dysk)</span>'}`}</div>
      <div class="row" style="justify-content:space-between">${isNew ? '<span></span>' : `<div class="row"><button type="button" class="btn danger" data-action="note-del" data-id="${n.id}" data-ret="${esc(preset.ret || '')}">⟨i:trash⟩ Usuń</button><button type="button" class="btn" data-action="note-dl" data-id="${n.id}">⟨i:download⟩ Pobierz .txt</button>${n.fileOk ? `<button type="button" class="btn" data-action="note-reveal" data-id="${n.id}">⟨i:folder⟩ Pokaż w folderze</button>` : ''}</div>`}
        <div class="row"><button type="button" class="btn" data-action="note-cancel" data-ret="${esc(preset.ret || '')}">Anuluj</button><button class="btn primary">⟨i:check⟩ Zapisz notatkę</button></div></div>
      <div class="hint" style="margin:0">Ctrl+S – zapisz</div>
    </form></div>`, m => { const b = $('[name=body]', m); (isNew ? $('[name=title]', m) : b).focus(); if (!isNew) b.setSelectionRange(b.value.length, b.value.length);
      m.addEventListener('keydown', e => { if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); $('form[data-form=note]', m).requestSubmit(); } }); });
}
function noteItem(n, ret, opts = {}) {
  const P_ = S.privacy ? pv : (x => x), o = noteObj(n);
  return `<div class="note-item" data-action="note-edit" data-id="${n.id}" data-ret="${esc(ret || '')}">
    <div class="ni-h">⟨i:note⟩ <b>${P_(esc(n.title))}</b><span class="muted">${dt(n.updatedAt)}</span>${n.fileOk ? '<span class="muted" title="Zapisany jako .txt">.txt</span>' : ''}</div>
    ${opts.link ? `<div class="muted ni-l">${n.kind === 'task' ? '⟨i:check-square⟩ Zadanie' : '⟨i:video⟩ Spotkanie'}: ${P_(esc(o ? objTitle(n.kind, o) : (n.refTitle || '') + ' (usunięte)'))}</div>` : ''}
    <div class="ni-s">${P_(esc((n.body || '').replace(/\s+/g, ' ').slice(0, 160)) || '<span class="muted">(pusta)</span>')}</div></div>`;
}
/* sekcja „Notatki” w oknie zadania / spotkania */
function notesSection(kind, id) {
  const l = notesOf(kind, id), ret = `${kind}:${id}`;
  return `<div class="notes-sec" style="grid-column:1/-1"><div class="row"><b>⟨i:note⟩ Notatki (${l.length})</b><div class="spacer"></div>
    ${l.length > 3 ? `<button type="button" class="btn sm ghost" data-action="notes-of" data-kind="${kind}" data-id="${id}">Wszystkie</button>` : ''}<button type="button" class="btn sm" data-action="note-new" data-kind="${kind}" data-id="${id}" data-ret="${ret}">⟨i:plus⟩ Dodaj notatkę</button></div>
    ${l.slice(0, 3).map(n => noteItem(n, ret)).join('') || '<div class="muted" style="font-size:12px;padding:6px 0">Brak notatek</div>'}</div>`;
}
/* okno z listą notatek obiektu */
function notesListModal(kind, id) {
  const o = kind === 'task' ? S.data.tasks.find(t => t.id === id) : S.data.meetings.find(m => m.id === id), l = notesOf(kind, id), ret = `list:${kind}:${id}`, P_ = S.privacy ? pv : (x => x);
  modal(`<h3>⟨i:note⟩ Notatki · ${kind === 'task' ? 'zadanie' : 'spotkanie'} „${P_(esc(o ? objTitle(kind, o) : ''))}”</h3>
    <div class="notes-list">${l.map(n => noteItem(n, ret)).join('') || '<div class="empty">Brak notatek – dodaj pierwszą</div>'}</div>
    <div class="row" style="justify-content:space-between;margin-top:12px"><button class="btn" data-action="${kind === 'task' ? 'edit-task' : 'edit-mtg'}" data-id="${id}">${kind === 'task' ? '⟨i:check-square⟩ Otwórz zadanie' : '⟨i:video⟩ Otwórz spotkanie'}</button>
      <div class="row"><button class="btn" data-action="modal-close">Zamknij</button><button class="btn primary" data-action="note-new" data-kind="${kind}" data-id="${id}" data-ret="${ret}">⟨i:plus⟩ Dodaj notatkę</button></div></div>`);
}
async function deleteNote(id, ret) {
  const i = notesAll().findIndex(n => n.id === id); if (i < 0) return; const n = notesAll()[i];
  S.data.notes.splice(i, 1); save(); await deleteNoteFile(n.file); retTo(ret)();
  toastAction(S.privacy ? 'Usunięto notatkę' : `Usunięto notatkę „${n.title}”`, '⟨i:undo⟩ Cofnij', async () => { S.data.notes.splice(Math.min(i, S.data.notes.length), 0, n); if (n.fileOk) await writeNoteFile(n); save(); render(); toast('Przywrócono notatkę'); });
}
/* sekcja „Notatki projektu” na ekranie Projekt */
function projNotesSection(p) {
  const P0 = S.proj, q = norm(P0.nq || ''), P_ = S.privacy ? pv : (x => x);
  let l = notesAll().filter(n => noteProj(n) === p.id);
  const cnt = {all: l.length, task: l.filter(n => n.kind === 'task').length, mtg: l.filter(n => n.kind === 'mtg').length};
  if (P0.nk && P0.nk !== 'all') l = l.filter(n => n.kind === P0.nk);
  if (q) l = l.filter(n => norm([n.title, n.body, objTitle(n.kind, noteObj(n)), n.refTitle].join(' ')).includes(q));
  l.sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
  const rows = l.map(n => { const o = noteObj(n); return `<tr class="nrow">
    <td class="nowrap">${dt(n.updatedAt)}</td>
    <td class="mt" data-action="note-edit" data-id="${n.id}"><b>${P_(esc(n.title))}</b><div class="muted" style="font-size:12px">${P_(esc((n.body || '').replace(/\s+/g, ' ').slice(0, 110)))}</div></td>
    <td>${o ? `<span class="crumb" data-action="${n.kind === 'task' ? 'edit-task' : 'edit-mtg'}" data-id="${o.id}">${n.kind === 'task' ? '⟨i:check-square⟩' : '⟨i:video⟩'} ${P_(esc(objTitle(n.kind, o)))}</span>` : `<span class="muted">${n.kind === 'task' ? 'Zadanie' : 'Spotkanie'} (usunięte): ${P_(esc(n.refTitle || ''))}</span>`}</td>
    <td class="nowrap acts">${n.fileOk ? `<button class="btn sm ghost" data-action="note-reveal" data-id="${n.id}" title="Pokaż plik .txt w folderze">⟨i:folder⟩</button>` : ''}<button class="btn sm ghost" data-action="note-dl" data-id="${n.id}" title="Pobierz .txt">⟨i:download⟩</button><button class="btn sm ghost" data-action="note-edit" data-id="${n.id}" title="Otwórz">⟨i:pencil⟩</button><button class="del btn sm ghost" data-action="note-del" data-id="${n.id}" title="Usuń notatkę">⟨i:trash⟩</button></td></tr>`; }).join('');
  return `<div class="card" id="psNotes" style="padding:0;margin-top:16px"><div class="row card-h"><h3 style="margin:0">⟨i:note⟩ Notatki projektu</h3><span class="muted">${cnt.all} · z zadań ${cnt.task} · ze spotkań ${cnt.mtg}</span><div class="spacer"></div>
    <input class="input" style="max-width:220px;padding:6px 10px 6px 34px" placeholder="Szukaj w notatkach…" data-bind="proj.nq" value="${esc(P0.nq || '')}">
    <div class="seg">${[['all', 'Wszystkie'], ['task', 'Zadania'], ['mtg', 'Spotkania']].map(([k, lb]) => `<button data-action="set" data-k="proj.nk" data-v="${k}" class="${(P0.nk || 'all') === k ? 'on' : ''}">${lb}</button>`).join('')}</div>
    <button class="btn sm" data-action="note-new" data-project="${p.id}">⟨i:plus⟩ Notatka</button></div>
    <table class="tbl mtbl"><thead><tr><th>Zmieniono</th><th>Notatka</th><th>Powiązanie</th><th></th></tr></thead>
    <tbody>${rows || `<tr><td colspan="4" class="empty">${cnt.all ? 'Brak notatek dla filtra' : 'Brak notatek. Dodaj notatkę w oknie zadania lub spotkania tego projektu albo przyciskiem „Notatka”.'}</td></tr>`}</tbody></table></div>`;
}

/* ================= REALIZACJA (stoper, wpisy czasu, raport) ================= */
const entriesAll = () => (S.data.timeEntries ||= []);
const entriesC = () => entriesAll().filter(inCtx);
const Lms = ms => { const d = new Date(ms); return `${iso(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const strMs = s => new Date(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10), +s.slice(11, 13), +s.slice(14, 16)).getTime();
const entMin = e => Math.max(0, ((e.endMs || strMs(e.end)) - (e.startMs || strMs(e.start))) / 60000);
const fmtH = min => { if (min > 0 && min < 1) return '<1m'; const t = Math.round(min), h = Math.floor(t / 60), m = t % 60; return h ? `${h}h${m ? ' ' + pad(m) + 'm' : ''}` : `${m}m`; };
const fmtClock = ms => { const s = Math.max(0, Math.floor(ms / 1000)); return `${pad(Math.floor(s / 3600))}:${pad(Math.floor(s / 60) % 60)}:${pad(s % 60)}`; };
const timerOn = () => S.data && S.data.timer;
/* bieżąca sesja stopera jako „wpis” (do kalendarza, wykresów, sum) */
const liveEntry = () => { const t = timerOn(); if (!t) return null; const now = Date.now(); return {id: '__live', live: true, taskId: t.taskId, title: t.title, projectId: t.projectId, ctx: t.ctx, startMs: t.startMs, endMs: now, start: Lms(t.startMs), end: Lms(Math.max(now, t.startMs + 60000)), note: t.note || ''}; };
const entriesWithLive = () => { const l = liveEntry(); return l && inCtx(l) ? [...entriesC(), l] : entriesC(); };
const taskSpent = id => entriesAll().filter(e => e.taskId === id).reduce((a, e) => a + entMin(e), 0) + (timerOn()?.taskId === id ? Math.round((Date.now() - timerOn().startMs) / 60000) : 0);
const entColor = e => projColor(projById(e.projectId)) || '#64748B';
/* start: istniejące zadanie lub nowe (opis), opcjonalnie projekt */
function startTimer({taskId, title, projectId, free, mtgId}) {
  if (mtgId) return mtgTimer(mtgId);
  { const t0 = taskId && S.data.tasks.find(x => x.id === taskId), pj0 = projById(projectId);
    if ((t0 && t0.ctx === 'priv') || (!t0 && pj0 && pj0.ctx === 'priv')) return toast('Rejestracja czasu (start/stop) działa tylko dla zadań JOB – zadań prywatnych nie mierzymy'); }
  if (timerOn()) stopTimer('switch', true);
  let t = taskId && S.data.tasks.find(x => x.id === taskId);
  if (!t && free) { const pj = projById(projectId); S.data.timer = {taskId: null, title: (title || 'Realizacja bez zadania').trim(), projectId: pj ? pj.id : null, ctx: 'job', startMs: Date.now(), note: '', free: true};
    S.trk = {}; save(); render(); return toast(`Start: ${S.data.timer.title} – czas trafi do RCP bez tworzenia zadania`); }
  if (!t) {
    title = (title || '').trim(); if (!title) return toast('Opisz, co robisz, albo wybierz zadanie');
    const pj = projById(projectId);
    t = {id: uid(), title, ctx: 'job', goalId: null, projectId: pj ? pj.id : null, start: TODAY(), due: TODAY(), status: 'doing', prio: 2, est: 0, notes: '', createdAt: TODAY(), doneAt: null};
    S.data.tasks.push(t);
  } else if (t.status !== 'doing') { t.status = 'doing'; t.doneAt = null; }
  S.data.timer = {taskId: t.id, title: t.title, projectId: projOf(t) || t.projectId || null, ctx: t.ctx, startMs: Date.now(), note: ''};
  S.trk = {}; save(); render(); toast(`Start: ${t.title}`);
}
/* zakończenie sesji: kind = 'pause' (przerwa), 'done' (zadanie zakończone), 'switch' (przełączenie) */
function stopTimer(kind, quiet) {
  const tm = timerOn(); if (!tm) return;
  const endMs = Date.now(), min = (endMs - tm.startMs) / 60000, t = S.data.tasks.find(x => x.id === tm.taskId);
  S.data.timer = null; if (tm.taskId) S.data.lastTask = tm.taskId; let saved = null;
  if (endMs - tm.startMs < 5000) { if (!quiet) toast('Sesja krótsza niż 5 sekund – nie zapisano wpisu'); }
  else {
    // podział na dni, jeśli sesja przeszła przez północ
    let s = tm.startMs;
    while (s < endMs) { const d = new Date(s), mid = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1).getTime(), e = Math.min(endMs, mid);
      if (e - s >= 5000) { saved = {id: uid(), taskId: tm.taskId, ...(tm.mtgId ? {mtgId: tm.mtgId} : {}), title: t ? t.title : tm.title, projectId: tm.projectId, ctx: tm.ctx, startMs: s, endMs: e, start: Lms(s), end: Lms(Math.max(e, s + 60000)), note: tm.note || '', kind}; entriesAll().push(saved); }
      s = e; }
  }
  if (t && kind === 'done') { t.status = 'done'; t.doneAt = TODAY(); }
  save(); render();
  if (!quiet && saved) toastAction(kind === 'done' ? `Zakończono „${t?.title || tm.title}” · zapisano ${fmtH(min)} ${min >= 1 ? `(${hm(saved.start)}–${hm(Lms(endMs))})` : ''}` : `Przerwa · zapisano ${fmtH(min)} realizacji ${min >= 1 ? `(${hm(saved.start)}–${hm(Lms(endMs))})` : ''}`, 'Pokaż w ewidencji', () => { S.tl = Object.assign(S.tl || {}, {period: 'today', proj: '', q: ''}); S.view = 'timelog'; history.replaceState(null, '', '#timelog'); render(); }, 7000);
}
/* zegar na żywo (pasek górny + ekran Stoper) */
setInterval(() => {
  const t = timerOn(); if (!t || !S.user) return;
  $$('[data-timer]').forEach(el => el.textContent = fmtClock(Date.now() - t.startMs));
  if (Math.floor(Date.now() / 1000) % 10 === 0) playerRender();
}, 1000);
setInterval(() => { // co minutę odśwież widoki z „żywym” paskiem realizacji
  if (!timerOn() || !S.user || $('#modalRoot').innerHTML) return;
  if (/INPUT|TEXTAREA|SELECT/.test(document.activeElement?.tagName)) return;
  if (['dashboard', 'calendar', 'track'].includes(S.view)) render();
}, 60000);
function timerChipHtml() {
  const t = timerOn(); if (!t) return '';
  return `<span class="tc-dot"></span><span data-timer>${fmtClock(Date.now() - t.startMs)}</span><span class="tc-t">${S.privacy ? pv(esc(t.title)) : esc(t.title)}</span>`;
}
/* wykres słupkowy godzin 8–16: minuty realizacji w każdej godzinie, w kolorach projektów */
function realChart(ds, H0 = 8, H1 = 16) {
  const es = entriesWithLive().filter(e => e.start.slice(0, 10) === ds), W = 640, Ht = 190, pl = 34, pb = 24, pt = 8, bw = (W - pl - 10) / (H1 - H0), y = v => pt + (Ht - pt - pb) * (1 - v / 60);
  let s = `<svg viewBox="0 0 ${W} ${Ht}" width="100%" class="realchart">`;
  [0, 15, 30, 45, 60].forEach(v => s += `<line x1="${pl}" x2="${W - 10}" y1="${y(v)}" y2="${y(v)}" stroke="var(--line)"/><text x="${pl - 6}" y="${y(v) + 4}" text-anchor="end">${v}</text>`);
  for (let h = H0; h < H1; h++) {
    const a = strMs(`${ds}T${pad(h)}:00`), b = a + 3600e3, x = pl + (h - H0) * bw + bw * .14, w = bw * .72; let acc = 0;
    const parts = {}; es.forEach(e => { const o = Math.min(b, e.endMs || strMs(e.end)) - Math.max(a, e.startMs || strMs(e.start)); if (o > 0) { const k = e.projectId || '-'; (parts[k] ||= {min: 0, titles: new Set(), e}).min += o / 60000; parts[k].titles.add(e.title); } });
    Object.entries(parts).forEach(([k, p]) => { const v = Math.min(60 - acc, p.min), y1 = y(acc + v), hgt = y(acc) - y1; acc += v;
      s += `<rect x="${x}" y="${y1}" width="${w}" height="${Math.max(0, hgt - 1)}" rx="3" fill="${entColor(p.e)}" opacity="${p.e.live ? .65 : .9}"><title>${pad(h)}:00–${pad(h + 1)}:00 · ${Math.round(p.min)} min · ${esc(projById(k)?.title || 'bez projektu')}${S.privacy ? '' : ' · ' + esc([...p.titles].join(', '))}</title></rect>`; });
    s += `<text x="${x + w / 2}" y="${Ht - 6}" text-anchor="middle">${h}:00</text>`;
  }
  return s + '</svg>';
}
function dayTotals(ds) {
  const es = entriesWithLive().filter(e => e.start.slice(0, 10) === ds), by = {};
  es.forEach(e => { const k = e.projectId || '-'; by[k] = (by[k] || 0) + entMin(e); });
  return {total: es.reduce((a, e) => a + entMin(e), 0), by: Object.entries(by).sort((a, b) => b[1] - a[1]), count: es.length};
}
/* karta na Pulpicie (pod kafelkami) */
function dashRealCard() {
  const td = TODAY(), tot = dayTotals(td), t = timerOn();
  return `<div class="card realcard" style="margin-bottom:16px"><h3>⟨i:timer⟩ Realizacja dziś · 8:00–16:00 <span class="muted" style="font-weight:500">· razem ${fmtH(tot.total)}</span>
    <span class="right row" style="gap:6px">${t ? `<button class="btn sm" data-action="timer-pause">⟨i:pause⟩ Przerwa</button><button class="btn sm" data-action="timer-done">⟨i:check⟩ Zakończ</button>` : `<button class="btn sm primary" data-action="nav" data-view="track">⟨i:play⟩ Start realizacji</button>`}<button class="btn sm" data-action="nav" data-view="timereport">Raport</button></span></h3>
    <div class="realgrid"><div>${realChart(td)}</div><div class="realside">
      ${t ? `<div class="live-box"><span class="tc-dot"></span><b data-timer>${fmtClock(Date.now() - t.startMs)}</b><div class="muted" style="font-size:12px">${S.privacy ? pv(esc(t.title)) : esc(t.title)}</div></div>` : ''}
      ${tot.by.map(([k, m]) => `<div class="rs-row"><span class="pcol" style="background:${projColor(projById(k)) || '#64748B'}"></span><span>${esc(projById(k)?.title || 'Bez projektu')}</span><b>${fmtH(m)}</b></div>`).join('') || '<div class="muted" style="font-size:13px">Brak realizacji dziś. Uruchom stoper w module Realizacja.</div>'}
    </div></div></div>`;
}
/* tygodniowa realizacja: godziny na dzień (słupki wg projektów) vs target 8h w dni robocze */
function weekChart(ws) {
  const days = [...Array(7)].map((_, i) => { const d = add(P(ws), i), ds = iso(d), t = dayTotals(ds); return {d, ds, t, tgt: capacityOfDay(d, 'job', true) * 60, type: dayType(d)}; });
  const maxH = Math.max(10, Math.ceil(Math.max(...days.map(x => x.t.total)) / 60) + 1);
  const W = 640, Ht = 230, pl = 34, pr = 10, pt = 22, pb = 40, cw = (W - pl - pr) / 7, y = m => pt + (Ht - pt - pb) * (1 - m / (maxH * 60));
  let s = `<svg viewBox="0 0 ${W} ${Ht}" width="100%" class="realchart weekchart">`;
  for (let h = 0; h <= maxH; h += 2) s += `<line x1="${pl}" x2="${W - pr}" y1="${y(h * 60)}" y2="${y(h * 60)}" stroke="var(--line)"/><text x="${pl - 6}" y="${y(h * 60) + 4}" text-anchor="end">${h}h</text>`;
  days.forEach((x, i) => {
    const x0 = pl + i * cw, bx = x0 + cw * .2, bw = cw * .6, isToday = x.ds === TODAY(), future = x.ds > TODAY();
    if (x.type !== 'work') s += `<rect x="${x0 + 2}" y="${pt}" width="${cw - 4}" height="${Ht - pt - pb}" fill="var(--line)" opacity=".35" rx="6"><title>${x.type === 'hol' ? esc(holName(x.d)) : x.type === 'vac' ? 'Urlop' : 'Weekend'}</title></rect>`;
    if (x.tgt) s += `<rect x="${bx - 4}" y="${y(x.tgt)}" width="${bw + 8}" height="${y(0) - y(x.tgt)}" rx="5" class="wk-target"/><line x1="${bx - 8}" x2="${bx + bw + 8}" y1="${y(x.tgt)}" y2="${y(x.tgt)}" class="wk-tline"/>`;
    let acc = 0;
    x.t.by.forEach(([k, m]) => { const y1 = y(acc + m), hgt = y(acc) - y1; acc += m;
      s += `<rect x="${bx}" y="${y1}" width="${bw}" height="${Math.max(0, hgt - 1)}" rx="3" fill="${projColor(projById(k)) || '#64748B'}" opacity=".9"><title>${['Pn', 'Wt', 'Śr', 'Cz', 'Pt', 'Sb', 'Nd'][i]} ${fmt(x.ds)} · ${fmtH(m)} · ${esc(projById(k)?.title || 'bez projektu')}</title></rect>`; });
    const tot = x.t.total, cls = !x.tgt ? (tot ? 'wk-ok' : '') : tot >= x.tgt ? 'wk-ok' : future ? 'wk-fut' : tot >= x.tgt * .75 ? 'wk-mid' : 'wk-low';
    if (tot || (x.tgt && !future)) s += `<text x="${bx + bw / 2}" y="${Math.min(y(Math.max(tot, 0)), y(x.tgt || 0)) - 6}" text-anchor="middle" class="wk-val ${cls}">${tot ? fmtH(tot) : '0h'}${x.tgt ? ` / ${x.tgt / 60}h` : ''}</text>`;
    s += `<text x="${x0 + cw / 2}" y="${Ht - 22}" text-anchor="middle" class="${isToday ? 'wk-today' : ''}">${['Pn', 'Wt', 'Śr', 'Cz', 'Pt', 'Sb', 'Nd'][i]}</text><text x="${x0 + cw / 2}" y="${Ht - 8}" text-anchor="middle" class="muted">${pad(x.d.getDate())}.${pad(x.d.getMonth() + 1)}</text>`;
  });
  return {svg: s + '</svg>', total: days.reduce((a, x) => a + x.t.total, 0), target: days.reduce((a, x) => a + x.tgt, 0), sofar: days.filter(x => x.ds <= TODAY()).reduce((a, x) => a + x.tgt, 0)};
}
function dashWeekCard() {
  const off = S.wkOff || 0, ws = iso(add(sow(today()), off * 7)), we = iso(add(P(ws), 6)), c = weekChart(ws), pct = c.target ? Math.round(100 * c.total / c.target) : 0;
  const due = off === 0 ? c.sofar : off < 0 ? c.target : 0, diff = c.total - due;
  return `<div class="card realcard" style="margin-bottom:16px"><h3>⟨i:chart⟩ Realizacja tydzień · ${fmt(ws)} – ${fmt(we)} <span class="muted" style="font-weight:500">· ${fmtH(c.total)} z ${c.target / 60}h (${pct}%)</span>
    <span class="right row" style="gap:6px"><button class="btn sm" data-action="wk-off" data-d="-1" title="Poprzedni tydzień">⟨i:chev-left⟩</button>${off ? '<button class="btn sm" data-action="wk-off" data-d="0">Ten tydzień</button>' : ''}<button class="btn sm" data-action="wk-off" data-d="1" title="Następny tydzień">⟨i:chev-right⟩</button></span></h3>
    ${c.svg}
    <div class="row wk-legend"><span><i class="wk-sw"></i> target ${CFG.HOURS_JOB_PER_DAY}h / dzień roboczy</span><span><b class="wk-ok">●</b> target osiągnięty</span><span><b class="wk-mid">●</b> ≥ 75%</span><span><b class="wk-low">●</b> poniżej 75%</span>
      ${due ? `<span class="spacer"></span><span class="${diff >= 0 ? 'wk-ok' : 'wk-low'}" style="font-weight:600">${diff >= 0 ? '+' : '−'}${fmtH(Math.abs(diff))} ${off === 0 ? 'względem targetu do dziś' : 'względem targetu tygodnia'}</span>` : ''}</div></div>`;
}
/* ---------- ekrany ---------- */
function trackTabs() {
  const t = timerOn();
  return ''; return `<div class="seg caltabs" style="margin-bottom:12px"><button data-action="nav" data-view="track" class="${S.view === 'track' ? 'on' : ''}">⟨i:timer⟩ Stoper${t ? ' <span class="tc-dot"></span>' : ''}</button><button data-action="nav" data-view="timelog" class="${S.view === 'timelog' ? 'on' : ''}">⟨i:list⟩ Ewidencja czasu</button><button data-action="nav" data-view="timereport" class="${S.view === 'timereport' ? 'on' : ''}">⟨i:chart⟩ Raport realizacji</button></div>`;
}
function vTrack() {
  const t = timerOn(), td = TODAY(), tot = dayTotals(td);
  let top;
  if (t) {
    const task = S.data.tasks.find(x => x.id === t.taskId), pj = projById(t.projectId), g = task && goalById(task.goalId);
    const spent = taskSpent(t.taskId), est = (task?.est || 0) * 60, pct = est ? Math.min(100, Math.round(100 * spent / est)) : 0, long = Date.now() - t.startMs > 10 * 3600e3;
    top = `<div class="card track-live ${t.ctx}">
      <div class="muted tl-lbl"><span class="tc-dot"></span> Realizujesz od ${pad(new Date(t.startMs).getHours())}:${pad(new Date(t.startMs).getMinutes())}</div>
      <div class="tl-title" data-action="edit-task" data-id="${t.taskId}">${esc(t.title)}</div>
      <div class="row" style="gap:6px;margin:6px 0 14px">${ctxTag(t.ctx)}${pj ? projTag(pj.id) : '<span class="tag src">bez projektu</span>'}${g ? `<span class="tag src">⟨i:layers⟩ ${esc(g.title.slice(0, 30))}</span>` : ''}${task?.due ? `<span class="tag src">⟨i:calendar⟩ termin ${fmt(task.due)}</span>` : ''}</div>
      <div class="tl-clock" data-timer>${fmtClock(Date.now() - t.startMs)}</div>
      ${long ? '<div class="hint lockinfo" style="margin:8px 0">Stoper działa ponad 10 godzin – po zakończeniu skoryguj wpis w Ewidencji czasu.</div>' : ''}
      <div class="tl-prog"><div class="row" style="justify-content:space-between;font-size:12px"><span class="muted">Łącznie na zadaniu: <b>${fmtH(spent)}</b>${est ? ` z ${fmtH(est)} (estymacja)` : ''}</span><span class="muted">Dziś razem: <b>${fmtH(tot.total)}</b></span></div>
        ${est ? `<div class="bar" style="margin-top:6px"><i style="width:${pct}%;background:${spent > est ? 'var(--bad)' : 'var(--brand)'}"></i></div>` : ''}</div>
      <label class="f" style="margin-top:14px">Notatka do tej sesji (trafi do wpisu w kalendarzu)<textarea class="input" id="tmNote" rows="2" placeholder="Co konkretnie robisz?">${esc(t.note || '')}</textarea></label>
      <div class="row tl-btns"><button class="btn big" data-action="timer-pause">⟨i:pause⟩ Wstrzymaj (przerwa)</button><button class="btn big primary" data-action="timer-done">⟨i:check⟩ Zakończ zadanie</button>
        <div class="spacer"></div><button class="btn" data-action="timer-switch">⟨i:switch⟩ Przełącz zadanie</button><button class="btn" data-action="note-new" data-kind="task" data-id="${t.taskId}">⟨i:note⟩ Notatka</button></div></div>`;
  } else {
    const paused = S.data.tasks.filter(x => inCtx(x) && x.status === 'doing' && entriesAll().some(e => e.taskId === x.id)).sort((a, b) => taskSpent(b.id) - taskSpent(a.id)).slice(0, 6);
    const recentIds = [...new Set(entriesC().sort((a, b) => (b.endMs || 0) - (a.endMs || 0)).map(e => e.taskId))].filter(id => !paused.some(p => p.id === id)).slice(0, 6);
    const recent = recentIds.map(id => S.data.tasks.find(x => x.id === id)).filter(x => x && x.status !== 'done');
    const sel = S.trk?.taskId && S.data.tasks.find(x => x.id === S.trk.taskId);
    top = `<div class="card track-start"><h3>⟨i:play⟩ Co teraz robisz?</h3>
      <div class="ts-in"><input class="input ts-q" id="trkQ" placeholder="Opisz, co robisz, albo wyszukaj istniejące zadanie…" autocomplete="off" value="${esc(sel ? sel.title : (S.trk?.q || ''))}">
        <div id="trkSug" class="ts-sug"></div></div>
      ${sel ? `<div class="qa-tplsel" style="margin-top:8px">⟨i:check-square⟩ Istniejące zadanie: <b>${esc(sel.title)}</b><button type="button" class="btn sm ghost" data-action="trk-clear">⟨i:x⟩</button></div>` : `<div class="hint" style="margin-top:6px">Wpisz opis, aby utworzyć <b>nowe zadanie</b>, albo wybierz z podpowiedzi <b>istniejące</b>.</div>`}
      <div class="row" style="margin-top:12px;align-items:flex-end"><label class="f" style="min-width:280px">Projekt (opcjonalnie)<select class="input" id="trkProj" ${sel ? 'disabled' : ''}>${projOptions(sel ? (projOf(sel) || sel.projectId) : (S.trk?.projectId || ''))}</select></label>
        <div class="spacer"></div><button class="btn big primary" data-action="trk-start">⟨i:play⟩ Start</button></div>
      ${paused.length ? `<div class="group-h" style="margin-top:16px">Wstrzymane – wznów</div><div class="tplchips">${paused.map(x => `<button class="tplchip" data-action="timer-start-task" data-id="${x.id}">⟨i:play⟩${esc(x.title)}<span>${fmtH(taskSpent(x.id))}</span></button>`).join('')}</div>` : ''}
      ${recent.length ? `<div class="group-h" style="margin-top:12px">Ostatnio realizowane</div><div class="tplchips">${recent.map(x => `<button class="tplchip" data-action="timer-start-task" data-id="${x.id}">⟨i:play⟩${esc(x.title)}</button>`).join('')}</div>` : ''}
      <div class="row" style="margin-top:14px"><button class="btn ghost" data-action="entry-new">⟨i:plus⟩ Dodaj wpis ręcznie (post factum)</button></div></div>`;
  }
  const es = entriesWithLive().filter(e => e.start.slice(0, 10) === td).sort((a, b) => a.start.localeCompare(b.start));
  return trackTabs() + `<div class="grid track-grid">${top}
    <div class="card"><h3>⟨i:timer⟩ Dziś · ${fmtH(tot.total)}<span class="right"><button class="btn sm" data-action="entry-new">⟨i:plus⟩ Wpis ręczny</button></span></h3>${realChart(td)}
      <div class="entlist">${es.map(entRow).join('') || '<div class="empty">Brak wpisów dziś</div>'}</div></div></div>`;
}
function entRow(e) {
  const pj = projById(e.projectId);
  return `<div class="entrow ${e.live ? 'live' : ''}" ${e.live ? '' : `data-action="entry-edit" data-id="${e.id}"`}><span class="mcol" style="background:${entColor(e)}"></span>
    <span class="er-t">${hm(e.start)}–${e.live ? 'teraz' : hm(e.end)}</span><b class="er-d">${fmtH(entMin(e))}</b>
    <span class="er-n">${esc(e.title)}${e.note ? `<span class="muted"> · ${esc(e.note.slice(0, 60))}</span>` : ''}</span>${pj ? projTag(pj.id) : ''}
    ${e.live ? '<span class="tag src"><span class="tc-dot"></span> trwa</span>' : `<span class="muted er-k" title="${e.manual ? 'wpis ręczny' : 'stoper'}">${e.manual ? '⟨i:pencil⟩' : e.kind === 'done' ? '⟨i:check⟩' : '⟨i:pause⟩'}</span>`}</div>`;
}
/* zakres okresu dla ewidencji i raportu */
function periodRange(p, from, to) {
  const t = today(), w = sow(t);
  switch (p) {
    case 'today': return [t, t];
    case 'yesterday': return [add(t, -1), add(t, -1)];
    case 'week': return [w, add(w, 6)];
    case 'lastweek': return [add(w, -7), add(w, -1)];
    case 'month': return [new Date(t.getFullYear(), t.getMonth(), 1), new Date(t.getFullYear(), t.getMonth() + 1, 0)];
    case 'lastmonth': return [new Date(t.getFullYear(), t.getMonth() - 1, 1), new Date(t.getFullYear(), t.getMonth(), 0)];
    case 'range': return [from ? P(from) : add(t, -30), to ? P(to) : t];
    default: return [w, add(w, 6)];
  }
}
const PERIODS = [['today', 'Dziś'], ['yesterday', 'Wczoraj'], ['week', 'Ten tydzień'], ['lastweek', 'Poprzedni tydzień'], ['month', 'Ten miesiąc'], ['lastmonth', 'Poprzedni miesiąc'], ['range', 'Zakres dat…']];
function periodBar(key) {
  const F = S[key] ||= {period: 'week', proj: '', q: ''};
  return `<select class="input" style="max-width:200px" data-bind="${key}.period">${PERIODS.map(([k, l]) => `<option value="${k}" ${F.period === k ? 'selected' : ''}>${l}</option>`).join('')}</select>
    ${F.period === 'range' ? `<input class="input" type="date" style="max-width:160px" data-bind="${key}.from" value="${F.from || ''}"><input class="input" type="date" style="max-width:160px" data-bind="${key}.to" value="${F.to || ''}">` : ''}
    <select class="input" style="max-width:240px" data-bind="${key}.proj"><option value="">Wszystkie projekty</option><option value="-" ${F.proj === '-' ? 'selected' : ''}>Bez projektu</option>${projectsC().map(p => `<option value="${p.id}" ${F.proj === p.id ? 'selected' : ''}>${esc(p.title)}</option>`).join('')}</select>`;
}
function periodEntries(key) {
  const F = S[key] ||= {period: 'week', proj: '', q: ''}, [a, b] = periodRange(F.period, F.from, F.to), A = iso(a), B = iso(b), q = norm(F.q || '');
  let es = entriesWithLive().filter(e => e.start.slice(0, 10) >= A && e.start.slice(0, 10) <= B);
  if (F.proj) es = es.filter(e => F.proj === '-' ? !projById(e.projectId) : e.projectId === F.proj);
  if (q) es = es.filter(e => norm([e.title, e.note, projById(e.projectId)?.title].join(' ')).includes(q));
  return {es, a, b};
}
function vTimeLog() {
  const {es, a, b} = periodEntries('tl'), by = {};
  es.sort((x, y) => y.start.localeCompare(x.start)).forEach(e => (by[e.start.slice(0, 10)] ||= []).push(e));
  const tot = es.reduce((s, e) => s + entMin(e), 0);
  return trackTabs() + `<div class="row" style="margin-bottom:12px">${periodBar('tl')}<input class="input" style="max-width:240px" placeholder="Szukaj…" data-bind="tl.q" value="${esc(S.tl.q || '')}">
    <div class="spacer"></div><button class="btn primary" data-action="entry-new">⟨i:plus⟩ Dodaj wpis ręcznie</button></div>
    <div class="muted" style="margin-bottom:10px;font-size:13px">${fmt(iso(a))} – ${fmt(iso(b))} · <b>${es.length}</b> wpisów · razem <b>${fmtH(tot)}</b></div>
    ${Object.entries(by).map(([d, l]) => `<div class="card" style="margin-bottom:12px;padding:10px 14px"><div class="row" style="margin-bottom:6px"><b class="crumb" data-action="show-day" data-d="${d}">⟨i:calendar⟩ ${DOWL[P(d).getDay()]}, ${fmt(d)}</b><div class="spacer"></div><b>${fmtH(l.reduce((s, e) => s + entMin(e), 0))}</b></div>
      ${l.sort((x, y) => x.start.localeCompare(y.start)).map(entRow).join('')}</div>`).join('') || '<div class="card empty">Brak wpisów w tym okresie. Uruchom stoper albo dodaj wpis ręcznie.</div>'}`;
}
function vTimeReport() {
  const F = S.tr ||= {period: 'week', proj: '', q: '', by: 'project'};
  const {es, a, b} = periodEntries('tr'), tot = es.reduce((s, e) => s + entMin(e), 0);
  const days = []; for (let d = new Date(a); d <= b && days.length < 62; d = add(d, 1)) days.push(iso(d));
  const keyOf = e => F.by === 'task' ? (e.taskId || e.title) : (e.projectId || '-');
  const nameOf = k => F.by === 'task' ? (S.data.tasks.find(t => t.id === k)?.title || k) : (projById(k)?.title || 'Bez projektu');
  const colOf = k => F.by === 'task' ? entColor(es.find(e => keyOf(e) === k) || {}) : (projColor(projById(k)) || '#64748B');
  const rows = {}; es.forEach(e => { const k = keyOf(e), d = e.start.slice(0, 10); rows[k] ||= {min: 0, n: 0, days: {}, tasks: new Set()}; rows[k].min += entMin(e); rows[k].n++; rows[k].days[d] = (rows[k].days[d] || 0) + entMin(e); rows[k].tasks.add(e.taskId || e.title); });
  const list = Object.entries(rows).sort((x, y) => y[1].min - x[1].min), dayTot = d => es.filter(e => e.start.slice(0, 10) === d).reduce((s, e) => s + entMin(e), 0);
  const workDays = days.filter(d => dayTot(d) > 0).length, showMatrix = days.length <= 31;
  // wykres: słupki dzienne skumulowane wg grup
  const W = 720, Ht = 200, pl = 34, pb = 26, mx = Math.max(60, ...days.map(dayTot)), mxH = Math.ceil(mx / 60), bw = (W - pl - 8) / Math.max(1, days.length), y = v => 8 + (Ht - 8 - pb) * (1 - v / (mxH * 60));
  let ch = `<svg viewBox="0 0 ${W} ${Ht}" width="100%" class="realchart">`;
  for (let h = 0; h <= mxH; h += Math.max(1, Math.ceil(mxH / 5))) ch += `<line x1="${pl}" x2="${W - 8}" y1="${y(h * 60)}" y2="${y(h * 60)}" stroke="var(--line)"/><text x="${pl - 6}" y="${y(h * 60) + 4}" text-anchor="end">${h}h</text>`;
  days.forEach((d, i) => { let acc = 0; const x = pl + i * bw + bw * .15, w = bw * .7;
    list.forEach(([k, r]) => { const v = r.days[d] || 0; if (!v) return; ch += `<rect x="${x}" y="${y(acc + v)}" width="${w}" height="${Math.max(0, y(acc) - y(acc + v) - 1)}" rx="2" fill="${colOf(k)}"><title>${fmt(d)} · ${esc(nameOf(k))}: ${fmtH(v)}</title></rect>`; acc += v; });
    if (days.length <= 14 || i % Math.ceil(days.length / 14) === 0) ch += `<text x="${x + w / 2}" y="${Ht - 8}" text-anchor="middle">${P(d).getDate()}.${pad(P(d).getMonth() + 1)}</text>`; });
  ch += '</svg>';
  return trackTabs() + `<div class="row" style="margin-bottom:12px">${periodBar('tr')}
    <div class="seg"><button data-action="set" data-k="tr.by" data-v="project" class="${F.by === 'project' ? 'on' : ''}">Wg projektów</button><button data-action="set" data-k="tr.by" data-v="task" class="${F.by === 'task' ? 'on' : ''}">Wg zadań</button></div>
    <div class="spacer"></div><button class="btn" data-action="tr-copy" title="Kopiuj tabelę do schowka (wklej do Excela / maila)">⟨i:list⟩ Kopiuj tabelę</button><button class="btn" data-action="tr-csv">⟨i:download⟩ Eksport CSV</button></div>
   <div class="grid g4 kpis" style="margin-bottom:16px">
     <div class="card kpi"><div class="v">${fmtH(tot)}</div><div class="l">Czas realizacji · ${fmt(iso(a))} – ${fmt(iso(b))}</div></div>
     <div class="card kpi"><div class="v">${workDays}</div><div class="l">Dni z realizacją</div></div>
     <div class="card kpi"><div class="v">${workDays ? fmtH(tot / workDays) : '—'}</div><div class="l">Średnio dziennie</div></div>
     <div class="card kpi"><div class="v" style="font-size:20px">${list[0] ? esc(nameOf(list[0][0])) : '—'}</div><div class="l">Najwięcej czasu${list[0] ? ' · ' + fmtH(list[0][1].min) : ''}</div></div></div>
   <div class="card" style="margin-bottom:16px"><h3>⟨i:chart⟩ Czas dziennie</h3>${ch}<div class="legend">${list.slice(0, 10).map(([k]) => `<span><i style="background:${colOf(k)}"></i>${esc(nameOf(k))}</span>`).join('')}</div></div>
   <div class="card" style="padding:0;overflow:auto"><table class="tbl ptbl" id="trTable"><thead><tr><th>${F.by === 'task' ? 'Zadanie' : 'Projekt'}</th><th class="num">Czas</th><th class="num">Udział</th><th class="num">Wpisy</th>${F.by === 'project' ? '<th class="num">Zadania</th>' : ''}
      ${showMatrix ? days.map(d => `<th class="num dcol ${dayType(P(d)) !== 'work' ? 'we' : ''}">${DOW[(P(d).getDay() + 6) % 7]}<br>${P(d).getDate()}.${pad(P(d).getMonth() + 1)}</th>`).join('') : ''}</tr></thead><tbody>
    ${list.map(([k, r]) => `<tr><td><span class="pcol" style="background:${colOf(k)}"></span>${esc(nameOf(k))}</td><td class="num"><b>${fmtH(r.min)}</b></td><td class="num">${tot ? Math.round(100 * r.min / tot) : 0}%</td><td class="num">${r.n}</td>${F.by === 'project' ? `<td class="num">${r.tasks.size}</td>` : ''}
      ${showMatrix ? days.map(d => `<td class="num dcol ${dayType(P(d)) !== 'work' ? 'we' : ''}">${r.days[d] ? (r.days[d] / 60).toFixed(1) : ''}</td>`).join('') : ''}</tr>`).join('') || `<tr><td colspan="${5 + (showMatrix ? days.length : 0)}" class="empty">Brak realizacji w tym okresie</td></tr>`}</tbody>
    ${list.length ? `<tfoot><tr><td><b>Razem</b></td><td class="num"><b>${fmtH(tot)}</b></td><td class="num">100%</td><td class="num">${es.length}</td>${F.by === 'project' ? '<td></td>' : ''}${showMatrix ? days.map(d => `<td class="num dcol">${dayTot(d) ? (dayTot(d) / 60).toFixed(1) : ''}</td>`).join('') : ''}</tr></tfoot>` : ''}</table></div>
   <div class="legend"><span>Wartości w kolumnach dni – godziny (h)</span><span>· raport uwzględnia przełącznik ALL/PRIV/JOB</span></div>`;
}
function reportRowsTSV() {
  const {es} = periodEntries('tr');
  return [['Data', 'Od', 'Do', 'Czas (h)', 'Projekt', 'Zadanie', 'IC case', 'Kontekst', 'Notatka', 'Typ'], ...es.sort((a, b) => a.start.localeCompare(b.start)).map(e => [e.start.slice(0, 10), hm(e.start), e.live ? '' : hm(e.end), (entMin(e) / 60).toFixed(2).replace('.', ','), projById(e.projectId)?.title || '', e.title, (e.taskId && S.data.tasks.find(t => t.id === e.taskId)?.ic) || '', (e.ctx || '').toUpperCase(), (e.note || '').replace(/\s+/g, ' '), e.live ? 'w toku' : e.manual ? 'ręczny' : 'stoper'])];
}
/* wpis ręczny / edycja wpisu */
function editEntry(e, preset = {}) {
  const isNew = !e;
  if (isNew) { const td = preset.date || TODAY(), last = entriesAll().filter(x => x.start.slice(0, 10) === td).sort((a, b) => b.end.localeCompare(a.end))[0];
    const from = last ? hm(last.end) : '08:00', nowS = `${pad(new Date().getHours())}:${pad(Math.floor(new Date().getMinutes() / 5) * 5)}`;
    e = {date: td, from: preset.from || from, to: preset.to || (td === TODAY() && nowS > from ? nowS : hm(addMin(`${td}T${from}`, 60))), taskId: preset.taskId || '', title: '', projectId: preset.projectId || '', note: ''}; }
  else e = Object.assign({}, e, {date: e.start.slice(0, 10), from: hm(e.start), to: hm(e.end)});
  const ts = S.data.tasks.filter(t => inCtx(t) || t.id === e.taskId).sort((a, b) => (a.status === 'done') - (b.status === 'done') || (b.due || '').localeCompare(a.due || ''));
  modal(`<h3>⟨i:timer⟩ ${isNew ? 'Wpis realizacji (post factum)' : 'Wpis realizacji'}</h3><form data-form="entry" data-id="${isNew ? '' : e.id}" class="grid g2">
    <label class="f">Data<input class="input" type="date" name="date" required value="${e.date}"></label>
    <div class="row"><label class="f">Od<input class="input" type="time" name="from" required value="${e.from}"></label><label class="f">Do<input class="input" type="time" name="to" required value="${e.to}"></label><span class="muted" id="entDur" style="padding-top:18px"></span></div>
    <label class="f" style="grid-column:1/-1">Zadanie<select class="input" name="taskId"><option value="">— nowe zadanie (wpisz opis poniżej) —</option>${ts.map(t => `<option value="${t.id}" data-proj="${projOf(t) || t.projectId || ''}" data-ctx="${t.ctx}" data-goal="${t.goalId || ''}" ${e.taskId === t.id ? 'selected' : ''}>${esc(t.title)}${t.status === 'done' ? ' (zrobione)' : ''}</option>`).join('')}</select></label>
    <label class="f entnew" style="grid-column:1/-1" ${e.taskId ? 'hidden' : ''}>Opis (utworzy nowe zadanie)<input class="input" name="title" value="${esc(e.taskId ? '' : e.title || '')}" placeholder="np. Analiza zgłoszenia klienta"></label>
    <label class="f">Projekt<select class="input" name="projectId">${projOptions(e.projectId)}</select></label>
    <label class="f" style="align-self:end"><span class="row" style="gap:6px;font-weight:500"><input type="checkbox" name="markDone"> Oznacz zadanie jako zrobione</span></label>
    <label class="f" style="grid-column:1/-1">Notatka<textarea class="input" name="note" rows="2">${esc(e.note || '')}</textarea></label>
    <div class="row" style="grid-column:1/-1;justify-content:space-between">${isNew ? '<span></span>' : `<button type="button" class="btn danger" data-action="entry-del" data-id="${e.id}">⟨i:trash⟩ Usuń wpis</button>`}
    <div class="row"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">Zapisz wpis</button></div></div></form>`, m => {
    const f = $('form[data-form=entry]', m), dur = () => { const a = f.from.value, b = f.to.value; if (a && b) { const d = (+b.slice(0, 2) * 60 + +b.slice(3)) - (+a.slice(0, 2) * 60 + +a.slice(3)); $('#entDur', m).textContent = d > 0 ? '= ' + fmtH(d) : 'koniec przed startem'; } };
    f.from.addEventListener('input', dur); f.to.addEventListener('input', dur); dur();
    f.taskId.addEventListener('change', () => { const o = f.taskId.selectedOptions[0]; $('.entnew', m).hidden = !!f.taskId.value; if (o && o.dataset.proj) f.projectId.value = o.dataset.proj; });
    linkForm(m); // projekt zawęża listę zadań
    (isNew && !e.taskId ? f.title : f.note).focus();
  });
}
function deleteEntry(id) {
  const i = entriesAll().findIndex(x => x.id === id); if (i < 0) return; const e = entriesAll()[i];
  entriesAll().splice(i, 1); save(); closeModal(); render();
  toastAction(`Usunięto wpis ${hm(e.start)}–${hm(e.end)} (${fmtH(entMin(e))})`, '⟨i:undo⟩ Cofnij', () => { entriesAll().splice(i, 0, e); save(); render(); });
}
afterRender.track = () => {
  const n = $('#tmNote'); if (n) n.addEventListener('input', () => { if (timerOn()) { S.data.timer.note = n.value; save(); } });
  const q = $('#trkQ'); if (!q) return;
  const sug = () => {
    const v = norm(q.value.trim()); S.trk ||= {}; S.trk.q = q.value;
    if (S.trk.taskId && S.data.tasks.find(x => x.id === S.trk.taskId)?.title !== q.value) S.trk.taskId = null;
    const l = S.data.tasks.filter(t => inCtx(t) && t.status !== 'done' && (!v || norm(t.title).includes(v))).sort(prioSort).slice(0, v ? 8 : 6);
    $('#trkSug').innerHTML = H(l.length && !S.trk.taskId ? `<div class="muted ts-h">${v ? 'Istniejące zadania' : 'Otwarte zadania'} – kliknij, aby wybrać${v ? ' · Enter = nowe zadanie „' + esc(q.value.trim()) + '”' : ''}</div>` + l.map(t => `<div class="ts-item" data-action="trk-pick" data-id="${t.id}"><span class="dot ${t.ctx}"></span><span class="ts-t">${esc(t.title)}</span>${projTag(projOf(t))}<span class="muted">${t.due ? relDay(t.due) : ''}${taskSpent(t.id) ? ' · ' + fmtH(taskSpent(t.id)) : ''}</span></div>`).join('') : '');
  };
  q.addEventListener('input', sug);
  q.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); A['trk-start'](); } });
  const ps = $('#trkProj'); if (ps) ps.addEventListener('change', () => { S.trk ||= {}; S.trk.projectId = ps.value; });
  sug(); if (!S.trk?.taskId) { q.focus(); q.setSelectionRange(q.value.length, q.value.length); }
};

/* ================= PROFIL UŻYTKOWNIKA (avatar w prawym górnym rogu) ================= */
function vProfile() {
  const T = S.profTab || 'account', u = S.user;
  const tabs = [['account', '⟨i:user⟩ Konto'], ['data', '⟨i:download⟩ Dane']];
  const head = `<div class="card prof-head"><div class="avatar big">${u.picture ? `<img src="${esc(u.picture)}" alt="">` : esc(u.name[0].toUpperCase())}</div>
    <div style="flex:1;min-width:0"><div class="gt" style="font-size:18px">${esc(u.name)}</div><div class="muted">${esc(u.email)} · ${u.provider === 'ms' ? 'konto Microsoft ' + esc(u.ms || '') + ' · synchronizacja przez OneDrive' : u.provider === 'google' ? 'konto Google' : 'laptop – dane w pliku w OneDrive'}</div></div>
    <button class="btn" data-action="logout">Wyloguj</button></div>
    <div class="seg" style="margin:16px 0 14px">${tabs.map(([k, l]) => `<button data-action="set" data-k="profTab" data-v="${k}" class="${T === k ? 'on' : ''}">${l}</button>`).join('')}</div>`;
  if (T !== 'account' && T !== 'data') S.profTab = 'account';
  if (T === 'account') return head + `<div class="card" style="max-width:640px"><h3>Dane konta</h3>
    <form data-form="profile" class="grid" style="gap:12px"><label class="f">Imię i nazwisko<input class="input" name="name" value="${esc(u.name)}" required></label>
      <label class="f">Adres e-mail (login)<input class="input" value="${esc(u.email)}" readonly></label>
      <div class="row"><button class="btn primary">Zapisz</button></div></form>
    <p class="hint" style="margin:14px 0 0">Dane TM są zapisywane osobno dla każdego konta na tym urządzeniu. Laptop i telefon pracują na tym samym pliku w OneDrive – zmiany scalają się automatycznie (Profil → Dane → Telefon).</p></div>`;
  if (T === 'data') return head + `<div class="grid prof-cal" style="margin-bottom:16px"><div class="card"><h3>⟨i:folder⟩ Gdzie są Twoje dane</h3>
    <p class="muted" style="margin-top:0">Dane TM zapisują się automatycznie do pliku w folderze aplikacji (<b>Dane</b>) – w OneDrive, z historią wersji. Przeglądarka trzyma kopię roboczą (działa także bez serwera).</p>
    <div id="diskStatus">${diskStatusHtml()}</div>
    <div class="row" style="margin-top:12px"><button class="btn sm" data-action="disk-reveal">⟨i:external⟩ Pokaż plik w Eksploratorze</button><button class="btn sm" data-action="disk-save">⟨i:upload⟩ Zapisz teraz</button></div></div>
    <div class="card"><h3>⟨i:rotate⟩ Kopie zapasowe (folder Kopie)</h3><p class="muted" style="margin-top:0">Codziennie przy pierwszej zmianie zapisywana jest kopia stanu z początku dnia (ostatnie 60). Przed przywróceniem tworzona jest dodatkowa kopia bieżących danych.</p>
    <div id="diskBackups"><p class="muted" style="margin:0">Wczytuję…</p></div></div>${CLOUD.on ? '' : phoneCard()}</div>
    <div class="card" style="max-width:640px"><h3>Kopia i przenoszenie danych</h3>
    <p class="muted" style="margin-top:0">Eksport zapisuje projekty, podprojekty, zadania, spotkania, urlopy i ustawienia importu do pliku JSON. Import zastępuje bieżące dane.</p>
    <div class="row"><button class="btn" data-action="export">⟨i:download⟩ Eksport JSON</button><label class="btn">⟨i:upload⟩ Import JSON<input type="file" accept=".json" id="importFile" hidden></label>
    <label class="row" style="gap:6px;font-size:13px"><input type="checkbox" data-action="demo-toggle" ${S.data.settings && S.data.settings.demoOn ? 'checked' : ''}> pokazuj przykładowe spotkania (demo)</label>
    <div class="spacer"></div><button class="btn danger" data-action="reset">⟨i:rotate⟩ Przywróć dane przykładowe</button></div></div>`;
  return head + `<div class="card">Wybierz zakładkę.</div>`;
}
/* Outlook: link ICS, harmonogram, zakres (Ustawienia → Integracje → Outlook) */
function vOutlookCfg() {
  const c = impCfg(), log = c.log || [], next = nextImport(c), n = S.data.meetings.filter(x => x.imported === 'outlook').length;
  const m = (k, l) => `<label class="radio-card ${c.method === k ? 'on' : ''}"><input type="radio" name="method" value="${k}" ${c.method === k ? 'checked' : ''}> ${l}</label>`;
  return `<div class="grid prof-cal">
   <div class="card"><div class="row" style="margin-bottom:12px"><h3 style="margin:0">${srcBadge('outlook')} Outlook / Microsoft 365</h3><span class="tag ${impReady(c) ? 'ok-tag' : 'src'}">${impReady(c) ? 'skonfigurowany' : 'nieskonfigurowany'}</span></div>
    <form data-form="import" class="grid" style="gap:14px">
      <div class="row" style="gap:8px">${m('ics', '⟨i:external⟩ Link ICS (opublikowany kalendarz)')}${m('graph', '⟨i:building⟩ Microsoft 365 (Graph API)')}</div>
      <div class="impm" data-m="ics" ${c.method !== 'ics' ? 'hidden' : ''}>
        <label class="f">Link ICS do kalendarza Outlook<input class="input" name="icsUrl" value="${esc(c.icsUrl || '')}" placeholder="https://outlook.office365.com/owa/calendar/…/calendar.ics" autocomplete="off" spellcheck="false"></label>
        <details class="hint" style="margin-top:6px"><summary style="cursor:pointer">Skąd wziąć link?</summary>
          <ol style="margin:6px 0 0 18px;padding:0"><li>Outlook → Ustawienia → Kalendarz → <b>Kalendarze udostępnione</b> (lub <a href="https://outlook.office.com/calendar/options/calendar/SharedCalendars" target="_blank" rel="noopener">w przeglądarce ↗</a>).</li>
          <li>Sekcja <b>Publikuj kalendarz</b>: wybierz kalendarz i „Może wyświetlać wszystkie szczegóły”, kliknij <b>Publikuj</b>.</li><li>Przy linku <b>ICS</b> wybierz „Kopiuj link” i wklej tutaj.</li></ol>
          <p style="margin:6px 0 0">Link jest prywatny – nie udostępniaj go. Aby go unieważnić: „Cofnij publikowanie” w Outlooku.</p></details></div>
      <div class="impm" data-m="graph" ${c.method !== 'graph' ? 'hidden' : ''}><p class="hint" style="margin:0">${CFG.MS_CLIENT_ID ? 'Aplikacja skonfigurowana (<code>MS_CLIENT_ID</code>). Przy pierwszym imporcie zalogujesz się kontem Microsoft 365.' : 'Wymaga rejestracji aplikacji w Microsoft Entra ID (<code>Calendars.Read</code>) i wpisania <code>MS_CLIENT_ID</code> w <code>config.js</code>.'}</p></div>
      <div class="sched"><h4>⟨i:clock⟩ Harmonogram importu</h4>
        <div class="row" style="gap:12px;align-items:flex-end">
          <label class="row" style="gap:8px;cursor:pointer;padding-bottom:9px"><input type="checkbox" name="auto" ${c.auto ? 'checked' : ''}> <b>Import automatyczny</b></label>
          <label class="f">Co<select class="input" name="every">${[[15, '15 minut'], [30, '30 minut'], [60, '1 godzinę'], [120, '2 godziny'], [240, '4 godziny'], [480, '8 godzin']].map(([v, l]) => `<option value="${v}" ${c.every === v ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
          <label class="f">Zakres wstecz<select class="input" name="back">${[0, 7, 14, 30, 90].map(v => `<option value="${v}" ${c.back === v ? 'selected' : ''}>${v} dni</option>`).join('')}</select></label>
          <label class="f">Zakres do przodu<select class="input" name="fwd">${[[14, '14 dni'], [30, '1 miesiąc'], [62, '2 miesiące'], [92, '3 miesiące'], [183, '6 miesięcy']].map(([v, l]) => `<option value="${v}" ${c.fwd === v ? 'selected' : ''}>${l}</option>`).join('')}</select></label></div>
        <p class="hint" style="margin:8px 0 0">Harmonogram działa, gdy TM jest otwarty (także zminimalizowany). Po uruchomieniu TM import wykona się od razu, jeśli minął ustawiony czas.</p></div>
      <div class="row" style="justify-content:space-between"><label class="btn" title="Jednorazowy import pliku .ics (np. Outlook klasyczny: Plik → Zapisz kalendarz)">⟨i:upload⟩ Wczytaj plik .ics<input type="file" accept=".ics,text/calendar" id="icsFile" hidden></label>
        <div class="row"><button class="btn" name="save" value="1">Zapisz ustawienia</button><button class="btn primary" name="run" value="1" ${importing ? 'disabled' : ''}>⟨i:refresh⟩ ${importing ? 'Importuję…' : 'Zapisz i importuj teraz'}</button></div></div>
    </form></div>
   <div class="grid" style="align-content:start">
    <div class="card"><h3>⟨i:refresh⟩ Status</h3><div class="kv">
      <span>Ostatni udany import</span><b>${ago(c.lastOk)}</b>
      <span>Następny import</span><b>${importing ? 'trwa…' : next ? `${pad(next.getHours())}:${pad(next.getMinutes())}` : c.auto ? 'po skonfigurowaniu' : 'wyłączony'}</b>
      <span>Spotkania z Outlooka w TM</span><b>${n}</b>
      <span>Zakres</span><b>${c.back} dni wstecz · ${c.fwd === 62 ? '2 miesiące' : c.fwd + ' dni'} do przodu</b></div></div>
    <div class="card"><h3>Historia importów</h3>${log.length ? `<table class="tbl implog"><tbody>${log.map(l => `<tr><td class="nowrap">${new Date(l.t).toLocaleString('pl-PL', {dateStyle: 'short', timeStyle: 'short'})}<div class="muted">${esc(l.via)}</div></td><td>${l.ok ? `<span class="p ok">●</span> ${l.total} spotkań · +${l.added} / zm. ${l.changed} / −${l.removed}${l.allDay ? `<div class="muted">pominięto całodniowe: ${l.allDay}</div>` : ''}` : `<span class="p bad">●</span> ${esc(l.err)}`}</td></tr>`).join('')}</tbody></table>` : '<div class="empty">Jeszcze nie importowano</div>'}</div>
    <div class="card srcopt off"><div>${srcBadge('google')} <b>Google Calendar</b><div class="muted" style="font-size:12px">Import prywatnego kalendarza Google – następny etap.</div></div><span class="tag src" style="margin-left:auto">wkrótce</span></div>
   </div></div>`;
}
function bindImportForm() {
  const f = $('form[data-form=import]');
  if (f) {
    $$('[name=method]', f).forEach(r => r.addEventListener('change', () => { $$('.impm', f).forEach(d => d.hidden = d.dataset.m !== r.value); $$('.radio-card', f).forEach(l => l.classList.toggle('on', l.querySelector('input').checked)); }));
    $$('button[name]', f).forEach(b => b.addEventListener('click', () => { f.dataset.submitter = b.name; }));
    $('#icsFile').onchange = e => { const x = e.target.files[0]; if (x) x.text().then(t => runImport({fileText: t})); };
  }
}
afterRender.profile = () => {
  bindImportForm();
  diskBackupsLoad();
  const jf = $('#importFile');
  if (jf) jf.onchange = e => { const x = e.target.files[0]; if (!x) return; x.text().then(t => { try { const d = JSON.parse(t); if (!d.tasks) throw 0; S.data = migrate(d); save(); render(); toast('Zaimportowano dane'); } catch { toast('Nieprawidłowy plik'); } }); };
};

/* ================= akcje (delegacja zdarzeń) ================= */
function setPath(path, v) { const ks = path.split('.'); let o = ks[0] === 'settings' ? S.data : S; for (let i = 0; i < ks.length - 1; i++) o = o[ks[i]]; o[ks.at(-1)] = v; }
const A = {
  nav: el => { closePop(); S.view = el.dataset.view; if (S.view === 'gantt') S.gantt.only = null; history.replaceState(null, '', '#' + S.view); render(); $('main').scrollTop = 0; },
  quick: () => quickAdd(),
  'trk-pick': el => { S.trk = {taskId: el.dataset.id}; render(); },
  'trk-clear': () => { S.trk = {}; render(); },
  'trk-start': () => startTimer({taskId: S.trk?.taskId, title: $('#trkQ')?.value, projectId: $('#trkProj')?.value || S.trk?.projectId}),
  'timer-pause': () => stopTimer('pause'),
  'timer-done': () => stopTimer('done'),
  'timer-switch': () => { stopTimer('switch', true); S.trk = {}; S.view = 'track'; history.replaceState(null, '', '#track'); render(); toast('Sesja zapisana – wybierz kolejne zadanie'); },
  'timer-start-task': (el, e) => { e.stopPropagation(); closeModal(); startTimer({taskId: el.dataset.id}); },
  'entry-new': (el, e) => { e.stopPropagation(); editEntry(null, {taskId: el.dataset.task || '', date: S.view === 'calendar' && S.cal.mode === 'day' ? iso(S.cal.date) : undefined}); },
  'entry-edit': (el, e) => { e.stopPropagation(); const x = entriesAll().find(y => y.id === el.dataset.id); if (x) editEntry(x); },
  'entry-del': el => deleteEntry(el.dataset.id),
  'tr-copy': async () => { const t = $('#trTable'); if (!t) return; const txt = [...t.rows].map(r => [...r.cells].map(c => c.innerText.replace(/\s+/g, ' ').trim()).join('\t')).join('\n');
    try { await navigator.clipboard.writeText(txt); toast('Skopiowano tabelę – wklej do Excela, maila lub prezentacji'); } catch { toast('Nie udało się skopiować do schowka'); } },
  'tr-csv': () => { const rows = reportRowsTSV(), csv = rows.map(r => r.map(v => /[;"\n]/.test(v) ? '"' + String(v).replace(/"/g, '""') + '"' : v).join(';')).join('\r\n');
    const [a, b] = periodRange(S.tr.period, S.tr.from, S.tr.to), el = document.createElement('a'); el.href = URL.createObjectURL(new Blob(['\ufeff' + csv], {type: 'text/csv;charset=utf-8'})); el.download = `realizacja_${iso(a)}_${iso(b)}.csv`; el.click(); },
  'qa-tpl': el => { S.qaTpl = el.dataset.id; const i = $('#qa'); if (i && i.value.trim().startsWith('/')) i.value = ''; qaRefresh(); i && i.focus(); },
  'qa-tpl-clear': () => { S.qaTpl = null; qaRefresh(); $('#qa')?.focus(); },
  'qa-cat': el => { S.qaCat = el.dataset.c; qaRefresh(); $('#qa')?.focus(); },
  'qa-manage': () => { closeModal(); S.view = 'settings'; S.setTab = 'tpl'; history.replaceState(null, '', '#settings'); render(); },
  'tpl-new': () => editTemplate(null),
  'tpl-edit': el => editTemplate(tplById(el.dataset.id)),
  'tpl-use': el => quickAdd(el.dataset.id),
  'tpl-dup': el => { const x = tplById(el.dataset.id); if (!x) return; const c = {...x, id: uid(), name: x.name + ' (kopia)', uses: 0, lastUsed: null}; tplAll().push(c); save(); render(); editTemplate(c); },
  'tpl-del': el => { const i = tplAll().findIndex(x => x.id === el.dataset.id); if (i < 0) return; const x = tplAll()[i]; tplAll().splice(i, 1); save(); closeModal(); render();
    toastAction(`Usunięto szablon „${x.name}”`, '⟨i:undo⟩ Cofnij', () => { tplAll().splice(i, 0, x); save(); render(); }); },
  'task-to-tpl': () => { const f = $('#modalRoot form[data-form=task]'); if (!f) return; const fd = Object.fromEntries(new FormData(f));
    editTemplate(null, {name: fd.title || 'Nowy szablon', title: fd.title || '', ctx: fd.ctx, projectId: fd.projectId || null, goalId: fd.goalId || null, prio: +fd.prio || 2, est: +fd.est || '', notes: fd.notes || ''}); }, 'qa-save': () => qaSave(false),
  'qa-more': () => { const r = quickResolve($('#qa').value); editTask(null, {task: {title: r.title, notes: r.notes || '', tplId: r.tpl?.id, ctx: r.ctx, goalId: r.goalId, projectId: r.projectId, start: r.start, due: r.due, prio: r.prio, est: r.est || 0}}); },
  toggle: el => { const t = S.data.tasks.find(x => x.id === el.dataset.id); setStatus(t, t.status === 'done' ? 'todo' : 'done'); },
  'edit-task': el => { if (!el.dataset.id) return editTask(null, {goal: el.dataset.goal, due: el.dataset.due, project: el.dataset.project}); editTask(S.data.tasks.find(x => x.id === el.dataset.id)); },
  'del-task': el => { closeModal(); deleteTasks([el.dataset.id]); },
  'del-task-row': (el, e) => { e.stopPropagation(); deleteTasks([el.dataset.id]); },
  'sel-mode': () => { S.tasks.selMode = !S.tasks.selMode; if (!S.tasks.selMode) S.tasks.sel.clear(); render(); },
  'sel-task': el => { const s = S.tasks.sel; el.checked ? s.add(el.dataset.id) : s.delete(el.dataset.id); render(); },
  'sel-all': () => { visibleTasks().forEach(t => S.tasks.sel.add(t.id)); render(); },
  'sel-none': () => { S.tasks.sel.clear(); render(); },
  'sel-group': el => { const ids = el.dataset.ids.split(','), s = S.tasks.sel, all = ids.every(id => s.has(id)); ids.forEach(id => all ? s.delete(id) : s.add(id)); render(); },
  'sel-done': () => { const s = S.tasks.sel; let n = 0; S.data.tasks.forEach(t => { if (s.has(t.id) && t.status !== 'done') { t.status = 'done'; t.doneAt = TODAY(); n++; } }); s.clear(); save(); render(); toast(`Oznaczono jako zrobione: ${n}`); },
  'sel-delete': () => { const ids = [...S.tasks.sel].filter(id => S.data.tasks.some(t => t.id === id)); if (!ids.length) return;
    const list = S.data.tasks.filter(t => ids.includes(t.id));
    confirmDlg('Usunąć zaznaczone zadania?', `<p>Zostanie usuniętych <b>${ids.length}</b> ${plural(ids.length)}${list.some(t => t.status !== 'done') ? `, w tym <b>${list.filter(t => t.status !== 'done').length}</b> niezakończonych` : ''}.</p><div class="hint">${list.slice(0, 8).map(t => '• ' + (S.privacy ? pv(esc(t.title)) : esc(t.title))).join('<br>')}${list.length > 8 ? `<br>… i ${list.length - 8} więcej` : ''}</div><p class="hint">Po usunięciu możesz jeszcze kliknąć „Cofnij”.</p>`,
      `⟨i:trash⟩ Usuń ${ids.length}`, () => { deleteTasks(ids); S.tasks.selMode = false; render(); }); },
  'del-done-ids': el => { const ids = el.dataset.ids.split(',').filter(Boolean); if (!ids.length) return;
    confirmDlg('Usunąć zrobione zadania?', `<p>Zostanie usuniętych <b>${ids.length}</b> zrobionych ${ids.length === 1 ? 'zadanie' : 'zadań'}. Po usunięciu możesz kliknąć „Cofnij”.</p>`, `⟨i:trash⟩ Usuń ${ids.length}`, () => deleteTasks(ids, `${ids.length} zrobionych zadań`)); },
  'del-done': () => { const F = S.tasks, keepStatus = F.status; F.status = 'all'; const done = visibleTasks().filter(t => t.status === 'done'); F.status = keepStatus; if (!done.length) return toast('Brak zrobionych zadań do usunięcia');
    const scope = [S.ctx !== 'all' ? S.ctx.toUpperCase() : 'ALL', F.proj && projById(F.proj) ? 'projekt: ' + projById(F.proj).title : '', F.goal && goalById(F.goal) ? 'podprojekt: ' + goalById(F.goal).title : '', F.q ? `szukaj: „${F.q}”` : ''].filter(Boolean).join(' · ');
    confirmDlg('Usunąć wszystkie zrobione zadania?', `<p>Zostanie usuniętych <b>${done.length}</b> zrobionych ${done.length === 1 ? 'zadanie' : 'zadań'}.</p><p class="hint">Zakres (bieżące filtry): ${esc(scope)}. Zadania otwarte nie zostaną usunięte. Po usunięciu możesz kliknąć „Cofnij”.</p>`,
      `⟨i:trash⟩ Usuń ${done.length}`, () => deleteTasks(done.map(t => t.id), `${done.length} zrobionych zadań`)); },
  'edit-goal': el => editGoal(el.dataset.id && goalById(el.dataset.id), {project: el.dataset.project}),
  'edit-project': el => editProject(el.dataset.id && projById(el.dataset.id)),
  'open-project': (el, e) => { if (e && e.target.closest('.acts') && !e.target.closest('[data-action="open-project"].btn')) return; S.proj.open = el.dataset.id; S.view = 'project'; history.replaceState(null, '', '#project'); render(); $('main').scrollTop = 0; },
  'close-project': () => { S.view = 'projects'; history.replaceState(null, '', '#projects'); render(); },
  'close-project-picker': () => { S.proj.open = null; render(); },
  psort: el => { const k = el.dataset.k; S.proj.dir = S.proj.sort === k ? -S.proj.dir : 1; S.proj.sort = k; render(); },
  psec: (el, e) => { e.preventDefault(); $('#' + el.dataset.t)?.scrollIntoView({behavior: 'smooth', block: 'start'}); },
  'proj-gantt': el => { const p = projById(el.dataset.id); S.gantt.group = 'project'; S.gantt.only = p.id; if (p.start) S.gantt.start = add(sow(P(p.start > TODAY() ? p.start : TODAY())), -7); S.view = 'gantt'; render(); },
  'del-project': el => { if (!confirm('Usunąć projekt? Podprojekty i zadania pozostaną, ale bez projektu.')) return; const id = el.dataset.id;
    S.data.projects = S.data.projects.filter(p => p.id !== id); S.data.goals.forEach(g => g.projectId === id && (g.projectId = null)); S.data.tasks.forEach(t => t.projectId === id && (t.projectId = null));
    S.proj.open = null; save(); closeModal(); render(); toast('Usunięto projekt'); },
  'del-goal': el => { if (!confirm('Usunąć podprojekt? Zadania pozostaną bez podprojektu.')) return; S.data.goals = S.data.goals.filter(g => g.id !== el.dataset.id); S.data.tasks.forEach(t => t.goalId === el.dataset.id && (t.goalId = null)); save(); closeModal(); render(); },
  set: el => { const v = el.dataset.v; setPath(el.dataset.k, /^\d+$/.test(v) ? +v : v); render(); },
  'cal-nav': el => { const n = +el.dataset.n, C = S.cal, d = C.date; C.date = C.mode === 'day' ? add(d, n) : C.mode === 'week' ? add(d, 7 * n) : C.mode === 'month' ? addM(d, n) : new Date(d.getFullYear() + n, d.getMonth(), 1); render(); },
  'cal-today': () => { S.cal.date = today(); render(); },
  'cal-goto': (el, e) => { if (e.target.closest('.chip[data-action="edit-task"]')) return; S.cal.date = P(el.dataset.d); S.cal.mode = el.dataset.mode || 'day'; render(); },
  'gantt-shift': el => { S.gantt.start = add(S.gantt.start, +el.dataset.n); render(); },
  'gantt-today': () => { S.gantt.start = add(sow(today()), -7); render(); },
  'vac-del': el => { S.data.vacations = S.data.vacations.filter(v => v.id !== el.dataset.id); save(); render(); },
  'connect-ms': async () => {
    if (!CFG.MS_CLIENT_ID) return infoModal('Outlook / Teams – konfiguracja', `Aby pobierać prawdziwe spotkania, dostępność i wysyłać pocztę z konta Microsoft 365:<ol><li>Entra ID → <b>App registrations</b> → New registration (typ <b>SPA</b>), Redirect URI = adres aplikacji (np. <code>http://localhost:8765/</code>).</li><li>API permissions (delegated): <code>User.Read</code>, <code>Calendars.ReadWrite</code>, <code>Mail.Send</code> → Grant admin consent (jeśli wymagane w InfoConsulting).</li><li>Wpisz <code>MS_CLIENT_ID</code> i <code>MS_TENANT</code> w <code>config.js</code>.</li></ol>Do tego czasu działa tryb demo: przykładowe spotkania, a tworzenie spotkań/maili otwiera Teams/Outlook z wypełnionymi polami.`);
    try { await MS.login(); S.conn.ms = true; toast('Połączono z Microsoft 365'); await syncExternal(); } catch (e) { toast('Logowanie Microsoft nieudane: ' + e.message); }
  },
  'connect-google': async () => {
    if (!CFG.GOOGLE_CLIENT_ID) return infoModal('Gmail / Google Calendar – konfiguracja', `<ol><li>Google Cloud Console → nowy projekt → włącz <b>Gmail API</b> i <b>Google Calendar API</b>.</li><li>OAuth consent screen (External/Testing, dodaj swój adres jako test user).</li><li>Credentials → OAuth client ID → <b>Web application</b>, Authorized JavaScript origins: <code>http://localhost:8765</code> + adres produkcyjny.</li><li>Wpisz <code>GOOGLE_CLIENT_ID</code> w <code>config.js</code> – włączy to też logowanie kontem Google.</li></ol>`);
    try { await GG.ensure(); S.conn.google = true; toast('Połączono z Google'); await syncExternal(); } catch (e) { toast('Google: autoryzacja nieudana'); }
  },
  'fb-person': el => { const I = S.integ; I.people = el.checked ? [...new Set([...I.people, el.value])] : I.people.filter(x => x !== el.value); I.fb = null; $('#fbGrid').innerHTML = H(fbGrid()); },
  'fb-load': async () => { const I = S.integ; if (S.conn.ms) { try { const rows = await MS.schedule([S.user.email, ...I.people], I.fbDate); rows[0].name = 'Ja'; I.fb = {date: I.fbDate, rows}; } catch (e) { toast(e.message); } } else toast('Tryb demo – dostępność przykładowa'); $('#fbGrid').innerHTML = H(fbGrid(I.fb)); },
  'fb-slot': el => { const I = S.integ; I.prefill = {date: I.fbDate, time: el.dataset.t, att: I.people, dur: 30, title: ''}; I.tab = 'new'; render(); },
  'new-mtg': (el, e) => { e.stopPropagation(); editMeeting(null, {date: el.dataset.d, project: el.dataset.project}); },
  'edit-mtg': (el, e) => { e.stopPropagation(); const m = S.data.meetings.find(x => x.id === el.dataset.id); if (m) editMeeting(m); },
  'del-mtg': (el, e) => { e.stopPropagation(); const m = S.data.meetings.find(x => x.id === el.dataset.id); if (!m) return;
    if (!canDelMtg(m)) return toast('Spotkania z Outlooka usuwasz w Outlooku');
    if (m.ext && (m.source === 'google' ? S.conn.google : S.conn.ms)) return confirmDlg('Odwołać spotkanie?', `<p>„${pv(esc(m.title))}” zostanie usunięte także w ${SRC[m.source][0]}${m.attendees?.length ? ', a uczestnicy dostaną informację o odwołaniu' : ''}.</p>`, '⟨i:trash⟩ Odwołaj', () => deleteMeeting(m.id));
    deleteMeeting(m.id); },
  privacy: () => { S.privacy = !S.privacy; const mm = $('#modalRoot form[data-form=mtg]'), id = mm?.dataset.id, tf = $('#modalRoot form[data-form=task]'), tid = tf?.dataset.id; render(); if (id) editMeeting(S.data.meetings.find(x => x.id === id)); if (tid) editTask(S.data.tasks.find(x => x.id === tid)); toast(S.privacy ? 'Dane spotkań i zadań ukryte' : 'Dane spotkań i zadań widoczne'); },
  'import-now': () => { if (!impReady()) { mailTmSync({silent: false}); S.view = 'settings'; S.setTab = 'integ'; history.replaceState(null, '', '#settile'); render(); return toast('Najpierw wklej link ICS Outlooka i zapisz ustawienia'); } runImport(); },
  'mail-open': el => { const t = S.data.tasks.find(x => x.id === el.dataset.id); if (t && t.mailId) outlookApp({action: 'mail-open', id: t.mailId}, null, 'Otworzono mail w Outlook'); },
  'wk-off': el => { const d = +el.dataset.d; S.wkOff = d === 0 ? 0 : (S.wkOff || 0) + d; render(); },
  'mail-now': () => mailTmSync({silent: false}).then(() => { if (S.view === 'settings') render(); }),
  'demo-toggle': el => { S.data.settings ||= {}; S.data.settings.demoOn = el.checked; S.data.demoDate = ''; refreshDemoMeetings(); save(); render(); toast(el.checked ? 'Włączono spotkania przykładowe' : 'Usunięto spotkania przykładowe'); },
  'cap-list': el => { const open = tasksC().filter(t => t.status !== 'done'), l = el.dataset.k === 'late' ? open.filter(isLate) : open.filter(t => !(t.est > 0));
    modal(`<h3>${el.dataset.k === 'late' ? 'Zadania zaległe' : 'Zadania bez estymacji'} (${l.length})</h3><p class="muted" style="margin-top:0">Kliknij zadanie, aby ${el.dataset.k === 'late' ? 'zmienić termin' : 'uzupełnić estymację'}.</p>${l.sort(prioSort).map(t => taskRow(t, {pv: true})).join('')}<div class="row" style="justify-content:flex-end;margin-top:10px"><button class="btn" data-action="modal-close">Zamknij</button></div>`); },
  'teams-open': el => teamsLaunch(el.dataset.url),
  'ol-open': el => { const m = S.data.meetings.find(x => x.id === el.dataset.id); if (m) outlookOpenMtg(m); },
  'disk-save': async () => { if (!DISK.on) return toast('Uruchom TM przez start-TM.bat'); DISK.ready = true; await diskWrite(); render(); toast(DISK.ok ? 'Zapisano dane do pliku w OneDrive' : 'Błąd: ' + DISK.err); },
  'disk-reveal': async () => { try { const r = await fetch(diskUrl('/__reveal-data'), {headers: diskHdr}); if (!r.ok) throw 0; } catch { toast('Uruchom ponownie start-TM.bat (nowa wersja serwera)'); } },
  'disk-restore': el => confirmDlg('Przywrócić kopię danych?', `<p>Bieżące dane zostaną zastąpione kopią <b>${esc(el.dataset.f)}</b>. Przed przywróceniem zapiszę dodatkową kopię obecnego stanu, więc możesz to odwrócić.</p>`, 'Przywróć', async () => {
    try {
      await fetch(diskUrl('/__snapshot'), {method: 'POST', headers: diskHdr});
      const r = await fetch(diskUrl('/__backup', '&f=' + encodeURIComponent(el.dataset.f)), {headers: diskHdr, cache: 'no-store'}); if (!r.ok) throw new Error(await r.text());
      const d = JSON.parse(await r.text()); if (!d.tasks) throw new Error('Plik nie zawiera danych TM');
      S.data = migrate(d); save(); render(); toast('Przywrócono dane z kopii');
    } catch (e) { toast('Nie udało się przywrócić: ' + (e.message || e)); }
  }),
  'open-profile': el => { const t = el.dataset.tab || 'account';
    if (['settings', 'calendars', 'notif', 'integ'].includes(t)) { S.view = 'settings'; S.setTab = t === 'calendars' || t === 'integ' ? 'integ' : t === 'notif' ? 'notif' : 'general'; history.replaceState(null, '', '#settings'); return render(); }
    S.view = 'profile'; S.profTab = t; history.replaceState(null, '', '#profile'); render(); },
  'import-open': () => { S.view = 'settings'; S.setTab = 'integ'; render(); },
  'cal-to-list': () => { S.mtgs.period = 'cal'; S.view = 'meetings'; history.replaceState(null, '', '#meetings'); render(); },
  'cal-nav-list': el => { A['cal-nav'](el); },
  'show-in-cal': el => { const m = S.data.meetings.find(x => x.id === el.dataset.id); if (!m) return; closeModal(); S.cal.date = P(m.start.slice(0, 10)); if (S.cal.mode === 'year' || S.cal.mode === 'month') S.cal.mode = 'week'; S.cal.flash = m.id; S.view = 'calendar'; history.replaceState(null, '', '#calendar'); render(); },
  'show-in-list': el => { const m = S.data.meetings.find(x => x.id === el.dataset.id); if (!m) return; closeModal(); S.cal.date = P(m.start.slice(0, 10)); S.mtgs.period = 'cal'; S.mtgs.q = ''; S.mtgs.lastOpened = m.id; S.view = 'meetings'; history.replaceState(null, '', '#meetings'); render(); },
  'show-day': el => { S.cal.date = P(el.dataset.d); S.cal.mode = 'day'; S.view = 'calendar'; history.replaceState(null, '', '#calendar'); render(); },
  msel: el => { el.checked ? S.mtgs.sel.add(el.dataset.id) : S.mtgs.sel.delete(el.dataset.id); render(); },
  'msel-all': el => { const ms = filteredMeetings(); if (el.checked) ms.forEach(m => S.mtgs.sel.add(m.id)); else ms.forEach(m => S.mtgs.sel.delete(m.id)); render(); },
  'msel-none': () => { S.mtgs.sel.clear(); render(); },
  msort: () => { S.mtgs.sort = S.mtgs.sort === 'asc' ? 'desc' : 'asc'; render(); },
  'msel-delete': () => { const ms = filteredMeetings().filter(m => S.mtgs.sel.has(m.id)), del = ms.filter(canDelMtg), lock = ms.length - del.length; if (!del.length) return;
    confirmDlg('Usunąć zaznaczone spotkania?', `<p>Zostanie usuniętych <b>${del.length}</b> spotkań.${lock ? ` <b>${lock}</b> spotkań z Outlooka zostanie pominiętych – usuwasz je w Outlooku.` : ''}</p><div class="hint">${del.slice(0, 8).map(m => `• ${fmt(m.start.slice(0, 10))} ${hm(m.start)} – ${pv(esc(m.title))}`).join('<br>')}${del.length > 8 ? `<br>… i ${del.length - 8} więcej` : ''}</div>${del.some(m => m.ext && m.source === 'google' && S.conn.google) ? '<p class="hint">Spotkania z Google Calendar zostaną odwołane także w Google.</p>' : ''}<p class="hint">Po usunięciu możesz kliknąć „Cofnij”.</p>`,
      `⟨i:trash⟩ Usuń ${del.length}`, () => deleteMeetings(del.map(m => m.id))); },
  'cal-slot': (el, e) => { const r = el.getBoundingClientRect(), y = e.clientY - r.top, mins = Math.floor((7 * 60 + y / 40 * 60) / 30) * 30; editMeeting(null, {date: el.dataset.d, time: `${pad(Math.floor(mins / 60))}:${pad(mins % 60)}`}); },
  'notes-of': (el, e) => { e.stopPropagation(); notesListModal(el.dataset.kind, el.dataset.id); },
  'note-new': (el, e) => { e.stopPropagation(); noteEditor(null, {kind: el.dataset.kind, id: el.dataset.id, ret: el.dataset.ret, project: el.dataset.project}); },
  'note-edit': (el, e) => { e.stopPropagation(); const n = notesAll().find(x => x.id === el.dataset.id); if (n) noteEditor(n, {ret: el.dataset.ret}); },
  'note-cancel': el => retTo(el.dataset.ret)(),
  'note-del': (el, e) => { e.stopPropagation(); deleteNote(el.dataset.id, el.dataset.ret || ($('#modalRoot form[data-form=note]')?.dataset.ret) || ''); },
  'note-dl': (el, e) => { e.stopPropagation(); const n = notesAll().find(x => x.id === el.dataset.id); if (n) downloadTxt(n); },
  'note-reveal': async (el, e) => { e.stopPropagation(); const n = notesAll().find(x => x.id === el.dataset.id); if (!n?.file) return;
    try { const r = await fetch('/__reveal?rel=' + encodeURIComponent(n.file)); if (!r.ok) throw 0; toast('Otworzono folder z notatką'); } catch { toast('Nie udało się otworzyć folderu – plik: ' + n.file.replace(/\//g, '\\')); } },
  'mtg-task': el => { const m = S.data.meetings.find(x => x.id === el.dataset.id); editTask(null, {task: {title: `Follow-up: ${m.title}`, ctx: m.ctx, start: m.start.slice(0, 10), due: iso(add(P(m.start.slice(0, 10)), 1)), est: 0.5}}); },
  'mail-tpl': el => {
    const f = $('form[data-form=mail]'), t = el.dataset.t, td = TODAY(), job = S.data.tasks.filter(x => x.ctx === 'job' && x.status !== 'done');
    const line = x => `• ${x.title} – termin ${fmt(x.due)}${x.status === 'doing' ? ' (w toku)' : ''}`;
    if (t === 'status') { f.subject.value = `Status zadań – ${fmt(td)}`; f.body.value = `Cześć,\n\nZrobione ostatnio:\n${S.data.tasks.filter(x => x.ctx === 'job' && x.doneAt && x.doneAt >= iso(add(today(), -7))).map(line).join('\n') || '• —'}\n\nW toku / na dziś:\n${job.filter(x => x.status === 'doing' || x.due <= td).map(line).join('\n') || '• —'}\n\nPozdrawiam,\n${S.user.name}`; }
    if (t === 'week') { const we = iso(add(sow(today()), 6)); f.subject.value = `Plan na tydzień ${fmt(iso(sow(today())))}–${fmt(we)}`; f.body.value = `Cześć,\n\nPlan na ten tydzień:\n${job.filter(x => x.due <= we).sort(prioSort).map(line).join('\n') || '• —'}\n\nPozdrawiam,\n${S.user.name}`; }
    if (t === 'late') { f.subject.value = 'Zaległe zadania – aktualizacja'; f.body.value = `Cześć,\n\nPoniższe zadania są po terminie:\n${job.filter(isLate).map(line).join('\n') || '• brak'}\n\nProponuję nowe terminy…\n\nPozdrawiam,\n${S.user.name}`; }
  },
  'modal-close': () => closeModal(), 'modal-bg': (el, e) => { if (e.target === el) closeModal(); },
  logout: () => logout(),
  export: () => { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([JSON.stringify(S.data, null, 2)], {type: 'application/json'})); a.download = `tm-${S.user.email}-${TODAY()}.json`; a.click(); },
  reset: () => confirmDlg('Przywrócić dane przykładowe?', '<p>Bieżące projekty, podprojekty, zadania i spotkania zostaną zastąpione danymi przykładowymi. Zrób wcześniej eksport JSON, jeśli chcesz je zachować.</p>', 'Przywróć', () => { const keep = S.data.integrations; S.data = seed(); S.data.integrations = keep; refreshDemoMeetings(); save(); render(); toast('Przywrócono dane przykładowe'); })
};
function infoModal(title, html) { modal(`<h3>${title}</h3><div class="hint" style="font-size:13px">${html}</div><div class="row" style="justify-content:flex-end;margin-top:12px"><button class="btn primary" data-action="modal-close">OK</button></div>`); }

const FORMS = {
  settings: (f, fd) => {
    const st = SET();
    [...f.elements].forEach(el => { if (!el.name) return; st[el.name] = el.type === 'checkbox' ? el.checked : el.type === 'number' ? (el.value === '' ? null : +el.value) : el.value.trim(); });
    applySettings(); save(); render(); toast('Zapisano ustawienia');
  },
  task: (f, fd) => {
    const id = f.dataset.id, t = id ? S.data.tasks.find(x => x.id === id) : {id: uid(), createdAt: TODAY(), doneAt: null};
    const was = t.status;
    const gp = fd.goalId && goalById(fd.goalId), pj = projById((gp && gp.projectId) || fd.projectId);
    Object.assign(t, {title: fd.title.trim(), ctx: pj ? pj.ctx : fd.ctx, goalId: fd.goalId || null, projectId: pj ? pj.id : null, start: fd.start || fd.due, due: fd.due || fd.start, status: fd.status, prio: +fd.prio, est: +fd.est || 0, notes: fd.notes});
    t.ic = t.ctx === 'job' ? icNorm(fd.ic) : ''; if (!t.ic) delete t.ic;
    ['assigneeId', 'requesterId'].forEach(k => { if (fd[k]) t[k] = fd[k]; else delete t[k]; });
    ['costEst', 'costAct'].forEach(k => { if (t.ctx === 'priv' && +fd[k] > 0) t[k] = +fd[k]; else delete t[k]; });
    if (t.start && t.due && t.start > t.due) t.start = t.due;
    if (t.status === 'done' && was !== 'done') t.doneAt = TODAY(); if (t.status !== 'done') t.doneAt = null;
    if (!id && (f.dataset.tpl || t.tplId)) { t.tplId = f.dataset.tpl || t.tplId; tplUsed(tplById(t.tplId)); }
    if (!id && f.dataset.trip && tripById(f.dataset.trip)) { const tr = tripById(f.dataset.trip); t.tripId = tr.id; t.tripAt = t.due > tr.to ? 'ret' : 'dep'; }
    if (!id) S.data.tasks.push(t); save(); closeModal(); render(); toast('Zapisano zadanie');
    if (!id && t.tripId) tripPanel(tripById(t.tripId));
  },
  goal: (f, fd) => {
    const id = f.dataset.id, g = id ? goalById(id) : {id: uid()};
    const pj = projById(fd.projectId);
    if (fd.start && fd.due && fd.due < fd.start) return toast('Termin podprojektu jest wcześniejszy niż jego start – popraw daty');
    Object.assign(g, {title: fd.title.trim(), ctx: pj ? pj.ctx : fd.ctx, projectId: pj ? pj.id : null, start: fd.start, due: fd.due, desc: fd.desc});
    if (g.ctx === 'priv' && +fd.budget > 0) g.budget = +fd.budget; else delete g.budget;
    if (g.ctx === 'job' && +fd.budgetH > 0) g.budgetH = +fd.budgetH; else delete g.budgetH;
    S.data.tasks.forEach(t => { if (t.goalId === g.id) { t.ctx = g.ctx; t.projectId = g.projectId; } }); if (fd.progress != null) g.progress = +fd.progress;
    if (!id) S.data.goals.push(g); save(); closeModal(); render(); toast('Zapisano podprojekt');
  },
  project: (f, fd) => {
    const id = f.dataset.id, p = id ? projById(id) : {id: uid(), createdAt: TODAY()};
    if (fd.start && fd.due && fd.due < fd.start) return toast('Termin zakończenia wcześniejszy niż start');
    Object.assign(p, {title: fd.title.trim(), ctx: fd.ctx, color: fd.color || p.color, status: fd.status, start: fd.start, due: fd.due, client: fd.client, desc: fd.desc});
    if (p.ctx === 'job' && +fd.budgetH > 0) p.budgetH = +fd.budgetH; else delete p.budgetH;
    if (!S.data.projects) S.data.projects = [];
    if (!id) S.data.projects.push(p);
    else { S.data.goals.forEach(g => g.projectId === p.id && (g.ctx = p.ctx)); S.data.tasks.forEach(t => projOf(t) === p.id && (t.ctx = p.ctx)); }
    save(); closeModal(); if (!id) { S.view = 'project'; S.proj.open = p.id; history.replaceState(null, '', '#project'); } render(); toast('Zapisano projekt');
  },
  mtg: async (f, fd) => {
    const id = f.dataset.id, old = id && S.data.meetings.find(x => x.id === id);
    const start = `${fd.date}T${fd.from}`, end = `${fd.date}T${fd.to}`;
    if (end <= start) return toast('Godzina zakończenia musi być późniejsza niż rozpoczęcia');
    const pj = projById(fd.projectId), att = (fd.att || '').split(/[,;\s]+/).map(s => s.trim()).filter(Boolean);
    const data = {title: fd.title.trim(), start, end, ctx: pj ? pj.ctx : fd.ctx, projectId: pj ? pj.id : null, attendees: att, location: fd.loc, body: fd.body};
    if (old && isOutlook(old)) {
      if (old.demo) { S.data.hiddenDemo = [...new Set([...(S.data.hiddenDemo || []), mtgKey(old)])]; old.demo = false; }
      Object.assign(old, {ctx: pj ? pj.ctx : fd.ctx, projectId: pj ? pj.id : null, body: fd.body});
      save(); closeModal(); render(); return toast('Zapisano projekt/notatki (temat i termin zmieniasz w Outlooku)');
    }
    if (old) {
      if (old.demo) { S.data.hiddenDemo = [...new Set([...(S.data.hiddenDemo || []), mtgKey(old)])]; old.demo = false; }
      Object.assign(old, data, fd.source ? {source: fd.source} : {});
      save(); closeModal(); render(); return toast(old.ext ? `Zapisano w TM (zmiany terminu wprowadź też w ${SRC[old.source][0]})` : 'Zapisano spotkanie');
    }
    const m = Object.assign({id: 'tm-' + uid(), source: fd.source, link: ''}, data);
    const api = m.source === 'google' ? S.conn.google : m.source === 'local' ? false : S.conn.ms;
    if (api) { try { await (m.source === 'google' ? GG : MS).createEvent(m); closeModal(); toast(`Utworzono w ${SRC[m.source][0]}`); return syncExternal(); } catch (e) { return toast('Błąd: ' + e.message); } }
    S.data.meetings.push(m); save(); closeModal(); render();
    const url = fd.invite && meetingUrl(m, fd.body);
    if (url && m.source === 'teams') return void teamsLaunch(url, 'Zapisano w TM · otworzono Teams do wysłania zaproszeń', {subject: m.title, start: m.start, end: m.end, attendees: (m.attendees || []).join('; '), body: fd.body});
    if (url && m.source === 'outlook') return void outlookApp({action: 'meeting', subject: m.title, start: m.start, end: m.end, attendees: (m.attendees || []).join('; '), location: m.location, body: fd.body}, url, 'Zapisano w TM · otworzono Outlook do wysłania zaproszeń');
    if (url) { openUrl(url); toast(`Zapisano w TM · otworzono ${SRC[m.source][0]} do wysłania zaproszeń`); } else toast('Dodano spotkanie');
  },
  entry: (f, fd) => {
    const id = f.dataset.id, s = `${fd.date}T${fd.from}`, e2 = `${fd.date}T${fd.to}`;
    if (e2 <= s) return toast('Godzina „do” musi być późniejsza niż „od”');
    let t = fd.taskId && S.data.tasks.find(x => x.id === fd.taskId); const pj = projById(fd.projectId);
    if (!t) { const title = (fd.title || '').trim(); if (!title) return toast('Wybierz zadanie albo wpisz opis');
      t = {id: uid(), title, ctx: pj ? pj.ctx : (S.ctx === 'all' ? 'job' : S.ctx), goalId: null, projectId: pj ? pj.id : null, start: fd.date, due: fd.date, status: 'doing', prio: 2, est: 0, notes: '', createdAt: TODAY(), doneAt: null}; S.data.tasks.push(t); }
    if (fd.markDone) { t.status = 'done'; t.doneAt = fd.date; }
    const rec = id ? entriesAll().find(x => x.id === id) : {id: uid(), manual: true, kind: 'manual'};
    Object.assign(rec, {taskId: t.id, title: t.title, projectId: (pj && pj.id) || projOf(t) || t.projectId || null, ctx: pj ? pj.ctx : t.ctx, start: s, end: e2, startMs: strMs(s), endMs: strMs(e2), note: fd.note || ''});
    if (!id) entriesAll().push(rec);
    const ov = entriesAll().filter(x => x !== rec && x.start.slice(0, 10) === fd.date && x.start < e2 && x.end > s).length;
    save(); closeModal(); render();
    toastAction(`Zapisano wpis ${fd.from}–${fd.to} (${fmtH(entMin(rec))})${ov ? ' · uwaga: nakłada się na inny wpis' : ''}`, 'Pokaż w ewidencji', () => { const d = P(fd.date), td = today(); S.tl = Object.assign(S.tl || {}, {period: iso(d) === iso(td) ? 'today' : 'range', from: fd.date, to: fd.date, proj: '', q: ''}); S.view = 'timelog'; history.replaceState(null, '', '#timelog'); render(); }, 7000);
  },
  tpl: (f, fd) => {
    const id = f.dataset.id, x = id ? tplById(id) : {id: uid(), uses: 0};
    const pj = projById(fd.projectId), g = goalById(fd.goalId);
    Object.assign(x, {name: fd.name.trim(), category: (fd.category || '').trim() || 'Bez kategorii', title: (fd.title || '').trim(), ctx: pj ? pj.ctx : fd.ctx, projectId: pj ? pj.id : (g?.projectId || null), goalId: g ? g.id : null,
      prio: +fd.prio || 2, est: fd.est === '' ? '' : +fd.est, dueRule: fd.dueRule || '', notes: fd.notes || ''});
    if (!id) tplAll().push(x); save(); closeModal(); if ((S.view === 'templates' || S.view === 'settings') && S.tplCat && S.tplCat !== '__all' && S.tplCat !== x.category) S.tplCat = x.category; render(); toast(`Zapisano szablon „${x.name}”`);
  },
  note: async (f, fd) => {
    const id = f.dataset.id, isNew = !id, now = new Date().toISOString();
    if (!fd.link) return toast('Wybierz zadanie lub spotkanie, do którego należy notatka');
    const [kind, refId] = fd.link.split(':');
    const n = isNew ? {id: uid(), createdAt: now} : notesAll().find(x => x.id === id); if (!n) return;
    const oldRel = n.file;
    Object.assign(n, {title: fd.title.trim() || 'Notatka', body: fd.body || '', kind, refId, updatedAt: now});
    const o = noteObj(n); if (o) { n.refTitle = objTitle(kind, o); n.projSnap = kind === 'task' ? projOf(o) : o.projectId; }
    if (isNew) notesAll().push(n);
    save();
    const how = await writeNoteFile(n, oldRel); save();
    retTo(f.dataset.ret)();
    toast(how === 'disk' ? `Zapisano notatkę · ${n.file.replace(/\//g, '\\')}` : 'Zapisano notatkę · plik .txt pobrano (uruchom ponownie start-TM.bat, aby zapisywać na dysk)');
  },
  profile: (f, fd) => { S.user.name = fd.name.trim() || S.user.name; LS.set('tm:user', S.user); render(); toast('Zapisano dane konta'); },
  import: (f, fd) => {
    const c = impCfg(), url = (fd.icsUrl || '').trim();
    if (url && !/^https:\/\/[^\s]+$/i.test(url)) return toast('Link ICS musi zaczynać się od https://');
    const changed = url !== c.icsUrl || (fd.method || c.method) !== c.method;
    Object.assign(c, {method: fd.method || c.method, icsUrl: url || (fd.method === 'ics' ? '' : c.icsUrl), back: +fd.back, fwd: +fd.fwd, every: +fd.every || 60, auto: !!fd.auto});
    if (changed) c.last = 0;
    save();
    if (f.dataset.submitter === 'run') return runImport();
    toast('Zapisano ustawienia importu'); render();
  },
  vac: (f, fd) => { if (fd.to < fd.from) return toast('Data „do” wcześniejsza niż „od”'); const half = !!f.half?.checked && fd.from === fd.to; if (f.half?.checked && !half) toast('Pół dnia można ustawić tylko dla urlopu jednodniowego');
    S.data.vacations.push({id: uid(), from: fd.from, to: fd.to, note: fd.note, ...(half ? {half: true} : {})}); save(); render(); toast(`Dodano urlop: ${vacationWorkdays({from: fd.from, to: fd.to, half})} dni roboczych`); },
  meeting: async (f, fd) => {
    const att = fd.att.split(/[,;\s]+/).map(s => s.trim()).filter(Boolean), st = P(fd.date); const [hh, mm] = fd.time.split(':').map(Number);
    st.setHours(hh, mm); const en = new Date(st.getTime() + fd.dur * 60000), L = d => `${iso(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
    const pjm = projById(fd.projectId);
    const m = {id: 'tm-' + uid(), title: fd.title, start: L(st), end: L(en), source: fd.via, ctx: pjm ? pjm.ctx : (fd.via === 'google' ? 'priv' : 'job'), projectId: pjm ? pjm.id : null, attendees: att, location: fd.loc, body: fd.body, link: ''};
    const api = fd.via === 'google' ? S.conn.google : S.conn.ms;
    if (api) { try { await (fd.via === 'google' ? GG : MS).createEvent(m); toast('Utworzono spotkanie'); S.integ.prefill = null; return syncExternal(); } catch (e) { return toast('Błąd: ' + e.message); } }
    const q = o => Object.entries(o).filter(([, v]) => v).map(([k, v]) => k + '=' + encodeURIComponent(v)).join('&');
    const gfmt = d => d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
    const url = fd.via === 'teams' ? 'msteams:/l/meeting/new?' + q({subject: m.title, attendees: att.join(','), startTime: st.toISOString(), endTime: en.toISOString(), content: fd.body})
      : fd.via === 'outlook' ? 'https://outlook.office.com/calendar/deeplink/compose?' + q({subject: m.title, to: att.join(','), startdt: st.toISOString(), enddt: en.toISOString(), body: fd.body, location: fd.loc})
        : 'https://calendar.google.com/calendar/render?' + q({action: 'TEMPLATE', text: m.title, dates: gfmt(st) + '/' + gfmt(en), add: att.join(','), details: fd.body, location: fd.loc});
    m.source = fd.via; S.data.meetings.push(m); save(); S.integ.prefill = null; S.integ.tab = 'meetings';
    if (fd.via === 'teams') { render(); return void teamsLaunch(url, 'Otworzono kreator spotkania w Teams · zapisano w kalendarzu TM', {subject: m.title, start: m.start, end: m.end, attendees: att.join('; '), body: fd.body}); }
    if (fd.via === 'outlook') { render(); return void outlookApp({action: 'meeting', subject: m.title, start: m.start, end: m.end, attendees: att.join('; '), location: fd.loc, body: fd.body}, url, 'Otworzono spotkanie w Outlook · zapisano w kalendarzu TM'); }
    openUrl(url); render(); toast('Otworzono kreator spotkania · zapisano w kalendarzu TM');
  },
  mail: async (f, fd) => {
    if (fd.via === 'outlook' && S.conn.ms) { try { await MS.sendMail(fd); return toast('Wysłano przez Outlook'); } catch (e) { return toast(e.message); } }
    if (fd.via === 'gmail' && S.conn.google) { try { await GG.sendMail(fd); return toast('Wysłano przez Gmail'); } catch (e) { return toast(e.message); } }
    const q = o => Object.entries(o).filter(([, v]) => v).map(([k, v]) => k + '=' + encodeURIComponent(v)).join('&');
    const url = fd.via === 'gmail' ? 'https://mail.google.com/mail/?view=cm&fs=1&' + q({to: fd.to, cc: fd.cc, su: fd.subject, body: fd.body})
      : fd.via === 'outlook' ? 'https://outlook.office.com/mail/deeplink/compose?' + q({to: fd.to, cc: fd.cc, subject: fd.subject, body: fd.body})
        : `mailto:${encodeURIComponent(fd.to)}?` + q({cc: fd.cc, subject: fd.subject, body: fd.body});
    if (fd.via === 'outlook') return void outlookApp({action: 'mail', to: fd.to, cc: fd.cc, subject: fd.subject, body: fd.body}, url, 'Otworzono nową wiadomość w Outlook');
    openUrl(url); toast('Otworzono okno nowej wiadomości');
  }
};

/* ================= start ================= */
document.addEventListener('click', e => {
  const el = e.target.closest('[data-action]'); if (!el) return;
  if (el.dataset.action === 'fb-person') return A['fb-person'](el, e);
  if (el.dataset.action === 'modal-bg' && e.target !== el) return; // klik wewnątrz okna – nie blokuj formularzy
  const f = A[el.dataset.action]; if (f) { if (el.tagName !== 'LABEL' && el.tagName !== 'INPUT') e.preventDefault(); f(el, e); }
});
document.addEventListener('submit', e => { const f = e.target.closest('form[data-form]'); if (!f) return; e.preventDefault(); FORMS[f.dataset.form](f, Object.fromEntries(new FormData(f))); });
document.addEventListener('change', e => {
  if (e.target.id !== 'bulkProj' || !e.target.value) return;
  const v = e.target.value, p = projById(v), ids = S.mtgs.sel; let n = 0;
  S.data.meetings.forEach(m => { if (ids.has(m.id)) { if (m.demo) { S.data.hiddenDemo = [...new Set([...(S.data.hiddenDemo || []), mtgKey(m)])]; m.demo = false; } m.projectId = p ? p.id : null; if (p) m.ctx = p.ctx; n++; } });
  save(); render(); toast(`Przypisano ${p ? '„' + p.title + '”' : 'brak projektu'} do ${n} spotk.`);
});
document.addEventListener('input', e => {
  const b = e.target.dataset?.bind; if (!b) return;
  setPath(b, e.target.type === 'number' ? +e.target.value : e.target.value);
  if (b.startsWith('settings')) save();
  if (b === 'tasks.q' || b === 'mtgs.q' || b === 'proj.q' || b === 'proj.pq' || b === 'proj.nq' || b === 'tplQ' || b === 'tl.q') { const pos = e.target.selectionStart; render(); const i = $(`[data-bind="${b}"]`); i.focus(); i.setSelectionRange(pos, pos); }
  else if (b.startsWith('mtgs.') || /^(tl|tr)\./.test(b)) render();
  else if (b === 'homeGoal') { const y = $('main').scrollTop; render(); $('main').scrollTop = y; }
  else if (b === 'integ.fbDate') { S.integ.fb = null; $('#fbGrid').innerHTML = H(fbGrid()); }
  else if (e.target.tagName === 'SELECT' || b.startsWith('settings')) render();
});
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') return closeModal();
  const typing = /INPUT|TEXTAREA|SELECT/.test(document.activeElement?.tagName);
  if (!S.user || typing || $('#modalRoot').innerHTML) return;
  if (e.key === 'n' || e.key === 'N') { e.preventDefault(); quickAdd(); }
  if (e.altKey && ['1', '2', '3'].includes(e.key)) { S.ctx = ['all', 'job', 'priv'][+e.key - 1]; LS.set('tm:ctx', S.ctx); render(); }
});
$('#ctxSwitch').addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; S.ctx = b.dataset.ctx; LS.set('tm:ctx', S.ctx); render(); });
$('#avatar').addEventListener('click', () => { S.view = 'profile'; S.profTab = 'account'; history.replaceState(null, '', '#profile'); render(); });
$('#navBack').addEventListener('click', () => navGo(-1)); $('#navFwd').addEventListener('click', () => navGo(1));
$('#bellBtn').addEventListener('click', () => { S.view = 'notifs'; history.replaceState(null, '', '#notifs'); render(); });
searchInit();
$('#loginForm').addEventListener('submit', e => { e.preventDefault(); const f = e.target; login({email: f.email.value, name: f.name.value.trim(), provider: 'email'}); });
window.addEventListener('hashchange', () => { const v = location.hash.slice(1); if (S.user && VIEWS[v] && v !== S.view) { S.view = v; render(); } });

/* ================= SZABLONY PROJEKTÓW (podprojekty + zadania; Projekty › Szablony projektów) =================
   szablon: {id, name, category, ctx, desc, days, goals:[{id, title, from, to, desc, tasks:[{id, title, day, est, prio, notes}]}], uses, lastUsed}
   from/to/day = dzień od startu projektu; przy tworzeniu projektu terminy przeliczane są od wybranej daty startu
   (opcjonalnie skalowane do innego czasu trwania i przesuwane na dni robocze – bez weekendów, świąt i urlopów). */
function seedProjTemplates() {
  const T = (title, day, est, prio = 2, notes = '') => ({id: uid(), title, day, est, prio, notes});
  const G = (title, from, to, tasks, desc = '') => ({id: uid(), title, from, to, desc, tasks});
  const std = (mod, role, cfg, mig, crp) => [
    G('Analiza przedwdrożeniowa i projekt rozwiązania', 0, 20, [
      T(`Kick-off modułu ${mod} – zakres, harmonogram, zespół klienta`, 2, 3, 1),
      ...cfg.ws.map((w, i) => T(w, 4 + i * 3, 8, 1)),
      T('Analiza fit-gap i lista luk / modyfikacji', 15, 12, 1),
      T('HLS dla zidentyfikowanych modyfikacji', 18, 16, 2),
      T('Akceptacja koncepcji rozwiązania przez klienta', 20, 2, 1)], `Zrozumienie procesów klienta ${mod.toLowerCase()} i uzgodnienie docelowego rozwiązania w IFS Cloud.`),
    G('Konfiguracja systemu', 18, 50, cfg.conf.map((c, i) => T(c, 24 + Math.round(i * 24 / Math.max(1, cfg.conf.length - 1)), 8, i < 2 ? 1 : 2)),
      'Ustawienia i dane podstawowe zgodnie z zaakceptowaną koncepcją.'),
    G('Migracja danych', 30, 72, [T('Szablony migracyjne i mapowanie danych', 34, 6, 2), ...mig.map((m, i) => T(m, 44 + i * 6, 6, 2)), T('Migracja – próba generalna i uzgodnienie z klientem', 70, 6, 1)],
      'Przeniesienie danych podstawowych i otwartych transakcji.'),
    G('Testy i akceptacja', 52, 86, [
      T(`Scenariusze testowe CRP – ${mod}`, 55, 6, 2), T(crp, 62, 8, 1), T('Poprawki i konfiguracja po CRP', 70, 8, 2),
      T('Testy akceptacyjne UAT z użytkownikami kluczowymi', 80, 12, 1), T('Protokół akceptacji UAT', 86, 2, 1)], 'Weryfikacja procesów end-to-end i formalna akceptacja.'),
    G('Szkolenia i dokumentacja', 72, 94, [
      T(`Instrukcje stanowiskowe – ${role}`, 80, 8, 2), T(`Szkolenie użytkowników kluczowych – ${mod}`, 86, 8, 1), T('Szkolenie użytkowników końcowych', 92, 8, 2), T('Dokumentacja powdrożeniowa (konfiguracja, procedury)', 94, 6, 3)],
      'Przygotowanie zespołu klienta do samodzielnej pracy.'),
    G('Uruchomienie produkcyjne i hypercare', 90, 112, [
      T('Plan przełączenia (cut-over) i checklista go-live', 90, 4, 1), T('Migracja produkcyjna i weryfikacja danych', 94, 6, 1), T(`Go-live modułu ${mod}`, 96, 4, 1),
      T('Wsparcie powdrożeniowe (hypercare)', 110, 16, 2), T('Odbiór etapu i przekazanie do utrzymania', 112, 2, 1)], 'Bezpieczny start i stabilizacja po uruchomieniu.')];
  const mk = (name, mod, role, cfg, mig, crp, desc) => ({id: uid(), name, category: 'Wdrożenia IFS Cloud', ctx: 'job', desc, days: 112, goals: std(mod, role, cfg, mig, crp), uses: 0});
  return [
    mk('Wdrożenie modułu Projekty', 'Projekty', 'kierownik projektu i kontroling', {
      ws: ['Warsztat: cykl życia projektu, planowanie i budżetowanie', 'Warsztat: struktura WBS, aktywności i raportowanie czasu', 'Warsztat: rozliczanie i fakturowanie projektów'],
      conf: ['Typy projektów, statusy i parametry modułu', 'Elementy kosztowe i struktura kontrolingowa', 'Budżetowanie i prognozy projektu', 'Raportowanie czasu i kosztów do aktywności', 'Fakturowanie projektowe (kamienie milowe / T&M)', 'Integracja z Finansami – rozpoznanie przychodu i kosztów']},
      ['Migracja otwartych projektów i aktywności', 'Migracja budżetów i kosztów otwarcia', 'Migracja zaawansowania i fakturowania'], 'CRP: projekt od założenia do rozliczenia',
      'Standardowe podprojekty i zadania wdrożenia IFS Cloud Project Management.'),
    mk('Wdrożenie modułu Produkcja', 'Produkcja', 'planista i mistrz produkcji', {
      ws: ['Warsztat: modele produkcji (MTO/MTS), typy zleceń', 'Warsztat: struktury produktowe (BOM) i marszruty', 'Warsztat: planowanie MRP / MS i zdolności produkcyjne'],
      conf: ['Części magazynowe i parametry planistyczne', 'Struktury produktowe (BOM) i warianty', 'Marszruty, gniazda produkcyjne i zasoby', 'Zlecenia produkcyjne i raportowanie operacji', 'Planowanie MRP / harmonogram główny', 'Kalkulacja kosztu wytworzenia (costing)']},
      ['Migracja kartoteki części', 'Migracja struktur BOM i marszrut', 'Migracja otwartych zleceń i stanów WIP'], 'CRP: od zapotrzebowania do przyjęcia wyrobu',
      'Standardowe podprojekty i zadania wdrożenia IFS Cloud Manufacturing.'),
    mk('Wdrożenie modułu Kontrola jakości', 'Kontrola jakości', 'kontroler jakości', {
      ws: ['Warsztat: kontrola przyjęć, międzyoperacyjna i końcowa', 'Warsztat: plany kontroli, charakterystyki i metody pomiaru', 'Warsztat: niezgodności, reklamacje i działania korygujące'],
      conf: ['Charakterystyki, metody i jednostki pomiarowe', 'Plany kontroli i reguły pobierania próbek', 'Kontrola przyjęć dostaw (receipt inspection)', 'Kontrola w produkcji i kontrola końcowa', 'Raporty niezgodności (NCR) i działania CAPA / 8D', 'Atesty, certyfikaty i ocena dostawców']},
      ['Migracja charakterystyk i planów kontroli', 'Migracja ocen i statusów dostawców', 'Migracja otwartych niezgodności'], 'CRP: od dostawy przez kontrolę do decyzji jakościowej',
      'Standardowe podprojekty i zadania wdrożenia IFS Cloud Quality Management.'),
    mk('Wdrożenie modułu Remonty i Serwis', 'Remonty i Serwis', 'planista utrzymania ruchu i serwisant', {
      ws: ['Warsztat: utrzymanie ruchu – prewencja i awarie', 'Warsztat: serwis zewnętrzny – zgłoszenia, umowy, SLA', 'Warsztat: obiekty, hierarchia urządzeń i części zamienne'],
      conf: ['Obiekty i hierarchia urządzeń', 'Plany prewencyjne (PM) i harmonogramy przeglądów', 'Zgłoszenia awarii i zlecenia pracy (Work Orders)', 'Planowanie zasobów, techników i części zamiennych', 'Umowy serwisowe, SLA i rozliczanie serwisu', 'Aplikacja mobilna dla techników (Mobile Work Order)']},
      ['Migracja obiektów i hierarchii', 'Migracja planów prewencyjnych', 'Migracja umów serwisowych i historii'], 'CRP: od zgłoszenia awarii do rozliczenia zlecenia',
      'Standardowe podprojekty i zadania wdrożenia IFS Cloud Maintenance i Service Management.')
  ];
}
const plN = (n, a, b, c) => `${n} ${n === 1 ? a : [2, 3, 4].includes(n % 10) && ![12, 13, 14].includes(n % 100) ? b : c}`;
const nCel = n => plN(n, 'podprojekt', 'podprojekty', 'podprojektów'), nZad = n => plN(n, 'zadanie', 'zadania', 'zadań');
const ptplAll = () => (S.data.projTemplates ||= []);
const ptplById = id => ptplAll().find(t => t.id === id);
const ptplCount = tp => tp.goals.reduce((a, g) => a + g.tasks.length, 0);
const ptplHours = tp => tp.goals.reduce((a, g) => a + g.tasks.reduce((b, t) => b + (+t.est || 0), 0), 0);
/* przeliczenie terminów szablonu od daty startu */
function ptplPlan(tp, startIso, days, work) {
  const base = P(startIso), k = (days || tp.days || 1) / (tp.days || 1);
  const D = off => { let d = add(base, Math.round((+off || 0) * k)); if (work) { let n = 0; while (!isWork(d) && n++ < 30) d = add(d, 1); } return iso(d); };
  return {start: startIso, due: D(tp.days), goals: tp.goals.map(g => ({g, start: D(g.from), due: D(g.to), tasks: g.tasks.map(t => ({t, due: D(t.day)}))}))};
}
function vProjTemplates() {
  const l = ptplAll().filter(inCtx), cats = [...new Set(l.map(t => t.category || 'Bez kategorii'))];
  const card = tp => `<div class="card tplcard ${tp.ctx}">
    <div class="row" style="gap:8px"><b class="tpl-name" data-action="ptpl-edit" data-id="${tp.id}">${esc(tp.name)}</b><div class="spacer"></div>${ctxTag(tp.ctx)}</div>
    ${tp.desc ? `<div class="muted" style="margin:6px 0 0;font-size:13px">${esc(tp.desc)}</div>` : ''}
    <div class="row ptpl-stats"><span>⟨i:layers⟩ ${nCel(tp.goals.length)}</span><span>⟨i:check-square⟩ ${nZad(ptplCount(tp))}</span><span>⟨i:clock⟩ ${r1(ptplHours(tp))} h</span><span>⟨i:calendar⟩ ${tp.days} dni</span>${tp.uses ? `<span class="muted">użyty ${tp.uses}×</span>` : ''}</div>
    <details class="ptpl-goals"><summary>Podprojekty i zadania</summary><ol>${tp.goals.map(g => `<li><b>${esc(g.title)}</b> <span class="muted">dzień ${g.from}–${g.to} · ${g.tasks.length} zad.</span><ul>${g.tasks.map(t => `<li>${esc(t.title)} <span class="muted">· dz. ${t.day}${t.est ? ` · ${t.est}h` : ''} · P${t.prio}</span></li>`).join('')}</ul></li>`).join('')}</ol></details>
    <div class="row" style="margin-top:10px;gap:6px"><button class="btn sm primary" data-action="ptpl-use" data-id="${tp.id}">⟨i:plus⟩ Utwórz projekt</button><button class="btn sm" data-action="ptpl-edit" data-id="${tp.id}">⟨i:pencil⟩ Edytuj</button><button class="btn sm" data-action="ptpl-dup" data-id="${tp.id}" title="Duplikuj szablon">⟨i:template⟩ Duplikuj</button><div class="spacer"></div><button class="btn sm ghost" data-action="ptpl-del" data-id="${tp.id}" title="Usuń szablon">⟨i:trash⟩</button></div></div>`;
  return projTabs() + `<div class="row" style="margin-bottom:14px"><p class="muted" style="margin:0;flex:1;min-width:240px">Gotowe zestawy podprojektów i zadań. Nowy projekt z szablonu dostaje wszystkie podprojekty i zadania z terminami liczonymi od daty startu.</p>
    <button class="btn" data-action="ptpl-edit">⟨i:plus⟩ Nowy szablon</button></div>
    ${l.length ? cats.map(c => `<h3 class="group-h">${esc(c)}</h3><div class="grid g2 ptpl-grid">${l.filter(t => (t.category || 'Bez kategorii') === c).map(card).join('')}</div>`).join('') : '<div class="card empty">Brak szablonów projektów w tym kontekście. Utwórz nowy albo zapisz istniejący projekt jako szablon (ekran Projekt → „Zapisz jako szablon”).</div>'}`;
}
/* --- edytor szablonu: podprojekty z listą zadań --- */
function ptplEditor(tp) {
  const isNew = !tp || !ptplById(tp.id);
  S.ptplDraft = JSON.parse(JSON.stringify(tp || {id: uid(), name: '', category: 'Wdrożenia IFS Cloud', ctx: S.ctx === 'priv' ? 'priv' : 'job', desc: '', days: 90, goals: [{id: uid(), title: '', from: 0, to: 30, tasks: [{id: uid(), title: '', day: 10, est: 4, prio: 2}]}], uses: 0}));
  S.ptplDraft._new = isNew; ptplEditorRender();
}
function ptplEditorRender(focusSel) {
  const d = S.ptplDraft, cats = [...new Set(ptplAll().map(t => t.category).filter(Boolean))];
  const trow = (gi, t, ti) => `<div class="pt-trow"><input class="input" data-g="${gi}" data-t="${ti}" data-k="title" value="${esc(t.title)}" placeholder="Zadanie…">
    <label title="Dzień od startu projektu (termin)">dz.<input class="input" type="number" min="0" data-g="${gi}" data-t="${ti}" data-k="day" value="${t.day}"></label>
    <label title="Estymacja w godzinach">h<input class="input" type="number" min="0" step="0.5" data-g="${gi}" data-t="${ti}" data-k="est" value="${t.est || ''}"></label>
    <select class="input" data-g="${gi}" data-t="${ti}" data-k="prio" title="Priorytet">${[1, 2, 3].map(n => `<option value="${n}" ${+t.prio === n ? 'selected' : ''}>P${n}</option>`).join('')}</select>
    <button type="button" class="iconbtn" data-action="ptpl-t-del" data-g="${gi}" data-t="${ti}" title="Usuń zadanie">⟨i:x⟩</button></div>`;
  const gblock = (g, gi) => `<div class="pt-goal"><div class="pt-ghead"><span class="pt-num">${gi + 1}</span><input class="input" data-g="${gi}" data-k="title" value="${esc(g.title)}" placeholder="Podprojekt, np. Konfiguracja systemu" style="font-weight:600">
      <label title="Dzień rozpoczęcia podprojektu">od dz.<input class="input" type="number" min="0" data-g="${gi}" data-k="from" value="${g.from}"></label><label title="Dzień zakończenia podprojektu">do dz.<input class="input" type="number" min="0" data-g="${gi}" data-k="to" value="${g.to}"></label>
      <button type="button" class="iconbtn" data-action="ptpl-g-move" data-g="${gi}" data-dir="-1" title="W górę" ${gi ? '' : 'disabled'}>⟨i:chev-up⟩</button><button type="button" class="iconbtn" data-action="ptpl-g-move" data-g="${gi}" data-dir="1" title="W dół" ${gi < d.goals.length - 1 ? '' : 'disabled'}>⟨i:chev-down⟩</button>
      <button type="button" class="iconbtn" data-action="ptpl-g-del" data-g="${gi}" title="Usuń podprojekt z zadaniami">⟨i:trash⟩</button></div>
    <div class="pt-tasks">${g.tasks.map((t, ti) => trow(gi, t, ti)).join('')}<button type="button" class="btn sm ghost" data-action="ptpl-t-add" data-g="${gi}">⟨i:plus⟩ Zadanie</button></div></div>`;
  modal(`<h3>⟨i:template⟩ ${d._new ? 'Nowy szablon projektu' : 'Edycja szablonu projektu'}</h3><form data-form="ptpl" class="grid g2 ptpl-form">
    <label class="f">Nazwa szablonu<input class="input" data-k="name" name="name" required value="${esc(d.name)}" placeholder="np. Wdrożenie modułu Finanse"></label>
    <label class="f">Kategoria<input class="input" data-k="category" value="${esc(d.category || '')}" list="ptplCats" placeholder="np. Wdrożenia IFS Cloud"><datalist id="ptplCats">${cats.map(c => `<option value="${esc(c)}">`).join('')}</datalist></label>
    <label class="f">Kontekst<select class="input" data-k="ctx"><option value="job" ${d.ctx === 'job' ? 'selected' : ''}>JOB – firmowy</option><option value="priv" ${d.ctx === 'priv' ? 'selected' : ''}>PRIV – prywatny</option></select></label>
    <label class="f">Czas trwania projektu (dni kalendarzowe)<input class="input" type="number" min="1" data-k="days" value="${d.days}"></label>
    <label class="f" style="grid-column:1/-1">Opis<input class="input" data-k="desc" value="${esc(d.desc || '')}" placeholder="opcjonalnie"></label>
    <div style="grid-column:1/-1"><div class="row" style="margin-bottom:8px"><b>Podprojekty i zadania</b><span class="muted" style="font-size:12px">„dz.” = dzień od startu projektu (0 = dzień startu)</span><div class="spacer"></div><span class="muted" style="font-size:12px">${nCel(d.goals.length)} · ${nZad(ptplCount(d))} · ${r1(ptplHours(d))} h</span></div>
      <div class="pt-goals">${d.goals.map(gblock).join('')}</div>
      <button type="button" class="btn sm" data-action="ptpl-g-add" style="margin-top:8px">⟨i:plus⟩ Dodaj podprojekt</button></div>
    <div class="row" style="grid-column:1/-1;justify-content:space-between"><span></span><div class="row"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">Zapisz szablon</button></div></div></form>`,
  m => { m.classList.add('wide'); const f = focusSel ? $(focusSel, m) : $('[data-k=name]', m); f && f.focus(); });
}
/* odczyt pól edytora do szkicu (przed każdą zmianą struktury) */
function ptplSync() {
  const d = S.ptplDraft; if (!d) return;
  $$('#modalRoot [data-k]').forEach(el => {
    const k = el.dataset.k, num = ['day', 'est', 'prio', 'from', 'to', 'days'].includes(k), v = num ? (el.value === '' ? 0 : +el.value) : el.value;
    if (el.dataset.t != null) d.goals[+el.dataset.g].tasks[+el.dataset.t][k] = v;
    else if (el.dataset.g != null) d.goals[+el.dataset.g][k] = v;
    else d[k] = v;
  });
}
/* --- utworzenie projektu z szablonu --- */
function ptplUse(id) {
  const list = ptplAll(); if (!list.length) return toast('Brak szablonów projektów – utwórz szablon w Projekty › Szablony projektów');
  const tp = ptplById(id) || list[0], st = TODAY();
  modal(`<h3>⟨i:folder⟩ Nowy projekt z szablonu</h3><form data-form="ptpluse" class="grid g2">
    <label class="f" style="grid-column:1/-1">Szablon<select class="input" name="tpl">${[...new Set(list.map(t => t.category || 'Bez kategorii'))].map(c => `<optgroup label="${esc(c)}">${list.filter(t => (t.category || 'Bez kategorii') === c).map(t => `<option value="${t.id}" ${t.id === tp.id ? 'selected' : ''}>${esc(t.name)} (${nCel(t.goals.length)}, ${nZad(ptplCount(t))})</option>`).join('')}</optgroup>`).join('')}</select></label>
    <label class="f">Klient<input class="input" name="client" placeholder="np. Alfa S.A."></label>
    <label class="f">Nazwa projektu<input class="input" name="title" required value="${esc(tp.name)}"></label>
    <label class="f">Start projektu<input class="input" type="date" name="start" required value="${st}"></label>
    <label class="f">Czas trwania (dni)<input class="input" type="number" min="1" name="days" value="${tp.days}"></label>
    <label class="f" style="grid-column:1/-1"><span class="row" style="gap:6px;font-weight:500"><input type="checkbox" name="work" checked> Terminy tylko w dni robocze (pomija weekendy, święta i Twoje urlopy)</span></label>
    <div id="ptplPrev" class="ptpl-prev" style="grid-column:1/-1"></div>
    <div class="row" style="grid-column:1/-1;justify-content:flex-end"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">⟨i:plus⟩ Utwórz projekt</button></div></form>`, m => {
    const f = $('form', m); let lastTpl = tp.id;
    const upd = () => {
      const t = ptplById(f.tpl.value); if (!t) return;
      if (t.id !== lastTpl) { const oldName = ptplById(lastTpl)?.name || ''; if (!f.title.value || f.title.value.startsWith(oldName)) f.title.value = t.name + (f.client.value ? ' – ' + f.client.value : ''); f.days.value = t.days; lastTpl = t.id; }
      const pl = ptplPlan(t, f.start.value || TODAY(), +f.days.value || t.days, f.work.checked);
      $('#ptplPrev', m).innerHTML = H(`<div class="row" style="gap:14px;margin-bottom:6px"><b>Podgląd</b><span class="muted">${fmt(pl.start)} → ${fmt(pl.due)} · ${nCel(t.goals.length)} · ${nZad(ptplCount(t))} · ${r1(ptplHours(t))} h</span></div>
        <table class="tbl"><tbody>${pl.goals.map(x => `<tr><td>⟨i:layers⟩ ${esc(x.g.title)}</td><td class="muted">${fmt(x.start)} – ${fmt(x.due)}</td><td class="num muted">${x.tasks.length} zad.</td></tr>`).join('')}</tbody></table>`);
    };
    f.client.addEventListener('input', () => { const t = ptplById(f.tpl.value); if (t && f.title.value.startsWith(t.name)) f.title.value = t.name + (f.client.value ? ' – ' + f.client.value : ''); });
    ['change', 'input'].forEach(ev => f.addEventListener(ev, e => { if (e.target.name !== 'title') upd(); }));
    upd(); f.client.focus();
  });
}
function ptplCreate(tp, fd) {
  const pl = ptplPlan(tp, fd.start, +fd.days || tp.days, !!fd.work), now = new Date().toISOString();
  const p = {id: uid(), createdAt: TODAY(), title: fd.title.trim() || tp.name, ctx: tp.ctx, color: PCOL[(S.data.projects || []).length % PCOL.length], status: fd.start > TODAY() ? 'planned' : 'active',
    start: pl.start, due: pl.due, client: fd.client || '', desc: tp.desc || '', fromTpl: tp.id};
  const goals = [], tasks = [];
  pl.goals.forEach(x => {
    const g = {id: uid(), title: x.g.title, ctx: tp.ctx, projectId: p.id, start: x.start, due: x.due, desc: x.g.desc || '', progress: 0};
    goals.push(g);
    x.tasks.forEach(y => { let s0 = startFor(y.due, +y.t.est || 0); if (s0 < g.start) s0 = g.start; if (s0 > y.due) s0 = y.due;
      tasks.push({id: uid(), title: y.t.title, ctx: tp.ctx, goalId: g.id, projectId: p.id, start: s0, due: y.due, status: 'todo', prio: +y.t.prio || 2, est: +y.t.est || 0, notes: y.t.notes || '', createdAt: TODAY(), doneAt: null}); });
  });
  (S.data.projects ||= []).push(p); S.data.goals.push(...goals); S.data.tasks.push(...tasks);
  tp.uses = (tp.uses || 0) + 1; tp.lastUsed = now;
  return {p, goals, tasks};
}
/* zapis istniejącego projektu jako szablon (terminy → dni od startu projektu) */
function ptplFromProject(p) {
  const gs = S.data.goals.filter(g => g.projectId === p.id), ts = S.data.tasks.filter(t => projOf(t) === p.id);
  const base = p.start || [...gs.map(g => g.start), ...ts.map(t => t.start)].filter(Boolean).sort()[0] || TODAY();
  const off = d => d ? Math.max(0, diffD(base, d)) : 0, tk = t => ({id: uid(), title: t.title, day: off(t.due), est: t.est || 0, prio: t.prio || 2, notes: t.notes || ''});
  const goals = gs.map(g => ({id: uid(), title: g.title, from: off(g.start), to: off(g.due), desc: g.desc || '', tasks: ts.filter(t => t.goalId === g.id).sort((a, b) => (a.due || '').localeCompare(b.due || '')).map(tk)}));
  const loose = ts.filter(t => !t.goalId || !gs.some(g => g.id === t.goalId));
  if (loose.length) goals.push({id: uid(), title: 'Pozostałe zadania', from: 0, to: Math.max(...loose.map(t => off(t.due))), tasks: loose.map(tk)});
  const days = Math.max(1, off(p.due) || Math.max(0, ...goals.map(g => g.to)));
  ptplEditor({id: uid(), name: p.title + ' (szablon)', category: 'Moje szablony', ctx: p.ctx, desc: p.desc || '', days, goals, uses: 0});
}
Object.assign(A, {
  'ptpl-edit': el => { const tp = ptplById(el.dataset.id); ptplEditor(tp); },
  'ptpl-use': el => ptplUse(el.dataset.id),
  'ptpl-dup': el => { const tp = ptplById(el.dataset.id); if (!tp) return; const c = JSON.parse(JSON.stringify(tp)); c.id = uid(); c.name = tp.name + ' (kopia)'; c.uses = 0; delete c.lastUsed; c.goals.forEach(g => { g.id = uid(); g.tasks.forEach(t => t.id = uid()); }); ptplAll().push(c); save(); render(); toast('Zduplikowano szablon'); },
  'ptpl-del': el => { const tp = ptplById(el.dataset.id); if (!tp) return; confirmDlg('Usunąć szablon projektu?', `<p>Szablon „${esc(tp.name)}” zostanie usunięty. Projekty utworzone z tego szablonu pozostaną bez zmian.</p>`, 'Usuń', () => { const i = ptplAll().indexOf(tp); ptplAll().splice(i, 1); save(); render(); toastAction('Usunięto szablon projektu', '⟨i:undo⟩ Cofnij', () => { ptplAll().splice(i, 0, tp); save(); render(); }); }); },
  'ptpl-g-add': () => { ptplSync(); const d = S.ptplDraft, last = d.goals[d.goals.length - 1]; d.goals.push({id: uid(), title: '', from: last ? last.to : 0, to: (last ? last.to : 0) + 20, tasks: [{id: uid(), title: '', day: (last ? last.to : 0) + 10, est: 4, prio: 2}]}); ptplEditorRender(`[data-g="${d.goals.length - 1}"][data-k=title]:not([data-t])`); },
  'ptpl-g-del': el => { ptplSync(); S.ptplDraft.goals.splice(+el.dataset.g, 1); ptplEditorRender(); },
  'ptpl-g-move': el => { ptplSync(); const gs = S.ptplDraft.goals, i = +el.dataset.g, j = i + +el.dataset.dir; if (j < 0 || j >= gs.length) return; [gs[i], gs[j]] = [gs[j], gs[i]]; ptplEditorRender(); },
  'ptpl-t-add': el => { ptplSync(); const g = S.ptplDraft.goals[+el.dataset.g], last = g.tasks[g.tasks.length - 1]; g.tasks.push({id: uid(), title: '', day: last ? last.day : g.from, est: 4, prio: 2}); ptplEditorRender(`[data-g="${el.dataset.g}"][data-t="${g.tasks.length - 1}"][data-k=title]`); },
  'ptpl-t-del': el => { ptplSync(); S.ptplDraft.goals[+el.dataset.g].tasks.splice(+el.dataset.t, 1); ptplEditorRender(); },
  'proj-to-tpl': el => { const p = projById(el.dataset.id); if (p) ptplFromProject(p); }
});
Object.assign(FORMS, {
  ptpl: () => {
    ptplSync(); const d = S.ptplDraft; if (!d) return;
    if (!d.name.trim()) return toast('Podaj nazwę szablonu');
    d.goals = d.goals.filter(g => g.title.trim() || g.tasks.some(t => t.title.trim()));
    d.goals.forEach(g => { g.title = g.title.trim() || 'Podprojekt bez nazwy'; g.tasks = g.tasks.filter(t => t.title.trim()); if (g.to < g.from) [g.from, g.to] = [g.to, g.from]; });
    if (!d.goals.length) return toast('Dodaj co najmniej jeden podprojekt');
    const isNew = d._new; delete d._new; d.days = Math.max(1, +d.days || Math.max(...d.goals.map(g => g.to), 1));
    if (isNew) ptplAll().push(d); else { const i = ptplAll().findIndex(t => t.id === d.id); ptplAll()[i] = d; }
    S.ptplDraft = null; save(); closeModal(); S.view = 'settings'; S.setTab = 'ptpl'; history.replaceState(null, '', '#settings'); render(); toast(isNew ? 'Zapisano nowy szablon projektu' : 'Zapisano szablon');
  },
  ptpluse: (f, fd) => {
    const tp = ptplById(fd.tpl); if (!tp) return;
    if (!fd.start) return toast('Podaj datę startu projektu');
    const r = ptplCreate(tp, fd);
    save(); closeModal(); S.view = 'project'; S.proj.open = r.p.id; history.replaceState(null, '', '#project'); render();
    toastAction(`Utworzono projekt: ${nCel(r.goals.length)}, ${nZad(r.tasks.length)}`, '⟨i:undo⟩ Cofnij', () => {
      const gid = new Set(r.goals.map(g => g.id)), tid = new Set(r.tasks.map(t => t.id));
      S.data.projects = S.data.projects.filter(p => p.id !== r.p.id); S.data.goals = S.data.goals.filter(g => !gid.has(g.id)); S.data.tasks = S.data.tasks.filter(t => !tid.has(t.id));
      tp.uses = Math.max(0, (tp.uses || 1) - 1); if (S.proj.open === r.p.id) S.proj.open = null; save(); S.view = 'settings'; S.setTab = 'ptpl'; render(); toast('Cofnięto utworzenie projektu');
    }, 12000);
  }
});
VIEWS.projtpl = ['⟨i:template⟩', 'Szablony projektów', vProjTemplates, null, null, true];
VIEWS.settings = ['⟨i:gear⟩', 'Ustawienia', vSettingsScreen, null, null, true];

/* ================= POWIADOMIENIA (dzwonek): nowe zadania, zaległe, plan dnia, spotkanie za X min, budżet ================= */
const NT = {new: ['⟨i:plus⟩', 'Nowe zadanie'], late: ['⟨i:alert⟩', 'Zadania przeterminowane'], day: ['⟨i:sun⟩', 'Twój dzień'], mtg: ['⟨i:clock⟩', 'Spotkanie'], budget: ['⟨i:wallet⟩', 'Budżet'], react: ['⟨i:note⟩', 'Reakcje na komentarze'], trip: ['⟨i:plane⟩', 'Delegacje'], person: ['⟨i:users⟩', 'Nowa osoba']};
const notifsAll = () => (S.data.notifs ||= []);
const notifUnread = () => notifsAll().filter(n => !n.read).length;
const nOn = k => SET()['n_' + k] !== false;
const hmNow = () => { const d = new Date(); return pad(d.getHours()) + ':' + pad(d.getMinutes()); };
function inQuiet() {
  const st = SET(); if (st.nQuiet === false) return false;
  const a = st.nQuietFrom || '21:00', b = st.nQuietTo || '07:00', n = hmNow();
  return a <= b ? n >= a && n < b : n >= a || n < b;
}
let notifPending = [];
/* dodaje powiadomienie (bez duplikatów wg klucza); sygnał wysyłany zbiorczo w notifFlush() */
function notifAdd(n) {
  if (tabInactive || !S.data || !nOn(n.type)) return false;
  const L = notifsAll(); if (n.key && L.some(x => x.key === n.key)) return false;
  const x = {id: uid(), t: Date.now(), read: false, ...n}; L.unshift(x); if (L.length > 200) L.length = 200;
  notifPending.push(x); return true;
}
async function notifFlush() {
  if (!notifPending.length) return;
  const l = notifPending; notifPending = []; save(); notifBadge();
  if (S.view === 'notifs') render();
  const one = l.length === 1, title = one ? l[0].title : `TM: ${plN(l.length, 'nowe powiadomienie', 'nowe powiadomienia', 'nowych powiadomień')}`, body = one ? (l[0].body || '') : l.slice(0, 3).map(n => '• ' + n.title).join('\n');
  if (document.visibilityState === 'visible') toastAction(title, 'Pokaż', () => { S.view = 'notifs'; history.replaceState(null, '', '#notifs'); render(); }, 9000);
  if (inQuiet()) return;
  const st = SET(), sound = st.nSound !== false;
  let sent = false;
  if (hasLocalSrv() && st.nWin !== false) {
    try { const sticky = st.nSticky !== false && l.some(n => n.type === 'mtg');
      const r = await fetch('/__notify', {method: 'POST', headers: {'X-TM': '1', 'Content-Type': 'application/json'}, body: JSON.stringify({title, body: body.slice(0, 300), sound, flash: st.nFlash !== false, duration: st.nDur === 'long' ? 'long' : 'short', tone: st.nTone || 'Notification.IM', reminder: sticky})});
      sent = r.ok; if (r.ok) { const j = await r.json(); st.notifLast = {t: Date.now(), toast: !!j.toast, flashed: j.flashed || 0}; save(); } } catch { }
  }
  if (!sent && sound) notifBeep();
}
function notifBeep() { try { const c = new (window.AudioContext || window.webkitAudioContext)(), o = c.createOscillator(), g = c.createGain(); o.type = 'sine'; o.frequency.value = 880; g.gain.setValueAtTime(.0001, c.currentTime); g.gain.exponentialRampToValueAtTime(.2, c.currentTime + .02); g.gain.exponentialRampToValueAtTime(.0001, c.currentTime + .35); o.connect(g).connect(c.destination); o.start(); o.stop(c.currentTime + .4); } catch { } }
function notifBadge() {
  const n = notifUnread(), b = $('#bellCnt');
  if (b) { b.textContent = n; b.hidden = !n; }
  document.title = (n ? `(${n}) ` : '') + document.title.replace(/^\(\d+\)\s*/, '');
}
/* sprawdzanie warunków – co minutę */
function notifTick() {
  if (!S.user || !S.data || tabInactive) return;
  const st = SET(), td = TODAY(), now = Date.now(), dayAt = st.nDayAt || '08:00';
  // plan dnia + zaległe: raz dziennie, od godziny z ustawień (albo przy pierwszym otwarciu później)
  if (hmNow() >= dayAt) {
    const mt = S.data.meetings.filter(m => m.start.slice(0, 10) === td).sort((a, b) => a.start.localeCompare(b.start));
    const dueT = S.data.tasks.filter(t => t.status !== 'done' && t.due === td);
    if (mt.length || dueT.length) notifAdd({type: 'day', key: 'day:' + td, title: `Twój dzień: ${mt.length} spotk., ${dueT.length} zad. z terminem dziś`,
      body: [...mt.slice(0, 4).map(m => `${hm(m.start)} ${m.title}`), ...dueT.slice(0, 4).map(t => `☐ ${t.title}`)].join('\n'), ref: {kind: 'view', id: 'dashboard'}});
    const late = S.data.tasks.filter(t => t.status !== 'done' && t.due && t.due < td);
    const rc = (S.data.comments || []).filter(c => !noReact(c.reaction) && !c.done && c.due && c.due <= td);
    if (rc.length) notifAdd({type: 'react', key: 'react:' + td, title: `Komentarze wymagające reakcji: ${rc.length}`, body: rc.slice(0, 5).map(c => { const o = c.kind === 'task' ? S.data.tasks.find(t => t.id === c.refId) : S.data.meetings.find(m => m.id === c.refId); return `${c.reaction} (${fmt(c.due)}) · ${o ? o.title : ''}`; }).join('\n'), ref: {kind: rc[0].kind === 'task' ? 'task' : 'mtg', id: rc[0].refId}});
    const lead = +(st.tripLead ?? 3);
    tripsAll().filter(t => t.from >= td && dDiff(td, t.from) <= lead).forEach(t => { const n = dDiff(td, t.from), op = tripTasks(t).filter(x => x.status !== 'done'), c = tripConflicts(t);
      notifAdd({type: 'trip', key: `trip:${t.id}:${td}`, title: `${n === 0 ? 'Dziś wyjazd' : n === 1 ? 'Jutro wyjazd' : `Za ${n} dni wyjazd`}: ${t.place}`,
        body: [op.length ? `Do zrobienia przed wyjazdem: ${op.length}` : 'Przygotowania zakończone', ...op.slice(0, 4).map(x => `☐ ${x.title}`), c.n ? `Kolizje w dniach delegacji: ${c.n}` : ''].filter(Boolean).join('\n'), ref: {kind: 'trip', id: t.id}}); });
    if (late.length) notifAdd({type: 'late', key: 'late:' + td, title: `Zadania przeterminowane: ${late.length}`, body: late.sort((a, b) => a.due.localeCompare(b.due)).slice(0, 5).map(t => `${fmt(t.due)} · ${t.title}`).join('\n'), ref: {kind: 'late'}});
  }
  // spotkanie za X minut
  const lead = +(st.nLead || 20);
  S.data.meetings.forEach(m => { const ms = strMs(m.start), dm = (ms - now) / 60000;
    if (dm > 0 && dm <= lead) notifAdd({type: 'mtg', key: 'mtg:' + m.id + ':' + m.start, title: `Za ${Math.round(dm)} min: ${m.title}`, body: `${hm(m.start)}–${hm(m.end)}${m.location ? ' · ' + m.location : ''}${m.link ? ' · Teams' : ''}`, ref: {kind: 'mtg', id: m.id}}); });
  // budżety: 90% i przekroczenie (PRIV zł, JOB godziny)
  const chk = (kind, x, b, money) => { if (!b.budget) return; const r = b.act / b.budget, lv = r > 1 ? 'bad' : r >= .9 ? 'warn' : ''; if (!lv) return; const f = money ? zl : hh;
    notifAdd({type: 'budget', key: `budget:${kind}:${x.id}:${lv}`, title: `${lv === 'bad' ? 'Przekroczony budżet' : 'Budżet wykorzystany w ' + Math.round(r * 100) + '%'}: ${x.title}`, body: `${money ? 'Koszt' : 'Realizacja'} ${f(b.act)} z ${f(b.budget)} (${kind === 'goal' ? 'podprojekt' : 'projekt'})`, ref: {kind, id: x.id}}); };
  S.data.goals.forEach(g => { if (g.ctx === 'priv' && g.budget) chk('goal', g, goalMoney(g), true); if (g.ctx === 'job' && g.budgetH) chk('goal', g, goalHours(g), false); });
  (S.data.projects || []).forEach(p => { if (p.status === 'done') return; if (p.ctx === 'priv') chk('project', p, projMoney(p), true); else chk('project', p, projHours(p), false); });
  notifFlush();
}
function notifOpen(n) {
  n.read = true; save(); const r = n.ref || {};
  if (r.kind === 'task') { const t = S.data.tasks.find(x => x.id === r.id); if (t) { render(); return editTask(t); } }
  if (r.kind === 'mtg') { const m = S.data.meetings.find(x => x.id === r.id); if (m) { render(); return editMeeting(m); } }
  if (r.kind === 'goal') { const g = goalById(r.id); if (g && g.projectId) { S.view = 'project'; S.proj.open = g.projectId; } else if (g) { render(); return editGoal(g); } }
  if (r.kind === 'project') { S.view = 'project'; S.proj.open = r.id; }
  if (r.kind === 'late') { S.view = 'tasks'; S.tasks.mode = 'list'; S.tasks.status = 'open'; S.tasks.q = ''; }
  if (r.kind === 'view') S.view = r.id;
  if (r.kind === 'person') { const pp = personById(r.id); S.view = 'people'; history.replaceState(null, '', '#people'); render(); return pp ? personPanel(pp) : null; }
  if (r.kind === 'trip') { S.view = 'trips'; history.replaceState(null, '', '#trips'); render(); return tripPanel(tripById(r.id)); }
  history.replaceState(null, '', '#' + S.view); render();
}
function vNotifs() {
  const F = S.nf ||= {f: 'all'}, L = notifsAll().filter(n => F.f === 'all' || (F.f === 'unread' ? !n.read : n.type === F.f)), td = TODAY();
  const grp = [['Dziś', L.filter(n => iso(new Date(n.t)) === td)], ['Wcześniej', L.filter(n => iso(new Date(n.t)) !== td)]];
  const item = n => `<div class="nitem ${n.read ? '' : 'unread'} nt-${n.type}" data-action="notif-open" data-id="${n.id}"><span class="nic">${NT[n.type]?.[0] || '⟨i:bell⟩'}</span>
    <div style="flex:1;min-width:0"><div class="row" style="gap:8px"><b class="pv-t">${S.privacy ? pv(esc(n.title)) : esc(n.title)}</b><span class="spacer"></span><span class="muted" style="font-size:12px;white-space:nowrap">${ago(n.t)}</span></div>
    ${n.body ? `<div class="muted nbody">${S.privacy ? pv(esc(n.body)) : esc(n.body).replace(/\n/g, '<br>')}</div>` : ''}</div>
    <button class="iconbtn" data-action="notif-del" data-id="${n.id}" title="Usuń">⟨i:x⟩</button></div>`;
  return `<div class="row" style="margin-bottom:12px;gap:8px;flex-wrap:wrap"><div class="seg">${[['all', 'Wszystkie'], ['unread', `Nieprzeczytane (${notifUnread()})`], ...Object.entries(NT).map(([k, [, l]]) => [k, l])].map(([k, l]) => `<button data-action="set" data-k="nf.f" data-v="${k}" class="${F.f === k ? 'on' : ''}">${l}</button>`).join('')}</div>
    <div class="spacer"></div>${eyeBtn('sm')}<button class="btn sm" data-action="notif-readall">⟨i:check⟩ Oznacz wszystkie jako przeczytane</button><button class="btn sm ghost" data-action="notif-clear" title="Usuń przeczytane">⟨i:trash⟩ Wyczyść przeczytane</button>
    <button class="btn sm ghost" data-action="open-profile" data-tab="notif" title="Ustawienia powiadomień">⟨i:sliders⟩</button></div>
    ${L.length ? grp.filter(([, l]) => l.length).map(([h, l]) => `<div class="group-h">${h}</div><div class="card" style="padding:4px 0">${l.map(item).join('')}</div>`).join('') : '<div class="card empty">Brak powiadomień.</div>'}`;
}
SETTINGS_SECTIONS.push(Object.assign(() => { const st = SET(), ck = (k, l) => `<label class="row" style="gap:8px;font-size:14px"><input type="checkbox" name="n_${k}" ${st['n_' + k] !== false ? 'checked' : ''}> ${l}</label>`;
  return `<div class="card"><h3>⟨i:bell⟩ Powiadomienia</h3><div class="grid" style="gap:6px">
    ${ck('new', 'Nowe zadanie (utworzone automatycznie, np. z maila TM)')}${ck('late', 'Zadania przeterminowane (raz dziennie)')}${ck('day', 'Twój dzień – spotkania i zadania na dziś')}${ck('mtg', 'Spotkanie za kilka minut')}${ck('budget', 'Budżet: 90% i przekroczenie')}${ck('react', 'Komentarze z reakcją na dziś lub zaległą')}${ck('trip', 'Delegacja: przygotowania i kolizje przed wyjazdem')}${ck('person', 'Utworzono nową osobę (z maili lub Teams)')}</div>
    <div class="grid g2" style="margin-top:10px"><label class="f">Spotkanie – ile minut wcześniej<input class="input" type="number" min="1" max="120" name="nLead" value="${st.nLead || 20}"></label>
    <label class="f">Plan dnia – od godziny<input class="input" type="time" name="nDayAt" value="${esc(st.nDayAt || '08:00')}"></label></div>
    <div class="grid" style="gap:6px;margin-top:8px"><label class="row" style="gap:8px;font-size:14px"><input type="checkbox" name="nWin" ${st.nWin !== false ? 'checked' : ''}> Powiadomienie Windows (dymek + miganie TM na pasku zadań)</label>
    <label class="row" style="gap:8px;font-size:14px"><input type="checkbox" name="nSound" ${st.nSound !== false ? 'checked' : ''}> Dźwięk</label>
    <label class="row" style="gap:8px;font-size:14px"><input type="checkbox" name="nQuiet" ${st.nQuiet !== false ? 'checked' : ''}> Ciche godziny (bez dźwięku i dymków)</label></div>
    <div class="grid g2"><label class="f">od<input class="input" type="time" name="nQuietFrom" value="${esc(st.nQuietFrom || '21:00')}"></label><label class="f">do<input class="input" type="time" name="nQuietTo" value="${esc(st.nQuietTo || '07:00')}"></label></div>
    <div class="row" style="margin-top:6px"><button type="button" class="btn sm" data-action="notif-test">⟨i:bell⟩ Wyślij powiadomienie testowe</button></div></div>`; }, {sec: 'notif'}));
SETTINGS_SECTIONS.push(Object.assign(() => { const st = SET(), r = st.notifLast;
  return `<div class="card"><h3>⟨i:monitor⟩ Dymki Windows (PowerShell)</h3>
    <p class="muted" style="margin-top:0;font-size:13px">Powiadomienia wyświetla serwer TM (okno „TM - serwer”) jako dymek Windows w prawym dolnym rogu i w Centrum powiadomień; okno TM miga na pasku zadań do czasu kliknięcia.</p>
    <div class="grid g2"><label class="f">Czas wyświetlania dymku<select class="input" name="nDur"><option value="short" ${st.nDur !== 'long' ? 'selected' : ''}>krótko (~7 s)</option><option value="long" ${st.nDur === 'long' ? 'selected' : ''}>długo (~25 s)</option></select></label>
    <label class="f">Dźwięk dymku<select class="input" name="nTone">${[['Notification.IM', 'wiadomość (jak Teams)'], ['Notification.Default', 'domyślny Windows'], ['Notification.Reminder', 'przypomnienie'], ['Notification.Mail', 'poczta']].map(([v, l]) => `<option value="${v}" ${(st.nTone || 'Notification.IM') === v ? 'selected' : ''}>${l}</option>`).join('')}</select></label></div>
    <label class="row" style="gap:8px;font-size:14px;margin-top:4px"><input type="checkbox" name="nSticky" ${st.nSticky !== false ? 'checked' : ''}> Przypomnienie o spotkaniu zostaje na ekranie do zamknięcia</label>
    <label class="row" style="gap:8px;font-size:14px;margin-top:4px"><input type="checkbox" name="nFlash" ${st.nFlash !== false ? 'checked' : ''}> Miganie okna TM na pasku zadań</label>
    <p class="hint" style="margin:8px 0 0">${r ? `Ostatni dymek: ${ago(r.t)} · ${r.toast ? 'wyświetlony' : 'niedostępny – Windows lub zasady firmy blokują dymki z PowerShella (zostaje dźwięk w TM i licznik)'}${r.flashed ? ' · okno zamigało' : ''}` : 'Kliknij „Wyślij powiadomienie testowe”, aby sprawdzić dymki.'}</p></div>`; }, {sec: 'notif'}));
Object.assign(A, {
  'notif-open': (el, e) => { if (e && e.target.closest('[data-action=notif-del]')) return; const n = notifsAll().find(x => x.id === el.dataset.id); if (n) notifOpen(n); },
  'notif-del': el => { S.data.notifs = notifsAll().filter(x => x.id !== el.dataset.id); save(); render(); },
  'notif-readall': () => { notifsAll().forEach(n => n.read = true); save(); render(); },
  'notif-clear': () => { S.data.notifs = notifsAll().filter(n => !n.read); save(); render(); },
  'notif-test': () => { notifAdd({type: 'day', key: 'test:' + Date.now(), title: 'Powiadomienie testowe TM', body: 'Tak wyglądają powiadomienia TM w Windows.'}); const q = inQuiet(); notifFlush(); if (q) toast('Trwają ciche godziny – dymek i dźwięk są wyciszone'); }
});

/* ================= KOMENTARZE (jak w CRM) do zadań i spotkań: tekst, data, konieczna reakcja (LoV), termin i status reakcji ================= */
const REACT_DEF = ['Brak – informacja', 'Odpowiedzieć', 'Zadzwonić', 'Wysłać materiały', 'Umówić spotkanie', 'Utworzyć zadanie', 'Decyzja / akceptacja', 'Eskalacja', 'Ponaglić klienta'];
const reactList = () => { const v = String(SET().reactions || '').split('\n').map(x => x.trim()).filter(Boolean); return v.length ? v : REACT_DEF; };
const noReact = r => !r || /^brak/i.test(r);
const cmtAll = () => (S.data.comments ||= []);
const cmtOf = (kind, id) => cmtAll().filter(c => c.kind === kind && c.refId === id).sort((a, b) => (b.date || '').localeCompare(a.date || '') || b.createdAt - a.createdAt);
const cmtOpen = (kind, id) => cmtOf(kind, id).filter(c => !noReact(c.reaction) && !c.done);
const cmtKey = (kind, id) => kind + ':' + id;
/* licznik komentarzy (przycisk rozwijający subtabelę) */
function cmtBadge(kind, id) {
  const n = cmtOf(kind, id).length, o = cmtOpen(kind, id).length, late = cmtOpen(kind, id).filter(c => c.due && c.due < TODAY()).length;
  return `<button type="button" class="cmt-badge ${(S.cx ||= new Set()).has(cmtKey(kind, id)) ? 'on' : ''} ${late ? 'late' : o ? 'open' : ''}" data-action="cmt-toggle" data-kind="${kind}" data-id="${id}" title="Komentarze${o ? ` · ${o} do reakcji` : ''}">⟨i:note⟩ ${n}${o ? ` · <b>${o}!</b>` : ''}</button>`;
}
/* subtabela komentarzy + dodawanie (bez zagnieżdżonego <form> – działa także wewnątrz okna zadania/spotkania) */
function cmtBlock(kind, id) {
  const l = cmtOf(kind, id), P_ = S.privacy ? pv : (x => x), rl = reactList();
  return `<div class="cmt-block" data-kind="${kind}" data-ref="${id}">
    <table class="tbl cmt-tbl"><thead><tr><th style="width:96px">Data</th><th>Komentarz</th><th style="width:170px">Konieczna reakcja</th><th style="width:110px">Termin reakcji</th><th style="width:70px">Status</th><th style="width:30px"></th></tr></thead><tbody>
    ${l.map(c => { const need = !noReact(c.reaction), late = need && !c.done && c.due && c.due < TODAY();
      return `<tr class="${c.done ? 'cdone' : ''}"><td class="nowrap">${fmt(c.date)}</td><td class="ctext">${P_(esc(c.text).replace(/\n/g, '<br>'))}</td>
      <td>${need ? `<span class="tag react">${esc(c.reaction)}</span>` : '<span class="muted">—</span>'}</td><td class="nowrap ${late ? 'bbad' : ''}">${need && c.due ? fmt(c.due) : '—'}</td>
      <td>${need ? `<label class="row" style="gap:4px;font-size:12px;cursor:pointer"><input type="checkbox" data-action="cmt-done" data-id="${c.id}" ${c.done ? 'checked' : ''}> ${c.done ? 'zrobione' : 'otwarte'}</label>` : ''}</td>
      <td><button type="button" class="iconbtn" data-action="cmt-del" data-id="${c.id}" title="Usuń komentarz">⟨i:x⟩</button></td></tr>`; }).join('') || `<tr><td colspan="6" class="muted" style="font-size:12px">Brak komentarzy</td></tr>`}
    <tr class="cmt-add"><td><input class="input cmt-date" type="date" value="${TODAY()}"></td><td><textarea class="input cmt-text" rows="1" placeholder="Nowy komentarz… (Enter = dodaj, Shift+Enter = nowa linia)"></textarea></td>
      <td><select class="input cmt-react">${rl.map(r => `<option>${esc(r)}</option>`).join('')}</select></td><td><input class="input cmt-due" type="date" title="Termin reakcji (opcjonalnie)"></td>
      <td colspan="2"><button type="button" class="btn sm primary" data-action="cmt-add">Dodaj</button></td></tr></tbody></table></div>`;
}
function cmtRefresh(kind, id) {
  const sel = `.cmt-block[data-kind="${kind}"][data-ref="${id}"]`, inModal = document.querySelector('#modalRoot ' + sel);
  if (inModal) { inModal.outerHTML = H(cmtBlock(kind, id)); const st = document.querySelector(`#modalRoot .cmt-sec[data-kind="${kind}"][data-ref="${id}"] .cmt-cnt`); if (st) st.textContent = cmtOf(kind, id).length; }
  if (!inModal || document.querySelector('#view ' + sel) || document.querySelector('#view .cmt-badge')) render();
}
function cmtAdd(box) {
  const kind = box.dataset.kind, id = box.dataset.ref, text = $('.cmt-text', box).value.trim();
  if (!text) { $('.cmt-text', box).focus(); return toast('Wpisz treść komentarza'); }
  const reaction = $('.cmt-react', box).value, due = $('.cmt-due', box).value;
  cmtAll().push({id: uid(), kind, refId: id, text, date: $('.cmt-date', box).value || TODAY(), reaction, due: noReact(reaction) ? '' : due, done: false, createdAt: Date.now()});
  save(); cmtRefresh(kind, id); toast('Dodano komentarz');
  setTimeout(() => { const nb = document.querySelector(`#modalRoot .cmt-block[data-kind="${kind}"][data-ref="${id}"] .cmt-text`) || document.querySelector(`.cmt-block[data-kind="${kind}"][data-ref="${id}"] .cmt-text`); nb && nb.focus(); });
}
/* sekcja w oknie zadania / spotkania */
const cmtSection = (kind, id) => `<div class="cmt-sec" data-kind="${kind}" data-ref="${id}" style="grid-column:1/-1"><b>⟨i:note⟩ Komentarze (<span class="cmt-cnt">${cmtOf(kind, id).length}</span>)</b>${cmtBlock(kind, id)}</div>`;
document.addEventListener('keydown', e => { if (e.target.classList?.contains('cmt-text') && e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); cmtAdd(e.target.closest('.cmt-block')); } });
Object.assign(A, {
  'cmt-toggle': el => { const k = cmtKey(el.dataset.kind, el.dataset.id), s = (S.cx ||= new Set()); s.has(k) ? s.delete(k) : s.add(k); render(); },
  'cmt-add': el => cmtAdd(el.closest('.cmt-block')),
  'cmt-done': el => { const c = cmtAll().find(x => x.id === el.dataset.id); if (!c) return; c.done = el.checked; c.doneAt = el.checked ? TODAY() : null; save(); cmtRefresh(c.kind, c.refId); },
  'cmt-del': el => { const c = cmtAll().find(x => x.id === el.dataset.id); if (!c) return; const i = cmtAll().indexOf(c); cmtAll().splice(i, 1); save(); cmtRefresh(c.kind, c.refId);
    toastAction('Usunięto komentarz', '⟨i:undo⟩ Cofnij', () => { cmtAll().splice(i, 0, c); save(); cmtRefresh(c.kind, c.refId); }); },
  'mtg-assign': el => { const m = S.data.meetings.find(x => x.id === el.dataset.id), p = projById(el.dataset.project); if (!m || !p) return; m.projectId = p.id; m.ctx = p.ctx; save(); render(); toast(`Przypisano spotkanie do projektu ${p.title}`); }
});
SETTINGS_SECTIONS.push(() => `<div class="card"><h3>⟨i:note⟩ Komentarze · lista „Konieczna reakcja”</h3>
  <label class="f">Wartości (jedna na linię; „Brak…” = bez reakcji)<textarea class="input" name="reactions" rows="8">${esc(reactList().join('\n'))}</textarea></label></div>`);
/* spotkania bez projektu, które pasują do projektu (nazwa, klient, IC case z zadań projektu) – do przypisania jednym kliknięciem */
function mtgSuggest(p) {
  const words = [p.title, p.client].filter(Boolean).map(norm).filter(w => w.length >= 3);
  const ics = projTasks(p).map(t => t.ic).filter(Boolean).map(norm);
  if (!words.length && !ics.length) return [];
  return S.data.meetings.filter(m => !m.projectId && !m.demo && m.end.slice(0, 10) >= iso(add(today(), -30)) && (() => { const t = norm([m.title, m.location, m.body].join(' ')); return words.some(w => t.includes(w)) || ics.some(w => t.includes(w)); })())
    .sort((a, b) => a.start.localeCompare(b.start)).slice(0, 8);
}

/* ================= OSOBY: klienci, koledzy, wykonawcy… – wewnętrzni / zewnętrzni, PRIV / JOB; przypisanie do zadań (odbiorca / zlecający) ================= */
const peopleAll = () => (S.data.people ||= []);
const personById = id => id ? peopleAll().find(p => p.id === id) : null;
const pName = p => p ? `${p.first || ''} ${p.last || ''}`.trim() || p.email || '(bez nazwy)' : '';
const pShort = p => p ? `${p.first || ''} ${(p.last || '')[0] ? p.last[0] + '.' : ''}`.trim() || p.email : '';
const pInitials = p => ((p.first || '?')[0] + ((p.last || '')[0] || '')).toUpperCase();
const pTasks = (p, role) => S.data.tasks.filter(t => role === 'to' ? t.assigneeId === p.id : role === 'from' ? t.requesterId === p.id : t.assigneeId === p.id || t.requesterId === p.id);
const personByEmail = em => { em = String(em || '').trim().toLowerCase(); return em ? peopleAll().find(p => (p.email || '').toLowerCase() === em) : null; };
/* opcje selecta osób (filtrowane kontekstem zadania przez linkForm – klasa ctx-filter) */
const personOptions = (sel, ctx) => `<option value="">— brak —</option>${[...peopleAll()].sort((a, b) => pName(a).localeCompare(pName(b), 'pl')).map(p => `<option value="${p.id}" data-ctx="${p.ctx}" ${sel === p.id ? 'selected' : ''}>${esc(pName(p))}${p.kind === 'ext' ? ' (zewn.)' : ''}${p.company ? ' · ' + esc(p.company) : ''}</option>`).join('')}`;
/* etykiety na zadaniu: → odbiorca (komu zleciłem), ← zlecający (od kogo) */
function personTags(t) {
  const a = personById(t.assigneeId), r = personById(t.requesterId), P_ = S.privacy ? pv : x => x;
  return (r ? `<span class="tag ptag from" title="Zlecający: ${esc(pName(r))}">← ${P_(esc(pShort(r)))}</span>` : '') + (a ? `<span class="tag ptag to" title="Odbiorca (zlecone osobie): ${esc(pName(a))}">→ ${P_(esc(pShort(a)))}</span>` : '');
}
function vPeople() {
  const F = S.pf ||= {q: '', kind: '', sort: 'name'}, P_ = S.privacy ? pv : x => x;
  let l = peopleAll().filter(inCtx);
  if (F.kind) l = l.filter(p => p.kind === F.kind);
  if (F.auto) l = l.filter(p => p.auto);
  if (F.q) l = l.filter(p => norm([p.first, p.last, p.email, p.company, p.role, p.phone].join(' ')).includes(norm(F.q)));
  const openC = (p, r) => pTasks(p, r).filter(t => t.status !== 'done').length;
  l.sort(F.sort === 'open' ? (a, b) => (openC(b) - openC(a)) || pName(a).localeCompare(pName(b), 'pl') : (a, b) => (a.last || a.first || '').localeCompare(b.last || b.first || '', 'pl') || pName(a).localeCompare(pName(b), 'pl'));
  const rows = l.map(p => { const to = openC(p, 'to'), fr = openC(p, 'from'), lateTo = pTasks(p, 'to').filter(isLate).length;
    return `<tr class="prow" data-action="person-open" data-id="${p.id}"><td><span class="pav ${p.ctx}">${pInitials(p)}</span> <b>${P_(esc(pName(p)))}</b>${p.auto ? ` <span class="tag src" title="Utworzono automatycznie – sprawdź dane">auto · ${esc(p.auto)}</span>` : ''}${p.role ? `<div class="muted" style="font-size:12px;margin-left:36px">${esc(p.role)}</div>` : ''}</td>
      <td>${p.email ? `<a href="mailto:${esc(p.email)}" class="plink">${P_(esc(p.email))}</a>` : '<span class="muted">—</span>'}</td><td>${P_(esc(p.company || ''))}</td>
      <td><span class="tag ${p.kind === 'ext' ? 'pext' : 'pint'}">${p.kind === 'ext' ? 'zewnętrzna' : 'wewnętrzna'}</span></td><td>${ctxTag(p.ctx)}</td>
      <td class="num">${to ? `<b>${to}</b>${lateTo ? ` <span class="bbad">(${lateTo} po term.)</span>` : ''}` : '<span class="muted">0</span>'}</td><td class="num">${fr || '<span class="muted">0</span>'}</td>
      <td class="nowrap acts p-acts"><button class="rbtn sm" data-action="person-task" data-id="${p.id}" data-role="from" title="Utwórz zadanie od tej osoby">⟨i:plus⟩ Zadanie od osoby</button><button class="rbtn sm ghost" data-action="person-task" data-id="${p.id}" data-role="to" title="Zleć zadanie tej osobie">⟨i:arrow-right⟩ Zleć</button><button class="rbtn sm ghost" data-action="person-mtg" data-id="${p.id}" title="Nowe spotkanie w Outlook / Teams">⟨i:video⟩ Spotkanie</button><button class="rbtn sm ghost" data-action="person-mail" data-id="${p.id}" title="${p.email ? 'Napisz wiadomość w aplikacji Outlook' : 'Brak adresu e-mail'}" ${p.email ? '' : 'disabled'}>⟨i:mail⟩ Wiadomość</button><button class="btn sm ghost" data-action="person-edit" data-id="${p.id}" title="Edytuj">⟨i:pencil⟩</button><button class="btn sm ghost del" data-action="person-del" data-id="${p.id}" title="Usuń">⟨i:trash⟩</button></td></tr>`; }).join('');
  return `<div class="row" style="margin-bottom:12px;gap:8px;flex-wrap:wrap"><input class="input" style="max-width:300px" placeholder="Szukaj: imię, nazwisko, e-mail, firma…" data-bind="pf.q" value="${esc(F.q)}">
    <div class="seg">${[['', 'Wszyscy'], ['int', 'Wewnętrzni'], ['ext', 'Zewnętrzni']].map(([k, lb]) => `<button data-action="set" data-k="pf.kind" data-v="${k}" class="${F.kind === k ? 'on' : ''}">${lb}</button>`).join('')}</div>
    <button class="fchip ${F.auto ? 'on' : ''}" data-action="ppl-auto" title="Osoby utworzone automatycznie z maili i Teams – do sprawdzenia">Utworzone automatycznie (${peopleAll().filter(p => p.auto).length})</button>
    <div class="seg">${[['name', 'A–Z'], ['open', 'Najwięcej zadań']].map(([k, lb]) => `<button data-action="set" data-k="pf.sort" data-v="${k}" class="${F.sort === k ? 'on' : ''}">${lb}</button>`).join('')}</div>
    <div class="spacer"></div>${eyeBtn('sm')}<button class="btn primary" data-action="person-edit">⟨i:plus⟩ Osoba</button></div>
    <div class="card" style="padding:0;overflow:auto"><table class="tbl ptbl"><thead><tr><th>Imię i nazwisko</th><th>E-mail</th><th>Firma</th><th>Typ</th><th>Kontekst</th><th class="num" title="Otwarte zadania zlecone tej osobie">Zlecone jej</th><th class="num" title="Otwarte zadania, które ta osoba zleciła">Od niej</th><th></th></tr></thead>
    <tbody>${rows || `<tr><td colspan="8" class="empty">${peopleAll().length ? 'Brak osób dla tego filtra / kontekstu.' : 'Brak osób. Dodaj klientów, kolegów i wykonawców przyciskiem „+ Osoba”.'}</td></tr>`}</tbody></table></div>
    <p class="hint">Osobę przypiszesz w oknie zadania: <b>Odbiorca</b> = komu zlecasz zadanie, <b>Zlecający</b> = od kogo zadanie przyszło. Zadania z maili TM same dostają zlecającego, jeśli adres nadawcy jest na tej liście.</p>`;
}
function editPerson(p, after) {
  const isNew = !p; p ||= {first: '', last: '', email: '', phone: '', company: '', role: '', kind: 'ext', ctx: S.ctx === 'priv' ? 'priv' : 'job', notes: ''};
  modal(`<h3>${isNew ? 'Nowa osoba' : 'Edycja osoby'}</h3><form data-form="person" data-id="${p.id || ''}" class="grid g2">
    <label class="f">Imię<input class="input" name="first" value="${esc(p.first)}" required></label><label class="f">Nazwisko<input class="input" name="last" value="${esc(p.last)}"></label>
    <label class="f">E-mail<input class="input" type="email" name="email" value="${esc(p.email)}"></label><label class="f">Telefon<input class="input" name="phone" value="${esc(p.phone || '')}"></label>
    <label class="f">Firma<input class="input" name="company" value="${esc(p.company || '')}" placeholder="np. ZRE / InfoConsulting / wykonawca"></label><label class="f">Rola / opis<input class="input" name="role" value="${esc(p.role || '')}" placeholder="np. Kierownik projektu klienta, hydraulik"></label>
    <label class="f">Typ<select class="input" name="kind"><option value="int" ${p.kind === 'int' ? 'selected' : ''}>wewnętrzna (z mojej firmy / bliscy)</option><option value="ext" ${p.kind !== 'int' ? 'selected' : ''}>zewnętrzna (klient, wykonawca…)</option></select></label>
    <label class="f">Kontekst<select class="input" name="ctx"><option value="job" ${p.ctx === 'job' ? 'selected' : ''}>JOB – relacja służbowa</option><option value="priv" ${p.ctx === 'priv' ? 'selected' : ''}>PRIV – relacja prywatna</option></select></label>
    <label class="f" style="grid-column:1/-1">Notatka<textarea class="input" name="notes" rows="2">${esc(p.notes || '')}</textarea></label>
    <div class="row" style="grid-column:1/-1;justify-content:flex-end"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">Zapisz</button></div></form>`, m => { $('[name=first]', m).focus(); });
  S.personAfter = after || null;
}
function personPanel(p) {
  const P_ = S.privacy ? pv : x => x, to = pTasks(p, 'to').sort(prioSort), fr = pTasks(p, 'from').sort(prioSort), open = l => l.filter(t => t.status !== 'done'), done = l => l.filter(t => t.status === 'done');
  const sec = (h, l, empty) => `<h4 style="margin:14px 0 6px">${h} <span class="muted">(${open(l).length} otwarte${done(l).length ? `, ${done(l).length} zrobione` : ''})</span></h4>${open(l).map(t => taskRow(t, {pv: true})).join('') || `<p class="muted" style="margin:0;font-size:13px">${empty}</p>`}`;
  modal(`<div class="row" style="gap:12px;align-items:center"><span class="pav big ${p.ctx}">${pInitials(p)}</span><div style="flex:1;min-width:0"><h3 style="margin:0">${P_(esc(pName(p)))}</h3>
    <div class="muted" style="font-size:13px">${[p.role, p.company].filter(Boolean).map(esc).join(' · ')}${p.email ? ` · <a href="mailto:${esc(p.email)}">${P_(esc(p.email))}</a>` : ''}${p.phone ? ' · ' + P_(esc(p.phone)) : ''}</div></div>
    ${ctxTag(p.ctx)}<span class="tag ${p.kind === 'ext' ? 'pext' : 'pint'}">${p.kind === 'ext' ? 'zewnętrzna' : 'wewnętrzna'}</span></div>
    ${p.notes ? `<p class="muted" style="margin:10px 0 0">${P_(esc(p.notes))}</p>` : ''}
    ${sec('→ Zadania zlecone tej osobie', to, 'Brak zadań zleconych tej osobie.')}${sec('← Zadania od tej osoby', fr, 'Brak zadań od tej osoby.')}
    <div class="row" style="justify-content:flex-end;margin-top:14px"><button class="btn" data-action="person-mtg" data-id="${p.id}">⟨i:video⟩ Spotkanie</button><button class="btn" data-action="person-mail" data-id="${p.id}" ${p.email ? '' : 'disabled'}>⟨i:mail⟩ Wiadomość</button><button class="btn" data-action="person-task" data-id="${p.id}" data-role="to">⟨i:plus⟩ Zleć zadanie</button><button class="btn" data-action="person-task" data-id="${p.id}" data-role="from">⟨i:plus⟩ Zadanie od tej osoby</button><button class="btn" data-action="person-edit" data-id="${p.id}">⟨i:pencil⟩ Edytuj</button><button class="btn primary" data-action="modal-close">Zamknij</button></div>`, m => m.classList.add('wide'));
}
Object.assign(A, {
  'person-open': (el, e) => { if (e && e.target.closest('a,[data-action=person-edit],[data-action=person-del],[data-action=person-task],[data-action=person-mtg],[data-action=person-mail]')) return; const p = personById(el.dataset.id); if (p) personPanel(p); },
  'person-edit': el => editPerson(personById(el.dataset.id)),
  'person-del': el => { const p = personById(el.dataset.id); if (!p) return; const n = pTasks(p).length;
    confirmDlg('Usunąć osobę?', `<p><b>${esc(pName(p))}</b> zostanie usunięta${n ? ` i odpięta od ${n} zadań` : ''}.</p>`, 'Usuń', () => {
      const i = peopleAll().indexOf(p), ta = pTasks(p, 'to').map(t => t.id), tf = pTasks(p, 'from').map(t => t.id);
      peopleAll().splice(i, 1); S.data.tasks.forEach(t => { if (t.assigneeId === p.id) delete t.assigneeId; if (t.requesterId === p.id) delete t.requesterId; }); save(); render();
      toastAction('Usunięto osobę', '⟨i:undo⟩ Cofnij', () => { peopleAll().splice(i, 0, p); S.data.tasks.forEach(t => { if (ta.includes(t.id)) t.assigneeId = p.id; if (tf.includes(t.id)) t.requesterId = p.id; }); save(); render(); }); }); },
  'person-mail': (el, e) => { e?.stopPropagation(); const p = personById(el.dataset.id); if (!p) return; if (!p.email) return toast('Ta osoba nie ma adresu e-mail – uzupełnij go w edycji osoby');
    outlookApp({action: 'mail', to: p.email, subject: '', body: ''}, 'https://outlook.office.com/mail/deeplink/compose?to=' + encodeURIComponent(p.email), `Otworzono nową wiadomość do: ${pName(p)}`); },
  'person-mtg': (el, e) => { e?.stopPropagation(); const p = personById(el.dataset.id); if (!p) return; closeModal(); editMeeting(null, {date: TODAY()}); setTimeout(() => { const f = $('#modalRoot form[data-form=mtg]'); if (!f) return;
    if (f.title) f.title.value = `Spotkanie: ${pName(p)}${p.company ? ' (' + p.company + ')' : ''}`; if (f.att && p.email) f.att.value = p.email; if (f.ctx) { f.ctx.value = p.ctx; f.ctx.dispatchEvent(new Event('change', {bubbles: true})); } if (f.source) { f.source.value = p.ctx === 'priv' ? 'outlook' : 'teams'; f.source.dispatchEvent(new Event('change', {bubbles: true})); } f.title?.focus(); }, 30); },
  'person-task': el => { const p = personById(el.dataset.id); if (!p) return; closeModal(); editTask(null); setTimeout(() => { const f = $('#modalRoot form[data-form=task]'); if (!f) return; f.ctx.value = p.ctx; f.ctx.dispatchEvent(new Event('change', {bubbles: true})); f[el.dataset.role === 'to' ? 'assigneeId' : 'requesterId'].value = p.id; f.title.focus(); }, 30); }
});
Object.assign(FORMS, {
  person: (f, fd) => {
    const id = f.dataset.id, p = id ? personById(id) : {id: uid(), createdAt: TODAY()};
    if (fd.email && peopleAll().some(x => x !== p && (x.email || '').toLowerCase() === fd.email.trim().toLowerCase())) return toast('Osoba z tym adresem e-mail już istnieje');
    Object.assign(p, {first: fd.first.trim(), last: fd.last.trim(), email: fd.email.trim(), phone: fd.phone.trim(), company: fd.company.trim(), role: fd.role.trim(), kind: fd.kind, ctx: fd.ctx, notes: fd.notes});
    if (!id) peopleAll().push(p); save(); closeModal(); render(); toast(id ? 'Zapisano osobę' : 'Dodano osobę');
  }
});

/* ================= STRONA DOMOWA: powitanie + 5 trybów (Dziś / Projekty / Zadania / Spotkania / Realizacja) ================= */
const tripsAll = () => (S.data.trips ||= []);
VIEWS.trips = ['⟨i:plane⟩', 'Delegacje', vTrips, 'avail'];
/* --- szybkie akcje --- */
function startPicker() {
  const td = TODAY(), open = S.data.tasks.filter(t => t.status !== 'done' && t.ctx !== 'priv'), pri = [...open.filter(t => t.due && t.due <= td), ...open.filter(t => t.status === 'doing')].filter((t, i, a) => a.indexOf(t) === i).sort(prioSort);
  const rest = open.filter(t => !pri.includes(t)).sort((a, b) => (a.due || '9').localeCompare(b.due || '9')).slice(0, 40);
  const row = t => `<div class="row pick-row" data-q="${esc(norm(t.title + ' ' + (t.ic || '') + ' ' + (projById(projOf(t))?.title || '')))}"><span class="dot ${t.ctx}"></span><span class="ell" style="flex:1">${pv(esc(t.title))}${t.ic ? ' ' + icTag(t) : ''}</span><span class="muted" style="font-size:12px">${esc(projById(projOf(t))?.title || '')}</span><button class="btn sm primary" data-action="timer-start-task" data-id="${t.id}">⟨i:play⟩ Start</button></div>`;
  modal(`<h3>⟨i:timer⟩ Start realizacji</h3>${mtgPickHtml()}<p class="hint" style="margin:0 0 8px">Rejestracja czasu obejmuje tylko zadania i spotkania JOB.</p><input class="input" id="pickQ" placeholder="Szukaj zadania…" style="margin-bottom:10px">
    <div class="pick-list">${pri.length ? `<div class="group-h">Na dziś i w toku</div>${pri.map(row).join('')}` : ''}<div class="group-h">Pozostałe otwarte</div>${rest.map(row).join('')}</div>
    <div class="row" style="margin-top:12px;gap:8px;border-top:1px solid var(--line);padding-top:12px"><input class="input" id="pickNew" placeholder="…albo realizacja bez zadania, np. Spotkanie z klientem" style="flex:1"><button class="btn" data-action="dq-start-free">⟨i:play⟩ Start</button><button class="btn" data-action="modal-close">Anuluj</button></div>`,
    m => { const q = $('#pickQ', m); q.focus(); q.addEventListener('input', () => { const v = norm(q.value); $$('.pick-row', m).forEach(r => r.hidden = v && !r.dataset.q.includes(v)); }); });
}
function newVacModal() {
  modal(`<h3>⟨i:sun⟩ Nowy urlop</h3><form data-form="vacm" class="grid g2"><label class="f">Od<input class="input" type="date" name="from" value="${TODAY()}" required></label><label class="f">Do<input class="input" type="date" name="to" value="${TODAY()}" required></label>
    <label class="f" style="grid-column:1/-1">Opis<input class="input" name="note" placeholder="np. Majówka"></label><label class="row" style="gap:6px;font-size:13px"><input type="checkbox" name="half"> pół dnia (urlop jednodniowy)</label>
    <div class="row" style="grid-column:1/-1;justify-content:flex-end"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">Dodaj urlop</button></div></form>`);
}
Object.assign(A, {
  'dash-mode': el => { S.dash = el.dataset.m; S.dashDay = 0; try { localStorage.setItem('tm:dash', S.dash); } catch { } if (S.view !== 'dashboard') { S.view = 'dashboard'; history.replaceState(null, '', '#dashboard'); } render(); $('main').scrollTop = 0; },
  'dq-start': () => startPicker(),
  'dq-start-free': () => { const v = ($('#pickNew')?.value || '').trim(); if (!v) return toast('Wpisz, co realizujesz'); closeModal(); startTimer({title: v}); },
  'dq-project': () => editProject(null),
  'dq-mtg': () => editMeeting(null, {date: TODAY()}),
  'dq-person': () => editPerson(null),
  'dq-vac': () => newVacModal(),
  'dq-trip': () => editTrip(null),
  'dq-entry': () => editEntry(null),
});
Object.assign(FORMS, {
  vacm: (f, fd) => { FORMS.vac(f, fd); closeModal(); }
});

/* ================= Delegacje ================= */
/* delegacja: {id, place, from, to, dep, ret, ctx, projectId, capJob, capPriv, note, tplId}; zadania przygotowawcze: task.tripId (+ tripAt 'dep'|'ret') */
const dDiff = (a, b) => Math.round((P(b) - P(a)) / 864e5);
const tripById = id => tripsAll().find(t => t.id === id);
const tripAt = ds => tripsAll().find(t => t.from <= ds && t.to >= ds); // niezależnie od przełącznika ALL/PRIV/JOB – capacity jest jedno
const tripDefCap = ctx => { const st = SET(); return ctx === 'priv' ? (st.tripCapPriv ?? 0) : (st.tripCapJob ?? 25); };
const tripCap = (t, ctx) => ctx === 'priv' ? (t.capPriv ?? tripDefCap('priv')) : (t.capJob ?? tripDefCap('job'));
const tripDays = t => dDiff(t.from, t.to) + 1;
const tripTasks = t => S.data.tasks.filter(x => x.tripId === t.id).sort((a, b) => (a.due || '').localeCompare(b.due || '') || a.title.localeCompare(b.title));
const tripRange = t => `${fmt(t.from)}${t.to !== t.from ? ' – ' + fmt(t.to) : ''}`;
const tripLabel = t => `${tripRange(t)} · ${plN(tripDays(t), 'dzień', 'dni', 'dni')}${t.dep ? ` · wyjazd ${t.dep}` : ''}${t.ret ? ` · powrót ${t.ret}` : ''}`;
/* kolizje: prywatne zadania/spotkania i urlopy w dniach delegacji + spotkania JOB spoza projektu delegacji */
function tripConflicts(t) {
  const inR = ds => ds && ds >= t.from && ds <= t.to;
  const privT = S.data.tasks.filter(x => x.status !== 'done' && x.ctx === 'priv' && !x.tripId && inR(x.due));
  const privM = S.data.meetings.filter(m => m.ctx === 'priv' && inR(m.start.slice(0, 10)));
  const jobM = S.data.meetings.filter(m => m.ctx === 'job' && inR(m.start.slice(0, 10)) && (!t.projectId || m.projectId !== t.projectId) && !/deleg/i.test(m.title));
  const vac = (S.data.vacations || []).filter(v => v.from <= t.to && v.to >= t.from);
  const other = tripsAll().filter(x => x.id !== t.id && x.from <= t.to && x.to >= t.from);
  return {privT, privM, jobM, vac, other, n: privT.length + privM.length + vac.length + other.length, soft: jobM.length};
}
/* --- szablony delegacji --- */
const dtplAll = () => (S.data.tripTemplates ||= seedTripTemplates());
const dtplById = id => dtplAll().find(t => t.id === id);
function seedTripTemplates() {
  const I = (title, ctx, at, d, est = 0.5) => ({title, ctx, at, d, est});
  return [
    {id: uid(), name: 'Delegacja krajowa – warsztat u klienta (2 dni)', place: '', days: 2, ctx: 'job', capJob: 20, capPriv: 0, dep: '06:30', ret: '19:00', note: 'Nocleg, dojazd, materiały na warsztat', uses: 0, items: [
      I('Zarezerwować nocleg', 'job', 'dep', 10), I('Zaplanować dojazd / kupić bilety', 'job', 'dep', 7), I('Przełożyć prywatne wizyty (lekarz, serwis…) z dni wyjazdu', 'priv', 'dep', 7, 0.25),
      I('Znaleźć zastępstwo do odbioru dzieci', 'priv', 'dep', 5), I('Potwierdzić agendę i salę z klientem', 'job', 'dep', 5), I('Przygotować materiały / prezentację na warsztat', 'job', 'dep', 2, 4),
      I('Sprawdzić laptop, ładowarki, VPN i dostępy', 'job', 'dep', 1), I('Spakować się', 'priv', 'dep', 1), I('Wysłać podsumowanie po warsztacie', 'job', 'ret', 1, 1), I('Rozliczyć delegację i koszty', 'job', 'ret', 2)]},
    {id: uid(), name: 'Delegacja jednodniowa', place: '', days: 1, ctx: 'job', capJob: 25, capPriv: 0, dep: '06:00', ret: '20:00', note: '', uses: 0, items: [
      I('Zaplanować dojazd / kupić bilety', 'job', 'dep', 3), I('Znaleźć zastępstwo do odbioru dzieci', 'priv', 'dep', 3), I('Przygotować materiały na spotkanie', 'job', 'dep', 1, 2), I('Rozliczyć delegację', 'job', 'ret', 2)]},
    {id: uid(), name: 'Delegacja zagraniczna (tydzień)', place: '', days: 5, ctx: 'job', capJob: 10, capPriv: 0, dep: '', ret: '', note: 'Lot, hotel, ubezpieczenie', uses: 0, items: [
      I('Zarezerwować lot i hotel', 'job', 'dep', 21, 1), I('Sprawdzić dokumenty, ubezpieczenie, EKUZ', 'priv', 'dep', 14), I('Zorganizować opiekę / odbiór dzieci na cały tydzień', 'priv', 'dep', 14, 1),
      I('Przełożyć prywatne wizyty z tygodnia wyjazdu', 'priv', 'dep', 14, 0.25), I('Przygotować materiały', 'job', 'dep', 3, 6), I('Waluta / karta płatnicza', 'priv', 'dep', 3),
      I('Ustawić zastępstwo w pracy i autoodpowiedź', 'job', 'dep', 1), I('Spakować się', 'priv', 'dep', 1, 1), I('Rozliczyć delegację i koszty', 'job', 'ret', 3, 1)]}
  ];
}
/* termin zadania przygotowawczego: dni przed wyjazdem / po powrocie; JOB tylko w dni robocze; nie wcześniej niż dziś */
function tripItemDue(trip, it) {
  let d = it.at === 'ret' ? add(P(trip.to), +it.d || 0) : add(P(trip.from), -(+it.d || 0)), n = 0;
  if (it.ctx === 'job') while (!isWork(d) && n++ < 10) d = add(d, it.at === 'ret' ? 1 : -1);
  return iso(d) < TODAY() ? TODAY() : iso(d);
}
function tripGenTasks(trip, items) {
  return items.map(it => { const due = tripItemDue(trip, it);
    const t = {id: uid(), createdAt: TODAY(), doneAt: null, title: it.title, ctx: it.ctx, goalId: null, projectId: trip.projectId && projById(trip.projectId)?.ctx === it.ctx ? trip.projectId : null,
      start: due, due, status: 'todo', prio: it.prio || 2, est: +it.est || 0, notes: `Delegacja: ${trip.place} (${tripRange(trip)})`, tripId: trip.id, tripAt: it.at === 'ret' ? 'ret' : 'dep'};
    S.data.tasks.push(t); return t; });
}
/* --- okno delegacji (szczegóły) --- */
function tripPanel(t) {
  if (!t) return;
  const ts = tripTasks(t), done = ts.filter(x => x.status === 'done').length, c = tripConflicts(t), p = projById(t.projectId), dTo = dDiff(TODAY(), t.from);
  const when = dTo > 1 ? `za ${dTo} dni` : dTo === 1 ? 'jutro' : dTo === 0 ? 'dziś' : t.to >= TODAY() ? 'w trakcie' : 'zakończona';
  const crow = (ic, txt, act, id, d) => `<div class="row trip-cf" ${act ? `data-action="${act}" data-id="${id}"` : ''}><span class="nowrap muted">${d ? fmt(d) : ''}</span>${ic} <span class="ell" style="flex:1">${txt}</span></div>`;
  modal(`<h3 data-trip-panel="${t.id}">⟨i:plane⟩ Delegacja: ${pv(esc(t.place))}</h3>
    <div class="row" style="gap:8px;flex-wrap:wrap;margin:-4px 0 12px">${ctxTag(t.ctx)}<b>${tripLabel(t)}</b><span class="tag src">${when}</span>${p ? `<a href="#" class="tag src" data-action="open-project" data-id="${p.id}">⟨i:folder⟩ ${pv(esc(p.title))}</a>` : ''}</div>
    ${t.note ? `<p class="muted" style="margin:0 0 12px">${pv(esc(t.note))}</p>` : ''}
    <div class="trip-capline">⟨i:gauge⟩ W dniach delegacji capacity: JOB <b>${tripCap(t, 'job')}%</b> · PRIV <b>${tripCap(t, 'priv')}%</b> normalnej dostępności</div>
    <h4 class="row" style="margin:14px 0 6px;gap:8px">⟨i:check-square⟩ Przygotowanie <span class="muted" style="font-weight:500">${done}/${ts.length}</span><span class="spacer"></span><button class="btn sm ghost" data-action="trip-task-add" data-id="${t.id}">⟨i:plus⟩ Zadanie</button></h4>
    ${ts.length ? `<div class="bar" style="margin-bottom:6px"><i style="width:${Math.round(100 * done / ts.length)}%;background:var(--brand)"></i></div><div class="trip-tasks">${ts.map(x => taskRow(x, {noProj: true, noTrip: true})).join('')}</div>` : '<div class="empty" style="padding:8px">Brak zadań przygotowawczych. Dodaj je ręcznie albo utwórz delegację z szablonu.</div>'}
    <h4 style="margin:14px 0 6px">⟨i:alert⟩ Kolizje w dniach delegacji ${c.n ? `<span class="tag bbad">${c.n}</span>` : '<span class="tag src">brak</span>'}${c.soft ? ` <span class="tag src" title="Spotkania firmowe w dniach wyjazdu – dołącz zdalnie albo przełóż">${c.soft} spotk. JOB do sprawdzenia</span>` : ''}</h4>
    ${c.n + c.soft ? `<div class="trip-cfs">${c.privT.map(x => crow('⟨i:check-square⟩', `${pv(esc(x.title))} <span class="muted">· PRIV – przełóż</span>`, 'edit-task', x.id, x.due)).join('')}
      ${c.privM.map(m => crow('⟨i:video⟩', `${hm(m.start)} ${pv(esc(m.title))} <span class="muted">· PRIV – przełóż</span>`, 'edit-mtg', m.id, m.start.slice(0, 10))).join('')}
      ${c.jobM.map(m => crow('⟨i:video⟩', `${hm(m.start)} ${pv(esc(m.title))} <span class="muted">· JOB – zdalnie albo przełóż</span>`, 'edit-mtg', m.id, m.start.slice(0, 10))).join('')}
      ${c.vac.map(v => crow('⟨i:sun⟩', `Urlop ${fmt(v.from)}–${fmt(v.to)} ${esc(v.note || '')}`, '', '', '')).join('')}${c.other.map(o => crow('⟨i:plane⟩', `Inna delegacja: ${pv(esc(o.place))} (${tripRange(o)})`, 'trip-open', o.id, '')).join('')}</div>` : '<p class="muted" style="margin:0">Nic prywatnego nie jest zaplanowane w tych dniach.</p>'}
    <div class="row" style="margin-top:16px;justify-content:space-between;flex-wrap:wrap;gap:6px"><div class="row" style="gap:6px"><button class="btn danger" data-action="trip-del" data-id="${t.id}">Usuń</button><button class="btn ghost" data-action="trip-to-tpl" data-id="${t.id}" title="Zapisz miejsce, długość, capacity i listę przygotowań jako szablon">⟨i:template⟩ Zapisz jako szablon</button></div>
      <div class="row" style="gap:6px"><button class="btn" data-action="trip-outlook" data-id="${t.id}" title="Zablokuj dni w kalendarzu Outlook (całodniowo, poza biurem)">⟨i:calendar⟩ Zablokuj w Outlook</button><button class="btn" data-action="trip-edit" data-id="${t.id}">⟨i:pencil⟩ Edytuj</button><button class="btn primary" data-action="modal-close">Zamknij</button></div></div>`);
}
/* --- formularz delegacji --- */
function editTrip(t, preset = {}) {
  const isNew = !t, ctx0 = preset.projectId ? projById(preset.projectId)?.ctx : null;
  t ||= {place: '', from: preset.from || TODAY(), to: preset.from || TODAY(), ctx: ctx0 || (S.ctx === 'priv' ? 'priv' : 'job'), note: '', projectId: preset.projectId || '', dep: '', ret: '', capJob: tripDefCap('job'), capPriv: tripDefCap('priv')};
  const projOpts = c => `<option value="">— bez projektu —</option>${(S.data.projects || []).filter(p => p.ctx === c && (p.status !== 'done' || p.id === t.projectId)).map(p => `<option value="${p.id}" ${p.id === t.projectId ? 'selected' : ''}>${esc(p.title)}</option>`).join('')}`;
  const tps = dtplAll(), open = isNew ? 0 : tripTasks(t).filter(x => x.status !== 'done').length;
  modal(`<h3>⟨i:plane⟩ ${isNew ? 'Nowa delegacja' : 'Edycja delegacji'}</h3><form data-form="trip" data-id="${t.id || ''}" class="grid g2">
    ${isNew ? `<label class="f" style="grid-column:1/-1">Szablon delegacji<select class="input" name="tpl" id="tripTpl"><option value="">— bez szablonu —</option>${tps.map(x => `<option value="${x.id}" ${preset.tpl === x.id ? 'selected' : ''}>${esc(x.name)} · ${plN(x.days, 'dzień', 'dni', 'dni')} · ${x.items.length} przygotowań</option>`).join('')}</select></label>
      <div id="tripItems" style="grid-column:1/-1" hidden></div>` : ''}
    <label class="f" style="grid-column:1/-1">Dokąd / podprojekt<input class="input" name="place" value="${esc(t.place)}" required placeholder="np. Gdańsk – warsztat u klienta ZRE"></label>
    <label class="f">Od<input class="input" type="date" name="from" value="${t.from}" required></label><label class="f">Do<input class="input" type="date" name="to" value="${t.to}" required></label>
    <label class="f">Godzina wyjazdu<input class="input" type="time" name="dep" value="${t.dep || ''}"></label><label class="f">Godzina powrotu<input class="input" type="time" name="ret" value="${t.ret || ''}"></label>
    <label class="f">Kontekst<select class="input" name="ctx"><option value="job" ${t.ctx === 'job' ? 'selected' : ''}>JOB – delegacja służbowa</option><option value="priv" ${t.ctx === 'priv' ? 'selected' : ''}>PRIV – wyjazd prywatny</option></select></label>
    <label class="f">Projekt<select class="input" name="projectId" id="tripProj">${projOpts(t.ctx)}</select></label>
    <label class="f" title="Ile procent normalnego czasu JOB zostaje w dniach delegacji na inne zadania">Capacity JOB w tych dniach (%)<input class="input" type="number" min="0" max="100" step="5" name="capJob" value="${tripCap(t, 'job')}"></label>
    <label class="f" title="Ile procent czasu prywatnego zostaje (zwykle 0 – jesteś poza domem)">Capacity PRIV w tych dniach (%)<input class="input" type="number" min="0" max="100" step="5" name="capPriv" value="${tripCap(t, 'priv')}"></label>
    <label class="f" style="grid-column:1/-1">Notatka<input class="input" name="note" value="${esc(t.note || '')}" placeholder="transport, nocleg, kontakt na miejscu…"></label>
    ${isNew ? `<label class="row" style="grid-column:1/-1;gap:8px;font-size:13px"><input type="checkbox" name="outlook" ${hasLocalSrv() ? 'checked' : ''}> Zablokuj dni w kalendarzu Outlook (całodniowo, „Poza biurem”)</label>`
      : open ? `<label class="row" style="grid-column:1/-1;gap:8px;font-size:13px"><input type="checkbox" name="shift" checked> Przy zmianie dat przesuń terminy otwartych zadań przygotowawczych (${open})</label>` : ''}
    <div class="row" style="grid-column:1/-1;justify-content:flex-end"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">${isNew ? 'Utwórz delegację' : 'Zapisz'}</button></div></form>`, m => {
    const f = $('form[data-form=trip]', m); f.place.focus();
    f.ctx.addEventListener('change', () => { f.projectId.innerHTML = projOpts(f.ctx.value); });
    const sel = $('#tripTpl', m), box = $('#tripItems', m); if (!sel) return;
    const preview = () => { const x = dtplById(sel.value); if (!x) { box.hidden = true; box.innerHTML = ''; return; }
      const tr = {from: f.from.value, to: f.to.value};
      box.hidden = false; box.innerHTML = H(`<div class="trip-prev"><b>Zadania przygotowawcze z szablonu</b> <span class="muted">– odznacz zbędne</span>${x.items.map((it, i) => `<label class="row"><input type="checkbox" name="it_${i}" checked><span class="dot ${it.ctx}"></span><span style="flex:1">${esc(it.title)}</span><span class="muted nowrap">${it.at === 'ret' ? `${it.d} dni po powrocie` : `${it.d} dni przed`} · ${fmt(tripItemDue(tr, it))}</span></label>`).join('')}</div>`); };
    const fill = () => { const x = dtplById(sel.value); if (x) { if (!f.place.value || f.place.dataset.auto) { f.place.value = x.place || ''; f.place.dataset.auto = x.place ? '1' : ''; }
        f.to.value = iso(add(P(f.from.value || TODAY()), Math.max(1, x.days) - 1)); f.ctx.value = x.ctx; f.projectId.innerHTML = projOpts(x.ctx); f.capJob.value = x.capJob ?? tripDefCap('job'); f.capPriv.value = x.capPriv ?? tripDefCap('priv');
        f.dep.value = x.dep || ''; f.ret.value = x.ret || ''; if (!f.note.value) f.note.value = x.note || ''; } preview(); };
    sel.addEventListener('change', fill); f.from.addEventListener('change', () => { const x = dtplById(sel.value); if (x) f.to.value = iso(add(P(f.from.value), Math.max(1, x.days) - 1)); preview(); }); f.to.addEventListener('change', preview);
    if (sel.value) fill();
  });
}
function tripOutlook(t) {
  return outlookApp({action: 'meeting', subject: `Delegacja: ${t.place}`, start: `${t.from}T00:00`, end: `${iso(add(P(t.to), 1))}T00:00`, location: t.place, oof: true,
    body: [t.dep && `Wyjazd ${t.dep}`, t.ret && `Powrót ${t.ret}`, t.note].filter(Boolean).join('\n')}, null, 'Otworzono termin w Outlook – zapisz go (całodniowo, Poza biurem)');
}
/* --- zapis jako szablon --- */
function tripToTpl(t) {
  const ts = tripTasks(t);
  modal(`<h3>⟨i:template⟩ Zapisz delegację jako szablon</h3><form data-form="trip2tpl" data-id="${t.id}" class="grid">
    <label class="f">Nazwa szablonu<input class="input" name="name" required value="${esc('Delegacja – ' + t.place)}"></label>
    <p class="muted" style="margin:0">Szablon zapamięta: miejsce, długość (${plN(tripDays(t), 'dzień', 'dni', 'dni')}), godziny wyjazdu i powrotu, capacity (JOB ${tripCap(t, 'job')}%, PRIV ${tripCap(t, 'priv')}%), notatkę oraz ${plN(ts.length, 'zadanie przygotowawcze', 'zadania przygotowawcze', 'zadań przygotowawczych')} z terminami liczonymi względem wyjazdu i powrotu. Projekt i daty nie są zapisywane.</p>
    <label class="row" style="gap:8px;font-size:13px"><input type="checkbox" name="keepPlace" checked> zapisz też miejsce („${pv(esc(t.place))}”)</label>
    <div class="row" style="justify-content:flex-end"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">Zapisz szablon</button></div></form>`, m => $('[name=name]', m).select());
}
/* --- edytor szablonu delegacji --- */
function dtplRow(it = {title: '', ctx: 'job', at: 'dep', d: 1, est: 0.5}) {
  return `<div class="dt-row"><input class="input" data-k="title" value="${esc(it.title)}" placeholder="Co trzeba zrobić…">
    <select class="input" data-k="ctx"><option value="job" ${it.ctx === 'job' ? 'selected' : ''}>JOB</option><option value="priv" ${it.ctx === 'priv' ? 'selected' : ''}>PRIV</option></select>
    <input class="input" type="number" min="0" max="90" data-k="d" value="${+it.d || 0}" title="Liczba dni">
    <select class="input" data-k="at"><option value="dep" ${it.at !== 'ret' ? 'selected' : ''}>dni przed wyjazdem</option><option value="ret" ${it.at === 'ret' ? 'selected' : ''}>dni po powrocie</option></select>
    <input class="input" type="number" min="0" step="0.25" data-k="est" value="${it.est || ''}" title="Estymacja (h)" placeholder="h">
    <button type="button" class="iconbtn" data-action="dtpl-row-del" title="Usuń">⟨i:trash⟩</button></div>`;
}
function dtplEditor(x) {
  const isNew = !x; x ||= {name: '', place: '', days: 2, ctx: 'job', capJob: tripDefCap('job'), capPriv: tripDefCap('priv'), dep: '', ret: '', note: '', items: [{title: '', ctx: 'job', at: 'dep', d: 3, est: 0.5}]};
  modal(`<h3>⟨i:plane⟩ ${isNew ? 'Nowy szablon delegacji' : 'Szablon delegacji'}</h3><form data-form="dtpl" data-id="${x.id || ''}" class="grid g2 dt-form">
    <label class="f" style="grid-column:1/-1">Nazwa<input class="input" name="name" required value="${esc(x.name)}" placeholder="np. Warsztat u klienta – 2 dni"></label>
    <label class="f">Miejsce (opcjonalnie)<input class="input" name="place" value="${esc(x.place || '')}"></label><label class="f">Długość (dni)<input class="input" type="number" min="1" max="60" name="days" value="${x.days}"></label>
    <label class="f">Godzina wyjazdu<input class="input" type="time" name="dep" value="${x.dep || ''}"></label><label class="f">Godzina powrotu<input class="input" type="time" name="ret" value="${x.ret || ''}"></label>
    <label class="f">Kontekst<select class="input" name="ctx"><option value="job" ${x.ctx === 'job' ? 'selected' : ''}>JOB</option><option value="priv" ${x.ctx === 'priv' ? 'selected' : ''}>PRIV</option></select></label>
    <label class="f">Notatka<input class="input" name="note" value="${esc(x.note || '')}"></label>
    <label class="f">Capacity JOB w dniach delegacji (%)<input class="input" type="number" min="0" max="100" step="5" name="capJob" value="${x.capJob ?? tripDefCap('job')}"></label>
    <label class="f">Capacity PRIV w dniach delegacji (%)<input class="input" type="number" min="0" max="100" step="5" name="capPriv" value="${x.capPriv ?? tripDefCap('priv')}"></label>
    <div style="grid-column:1/-1"><div class="group-h" style="margin-top:4px">Zadania przygotowawcze i po powrocie</div><div id="dtRows">${x.items.map(dtplRow).join('')}</div>
      <button type="button" class="btn sm" data-action="dtpl-row-add" style="margin-top:6px">⟨i:plus⟩ Dodaj pozycję</button></div>
    <div class="row" style="grid-column:1/-1;justify-content:flex-end"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">Zapisz szablon</button></div></form>`, m => $('[name=name]', m).focus());
}
function vTripTemplates() {
  const st = SET(), l = dtplAll().filter(inCtx);
  const card = x => `<div class="card tplcard ${x.ctx}"><div class="row" style="gap:8px"><b class="tpl-name" data-action="dtpl-edit" data-id="${x.id}">${esc(x.name)}</b><div class="spacer"></div>${ctxTag(x.ctx)}</div>
    <div class="row ptpl-stats"><span>⟨i:calendar⟩ ${plN(x.days, 'dzień', 'dni', 'dni')}</span><span>⟨i:check-square⟩ ${x.items.length} przygotowań</span><span>⟨i:gauge⟩ JOB ${x.capJob ?? tripDefCap('job')}% · PRIV ${x.capPriv ?? tripDefCap('priv')}%</span>${x.place ? `<span>⟨i:pin⟩ ${esc(x.place)}</span>` : ''}${x.uses ? `<span class="muted">użyty ${x.uses}×</span>` : ''}</div>
    <details class="ptpl-goals"><summary>Lista przygotowań</summary><ul>${x.items.map(it => `<li><span class="dot ${it.ctx}"></span> ${esc(it.title)} <span class="muted">· ${it.d} ${it.at === 'ret' ? 'dni po powrocie' : 'dni przed wyjazdem'}${it.est ? ` · ${it.est}h` : ''}</span></li>`).join('')}</ul></details>
    <div class="row" style="margin-top:10px;gap:6px"><button class="btn sm primary" data-action="dtpl-use" data-id="${x.id}">⟨i:plus⟩ Utwórz delegację</button><button class="btn sm" data-action="dtpl-edit" data-id="${x.id}">⟨i:pencil⟩ Edytuj</button><button class="btn sm" data-action="dtpl-dup" data-id="${x.id}">⟨i:template⟩ Duplikuj</button><div class="spacer"></div><button class="iconbtn" data-action="dtpl-del" data-id="${x.id}" title="Usuń szablon">⟨i:trash⟩</button></div></div>`;
  return `<div class="card" style="margin-bottom:14px"><h3>⟨i:sliders⟩ Domyślne ustawienia delegacji</h3><form data-form="dtplset" class="row" style="gap:14px;flex-wrap:wrap;align-items:flex-end">
      <label class="f">Capacity JOB w dniach delegacji (%)<input class="input" type="number" min="0" max="100" step="5" name="tripCapJob" value="${st.tripCapJob ?? 25}"></label>
      <label class="f">Capacity PRIV w dniach delegacji (%)<input class="input" type="number" min="0" max="100" step="5" name="tripCapPriv" value="${st.tripCapPriv ?? 0}"></label>
      <label class="f">Przypomnienie przed wyjazdem (dni)<input class="input" type="number" min="0" max="30" name="tripLead" value="${st.tripLead ?? 3}"></label><button class="btn">Zapisz</button></form></div>
    <div class="row" style="margin-bottom:14px"><p class="muted" style="margin:0;flex:1;min-width:240px">Szablon delegacji podpowiada długość, godziny, capacity i listę rzeczy do zrobienia przed wyjazdem i po powrocie. Najprościej: zrób jedną delegację ręcznie, a potem w jej oknie kliknij „Zapisz jako szablon”.</p>
    <button class="btn" data-action="dtpl-edit">⟨i:plus⟩ Nowy szablon</button></div>
    ${l.length ? `<div class="grid g2 ptpl-grid">${l.map(card).join('')}</div>` : '<div class="card empty">Brak szablonów delegacji w tym kontekście.</div>'}`;
}
/* --- widok modułu --- */
function tripCard(t) {
  const ts = tripTasks(t), done = ts.filter(x => x.status === 'done').length, c = tripConflicts(t), p = projById(t.projectId), dTo = dDiff(TODAY(), t.from), nxt = ts.find(x => x.status !== 'done');
  const when = dTo > 1 ? `za ${dTo} dni` : dTo === 1 ? 'jutro' : dTo === 0 ? 'dziś' : 'w trakcie', pct = ts.length ? Math.round(100 * done / ts.length) : 0;
  return `<div class="card tripcard ${t.ctx}" data-action="trip-open" data-id="${t.id}"><div class="row" style="gap:8px"><span class="tc-plane">⟨i:plane⟩</span><b class="ell" style="flex:1">${pv(esc(t.place))}</b><span class="tag src">${when}</span>${ctxTag(t.ctx)}</div>
    <div class="muted" style="font-size:12px;margin-top:4px">${tripLabel(t)}${p ? ` · ⟨i:folder⟩ ${pv(esc(p.title))}` : ''}</div>
    <div class="row" style="justify-content:space-between;margin-top:10px;font-size:13px"><span>Przygotowanie</span><b>${done}/${ts.length}</b></div>
    <div class="bar" style="margin-top:4px"><i style="width:${pct}%;background:${pct === 100 ? '#059669' : 'var(--brand)'}"></i></div>
    <div class="pc-meta">${nxt ? `<span class="${isLate(nxt) ? 'bbad' : ''}">☐ ${pv(esc(nxt.title))} · ${relDay(nxt.due)}</span>` : ts.length ? '<span>✓ wszystko gotowe</span>' : '<span class="muted">brak listy przygotowań</span>'}
      ${c.n ? `<span class="bbad">⟨i:alert⟩ ${plN(c.n, 'kolizja', 'kolizje', 'kolizji')}</span>` : ''}${c.soft ? `<span class="muted">⟨i:video⟩ ${c.soft} spotk. JOB</span>` : ''}<span class="muted">capacity JOB ${tripCap(t, 'job')}%</span></div></div>`;
}
function vTrips() {
  const td = TODAY(), all = tripsAll().filter(inCtx), up = all.filter(t => t.to >= td).sort((a, b) => a.from.localeCompare(b.from)), past = all.filter(t => t.to < td).sort((a, b) => b.from.localeCompare(a.from));
  const y = String(today().getFullYear()), wd = t => { let n = 0; for (let d = P(t.from); iso(d) <= t.to; d = add(d, 1)) if (iso(d).startsWith(y) && isWork(d)) n++; return n; };
  const openPrep = up.reduce((a, t) => a + tripTasks(t).filter(x => x.status !== 'done').length, 0), cfl = up.reduce((a, t) => a + tripConflicts(t).n, 0);
  return `<div class="grid g4 kpis" style="margin-bottom:16px">
    <div class="card kpi"><div class="v">${up.length}</div><div class="l">Delegacje zaplanowane</div></div>
    <div class="card kpi"><div class="v">${all.reduce((a, t) => a + wd(t), 0)}</div><div class="l">Dni roboczych w delegacji (${y})</div></div>
    <div class="card kpi"><div class="v">${openPrep}</div><div class="l">Otwarte przygotowania</div></div>
    <div class="card kpi"><div class="v" style="color:${cfl ? 'var(--bad)' : 'inherit'}">${cfl}</div><div class="l">Kolizje do rozwiązania</div></div></div>
   <div class="row" style="margin-bottom:12px;gap:8px">${eyeBtn('sm')}<div class="spacer"></div><button class="btn" data-action="trip-tpls">⟨i:template⟩ Szablony delegacji</button><button class="btn primary" data-action="dq-trip">⟨i:plus⟩ Nowa delegacja</button></div>
   ${up.length ? `<div class="pcards">${up.map(tripCard).join('')}</div>` : '<div class="card empty">Brak zaplanowanych delegacji. Kliknij „Nowa delegacja” i wybierz szablon – lista przygotowań, blokada kalendarza i capacity ustawią się same.</div>'}
   ${past.length ? `<div class="card" style="margin-top:16px;padding:0"><div class="row card-h"><h3 style="margin:0">⟨i:list⟩ Minione delegacje</h3><span class="muted">${past.length}</span></div>
     <table class="tbl"><thead><tr><th>Termin</th><th>Podprojekt</th><th>Projekt</th><th class="num">Dni</th><th class="num">Zadania</th><th></th></tr></thead><tbody>${past.map(t => { const ts = tripTasks(t), p = projById(t.projectId);
       return `<tr><td class="nowrap">${tripRange(t)}</td><td class="pt" data-action="trip-open" data-id="${t.id}">${ctxIcon(t.ctx)}${pv(esc(t.place))}</td><td>${p ? pv(esc(p.title)) : '<span class="muted">—</span>'}</td><td class="num">${tripDays(t)}</td><td class="num">${ts.filter(x => x.status === 'done').length}/${ts.length}</td>
       <td class="nowrap acts"><button class="btn sm ghost" data-action="trip-to-tpl" data-id="${t.id}" title="Zapisz jako szablon">⟨i:template⟩</button><button class="btn sm ghost" data-action="trip-dup" data-id="${t.id}" title="Zaplanuj podobną delegację">⟨i:plus⟩</button></td></tr>`; }).join('')}</tbody></table></div>` : ''}`;
}
/* sekcja w projekcie */
function projTripsSection(p) {
  const l = tripsAll().filter(t => t.projectId === p.id).sort((a, b) => b.from.localeCompare(a.from));
  return `<div class="card" id="psTrips" style="margin-top:16px;padding:0"><div class="row card-h"><h3 style="margin:0">⟨i:plane⟩ Delegacje w projekcie</h3><span class="muted">${l.length} · ${l.reduce((a, t) => a + tripDays(t), 0)} dni</span><div class="spacer"></div><button class="btn sm" data-action="trip-new-proj" data-project="${p.id}">⟨i:plus⟩ Delegacja</button></div>
    <table class="tbl"><thead><tr><th>Termin</th><th>Podprojekt</th><th class="num">Dni</th><th>Przygotowanie</th><th></th></tr></thead><tbody>${l.map(t => { const ts = tripTasks(t), dn = ts.filter(x => x.status === 'done').length, c = tripConflicts(t);
      return `<tr><td class="nowrap"><b>${tripRange(t)}</b>${t.to < TODAY() ? ' <span class="muted">minęła</span>' : ''}</td><td class="pt" data-action="trip-open" data-id="${t.id}">${pv(esc(t.place))}${c.n && t.to >= TODAY() ? ` <span class="tag bbad">⟨i:alert⟩ ${c.n}</span>` : ''}</td><td class="num">${tripDays(t)}</td><td>${dn}/${ts.length}</td>
      <td class="nowrap acts"><button class="btn sm ghost" data-action="trip-open" data-id="${t.id}" title="Szczegóły">⟨i:eye⟩</button><button class="btn sm ghost" data-action="trip-edit" data-id="${t.id}" title="Edytuj">⟨i:pencil⟩</button></td></tr>`; }).join('') || '<tr><td colspan="5" class="empty">Brak delegacji w tym projekcie.</td></tr>'}</tbody></table></div>`;
}
function tripDelete(t) {
  const ts = tripTasks(t), open = ts.filter(x => x.status !== 'done');
  const go = delTasks => { const i = tripsAll().indexOf(t); tripsAll().splice(i, 1); const rm = delTasks ? open : [];
    S.data.tasks = S.data.tasks.filter(x => !rm.includes(x)); ts.forEach(x => { if (!rm.includes(x)) delete x.tripId; });
    save(); closeModal(); render(); toastAction('Usunięto delegację', '⟨i:undo⟩ Cofnij', () => { tripsAll().splice(i, 0, t); S.data.tasks.push(...rm); ts.forEach(x => x.tripId = t.id); save(); render(); }); };
  if (!open.length) return go(false);
  confirmDlg('Usunąć delegację?', `<p>Delegacja „${esc(t.place)}” ma ${plN(open.length, 'otwarte zadanie', 'otwarte zadania', 'otwartych zadań')} przygotowawcze. Usunąć je razem z delegacją?</p>`, 'Usuń z zadaniami', () => go(true));
}
Object.assign(A, {
  'trip-open': el => { closeModal(); tripPanel(tripById(el.dataset.id)); },
  'trip-edit': el => { closeModal(); editTrip(tripById(el.dataset.id)); },
  'trip-del': el => { const t = tripById(el.dataset.id); if (t) tripDelete(t); },
  'trip-outlook': el => { const t = tripById(el.dataset.id); if (t) { tripOutlook(t); t.booked = true; save(); } },
  'trip-to-tpl': el => { const t = tripById(el.dataset.id); if (t) { closeModal(); tripToTpl(t); } },
  'trip-dup': el => { const t = tripById(el.dataset.id); if (!t) return; editTrip(null, {projectId: t.projectId}); const f = $('form[data-form=trip]'); if (f) { f.place.value = t.place; f.ctx.value = t.ctx; f.capJob.value = tripCap(t, 'job'); f.capPriv.value = tripCap(t, 'priv'); f.dep.value = t.dep || ''; f.ret.value = t.ret || ''; f.note.value = t.note || ''; } },
  'trip-new-proj': el => editTrip(null, {projectId: el.dataset.project}),
  'trip-tpls': () => { S.view = 'settings'; S.setTab = 'dtpl'; history.replaceState(null, '', '#settings'); render(); },
  'trip-task-add': el => { const t = tripById(el.dataset.id); if (!t) return; closeModal(); const d = iso(add(P(t.from), -1)) < TODAY() ? TODAY() : iso(add(P(t.from), -1));
    editTask(null, {project: t.projectId || undefined, due: d}); const f = $('#modalRoot form[data-form=task]') || $('form[data-form=task]'); if (!f) return;
    f.dataset.trip = t.id; if (!t.projectId && f.ctx) { f.ctx.value = t.ctx; f.ctx.dispatchEvent(new Event('change', {bubbles: true})); } if (f.notes) f.notes.value = `Delegacja: ${t.place} (${tripRange(t)})`; },
  'dtpl-edit': el => dtplEditor(dtplById(el.dataset.id)),
  'dtpl-use': el => editTrip(null, {tpl: el.dataset.id}),
  'dtpl-dup': el => { const x = dtplById(el.dataset.id); if (!x) return; const c = JSON.parse(JSON.stringify(x)); c.id = uid(); c.name += ' (kopia)'; c.uses = 0; dtplAll().push(c); save(); render(); toast('Zduplikowano szablon'); },
  'dtpl-del': el => { const x = dtplById(el.dataset.id); if (!x) return; const i = dtplAll().indexOf(x); dtplAll().splice(i, 1); save(); render(); toastAction('Usunięto szablon delegacji', '⟨i:undo⟩ Cofnij', () => { dtplAll().splice(i, 0, x); save(); render(); }); },
  'dtpl-row-add': () => { const box = $('#dtRows'); box.insertAdjacentHTML('beforeend', H(dtplRow())); box.lastElementChild.querySelector('[data-k=title]').focus(); },
  'dtpl-row-del': el => el.closest('.dt-row').remove()
});
{ const tg = A.toggle; A.toggle = (el, ...r) => { const res = tg(el, ...r), tp = $('#modalRoot [data-trip-panel]'); if (tp) tripPanel(tripById(tp.dataset.tripPanel)); return res; }; }
Object.assign(FORMS, {
  trip: (f, fd) => {
    if (fd.to < fd.from) return toast('Data „do” jest wcześniejsza niż „od”');
    const id = f.dataset.id, t = id ? tripById(id) : {id: uid(), createdAt: TODAY()}, old = id ? {from: t.from, to: t.to} : null;
    const pj = projById(fd.projectId);
    Object.assign(t, {place: fd.place.trim(), from: fd.from, to: fd.to, dep: fd.dep || '', ret: fd.ret || '', ctx: pj ? pj.ctx : fd.ctx, projectId: pj ? pj.id : null,
      capJob: Math.max(0, Math.min(100, +fd.capJob || 0)), capPriv: Math.max(0, Math.min(100, +fd.capPriv || 0)), note: fd.note || ''});
    let made = [];
    if (!id) { tripsAll().push(t); const x = dtplById(fd.tpl); if (x) { t.tplId = x.id; x.uses = (x.uses || 0) + 1; made = tripGenTasks(t, x.items.filter((_, i) => fd['it_' + i] != null)); } }
    else if (fd.shift != null && (old.from !== t.from || old.to !== t.to)) { const dF = dDiff(old.from, t.from), dT = dDiff(old.to, t.to);
      tripTasks(t).forEach(x => { if (x.status === 'done' || !x.due) return; const k = x.tripAt === 'ret' ? dT : dF; let nd = iso(add(P(x.due), k)); if (nd < TODAY()) nd = TODAY(); x.due = nd; if (x.start) x.start = nd; }); }
    save(); closeModal(); render();
    if (!id && fd.outlook != null) tripOutlook(t);
    toast(id ? 'Zapisano delegację' : `Utworzono delegację${made.length ? ` · ${plN(made.length, 'zadanie przygotowawcze', 'zadania przygotowawcze', 'zadań przygotowawczych')}` : ''}`);
    if (!id) tripPanel(t);
  },
  trip2tpl: (f, fd) => { const t = tripById(f.dataset.id); if (!t) return;
    const items = tripTasks(t).map(x => ({title: x.title, ctx: x.ctx, est: x.est || 0, prio: x.prio || 2, at: x.tripAt === 'ret' || x.due > t.to ? 'ret' : 'dep', d: x.tripAt === 'ret' || x.due > t.to ? Math.max(0, dDiff(t.to, x.due)) : Math.max(0, dDiff(x.due, t.from))}));
    const x = {id: uid(), name: fd.name.trim(), place: fd.keepPlace != null ? t.place : '', days: tripDays(t), ctx: t.ctx, capJob: tripCap(t, 'job'), capPriv: tripCap(t, 'priv'), dep: t.dep || '', ret: t.ret || '', note: t.note || '', items, uses: 0};
    dtplAll().push(x); save(); closeModal(); render(); toastAction(`Zapisano szablon „${x.name}” (${items.length} pozycji)`, 'Edytuj', () => dtplEditor(x), 8000); },
  dtpl: (f, fd) => { const id = f.dataset.id, x = id ? dtplById(id) : {id: uid(), uses: 0};
    const items = $$('.dt-row', f).map(r => { const v = k => r.querySelector(`[data-k=${k}]`).value; return {title: v('title').trim(), ctx: v('ctx'), at: v('at'), d: Math.max(0, +v('d') || 0), est: +v('est') || 0}; }).filter(it => it.title);
    Object.assign(x, {name: fd.name.trim(), place: fd.place.trim(), days: Math.max(1, +fd.days || 1), ctx: fd.ctx, dep: fd.dep || '', ret: fd.ret || '', note: fd.note || '', capJob: Math.max(0, Math.min(100, +fd.capJob || 0)), capPriv: Math.max(0, Math.min(100, +fd.capPriv || 0)), items});
    if (!id) dtplAll().push(x); save(); closeModal(); render(); toast('Zapisano szablon delegacji'); },
  dtplset: (f, fd) => { const st = SET(); st.tripCapJob = Math.max(0, Math.min(100, +fd.tripCapJob || 0)); st.tripCapPriv = Math.max(0, Math.min(100, +fd.tripCapPriv || 0)); st.tripLead = Math.max(0, +fd.tripLead || 0); save(); render(); toast('Zapisano ustawienia delegacji'); }
});

/* ================= UI 2026 – powłoka w stylu Spotify ================= */
const PLAY_F = '<svg class="ic f" viewBox="0 0 24 24"><path fill="currentColor" d="M8 5.6v12.8a1 1 0 0 0 1.52.85l10.2-6.4a1 1 0 0 0 0-1.7L9.52 4.75A1 1 0 0 0 8 5.6z"/></svg>';
const PAUSE_F = '<svg class="ic f" viewBox="0 0 24 24"><rect fill="currentColor" x="6" y="5" width="4" height="14" rx="1.2"/><rect fill="currentColor" x="14" y="5" width="4" height="14" rx="1.2"/></svg>';
const STOP_F = '<svg class="ic f" viewBox="0 0 24 24"><rect fill="currentColor" x="6" y="6" width="12" height="12" rx="2"/></svg>';
const EQ = '<span class="live" title="Realizacja w toku"></span>';
const PRIO_L = {1: 'Wysoki', 2: 'Średni', 3: 'Niski'};
const NAV_MAIN = ['dashboard', 'projects', 'tasks', 'meetings', 'calendar', 'people', 'avail'];
const VIEW_HUE = {dashboard: '#5c1630', projects: '#3d2a66', goals: '#3d2a66', tasks: '#1d3d78', meetings: '#0d5560', integ: '#0d5560', calendar: '#124a6b', people: '#634312', avail: '#16573f', capacity: '#16573f', plan: '#16573f', track: '#16573f', timelog: '#16573f', timereport: '#16573f', trips: '#1f4a66', gantt: '#33335e', notifs: '#5c1630', settings: '#2e2e2e', profile: '#2e2e2e'};
const navRoot = v => VIEWS[v]?.[3] || v;

/* --- artwork: deterministyczne „okładki” jak na Spotify (fale w kolorze obiektu) --- */
const hashS = s => { let h = 2166136261; for (const c of String(s)) h = Math.imul(h ^ c.charCodeAt(0), 16777619); return h >>> 0; };
const rngOf = seed => () => { seed |= 0; seed = seed + 0x6D2B79F5 | 0; let t = Math.imul(seed ^ seed >>> 15, 1 | seed); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; };
const _artC = new Map();
function waves(seed, col, n = 16) {
  const k = seed + col + n; if (_artC.has(k)) return _artC.get(k);
  const r = rngOf(hashS(seed)), y0 = 15 + r() * 40, amp = 25 + r() * 35, tilt = -30 + r() * 60, sp = 1.6 + r() * 1.6;
  let p = '';
  for (let i = 0; i < n; i++) { const y = y0 + i * sp, c1 = y - amp + r() * 8, c2 = y + amp * .6 + tilt, o = (.10 + .5 * (i / n) * r()).toFixed(2);
    p += `<path d="M-10 ${(y + tilt * .3).toFixed(1)} C 30 ${c1.toFixed(1)}, 60 ${c2.toFixed(1)}, 110 ${(y - tilt * .4 + i).toFixed(1)}" stroke="${i % 5 === 0 ? col : '#fff'}" stroke-opacity="${o}" stroke-width="${i % 5 === 0 ? .9 : .45}" fill="none"/>`; }
  const s = `<svg class="waves" viewBox="0 0 100 100" preserveAspectRatio="xMidYMid slice" aria-hidden="true">${p}</svg>`;
  _artC.set(k, s); return s;
}

/* --- klienci (firmy z osób i projektów) --- */
function companiesAll() {
  const m = new Map(), key = s => norm(String(s || '').trim()), add = (name, kind, x) => { const k = key(name); if (!k) return; const c = m.get(k) || {id: k, name: String(name).trim(), people: [], projects: [], ctx: new Set()}; c[kind].push(x); c.ctx.add(x.ctx); m.set(k, c); };
  peopleAll().forEach(p => add(p.company, 'people', p)); (S.data.projects || []).forEach(p => add(p.client, 'projects', p));
  return [...m.values()].filter(c => S.ctx === 'all' || c.ctx.has(S.ctx)).sort((a, b) => (b.projects.length + b.people.length) - (a.projects.length + a.people.length) || a.name.localeCompare(b.name, 'pl'));
}
const coInitials = n => n.split(/[\s.\-–]+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase();

/* --- historia nawigacji (wstecz / dalej) --- */
const NAVH = {back: [], fwd: [], cur: null};
const navKey = () => S.view + '|' + (S.view === 'project' ? S.proj.open || '' : '');
function navTrack() { const k = navKey(); if (NAVH.cur && NAVH.cur !== k) { NAVH.back.push(NAVH.cur); if (NAVH.back.length > 60) NAVH.back.shift(); NAVH.fwd = []; } NAVH.cur = k; }
function navGo(dir) {
  const from = dir < 0 ? NAVH.back : NAVH.fwd, to = dir < 0 ? NAVH.fwd : NAVH.back; if (!from.length) return;
  to.push(NAVH.cur); const k = from.pop(), [v, p] = k.split('|'); NAVH.cur = k;
  if (!VIEWS[v]) return; S.view = v; if (v === 'project') S.proj.open = p || null; history.replaceState(null, '', '#' + v); closeModal(); render(); $('main').scrollTop = 0;
}

/* --- pasek boczny --- */
/* --- pasek „odtwarzacza” = stoper realizacji (RCP) --- */
/* --- render --- */
/* --- menu „Utwórz” --- */
function popMenu(el, items) {
  const r = el.getBoundingClientRect(), root = $('#popRoot');
  root.innerHTML = H(`<div class="pop" style="left:${Math.min(r.left, innerWidth - 270)}px;top:${r.bottom + 6}px">${items.map(([a, i, l, s]) => `<button data-action="${a}" class="pop-i"><span class="ico">${i}</span><span><b>${l}</b>${s ? `<small>${s}</small>` : ''}</span></button>`).join('')}</div>`);
}
const closePop = () => { const r = $('#popRoot'); if (r) r.innerHTML = ''; };
document.addEventListener('mousedown', e => { if (!e.target.closest('.pop,[data-action=create-menu]')) closePop(); if (!e.target.closest('.search')) searchClose(); });

/* --- wyszukiwarka globalna --- */
const SR = {q: '', i: 0, list: []};
function searchAll(q) {
  const n = norm(q).trim(); if (!n) return [];
  const sc = s => { const x = norm(s || ''); return !x ? 0 : x.startsWith(n) ? 3 : x.split(/[\s\-–/·,.()]+/).some(w => w.startsWith(n)) ? 2 : x.includes(n) ? 1 : 0; };
  const R = [], push = (o, s) => { if (s > 0) R.push({...o, s}); };
  Object.entries(VIEWS).filter(([k, v]) => !v[5] || ['settings', 'notifs'].includes(k)).forEach(([k, v]) => push({k: 'nav', id: k, t: v[1], sub: v[3] ? `Moduł · ${VIEWS[v[3]][1]}` : 'Moduł', ic: v[0]}, sc(v[1]) * 1.1));
  (S.data.projects || []).forEach(p => push({k: 'project', id: p.id, t: p.title, sub: `Projekt · ${p.client || (p.ctx === 'priv' ? 'PRIV' : 'JOB')}`, art: projArt(p, 'sq')}, Math.max(sc(p.title) + .5, sc(p.client), sc(p.desc) * .5)));
  S.data.goals.forEach(g => push({k: 'goal', id: g.id, t: g.title, sub: `Podprojekt · ${projById(g.projectId)?.title || 'bez projektu'}`, art: art(g.id, projColor(projById(g.projectId)) || ctxColor(g.ctx), '⟨i:layers⟩', 'sq')}, Math.max(sc(g.title) + .3, sc(g.desc) * .5)));
  S.data.tasks.forEach(t => push({k: 'task', id: t.id, t: t.title, sub: `Zadanie · ${projById(projOf(t))?.title || (t.ctx === 'priv' ? 'PRIV' : 'JOB')}${t.due ? ' · ' + relDay(t.due) : ''}${t.status === 'done' ? ' · zrobione' : ''}`, ic: t.status === 'done' ? '⟨i:check⟩' : '⟨i:check-square⟩', done: t.status === 'done'}, Math.max(sc(t.title), sc(t.ic) * 1.2, sc(t.notes) * .4) - (t.status === 'done' ? .6 : 0)));
  S.data.meetings.forEach(m => push({k: 'mtg', id: m.id, t: m.title, sub: `Spotkanie · ${fmt(m.start.slice(0, 10))} ${hm(m.start)}${m.location ? ' · ' + m.location : ''}`, ic: '⟨i:video⟩'}, Math.max(sc(m.title), sc(m.location) * .6, sc((m.attendees || []).join(' ')) * .5) - (m.end.slice(0, 10) < TODAY() ? .4 : 0)));
  peopleAll().forEach(p => push({k: 'person', id: p.id, t: pName(p), sub: `Osoba · ${[p.role, p.company].filter(Boolean).join(' · ') || p.email || ''}`, round: pInitials(p), hue: hashS(p.id) % 360}, Math.max(sc(pName(p)) + .2, sc(p.email), sc(p.company) * .7)));
  companiesAll().forEach(c => push({k: 'client', id: c.name, t: c.name, sub: `Klient · ${plN(c.projects.length, 'projekt', 'projekty', 'projektów')}`, round: coInitials(c.name), hue: hashS(c.id) % 360}, sc(c.name) + .4));
  tripsAll().forEach(t => push({k: 'trip', id: t.id, t: t.place, sub: `Delegacja · ${tripRange(t)}`, ic: '⟨i:plane⟩'}, Math.max(sc(t.place), sc(t.note) * .5)));
  notesAll().forEach(x => push({k: 'note', id: x.id, t: x.title, sub: 'Notatka', ic: '⟨i:note⟩'}, Math.max(sc(x.title), sc(x.text || x.body || '') * .4)));
  return R.sort((a, b) => b.s - a.s);
}
const SR_G = [['project', 'Projekty'], ['goal', 'Podprojekty'], ['task', 'Zadania'], ['mtg', 'Spotkania'], ['person', 'Osoby'], ['client', 'Klienci'], ['trip', 'Delegacje'], ['note', 'Notatki'], ['nav', 'Moduły']];
function searchOpen(r) {
  if (!r) return; searchClose(true);
  const rec = (LS.get('tm:recent', []) || []).filter(x => !(x.k === r.k && x.id === r.id)); rec.unshift({k: r.k, id: r.id, t: r.t, sub: r.sub, art: r.art, ic: r.ic, round: r.round, hue: r.hue}); LS.set('tm:recent', rec.slice(0, 8));
  const go = (v, f) => { S.view = v; f && f(); history.replaceState(null, '', '#' + v); render(); $('main').scrollTop = 0; };
  switch (r.k) {
    case 'cmd': return A[r.id]?.({dataset: {}}, {stopPropagation() { }});
    case 'nav': return go(r.id);
    case 'project': if (!projById(r.id)) return toast('Projekt nie istnieje'); return go('project', () => S.proj.open = r.id);
    case 'goal': { const g = goalById(r.id); if (!g) return; if (g.projectId && projById(g.projectId)) go('project', () => S.proj.open = g.projectId); return editGoal(g); }
    case 'task': { const t = S.data.tasks.find(x => x.id === r.id); return t ? editTask(t) : toast('Zadanie zostało usunięte'); }
    case 'mtg': { const m = S.data.meetings.find(x => x.id === r.id); return m ? editMeeting(m) : toast('Spotkanie nie istnieje'); }
    case 'person': { const p = personById(r.id); return p ? personPanel(p) : null; }
    case 'client': return go('people', () => { (S.pf ||= {q: '', kind: '', sort: 'name'}).q = r.id; });
    case 'trip': return tripPanel(tripById(r.id));
    case 'note': { const n = notesAll().find(x => x.id === r.id); return n ? noteEditor(n) : null; }
  }
}
function searchClose(clear) { const b = $('#searchDrop'); if (b) b.hidden = true; if (clear) { const i = $('#gsearch'); if (i) { i.value = ''; i.blur(); } } $('.search')?.classList.remove('open'); }
function searchInit() {
  const i = $('#gsearch'); if (!i || i.dataset.ok) return; i.dataset.ok = 1;
  i.addEventListener('focus', () => { $('.search').classList.add('open'); SR.i = 0; searchRender(); });
  i.addEventListener('input', () => { SR.i = 0; searchRender(); });
  i.addEventListener('keydown', e => {
    if (e.key === 'Escape') { e.stopPropagation(); return searchClose(true); }
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); const n = SR.list.length; if (!n) return; SR.i = (SR.i + (e.key === 'ArrowDown' ? 1 : -1) + n) % n; $$('#searchDrop .sr-i').forEach(x => x.classList.toggle('on', +x.dataset.n === SR.i)); $(`#searchDrop .sr-i[data-n="${SR.i}"]`)?.scrollIntoView({block: 'nearest'}); }
    if (e.key === 'Enter') { e.preventDefault(); searchOpen(SR.list[SR.i]); }
    if (e.key === 'Tab') { e.preventDefault(); }
  });
}
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k' && S.user) { e.preventDefault(); closeModal(); $('#gsearch')?.focus(); $('#gsearch')?.select(); }
  else if (e.key === '/' && S.user && !/INPUT|TEXTAREA|SELECT/.test(document.activeElement?.tagName) && !$('#modalRoot').innerHTML) { e.preventDefault(); $('#gsearch')?.focus(); }
  else if (e.altKey && (e.key === 'ArrowLeft' || e.key === 'ArrowRight') && S.user) { e.preventDefault(); navGo(e.key === 'ArrowLeft' ? -1 : 1); }
});

/* --- lista zadań w stylu tracklisty --- */
function clientCard(c) {
  return `<div class="pfcard" data-action="lib-client" data-id="${esc(c.name)}"><span class="pf-tile" style="--a:hsl(${hashS(c.id) % 360} 52% 42%)">${esc(coInitials(c.name))}</span>
    <div class="pf-1">${S.privacy ? pv(esc(c.name)) : esc(c.name)}</div><div class="pf-2">${plN(c.projects.length, 'projekt', 'projekty', 'projektów')} · ${plN(c.people.length, 'osoba', 'osoby', 'osób')}</div></div>`;
}
function contCard(t) {
  const p = projById(projOf(t)), sp = taskSpent(t.id), est = (t.est || 0) * 60, pct = est ? Math.min(100, Math.round(100 * sp / est)) : 0, on = timerOn()?.taskId === t.id, P_ = S.privacy ? pv : y => y;
  return `<div class="ncard cont" data-action="edit-task" data-id="${t.id}"><div class="nc-thumb">${art(p ? p.id : t.id, p ? projColor(p) : ctxColor(t.ctx), '', 'wide')}
      ${on ? '<span class="nc-flag onair">W toku</span>' : isLate(t) ? '<span class="nc-flag bad">Po terminie</span>' : ''}<div class="nc-title sm">${P_(esc(t.title))}</div><span class="nc-prog"><i style="width:${pct}%"></i></span></div>
    <div class="nc-info"><div class="nc-meta"><span>${fmtH(sp) === '<1m' && !sp ? '0m' : fmtH(sp)}${est ? ' z ' + fmtH(est) : ''}</span>${t.due ? `<span class="${isLate(t) ? 'late' : ''}">${relDay(t.due)}</span>` : ''}</div>
      <div class="nc-sub">${p ? P_(esc(p.title)) : t.ctx === 'priv' ? 'PRIV' : 'JOB'}</div>
      <div class="nc-acts">${on ? `<button class="rbtn sm" data-action="timer-pause">${PAUSE_F} Pauza</button>` : `<button class="rbtn light sm" data-action="timer-start-task" data-id="${t.id}">${PLAY_F} Wznów</button>`}</div></div></div>`;
}
function contTasks() {
  const run = timerOn()?.taskId, last = new Map(); entriesAll().forEach(e => { if (e.taskId && (e.endMs || 0) > (last.get(e.taskId) || 0)) last.set(e.taskId, e.endMs || 0); });
  return tasksC().filter(t => t.status !== 'done' && (t.id === run || t.status === 'doing' || last.has(t.id))).sort((a, b) => (b.id === run) - (a.id === run) || (last.get(b.id) || 0) - (last.get(a.id) || 0)).slice(0, 12);
}
const rowNav = id => `<span class="row-nav"><button class="rn" data-action="row-scroll" data-row="${id}" data-d="-1" title="Przewiń w lewo">⟨i:chev-left⟩</button><button class="rn" data-action="row-scroll" data-row="${id}" data-d="1" title="Przewiń w prawo">⟨i:chev-right⟩</button></span>`;
const shelf = (title, body, more) => `<section class="shelf"><div class="sh-h"><h3>${title}</h3>${more ? `<a href="#" ${more}>Pokaż wszystko</a>` : ''}</div>${body}</section>`;

/* --- Start --- */
const DASH_MODES = [['over', 'Przegląd'], ['projects', 'Projekty'], ['tasks', 'Zadania'], ['meetings', 'Spotkania'], ['clients', 'Osoby']];
const tripOf = d => { const s = typeof d === 'string' ? d : iso(d); return tripsAll().find(t => t.from <= s && t.to >= s && inCtx(t)); };
const firstName = () => String(S.user.name || '').trim().split(/\s+/)[0] || '';
function dashHero(mode) {
  const td = TODAY(), d = new Date(), wd = d.toLocaleDateString('pl-PL', {weekday: 'long', day: 'numeric', month: 'long'}), P_ = S.privacy ? pv : y => y;
  const act = projectsC().filter(p => p.status === 'active').sort((a, b) => (a.due || '9').localeCompare(b.due || '9')), tToday = tasksC().filter(t => t.status !== 'done' && t.due && t.due <= td).length;
  const mToday = meetingsC().filter(m => m.start.slice(0, 10) === td).length, h = d.getHours(), tr = tripOf(td);
  const hi = h < 5 ? 'Dobranoc' : h < 18 ? 'Dzień dobry' : 'Dobry wieczór';
  const i = act.length ? ((S.heroI || 0) % act.length + act.length) % act.length : 0, p = act[i], x = p && projStats(p);
  const nx = p && S.data.tasks.filter(t => projOf(t) === p.id && t.status !== 'done').sort((a, b) => a.prio - b.prio || prioSort(a, b))[0];
  const tabs = `<nav class="hero-tabs">${DASH_MODES.map(([k, l]) => `<button class="htab ${mode === k ? 'on' : ''}" data-action="dash-mode" data-m="${k}">${l}</button>`).join('')}</nav>`;
  const greet = `<div class="hero-greet"><b>${hi}${firstName() ? ', ' + esc(firstName()) : ''}.</b> <span>${wd} · ${plN(act.length, 'aktywny projekt', 'aktywne projekty', 'aktywnych projektów')} · ${plN(tToday, 'zadanie', 'zadania', 'zadań')} na dziś · ${plN(mToday, 'spotkanie', 'spotkania', 'spotkań')}</span>${tr ? ` · <a href="#" class="hero-trip" data-action="trip-open" data-id="${tr.id}">⟨i:plane⟩ delegacja: ${P_(esc(tr.place))}</a>` : ''}</div>`;
  if (mode !== 'over' || !p) return `<div class="billboard slim">${greet}<h1 class="bb-h">${{over: 'Przegląd', projects: 'Projekty', tasks: 'Zadania', meetings: 'Spotkania', clients: 'Osoby'}[mode]}</h1>${tabs}</div>`;
  return `<div class="billboard" style="--a:${projColor(p) || ctxColor(p.ctx)}"><div class="bb-art">${waves('bb-' + p.id, projColor(p) || '#C8102E', 30)}</div>
    ${greet}<div class="bb-k"><span class="bb-mark"></span>${p.ctx === 'priv' ? 'Projekt prywatny' : 'Projekt'} · w realizacji${act.length > 1 ? ` · ${i + 1} z ${act.length}` : ''}</div>
    <h1 class="bb-h">${P_(esc(p.title))}</h1>
    <div class="bb-meta"><b class="ok">${x.pr}% realizacji</b><span>plan ${x.ex}%</span><span>${x.ts.length - x.open.length}/${x.ts.length} zadań</span>${p.due ? `<span>koniec ${fmt(p.due)}</span>` : ''}${p.client ? `<span class="bb-tag">${P_(esc(p.client))}</span>` : ''}${x.late.length ? `<span class="late">${x.late.length} po terminie</span>` : ''}</div>
    ${nx ? `<p class="bb-next">Następne zadanie: <b>${P_(esc(nx.title))}</b>${nx.due ? ` · ${relDay(nx.due)}` : ''}</p>` : '<p class="bb-next">Brak otwartych zadań w tym projekcie.</p>'}
    <div class="bb-acts"><button class="rbtn light lg" data-action="proj-play" data-id="${p.id}">${PLAY_F} ${timerOn() && projById(timerOn().projectId) === p ? 'W realizacji' : 'Kontynuuj realizację'}</button><button class="rbtn ghost lg" data-action="open-project" data-id="${p.id}">⟨i:folder⟩ Otwórz projekt</button>
      ${act.length > 1 ? `<span class="bb-nav"><button class="rn" data-action="hero-i" data-d="-1" title="Poprzedni projekt">⟨i:chev-left⟩</button><button class="rn" data-action="hero-i" data-d="1" title="Następny projekt">⟨i:chev-right⟩</button></span>` : ''}</div>
    ${tabs}</div>`;
}
function todayTasks() { const td = TODAY(), open = tasksC().filter(t => t.status !== 'done');
  return [...open.filter(isLate), ...open.filter(t => t.due === td), ...open.filter(t => t.status === 'doing' && t.due !== td && !isLate(t))].filter((t, i, a) => a.indexOf(t) === i).sort((a, b) => a.prio - b.prio || prioSort(a, b)); }
function dayAgenda() {
  const off = S.dashDay || 0, ds = iso(add(today(), off)), now = new Date().toTimeString().slice(0, 5), P_ = S.privacy ? pv : x => x;
  const ms = meetingsC().filter(m => m.start.slice(0, 10) === ds).sort((a, b) => a.start.localeCompare(b.start)), tr = tripOf(ds);
  const lbl = off === 0 ? 'Dzisiaj' : off === 1 ? 'Jutro' : off === -1 ? 'Wczoraj' : DOWL[P(ds).getDay()] + ' ' + fmt(ds);
  return `<section class="pane"><div class="pane-h"><h3>⟨i:calendar⟩ ${lbl}</h3><span class="spacer"></span>${eyeBtn('sm')}<button class="circ sm" data-action="dash-day" data-d="-1" title="Poprzedni dzień">⟨i:chev-left⟩</button><button class="circ sm" data-action="dash-day" data-d="1" title="Następny dzień">⟨i:chev-right⟩</button></div>
    ${tr ? `<div class="ag-trip" data-action="trip-open" data-id="${tr.id}">⟨i:plane⟩ Delegacja: ${P_(esc(tr.place))}</div>` : ''}
    <div class="agenda">${ms.map(m => { const past = off < 0 || (off === 0 && m.end.slice(11) < now), live = off === 0 && m.start.slice(11) <= now && m.end.slice(11) >= now;
      return `<div class="ag ${past ? 'past' : ''} ${live ? 'live' : ''}" data-action="edit-mtg" data-id="${m.id}"><span class="ag-dot"></span><div class="ag-time"><b>${hm(m.start)}</b><small>– ${hm(m.end)}</small></div><div class="ag-t"><div class="ag-1">${P_(esc(m.title))}</div><div class="ag-2">${m.location ? P_(esc(m.location)) + ' · ' : ''}${m.link ? 'Online ⟨i:video⟩' : m.source === 'outlook' ? 'Outlook' : 'TM'}</div></div>${m.link && !past ? joinLink(m, 'Dołącz', 'btn sm join') : ''}</div>`; }).join('') || `<div class="ag-empty">Brak spotkań${off === 0 ? ' dziś' : ''}.</div>`}</div>
    <button class="linkbtn" data-action="nav" data-view="calendar">Otwórz kalendarz ⟨i:arrow-right⟩</button></section>`;
}
function quickTiles() {
  const qa = [['quick', '⟨i:plus⟩', 'Nowe zadanie'], ['dq-project', '⟨i:folder⟩', 'Nowy projekt'], ['dq-mtg', '⟨i:video⟩', 'Spotkanie'], ['dq-person', '⟨i:users⟩', 'Kontakt'], ['dq-start', '⟨i:play⟩', 'Start realizacji'], ['dq-trip', '⟨i:plane⟩', 'Delegacja']];
  return `<section class="pane"><div class="pane-h"><h3>⟨i:bolt⟩ Szybkie akcje</h3></div><div class="qtiles">${qa.map(([a, i, l]) => `<button class="qtile" data-action="${a}">${i}<span>${l}</span></button>`).join('')}</div></section>`;
}
function prioPane() {
  const l = tasksC().filter(t => t.prio === 1 && (t.status !== 'done' || t.doneAt === TODAY())).sort((a, b) => (a.status === 'done') - (b.status === 'done') || prioSort(a, b)).slice(0, 5), P_ = S.privacy ? pv : y => y;
  return `<section class="pane"><div class="pane-h"><h3>Top 5 priorytetów</h3><span class="spacer"></span><span class="muted sm">P1</span></div>
    ${l.map((t, i) => `<div class="rank ${t.status === 'done' ? 'done' : ''}"><span class="rk-n">${i + 1}</span><div class="rk-t" data-action="edit-task" data-id="${t.id}"><div class="rk-1">${P_(esc(t.title))}</div><div class="rk-2">${projById(projOf(t)) ? P_(esc(projById(projOf(t)).title)) + ' · ' : ''}${isLate(t) ? '<span class="late">po terminie</span>' : t.due ? relDay(t.due) : 'bez terminu'}</div></div><span class="tchk ${t.status === 'done' ? 'on' : ''}" data-action="toggle" data-id="${t.id}" title="Zrobione">${t.status === 'done' ? '⟨i:check⟩' : ''}</span></div>`).join('') || '<div class="ag-empty">Brak zadań z priorytetem P1.</div>'}</section>`;
}
function rcpToday() {
  const td = TODAY(), es = entriesWithLive().filter(e => e.start.slice(0, 10) === td), m = new Map();
  es.forEach(e => { const k = (e.projectId || '-') + '|' + (e.taskId || 'free:' + (e.title || '')); const t = e.taskId && S.data.tasks.find(x => x.id === e.taskId);
    const r = m.get(k) || {pid: e.projectId, title: t ? t.title : e.title || 'Realizacja bez zadania', free: !e.taskId, min: 0, live: false, n: 0}; r.min += entMin(e); r.n++; if (e.live) r.live = true; m.set(k, r); });
  const rows = [...m.values()].sort((a, b) => b.min - a.min), tot = rows.reduce((a, r) => a + r.min, 0), tgt = capacityOfDay(today(), 'job', true) * 60;
  return {rows, tot, tgt, n: es.length, projects: new Set(rows.map(r => r.pid || '-')).size};
}
const hDec = m => (Math.round(m / 6) / 10).toFixed(1).replace('.', ',');
function rcpText() {
  const r = rcpToday(), pn = id => projById(id)?.title || 'Bez projektu';
  return [`RCP ${fmt(TODAY())}`, ...r.rows.map(x => `${pn(x.pid)} – ${x.title}: ${fmtH(x.min)} (${hDec(x.min)} h)`), `Razem: ${fmtH(r.tot)} (${hDec(r.tot)} h)`].join('\n');
}
function rcpPane() {
  const r = rcpToday(), P_ = S.privacy ? pv : x => x, pct = r.tgt ? Math.min(100, Math.round(100 * r.tot / r.tgt)) : 0;
  return `<section class="pane rcp-pane"><div class="pane-h"><h3>Do RCP – dziś</h3><span class="spacer"></span><b class="rcp-sum">${hDec(r.tot)} h</b></div>
    <span class="pbar"><i style="width:${pct}%"></i></span><div class="rcp-sub">${fmtH(r.tot) === '<1m' && !r.tot ? '0m' : fmtH(r.tot)} z ${r1(r.tgt / 60)}h · ${plN(r.n, 'wpis', 'wpisy', 'wpisów')}</div>
    ${r.rows.map(x => `<div class="rcp-row"><div class="rr-t"><div class="rr-1">${x.live ? EQ : ''}${P_(esc(x.title))}${x.free ? ' <span class="tag src">bez zadania</span>' : ''}</div><div class="rr-2">${P_(esc(projById(x.pid)?.title || 'Bez projektu'))}</div></div><b class="rr-h">${fmtH(x.min)}</b></div>`).join('') || '<div class="ag-empty">Dziś nic jeszcze nie zaraportowano.</div>'}
    <div class="rcp-acts"><button class="rbtn sm light" data-action="rcp-copy" ${r.rows.length ? '' : 'disabled'} title="Skopiuj zestawienie do schowka – wklej do RCP">⟨i:check-square⟩ Kopiuj do RCP</button><button class="rbtn sm" data-action="dq-entry">⟨i:plus⟩ Wpis</button><button class="rbtn sm ghost" data-action="rcp-log">Ewidencja</button></div></section>`;
}
function homeStats() {
  const r = rcpToday(), wk = weekChart(iso(sow(today()))), miss = r.tgt - r.tot, tf = m => fmtH(m) === '<1m' && !m ? '0h' : fmtH(m);
  const st = (v, l, s, act, cls = '') => `<div class="stat ${cls}" data-action="${act}"><div class="st-v">${v}</div><div class="st-l">${l}</div><div class="st-s">${s}</div></div>`;
  return `<div class="stats home-stats">${st(tf(r.tot), 'Zaraportowane dziś', `${r.tgt ? Math.round(100 * r.tot / r.tgt) : 0}% z ${r1(r.tgt / 60)}h${timerOn() ? ' · stoper w toku' : ''}`, 'rcp-log', r.tot >= r.tgt && r.tgt ? 'ok' : '')}
    ${st(hDec(r.tot) + ' h', 'Do przekazania do RCP', `${plN(r.n, 'wpis', 'wpisy', 'wpisów')} · ${plN(r.projects, 'projekt', 'projekty', 'projektów')}`, 'rcp-copy')}
    ${st(miss > 0 ? tf(miss) : '+' + tf(-miss), miss > 0 ? 'Brakuje do celu dnia' : 'Cel dnia osiągnięty', miss > 0 ? 'uzupełnij stoperem lub wpisem' : 'nadwyżka ponad plan', miss > 0 ? 'dq-entry' : 'rcp-log', miss > 0 ? 'warn' : 'ok')}
    ${st(tf(wk.total), 'Ten tydzień', `z ${wk.target / 60}h · ${wk.target ? Math.round(100 * wk.total / wk.target) : 0}%`, 'rcp-week')}</div>`;
}
function freeRow(p) {
  return `<div class="ep free"><span class="ep-n">⟨i:timer⟩</span><div class="ep-t"><div class="ep-1">Realizacja bez zadania</div><div class="ep-2">${p ? `Czas trafi do RCP na projekt „${S.privacy ? pv(esc(p.title)) : esc(p.title)}” – bez tworzenia zadania.` : 'Czas trafi do RCP jako ogólna praca – bez tworzenia zadania.'}</div></div>
    <button class="rbtn light sm" data-action="free-start" ${p ? `data-project="${p.id}"` : ''}>${PLAY_F} Start</button></div>`;
}
function episodes(p) {
  const G = S.homeGoal || '', gs = projGoals(p).filter(g => goalProgress(g) < 100 || G === g.id), run = timerOn()?.taskId, P_ = S.privacy ? pv : x => x;
  const ts = S.data.tasks.filter(t => projOf(t) === p.id && t.status !== 'done' && (!G || t.goalId === G)).sort((a, b) => (a.status === 'doing' ? -1 : 0) - (b.status === 'doing' ? -1 : 0) || prioSort(a, b));
  const ep = (t, i) => { const g = t.goalId && goalById(t.goalId), on = run === t.id, sp = taskSpent(t.id), est = (t.est || 0) * 60;
    return `<div class="ep pr${t.prio} ${on ? 'playing' : ''}"><span class="ep-n">${i + 1}</span><span class="ep-th">${art(t.id, projColor(p) || ctxColor(p.ctx), '', 'wide')}${est ? `<span class="nc-prog"><i style="width:${Math.min(100, 100 * sp / est)}%"></i></span>` : ''}</span>
      <div class="ep-t" data-action="edit-task" data-id="${t.id}"><div class="ep-1">${on ? EQ : ''}${P_(esc(t.title))}<span class="ep-d">${t.est ? r1(t.est) + 'h' : ''}</span></div>
        <div class="ep-2">${g ? P_(esc(g.title)) + ' · ' : ''}<span class="${isLate(t) ? 'late' : ''}">${t.due ? relDay(t.due) : 'bez terminu'}</span> · ${PRIO_L[t.prio]}${sp ? ' · zaraportowano ' + fmtH(sp) : ''}</div>
        ${t.notes ? `<div class="ep-3">${P_(esc(t.notes.slice(0, 160)))}${t.notes.length > 160 ? '…' : ''}</div>` : ''}</div>
      <div class="ep-a">${on ? `<button class="rbtn sm" data-action="timer-pause">${PAUSE_F} Pauza</button>` : `<button class="rbtn sm" data-action="timer-start-task" data-id="${t.id}">${PLAY_F} Start</button>`}<span class="tchk" data-action="toggle" data-id="${t.id}" title="Oznacz jako zrobione"></span></div></div>`; };
  return `<section class="shelf"><div class="sh-h"><h3>${P_(esc(p.title))} <span class="cnt">${ts.length}</span></h3><span class="spacer"></span>
      ${gs.length ? `<select class="input season" data-bind="homeGoal"><option value="">Wszystkie podprojekty</option>${gs.map(g => `<option value="${g.id}" ${G === g.id ? 'selected' : ''}>${esc(g.title)}</option>`).join('')}</select>` : ''}
      <button class="rbtn sm ghost" data-action="home-proj" data-id="">⟨i:x⟩ Zadania na dziś</button></div>
    <div class="eps">${freeRow(p)}${ts.map(ep).join('') || '<div class="tl-empty">Brak otwartych zadań w tym projekcie.</div>'}</div></section>`;
}
function dashOver() {
  const ps = projectsC().filter(p => p.status === 'active').sort((a, b) => (a.due || '9').localeCompare(b.due || '9')), tt = todayTasks(), F = S.dtf || 'today', ct = contTasks();
  const td = TODAY(), we = iso(add(sow(today()), 6)), open = tasksC().filter(t => t.status !== 'done'), sel = projById(S.homeProj) && ps.some(p => p.id === S.homeProj) ? S.homeProj : '';
  const lists = {today: tt, late: open.filter(isLate).sort(prioSort), week: open.filter(t => t.due && t.due >= td && t.due <= we).sort(prioSort)};
  const qa = [['quick', '⟨i:plus⟩', 'Nowe zadanie', 'szybkie dodanie (N)'], ['dq-start', '⟨i:play⟩', 'Start realizacji', 'stoper na zadaniu'], ['free-start', '⟨i:timer⟩', 'Bez zadania', 'stoper do RCP'], ['dq-entry', '⟨i:clock⟩', 'Wpis czasu', 'post factum'], ['dq-mtg', '⟨i:video⟩', 'Nowe spotkanie', 'Outlook / Teams'], ['dq-person', '⟨i:users⟩', 'Nowa osoba', 'kontakt'], ['dq-project', '⟨i:folder⟩', 'Nowy projekt', 'pusty lub z szablonu'], ['dq-vac', '⟨i:sun⟩', 'Urlop', 'nieobecność'], ['dq-trip', '⟨i:plane⟩', 'Delegacja', 'z szablonu']];
  const main = sel ? episodes(projById(sel)) : `<section class="shelf"><div class="sh-h"><h3>Zadania na dziś <span class="cnt">${tt.length}</span></h3><span class="spacer"></span>
        <div class="chips">${[['today', 'Dziś'], ['late', 'Zaległe'], ['week', 'Ten tydzień']].map(([k, l]) => `<button class="fchip ${F === k ? 'on' : ''}" data-action="set" data-k="dtf" data-v="${k}">${l}${k === 'late' && lists.late.length ? ` (${lists.late.length})` : ''}</button>`).join('')}</div></div>
        <div class="eps">${freeRow(null)}</div>${trackList(lists[F].slice(0, 12), {empty: F === 'today' ? 'Na dziś nic nie zaplanowano.' : 'Pusto.'})}${lists[F].length > 12 ? `<a href="#" class="linkbtn" data-action="nav" data-view="tasks">Pokaż wszystkie (${lists[F].length}) ⟨i:arrow-right⟩</a>` : ''}</section>`;
  return `<section class="qa-sec"><div class="sh-h"><h3>Szybkie akcje</h3></div><div class="qrow">${qa.map(([a, i, l, d]) => `<button class="qbtn" data-action="${a}" title="${d}">${i}<span><b>${l}</b><small>${d}</small></span></button>`).join('')}</div></section>${homeStats()}
    <section class="row-sec"><div class="sh-h"><h3>Aktywne projekty</h3><span class="muted sm">kliknij projekt, aby zobaczyć jego zadania poniżej</span><a href="#" data-action="nav" data-view="projects">Wszystkie ⟨i:chev-right⟩</a>${rowNav('rowProj')}</div>${ps.length ? `<div class="nrow" id="rowProj">${ps.map(p => projCard(p, sel)).join('')}</div>` : '<div class="tl-empty">Brak aktywnych projektów. <a href="#" data-action="dq-project">Utwórz projekt</a></div>'}</section>
    <div class="home-grid" id="homeMain"><div class="home-main">${main}</div><aside class="home-side">${rcpPane()}${dayAgenda()}${prioPane()}</aside></div>
    ${ct.length ? `<section class="row-sec" style="margin-top:10px"><div class="sh-h"><h3>Kontynuuj realizację</h3>${rowNav('rowCont')}</div><div class="nrow" id="rowCont">${ct.map(contCard).join('')}</div></section>` : ''}`;
}
function dashProjects() {
  const ps = projectsC().filter(p => p.status !== 'done').sort((a, b) => (a.status === 'active' ? 0 : 1) - (b.status === 'active' ? 0 : 1) || (a.due || '9').localeCompare(b.due || '9'));
  const jobs = ps.filter(p => p.ctx === 'job'), privs = ps.filter(p => p.ctx === 'priv');
  return `<div class="row" style="margin-bottom:8px;gap:8px">${eyeBtn('sm')}<span class="spacer"></span><button class="btn" data-action="ptpl-use">⟨i:template⟩ Z szablonu</button><button class="btn primary" data-action="dq-project">⟨i:plus⟩ Nowy projekt</button></div>
    ${jobs.length ? shelf('Firmowe', `<div class="ngrid">${jobs.map(projCard).join('')}</div>`, 'data-action="nav" data-view="projects"') : ''}${privs.length ? shelf('Prywatne', `<div class="ngrid">${privs.map(projCard).join('')}</div>`) : ''}${ps.length ? '' : '<div class="tl-empty">Brak projektów.</div>'}`;
}
function dashTasks() {
  const td = TODAY(), we = iso(add(sow(today()), 6)), open = tasksC().filter(t => t.status !== 'done');
  const late = open.filter(isLate).sort(prioSort), tdy = open.filter(t => t.due === td).sort((a, b) => a.prio - b.prio), wk = open.filter(t => t.due > td && t.due <= we).sort(prioSort);
  return `<div class="row" style="margin-bottom:8px;gap:8px"><span class="spacer"></span><button class="btn" data-action="nav" data-view="tasks">⟨i:list⟩ Wszystkie zadania</button><button class="btn primary" data-action="quick">⟨i:plus⟩ Zadanie</button></div>
    ${late.length ? shelf(`<span class="late">Zaległe</span> <span class="cnt">${late.length}</span>`, trackList(late)) : ''}${shelf(`Dziś <span class="cnt">${tdy.length}</span>`, trackList(tdy, {empty: 'Nic na dziś.'}))}${shelf(`Do końca tygodnia <span class="cnt">${wk.length}</span>`, trackList(wk, {empty: 'Nic więcej w tym tygodniu.'}))}`;
}
function dashMeetings() {
  const td = TODAY(), days = [...Array(7)].map((_, i) => iso(add(today(), i))), ms = meetingsC().filter(m => m.start.slice(0, 10) >= td && m.start.slice(0, 10) <= days[6]).sort((a, b) => a.start.localeCompare(b.start));
  const ups = tripsAll().filter(t => t.to >= td && inCtx(t)).sort((a, b) => a.from.localeCompare(b.from)).slice(0, 4);
  return `<div class="row" style="margin-bottom:8px;gap:8px">${eyeBtn('sm')}${importBtn('sm')}<span class="spacer"></span><button class="btn" data-action="nav" data-view="calendar">⟨i:calendar⟩ Kalendarz</button><button class="btn primary" data-action="dq-mtg">⟨i:plus⟩ Spotkanie</button></div>
    <div class="home-grid"><div class="home-main">${days.map(d => { const l = ms.filter(m => m.start.slice(0, 10) === d); if (!l.length) return '';
      return shelf(`${d === td ? 'Dziś' : relDay(d)[0].toUpperCase() + relDay(d).slice(1)} <span class="muted sm">${fmt(d)}</span>`, `<div class="pane flat">${l.map(m => meetingRow(m)).join('')}</div>`); }).join('') || '<div class="tl-empty">Brak spotkań w najbliższych 7 dniach.</div>'}</div>
    <aside class="home-side">${ups.length ? `<section class="pane"><div class="pane-h"><h3>⟨i:plane⟩ Delegacje</h3></div>${ups.map(t => `<div class="ag" data-action="trip-open" data-id="${t.id}"><span class="ag-dot"></span><div class="ag-time"><b>${fmt(t.from)}</b><small>${plN(tripDays(t), 'dzień', 'dni', 'dni')}</small></div><div class="ag-t"><div class="ag-1">${S.privacy ? pv(esc(t.place)) : esc(t.place)}</div><div class="ag-2">${tripTasks(t).filter(x => x.status === 'done').length}/${tripTasks(t).length} przygotowań</div></div></div>`).join('')}</section>` : ''}${quickTiles()}</aside></div>`;
}
function dashClients() {
  const ps = peopleAll().filter(inCtx).sort((a, b) => (a.last || a.first || '').localeCompare(b.last || b.first || '', 'pl')), cs = companiesAll(), P_ = S.privacy ? pv : x => x;
  const tile = x => `<div class="pfcard" data-action="person-open" data-id="${x.id}"><span class="pf-tile" style="--a:hsl(${hashS(x.id) % 360} 52% 42%)">${esc(pInitials(x))}</span>
    <div class="pf-1">${P_(esc(pName(x)))}</div><div class="pf-2">${P_(esc(x.company || (x.kind === 'ext' ? 'zewnętrzna' : 'wewnętrzna')))}</div>
    <div class="pf-acts"><button class="rbtn sm" data-action="person-task" data-id="${x.id}" data-role="from" title="Zadanie od tej osoby">⟨i:plus⟩</button><button class="rbtn sm" data-action="person-mtg" data-id="${x.id}" title="Spotkanie Outlook / Teams">⟨i:video⟩</button><button class="rbtn sm" data-action="person-mail" data-id="${x.id}" title="Wiadomość w Outlook" ${x.email ? '' : 'disabled'}>⟨i:mail⟩</button></div></div>`;
  return `<div class="row" style="margin-bottom:12px;gap:8px">${eyeBtn('sm')}<span class="spacer"></span><button class="btn" data-action="nav" data-view="people">⟨i:list⟩ Lista osób</button><button class="btn primary" data-action="dq-person">⟨i:plus⟩ Osoba</button></div>
    ${ps.length ? shelf('Osoby', `<div class="pfgrid">${ps.map(tile).join('')}</div>`) : '<div class="tl-empty">Brak osób. Dodaj klientów, kolegów i wykonawców.</div>'}
    ${cs.length ? shelf('Firmy', `<div class="pfgrid">${cs.map(clientCard).join('')}</div>`) : ''}`;
}
/* --- Dostępność (CRP): wyniki + wejście do podmodułów --- */
function vAvail() {
  const td = TODAY(), ws = iso(sow(today())), wk = weekChart(ws), tot = dayTotals(td).total, tgt = capacityOfDay(today(), 'job', true) * 60;
  const c14 = capacityRange(today(), 14), mtW = meetingHoursByDay(meetingsC().filter(m => m.start.slice(0, 10) >= ws && m.start.slice(0, 10) <= iso(add(P(ws), 6)))), mtH = Object.values(mtW).reduce((a, b) => a + b, 0);
  const ups = tripsAll().filter(t => t.to >= td && inCtx(t)).sort((a, b) => a.from.localeCompare(b.from)), vacs = (S.data.vacations || []).filter(v => v.to >= td).sort((a, b) => a.from.localeCompare(b.from)).slice(0, 3);
  const cd = capDayChart(14), pct = v => Math.round(v * 100);
  const stat = (v, l, s, view, cls = '') => `<div class="stat ${cls}" data-action="nav" data-view="${view}"><div class="st-v">${v}</div><div class="st-l">${l}</div><div class="st-s">${s}</div></div>`;
  const plan = [...Array(7)].map((_, i) => { const d = add(today(), i), ds = iso(d); const x = dayLoadInfo(ds, S.ctx === 'all' ? 'job' : S.ctx); return {d, ds, x}; });
  return `<div class="stats">
      ${stat(fmtH(tot) === '<1m' && !tot ? '0h' : fmtH(tot), 'Realizacja dziś', `target ${r1(tgt / 60)}h · ${tgt ? Math.round(100 * tot / tgt) : 0}%`, 'track')}
      ${stat(fmtH(wk.total), 'Realizacja – ten tydzień', `target ${wk.target / 60}h · ${wk.target ? Math.round(100 * wk.total / wk.target) : 0}%`, 'timelog')}
      ${stat(pct(c14.util) + '%', 'Obciążenie – 14 dni', `${r1(c14.tasks + c14.mtg)}h z ${r1(c14.cap)}h capacity`, 'capacity', c14.util > 1 ? 'bad' : c14.util > .85 ? 'warn' : 'ok')}
      ${stat(r1(mtH) + 'h', 'Spotkania w tym tygodniu', plN(meetingsC().filter(m => m.start.slice(0, 10) >= ws && m.start.slice(0, 10) <= iso(add(P(ws), 6))).length, 'spotkanie', 'spotkania', 'spotkań'), 'meetings')}
      ${stat(ups.length, 'Delegacje', ups[0] ? `${esc(ups[0].place)} · ${relDay(ups[0].from)}` : 'brak zaplanowanych', 'trips')}</div>
    <div class="avail-grid">
      <section class="pane wide"><div class="pane-h"><h3>Obciążenie – kolejne 14 dni</h3><span class="spacer"></span><button class="linkbtn" data-action="nav" data-view="capacity">Capacity ⟨i:arrow-right⟩</button></div>${cd.svg}
        <div class="row wk-legend"><span><b style="color:#5E8BFF">■</b> spotkania</span><span><b style="color:#8E929B">■</b> zadania</span><span><b style="color:#E5484D">■</b> zaległe</span><span><i class="wk-sw"></i> capacity</span><span class="spacer"></span>${cd.overDays.length ? `<span class="wk-low">Przeładowane dni: ${cd.overDays.length}</span>` : '<span class="wk-ok">Bez przeładowań</span>'}</div></section>
      <section class="pane"><div class="pane-h"><h3>Realizacja – ten tydzień</h3><span class="spacer"></span><button class="linkbtn" data-action="nav" data-view="timereport">Raport ⟨i:arrow-right⟩</button></div>${wk.svg}</section>
      <section class="pane"><div class="pane-h"><h3>Zaplanowane – 7 dni</h3><span class="spacer"></span><button class="linkbtn" data-action="nav" data-view="plan">Plan ⟨i:arrow-right⟩</button></div>
        ${plan.map(({d, ds, x}) => { const u = x.cap ? x.tot / x.cap : x.tot ? 2 : 0; return `<div class="pl-row" data-action="plan-day" data-d="${ds}"><span class="pr-d">${['Nd', 'Pn', 'Wt', 'Śr', 'Cz', 'Pt', 'Sb'][d.getDay()]} <b>${d.getDate()}.${pad(d.getMonth() + 1)}</b></span><span class="pbar ${u > 1 ? 'bad' : u > .85 ? 'warn' : ''}"><i style="width:${Math.min(100, u * 100)}%"></i></span><span class="pr-v">${r1(x.tot)} / ${r1(x.cap)}h</span>${x.trip ? '<span class="pr-tag">⟨i:plane⟩</span>' : x.type === 'vac' ? '<span class="pr-tag">⟨i:sun⟩</span>' : x.type === 'hol' ? '<span class="pr-tag">⟨i:flag⟩</span>' : '<span class="pr-tag"></span>'}</div>`; }).join('')}</section>
      <section class="pane"><div class="pane-h"><h3>Delegacje i urlopy</h3><span class="spacer"></span><button class="linkbtn" data-action="nav" data-view="trips">Delegacje ⟨i:arrow-right⟩</button></div>
        ${ups.slice(0, 3).map(t => `<div class="ag" data-action="trip-open" data-id="${t.id}"><span class="ag-dot trip"></span><div class="ag-time"><b>${fmt(t.from)}</b><small>${plN(tripDays(t), 'dzień', 'dni', 'dni')}</small></div><div class="ag-t"><div class="ag-1">⟨i:plane⟩ ${esc(t.place)}</div><div class="ag-2">przygotowanie ${tripTasks(t).filter(x => x.status === 'done').length}/${tripTasks(t).length} · capacity JOB ${tripCap(t, 'job')}%</div></div></div>`).join('')}
        ${vacs.map(v => `<div class="ag"><span class="ag-dot vac"></span><div class="ag-time"><b>${fmt(v.from)}</b><small>${vacationWorkdays(v)} dni rob.</small></div><div class="ag-t"><div class="ag-1">⟨i:sun⟩ ${esc(v.note || 'Urlop')}</div><div class="ag-2">${fmt(v.from)} – ${fmt(v.to)}</div></div></div>`).join('')}
        ${ups.length || vacs.length ? '' : '<div class="ag-empty">Brak delegacji i urlopów.</div>'}</section>
      <section class="pane"><div class="pane-h"><h3>⟨i:bolt⟩ Wprowadź dane</h3></div><div class="qtiles">${[['dq-start', '⟨i:play⟩', 'Start stopera'], ['dq-entry', '⟨i:clock⟩', 'Wpis czasu'], ['dq-trip', '⟨i:plane⟩', 'Delegacja'], ['dq-vac', '⟨i:sun⟩', 'Urlop'], ['quick', '⟨i:plus⟩', 'Zaplanuj zadanie'], ['dq-mtg', '⟨i:video⟩', 'Spotkanie']].map(([a, i, l]) => `<button class="qtile" data-action="${a}">${i}<span>${l}</span></button>`).join('')}</div></section>
    </div>`;
}
/* Zaplanowane zadania – CRP dzień po dniu: capacity vs spotkania vs zadania */
function vPlan() {
  const n = S.planDays || 14, td = TODAY(), by = {};
  tasksC().filter(t => t.status !== 'done').forEach(t => Object.entries(taskLoad(t)).forEach(([d, h]) => { if (h > .01) (by[d] ||= []).push([t, h]); }));
  const mt = meetingHoursByDay(meetingsC()), mts = {}; meetingsC().forEach(m => (mts[m.start.slice(0, 10)] ||= []).push(m));
  const rows = [...Array(n)].map((_, i) => { const d = add(today(), i), ds = iso(d), cap = capacityOfDay(d, S.ctx), l = (by[ds] || []).sort((a, b) => b[1] - a[1]), th = l.reduce((a, x) => a + x[1], 0), mh = mt[ds] || 0, free = cap - th - mh, tr = tripAt(ds);
    return `<div class="plan-row ${ds === td ? 'today' : ''} ${dayType(d) !== 'work' ? 'off' : ''}" id="pd-${ds}"><div class="pr-day"><b>${['Nd', 'Pn', 'Wt', 'Śr', 'Cz', 'Pt', 'Sb'][d.getDay()]}</b> ${d.getDate()}.${pad(d.getMonth() + 1)}${holName(d) ? `<small>${esc(holName(d))}</small>` : vacOf(d) ? '<small>urlop</small>' : tr ? `<small>⟨i:plane⟩ ${esc(tr.place)}</small>` : ''}</div>
      <div class="pr-bar"><span class="pbar ${free < -.01 ? 'bad' : cap && free < cap * .15 ? 'warn' : ''}"><i style="width:${cap ? Math.min(100, 100 * (th + mh) / cap) : th + mh ? 100 : 0}%"></i></span><small>${r1(th + mh)} / ${r1(cap)}h · <span class="${free < -.01 ? 'late' : 'muted'}">${free < -.01 ? 'brakuje ' + r1(-free) : 'wolne ' + r1(free)}h</span></small></div>
      <div class="pr-items">${(mts[ds] || []).sort((a, b) => a.start.localeCompare(b.start)).map(m => `<span class="pchip m" data-action="edit-mtg" data-id="${m.id}">⟨i:video⟩ ${hm(m.start)} ${S.privacy ? pv(esc(m.title)) : esc(m.title)}</span>`).join('')}${l.map(([t, h]) => `<span class="pchip ${isLate(t) ? 'late' : ''}" data-action="edit-task" data-id="${t.id}"><span class="dot ${t.ctx}"></span>${esc(t.title)} <b>${r1(h)}h</b></span>`).join('')}${!(mts[ds] || []).length && !l.length ? '<span class="muted sm">—</span>' : ''}</div></div>`; }).join('');
  return `<div class="row" style="margin-bottom:12px;gap:8px"><div class="chips">${[7, 14, 28].map(k => `<button class="fchip ${n === k ? 'on' : ''}" data-action="set" data-k="planDays" data-v="${k}">${k} dni</button>`).join('')}</div><span class="muted sm">Estymacja zadań rozłożona na dni od startu do terminu, minus zaraportowany czas. Spotkania nakładające się liczone raz.</span><span class="spacer"></span><button class="btn primary" data-action="quick">⟨i:plus⟩ Zaplanuj zadanie</button></div><div class="plan">${rows}</div>`;
}

Object.assign(A, {
  'nav-back': () => navGo(-1), 'nav-fwd': () => navGo(1),
  'create-menu': el => popMenu(el, [['quick', '⟨i:check-square⟩', 'Zadanie', 'Szybkie dodanie (N)'], ['dq-project', '⟨i:folder⟩', 'Projekt', 'Pusty albo z szablonu'], ['new-goal-pop', '⟨i:layers⟩', 'Podprojekt', 'Etap w projekcie'], ['dq-mtg', '⟨i:video⟩', 'Spotkanie', 'TM, Outlook lub Teams'], ['dq-person', '⟨i:users⟩', 'Osoba / kontakt', 'Klient, kolega, wykonawca'], ['dq-trip', '⟨i:plane⟩', 'Delegacja', 'Z szablonu przygotowań'], ['dq-vac', '⟨i:sun⟩', 'Urlop', ''], ['note-new-pop', '⟨i:note⟩', 'Notatka', '']]),
  'new-goal-pop': () => { closePop(); editGoal(null, S.view === 'project' && S.proj.open ? {project: S.proj.open} : {}); },
  'note-new-pop': () => { closePop(); noteEditor(null, {}); },
  'lib-today': () => { S.view = 'dashboard'; S.dash = 'over'; S.dtf = 'today'; history.replaceState(null, '', '#dashboard'); render(); },
  'lib-goal': el => { const g = goalById(el.dataset.id); if (!g) return; if (g.projectId && projById(g.projectId)) { S.view = 'project'; S.proj.open = g.projectId; history.replaceState(null, '', '#project'); render(); } editGoal(g); },
  'lib-client': el => { S.view = 'people'; (S.pf ||= {q: '', kind: '', sort: 'name'}).q = el.dataset.id; history.replaceState(null, '', '#people'); render(); $('main').scrollTop = 0; },
  'proj-play': (el, e) => { e.stopPropagation(); const p = projById(el.dataset.id); if (!p) return; const nx = S.data.tasks.filter(t => projOf(t) === p.id && t.status !== 'done').sort((a, b) => (a.status === 'doing' ? -1 : 0) - (b.status === 'doing' ? -1 : 0) || a.prio - b.prio || prioSort(a, b))[0];
    if (nx) startTimer({taskId: nx.id}); else startTimer({title: 'Praca: ' + p.title, projectId: p.id}); },
  'sr-open': el => searchOpen(SR.list[+el.dataset.n]),
  'home-proj': el => { S.homeProj = S.homeProj === el.dataset.id ? '' : el.dataset.id; S.homeGoal = ''; const y = $('main').scrollTop; render(); $('main').scrollTop = y; if (S.homeProj) $('#homeMain')?.scrollIntoView({behavior: 'smooth', block: 'start'}); },
  'free-start': el => { const p = projById(el.dataset.project || S.homeProj); startTimer({free: true, title: p ? `Praca w projekcie: ${p.title}` : 'Realizacja bez zadania', projectId: p ? p.id : null}); },
  'rcp-copy': async () => { const t = rcpText(); try { await navigator.clipboard.writeText(t); toast('Skopiowano zestawienie dnia – wklej do RCP'); } catch { const a = document.createElement('textarea'); a.value = t; document.body.append(a); a.select(); document.execCommand('copy'); a.remove(); toast('Skopiowano zestawienie dnia – wklej do RCP'); } },
  'rcp-log': () => { S.tl = Object.assign(S.tl || {}, {period: 'today', proj: '', q: ''}); S.view = 'timelog'; history.replaceState(null, '', '#timelog'); render(); $('main').scrollTop = 0; },
  'rcp-week': () => { S.view = 'timereport'; history.replaceState(null, '', '#timereport'); render(); $('main').scrollTop = 0; },
  'hero-i': el => { S.heroI = (S.heroI || 0) + +el.dataset.d; render(); },
  'row-scroll': el => { const r = $('#' + el.dataset.row); if (r) r.scrollBy({left: +el.dataset.d * r.clientWidth * .85, behavior: 'smooth'}); },
  'dash-day': el => { S.dashDay = (S.dashDay || 0) + +el.dataset.d; if (Math.abs(S.dashDay) > 60) S.dashDay = 0; render(); },
  'plan-day': el => { S.view = 'plan'; S.planDays = Math.max(S.planDays || 14, diffD(TODAY(), el.dataset.d) + 1 > 14 ? 28 : 14); history.replaceState(null, '', '#plan'); render(); $('#pd-' + el.dataset.d)?.scrollIntoView({block: 'center'}); },
  'gantt-open': el => { S.gantt.group = el.dataset.group || 'project'; S.gantt.only = el.dataset.only || null; if (el.dataset.only && projById(el.dataset.only)?.start) S.gantt.start = add(sow(P(projById(el.dataset.only).start > TODAY() ? projById(el.dataset.only).start : TODAY())), -7); S.view = 'gantt'; history.replaceState(null, '', '#gantt'); render(); $('main').scrollTop = 0; }
});

/* ================= InfoConsulting Workspace 2027 – powłoka i kokpit ================= */
const IC_MARK = '<svg viewBox="0 0 128 124" aria-hidden="true"><rect x="0" y="0" width="28" height="26"/><rect x="0" y="33" width="28" height="91"/><path d="M55 0H128V26H41V14Q41 0 55 0Z"/><path d="M41 33H69V98H128V124H56Q41 124 41 109Z"/></svg>';
const fmtHM = m => { m = Math.max(0, Math.round(m || 0)); return `${Math.floor(m / 60)}:${pad(m % 60)}`; };
const STATUS_L = {todo: 'Do zrobienia', doing: 'W toku', done: 'Zrobione'};
/* kod projektu: 2–3 litery z nazwy (znak rozpoznawczy, nie dekoracja) */
function projCode(p) {
  if (!p) return '—'; if (p.code) return String(p.code).slice(0, 4).toUpperCase();
  const w = String(p.title || '').replace(/[–—\-/:+&,.()]/g, ' ').split(/\s+/).filter(x => x.length > 1 && !/^(i|w|z|na|do|dla|od|ze|oraz|the|and|of)$/i.test(x));
  let c = w.filter(x => /^[A-ZĄĆĘŁŃÓŚŹŻ0-9]/.test(x)).map(x => x[0]).join(''); if (c.length < 2) c = w.slice(0, 3).map(x => x[0]).join('');
  return (c || '?').slice(0, 3).toUpperCase();
}
const mono = (p, cls = '') => `<span class="mono-tag ${cls}" style="--c:${p ? projColor(p) || ctxColor(p.ctx) : '#72757D'}" title="${p ? esc(p.title) : 'Bez projektu'}">${esc(projCode(p))}</span>`;
function art(seed, col, icon, cls = '') { return `<span class="tile ${cls}" style="--c:${col || '#72757D'}">${icon || ''}</span>`; }
const projArt = (p, cls) => mono(p, cls);
const kbd = k => `<kbd>${k}</kbd>`;

/* --- pasek boczny --- */
function sideHtml() {
  const td = TODAY(), late = tasksC().filter(isLate).length, mt = meetingsC().filter(m => m.start.slice(0, 10) === td).length, P_ = S.privacy ? pv : x => x;
  const badge = {tasks: late ? `<span class="nb bad" title="Zadania po terminie">${late}</span>` : '', meetings: mt ? `<span class="nb" title="Spotkania dziś">${mt}</span>` : ''};
  const groups = [['Praca', ['dashboard', 'tasks', 'meetings', 'calendar']], ['Portfolio', ['projects', 'people']], ['Zasoby', ['avail']]];
  const nav = groups.map(([g, ks]) => `<div class="sb-group"><div class="sb-gl">${g}</div>${ks.map(k => { const [i, l] = VIEWS[k], on = navRoot(S.view) === k || (k === 'tasks' && S.view === 'gantt');
    return `<button class="nv ${on ? 'on' : ''}" data-action="nav" data-view="${k}" title="${l}"><span class="ico">${i}</span><span class="nl">${k === 'dashboard' ? 'Kokpit' : l}</span>${badge[k] || ''}</button>`; }).join('')}</div>`).join('');
  const ps = projectsC().filter(p => p.status === 'active').sort((a, b) => (a.due || '9').localeCompare(b.due || '9')).slice(0, 8);
  const pins = ps.map(p => { const o = projTasks(p).filter(t => t.status !== 'done').length, lt = projTasks(p).filter(isLate).length;
    return `<button class="pin ${S.view === 'project' && S.proj.open === p.id ? 'on' : ''}" data-action="open-project" data-id="${p.id}" title="${esc(p.title)}">${mono(p, 'sm')}<span class="pin-t">${P_(esc(p.title))}</span><span class="pin-n ${lt ? 'bad' : ''}">${lt ? lt + '!' : o}</span></button>`; }).join('');
  const u = S.user;
  return `<div class="sb-brand"><span class="sb-mark" title="InfoConsulting">${IC_MARK}</span><div class="sb-name"><b>InfoConsulting</b><small>Workspace</small></div><button class="sb-coll" data-action="side-toggle" title="${S.sideColl ? 'Rozwiń' : 'Zwiń'} panel ( [ )" aria-label="Zwiń lub rozwiń panel">⟨i:sidebar⟩</button></div>
    <div class="sb-nav">${nav}
      <div class="sb-group pins"><div class="sb-gl">Aktywne projekty<button class="sb-add" data-action="create-menu" title="Utwórz nowy obiekt">⟨i:plus⟩</button></div>${pins || '<div class="sb-empty">Brak aktywnych projektów</div>'}</div></div>
    <div class="sb-foot"><button class="me" data-action="open-profile" data-tab="account" title="Konto"><span class="me-av">${u.picture ? `<img src="${esc(u.picture)}" alt="">` : esc((u.name || '?').split(/\s+/).map(x => x[0]).join('').slice(0, 2).toUpperCase())}</span><span class="me-t"><b>${esc(u.name || '')}</b><small>${esc(u.email || '')}</small></span></button>
      <button class="gearbtn ${S.view === 'settings' ? 'on' : ''}" id="gearBtn" data-action="nav" data-view="settings" title="Ustawienia aplikacji" aria-label="Ustawienia aplikacji">⟨i:gear⟩</button></div>`;
}

/* --- pasek ewidencji czasu (dół) --- */
function lastPlayed() {
  const lt = S.data.lastTask && S.data.tasks.find(x => x.id === S.data.lastTask); if (lt && lt.status !== 'done') return lt;
  const e = entriesAll().filter(x => x.taskId).sort((a, b) => (b.endMs || 0) - (a.endMs || 0))[0], t = e && S.data.tasks.find(x => x.id === e.taskId);
  return t && t.status !== 'done' ? t : null;
}
function playerHtml() {
  const t = timerOn(), r = rcpToday(), P_ = S.privacy ? pv : x => x, dp = r.tgt ? Math.min(100, 100 * r.tot / r.tgt) : 0;
  const day = `<span class="sb-day" title="Zarejestrowane dziś względem ${r1(r.tgt / 60)} h"><span class="sb-k">Dziś</span><b class="num">${fmtHM(r.tot)}</b><span class="sb-of">/ ${fmtHM(r.tgt)}</span><span class="meter"><i style="width:${dp}%"></i></span></span>`;
  if (t) {
    const task = S.data.tasks.find(x => x.id === t.taskId), p = projById(t.projectId), spent = task ? taskSpent(task.id) : (Date.now() - t.startMs) / 60000, est = (task?.est || 0) * 60;
    return `<div class="sb-l"><span class="rec" title="Rejestracja w toku"></span><span class="sb-k rec-k">Rejestracja</span>${p ? mono(p, 'xs') : ''}<span class="sb-task" data-action="${task ? 'edit-task' : 'nav'}" ${task ? `data-id="${task.id}"` : 'data-view="track"'} title="Otwórz szczegóły">${P_(esc(t.title))}</span>${p ? `<span class="sb-proj">${P_(esc(p.title))}</span>` : t.free ? '<span class="sb-proj">bez zadania</span>' : ''}</div>
      <div class="sb-m"><span class="sb-clock num" data-timer>${fmtClock(Date.now() - t.startMs)}</span>${est ? `<span class="sb-est num" title="Zarejestrowano / estymacja">${fmtHM(spent)} / ${fmtHM(est)}</span><span class="meter"><i style="width:${Math.min(100, 100 * spent / est)}%"></i></span>` : ''}</div>
      <div class="sb-r"><button class="sbtn" data-action="timer-pause" title="Zapisz czas i wstrzymaj rejestrację">Przerwa</button><button class="sbtn" data-action="timer-done" title="Zapisz czas i oznacz zadanie jako zrobione">Zakończ zadanie</button><button class="sbtn ghost" data-action="timer-switch" title="Zapisz czas i wybierz inne zadanie">Przełącz</button><span class="sb-sep"></span>${day}</div>`;
  }
  const lt = lastPlayed();
  return `<div class="sb-l"><span class="idle"></span><span class="sb-k">Brak aktywnej rejestracji</span>${lt ? `<button class="sbtn ghost" data-action="timer-start-task" data-id="${lt.id}" title="Wznów rejestrację ostatniego zadania">⟨i:timer⟩ Wznów: <span class="sb-lt">${P_(esc(lt.title))}</span></button>` : ''}</div>
    <div class="sb-m"></div>
    <div class="sb-r"><button class="sbtn" data-action="dq-start" title="Wybierz zadanie i rozpocznij rejestrację czasu">⟨i:timer⟩ Rejestruj czas ${kbd('R')}</button><button class="sbtn ghost" data-action="free-start" title="Rejestracja bez przypisanego zadania">Bez zadania</button><button class="sbtn ghost" data-action="dq-entry" title="Dodaj wpis czasu po fakcie">Wpis ręczny ${kbd('W')}</button><span class="sb-sep"></span>${day}</div>`;
}
function playerRender() { const el = $('#player'); if (el && S.user && S.data) { el.innerHTML = H(playerHtml()); el.classList.toggle('active', !!timerOn()); } }

/* --- render --- */
function render() {
  try { recurGen(); } catch (e) { console.warn('recurGen', e); }
  if (S.view === 'templates' || S.view === 'projtpl') { S.setTab = S.view === 'templates' ? 'tpl' : 'ptpl'; S.view = 'settings'; history.replaceState(null, '', '#settings'); }
  if (!VIEWS[S.view]) { S.view = 'dashboard'; history.replaceState(null, '', '#dashboard'); }
  navTrack();
  document.body.classList.toggle('side-coll', !!S.sideColl);
  $('#nav').innerHTML = H(sideHtml());
  $$('#ctxSwitch button').forEach(b => b.classList.toggle('on', b.dataset.ctx === S.ctx));
  document.body.classList.toggle('privacy', !!S.privacy);
  const av = $('#avatar'); av.innerHTML = S.user.picture ? `<img src="${esc(S.user.picture)}" alt="">` : esc((S.user.name || '?').split(/\s+/).map(x => x[0]).join('').slice(0, 2).toUpperCase());
  $('#navBack').disabled = !NAVH.back.length; $('#navFwd').disabled = !NAVH.fwd.length;
  $('#bellBtn').classList.toggle('on', S.view === 'notifs');
  { const eb = $('#eyeBtn'); if (eb) { eb.innerHTML = H(S.privacy ? '⟨i:eye-off⟩' : '⟨i:eye⟩'); eb.classList.toggle('on', !!S.privacy); eb.setAttribute('aria-pressed', S.privacy ? 'true' : 'false'); eb.title = S.privacy ? 'Dane ukryte – kliknij, aby pokazać (H)' : 'Dane widoczne – kliknij, aby ukryć (H)'; } }
  const [, title, fn, parent] = VIEWS[S.view], root = navRoot(S.view), op = S.view === 'project' && projById(S.proj.open);
  const rootL = root === 'dashboard' ? 'Kokpit' : VIEWS[root][1];
  $('#crumbs').innerHTML = H(`<button class="cr" data-action="nav" data-view="${root}">${VIEWS[root][0]}<span>${rootL}</span></button>${parent || op ? `<span class="cr-sep">/</span><span class="cr cur">${esc(op ? op.title : title)}</span>` : S.view === 'gantt' ? '<span class="cr-sep">/</span><span class="cr cur">Gantt</span>' : ''}`);
  const kids = Object.entries(VIEWS).filter(([k, v]) => v[3] === root && !v[5] && (k !== 'project' || projById(S.proj.open)));
  const subtabs = kids.length ? `<nav class="tabs">${[[root, VIEWS[root][4] || VIEWS[root][1]], ...kids.map(([k, v]) => [k, k === 'project' ? projById(S.proj.open).title : v[1]])].map(([k, l]) => `<button class="tab ${S.view === k ? 'on' : ''}" data-action="nav" data-view="${k}">${esc(l)}</button>`).join('')}</nav>` : '';
  let head = '';
  if (S.view !== 'dashboard') {
    const meta = op ? (() => { const x = projStats(op); return `${op.client ? esc(op.client) + ' · ' : ''}${op.ctx === 'priv' ? 'PRIV' : 'JOB'} · ${PSTAT[op.status]?.[0] || ''} · ${plN(x.open.length, 'otwarte zadanie', 'otwarte zadania', 'otwartych zadań')} · realizacja <span class="num">${x.pr}%</span>${op.due ? ' · koniec ' + fmt(op.due) : ''}`; })() : ctxLabel();
    const back = S.view === 'gantt' && NAVH.back.length ? `<button class="tab" data-action="nav-back">⟨i:chev-left⟩ Wróć</button>` : '';
    head = `<header class="phd">${op ? mono(op, 'lg') : `<span class="phd-ic">${VIEWS[S.view][0]}</span>`}<div class="phd-t"><h2 class="view-title">${esc(op ? op.title : title)}</h2><div class="phd-m">${meta}</div></div></header>${back || subtabs ? `<div class="phd-tabs">${back}${subtabs}</div>` : ''}`;
  }
  $('#view').innerHTML = H(head + `<div class="vbody">${fn()}</div>`);
  playerRender();
  afterRender[S.view]?.(); trkGuard(document); cloudBadge();
  document.title = `${op ? op.title : S.view === 'dashboard' ? 'Kokpit' : title} · InfoConsulting Workspace`; notifBadge();
}

/* --- paleta poleceń / wyszukiwarka (lista w stylu narzędzi deweloperskich) --- */
const CMDS = [['quick', '⟨i:check-square⟩', 'Nowe zadanie', 'N'], ['dq-start', '⟨i:timer⟩', 'Rejestruj czas', 'R'], ['free-start', '⟨i:timer⟩', 'Czas bez zadania', ''], ['dq-entry', '⟨i:clock⟩', 'Wpis ręczny', 'W'],
  ['dq-mtg', '⟨i:video⟩', 'Nowe spotkanie', 'M'], ['dq-person', '⟨i:users⟩', 'Nowa osoba', 'O'], ['dq-project', '⟨i:folder⟩', 'Nowy projekt', 'P'], ['new-goal-pop', '⟨i:layers⟩', 'Nowy podprojekt', ''], ['dq-vac', '⟨i:sun⟩', 'Urlop', ''], ['dq-trip', '⟨i:plane⟩', 'Delegacja', ''], ['rcp-copy', '⟨i:check⟩', 'Kopiuj dzień do RCP', ''], ['sch-open', '⟨i:schedule⟩', 'Harmonogram zadań', ''], ['sch-auto', '⟨i:refresh⟩', 'Ułóż harmonogram (2 tygodnie)', ''], ['rec-new', '⟨i:repeat⟩', 'Nowe zadanie / spotkanie cykliczne', ''], ['tl-sync', '⟨i:video⟩', 'Synchronizuj Teams (spotkania, osoby, zadania)', '']];
function srItem(r, i) {
  const P_ = S.privacy && !['nav', 'client', 'cmd'].includes(r.k) ? pv : x => x;
  const a = r.k === 'project' ? mono(projById(r.id), 'sm') : r.round ? `<span class="av-i" style="--c:hsl(${r.hue} 40% 45%)">${esc(r.round)}</span>` : `<span class="sr-ic">${r.ic || ''}</span>`;
  return `<div class="sr-i ${i === SR.i ? 'on' : ''} ${r.done ? 'done' : ''}" data-action="sr-open" data-n="${i}">${a}<div class="sr-t"><span class="sr-1">${P_(esc(r.t))}</span><span class="sr-2">${esc(r.k === 'cmd' ? '' : r.sub.split(' · ').slice(1).join(' · '))}</span></div><span class="sr-type">${r.key ? kbd(r.key) : esc(r.k === 'cmd' ? 'Polecenie' : r.sub.split(' · ')[0])}</span></div>`;
}
function searchRender() {
  const box = $('#searchDrop'); if (!box) return;
  const q = $('#gsearch').value; SR.q = q; const n = norm(q).trim(), list = [];
  const cmds = CMDS.filter(c => !n || norm(c[2]).includes(n)).map(([id, ic, t, key]) => ({k: 'cmd', id, ic, t, sub: 'Polecenie', key}));
  const gh = (l, items) => items.length ? `<div class="sr-h">${l}</div>` + items.map(r => { list.push(r); return srItem(r, list.length - 1); }).join('') : '';
  let html = '';
  if (!n) { const rec = (LS.get('tm:recent', []) || []).slice(0, 5); html = gh('Ostatnio otwierane', rec) + gh('Polecenia', cmds.slice(0, 8)); }
  else { const all = searchAll(q), cmdFirst = cmds.some(c => norm(c.t).startsWith(n)), cg = () => gh('Polecenia', cmds.slice(0, 5)); html = (cmdFirst ? cg() : '') + SR_G.map(([k, l]) => gh(l, all.filter(r => r.k === k).slice(0, k === 'task' ? 6 : 4))).join('') + (cmdFirst ? '' : cg()); if (!all.length && !cmds.length) html = `<div class="sr-empty">Brak wyników dla „${esc(q)}”. Spróbuj nazwy projektu, klienta, IC case lub osoby.</div>`; }
  if (SR.i >= list.length) SR.i = 0;
  SR.list = list; box.hidden = false;
  box.innerHTML = H(html + `<div class="sr-foot"><span>${kbd('↑')}${kbd('↓')} wybór</span><span>${kbd('Enter')} otwórz</span><span>${kbd('Esc')} zamknij</span><span class="spacer"></span><span>${kbd('Ctrl')}${kbd('K')} w dowolnym miejscu</span></div>`);
  $$('#searchDrop .sr-i').forEach(x => x.classList.toggle('on', +x.dataset.n === SR.i));
}

/* --- lista zadań (tabela operacyjna) --- */
function trackList(ts, opts = {}) {
  const run = timerOn()?.taskId, P_ = S.privacy ? pv : x => x;
  const head = opts.noHead ? '' : `<div class="trow th"><span></span><span>Zadanie</span><span class="c-proj">Projekt</span><span class="c-st">Status</span><span class="c-prio">Prior.</span><span class="c-due">Termin</span><span class="c-time">Czas</span><span></span></div>`;
  return `<div class="tlist">${head}${ts.map(t => { const p = projById(projOf(t)), g = t.goalId && goalById(t.goalId), on = run === t.id, done = t.status === 'done', sp = taskSpent(t.id), est = (t.est || 0) * 60, ap = personById(t.assigneeId), rq = personById(t.requesterId);
    return `<div class="trow ${done ? 'done' : ''} ${on ? 'rec-on' : ''} ${isLate(t) ? 'is-late' : ''}"><span class="tchk ${done ? 'on' : ''}" data-action="toggle" data-id="${t.id}" title="${done ? 'Przywróć do realizacji' : 'Oznacz jako zrobione'}" role="checkbox" aria-checked="${done}">${done ? '⟨i:check⟩' : ''}</span>
      <div class="ttl" data-action="edit-task" data-id="${t.id}" title="Otwórz szczegóły zadania"><div class="t1">${on ? '<span class="rec"></span>' : ''}${P_(esc(t.title))}</div><div class="t2">${g ? `<span>${P_(esc(g.title))}</span>` : ''}${t.ic ? `<span class="num">${esc(t.ic)}</span>` : ''}${ap ? `<span>→ ${P_(esc(pName(ap)))}</span>` : ''}${rq ? `<span>← ${P_(esc(pName(rq)))}</span>` : ''}${!g && !t.ic && !ap && !rq ? `<span>${t.ctx === 'priv' ? 'PRIV' : 'JOB'}</span>` : ''}</div></div>
      <span class="c-proj">${p ? `<a href="#" data-action="open-project" data-id="${p.id}" title="${esc(p.title)}">${mono(p, 'xs')}<span>${P_(esc(p.title))}</span></a>` : '<span class="dim">—</span>'}</span>
      <span class="c-st"><button class="st st-${t.status}" data-action="task-status-menu" data-id="${t.id}" title="Zmień status">${STATUS_L[t.status] || t.status}</button></span>
      <span class="c-prio"><span class="pr pr${t.prio}" title="Priorytet ${PRIO_L[t.prio] || ''}">P${t.prio}</span></span>
      <span class="c-due ${isLate(t) ? 'late' : t.due === TODAY() ? 'today' : ''}">${t.due ? relDay(t.due) : '<span class="dim">—</span>'}</span>
      <span class="c-time num">${sp || est ? `${sp ? fmtHM(sp) : '0:00'}${est ? `<span class="dim"> / ${fmtHM(est)}</span>` : ''}${est ? `<span class="meter xs"><i style="width:${Math.min(100, 100 * sp / est)}%"></i></span>` : ''}` : '<span class="dim">—</span>'}</span>
      <span class="t-act">${done ? '' : on ? `<button class="sbtn sm" data-action="timer-pause" title="Zapisz czas i wstrzymaj">Przerwa</button>` : `<button class="sbtn sm" data-action="timer-start-task" data-id="${t.id}" title="Rozpocznij rejestrację czasu na tym zadaniu">⟨i:timer⟩ Rejestruj</button>`}</span></div>`; }).join('') || `<div class="empty-st">${opts.empty || 'Brak zadań'}</div>`}</div>`;
}

/* --- KOKPIT --- */
function vDashboard() {
  return cockHead() + `<div class="ck"><div class="ck-main">${nextAction()}${schDayPanel()}${taskPanel()}${portfolio()}</div><div class="ck-side">${timePanel()}${meetPanel()}${quickPanel()}${awayPanel()}</div></div>`;
}
function cockHead() {
  const d = new Date(), td = TODAY(), h = d.getHours(), hi = h < 5 ? 'Dobranoc' : h < 12 ? 'Dzień dobry' : h < 18 ? 'Dzień dobry' : 'Dobry wieczór';
  const open = tasksC().filter(t => t.status !== 'done'), tdy = open.filter(t => t.due === td).length, late = open.filter(isLate).length;
  const ms = meetingsC().filter(m => m.start.slice(0, 10) === td).sort((a, b) => a.start.localeCompare(b.start)), nowT = d.toTimeString().slice(0, 5), nm = ms.find(m => m.end.slice(11) > nowT), tr = tripOf(td), P_ = S.privacy ? pv : x => x;
  const {load} = capLoads(), cap = capacityOfDay(today(), S.ctx), mh = meetingHoursByDay(meetingsC())[td] || 0, th = load[td] || 0, tot = mh + th, u = cap ? tot / cap : 0;
  const ctx = [`${plN(tdy, 'zadanie', 'zadania', 'zadań')} na dziś${late ? `, <span class="late">${late} po terminie</span>` : ''}`, ms.length ? `${plN(ms.length, 'spotkanie', 'spotkania', 'spotkań')}${nm ? `, najbliższe o <span class="num">${hm(nm.start)}</span>` : ', wszystkie za Tobą'}` : 'bez spotkań', tr ? `<a href="#" data-action="trip-open" data-id="${tr.id}">delegacja: ${P_(esc(tr.place))}</a>` : ''].filter(Boolean).join('<span class="dot-sep">·</span>');
  const segs = Math.max(1, Math.round(cap)), fill = tot;
  return `<header class="ck-head"><div class="ck-hl"><div class="ck-date">${DOWL[d.getDay()]}, ${d.getDate()} ${MONS[d.getMonth()]} ${d.getFullYear()}<span class="dot-sep">·</span>tydzień <span class="num">${isoWeek(d)}</span></div>
      <h1 class="ck-greet">${hi}${firstName() ? ', ' + esc(firstName()) : ''}</h1><div class="ck-ctx">${ctx}</div></div>
    <div class="ck-load" title="Planowane obciążenie dnia: spotkania ${r1(mh)} h + zadania ${r1(th)} h przy dostępności ${r1(cap)} h"><div class="ld-top"><span class="ld-l">Obciążenie dnia</span><span class="ld-v num ${u > 1 ? 'late' : ''}">${r1(tot).toString().replace('.', ',')}<span class="dim"> / ${r1(cap).toString().replace('.', ',')} h</span></span></div>
      <div class="ld-bar">${Array.from({length: segs}, (_, i) => { const f = Math.max(0, Math.min(1, fill - i)), m = Math.max(0, Math.min(1, mh - i)); return `<i><b style="width:${f * 100}%"></b>${m > 0 ? `<em style="width:${m * 100}%"></em>` : ''}</i>`; }).join('')}</div>
      <div class="ld-s"><span><em class="lg-m"></em>spotkania ${r1(mh)} h</span><span><em class="lg-t"></em>zadania ${r1(th)} h</span>${u > 1 ? '<span class="late">przeładowany</span>' : `<span>wolne ${r1(Math.max(0, cap - tot))} h</span>`}</div></div></header>`;
}
/* jedno dominujące pole: co zrobić teraz */
function naCandidates() {
  const td = TODAY(), open = tasksC().filter(t => t.status !== 'done');
  return [...open.filter(isLate).sort((a, b) => a.prio - b.prio || a.due.localeCompare(b.due)), ...open.filter(t => t.due === td).sort((a, b) => a.prio - b.prio), ...open.filter(t => t.status === 'doing' && !isLate(t) && t.due !== td), ...open.filter(t => t.due > td).sort((a, b) => a.due.localeCompare(b.due) || a.prio - b.prio)].filter((t, i, a) => a.indexOf(t) === i);
}
function nextAction() {
  const t = timerOn(), now = new Date(), nowS = `${TODAY()}T${now.toTimeString().slice(0, 5)}`, P_ = S.privacy ? pv : x => x;
  const head = (k, why) => `<div class="na-k"><span class="na-mark"></span>${k}${why ? `<span class="na-why">${why}</span>` : ''}</div>`;
  if (t) {
    const task = S.data.tasks.find(x => x.id === t.taskId), p = projById(t.projectId), sp = task ? taskSpent(task.id) : 0, est = (task?.est || 0) * 60;
    return `<section class="na na-rec">${head('Trwa rejestracja czasu', `od ${new Date(t.startMs).toTimeString().slice(0, 5)}`)}
      <div class="na-b"><div class="na-main">${mono(p, 'md')}<div class="na-txt"><h2 class="na-t">${P_(esc(t.title))}</h2><div class="na-m">${p ? P_(esc(p.title)) : t.mtgId ? 'Spotkanie – czas trafi do RCP' : t.free ? 'Realizacja bez zadania' : 'Bez projektu'}${est ? `<span class="dot-sep">·</span><span class="num">${fmtHM(sp)} / ${fmtHM(est)}</span> zarejestrowano` : ''}</div></div></div>
        <div class="na-clock num" data-timer>${fmtClock(Date.now() - t.startMs)}</div></div>
      <div class="na-acts"><button class="pbtn" data-action="timer-pause">Przerwa – zapisz czas</button>${task ? `<button class="sbtn" data-action="timer-done">Zakończ zadanie</button><button class="sbtn ghost" data-action="edit-task" data-id="${task.id}">Szczegóły zadania</button>` : ''}<button class="sbtn ghost" data-action="timer-switch">Przełącz zadanie</button></div></section>`;
  }
  const ms = meetingsC().filter(m => m.start.slice(0, 10) === TODAY() && m.end > nowS).sort((a, b) => a.start.localeCompare(b.start)), m = ms[0];
  const mins = m ? Math.round((strMs(m.start) - Date.now()) / 60000) : 999;
  if (m && mins <= 30) {
    const p = projById(m.projectId), nn = notesOf('mtg', m.id).length;
    return `<section class="na na-mtg">${head(mins <= 0 ? 'Spotkanie trwa' : `Spotkanie za ${mins} min`, `${hm(m.start)}–${hm(m.end)}`)}
      <div class="na-b"><div class="na-main"><span class="na-ic">⟨i:video⟩</span><div class="na-txt"><h2 class="na-t">${P_(esc(m.title))}</h2><div class="na-m">${p ? P_(esc(p.title)) : 'bez projektu'}${m.location ? `<span class="dot-sep">·</span>${P_(esc(m.location))}` : ''}${m.link ? '<span class="dot-sep">·</span>online' : ''}${(m.attendees || []).length ? `<span class="dot-sep">·</span>${plN(m.attendees.length, 'uczestnik', 'uczestników', 'uczestników')}` : ''}</div></div></div></div>
      <div class="na-acts">${m.link ? joinLink(m, 'Dołącz do spotkania', 'pbtn') : ''}${mtgTrkBtn(m, 'sbtn')}<button class="sbtn" data-action="edit-mtg" data-id="${m.id}">Agenda i notatki${nn ? ` (${nn})` : ''}</button>${p ? `<button class="sbtn ghost" data-action="open-project" data-id="${p.id}">Projekt</button>` : ''}</div></section>`;
  }
  { const {cur, next} = schNow(), sb = cur || (next && hmToMin(next.s) - nowMin() <= 20 ? next : null), st = sb && schTask(sb.taskId);
    if (st && st.status !== 'done') { const p = projById(projOf(st)), sp = taskSpent(st.id), est = (st.est || 0) * 60, inM = hmToMin(sb.s) - nowMin();
      return `<section class="na na-sch">${head(cur ? 'Teraz według harmonogramu' : `Za ${inM} min według harmonogramu`, `${sb.s}–${sb.e}`)}
      <div class="na-b"><div class="na-main">${mono(p, 'md')}<div class="na-txt"><h2 class="na-t" data-action="edit-task" data-id="${st.id}">${P_(esc(st.title))}</h2><div class="na-m">${p ? P_(esc(p.title)) : 'Bez projektu'}${st.due ? `<span class="dot-sep">·</span>termin ${relDay(st.due)}` : ''}${est || sp ? `<span class="dot-sep">·</span><span class="num">${fmtHM(sp)}${est ? ' / ' + fmtHM(est) : ''}</span>` : ''}<span class="dot-sep">·</span>blok to plan – czas liczy się po starcie realizacji</div></div></div></div>
      <div class="na-acts"><button class="pbtn" data-action="timer-start-task" data-id="${st.id}">⟨i:timer⟩ Rozpocznij realizację ${kbd('R')}</button><button class="sbtn" data-action="sch-edit" data-id="${sb.id}">Przesuń blok</button><button class="sbtn ghost" data-action="edit-task" data-id="${st.id}">Szczegóły</button></div></section>`; } }
  const c = naCandidates(), x = c[0];
  if (!x) return `<section class="na na-empty">${head('Następne działanie')}<div class="na-b"><div class="na-main"><span class="na-ic">⟨i:check⟩</span><div class="na-txt"><h2 class="na-t">Brak otwartych zadań</h2><div class="na-m">Plan jest pusty. Dodaj zadanie albo zarejestruj czas bez zadania.</div></div></div></div><div class="na-acts"><button class="pbtn" data-action="quick">Nowe zadanie ${kbd('N')}</button><button class="sbtn" data-action="free-start">Rejestruj bez zadania</button></div></section>`;
  const p = projById(projOf(x)), g = x.goalId && goalById(x.goalId), sp = taskSpent(x.id), est = (x.est || 0) * 60;
  const why = isLate(x) ? `<span class="late">po terminie ${relDay(x.due)}</span>` : x.due === TODAY() ? 'termin dziś' : x.due ? 'termin ' + relDay(x.due) : 'w toku';
  return `<section class="na">${head('Następne działanie', `${why} · P${x.prio}`)}
    <div class="na-b"><div class="na-main">${mono(p, 'md')}<div class="na-txt"><h2 class="na-t" data-action="edit-task" data-id="${x.id}">${P_(esc(x.title))}</h2><div class="na-m">${p ? P_(esc(p.title)) : 'Bez projektu'}${g ? `<span class="dot-sep">·</span>${P_(esc(g.title))}` : ''}${est || sp ? `<span class="dot-sep">·</span><span class="num">${fmtHM(sp)}${est ? ' / ' + fmtHM(est) : ''}</span>` : ''}${m ? `<span class="dot-sep">·</span>do spotkania <span class="num">${hm(m.start)}</span> zostało ${fmtH(mins)}` : ''}</div></div></div></div>
    <div class="na-acts"><button class="pbtn" data-action="timer-start-task" data-id="${x.id}">⟨i:timer⟩ Rozpocznij rejestrację ${kbd('R')}</button><button class="sbtn" data-action="edit-task" data-id="${x.id}">Szczegóły</button><button class="sbtn ghost" data-action="toggle" data-id="${x.id}">⟨i:check⟩ Oznacz jako zrobione</button>
      ${c.length > 1 ? `<span class="na-then">Potem: ${c.slice(1, 3).map(y => `<a href="#" data-action="edit-task" data-id="${y.id}">${P_(esc(y.title))}</a>`).join(', ')}</span>` : ''}</div></section>`;
}
function taskPanel() {
  const td = TODAY(), we = iso(add(sow(today()), 6)), open = tasksC().filter(t => t.status !== 'done'), F = ['today', 'late', 'week', 'p1'].includes(S.dtf) ? S.dtf : 'today';
  const L = {today: [...open.filter(isLate), ...open.filter(t => t.due === td), ...open.filter(t => t.status === 'doing' && !isLate(t) && t.due !== td)].filter((t, i, a) => a.indexOf(t) === i).sort((a, b) => a.prio - b.prio || prioSort(a, b)),
    late: open.filter(isLate).sort((a, b) => a.prio - b.prio || prioSort(a, b)), week: open.filter(t => t.due && t.due >= td && t.due <= we).sort(prioSort), p1: open.filter(t => t.prio === 1).sort(prioSort)};
  const tabs = [['today', 'Na dziś'], ['late', 'Po terminie'], ['week', 'Ten tydzień'], ['p1', 'Priorytet P1']];
  return `<section class="panel"><div class="pn-h"><h3>Zadania</h3><div class="seg2">${tabs.map(([k, l]) => `<button class="${F === k ? 'on' : ''} ${k === 'late' && L.late.length ? 'has-late' : ''}" data-action="set" data-k="dtf" data-v="${k}">${l}<span class="num">${L[k].length}</span></button>`).join('')}</div><span class="spacer"></span><button class="sbtn ghost sm" data-action="quick">⟨i:plus⟩ Nowe ${kbd('N')}</button><a href="#" class="pn-link" data-action="nav" data-view="tasks">Wszystkie zadania ⟨i:arrow-right⟩</a></div>
    ${trackList(L[F].slice(0, 10), {empty: {today: 'Na dziś nie ma zaplanowanych zadań. Sprawdź „Ten tydzień” albo dodaj zadanie.', late: 'Nic nie jest po terminie.', week: 'Brak zadań z terminem w tym tygodniu.', p1: 'Brak zadań z priorytetem P1.'}[F]})}
    ${L[F].length > 10 ? `<a href="#" class="pn-more" data-action="nav" data-view="tasks">Pokaż pozostałe ${L[F].length - 10} ⟨i:arrow-right⟩</a>` : ''}</section>`;
}
function timePanel() {
  const r = rcpToday(), wk = weekChart(iso(sow(today()))), miss = r.tgt - r.tot, P_ = S.privacy ? pv : x => x, segs = Math.max(1, Math.round(r.tgt / 60));
  return `<section class="panel tp"><div class="pn-h"><h3>Czas pracy</h3><span class="spacer"></span><button class="sbtn ghost sm" data-action="rcp-log" title="Ewidencja czasu z dziś">Ewidencja</button></div>
    <div class="tp-top"><div class="tp-big num">${fmtHM(r.tot)}<span class="dim"> / ${fmtHM(r.tgt)}</span></div><div class="tp-side"><div><span class="dim">Do RCP</span> <b class="num">${hDec(r.tot)} h</b></div><div class="${miss > 0 ? '' : 'ok'}"><span class="dim">${miss > 0 ? 'Brakuje' : 'Nadwyżka'}</span> <b class="num">${fmtHM(Math.abs(miss))}</b></div></div></div>
    <div class="hbar" title="Każdy segment to 1 h">${Array.from({length: segs}, (_, i) => `<i><b style="width:${Math.max(0, Math.min(1, r.tot / 60 - i)) * 100}%"></b></i>`).join('')}</div>
    <div class="tp-week"><span class="dim">Tydzień</span><span class="num">${fmtHM(wk.total)} / ${fmtHM(wk.target)}</span><span class="meter"><i style="width:${wk.target ? Math.min(100, 100 * wk.total / wk.target) : 0}%"></i></span></div>
    <div class="tp-list">${r.rows.slice(0, 5).map(x => `<div class="tp-row">${x.live ? '<span class="rec"></span>' : mono(projById(x.pid), 'xs')}<span class="tp-t" title="${esc(x.title)}">${P_(esc(x.title))}${x.free ? '<span class="dim"> · bez zadania</span>' : ''}</span><span class="num">${fmtHM(x.min)}</span></div>`).join('') || '<div class="empty-st sm">Dziś nie ma jeszcze wpisów. Rozpocznij rejestrację albo dodaj wpis ręczny.</div>'}${r.rows.length > 5 ? `<div class="dim sm">+ ${r.rows.length - 5} kolejnych pozycji</div>` : ''}</div>
    <div class="tp-acts"><button class="sbtn" data-action="dq-start">⟨i:timer⟩ Rejestruj ${kbd('R')}</button><button class="sbtn ghost" data-action="free-start">Bez zadania</button><button class="sbtn ghost" data-action="dq-entry">Wpis ${kbd('W')}</button><button class="sbtn ghost" data-action="rcp-copy" ${r.rows.length ? '' : 'disabled'} title="Kopiuj zestawienie dnia do schowka">Kopiuj do RCP</button></div></section>`;
}
function meetPanel() {
  const off = S.dashDay || 0, ds = iso(add(today(), off)), nowS = new Date().toTimeString().slice(0, 5), P_ = S.privacy ? pv : x => x;
  const ms = meetingsC().filter(m => m.start.slice(0, 10) === ds).sort((a, b) => a.start.localeCompare(b.start));
  const lbl = off === 0 ? 'Dziś' : off === 1 ? 'Jutro' : off === -1 ? 'Wczoraj' : `${DOWL[P(ds).getDay()]} ${fmt(ds)}`;
  let nowDone = off !== 0;
  const nowLine = () => `<div class="tl-now"><span class="num">${nowS}</span></div>`;
  const items = ms.map(m => { const past = off < 0 || (off === 0 && m.end.slice(11) <= nowS), live = off === 0 && m.start.slice(11) <= nowS && m.end.slice(11) > nowS, p = projById(m.projectId), nn = notesOf('mtg', m.id).length;
    const att = (m.attendees || []).map(e => { const ps = peopleAll().find(x => x.email && norm(x.email) === norm(e)); return ps ? pInitials(ps) : String(e)[0].toUpperCase(); });
    let pre = ''; if (!nowDone && !live && m.start.slice(11) > nowS) { pre = nowLine(); nowDone = true; }
    if (live) nowDone = true;
    return pre + `<div class="tl-i ${past ? 'past' : ''} ${live ? 'live' : ''}" data-action="edit-mtg" data-id="${m.id}"><div class="tl-time num">${hm(m.start)}<span>${hm(m.end)}</span></div><span class="tl-node"></span>
      <div class="tl-c"><div class="tl-1">${P_(esc(m.title))}</div><div class="tl-2">${p ? `${mono(p, 'xs')}<span>${P_(esc(p.client || p.title))}</span>` : `<span>${m.ctx === 'priv' ? 'PRIV' : 'JOB'}</span>`}${m.location ? `<span>${P_(esc(m.location))}</span>` : m.link ? '<span>online</span>' : ''}</div>
        <div class="tl-3">${att.length ? `<span class="avs">${att.slice(0, 4).map(a => `<i>${esc(a)}</i>`).join('')}${att.length > 4 ? `<i class="more">+${att.length - 4}</i>` : ''}</span>` : ''}${nn ? `<span class="tl-n" title="Notatki">⟨i:note⟩ ${nn}</span>` : ''}<span class="spacer"></span>${mtgTrkBtn(m, 'sbtn sm')}${m.link && !past ? joinLink(m, 'Dołącz', 'sbtn sm') : ''}</div></div></div>`; }).join('');
  return `<section class="panel mp"><div class="pn-h"><h3>Spotkania</h3><span class="pn-sub">${lbl}</span><span class="spacer"></span><button class="ib" data-action="dash-day" data-d="-1" title="Poprzedni dzień">⟨i:chev-left⟩</button><button class="ib" data-action="dash-day" data-d="1" title="Następny dzień">⟨i:chev-right⟩</button></div>
    ${ms.length ? `<div class="tl">${items}${!nowDone ? nowLine() : ''}</div>` : `<div class="empty-st">${off === 0 ? 'Brak spotkań dziś – dobry dzień na pracę skupioną.' : 'Brak spotkań w tym dniu.'}</div>`}
    <div class="pn-f"><button class="sbtn ghost sm" data-action="dq-mtg">⟨i:plus⟩ Spotkanie ${kbd('M')}</button><a href="#" class="pn-link" data-action="nav" data-view="calendar">Kalendarz ⟨i:arrow-right⟩</a></div></section>`;
}
function portfolio() {
  const ps = projectsC().filter(p => p.status === 'active' || p.status === 'planned'), P_ = S.privacy ? pv : x => x, td = TODAY();
  if (!ps.length) return `<section class="panel"><div class="pn-h"><h3>Portfolio projektów</h3></div><div class="empty-st">Brak aktywnych projektów. <button class="sbtn sm" data-action="dq-project">Nowy projekt</button></div></section>`;
  const info = p => { const x = projStats(p), ng = projGoals(p).filter(g => goalProgress(g) < 100).sort(goalSort)[0], gap = x.ex - x.pr, days = p.due ? diffD(td, p.due) : null;
    const sig = [x.late.length ? `${plN(x.late.length, 'zadanie', 'zadania', 'zadań')} po terminie` : '', p.status === 'active' && gap >= 10 ? `opóźnienie względem planu ${gap} pp` : '', days != null && days >= 0 && days <= 14 && p.status === 'active' ? `koniec za ${plN(days, 'dzień', 'dni', 'dni')}` : '', days != null && days < 0 && p.status !== 'done' ? 'termin projektu minął' : ''].filter(Boolean);
    return {p, x, ng, gap, days, sig, score: x.late.length * 3 + (gap >= 10 ? 2 : 0) + (days != null && days <= 14 ? 1 : 0) + (days != null && days < 0 ? 3 : 0)}; };
  const all = ps.map(info).sort((a, b) => b.score - a.score || (a.p.due || '9').localeCompare(b.p.due || '9')), f = all[0].score > 0 ? all[0] : null, rest = f ? all.slice(1) : all;
  const prog = (x, w) => `<span class="pg ${w || ''}" title="Realizacja ${x.pr}% · plan wg czasu ${x.ex}%"><i style="width:${x.pr}%"></i><b style="left:${x.ex}%"></b></span>`;
  const focus = f ? `<div class="pf-focus"><div class="pf-fk"><span class="flag">Wymaga uwagi</span>${f.sig.map(s => `<span class="sig">${s}</span>`).join('')}</div>
    <div class="pf-fb">${mono(f.p, 'md')}<div class="pf-ft"><h4 data-action="open-project" data-id="${f.p.id}">${P_(esc(f.p.title))}</h4><div class="dim">${f.p.client ? P_(esc(f.p.client)) + ' · ' : ''}${PSTAT[f.p.status]?.[0] || ''}${f.p.due ? ' · koniec ' + fmt(f.p.due) : ''}</div></div>
      <div class="pf-kv"><div><span class="dim">Realizacja</span><b class="num">${f.x.pr}%</b></div><div><span class="dim">Plan</span><b class="num">${f.x.ex}%</b></div><div><span class="dim">Otwarte</span><b class="num">${f.x.open.length}</b></div><div><span class="dim">Pozostało</span><b class="num">${r1(f.x.rest)} h</b></div></div></div>
    ${prog(f.x, 'lg')}${f.ng ? `<div class="pf-ng"><span class="dim">Najbliższy podprojekt</span> ${P_(esc(f.ng.title))} · <span class="num">${fmt(f.ng.due)}</span> · ${goalProgress(f.ng)}%</div>` : ''}
    <div class="pf-acts"><button class="sbtn" data-action="open-project" data-id="${f.p.id}">Otwórz projekt</button><button class="sbtn ghost" data-action="proj-tasks" data-id="${f.p.id}">Zadania projektu</button><button class="sbtn ghost" data-action="proj-play" data-id="${f.p.id}">⟨i:timer⟩ Rejestruj w projekcie</button></div></div>` : '';
  const rows = rest.map(({p, x, ng, sig}) => `<div class="pf-row" data-action="open-project" data-id="${p.id}">${mono(p, 'sm')}<div class="pf-n"><b>${P_(esc(p.title))}</b><span class="dim">${P_(esc(p.client || (p.ctx === 'priv' ? 'prywatny' : 'firmowy')))}</span></div>
      <span class="pf-st st-${p.status}">${PSTAT[p.status]?.[0] || p.status}</span><span class="pf-pg">${prog(x)}<span class="num">${x.pr}%</span></span><span class="num pf-o" title="Otwarte zadania">${x.open.length}</span>
      <span class="pf-d num">${ng ? fmt(ng.due) : p.due ? fmt(p.due) : '—'}</span><span class="pf-sig">${sig.length ? `<span class="sig" title="${esc(sig.join(' · '))}">${sig[0]}</span>` : '<span class="dim">bez uwag</span>'}</span></div>`).join('');
  return `<section class="panel"><div class="pn-h"><h3>Portfolio projektów</h3><span class="pn-sub">${plN(ps.length, 'projekt', 'projekty', 'projektów')}</span><span class="spacer"></span><a href="#" class="pn-link" data-action="nav" data-view="projects">Wszystkie projekty ⟨i:arrow-right⟩</a></div>
    ${focus}${rows ? `<div class="pf-list"><div class="pf-row th"><span></span><span>Projekt</span><span>Status</span><span>Postęp</span><span title="Otwarte zadania">Otw.</span><span>Najbl. termin</span><span>Sygnał</span></div>${rows}</div>` : ''}</section>`;
}
function quickPanel() {
  return `<section class="panel qp"><div class="pn-h"><h3>Szybkie akcje</h3><span class="spacer"></span><span class="dim sm">${kbd('Ctrl')}${kbd('K')} wszystkie</span></div>
    <div class="qlist">${CMDS.filter(c => c[0] !== 'rcp-copy' && c[0] !== 'new-goal-pop').map(([a, i, l, k]) => `<button class="qi" data-action="${a}"><span class="qi-ic">${i}</span><span class="qi-l">${l}</span>${k ? kbd(k) : ''}</button>`).join('')}</div></section>`;
}
function awayPanel() {
  const td = TODAY(), lim = iso(add(today(), 21)), tr = tripsAll().filter(t => t.to >= td && t.from <= lim && inCtx(t)).sort((a, b) => a.from.localeCompare(b.from)), vc = (S.data.vacations || []).filter(v => v.to >= td && v.from <= lim).sort((a, b) => a.from.localeCompare(b.from));
  if (!tr.length && !vc.length) return '';
  return `<section class="panel"><div class="pn-h"><h3>Nieobecności</h3><span class="pn-sub">21 dni</span><span class="spacer"></span><a href="#" class="pn-link" data-action="nav" data-view="avail">Dostępność ⟨i:arrow-right⟩</a></div>
    ${tr.map(t => { const ts = tripTasks(t), dn = ts.filter(x => x.status === 'done').length; return `<div class="aw" data-action="trip-open" data-id="${t.id}"><span class="aw-ic">⟨i:plane⟩</span><div class="aw-t"><b>${S.privacy ? pv(esc(t.place)) : esc(t.place)}</b><span class="dim">${tripRange(t)} · przygotowanie ${dn}/${ts.length}</span></div><span class="num dim">${relDay(t.from)}</span></div>`; }).join('')}
    ${vc.map(v => `<div class="aw"><span class="aw-ic">⟨i:sun⟩</span><div class="aw-t"><b>${esc(v.note || 'Urlop')}</b><span class="dim">${fmt(v.from)} – ${fmt(v.to)} · ${vacationWorkdays(v)} dni rob.</span></div><span class="num dim">${relDay(v.from)}</span></div>`).join('')}</section>`;
}
/* menu statusu zadania (osobna akcja od otwierania szczegółów) */
function statusMenu(el) {
  const t = S.data.tasks.find(x => x.id === el.dataset.id); if (!t) return;
  const r = el.getBoundingClientRect();
  $('#popRoot').innerHTML = H(`<div class="pop" style="left:${Math.min(r.left, innerWidth - 220)}px;top:${r.bottom + 4}px;min-width:200px">${Object.entries(STATUS_L).map(([k, l]) => `<button class="pop-i ${t.status === k ? 'cur' : ''}" data-action="task-set-status" data-id="${t.id}" data-s="${k}"><span class="st st-${k}">${l}</span>${t.status === k ? '<span class="spacer"></span>⟨i:check⟩' : ''}</button>`).join('')}</div>`);
}
Object.assign(A, {
  'side-toggle': () => { S.sideColl = !S.sideColl; LS.set('tm:side', S.sideColl ? 1 : 0); render(); },
  'task-status-menu': (el, e) => { e?.stopPropagation(); statusMenu(el); },
  'task-set-status': el => { closePop(); const t = S.data.tasks.find(x => x.id === el.dataset.id); if (!t || t.status === el.dataset.s) return; if (el.dataset.s === 'done' && timerOn()?.taskId === t.id) return stopTimer('done'); setStatus(t, el.dataset.s); },
  'proj-tasks': el => { S.tasks.proj = el.dataset.id; S.tasks.goal = ''; S.tasks.status = 'open'; S.view = 'tasks'; history.replaceState(null, '', '#tasks'); render(); $('main').scrollTop = 0; }
});
S.sideColl = !!LS.get('tm:side', 0);
document.addEventListener('keydown', e => {
  if (!S.user || e.ctrlKey || e.metaKey || e.altKey || /INPUT|TEXTAREA|SELECT/.test(document.activeElement?.tagName) || $('#modalRoot').innerHTML) return;
  const k = e.key.toLowerCase(), m = {h: () => A.privacy(), r: () => startPicker(), w: () => editEntry(null), m: () => editMeeting(null, {date: TODAY()}), o: () => editPerson(null), p: () => editProject(null), '[': () => A['side-toggle']()}[k];
  if (m) { e.preventDefault(); m(); }
});

/* ================= OSOBY AUTOMATYCZNIE (maile, Teams) + TEAMS (Microsoft Graph) ================= */
const FREE_DOMAINS = new Set(['gmail.com', 'googlemail.com', 'outlook.com', 'outlook.pl', 'hotmail.com', 'live.com', 'msn.com', 'yahoo.com', 'icloud.com', 'me.com', 'wp.pl', 'o2.pl', 'onet.pl', 'op.pl', 'onet.eu', 'interia.pl', 'interia.eu', 'poczta.fm', 'vp.pl', 'gazeta.pl', 'tlen.pl', 'protonmail.com', 'proton.me', 'gmx.com', 'gmx.net', 'aol.com', 'yandex.com']);
const AUTO_LOCAL = /^(no-?reply|do-?not-?reply|noreply|mailer-daemon|postmaster|bounces?|notifications?|notify|newsletter|news|marketing|alerts?|automat|system|calendar|invitations?|powiadomienia|faktury|billing|support-noreply|reply)$|(^|[.\-_])(noreply|no-reply|donotreply|bounce)([.\-_]|$)/i;
const AUTO_DOMAIN = /(mailchimp|sendgrid|amazonses|mandrillapp|hubspot|salesforce|zendesk|atlassian\.net|github\.com|linkedin\.com|facebookmail|microsoft\.com|sharepointonline|office365|accounts\.google|docusign|calendly|bounce)/i;
const capW = w => w ? w.charAt(0).toUpperCase() + w.slice(1).toLowerCase() : '';
function nameFromEmail(em) {
  const loc = String(em).split('@')[0].replace(/\+.*$/, '').replace(/[0-9]+$/, ''), parts = loc.split(/[._\-]+/).filter(x => x.length > 1);
  return parts.length >= 2 ? {first: capW(parts[0]), last: parts.slice(1).map(capW).join('-')} : {first: capW(parts[0] || loc), last: ''};
}
function nameSplit(n) {
  n = String(n || '').replace(/\s*[([].*?[)\]]\s*/g, ' ').replace(/["']/g, '').replace(/\s+/g, ' ').trim();
  if (!n || /@/.test(n)) return null;
  if (n.includes(',')) { const [l, f] = n.split(',').map(x => x.trim()); return {first: f || l, last: f ? l : ''}; }
  const w = n.split(' '); return w.length === 1 ? {first: w[0], last: ''} : {first: w.slice(0, -1).join(' '), last: w[w.length - 1]};
}
const myDomain = () => (String(S.user?.email || '').split('@')[1] || '').toLowerCase();
function companyForDomain(dom) {
  if (!dom || FREE_DOMAINS.has(dom)) return '';
  const same = peopleAll().find(p => p.company && String(p.email || '').toLowerCase().endsWith('@' + dom)); if (same) return same.company;
  const L = dom.split('.'), stem = (L.length > 2 && ['com', 'co', 'org', 'net', 'gov', 'edu'].includes(L[L.length - 2]) ? L[L.length - 3] : L[L.length - 2]) || L[0], k = norm(stem).replace(/[^a-z0-9]/g, '');
  const pj = k.length >= 3 && (S.data.projects || []).find(p => p.client && norm(p.client).replace(/[^a-z0-9]/g, '').includes(k)); if (pj) return pj.client;
  const pc = k.length >= 3 && peopleAll().find(p => p.company && norm(p.company).replace(/[^a-z0-9]/g, '').includes(k)); if (pc) return pc.company;
  return stem.length <= 4 ? stem.toUpperCase() : capW(stem);
}
/* jedna osoba = jeden e-mail; bez e-maila – dopasowanie po imieniu i nazwisku; alias tej samej osoby w tej samej domenie nie tworzy duplikatu */
function personUpsert({name, email, source}) {
  email = String(email || '').trim().toLowerCase(); if (email && !/^[^@\s]+@[^@\s]+\.[a-z]{2,}$/.test(email)) email = '';
  let ns = nameSplit(name); const dom = email.split('@')[1] || '';
  if (ns && ns.first && ns.last && email) { const loc = email.split('@')[0].split(/[._\-]+/).map(norm); if (loc.length >= 2 && loc[0] === norm(ns.last) && loc[loc.length - 1] === norm(ns.first)) ns = {first: ns.last, last: ns.first}; else if (!String(name).includes(',') && loc[0] && loc[0] === norm(ns.last).split(' ')[0] && loc[0] !== norm(ns.first)) ns = {first: ns.last, last: ns.first}; }
  const full = ns ? norm(`${ns.first} ${ns.last}`.trim()) : '';
  let p = email ? personByEmail(email) || peopleAll().find(x => (x.aliases || []).includes(email)) : null;
  if (!p && full && full.includes(' ')) p = peopleAll().find(x => norm(pName(x)) === full && (!x.email || !email || x.email.split('@')[1] === dom));
  if (p) { if (email && !p.email) p.email = email; else if (email && p.email !== email && !(p.aliases || []).includes(email)) (p.aliases ||= []).push(email); if (ns && ns.last && !p.last && norm(p.first) === norm(ns.first)) p.last = ns.last; return {p, created: false}; }
  if (!email && !full.includes(' ')) return {p: null};
  if (email && (email === String(S.user.email || '').toLowerCase() || AUTO_LOCAL.test(email.split('@')[0]) || AUTO_DOMAIN.test(dom))) return {p: null};
  const nm = ns && ns.first && ns.last ? ns : email ? nameFromEmail(email) : ns; if (!nm || !nm.first) return {p: null};
  const own = dom && dom === myDomain();
  p = {id: uid(), first: nm.first, last: nm.last || '', email, phone: '', company: companyForDomain(dom), role: '', kind: own ? 'int' : 'ext', ctx: FREE_DOMAINS.has(dom) ? 'priv' : 'job',
    notes: `Utworzono automatycznie (${source}) ${fmt(TODAY())} – sprawdź imię, nazwisko i firmę.`, auto: source, createdAt: TODAY()};
  peopleAll().push(p);
  notifAdd({type: 'person', key: 'person:' + (email || p.id), title: `Utworzono nową osobę: ${pName(p)}`, body: [p.company, email, 'źródło: ' + source].filter(Boolean).join(' · '), ref: {kind: 'person', id: p.id}});
  return {p, created: true};
}
const pplCfg = () => { const st = SET(); return {mail: st.pplMail !== false, days: +st.pplDays || 30, minRecv: st.pplMinRecv == null ? 2 : +st.pplMinRecv, teams: st.pplTeams !== false}; };
let pplBusy = false;
async function mailPeopleSync({silent = true, force = false} = {}) {
  if (!S.user || !S.data) return;
  const c = pplCfg(), st = SET(); if (pplBusy || !S.user || !hasLocalSrv() || (!c.mail && !force)) return;
  if (!force && st.pplLast && Date.now() - st.pplLast < 60 * 60e3) return;
  pplBusy = true;
  try {
    const r = await fetch('/__mail-people?days=' + c.days, {headers: {'X-TM': '1'}, cache: 'no-store'});
    if (r.status === 404) throw new Error('Uruchom ponownie start-TM.bat (nowa wersja serwera)');
    if (!r.ok) throw new Error((await r.text()).slice(0, 200));
    const j = await r.json(), me = String(j.me || '').toLowerCase(), list = Array.isArray(j.people) ? j.people : j.people ? [j.people] : [];
    let n = 0;
    list.filter(x => x && x.email && x.email !== me && ((+x.sent || 0) >= 1 || (+x.recv || 0) >= c.minRecv)).sort((a, b) => String(b.last).localeCompare(String(a.last)))
      .forEach(x => { if (personUpsert({name: x.name, email: x.email, source: +x.sent ? 'mail wysłany' : 'mail odebrany'}).created) n++; });
    st.pplLast = Date.now(); st.pplErr = ''; st.pplAdded = n; save();
    if (n) { notifFlush(); render(); toastAction(`Utworzono ${plN(n, 'nową osobę', 'nowe osoby', 'nowych osób')} z maili`, 'Pokaż', () => { S.view = 'people'; S.pf = Object.assign(S.pf || {}, {q: '', kind: '', sort: 'name', auto: true}); render(); }); }
    else if (!silent) toast(`Sprawdzono ${list.length} adresów – nie ma nowych osób`);
  } catch (e) { st.pplErr = e.message; save(); if (!silent) toast('Osoby z maili: ' + e.message); }
  pplBusy = false;
}

/* ---------- Teams przez Microsoft Graph (czaty: osoby + zadania) ---------- */
const teamsCfg = () => { const st = SET(); return {on: !!st.teamsOn, clientId: String(CFG.MS_CLIENT_ID || st.msClientId || '').trim(), tenant: String(st.msTenant || CFG.MS_TENANT || '').trim() || 'organizations', days: +st.teamsDays || 7,
  kw: st.teamsKw == null ? '#tm' : String(st.teamsKw).trim(), self: st.teamsSelf !== false, mention: st.teamsMention === true, react: st.teamsReact || 'none', every: +st.teamsEvery || 15}; };
const teamsState = () => ((S.data.integrations ||= {}).teams ||= {});
const TEAMS = {
  app: null, acct: null, cid: '', scopes: ['User.Read', 'Chat.Read'],
  async init() {
    const c = teamsCfg(); if (!c.clientId) throw new Error('Brak identyfikatora aplikacji (Client ID) – uzupełnij w Ustawienia → Integracje → Teams');
    if (this.app && this.cid === c.clientId + c.tenant) return;
    await loadScript('https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js');
    this.app = new msal.PublicClientApplication({auth: {clientId: c.clientId, authority: 'https://login.microsoftonline.com/' + c.tenant, redirectUri: location.origin + '/'}, cache: {cacheLocation: 'localStorage'}});
    if (this.app.initialize) await this.app.initialize();
    this.cid = c.clientId + c.tenant; this.acct = this.app.getAllAccounts()[0] || null;
  },
  async login() { await this.init(); const r = await this.app.loginPopup({scopes: this.scopes, loginHint: S.user.email, prompt: 'select_account'}); this.acct = r.account; },
  async token(interactive) {
    await this.init();
    if (!this.acct) { if (!interactive) throw new Error('Wymagane logowanie do Microsoft 365 – kliknij „Połącz z Teams”'); await this.login(); }
    try { return (await this.app.acquireTokenSilent({scopes: this.scopes, account: this.acct})).accessToken; }
    catch { if (!interactive) throw new Error('Sesja Microsoft 365 wygasła – kliknij „Połącz z Teams”'); return (await this.app.acquireTokenPopup({scopes: this.scopes})).accessToken; }
  },
  async g(path, interactive) {
    const r = await fetch(path.startsWith('http') ? path : 'https://graph.microsoft.com/v1.0' + path, {headers: {Authorization: 'Bearer ' + await this.token(interactive)}});
    if (!r.ok) { const t = await r.text(); throw new Error(r.status === 403 ? 'Brak zgody na uprawnienie Chat.Read – poproś administratora o „Udziel zgody” w rejestracji aplikacji' : r.status === 401 ? 'Logowanie wygasło – kliknij „Połącz z Teams”' : `Graph ${r.status}: ${t.slice(0, 160)}`); }
    return r.json();
  }
};
const htmlText = h => { const t = document.createElement('textarea'); t.innerHTML = String(h || '').replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div|li|h\d)>/gi, '\n').replace(/<[^>]+>/g, ''); return t.value.replace(/ /g, ' ').replace(/\n{3,}/g, '\n\n').trim(); };
let teamsBusy = false;
async function teamsSync({silent = true, interactive = false} = {}) {
  const c = teamsCfg(), st = teamsState(); if (teamsBusy || !S.user || (!c.on && silent)) return;
  teamsBusy = true; st.running = true;
  try {
    const me = await TEAMS.g('/me?$select=id,displayName,mail,userPrincipalName', interactive), since = Date.now() - c.days * 864e5;
    let url = '/me/chats?$expand=members,lastMessagePreview&$top=50', chats = [];
    for (let i = 0; url && i < 3; i++) { const j = await TEAMS.g(url); chats.push(...(j.value || [])); url = j['@odata.nextLink'] || ''; }
    chats = chats.map(ch => ({ch, t: Date.parse(ch.lastMessagePreview?.createdDateTime || ch.lastUpdatedDateTime || 0) || 0})).filter(x => x.t >= since).sort((a, b) => b.t - a.t).slice(0, 40).map(x => x.ch);
    let np = 0, nt = 0; const created = [], kw = c.kw ? norm(c.kw) : '';
    for (const ch of chats) {
      const mem = (ch.members || []).map(m => ({id: m.userId, name: m.displayName, email: String(m.email || '').toLowerCase()})), byId = new Map(mem.map(m => [m.id, m]));
      if (c.people !== false && pplCfg().teams) mem.filter(m => m.id !== me.id && m.email).forEach(m => { if (personUpsert({name: m.name, email: m.email, source: 'czat Teams'}).created) np++; });
      const others = mem.filter(m => m.id !== me.id), isSelf = ch.chatType === 'oneOnOne' && !others.length;
      const chatName = ch.topic || others.map(m => m.name).filter(Boolean).join(', ') || (isSelf ? 'Czat ze sobą' : 'Czat Teams');
      if (!kw && !c.mention && c.react === 'none' && !(c.self && isSelf)) continue;
      const msgs = (await TEAMS.g(`/me/chats/${encodeURIComponent(ch.id)}/messages?$top=50`)).value || [];
      for (const m of msgs) {
        if (m.messageType !== 'message' || m.deletedDateTime || Date.parse(m.createdDateTime) < since) continue;
        if (S.data.tasks.some(t => t.teamsId === m.id)) continue;
        const text = htmlText(m.body?.content || ''), fromId = m.from?.user?.id, fromMe = fromId === me.id, plain = htmlText(String(m.body?.content || '').replace(/<at\b[^>]*>[\s\S]*?<\/at>\s*[,:]?/gi, ''));
        if (!text) continue;
        const trig = (c.self && isSelf && fromMe) || (kw && norm(text).includes(kw)) || (c.mention && !fromMe && (m.mentions || []).some(x => x.mentioned?.user?.id === me.id))
          || (c.react !== 'none' && (m.reactions || []).some(r => r.user?.user?.id === me.id && (c.react === 'any' || r.reactionType === c.react)));
        if (!trig) continue;
        const who = byId.get(fromId) || {name: m.from?.user?.displayName || '', email: ''}, rq = !fromMe ? personUpsert({name: who.name, email: who.email, source: 'czat Teams'}) : {p: null};
        if (rq.created) np++;
        const clean = c.kw ? (plain || text).replace(new RegExp(c.kw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'ig'), '').trim() : (plain || text), first = clean.split('\n').map(x => x.trim()).find(Boolean) || '(wiadomość Teams)';
        const due = nextWorkday(), when = String(m.createdDateTime || '').slice(0, 16).replace('T', ' ');
        const t = {id: uid(), title: first.length > 140 ? first.slice(0, 137) + '…' : first, ctx: 'job', goalId: null, projectId: null, start: due, due, status: 'todo', prio: 2, est: 0,
          notes: `Teams – ${chatName}\nOd: ${fromMe ? 'mnie' : who.name || '?'}${who.email && !fromMe ? ' <' + who.email + '>' : ''}${when ? '\nWysłano: ' + when : ''}\n\n${text}`, createdAt: TODAY(), doneAt: null,
          teamsId: m.id, teamsChat: ch.id, teamsUrl: m.webUrl || ch.webUrl || '', source: 'teams'};
        if (rq.p) t.requesterId = rq.p.id;
        const ic = icFind(text.slice(0, 400)); if (ic) t.ic = ic;
        S.data.tasks.push(t); created.push(t); nt++;
        notifAdd({type: 'new', key: 'new:' + t.id, title: 'Nowe zadanie z Teams: ' + t.title, body: `${chatName}${who.name && !fromMe ? ' · od ' + who.name : ''} · termin ${fmt(due)}`, ref: {kind: 'task', id: t.id}});
      }
    }
    Object.assign(st, {last: Date.now(), err: '', me: me.displayName || me.mail || me.userPrincipalName, chats: chats.length, addedTasks: nt, addedPeople: np, running: false}); save();
    if (nt || np) { notifFlush(); render(); toastAction(`Teams: ${[nt ? plN(nt, 'nowe zadanie', 'nowe zadania', 'nowych zadań') : '', np ? plN(np, 'nowa osoba', 'nowe osoby', 'nowych osób') : ''].filter(Boolean).join(' · ')}`, 'Pokaż', () => { S.view = nt ? 'tasks' : 'people'; render(); }); }
    else if (!silent) { render(); toast(`Teams: sprawdzono ${plN(chats.length, 'czat', 'czaty', 'czatów')} – nic nowego`); }
  } catch (e) { Object.assign(st, {err: e.message, running: false, lastTry: Date.now()}); save(); if (!silent) { render(); toast('Teams: ' + e.message); } }
  teamsBusy = false;
}
setInterval(() => { if (!S.user || !S.data) return; const c = teamsCfg(), st = teamsState(); if (!c.on || !c.clientId || !TEAMS.acct || tabInactive) return; if (Date.now() - Math.max(st.last || 0, st.lastTry || 0) >= c.every * 60e3) teamsSync({silent: true}); }, 60e3);
async function teamsBoot(n = 0) { if (!S.user || !S.data) { if (n < 120) setTimeout(() => teamsBoot(n + 1), 5000); return; } const c = teamsCfg(); if (!c.on || !c.clientId) return; try { await TEAMS.init(); if (TEAMS.acct) setTimeout(() => teamsSync({silent: true}), 4000); } catch { } }

/* ---------- ustawienia: Teams + osoby automatycznie ---------- */
function pplCfgCard() {
  const c = pplCfg(), st = SET();
  return `<form data-form="settings" class="card"><h3>⟨i:users⟩ Osoby tworzone automatycznie</h3>
    <p class="muted" style="margin-top:0;font-size:13px">Nowa osoba powstaje tylko wtedy, gdy jej adresu jeszcze nie ma (bez duplikatów; inny adres tej samej osoby w tej samej domenie zapisuje się jako alias). Imię i nazwisko – z nazwy nadawcy, a gdy jej brak – z adresu (np. jan.kowalski@firma.pl → Jan Kowalski). Firma – z domeny: najpierw z innych osób z tej domeny, potem z klienta projektu, inaczej z nazwy domeny. Pomijane są adresy automatyczne (noreply, newsletter, powiadomienia).</p>
    <label class="row" style="gap:8px;font-size:14px;margin-bottom:6px"><input type="checkbox" name="pplMail" ${c.mail ? 'checked' : ''}> <b>Twórz osoby z maili</b> (adresaci maili wysłanych i nadawcy odebranych)</label>
    <label class="row" style="gap:8px;font-size:14px;margin-bottom:8px"><input type="checkbox" name="pplTeams" ${c.teams ? 'checked' : ''}> <b>Twórz osoby z czatów Teams</b></label>
    <div class="grid g2"><label class="f">Sprawdzaj maile z ostatnich (dni)<input class="input" type="number" min="1" max="365" name="pplDays" value="${c.days}"></label>
      <label class="f">Nadawca bez mojej odpowiedzi – minimum odebranych maili<input class="input" type="number" min="1" max="20" name="pplMinRecv" value="${c.minRecv}"></label></div>
    <p class="hint" style="margin:6px 0 0">Sprawdzane raz na godzinę razem z mailami TM. Ostatnio: ${st.pplLast ? ago(st.pplLast) + (st.pplAdded ? ` · utworzono ${st.pplAdded}` : '') : '—'}${st.pplErr ? ` · <span class="late">${esc(st.pplErr)}</span>` : ''}</p>
    <div class="row" style="justify-content:flex-end;margin-top:8px"><button type="button" class="btn" data-action="ppl-now">⟨i:refresh⟩ Utwórz osoby z maili teraz</button><button class="btn primary">Zapisz</button></div></form>`;
}
function vTeamsCfg() {
  const c = teamsCfg(), st = teamsState(), conn = !!TEAMS.acct;
  const reacts = [['none', 'nie używaj reakcji'], ['like', '👍 Lubię to'], ['heart', '❤️ Serce'], ['surprised', '😮 Zaskoczenie'], ['laugh', '😆 Śmiech'], ['any', 'dowolna moja reakcja']];
  return `<div class="grid settings-grid">${tlCard()}
    <form data-form="settings" class="card"><h3>⟨i:video⟩ Czaty Teams (Microsoft Graph – wymaga rejestracji w Azure)</h3>
      <p class="muted" style="margin-top:0;font-size:13px">Czaty Teams są dostępne tylko przez Microsoft Graph (Teams nie udostępnia ich innym programom na komputerze). TM czyta czaty <b>tylko do odczytu</b> (uprawnienie Chat.Read), bezpośrednio z przeglądarki – dane nie trafiają nigdzie poza Twój komputer i Microsoft 365.</p>
      <div class="row" style="gap:10px;margin:4px 0 12px;font-size:13px"><span class="${conn ? 'ok' : 'muted'}">●</span> ${conn ? `Połączono${st.me ? ' jako <b>' + esc(st.me) + '</b>' : ''}` : c.clientId ? 'Nie połączono – kliknij „Połącz z Teams”' : 'Brak identyfikatora aplikacji – patrz instrukcja poniżej'}${st.last ? ` · ostatnia synchronizacja ${ago(st.last)} (${plN(st.chats || 0, 'czat', 'czaty', 'czatów')}${st.addedTasks ? ', zadania: ' + st.addedTasks : ''}${st.addedPeople ? ', osoby: ' + st.addedPeople : ''})` : ''}${st.err ? `<br><span class="late">${esc(st.err)}</span>` : ''}</div>
      <label class="row" style="gap:8px;font-size:14px;margin-bottom:8px"><input type="checkbox" name="teamsOn" ${c.on ? 'checked' : ''}> <b>Włącz integrację z Teams</b> (osoby i zadania z czatów)</label>
      <div class="grid g2"><label class="f">Identyfikator aplikacji (Client ID)<input class="input" name="msClientId" value="${esc(SET().msClientId || CFG.MS_CLIENT_ID || '')}" placeholder="00000000-0000-0000-0000-000000000000" spellcheck="false"></label>
        <label class="f">Identyfikator katalogu (Tenant ID)<input class="input" name="msTenant" value="${esc(SET().msTenant || '')}" placeholder="organizations albo ID dzierżawy" spellcheck="false"></label></div>
      <details style="margin-top:10px"><summary class="muted" style="cursor:pointer;font-size:13px">Jak uzyskać identyfikator aplikacji (jednorazowo, ok. 5 minut)</summary>
        <ol class="muted" style="font-size:13px;line-height:1.6;padding-left:18px;margin:8px 0 0">
          <li>Otwórz <b>portal.azure.com</b> → <b>Microsoft Entra ID</b> → <b>Rejestracje aplikacji</b> → <b>Nowa rejestracja</b>.</li>
          <li>Nazwa: <code>TM Workspace</code>; typ kont: <i>tylko konta w tym katalogu organizacyjnym</i>.</li>
          <li>Identyfikator URI przekierowania: platforma <b>Aplikacja jednostronicowa (SPA)</b>, adres <code>http://localhost:8765/</code> → Zarejestruj.</li>
          <li><b>Uprawnienia interfejsu API</b> → Dodaj → Microsoft Graph → <b>Uprawnienia delegowane</b>: <code>User.Read</code>, <code>Chat.Read</code>. Jeśli pojawi się „wymagana zgoda administratora”, poproś IT o kliknięcie „Udziel zgody administratora”.</li>
          <li>Na stronie <b>Przegląd</b> skopiuj <b>Identyfikator aplikacji (klienta)</b> i <b>Identyfikator katalogu (dzierżawy)</b> do pól powyżej, zapisz i kliknij „Połącz z Teams”.</li>
        </ol>
        <p class="hint" style="margin:6px 0 0">Jeśli w Twojej firmie użytkownicy nie mogą rejestrować aplikacji, przekaż tę instrukcję działowi IT.</p></details>
      <div class="row" style="justify-content:flex-end;margin-top:12px;gap:6px"><button type="button" class="btn" data-action="teams-connect" ${c.clientId ? '' : 'disabled'}>⟨i:plug⟩ ${conn ? 'Połącz ponownie' : 'Połącz z Teams'}</button><button type="button" class="btn" data-action="teams-sync" ${c.clientId ? '' : 'disabled'}>⟨i:refresh⟩ Synchronizuj teraz</button><button class="btn primary">Zapisz</button></div></form>
    <form data-form="settings" class="card"><h3>⟨i:check-square⟩ Zadania z wiadomości Teams</h3>
      <p class="muted" style="margin-top:0;font-size:13px">Wiadomość z czatu staje się zadaniem JOB (termin: następny dzień roboczy), gdy spełnia jeden z warunków poniżej. Nadawca jest przypisany jako zlecający, a w zadaniu jest link „Otwórz w Teams”. Ta sama wiadomość nie utworzy drugiego zadania.</p>
      <div class="grid g2"><label class="f">Słowo kluczowe w treści (puste = wyłączone)<input class="input" name="teamsKw" value="${esc(c.kw)}" placeholder="#tm"></label>
        <label class="f">Moja reakcja na wiadomość<select class="input" name="teamsReact">${reacts.map(([k, l]) => `<option value="${k}" ${c.react === k ? 'selected' : ''}>${l}</option>`).join('')}</select></label></div>
      <label class="row" style="gap:8px;font-size:14px;margin:10px 0 4px"><input type="checkbox" name="teamsSelf" ${c.self ? 'checked' : ''}> Każda moja wiadomość w czacie ze sobą (notatki)</label>
      <label class="row" style="gap:8px;font-size:14px;margin-bottom:8px"><input type="checkbox" name="teamsMention" ${c.mention ? 'checked' : ''}> Każda wiadomość, w której ktoś mnie @oznaczył</label>
      <div class="grid g2"><label class="f">Sprawdzaj czaty z ostatnich (dni)<input class="input" type="number" min="1" max="30" name="teamsDays" value="${c.days}"></label>
        <label class="f">Synchronizuj co (minut)<input class="input" type="number" min="5" max="240" name="teamsEvery" value="${c.every}"></label></div>
      <div class="row" style="justify-content:flex-end;margin-top:8px"><button class="btn primary">Zapisz</button></div></form>
    ${pplCfgCard()}</div>`;
}
Object.assign(A, {
  'ppl-auto': () => { (S.pf ||= {q: '', kind: '', sort: 'name'}).auto = !S.pf.auto; render(); },
  'ppl-now': () => mailPeopleSync({silent: false, force: true}),
  'teams-connect': async () => { try { TEAMS.app = null; await TEAMS.login(); toast('Połączono z Microsoft Teams'); SET().teamsOn = true; save(); await teamsSync({silent: false, interactive: true}); } catch (e) { teamsState().err = e.message; save(); render(); toast('Teams: ' + (e.errorMessage || e.message)); } },
  'teams-sync': () => teamsSync({silent: false, interactive: true}),
  'teams-open': el => { const t = S.data.tasks.find(x => x.id === el.dataset.id); if (t?.teamsUrl) teamsLaunch(teamsApp(t.teamsUrl), 'Otwieram wiadomość w Teams'); }
});

/* ================= HARMONOGRAM ZADAŃ =================
   Plan        = terminy i estymacje (capacity nieograniczone)  → vPlan / Capacity
   Harmonogram = kiedy chcę się zajmować zadaniem (bloki czasu w ograniczonym capacity) → ten moduł
   Realizacja  = rzeczywiste wpisy czasu (stoper lub wpis post factum) → Realizacja / Ewidencja
   Blok harmonogramu NIE jest realizacją – czas powstaje dopiero po starcie stopera albo wpisie ręcznym. */
const schAll = () => (S.data.schedule ||= []);
const SCH_H = 14;                                                    // horyzont automatu: dziś + 13 dni
const hmToMin = s => { const m = /^(\d{1,2}):(\d{2})$/.exec(String(s || '').trim()); return m ? +m[1] * 60 + +m[2] : null; };
const minToHm = m => `${pad(Math.floor(m / 60))}:${pad(m % 60)}`;
const schDur = b => Math.max(0, hmToMin(b.e) - hmToMin(b.s));
const parseWins = (str, def) => { const r = String(str || def).split(/[,;]/).map(x => x.split('-').map(hmToMin)).filter(([a, b]) => a != null && b != null && b > a); return r.length ? r : parseWins(def, def); };
const schCfg = () => { const st = SET(); return {job: parseWins(st.schJob, '08:30-16:00'), priv: parseWins(st.schPriv, '06:00-08:30,16:00-22:30'), min: +st.schMin || 30, jobStr: st.schJob || '08:30-16:00', privStr: st.schPriv || '06:00-08:30,16:00-22:30'}; };
const nowMin = () => { const d = new Date(); return d.getHours() * 60 + d.getMinutes(); };
/* okna pracy dnia dla kontekstu (z urlopami, świętami, delegacjami) */
function schWindows(ds, ctx) {
  const c = schCfg(), d = P(ds), ty = dayType(d), tr = tripAt(ds), scale = (wins, pct) => { if (pct >= 100) return wins; let left = wins.reduce((a, [x, y]) => a + y - x, 0) * pct / 100; return wins.map(([x, y]) => { const take = Math.max(0, Math.min(y - x, left)); left -= take; return [x, x + take]; }).filter(([x, y]) => y - x >= 15); };
  let w;
  if (ctx === 'job') { if (ty !== 'work') return []; w = c.job.map(x => [...x]); if (halfVac(d)) { const tot = w.reduce((a, [x, y]) => a + y - x, 0); w = scale(w, 50); void tot; } if (tr) w = scale(w, tripCap(tr, 'job')); }
  else { w = c.priv.map(x => [...x]); if (ty !== 'work') w = mergeIv([...w, ...c.job.map(x => [...x])]); if (tr) w = scale(w, tripCap(tr, 'priv')); }
  return w;
}
function mergeIv(iv) { const a = iv.filter(([x, y]) => y > x).sort((p, q) => p[0] - q[0]), out = []; a.forEach(([x, y]) => { const l = out[out.length - 1]; if (l && x <= l[1]) l[1] = Math.max(l[1], y); else out.push([x, y]); }); return out; }
function subIv(wins, busy) { let res = wins.map(x => [...x]); mergeIv(busy).forEach(([bx, by]) => { res = res.flatMap(([x, y]) => by <= x || bx >= y ? [[x, y]] : [[x, Math.min(bx, y)], [Math.max(by, x), y]]).filter(([x, y]) => y - x > 0); }); return res; }
const mtgBusy = ds => S.data.meetings.filter(m => m.start.slice(0, 10) === ds).map(m => [toMin(m.start), Math.max(toMin(m.start) + 15, m.end.slice(0, 10) === ds ? toMin(m.end) : 24 * 60)]);
const schOfDay = ds => schAll().filter(b => b.d === ds).sort((a, b) => a.s.localeCompare(b.s));
const schTask = id => S.data.tasks.find(t => t.id === id);
const schVisible = b => { const t = schTask(b.taskId); return t && (t.status !== 'done' || b.d <= TODAY()); };

/* AUTOMAT: usuwa przyszłe bloki (opcjonalnie zostawia zatwierdzone) i układa zadania od teraz na 14 dni */
function schAuto({keepLocked = false} = {}) {
  const td = TODAY(), nm = Math.ceil(nowMin() / 15) * 15, c = schCfg(), days = [...Array(SCH_H)].map((_, i) => iso(add(today(), i))), last = days[days.length - 1];
  const keep = b => b.d < td || (b.d === td && hmToMin(b.s) < nm) || (keepLocked && b.lock) || b.recur || b.d > last;
  S.data.schedule = schAll().filter(keep);
  const placed = new Map(), busyOf = {}; days.forEach(d => { busyOf[d] = [...mtgBusy(d), ...schOfDay(d).map(b => [hmToMin(b.s), hmToMin(b.e)])]; });
  schAll().filter(b => b.d >= td).forEach(b => placed.set(b.taskId, (placed.get(b.taskId) || 0) + schDur(b)));
  const tasks = S.data.tasks.filter(t => t.status !== 'done').map(t => ({t, need: Math.ceil(Math.max(0, taskRemain(t) * 60 - (placed.get(t.id) || 0)) / 15) * 15}))
    .filter(x => x.need >= 15).sort((a, b) => (a.t.due || '9999').localeCompare(b.t.due || '9999') || a.t.prio - b.t.prio || (a.t.status === 'doing' ? -1 : 0) - (b.t.status === 'doing' ? -1 : 0));
  const res = {placed: 0, minutes: 0, late: [], short: []};
  for (const x of tasks) {
    const t = x.t, from = t.start && t.start > td ? t.start : td; let need = x.need, lateMin = 0;
    for (const d of days) {
      if (need <= 0) break; if (d < from) continue;
      let free = subIv(schWindows(d, t.ctx), busyOf[d]); if (d === td) free = subIv(free, [[0, nm]]);
      for (const [a, b] of free) {
        if (need <= 0) break; const len = b - a; if (len < Math.min(c.min, need)) continue;
        const take = Math.min(need, len), blk = {id: uid(), taskId: t.id, d, s: minToHm(a), e: minToHm(a + take), auto: true, lock: false};
        schAll().push(blk); busyOf[d].push([a, a + take]); need -= take; res.minutes += take; if (t.due && d > t.due) lateMin += take;
      }
    }
    if (x.need > need) res.placed++;
    if (lateMin) res.late.push({t, min: lateMin}); if (need > 0) res.short.push({t, min: need});
  }
  SET().schLast = Date.now(); save(); return res;
}
/* KONFLIKTY: nakładanie się, bloki po terminie, poza oknami kontekstu, zadania bez miejsca w harmonogramie */
function schConflicts() {
  const td = TODAY(), last = iso(add(today(), SCH_H - 1)), out = [], vis = schAll().filter(b => b.d >= td && schVisible(b));
  const byDay = {}; vis.forEach(b => (byDay[b.d] ||= []).push(b));
  Object.entries(byDay).forEach(([d, l]) => {
    const ms = S.data.meetings.filter(m => m.start.slice(0, 10) === d);
    l.forEach(b => { const s = hmToMin(b.s), e = hmToMin(b.e), t = schTask(b.taskId);
      const mm = ms.find(m => toMin(m.start) < e && toMin(m.end) > s); if (mm) out.push({k: 'mtg', b, t, txt: `nakłada się na spotkanie „${mm.title}” (${hm(mm.start)})`});
      const ob = l.find(o => o !== b && o.id < b.id && hmToMin(o.s) < e && hmToMin(o.e) > s); if (ob) out.push({k: 'blk', b, t, txt: `nakłada się na blok „${schTask(ob.taskId)?.title || '?'}” (${ob.s})`});
      if (t?.due && d > t.due && t.due >= td) out.push({k: 'due', b, t, txt: `zaplanowane po terminie zadania (${fmt(t.due)})`});
      const w = schWindows(d, t?.ctx || 'job'); if (!w.some(([x, y]) => s >= x && e <= y)) out.push({k: 'win', b, t, txt: t?.ctx === 'priv' ? 'poza oknem prywatnym' : 'poza godzinami pracy', soft: true}); });
  });
  const sched = new Map(); vis.forEach(b => sched.set(b.taskId, (sched.get(b.taskId) || 0) + schDur(b)));
  S.data.tasks.filter(t => t.status !== 'done' && t.due && t.due < td && sched.has(t.id)).forEach(t => out.push({k: 'due', t, soft: true, txt: `termin minął ${fmt(t.due)} – ustaw nowy termin (zaplanowano ${fmtH(sched.get(t.id))})`}));
  S.data.tasks.filter(t => t.status !== 'done' && t.due && t.due >= td && t.due <= last).forEach(t => { const need = taskRemain(t) * 60 - (sched.get(t.id) || 0); if (need >= 15) out.push({k: 'short', t, txt: `brakuje ${fmtH(need)} w harmonogramie przed terminem ${fmt(t.due)}`}); });
  return out;
}
const schToday = () => schOfDay(TODAY()).filter(schVisible);
function schNow() { const n = nowMin(), l = schToday(); return {cur: l.find(b => hmToMin(b.s) <= n && hmToMin(b.e) > n), next: l.find(b => hmToMin(b.s) > n)}; }

/* ---------- widok: Dostępność › Harmonogram ---------- */
const SCH_H0 = 6 * 60, SCH_H1 = 22 * 60 + 30, SCH_PPM = .8;
function vSched() {
  const G = S.sch ||= {off: 0}, ws = add(today(), G.off * 7), days = [...Array(7)].map((_, i) => iso(add(ws, i))), td = TODAY(), c = schCfg(), P_ = S.privacy ? pv : x => x;
  const cf = schConflicts(), hard = cf.filter(x => !x.soft), y = m => (m - SCH_H0) * SCH_PPM, H = (SCH_H1 - SCH_H0) * SCH_PPM;
  const ent = entriesWithLive().filter(e => days.includes(e.start.slice(0, 10)));
  const hours = []; for (let m = SCH_H0; m <= SCH_H1; m += 60) hours.push(m);
  const col = d => {
    const bands = [...schWindows(d, 'job').map(w => ['job', w]), ...schWindows(d, 'priv').filter(([a, b]) => !schWindows(d, 'job').some(([x, z]) => a >= x && b <= z)).map(w => ['priv', w])];
    const ms = S.data.meetings.filter(m => m.start.slice(0, 10) === d), bl = schOfDay(d).filter(schVisible), es = ent.filter(e => e.start.slice(0, 10) === d), cfIds = new Set(hard.filter(x => x.b).map(x => x.b.id));
    const avail = subIv(mergeIv([...schWindows(d, 'job'), ...schWindows(d, 'priv')]), mtgBusy(d)).reduce((a, [p, q]) => a + q - p, 0), used = bl.reduce((a, b) => a + schDur(b), 0);
    return `<div class="sch-col ${d === td ? 'today' : ''} ${dayType(P(d)) !== 'work' ? 'off' : ''}" data-d="${d}" style="height:${H}px">
      ${bands.map(([k, [a, b]]) => `<i class="sch-band ${k}" style="top:${y(Math.max(a, SCH_H0))}px;height:${Math.max(0, y(Math.min(b, SCH_H1)) - y(Math.max(a, SCH_H0)))}px"></i>`).join('')}
      ${ms.map(m => { const a = Math.max(toMin(m.start), SCH_H0), b = Math.min(m.end.slice(0, 10) === d ? toMin(m.end) : SCH_H1, SCH_H1); return b > a ? `<div class="sch-mtg ${m.ctx} ${mtgLogged(m) ? 'rcp' : ''}" style="top:${y(a)}px;height:${Math.max(14, y(b) - y(a))}px" data-action="edit-mtg" data-id="${m.id}" title="Spotkanie ${hm(m.start)}–${hm(m.end)}: ${esc(m.title)}${mtgLogged(m) ? ' · czas w RCP' : ''}"><b>${hm(m.start)}</b> ${m.recurId ? '⟨i:repeat⟩' : ''}${P_(esc(m.title))}${m.ctx !== 'priv' ? `<button class="sch-mt" data-action="mtg-trk" data-id="${m.id}" title="${mtgLogged(m) ? 'Czas spotkania jest w RCP' : 'Rejestruj czas spotkania (RCP)'}">⟨i:${mtgLogged(m) ? 'check' : 'timer'}⟩</button>` : ''}</div>` : ''; }).join('')}
      ${bl.map(b => { const t = schTask(b.taskId), p = projById(projOf(t)), a = hmToMin(b.s), e = hmToMin(b.e); return `<div class="sch-blk ${t.ctx} ${b.lock ? 'lock' : ''} ${cfIds.has(b.id) ? 'conf' : ''} ${t.status === 'done' ? 'done' : ''}" data-bid="${b.id}" style="top:${y(a)}px;height:${Math.max(16, y(e) - y(a))}px" title="${esc(t.title)} · ${b.s}–${b.e}${b.lock ? ' · zatwierdzony' : ''}">
        <div class="sb-1">${b.lock ? '⟨i:lock⟩' : ''}<span class="num">${b.s}</span>${p ? mono(p, 'xs') : ''}</div><div class="sb-2">${P_(esc(t.title))}</div><i class="sch-rs" title="Zmień długość"></i></div>`; }).join('')}
      ${es.map(e => { const a = Math.max(toMin(e.start), SCH_H0), b = Math.min(toMin(e.end), SCH_H1); return b > a ? `<i class="sch-real ${e.live ? 'live' : ''}" style="top:${y(a)}px;height:${Math.max(3, y(b) - y(a))}px" title="Realizacja ${hm(e.start)}–${hm(e.end)}: ${esc(e.title || '')}"></i>` : ''; }).join('')}
      ${d === td && nowMin() >= SCH_H0 && nowMin() <= SCH_H1 ? `<i class="sch-now" style="top:${y(nowMin())}px"></i>` : ''}</div>`;
  };
  const head = d => { const bl = schOfDay(d).filter(schVisible), used = bl.reduce((a, b) => a + schDur(b), 0) + S.data.meetings.filter(m => m.start.slice(0, 10) === d).reduce((a, m) => a + mtgMinOn(m, d), 0), avail = subIv(mergeIv([...schWindows(d, 'job'), ...schWindows(d, 'priv')]), mtgBusy(d)).reduce((a, [p, q]) => a + q - p, 0), n = hard.filter(x => x.b?.d === d).length;
    return `<div class="sch-dh ${d === td ? 'today' : ''}"><b>${['Nd', 'Pn', 'Wt', 'Śr', 'Cz', 'Pt', 'Sb'][P(d).getDay()]} ${P(d).getDate()}.${pad(P(d).getMonth() + 1)}</b><span class="num">${fmtHM(used)} / ${fmtHM(avail + used)}</span>${n ? `<span class="late" title="Konflikty">${n}!</span>` : ''}${holName(P(d)) ? `<small>${esc(holName(P(d)))}</small>` : vacOf(P(d)) ? '<small>urlop</small>' : tripAt(d) ? `<small>⟨i:plane⟩ ${esc(tripAt(d).place)}</small>` : ''}</div>`; };
  const sched = new Map(); schAll().filter(b => b.d >= td).forEach(b => sched.set(b.taskId, (sched.get(b.taskId) || 0) + schDur(b)));
  const unsch = S.data.tasks.filter(t => t.status !== 'done' && taskRemain(t) * 60 - (sched.get(t.id) || 0) >= 15).sort((a, b) => (a.due || '9').localeCompare(b.due || '9') || a.prio - b.prio);
  const last = SET().schLast;
  return `<div class="sch-top">
      <div class="sch-actions"><button class="pbtn" data-action="sch-auto">⟨i:refresh⟩ Ułóż automatycznie – 2 tygodnie</button><button class="sbtn" data-action="sch-lock" title="Zatwierdź (przypnij) bloki bez konfliktów w widocznym tygodniu">⟨i:lock⟩ Zatwierdź tydzień</button><button class="sbtn ghost" data-action="sch-clear">Wyczyść przyszłe</button>
        <span class="spacer"></span><button class="ib" data-action="sch-week" data-d="-1" title="Poprzedni tydzień">⟨i:chev-left⟩</button><button class="sbtn ghost sm" data-action="sch-week" data-d="0">Ten tydzień</button><button class="ib" data-action="sch-week" data-d="1" title="Następny tydzień">⟨i:chev-right⟩</button><span class="sch-range num">${fmt(days[0])} – ${fmt(days[6])}</span></div>
      <div class="sch-legend"><span><i class="lg job"></i>praca ${esc(c.jobStr)}</span><span><i class="lg priv"></i>prywatne ${esc(c.privStr)}</span><span><i class="lg mtg"></i>spotkania (zajmują czas jako pierwsze)</span><span><i class="lg real"></i>realizacja (wpisy czasu)</span><span>⟨i:lock⟩ zatwierdzony</span><span class="spacer"></span><span class="dim">${last ? 'ostatni automat ' + ago(last) : 'automat jeszcze nie był uruchamiany'} · przeciągnij blok, aby przesunąć; dolną krawędź, aby zmienić długość; kliknij puste pole, aby dodać</span></div></div>
    <div class="sch-wrap"><div class="sch-grid" style="--h:${H}px">
      <div class="sch-hd0"></div>${days.map(head).join('')}
      <div class="sch-axis" style="height:${H}px">${hours.map(m => `<span style="top:${y(m)}px" class="num">${minToHm(m)}</span>`).join('')}</div>${days.map(col).join('')}</div></div>
    <div class="sch-side">
      <section class="panel"><div class="pn-h"><h3>Konflikty do rozwiązania</h3><span class="pn-sub">${hard.length}${cf.length > hard.length ? ` + ${cf.length - hard.length} ostrzeżeń` : ''}</span></div>
        ${cf.length ? cf.slice(0, 30).map(x => `<div class="sch-cf ${x.soft ? 'soft' : ''}" ${x.b ? `data-action="sch-edit" data-id="${x.b.id}"` : `data-action="edit-task" data-id="${x.t.id}"`}><span class="cf-k">${{mtg: 'Spotkanie', blk: 'Nakładanie', due: 'Po terminie', win: 'Okno', short: 'Brak czasu'}[x.k]}</span><div class="cf-t"><b>${P_(esc(x.t?.title || '?'))}</b><span>${x.b ? `${fmt(x.b.d)} ${x.b.s}–${x.b.e} · ` : ''}${esc(x.txt)}</span></div></div>`).join('') : '<div class="empty-st sm">Brak konfliktów – harmonogram mieści się w capacity i przed terminami.</div>'}</section>
      ${recPanel()}
      <section class="panel"><div class="pn-h"><h3>Bez miejsca w harmonogramie</h3><span class="pn-sub">${unsch.length}</span></div>
        ${unsch.slice(0, 20).map(t => `<div class="sch-un"><span class="dot ${t.ctx}"></span><div class="cf-t" data-action="edit-task" data-id="${t.id}"><b>${P_(esc(t.title))}</b><span>${fmtH(taskRemain(t) * 60 - (sched.get(t.id) || 0))} do ułożenia · ${t.due ? 'termin ' + fmt(t.due) : 'bez terminu'} · P${t.prio}</span></div><button class="sbtn sm" data-action="sch-add" data-task="${t.id}">⟨i:plus⟩ Blok</button></div>`).join('') || '<div class="empty-st sm">Wszystkie otwarte zadania mają czas w harmonogramie.</div>'}</section>
      <form data-form="settings" class="panel"><div class="pn-h"><h3>Okna czasu</h3></div>
        <div class="grid g2"><label class="f">Praca (JOB)<input class="input" name="schJob" value="${esc(c.jobStr)}" placeholder="08:30-16:00"></label><label class="f">Prywatne (PRIV)<input class="input" name="schPriv" value="${esc(c.privStr)}" placeholder="06:00-08:30,16:00-22:30"></label>
        <label class="f">Najkrótszy blok (min)<input class="input" type="number" min="15" max="120" step="15" name="schMin" value="${c.min}"></label></div>
        <p class="hint" style="margin:6px 0 0">W dni wolne okno pracy jest dostępne dla zadań prywatnych. Delegacje i pół dnia urlopu skracają okna. Zadania bez estymacji liczone po ${defEst()} h.</p>
        <div class="row" style="justify-content:flex-end;margin-top:8px"><button class="btn primary">Zapisz</button></div></form></div>`;
}
afterRender.sched = () => schDragInit();
function schDragInit() {
  const grid = $('.sch-grid'); if (!grid) return;
  grid.addEventListener('mousedown', e => {
    const el = e.target.closest('.sch-blk'); if (!el || e.button !== 0) return; e.preventDefault();
    const b = schAll().find(x => x.id === el.dataset.bid); if (!b) return;
    const resize = e.target.classList.contains('sch-rs'), y0 = e.clientY, s0 = hmToMin(b.s), e0 = hmToMin(b.e); let moved = false, nd = b.d, ns = s0, ne = e0;
    el.classList.add('drag');
    const mv = ev => { const dm = Math.round((ev.clientY - y0) / SCH_PPM / 15) * 15; if (Math.abs(ev.clientY - y0) > 3) moved = true;
      if (resize) { ne = Math.max(s0 + 15, Math.min(SCH_H1, e0 + dm)); el.style.height = Math.max(16, (ne - s0) * SCH_PPM) + 'px'; }
      else { const len = e0 - s0; ns = Math.max(SCH_H0, Math.min(SCH_H1 - len, s0 + dm)); ne = ns + len; el.style.top = (ns - SCH_H0) * SCH_PPM + 'px';
        const c = document.elementsFromPoint(ev.clientX, ev.clientY).find(x => x.classList?.contains('sch-col')); if (c && c.dataset.d !== nd) { nd = c.dataset.d; c.append(el); } } };
    const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up);
      if (!moved) return schEdit(b);
      Object.assign(b, {d: nd, s: minToHm(ns), e: minToHm(ne), auto: false}); save(); render(); };
    document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
  });
  grid.addEventListener('click', e => { const c = e.target.closest('.sch-col'); if (!c || e.target.closest('.sch-blk,.sch-mtg')) return; const r = c.getBoundingClientRect(), m = Math.floor(((e.clientY - r.top) / SCH_PPM + SCH_H0) / 15) * 15; schAdd(null, c.dataset.d, m); });
}
function schEdit(b) {
  const t = schTask(b.taskId); if (!t) return;
  modal(`<h3>Blok harmonogramu</h3><form data-form="schblk" data-id="${b.id}" class="grid g2">
    <div class="sch-mt" style="grid-column:1/-1">${mono(projById(projOf(t)), 'sm')}<a href="#" data-action="edit-task" data-id="${t.id}">${esc(t.title)}</a><span class="dim">${t.due ? 'termin ' + fmt(t.due) : 'bez terminu'} · P${t.prio} · pozostało ${fmtH(taskRemain(t) * 60)}</span></div>
    <label class="f">Dzień<input class="input" type="date" name="d" value="${b.d}" required></label><span></span>
    <label class="f">Od<input class="input" type="time" step="900" name="s" value="${b.s}" required></label><label class="f">Do<input class="input" type="time" step="900" name="e" value="${b.e}" required></label>
    <label class="row" style="grid-column:1/-1;gap:8px;font-size:13px"><input type="checkbox" name="lock" ${b.lock ? 'checked' : ''}> Zatwierdzony (przypięty – zostaje, gdy automat ma zachować zatwierdzone bloki)</label>
    <p class="hint" style="grid-column:1/-1;margin:0">Blok to plan, nie realizacja. Czas trafi do RCP dopiero po „Rozpocznij realizację” albo wpisie ręcznym.</p>
    <div class="row" style="grid-column:1/-1;justify-content:space-between"><div class="row"><button type="button" class="btn danger" data-action="sch-del" data-id="${b.id}">Usuń blok</button><button type="button" class="btn" data-action="sch-start" data-id="${b.id}">⟨i:timer⟩ Rozpocznij realizację</button></div>
      <div class="row"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">Zapisz</button></div></div></form>`);
}
function schAdd(taskId, d, m) {
  const td = TODAY(), open = S.data.tasks.filter(t => t.status !== 'done').sort((a, b) => (a.due || '9').localeCompare(b.due || '9') || a.prio - b.prio), s = m != null ? Math.max(SCH_H0, Math.min(SCH_H1 - 30, m)) : Math.max(SCH_H0, Math.ceil(nowMin() / 15) * 15);
  modal(`<h3>Nowy blok w harmonogramie</h3><form data-form="schblk" class="grid g2">
    <label class="f" style="grid-column:1/-1">Zadanie<select class="input" name="taskId" required>${open.map(t => `<option value="${t.id}" ${t.id === taskId ? 'selected' : ''}>${esc(t.title)}${t.due ? ' · ' + fmt(t.due) : ''} · ${t.ctx === 'priv' ? 'PRIV' : 'JOB'}</option>`).join('')}</select></label>
    <label class="f">Dzień<input class="input" type="date" name="d" value="${d || td}" required></label><span></span>
    <label class="f">Od<input class="input" type="time" step="900" name="s" value="${minToHm(s)}" required></label><label class="f">Do<input class="input" type="time" step="900" name="e" value="${minToHm(Math.min(SCH_H1, s + 60))}" required></label>
    <label class="row" style="grid-column:1/-1;gap:8px;font-size:13px"><input type="checkbox" name="lock"> Zatwierdzony</label>
    <div class="row" style="grid-column:1/-1;justify-content:flex-end"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">Dodaj blok</button></div></form>`);
}
Object.assign(A, {
  'sch-auto': () => {
    const fut = schAll().filter(b => b.d > TODAY() || (b.d === TODAY() && hmToMin(b.s) >= nowMin())), locked = fut.filter(b => b.lock).length;
    const run = keep => { const r = schAuto({keepLocked: keep}); render(); toastAction(`Harmonogram: ${plN(r.placed, 'zadanie', 'zadania', 'zadań')} · ${fmtH(r.minutes)}${r.late.length ? ` · ${r.late.length} po terminie` : ''}${r.short.length ? ` · ${r.short.length} bez miejsca` : ''}`, 'Konflikty', () => { $('.sch-side')?.scrollIntoView({behavior: 'smooth'}); }, 9000); };
    if (!fut.length) return run(false);
    S._schKeep = false;
    confirmDlg('Ułożyć harmonogram od nowa?', `<p>Automat usunie ${plN(fut.length, 'przyszły blok', 'przyszłe bloki', 'przyszłych bloków')} (także edytowane ręcznie) i ułoży zadania od teraz na 14 dni: najpierw spotkania z kalendarza, potem zadania według terminów i priorytetów w oknach pracy i prywatnych.</p>${locked ? `<label class="row" style="gap:8px;font-size:13px"><input type="checkbox" id="schKeep" onchange="S._schKeep=this.checked"> zachowaj ${plN(locked, 'zatwierdzony blok', 'zatwierdzone bloki', 'zatwierdzonych bloków')}</label>` : ''}`, 'Ułóż od nowa', () => run(!!S._schKeep));
  },
  'sch-clear': () => confirmDlg('Wyczyścić przyszły harmonogram?', '<p>Usunę wszystkie bloki od teraz (także zatwierdzone). Realizacja i wpisy czasu zostają bez zmian.</p>', 'Wyczyść', () => { const n = nowMin(); S.data.schedule = schAll().filter(b => b.d < TODAY() || (b.d === TODAY() && hmToMin(b.s) < n)); save(); render(); toast('Wyczyszczono przyszły harmonogram'); }),
  'sch-lock': () => { const G = S.sch ||= {off: 0}, ws = iso(add(today(), G.off * 7)), we = iso(add(today(), G.off * 7 + 6)), bad = new Set(schConflicts().filter(x => !x.soft && x.b).map(x => x.b.id)); let n = 0;
    schAll().forEach(b => { if (b.d >= ws && b.d <= we && b.d >= TODAY() && !bad.has(b.id) && !b.lock) { b.lock = true; n++; } }); save(); render(); toast(n ? `Zatwierdzono ${plN(n, 'blok', 'bloki', 'bloków')}${bad.size ? ` · ${bad.size} z konfliktem pozostaje do przeplanowania` : ''}` : 'Brak bloków do zatwierdzenia'); },
  'sch-week': el => { const G = S.sch ||= {off: 0}; G.off = +el.dataset.d === 0 ? 0 : G.off + +el.dataset.d; render(); },
  'sch-edit': el => { const b = schAll().find(x => x.id === el.dataset.id); if (b) schEdit(b); },
  'sch-add': el => schAdd(el.dataset.task || null),
  'sch-del': el => { const i = schAll().findIndex(x => x.id === el.dataset.id); if (i < 0) return; const b = schAll()[i]; schAll().splice(i, 1); save(); closeModal(); render(); toastAction('Usunięto blok', '⟨i:undo⟩ Cofnij', () => { schAll().splice(i, 0, b); save(); render(); }); },
  'sch-start': el => { const b = schAll().find(x => x.id === el.dataset.id); if (!b) return; closeModal(); startTimer({taskId: b.taskId}); },
  'sch-open': () => { S.view = 'sched'; history.replaceState(null, '', '#sched'); render(); $('main').scrollTop = 0; }
});
Object.assign(FORMS, {
  schblk: (f, fd) => { const s = hmToMin(fd.s), e = hmToMin(fd.e); if (s == null || e == null || e <= s) return toast('Godzina „do” musi być późniejsza niż „od”');
    const id = f.dataset.id, b = id ? schAll().find(x => x.id === id) : {id: uid(), taskId: fd.taskId, auto: false};
    Object.assign(b, {d: fd.d, s: minToHm(Math.round(s / 15) * 15), e: minToHm(Math.round(e / 15) * 15), lock: fd.lock != null, auto: false}); if (!id) schAll().push(b);
    save(); closeModal(); render(); const cf = schConflicts().filter(x => !x.soft && x.b?.id === b.id); toast(cf.length ? `Zapisano – uwaga: ${cf[0].txt}` : 'Zapisano blok harmonogramu'); }
});

/* ---------- kokpit: harmonogram i realizacja dnia ---------- */
function schDayPanel() {
  const td = TODAY(), bl = schToday(), es = entriesWithLive().filter(e => e.start.slice(0, 10) === td), ms = meetingsC().filter(m => m.start.slice(0, 10) === td), P_ = S.privacy ? pv : x => x;
  const a0 = Math.min(SCH_H0, ...bl.map(b => hmToMin(b.s)), ...es.map(e => toMin(e.start))), a1 = Math.max(SCH_H1, ...bl.map(b => hmToMin(b.e)), ...es.map(e => toMin(e.end))), x = m => ((m - a0) / (a1 - a0) * 100).toFixed(2) + '%', w = (a, b) => ((b - a) / (a1 - a0) * 100).toFixed(2) + '%';
  const planned = bl.reduce((a, b) => a + schDur(b), 0) + ms.reduce((a, m) => a + mtgMinOn(m, td), 0), real = es.reduce((a, e) => a + entMin(e), 0), schIds = new Set(bl.map(b => b.taskId)), realOnPlan = es.filter(e => schIds.has(e.taskId)).reduce((a, e) => a + entMin(e), 0);
  const {cur, next} = schNow(), nb = cur || next, nt = nb && schTask(nb.taskId), n = nowMin(), cfN = schConflicts().filter(c => !c.soft && (c.b ? c.b.d === td : false)).length;
  const ticks = []; for (let m = Math.ceil(a0 / 120) * 120; m <= a1; m += 120) ticks.push(m);
  return `<section class="panel sd"><div class="pn-h"><h3>Dzień: harmonogram i realizacja</h3><span class="spacer"></span><a href="#" class="pn-link" data-action="sch-open">Harmonogram ⟨i:arrow-right⟩</a></div>
    <div class="sd-kpi"><div><span class="dim">Harmonogram</span><b class="num">${fmtHM(planned)}</b></div><div><span class="dim">Realizacja</span><b class="num">${fmtHM(real)}</b></div><div><span class="dim">Realizacja wg harmonogramu</span><b class="num">${planned ? Math.round(100 * Math.min(realOnPlan, planned) / planned) : 0}%</b></div>${cfN ? `<div><span class="dim">Konflikty</span><b class="num late">${cfN}</b></div>` : ''}
      <div class="sd-next">${sdMtgNext(nb) || (nt ? `<span class="dim">${cur ? 'Teraz wg harmonogramu' : 'Następny blok'} · <span class="num">${nb.s}–${nb.e}</span></span><b>${P_(esc(nt.title))}</b>${timerOn()?.taskId === nt.id ? '<span class="sd-on"><span class="rec"></span> realizujesz</span>' : `<button class="sbtn sm" data-action="timer-start-task" data-id="${nt.id}">⟨i:timer⟩ Rozpocznij realizację</button>`}` : bl.length ? '<span class="dim">Dzisiejsze bloki są już za Tobą.</span>' : `<span class="dim">Brak harmonogramu na dziś.</span><button class="sbtn sm" data-action="sch-auto">⟨i:refresh⟩ Ułóż automatycznie</button>`)}</div></div>
    <div class="sd-tl"><div class="sd-ax">${ticks.map(m => `<span style="left:${x(m)}" class="num">${minToHm(m)}</span>`).join('')}</div>
      <div class="sd-lane"><span class="sd-l">Spotkania</span><div class="sd-tr">${ms.map(m => `<i class="sd-m" data-action="edit-mtg" data-id="${m.id}" style="left:${x(toMin(m.start))};width:${w(toMin(m.start), toMin(m.end))}" title="${hm(m.start)}–${hm(m.end)} ${esc(m.title)}"></i>`).join('')}</div></div>
      <div class="sd-lane"><span class="sd-l">Harmonogram</span><div class="sd-tr">${ms.map(m => `<i class="sd-b mtg ${m.ctx}" style="left:${x(toMin(m.start))};width:${w(toMin(m.start), toMin(m.end))}" data-action="edit-mtg" data-id="${m.id}" title="Spotkanie ${hm(m.start)}–${hm(m.end)} ${esc(m.title)}"></i>`).join('')}${bl.map(b => { const t = schTask(b.taskId); return `<i class="sd-b ${t.ctx} ${b === cur ? 'cur' : ''}" style="left:${x(hmToMin(b.s))};width:${w(hmToMin(b.s), hmToMin(b.e))}" data-action="sch-edit" data-id="${b.id}" title="${b.s}–${b.e} ${esc(t.title)}"></i>`; }).join('')}</div></div>
      <div class="sd-lane"><span class="sd-l">Realizacja</span><div class="sd-tr">${es.map(e => `<i class="sd-r ${e.live ? 'live' : ''}" style="left:${x(toMin(e.start))};width:${w(toMin(e.start), Math.max(toMin(e.end), toMin(e.start) + 5))}" title="${hm(e.start)}–${hm(e.end)} ${esc(e.title || '')}"></i>`).join('')}</div></div>
      ${n >= a0 && n <= a1 ? `<i class="sd-now" style="left:calc(84px + (100% - 84px) * ${((n - a0) / (a1 - a0)).toFixed(4)})"></i>` : ''}</div></section>`;
}

/* ================= RCP: SPOTKANIA + TYLKO JOB =================
   Stoper (start/stop) mierzy wyłącznie pracę JOB. Zadania i spotkania PRIV nie są rejestrowane.
   Spotkanie JOB można rejestrować stoperem (od teraz albo od początku trwającego spotkania) lub zapisać post factum. */
const mtgById = id => S.data.meetings.find(x => x.id === id);
const mtgEntries = m => entriesAll().filter(e => e.mtgId === m.id);
const mtgLogged = m => mtgEntries(m).length > 0 || timerOn()?.mtgId === m.id;
const mtgMinOn = (m, d) => m.start.slice(0, 10) !== d ? 0 : Math.max(0, (m.end.slice(0, 10) === d ? toMin(m.end) : 24 * 60) - toMin(m.start));
function mtgTimer(id) {
  const m = mtgById(id); if (!m) return;
  if (m.ctx === 'priv') return toast('Spotkania prywatne nie są rejestrowane w RCP');
  const tm = timerOn(); if (tm?.mtgId === id) return toast('Czas tego spotkania już się rejestruje');
  if (m.start.slice(0, 10) > TODAY()) return toast('Rejestrację czasu spotkania można włączyć w dniu spotkania');
  if (strMs(m.end) <= Date.now()) return mtgLog(id);
  if (tm) stopTimer('switch', true);
  const now = Date.now(), st = strMs(m.start); let startMs = now;
  if (st < now) { const lastEnd = Math.max(0, ...entriesAll().filter(e => e.endMs && e.endMs <= now).map(e => e.endMs)); startMs = Math.max(st, lastEnd); }
  S.data.timer = {taskId: null, mtgId: m.id, title: m.title, projectId: m.projectId || null, ctx: 'job', startMs, note: '', free: true};
  S.trk = {}; save(); closeModal(); render();
  toast(`Rejestracja spotkania: ${S.privacy ? '' : '„' + m.title + '” '}od ${new Date(startMs).toTimeString().slice(0, 5)} – zatrzymaj po zakończeniu`);
}
function mtgLog(id) {
  const m = mtgById(id); if (!m) return;
  if (m.ctx === 'priv') return toast('Spotkania prywatne nie są rejestrowane w RCP');
  const s = strMs(m.start), e = Math.min(strMs(m.end), Date.now());
  if (e - s < 60000) return toast('Spotkanie jeszcze się nie zaczęło');
  if (mtgEntries(m).length) return toast('Czas tego spotkania jest już w RCP');
  const ov = entriesAll().filter(x => x.startMs < e && x.endMs > s);
  const doIt = () => { const en = {id: uid(), taskId: null, mtgId: m.id, title: m.title, projectId: m.projectId || null, ctx: 'job', startMs: s, endMs: e, start: Lms(s), end: Lms(e), note: 'Spotkanie', kind: 'mtg'};
    entriesAll().push(en); save(); closeModal(); render();
    toastAction(`Zapisano w RCP: ${fmtH((e - s) / 60000)} (${hm(en.start)}–${hm(en.end)})`, '⟨i:undo⟩ Cofnij', () => { const i = entriesAll().indexOf(en); if (i >= 0) entriesAll().splice(i, 1); save(); render(); }); };
  if (!ov.length) return doIt();
  confirmDlg('Czas nakłada się na inne wpisy', `<p>W godzinach ${hm(m.start)}–${hm(Lms(e))} masz już ${plN(ov.length, 'wpis', 'wpisy', 'wpisów')} realizacji (${ov.slice(0, 3).map(x => esc(x.title || '')).join(', ')}). Zapisać mimo to? Wpisy możesz potem poprawić w Ewidencji.</p>`, 'Zapisz', doIt);
}
/* przycisk RCP dla spotkania (pusty dla PRIV i spotkań z przyszłych dni) */
function mtgTrkBtn(m, cls = 'sbtn sm', big = false) {
  if (!m || !m.id || m.ctx === 'priv' || m.start.slice(0, 10) > TODAY()) return '';
  const tm = timerOn(), lg = mtgEntries(m), ended = strMs(m.end) <= Date.now();
  if (tm?.mtgId === m.id) return `<button type="button" class="${cls}" data-action="timer-pause" title="Zatrzymaj i zapisz czas spotkania">⟨i:pause⟩ ${big ? 'Zatrzymaj rejestrację spotkania' : 'Zatrzymaj'}</button>`;
  if (lg.length) return `<span class="mtg-rcp" title="Czas spotkania jest w RCP">⟨i:check⟩ RCP ${fmtH(lg.reduce((a, e) => a + entMin(e), 0))}</span>`;
  return `<button type="button" class="${cls}" data-action="mtg-trk" data-id="${m.id}" title="${ended ? 'Zapisz czas spotkania jako wpis realizacji (RCP)' : 'Start stopera na spotkaniu – czas trafi do RCP'}">⟨i:timer⟩ ${ended ? (big ? 'Zapisz czas spotkania do RCP' : 'Do RCP') : (big ? 'Rejestruj czas spotkania' : 'Rejestruj')}</button>`;
}
/* spotkania do wyboru w oknie „Start realizacji” */
function mtgPickHtml() {
  const td = TODAY(), now = Date.now(), ms = S.data.meetings.filter(m => m.ctx !== 'priv' && m.start.slice(0, 10) === td && !mtgLogged(m) && strMs(m.start) - now <= 30 * 60e3).sort((a, b) => a.start.localeCompare(b.start));
  return ms.length ? `<div class="group-h">Spotkania dziś – czas do RCP</div><div class="tplchips" style="margin-bottom:10px">${ms.map(m => `<button class="tplchip" data-action="mtg-trk" data-id="${m.id}">⟨i:video⟩ <span class="num">${hm(m.start)}–${hm(m.end)}</span> ${pv(esc(m.title))}${strMs(m.end) <= now ? ' · zapisz' : ''}</button>`).join('')}</div>` : '';
}
/* kokpit: spotkanie jako pozycja harmonogramu (teraz / następne) */
function sdMtgNext(nb) {
  const n = nowMin(), td = TODAY(), m = meetingsC().filter(x => x.start.slice(0, 10) === td && (x.end.slice(0, 10) > td || toMin(x.end) > n)).sort((a, b) => a.start.localeCompare(b.start))[0];
  if (!m) return ''; const s = toMin(m.start), live = s <= n; if (!live && nb && hmToMin(nb.s) < s) return '';
  return `<span class="dim">${live ? 'Teraz: spotkanie' : 'Następne: spotkanie'} · <span class="num">${hm(m.start)}–${hm(m.end)}</span></span><b data-action="edit-mtg" data-id="${m.id}" style="cursor:pointer">${(S.privacy ? pv : x => x)(esc(m.title))}</b>${timerOn()?.mtgId === m.id ? '<span class="sd-on"><span class="rec"></span> rejestrujesz</span>' : mtgTrkBtn(m, 'sbtn sm')}`;
}
/* ukrywa start stopera przy zadaniach/projektach PRIV (wszystkie widoki i okna) */
function trkGuard(root) {
  if (!S.data || !root) return;
  root.querySelectorAll('[data-action="timer-start-task"],[data-action="sch-start"],[data-action="proj-play"],[data-action="free-start"]').forEach(b => {
    const a = b.dataset.action; let priv;
    if (a === 'timer-start-task') priv = S.data.tasks.find(t => t.id === b.dataset.id)?.ctx === 'priv';
    else if (a === 'sch-start') priv = schTask(schAll().find(x => x.id === b.dataset.id)?.taskId)?.ctx === 'priv';
    else if (a === 'proj-play') priv = projById(b.dataset.id)?.ctx === 'priv';
    else priv = projById(b.dataset.project || S.homeProj)?.ctx === 'priv';
    if (!priv) return;
    if (b.parentElement?.classList.contains('t-act')) { const sp = document.createElement('span'); sp.className = 'no-trk'; sp.title = 'Zadania prywatne nie są rejestrowane w czasie (RCP tylko dla JOB)'; sp.textContent = 'bez RCP'; b.replaceWith(sp); }
    else b.remove();
  });
}
Object.assign(A, {
  'mtg-trk': (el, e) => { e?.stopPropagation(); const m = mtgById(el.dataset.id); if (!m) return; if (strMs(m.end) <= Date.now()) mtgLog(m.id); else mtgTimer(m.id); }
});

/* ================= ZADANIA I SPOTKANIA CYKLICZNE =================
   Seria (S.data.recurring) generuje wystąpienia: zadania na 14 dni (horyzont harmonogramu), spotkania na 62 dni.
   Zadanie z godziną dostaje zatwierdzony blok harmonogramu (automat go nie przesuwa); bez godziny – układa je automat. */
const recAll = () => (S.data.recurring ||= []);
const REC_F = [['daily', 'Codziennie'], ['work', 'W dni robocze'], ['weekly', 'Co tydzień'], ['monthly', 'Co miesiąc']];
const DOW_S = ['Pn', 'Wt', 'Śr', 'Cz', 'Pt', 'Sb', 'Nd'];
const isoDow = ds => (P(ds).getDay() + 6) % 7 + 1;                    // 1 = pon … 7 = nd
const recById = id => recAll().find(r => r.id === id);
function recDesc(r) {
  const ev = +r.every > 1 ? +r.every : 1;
  const f = r.freq === 'daily' ? (ev > 1 ? `co ${ev} dni` : 'codziennie') : r.freq === 'work' ? 'w dni robocze' : r.freq === 'weekly' ? `${ev > 1 ? `co ${ev} tyg.` : 'co tydzień'}: ${(r.dows || []).map(d => DOW_S[d - 1]).join(', ')}` : `${ev > 1 ? `co ${ev} mies.` : 'co miesiąc'}, ${r.dom || 1}. dnia`;
  return `${f}${r.time ? ' · ' + r.time : ''}${r.dur ? ' · ' + fmtH(r.dur) : ''}${r.until ? ' · do ' + fmt(r.until) : ''}`;
}
function recDates(r, from, to) {
  const out = [], a = r.from && r.from > from ? r.from : from, b = r.until && r.until < to ? r.until : to, ev = Math.max(1, +r.every || 1);
  if (!r.from || a > b) return out;
  const w0 = iso(add(P(r.from), -(isoDow(r.from) - 1)));
  for (let d = a; d <= b; d = iso(add(P(d), 1))) {
    let ok = false; const dt = P(d);
    if (r.freq === 'daily') ok = diffD(r.from, d) % ev === 0;
    else if (r.freq === 'work') ok = isWork(dt);
    else if (r.freq === 'weekly') ok = (r.dows || []).includes(isoDow(d)) && Math.floor(diffD(w0, d) / 7) % ev === 0;
    else if (r.freq === 'monthly') { const last = new Date(dt.getFullYear(), dt.getMonth() + 1, 0).getDate(), f = P(r.from), md = (dt.getFullYear() - f.getFullYear()) * 12 + dt.getMonth() - f.getMonth();
      ok = dt.getDate() === Math.min(+r.dom || 1, last) && md % ev === 0; }
    if (ok) out.push(d);
  }
  return out;
}
/* generowanie wystąpień (idempotentne: daty już wygenerowane są zapamiętane w serii) */
function recurGen() {
  if (!S.data || !S.data.tasks) return 0; const td = TODAY(); let n = 0;
  recAll().forEach(r => {
    const gs = new Set(r.gen || []);
    if (r.active !== false) recDates(r, td, iso(add(today(), (r.kind === 'mtg' ? 62 : SCH_H) - 1))).forEach(d => {
      if (gs.has(d)) return; gs.add(d); n++;
      const dur = Math.max(15, +r.dur || (r.kind === 'mtg' ? 30 : 60));
      if (r.kind === 'mtg') { const st = `${d}T${r.time || '09:00'}`;
        S.data.meetings.push({id: uid(), title: r.title, start: st, end: addMin(st, dur), ctx: r.ctx, projectId: r.projectId || null, source: 'local', attendees: [...(r.attendees || [])], location: r.location || '', body: r.notes || '', recurId: r.id, recurDate: d}); }
      else { const t = {id: uid(), title: r.title, ctx: r.ctx, goalId: null, projectId: r.projectId || null, start: d, due: d, status: 'todo', prio: +r.prio || 2, est: dur / 60, notes: r.notes || '', createdAt: td, doneAt: null, recurId: r.id, recurDate: d};
        S.data.tasks.push(t);
        if (r.time) { const s = hmToMin(r.time); if (s != null) schAll().push({id: uid(), taskId: t.id, d, s: minToHm(s), e: minToHm(Math.min(s + dur, 23 * 60 + 45)), auto: false, lock: true, recur: true}); } }
    });
    // pominięte wystąpienia z przeszłości (bez czasu realizacji) nie zalegają jako zadania po terminie
    if (r.kind !== 'mtg' && r.skipMissed !== false) S.data.tasks = S.data.tasks.filter(t => { const drop = t.recurId === r.id && t.status !== 'done' && t.due && t.due < td && !taskSpent(t.id) && !cmtAll().some(c => c.refId === t.id); if (drop) n++; return !drop; });
    r.gen = [...gs].filter(d => d >= iso(add(today(), -45))).sort();
  });
  if (n) { const ids = new Set(S.data.tasks.map(t => t.id)); S.data.schedule = schAll().filter(b => ids.has(b.taskId)); save(); }
  return n;
}
/* usuwa przyszłe (nierealizowane) wystąpienia serii, np. po zmianie lub usunięciu serii */
function recPurge(r) {
  const td = TODAY(), nowS = `${td}T${new Date().toTimeString().slice(0, 5)}`; let n = 0;
  S.data.tasks = S.data.tasks.filter(t => { const d = t.recurId === r.id && t.status !== 'done' && (t.due || td) >= td && !taskSpent(t.id) && !cmtAll().some(c => c.refId === t.id); if (d) n++; return !d; });
  S.data.meetings = S.data.meetings.filter(m => { const d = m.recurId === r.id && m.start > nowS && !mtgEntries(m).length && !notesOf('mtg', m.id).length; if (d) n++; return !d; });
  const ids = new Set(S.data.tasks.map(t => t.id)); S.data.schedule = schAll().filter(b => ids.has(b.taskId));
  const keep = new Set([...S.data.tasks.filter(t => t.recurId === r.id).map(t => t.recurDate), ...S.data.meetings.filter(m => m.recurId === r.id).map(m => m.recurDate)]);
  r.gen = (r.gen || []).filter(d => d < td || keep.has(d)); return n;
}
function editSeries(r, preset = {}) {
  const isNew = !r, td = TODAY();
  r = r || {kind: 'task', title: '', ctx: S.ctx === 'priv' ? 'priv' : 'job', projectId: null, prio: 2, freq: 'weekly', every: 1, dows: [isoDow(td)], dom: P(td).getDate(), time: '', dur: 60, from: td, until: '', notes: '', attendees: [], location: '', active: true, skipMissed: true, ...preset};
  const fr = REC_F.map(([k, l]) => `<option value="${k}" ${r.freq === k ? 'selected' : ''}>${l}</option>`).join('');
  modal(`<h3>⟨i:repeat⟩ ${isNew ? 'Nowa seria cykliczna' : 'Seria cykliczna'}</h3><form data-form="recur" data-id="${isNew ? '' : r.id}" class="grid g2 rec-f">
    <label class="f">Rodzaj<select class="input" name="kind"><option value="task" ${r.kind !== 'mtg' ? 'selected' : ''}>Zadanie</option><option value="mtg" ${r.kind === 'mtg' ? 'selected' : ''}>Spotkanie (tylko w TM)</option></select></label>
    <label class="f">Kontekst<select class="input" name="ctx"><option value="job" ${r.ctx !== 'priv' ? 'selected' : ''}>JOB – firmowe</option><option value="priv" ${r.ctx === 'priv' ? 'selected' : ''}>PRIV – prywatne</option></select></label>
    <label class="f" style="grid-column:1/-1">Nazwa<input class="input" name="title" required value="${esc(r.title)}" placeholder="np. Raport tygodniowy, Daily, Trening"></label>
    <label class="f">Projekt<select class="input" name="projectId">${projOptions(r.projectId)}</select></label>
    <label class="f" data-rk="task">Priorytet<select class="input" name="prio">${[1, 2, 3].map(n => `<option ${+r.prio === n ? 'selected' : ''}>${n}</option>`).join('')}</select></label>
    <label class="f">Powtarzaj<select class="input" name="freq">${fr}</select></label>
    <label class="f">Co ile (dni/tyg./mies.)<input class="input" type="number" min="1" max="12" name="every" value="${+r.every || 1}"></label>
    <div class="f" data-rf="weekly" style="grid-column:1/-1"><span>Dni tygodnia</span><div class="row rec-dows">${DOW_S.map((l, i) => `<label class="rec-dow"><input type="checkbox" name="dow${i + 1}" ${(r.dows || []).includes(i + 1) ? 'checked' : ''}><span>${l}</span></label>`).join('')}</div></div>
    <label class="f" data-rf="monthly">Dzień miesiąca<input class="input" type="number" min="1" max="31" name="dom" value="${+r.dom || 1}"></label>
    <label class="f">Godzina <span class="dim" data-rk="task">(opcjonalnie – stały blok w harmonogramie)</span><input class="input" type="time" step="900" name="time" value="${r.time || ''}"></label>
    <label class="f">Czas trwania (min)<input class="input" type="number" min="15" step="15" name="dur" value="${+r.dur || 60}"></label>
    <label class="f">Od<input class="input" type="date" name="from" required value="${r.from || td}"></label>
    <label class="f">Do (puste = bez końca)<input class="input" type="date" name="until" value="${r.until || ''}"></label>
    <label class="f" data-rk="mtg" style="grid-column:1/-1">Uczestnicy (e-maile, przecinek)<input class="input" name="att" value="${esc((r.attendees || []).join(', '))}"></label>
    <label class="f" data-rk="mtg" style="grid-column:1/-1">Lokalizacja<input class="input" name="loc" value="${esc(r.location || '')}"></label>
    <label class="f" style="grid-column:1/-1">Opis / agenda<textarea class="input" name="notes" rows="2">${esc(r.notes || '')}</textarea></label>
    <label class="row" style="grid-column:1/-1;gap:8px;font-size:13px"><input type="checkbox" name="active" ${r.active !== false ? 'checked' : ''}> Seria aktywna</label>
    <label class="row" data-rk="task" style="grid-column:1/-1;gap:8px;font-size:13px"><input type="checkbox" name="skipMissed" ${r.skipMissed !== false ? 'checked' : ''}> Usuwaj niewykonane wystąpienia z przeszłości (bez czasu realizacji), żeby nie zalegały po terminie</label>
    <p class="hint" style="grid-column:1/-1;margin:0">Zadania tworzą się na ${SCH_H} dni do przodu (horyzont harmonogramu), spotkania na 62 dni. Zadanie z godziną dostaje zatwierdzony blok w harmonogramie; bez godziny – układa je automat w oknie ${r.ctx === 'priv' ? 'prywatnym' : 'pracy'}.${isNew ? '' : ` Zmiana serii odtwarza przyszłe, jeszcze nierealizowane wystąpienia.`}</p>
    <div class="row" style="grid-column:1/-1;justify-content:space-between"><div class="row">${isNew ? '' : `<button type="button" class="btn danger" data-action="rec-del" data-id="${r.id}">⟨i:trash⟩ Usuń serię</button>`}</div>
      <div class="row"><button type="button" class="btn" data-action="modal-close">Anuluj</button><button class="btn primary">${isNew ? 'Utwórz serię' : 'Zapisz serię'}</button></div></div></form>`, mm => {
    const upd = () => { const k = $('[name=kind]', mm).value, f = $('[name=freq]', mm).value;
      $$('[data-rk]', mm).forEach(x => x.hidden = x.dataset.rk !== k); $$('[data-rf]', mm).forEach(x => x.hidden = x.dataset.rf !== f);
      $('[name=every]', mm).closest('label').hidden = f === 'work'; };
    $('[name=kind]', mm).addEventListener('change', upd); $('[name=freq]', mm).addEventListener('change', upd); upd(); $('[name=title]', mm).focus();
  });
}
/* sekcja w formularzu zadania / spotkania */
function recurLine(o, kind, isNew) {
  const r = o && o.recurId && recById(o.recurId);
  if (r) return `<div class="trk-line rec-line" style="grid-column:1/-1">⟨i:repeat⟩ ${kind === 'mtg' ? 'Spotkanie' : 'Zadanie'} cykliczne: <b>${esc(recDesc(r))}</b>${r.active === false ? ' · <span class="late">seria wstrzymana</span>' : ''}<span class="spacer"></span><button type="button" class="btn sm" data-action="rec-edit" data-id="${r.id}">Edytuj serię</button></div>`;
  return `<div class="trk-line rec-line" style="grid-column:1/-1">⟨i:repeat⟩ <span class="dim">Jednorazowe</span><span class="spacer"></span><button type="button" class="btn sm" data-action="rec-from" data-kind="${kind}" data-id="${isNew ? '' : o.id}">Powtarzaj…</button></div>`;
}
function recPanel() {
  const L = recAll(), P_ = S.privacy ? pv : x => x;
  return `<section class="panel"><div class="pn-h"><h3>Cykliczne</h3><span class="pn-sub">${L.filter(r => r.active !== false).length}</span><span class="spacer"></span><button class="sbtn sm" data-action="rec-new">⟨i:plus⟩ Seria</button></div>
    ${L.length ? L.map(r => `<div class="sch-un rec-row ${r.active === false ? 'off' : ''}" data-action="rec-edit" data-id="${r.id}"><span class="dot ${r.ctx}"></span><div class="cf-t"><b>${r.kind === 'mtg' ? '⟨i:video⟩ ' : ''}${P_(esc(r.title))}</b><span>${esc(recDesc(r))}${r.active === false ? ' · wstrzymana' : ''}</span></div></div>`).join('') : '<div class="empty-st">Brak serii. Zadania i spotkania powtarzalne (np. raport tygodniowy, daily, trening) pojawią się w harmonogramie automatycznie.</div>'}</section>`;
}
Object.assign(A, {
  'rec-new': () => editSeries(null),
  'rec-edit': (el, e) => { e?.stopPropagation(); const r = recById(el.dataset.id); if (r) editSeries(r); },
  'rec-from': el => {
    const f = el.closest('form'), v = n => f?.elements[n]?.value || '', kind = el.dataset.kind, id = el.dataset.id;
    let pre;
    if (kind === 'mtg') { const st = `${v('date') || TODAY()}T${v('from') || '09:00'}`, en = `${v('date') || TODAY()}T${v('to') || '10:00'}`;
      pre = {kind: 'mtg', title: v('title'), ctx: v('ctx') || 'job', projectId: v('projectId') || null, time: v('from') || '09:00', dur: Math.max(15, toMin(en) - toMin(st)), from: st.slice(0, 10), dows: [isoDow(st.slice(0, 10))], dom: P(st.slice(0, 10)).getDate(),
        attendees: v('att').split(/[,;\s]+/).filter(x => x.includes('@')), location: v('loc'), notes: v('body')}; }
    else { const d = v('due') || v('start') || TODAY();
      pre = {kind: 'task', title: v('title'), ctx: v('ctx') || 'job', projectId: v('projectId') || null, prio: +v('prio') || 2, dur: Math.round((+v('est') || 1) * 60 / 15) * 15 || 60, from: d, dows: [isoDow(d)], dom: P(d).getDate(), notes: v('notes')}; }
    S._recSrc = id ? {kind, id} : null; editSeries(null, pre);
  },
  'rec-del': el => { const r = recById(el.dataset.id); if (!r) return;
    confirmDlg('Usunąć serię cykliczną?', `<p>Usunę serię „${esc(r.title)}” i jej przyszłe wystąpienia, które nie mają jeszcze realizacji ani notatek. Wystąpienia z przeszłości i zrealizowane zostają.</p>`, 'Usuń serię', () => {
      const n = recPurge(r); S.data.tasks.forEach(t => { if (t.recurId === r.id) t.recurId = null; }); S.data.meetings.forEach(m => { if (m.recurId === r.id) m.recurId = null; });
      S.data.recurring = recAll().filter(x => x !== r); save(); render(); toast(`Usunięto serię${n ? ` i ${plN(n, 'przyszłe wystąpienie', 'przyszłe wystąpienia', 'przyszłych wystąpień')}` : ''}`); }); }
});
Object.assign(FORMS, {
  recur: (f, fd) => {
    const id = f.dataset.id, old = id && recById(id), dows = DOW_S.map((_, i) => i + 1).filter(i => fd['dow' + i] != null);
    if (fd.freq === 'weekly' && !dows.length) return toast('Wybierz co najmniej jeden dzień tygodnia');
    if (fd.until && fd.until < fd.from) return toast('Data końca serii jest wcześniejsza niż początek');
    if (fd.kind === 'mtg' && !fd.time) return toast('Podaj godzinę spotkania');
    const r = old || {id: uid(), gen: [], createdAt: TODAY()};
    if (old) recPurge(old);
    Object.assign(r, {kind: fd.kind, title: fd.title.trim(), ctx: fd.ctx, projectId: fd.projectId || null, prio: +fd.prio || 2, freq: fd.freq, every: Math.max(1, +fd.every || 1), dows, dom: +fd.dom || 1, time: fd.time || '', dur: Math.max(15, +fd.dur || 60),
      from: fd.from, until: fd.until || '', notes: fd.notes || '', attendees: String(fd.att || '').split(/[,;\s]+/).filter(x => x.includes('@')), location: fd.loc || '', active: fd.active != null, skipMissed: fd.skipMissed != null});
    if (!old) recAll().push(r);
    // obiekt, z którego utworzono serię, staje się jej pierwszym wystąpieniem (bez duplikatu w tym dniu)
    const src = S._recSrc; S._recSrc = null;
    if (!old && src) { const o = src.kind === 'mtg' ? mtgById(src.id) : S.data.tasks.find(t => t.id === src.id);
      if (o) { const d = src.kind === 'mtg' ? o.start.slice(0, 10) : (o.due || o.start); if (recDates(r, d, d).length) { o.recurId = r.id; o.recurDate = d; r.gen.push(d); } } }
    const n = recurGen(); closeModal(); render();
    toast(`${old ? 'Zapisano serię' : 'Utworzono serię'} „${r.title}” · ${recDesc(r)}${n ? ` · ${plN(n, 'wystąpienie', 'wystąpienia', 'wystąpień')} w horyzoncie` : ''}`);
  }
});

/* ================= TEAMS – PROSTA INTEGRACJA (lokalnie, jak Outlook) =================
   Bez Azure: lokalny serwer czyta przez klasycznego Outlooka (COM)
   • spotkania Teams z kalendarza (z linkiem „Dołącz”) + uczestników → Osoby
   • zadania z Microsoft To Do / Teams (folder Zadania w Outlooku i jego listy) → Zadania JOB
   Zakończenie zadania w TM oznacza je jako wykonane w To Do. Czaty Teams wymagają nadal połączenia Graph (niżej). */
const tlCfg = () => { const st = SET(); return {on: st.tlOn !== false, mtg: st.tlMtg !== false, todo: st.tlTodo !== false, ppl: st.tlPpl !== false, done: st.tlDone !== false, every: +st.tlEvery || 30}; };
let tlBusy = false;
async function teamsLocalSync({silent = true, force = false} = {}) {
  if (!S.user || !S.data || tlBusy || !hasLocalSrv()) return;
  const c = tlCfg(), st = SET();
  if (!force && (!c.on || (st.tlLast && Date.now() - st.tlLast < c.every * 60e3))) return;
  tlBusy = true; const arr = x => Array.isArray(x) ? x : x ? [x] : [];
  try {
    const doneIds = c.done ? S.data.tasks.filter(t => t.todoId && t.status === 'done' && !t.todoClosed).map(t => t.todoId) : [];
    if (doneIds.length) { const r = await fetch('/__todo-done', {method: 'POST', headers: {'X-TM': '1', 'Content-Type': 'application/json'}, body: JSON.stringify({ids: doneIds})}); if (r.ok) S.data.tasks.forEach(t => { if (doneIds.includes(t.todoId)) t.todoClosed = true; }); }
    const r = await fetch(`/__teams-local?back=7&fwd=${IMP_FWD_DEF}&mtg=${c.mtg ? 1 : 0}&todo=${c.todo ? 1 : 0}`, {headers: {'X-TM': '1'}, cache: 'no-store'});
    if (r.status === 404) throw new Error('Uruchom ponownie start-TM.bat (nowa wersja serwera)');
    if (!r.ok) throw new Error((await r.text()).slice(0, 200));
    const j = await r.json(); let nm = 0, nt = 0, np = 0;
    if (c.mtg) {
      const from = add(today(), -7), to = add(today(), IMP_FWD_DEF + 1), key = m => norm(m.title) + '|' + m.start, ics = new Set(S.data.meetings.filter(m => m.imported === 'outlook').map(key));
      const list = arr(j.meetings).filter(x => x && x.start).map(x => ({id: 'tl-' + hash((x.gid || x.title || '') + '|' + x.start).toString(36), title: x.title || '(bez tytułu)', start: x.start, end: x.end || addMin(x.start, 30), source: 'teams', ctx: 'job', projectId: null,
        attendees: arr(x.attendees).map(a => String(a.email || '').toLowerCase()).filter(Boolean), location: x.location || '', body: String(x.body || '').slice(0, 2000), link: x.link || '', imported: 'teamsl'}));
      const fresh = list.filter(m => !ics.has(key(m))), before = new Set(S.data.meetings.filter(m => m.imported === 'teamsl').map(m => m.id));
      applyImport(fresh, 'teamsl', from, to); nm = fresh.filter(m => !before.has(m.id)).length;
      if (c.ppl) arr(j.meetings).forEach(x => arr(x.attendees).forEach(a => { if (personUpsert({name: a.name, email: a.email, source: 'Teams – spotkanie'}).created) np++; }));
    }
    if (c.todo) { const seen = S.data.todoSeen ||= [];
      arr(j.tasks).forEach(x => { if (!x || !x.id || !x.title) return; const t0 = S.data.tasks.find(y => y.todoId === x.id);
        if (t0) { if (x.due && t0.todoDue !== x.due && t0.status !== 'done') { t0.due = x.due; t0.todoDue = x.due; } return; }
        if (seen.includes(x.id)) return;                                    // usunięte w TM – nie wraca
        const t = {id: uid(), title: String(x.title).trim(), ctx: 'job', goalId: null, projectId: null, start: TODAY(), due: x.due || nextWorkday(), status: 'todo', prio: +x.imp === 2 ? 1 : +x.imp === 0 ? 3 : 2, est: 0,
          notes: [x.list ? 'Lista To Do: ' + x.list : '', String(x.body || '').trim()].filter(Boolean).join('\n\n'), createdAt: TODAY(), doneAt: null, todoId: x.id, todoDue: x.due || '', source: 'todo'};
        const ic = icFind(t.title); if (ic) t.ic = ic;
        S.data.tasks.push(t); seen.push(x.id); nt++;
        notifAdd({type: 'new', key: 'new:' + t.id, title: 'Nowe zadanie z Teams / To Do: ' + t.title, body: (x.list ? 'lista ' + x.list + ' · ' : '') + 'termin ' + fmt(t.due), ref: {kind: 'task', id: t.id}}); });
      if (seen.length > 3000) S.data.todoSeen = seen.slice(-3000); }
    st.tlLast = Date.now(); st.tlErr = ''; st.tlStat = {m: arr(j.meetings).length, t: arr(j.tasks).length, nm, nt, np}; save();
    if (nm || nt || np) { notifFlush(); render(); toast(`Teams: ${[nm ? plN(nm, 'nowe spotkanie', 'nowe spotkania', 'nowych spotkań') : '', nt ? plN(nt, 'nowe zadanie', 'nowe zadania', 'nowych zadań') : '', np ? plN(np, 'nowa osoba', 'nowe osoby', 'nowych osób') : ''].filter(Boolean).join(' · ')}`); }
    else { if (S.view === 'settings') render(); if (!silent) toast(`Teams: sprawdzono ${plN(arr(j.meetings).length, 'spotkanie', 'spotkania', 'spotkań')} i ${plN(arr(j.tasks).length, 'zadanie', 'zadania', 'zadań')} – nic nowego`); }
  } catch (e) { st.tlErr = e.message; save(); if (!silent) toast('Teams (Outlook): ' + e.message); }
  tlBusy = false;
}
function tlCard() {
  const c = tlCfg(), st = SET(), s = st.tlStat, srv = hasLocalSrv();
  return `<form data-form="settings" class="card tl-card"><h3>⟨i:plug⟩ Teams – prosta integracja (przez Outlook)</h3>
    <p class="muted" style="margin-top:0;font-size:13px">Działa jak integracja z Outlookiem: lokalny serwer TM czyta dane z klasycznego Outlooka na tym komputerze. Bez rejestracji aplikacji w Azure i bez logowania. Pobiera <b>spotkania Teams</b> z kalendarza (z linkiem „Dołącz”), <b>uczestników jako Osoby</b> (bez duplikatów) i <b>zadania z Microsoft To Do / Teams</b> (folder Zadania w Outlooku i jego listy).</p>
    <div class="row" style="gap:10px;margin:4px 0 12px;font-size:13px"><span class="${st.tlErr ? 'late' : st.tlLast ? 'ok' : 'muted'}">●</span> ${!srv ? 'Wymaga uruchomienia TM przez start-TM.bat' : st.tlErr ? 'Błąd: ' + esc(st.tlErr) : st.tlLast ? `Ostatnia synchronizacja ${ago(st.tlLast)}${s ? ` · w Outlooku: ${plN(s.m || 0, 'spotkanie Teams', 'spotkania Teams', 'spotkań Teams')}, ${plN(s.t || 0, 'otwarte zadanie', 'otwarte zadania', 'otwartych zadań')}` : ''}` : 'Jeszcze nie synchronizowano'}</div>
    <label class="row" style="gap:8px;font-size:14px;margin-bottom:8px"><input type="checkbox" name="tlOn" ${c.on ? 'checked' : ''}> <b>Synchronizuj automatycznie</b></label>
    <label class="row" style="gap:8px;font-size:14px;margin-bottom:6px"><input type="checkbox" name="tlMtg" ${c.mtg ? 'checked' : ''}> Spotkania Teams z kalendarza (7 dni wstecz, ${IMP_FWD_DEF} dni do przodu; bez duplikatów ze spotkaniami z linku ICS)</label>
    <label class="row" style="gap:8px;font-size:14px;margin-bottom:6px"><input type="checkbox" name="tlPpl" ${c.ppl ? 'checked' : ''}> Uczestnicy spotkań Teams jako Osoby (powiadomienie „Utworzono nową osobę”)</label>
    <label class="row" style="gap:8px;font-size:14px;margin-bottom:6px"><input type="checkbox" name="tlTodo" ${c.todo ? 'checked' : ''}> Zadania z To Do / Teams jako zadania JOB (priorytet z ważności, termin z To Do)</label>
    <label class="row" style="gap:8px;font-size:14px;margin-bottom:8px"><input type="checkbox" name="tlDone" ${c.done ? 'checked' : ''}> Zakończenie zadania w TM oznacza je jako wykonane w To Do</label>
    <label class="f" style="max-width:240px">Synchronizuj co (minut)<input class="input" type="number" min="5" max="240" name="tlEvery" value="${c.every}"></label>
    <p class="hint" style="margin:8px 0 0">Wymaga klasycznego Outlooka (nowy Outlook nie udostępnia danych innym programom). Zadania z Plannera widoczne w Teams trafiają do To Do dopiero po przypisaniu do Ciebie.</p>
    <div class="row" style="justify-content:flex-end;margin-top:12px;gap:6px"><button type="button" class="btn" data-action="tl-sync" ${srv ? '' : 'disabled'}>⟨i:refresh⟩ Synchronizuj teraz</button><button class="btn primary">Zapisz</button></div></form>`;
}
Object.assign(A, { 'tl-sync': () => teamsLocalSync({silent: false, force: true}) });

Object.assign(A, {
  'cloud-sync': () => CLOUD.ready ? cloudPull({silent: false}) : cloudLoad(LS.get('tm:data:' + S.user.email, null)),
  'cloud-relogin': () => CLOUD.app ? CLOUD.login() : cloudBoot(null),
  'phone-copy': async el => { try { await navigator.clipboard.writeText(el.dataset.l); toast('Skopiowano link – otwórz go na telefonie'); } catch { prompt('Skopiuj link:', el.dataset.l); } },
  'phone-mail': el => { location.href = `mailto:?subject=${encodeURIComponent('TM na telefon – link konfiguracyjny')}&body=${encodeURIComponent('Otwórz na telefonie (Edge/Chrome/Safari), zaloguj się kontem InfoConsulting, potem „Dodaj do ekranu głównego”:\n\n' + el.dataset.l)}`; }
});

// PWA: instalacja (pasek zadań Windows / ekran główny telefonu)
let deferredPrompt;
window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); deferredPrompt = e; $('#installBtn').style.display = 'inline-flex'; });
$('#installBtn').addEventListener('click', async () => { if (!deferredPrompt) return; deferredPrompt.prompt(); await deferredPrompt.userChoice; deferredPrompt = null; $('#installBtn').style.display = 'none'; });
if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) navigator.serviceWorker.register('sw.js').catch(() => { });

(async () => {
  tabClaim();
  const u = LS.get('tm:user', null);
  if (CLOUD.on) await cloudBoot(u); else if (u) login(u); else showLogin();
  if (CFG.MS_CLIENT_ID) { try { if (await MS.init() && MS.acct) { S.conn.ms = true; syncExternal(); } } catch { } }
  teamsBoot(); setTimeout(() => teamsLocalSync({silent: true}), 12000); setInterval(() => teamsLocalSync({silent: true}), 5 * 60e3); setTimeout(() => mailPeopleSync({silent: true}), 20000); setTimeout(() => mailPeopleSync({silent: true}), 90000);
})();
