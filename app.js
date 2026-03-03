/* ================================================
   Work Order Map PWA — app.js
   ================================================ */

'use strict';

// ── Version ───────────────────────────────────
const APP_VERSION = 'v2.5';

// ── Google Sheets published CSV URL ───────────
// Dispatcher: File → Share → Publish to web → CSV → paste the URL here
const SHEETS_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTmjcAZ6v2j5Lrs_XhyPovwduIdtVjfnQKr0bqOau-MSyW3nuePnfoHsFAU4-OJWxilBqxCL3DKe2AA/pub?gid=0&single=true&output=csv';

// ── Storage keys ──────────────────────────────
const RECORDS_KEY = 'wo_records';
const GEOCACHE_KEY = 'wo_geocache';
const COMPLETIONS_KEY = 'wo_completions';
const MAP_STYLE_KEY = 'wo_map_style';
const ENGINEER_KEY = 'wo_engineer';
const POINTS_KEY = 'wo_points';


// ── State ─────────────────────────────────────
let workOrders = [];       // parsed CSV rows
let geocodedPoints = [];       // { lat, lng, row }
let geocodeFailures = [];       // { row, query }
let completions = {};       // { [workorder]: { date } }
let mapInitialized = false;
let leafletMap = null;
let mapMarkers = [];
let userLocationMarker = null;
let gpsWatching = false;
let gpsAutoStopTimer = null;
let activeRow = null;     // row shown in detail sheet
let sheetJustOpened = false;    // guard: prevents same tap from immediately closing the sheet
let dueTodayActive = false;
let selectedEngineer = '';
let pendingRecords = null;
let mapStyle = localStorage.getItem(MAP_STYLE_KEY) || 'auto';
let darkMQ = null;

// ── DOM refs ──────────────────────────────────
const splash = document.getElementById('splash');
const viewHome = document.getElementById('view-home');
const viewMap = document.getElementById('view-map');
const csvFileInput = document.getElementById('csv-file-input');
const csvReloadInput = document.getElementById('csv-reload-input');
const btnLoadNew = document.getElementById('btn-load-new');
const woCountBadge = document.getElementById('wo-count-badge');
const geocodeBar = document.getElementById('geocode-bar');
const geocodeBarText = document.getElementById('geocode-bar-text');
const geocodeBarFill = document.getElementById('geocode-bar-fill');
const notFoundBanner = document.getElementById('not-found-banner');
const notFoundText = document.getElementById('not-found-text');
const btnFixAddresses = document.getElementById('btn-fix-addresses');
const detailSheet = document.getElementById('detail-sheet');
const detailClose = document.getElementById('detail-close');
const detailNotifChip = document.getElementById('detail-notif-chip');
const detailWoNum = document.getElementById('detail-wo-num');
const detailAddress = document.getElementById('detail-address');
const detailMis = document.getElementById('detail-mis');
const detailCity = document.getElementById('detail-city');
const detailLocRow = document.getElementById('detail-loc-row');
const detailLoc = document.getElementById('detail-loc');
const detailWorkType = document.getElementById('detail-work-type');
const detailNotifCode = document.getElementById('detail-notif-code');
const detailMeterNum = document.getElementById('detail-meter-num');
const detailMeterSize = document.getElementById('detail-meter-size');
const detailLastReadRow = document.getElementById('detail-last-read-row');
const detailLastRead = document.getElementById('detail-last-read');
const detailRefErt = document.getElementById('detail-ref-ert');
const detailDatesRow = document.getElementById('detail-dates-row');
const detailDates = document.getElementById('detail-dates');
const detailNavLink = document.getElementById('detail-nav-link');
const geocodeFixModal = document.getElementById('geocode-fix-modal');
const geocodeFixList = document.getElementById('geocode-fix-list');
const geocodeFixClose = document.getElementById('geocode-fix-close');
const engineerFilterSel = document.getElementById('engineer-filter');
const btnDueToday = document.getElementById('btn-due-today');
const toast = document.getElementById('toast');
const btnComplete = document.getElementById('btn-complete');
const overdueWarning = document.getElementById('detail-overdue-warning');
const overdueText = document.getElementById('detail-overdue-text');
const overdueDismiss = document.getElementById('detail-overdue-dismiss');
const statusBar = document.getElementById('status-bar');
const btnMapStyle = document.getElementById('btn-map-style');
const mapStyleMenu = document.getElementById('map-style-menu');
const viewEngineer = document.getElementById('view-engineer');
const engineerList = document.getElementById('engineer-list');
const mergeModal = document.getElementById('merge-modal');
const mergeModalDesc = document.getElementById('merge-modal-desc');
const btnMergeKeep = document.getElementById('btn-merge-keep');
const btnMergeFresh = document.getElementById('btn-merge-fresh');

// ── Helpers ───────────────────────────────────
function fmtDate(str) {
  if (!str) return str;
  try {
    const d = new Date(str);
    if (isNaN(d)) return str;
    const mon = d.toLocaleDateString('en-CA', { month: 'short' });
    const day = String(d.getDate()).padStart(2, '0');
    const wday = d.toLocaleDateString('en-CA', { weekday: 'short' });
    return `${mon} ${day} (${wday})`;
  } catch (_) { return str; }
}

function esc(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

let toastTimer = null;
function showToast(msg, isError = false) {
  toast.textContent = msg;
  toast.style.background = isError ? 'rgba(180,30,30,0.92)' : 'rgba(26,36,56,0.92)';
  toast.classList.remove('hidden', 'fade-out');
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.classList.add('fade-out');
    setTimeout(() => toast.classList.add('hidden'), 400);
  }, 2800);
}

// ── CSV Parsing ───────────────────────────────
function parseCSV(text) {
  const lines = [];
  let cur = '', inQ = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') { inQ = !inQ; cur += ch; }
    else if ((ch === '\n' || ch === '\r') && !inQ) {
      if (cur.trim()) lines.push(cur);
      cur = '';
      if (ch === '\r' && text[i + 1] === '\n') i++;
    } else { cur += ch; }
  }
  if (cur.trim()) lines.push(cur);

  const parseLine = (line) => {
    const fields = []; let f = '', q = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (q && line[i + 1] === '"') { f += '"'; i++; }
        else q = !q;
      } else if (ch === ',' && !q) {
        fields.push(f.trim());
        f = '';
      } else { f += ch; }
    }
    fields.push(f.trim());
    return fields;
  };

  const headers = parseLine(lines[0]);
  return lines.slice(1).map(line => {
    const vals = parseLine(line);
    const obj = {};
    headers.forEach((h, i) => {
      obj[h.trim()] = (vals[i] || '').replace(/^"|"$/g, '').trim();
    });
    return obj;
  });
}

// ── Google Sheets CSV fetch ────────────────────
async function fetchCSVFromSheets() {
  if (!SHEETS_CSV_URL || SHEETS_CSV_URL.startsWith('PASTE_')) {
    console.log('[Sheets] URL not configured — skipping fetch');
    return null;
  }
  console.log('[Sheets] Fetching:', SHEETS_CSV_URL);
  try {
    const res = await fetch(SHEETS_CSV_URL, { cache: 'no-store' });
    console.log('[Sheets] Response status:', res.status, res.ok);
    if (!res.ok) return null;
    const text = await res.text();
    console.log('[Sheets] CSV length:', text.length, '— first 200 chars:', text.slice(0, 200));
    return text;
  } catch (err) {
    console.error('[Sheets] Fetch error:', err);
    return null;
  }
}

// ── Completions persistence ───────────────────
function loadCompletions() {
  try { return JSON.parse(localStorage.getItem(COMPLETIONS_KEY) || '{}'); } catch (_) { return {}; }
}
function saveCompletions() {
  try { localStorage.setItem(COMPLETIONS_KEY, JSON.stringify(completions)); } catch (_) { }
}

// ── Geocache ──────────────────────────────────
function loadGeoCache() {
  try { return JSON.parse(localStorage.getItem(GEOCACHE_KEY) || '{}'); } catch (_) { return {}; }
}
function saveGeoCache(cache) {
  try { localStorage.setItem(GEOCACHE_KEY, JSON.stringify(cache)); } catch (_) { }
}

// ── Persisted geocoded points ─────────────────
function savePoints() {
  try { localStorage.setItem(POINTS_KEY, JSON.stringify(geocodedPoints)); } catch (_) { }
}
function loadPoints() {
  try { return JSON.parse(localStorage.getItem(POINTS_KEY) || '[]'); } catch (_) { return []; }
}

// ── Address cleaning for geocoding ───────────
// Strip the Mis Address portion (unit/suite info) from Street Address
function cleanAddressForGeocode(streetAddress, misAddress) {
  let addr = (streetAddress || '').trim();
  if (misAddress) {
    // Strip /U: prefix from Mis Address to get the unit identifier text
    const unitText = misAddress.replace(/^\/U:/i, '').trim();
    if (unitText) {
      // Escape any regex special chars in the unit text before substituting
      const escaped = unitText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      addr = addr.replace(new RegExp(escaped, 'gi'), '');
    }
  }
  // Clean up leftover double-commas, extra spaces, trailing commas
  return addr
    .replace(/,\s*,/g, ',')
    .replace(/\s{2,}/g, ' ')
    .trim()
    .replace(/[,\s]+$/, '');
}

// ── Single-address geocoder (Nominatim) ───────
function geocodeAddress(streetAddress, misAddress, city, cache) {
  const cleanAddr = cleanAddressForGeocode(streetAddress, misAddress);
  const cacheKey = `${cleanAddr},${city}`.toLowerCase();

  if (cache[cacheKey]) return Promise.resolve({ coords: cache[cacheKey], count: 0 });

  const doFetch = (q) =>
    fetch(
      `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(q)}&format=json&limit=1&countrycodes=ca`,
      { headers: { 'User-Agent': 'WorkOrderMapPWA/1.0' } }
    )
      .then(r => r.json())
      .then(d => (d && d[0]) ? { lat: parseFloat(d[0].lat), lng: parseFloat(d[0].lon) } : null)
      .catch(() => null);

  // Attempt 1: cleanAddr + Ontario
  return doFetch(`${cleanAddr}, Ontario`).then(coords => {
    if (coords) {
      cache[cacheKey] = coords;
      saveGeoCache(cache);
      return { coords, count: 1 };
    }
    // Attempt 2: cleanAddr + City + Ontario
    return doFetch(`${cleanAddr}, ${city}, Ontario`).then(coords2 => {
      if (coords2) { cache[cacheKey] = coords2; saveGeoCache(cache); }
      return { coords: coords2, count: 2 };
    });
  });
}

// ── Batch geocoding with rate-limiting ────────
function geocodeAllRecords(progressCb) {
  const cache = loadGeoCache();
  const tasks = workOrders.map(row => ({
    row,
    streetAddress: (row['Street Address'] || '').trim(),
    misAddress: (row['Mis Address'] || '').trim(),
    city: (row['City'] || '').trim(),
  })).filter(t => t.streetAddress &&
    (!selectedEngineer || (t.row['engineer'] || '').trim() === selectedEngineer));

  const results = [], failures = [];
  let done = 0;
  const total = tasks.length;

  function processNext(i) {
    if (i >= total) return Promise.resolve({ points: results, failures });
    const { row, streetAddress, misAddress, city } = tasks[i];
    return geocodeAddress(streetAddress, misAddress, city, cache).then(({ coords, count }) => {
      done++;
      progressCb(done, total);
      if (coords) results.push({ lat: coords.lat, lng: coords.lng, row });
      else failures.push({
        row, streetAddress, misAddress, city,
        query: `${cleanAddressForGeocode(streetAddress, misAddress)}, Ontario`
      });
      const delay = count * 1050;
      return new Promise(res => setTimeout(res, delay)).then(() => processNext(i + 1));
    });
  }

  return processNext(0);
}

// ── Marker icons ──────────────────────────────
const NOTIF_CODE = (row) => (row['Notification Code'] || '').trim().toUpperCase();

function isRedLock(row) { return NOTIF_CODE(row) === 'RDLK'; }
function isBlackLock(row) { return ['LKFS', 'TLOC', 'LOCK'].includes(NOTIF_CODE(row)); }
function isOpenLock(row) { return NOTIF_CODE(row) === 'LKOO'; }
function isBattery(row) { return NOTIF_CODE(row) === 'RMBE'; }
function isTamper(row) { return NOTIF_CODE(row) === 'TC01'; }
function isMove(row) { return NOTIF_CODE(row) === 'MOVE'; }
function isSpecialRead(row) { return ['MT31', 'ESTS', 'CKRD'].includes(NOTIF_CODE(row)); }

function getMarkerColor(row) {
  if (isRedLock(row)) return '#ef4444';   // red lock
  if (isBlackLock(row)) return '#1a1a1a';   // black lock
  if (isOpenLock(row)) return '#4b5563';   // dark grey open lock
  if (isBattery(row)) return '#f97316';   // orange battery
  if (isTamper(row)) return '#0d9488';   // turquoise tamper
  if (isSpecialRead(row)) return '#eab308';   // yellow special read
  return '#3b82f6';                           // default blue
}

function makeCircleIcon(color) {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 22 22">
    <circle cx="11" cy="11" r="9" fill="${color}" stroke="#fff" stroke-width="2.5"/>
  </svg>`;
  return L.divIcon({ html: svg, className: '', iconSize: [22, 22], iconAnchor: [11, 11], popupAnchor: [0, -14] });
}

// RDLK must be done Mon–Thu; if targetfinish is Fri/Sat/Sun, shift back to that Thursday
function effectiveLockDate(row) {
  const tf = (row['targetfinish'] || '').trim();
  if (!tf) return null;
  try {
    const d = new Date(tf);
    d.setHours(0, 0, 0, 0);
    const day = d.getDay(); // 0=Sun,1=Mon,...,5=Fri,6=Sat
    if (day === 5) d.setDate(d.getDate() - 1); // Fri → Thu
    if (day === 6) d.setDate(d.getDate() - 2); // Sat → Thu
    if (day === 0) d.setDate(d.getDate() - 3); // Sun → Thu (prev week)
    return d;
  } catch (_) { return null; }
}

function isLockEndToday(row) {
  const effective = effectiveLockDate(row);
  if (!effective) return false;
  const today = new Date();
  return effective.getFullYear() === today.getFullYear() &&
    effective.getMonth() === today.getMonth() &&
    effective.getDate() === today.getDate();
}

function isLockEndPast(row) {
  const effective = effectiveLockDate(row);
  if (!effective) return false;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  return effective < today;
}

function isAptToday(row) {
  const val = (row['aptstart'] || '').trim();
  if (!val) return false;
  try {
    const d = new Date(val);
    const today = new Date();
    return d.getFullYear() === today.getFullYear() &&
      d.getMonth() === today.getMonth() &&
      d.getDate() === today.getDate();
  } catch (_) { return false; }
}

function isDueToday(row) {
  if (isRedLock(row)) return isLockEndToday(row);
  if (isTamper(row)) return true;                    // TC01
  if (isBlackLock(row)) return true;                    // LKFS, TLOC, LOCK
  if (isSpecialRead(row)) return isAptToday(row);        // MT31, ESTS, CKRD — only if aptstart = today
  return false;
}

// badge: null | 'star' | 'exclamation'
function makeLockIcon(bgColor, keyColor, badge = null) {
  let badgeSvg = '';
  if (badge === 'star') {
    badgeSvg = `<polygon points="20,2 20.88,4.29 23.33,4.42 21.43,5.96 22.06,8.33 20,7 17.94,8.33 18.57,5.96 16.67,4.42 19.12,4.29"
         fill="#fbbf24" stroke="#fff" stroke-width="0.6"/>`;
  } else if (badge === 'exclamation') {
    badgeSvg = `<circle cx="21" cy="5" r="3.8" fill="#fff" stroke="#dc2626" stroke-width="1"/>
      <text x="21" y="7.8" text-anchor="middle" font-size="6" font-weight="bold" fill="#dc2626">!</text>`;
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="26" height="26" viewBox="0 0 26 26">
    <circle cx="13" cy="13" r="11" fill="${bgColor}" stroke="#fff" stroke-width="2.5"/>
    <rect x="9" y="12" width="8" height="6" rx="1.2" fill="#fff"/>
    <path d="M10.5 12v-2a2.5 2.5 0 0 1 5 0v2" fill="none" stroke="#fff" stroke-width="1.6" stroke-linecap="round"/>
    <circle cx="13" cy="15" r="1" fill="${keyColor}"/>
    ${badgeSvg}
  </svg>`;
  return L.divIcon({ html: svg, className: '', iconSize: [26, 26], iconAnchor: [13, 13], popupAnchor: [0, -16] });
}

function makeOpenLockIcon() {
  // Dark grey circle with an open padlock (shackle raised on right side)
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="26" height="26" viewBox="0 0 26 26">
    <circle cx="13" cy="13" r="11" fill="#4b5563" stroke="#fff" stroke-width="2.5"/>
    <rect x="9" y="13" width="8" height="6" rx="1.2" fill="#fff"/>
    <path d="M10.5 13v-3a2.5 2.5 0 0 1 5 0" fill="none" stroke="#fff" stroke-width="1.6" stroke-linecap="round"/>
    <circle cx="13" cy="16" r="1" fill="#4b5563"/>
  </svg>`;
  return L.divIcon({ html: svg, className: '', iconSize: [26, 26], iconAnchor: [13, 13], popupAnchor: [0, -16] });
}

function makeBatteryIcon() {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="26" height="26" viewBox="0 0 26 26">
    <circle cx="13" cy="13" r="11" fill="#f97316" stroke="#fff" stroke-width="2.5"/>
    <rect x="7" y="10.5" width="10" height="5.5" rx="1.2" fill="none" stroke="#fff" stroke-width="1.5"/>
    <rect x="17" y="12" width="2" height="2.5" rx="0.5" fill="#fff"/>
    <line x1="9.5" y1="13.25" x2="11.5" y2="13.25" stroke="#fff" stroke-width="1.5" stroke-linecap="round"/>
  </svg>`;
  return L.divIcon({ html: svg, className: '', iconSize: [26, 26], iconAnchor: [13, 13], popupAnchor: [0, -16] });
}

function makeTamperIcon() {
  // Turquoise circle: closed lock body with a "!" exclamation — tamper alert
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="26" height="26" viewBox="0 0 26 26">
    <circle cx="13" cy="13" r="11" fill="#0d9488" stroke="#fff" stroke-width="2.5"/>
    <rect x="9.5" y="12.5" width="7" height="5.5" rx="1.2" fill="#fff"/>
    <path d="M11 12.5v-1.8a2 2 0 0 1 4 0v1.8" fill="none" stroke="#fff" stroke-width="1.5" stroke-linecap="round"/>
    <line x1="13" y1="14.2" x2="13" y2="16" stroke="#0d9488" stroke-width="1.4" stroke-linecap="round"/>
    <circle cx="13" cy="17" r="0.6" fill="#0d9488"/>
  </svg>`;
  return L.divIcon({ html: svg, className: '', iconSize: [26, 26], iconAnchor: [13, 13], popupAnchor: [0, -16] });
}

function makeMoveIcon() {
  // Blue circle with a bold right-pointing arrow — represents relocation/move
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="26" height="26" viewBox="0 0 26 26">
    <circle cx="13" cy="13" r="11" fill="#3b82f6" stroke="#fff" stroke-width="2.5"/>
    <line x1="7" y1="13" x2="17" y2="13" stroke="#fff" stroke-width="2" stroke-linecap="round"/>
    <polyline points="13,9 17,13 13,17" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
  </svg>`;
  return L.divIcon({ html: svg, className: '', iconSize: [26, 26], iconAnchor: [13, 13], popupAnchor: [0, -16] });
}

function makeSpecialReadIcon() {
  // Yellow circle with a checkmark — special/check read
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="26" height="26" viewBox="0 0 26 26">
    <circle cx="13" cy="13" r="11" fill="#eab308" stroke="#fff" stroke-width="2.5"/>
    <polyline points="7.5,13.5 11,17 18.5,9" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
  </svg>`;
  return L.divIcon({ html: svg, className: '', iconSize: [26, 26], iconAnchor: [13, 13], popupAnchor: [0, -16] });
}

function makeBaseMarkerIcon(row) {
  if (isRedLock(row)) {
    const badge = isLockEndPast(row) ? 'exclamation' : isLockEndToday(row) ? 'star' : null;
    return makeLockIcon('#ef4444', '#ef4444', badge);
  }
  if (isBlackLock(row)) return makeLockIcon('#1a1a1a', '#1a1a1a');
  if (isOpenLock(row)) return makeOpenLockIcon();
  if (isBattery(row)) return makeBatteryIcon();
  if (isTamper(row)) return makeTamperIcon();
  if (isMove(row)) return makeMoveIcon();
  if (isSpecialRead(row)) return makeSpecialReadIcon();
  return makeCircleIcon(getMarkerColor(row));
}

function makeMarkerIcon(row) {
  if (completions[row['Workorder'] || '']) return makeCircleIcon('#d1d5db');
  return makeBaseMarkerIcon(row);
}

function refreshMarkerIcon(workorder) {
  const entry = mapMarkers.find(m => (m.row['Workorder'] || '') === workorder);
  if (entry) entry.marker.setIcon(makeMarkerIcon(entry.row));
}

// ── Map initialisation ────────────────────────
const TILES = {
  light: {
    url: 'https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png',
    attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> © <a href="https://carto.com/">CARTO</a>',
  },
  dark: {
    url: 'https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png',
    attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> © <a href="https://carto.com/">CARTO</a>',
  },
  satellite: {
    url: 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
    attribution: '© <a href="https://www.esri.com/">Esri</a> © OpenStreetMap',
  },
  standard: {
    url: 'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
    attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
  },
};

let tileLayerRef = null;

function applyTileTheme(style) {
  if (!leafletMap) return;
  mapStyle = style;
  let cfg;
  if (style === 'auto') {
    cfg = (darkMQ && darkMQ.matches) ? TILES.dark : TILES.light;
  } else {
    cfg = TILES[style] || TILES.light;
  }
  if (tileLayerRef) leafletMap.removeLayer(tileLayerRef);
  tileLayerRef = L.tileLayer(cfg.url, { attribution: cfg.attribution, maxZoom: 19 });
  tileLayerRef.addTo(leafletMap);
  tileLayerRef.bringToBack();
}

function initLeafletMap() {
  if (mapInitialized) return;

  leafletMap = L.map('map-container', { zoomControl: true });

  darkMQ = window.matchMedia('(prefers-color-scheme: dark)');
  applyTileTheme(mapStyle);
  darkMQ.addEventListener('change', () => { if (mapStyle === 'auto') applyTileTheme('auto'); });

  // GPS locate button
  const LocateControl = L.Control.extend({
    options: { position: 'topright' },
    onAdd() {
      const btn = L.DomUtil.create('button', 'locate-btn leaflet-bar');
      btn.title = 'Show my location';
      btn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24"
        fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="7" r="4"/><path d="M4 21v-1a8 8 0 0 1 16 0v1"/></svg>`;
      L.DomEvent.on(btn, 'click', startLocating);
      return btn;
    },
  });
  new LocateControl().addTo(leafletMap);

  leafletMap.on('locationfound', onLocationFound);
  leafletMap.on('locationerror', () => showToast('Could not get your location', true));
  leafletMap.on('move', resetGpsTimer);
  leafletMap.on('zoom', resetGpsTimer);

  mapInitialized = true;
}

// ── GPS tracking ──────────────────────────────
const GPS_TIMEOUT_MS = 5 * 60 * 1000;

function startLocating() {
  if (!mapInitialized || !leafletMap) return;
  if (!('geolocation' in navigator)) { showToast('Geolocation not supported', true); return; }
  leafletMap.locate({ setView: false, watch: true });
  gpsWatching = true;
  resetGpsTimer();
}

function onLocationFound(e) {
  if (userLocationMarker) {
    userLocationMarker.setLatLng(e.latlng);
  } else {
    const icon = L.divIcon({
      html: `<svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24"
        fill="#1e50a2" stroke="#fff" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="7" r="4"/><path d="M4 21v-1a8 8 0 0 1 16 0v1"/></svg>`,
      className: '',
      iconSize: [28, 28],
      iconAnchor: [14, 28],
      popupAnchor: [0, -30],
    });
    userLocationMarker = L.marker(e.latlng, { icon, zIndexOffset: 1000 })
      .addTo(leafletMap)
      .bindPopup('<strong>Your Location</strong>');
    leafletMap.setView(e.latlng, 16);
  }
}

function resetGpsTimer() {
  if (!gpsWatching) return;
  clearTimeout(gpsAutoStopTimer);
  gpsAutoStopTimer = setTimeout(() => {
    if (leafletMap) leafletMap.stopLocate();
    gpsWatching = false;
    showToast('GPS stopped after 5 minutes of inactivity.');
  }, GPS_TIMEOUT_MS);
}

// ── Place markers ─────────────────────────────
function clearMapMarkers() {
  mapMarkers.forEach(({ marker }) => marker.remove());
  mapMarkers = [];
}

function placeMarkers(points, zoomToFit = true) {
  if (!mapInitialized) return;
  clearMapMarkers();
  const bounds = [];

  points.forEach(({ lat, lng, row }) => {
    const icon = makeMarkerIcon(row);
    const addr = (row['Street Address'] || '').trim();
    const mapsUrl = `https://www.google.com/maps/dir/?api=1&destination=${lat},${lng}`;

    const popup = `<strong>${esc(addr)}</strong>
      <br><a href="${mapsUrl}" target="_blank" rel="noopener" class="popup-nav-link">↪ Navigate</a>`;

    const marker = L.marker([lat, lng], { icon })
      .addTo(leafletMap)
      .bindPopup(popup, { autoClose: false, closeOnClick: false });

    marker.on('click', (e) => {
      L.DomEvent.stopPropagation(e);
      if (e.originalEvent) e.originalEvent.stopPropagation();
      marker.closePopup();
      openDetailSheet(row, lat, lng);
    });

    mapMarkers.push({ marker, row });
    bounds.push([lat, lng]);
  });

  if (bounds.length && zoomToFit) {
    leafletMap.invalidateSize();
    leafletMap.fitBounds(bounds, { padding: [48, 48] });
  }
}

// ── Add a single marker (from geocode fix) ────
function addSingleMarker(coords, row) {
  const icon = makeMarkerIcon(row);
  const addr = (row['Street Address'] || '').trim();
  const mapsUrl = `https://www.google.com/maps/dir/?api=1&destination=${coords.lat},${coords.lng}`;
  const popup = `<strong>${esc(addr)}</strong>
    <br><a href="${mapsUrl}" target="_blank" rel="noopener" class="popup-nav-link">↪ Navigate</a>`;

  const marker = L.marker([coords.lat, coords.lng], { icon })
    .addTo(leafletMap)
    .bindPopup(popup, { autoClose: false, closeOnClick: false });

  marker.on('click', (e) => {
    L.DomEvent.stopPropagation(e);
    if (e.originalEvent) e.originalEvent.stopPropagation();
    marker.closePopup();
    openDetailSheet(row, coords.lat, coords.lng);
  });

  mapMarkers.push({ marker, row });
}

// ── Not-found banner ──────────────────────────
function updateNotFoundBar() {
  const n = geocodeFailures.length;
  if (n === 0) {
    notFoundBanner.classList.add('hidden');
  } else {
    notFoundText.textContent = `${n} address${n > 1 ? 'es' : ''} not found`;
    notFoundBanner.classList.remove('hidden');
  }
}

// ── Chip class for notification type ──────────
function chipClass(type) {
  const t = (type || '').toUpperCase();
  if (t === 'ZB') return 'chip-zb';
  if (t === 'YD') return 'chip-yd';
  if (t === 'YE') return 'chip-ye';
  return 'chip-default';
}

// ── Detail sheet ──────────────────────────────
function openDetailSheet(row, lat, lng) {
  sheetJustOpened = true;
  clearTimeout(openDetailSheet._guard);
  openDetailSheet._guard = setTimeout(() => { sheetJustOpened = false; }, 600);
  activeRow = row;

  const notifType = (row['Notification Type'] || '').trim();
  detailNotifChip.textContent = notifType || '—';
  detailNotifChip.className = `notif-chip ${chipClass(notifType)}`;
  detailWoNum.textContent = `WO# ${row['Workorder'] || '—'}`;

  detailAddress.textContent = (row['Street Address'] || '').trim() || '—';

  const mis = (row['Mis Address'] || '').trim();
  if (mis) {
    detailMis.textContent = mis;
    detailMis.classList.remove('hidden');
  } else {
    detailMis.classList.add('hidden');
  }

  detailCity.textContent = (row['City'] || '').trim();

  const loc = (row['Device Location Note'] || '').trim();
  if (loc) {
    detailLoc.textContent = loc;
    detailLocRow.classList.remove('hidden');
  } else {
    detailLocRow.classList.add('hidden');
  }

  detailWorkType.textContent = (row['Meter Location'] || '').trim() || '—';
  detailNotifCode.textContent = (row['Notification Code'] || '').trim() || '—';
  detailMeterNum.textContent = (row['Meter Number'] || '').trim() || '—';
  detailMeterSize.textContent = (row['Meter Size'] || '').trim() || '—';

  const grid = (row['Grid'] || '').trim();
  const refErt = (row['Reference ERT'] || '').trim();
  detailLastRead.textContent = grid || '—';
  detailRefErt.textContent = refErt || '—';
  if (grid || refErt) {
    detailLastReadRow.classList.remove('hidden');
  } else {
    detailLastReadRow.classList.add('hidden');
  }

  const ts = (row['targetstart'] || '').trim();
  const tf = (row['targetfinish'] || '').trim();
  if (ts || tf) {
    detailDates.textContent = ts && tf ? `${fmtDate(ts)} → ${fmtDate(tf)}` : fmtDate(ts || tf);
    detailDatesRow.classList.remove('hidden');
  } else {
    detailDatesRow.classList.add('hidden');
  }

  detailNavLink.href = `https://www.google.com/maps/dir/?api=1&destination=${lat},${lng}`;

  // Set complete button state
  const wo = (row['Workorder'] || '').trim();
  if (completions[wo]) {
    btnComplete.classList.add('done');
    btnComplete.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24"
      fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
      <polyline points="20 6 9 17 4 12"/></svg> Completed`;
  } else {
    btnComplete.classList.remove('done');
    btnComplete.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24"
      fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
      <polyline points="20 6 9 17 4 12"/></svg> Complete`;
  }

  if (isRedLock(row) && isLockEndPast(row)) {
    const tf = (row['targetfinish'] || '').trim();
    overdueText.textContent = `⚠ Lock end date has passed (${tf})`;
    overdueWarning.classList.remove('hidden');
  } else {
    overdueWarning.classList.add('hidden');
  }

  detailSheet.classList.remove('hidden');
  requestAnimationFrame(() => detailSheet.classList.add('open'));
}

function closeDetailSheet() {
  detailSheet.classList.remove('open');
  setTimeout(() => detailSheet.classList.add('hidden'), 310);
  activeRow = null;
}

// ── Geocode fix modal ─────────────────────────
function showGeocodeFix() {
  geocodeFixList.innerHTML = '';
  const cache = loadGeoCache();

  geocodeFailures.forEach((f, idx) => {
    const item = document.createElement('div');
    item.className = 'geocode-fix-item';
    item.innerHTML = `
      <div class="geocode-fix-addr">${esc((f.row['Street Address'] || '').trim())}, ${esc((f.row['City'] || '').trim())}</div>
      <div class="geocode-fix-row">
        <input type="text" class="geocode-fix-input" value="${esc(f.query)}" />
        <button class="geocode-fix-btn">Retry</button>
      </div>
      <div class="geocode-fix-status"></div>
    `;

    const input = item.querySelector('.geocode-fix-input');
    const btn = item.querySelector('.geocode-fix-btn');
    const status = item.querySelector('.geocode-fix-status');

    btn.addEventListener('click', () => {
      const q = input.value.trim();
      if (!q) return;
      btn.disabled = true;
      btn.textContent = '…';
      status.textContent = '';

      fetch(
        `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(q)}&format=json&limit=1&countrycodes=ca`,
        { headers: { 'User-Agent': 'WorkOrderMapPWA/1.0' } }
      )
        .then(r => r.json())
        .then(data => {
          if (!data || !data[0]) {
            status.textContent = 'Not found — try a different query';
            status.className = 'geocode-fix-status geocode-fix-fail';
            btn.disabled = false;
            btn.textContent = 'Retry';
            return;
          }
          const coords = { lat: parseFloat(data[0].lat), lng: parseFloat(data[0].lon) };
          const cacheKey = `${cleanAddressForGeocode(f.streetAddress, f.misAddress)},${f.city}`.toLowerCase();
          cache[cacheKey] = coords;
          saveGeoCache(cache);

          geocodedPoints.push({ lat: coords.lat, lng: coords.lng, row: f.row });
          geocodeFailures = geocodeFailures.filter((_, i) => i !== idx);
          savePoints();
          addSingleMarker(coords, f.row);
          updateNotFoundBar();

          status.textContent = 'Located!';
          status.className = 'geocode-fix-status geocode-fix-ok';
          btn.textContent = '✓';
          input.disabled = true;
        })
        .catch(() => {
          status.textContent = 'Network error — try again';
          status.className = 'geocode-fix-status geocode-fix-fail';
          btn.disabled = false;
          btn.textContent = 'Retry';
        });
    });

    geocodeFixList.appendChild(item);
  });

  geocodeFixModal.classList.remove('hidden');
}

// ── Engineer filter ───────────────────────────
function buildEngineerFilter() {
  const engineers = [...new Set(
    workOrders.map(r => (r['engineer'] || '').trim()).filter(Boolean)
  )].sort();
  engineerFilterSel.innerHTML = '<option value="">All Engineers</option>';
  engineers.forEach(eng => {
    const opt = document.createElement('option');
    opt.value = eng;
    opt.textContent = eng;
    engineerFilterSel.appendChild(opt);
  });
  engineerFilterSel.classList.toggle('hidden', engineers.length === 0);
}

function getFilteredPoints() {
  const eng = engineerFilterSel.value;
  return geocodedPoints.filter(p => {
    if (eng && (p.row['engineer'] || '').trim() !== eng) return false;
    if (dueTodayActive && !isDueToday(p.row)) return false;
    return true;
  });
}

function updateBadge() {
  const n = getFilteredPoints().length;
  woCountBadge.textContent = `${n} job${n !== 1 ? 's' : ''}`;
}

function updateStatusBar() {
  const pts = getFilteredPoints();
  const total = pts.length;
  if (total === 0) { statusBar.classList.add('hidden'); return; }
  const done = pts.filter(p => completions[p.row['Workorder'] || '']).length;
  const remaining = total - done;
  statusBar.innerHTML = `<span class="status-bar-version">${APP_VERSION}</span><span>${done} complete · ${remaining} remaining</span>`;
  statusBar.classList.remove('hidden');
}

// ── Load CSV and kick off geocoding ───────────
function applyNewCSV(records, keepPrev) {
  let finalRecords = records;

  if (keepPrev && workOrders.length && selectedEngineer) {
    const prevIncomplete = workOrders.filter(r => {
      const wo = (r['Workorder'] || '').trim();
      return (r['engineer'] || '').trim() === selectedEngineer && wo && !completions[wo];
    });
    if (prevIncomplete.length) {
      const newWoIds = new Set(records.map(r => (r['Workorder'] || '').trim()));
      const extras = prevIncomplete.filter(r => !newWoIds.has((r['Workorder'] || '').trim()));
      finalRecords = [...records, ...extras];
    }
  }

  workOrders = finalRecords;
  geocodedPoints = [];
  geocodeFailures = [];
  selectedEngineer = '';
  clearMapMarkers();

  try { localStorage.setItem(RECORDS_KEY, JSON.stringify(workOrders)); } catch (_) { }
  try { localStorage.removeItem(ENGINEER_KEY); } catch (_) { }
  try { localStorage.removeItem(POINTS_KEY); } catch (_) { }

  showEngineerView();
}

function loadCSV(file) {
  const reader = new FileReader();
  reader.onload = (ev) => {
    const records = parseCSV(ev.target.result);
    if (!records.length || !records[0].hasOwnProperty('Street Address')) {
      showToast('Not a valid work order CSV — missing "Street Address" column', true);
      return;
    }

    // If there are incomplete work orders from a previous session, offer to keep them
    if (workOrders.length && selectedEngineer) {
      const prevIncomplete = workOrders.filter(r => {
        const wo = (r['Workorder'] || '').trim();
        return (r['engineer'] || '').trim() === selectedEngineer && wo && !completions[wo];
      });
      if (prevIncomplete.length) {
        pendingRecords = records;
        mergeModalDesc.textContent =
          `You have ${prevIncomplete.length} incomplete work order${prevIncomplete.length !== 1 ? 's' : ''} from your previous session. Keep them alongside the new file?`;
        mergeModal.classList.remove('hidden');
        return;
      }
    }

    applyNewCSV(records, false);
  };
  reader.readAsText(file);
}

// ── Engineer picker ───────────────────────────
function showEngineerView() {
  viewHome.classList.add('hidden');
  viewMap.classList.add('hidden');
  viewEngineer.classList.remove('hidden');

  const names = [...new Set(
    workOrders.map(r => (r['engineer'] || '').trim()).filter(Boolean)
  )].sort();

  engineerList.innerHTML = '';
  names.forEach(name => {
    const btn = document.createElement('button');
    btn.className = 'engineer-pick-btn';
    btn.textContent = name;
    btn.addEventListener('click', () => selectEngineer(name));
    engineerList.appendChild(btn);
  });
}

function selectEngineer(name) {
  selectedEngineer = name;
  try { localStorage.setItem(ENGINEER_KEY, name); } catch (_) { }

  // Narrow the saved records down to just this engineer's rows so the
  // localStorage payload stays small regardless of how large the full CSV is.
  const myJobs = workOrders.filter(r => (r['engineer'] || '').trim() === name);
  try { localStorage.setItem(RECORDS_KEY, JSON.stringify(myJobs)); } catch (_) { }

  viewEngineer.classList.add('hidden');
  showMapView();
  woCountBadge.textContent = `${myJobs.length} job${myJobs.length !== 1 ? 's' : ''}`;

  geocodeBar.classList.remove('hidden');
  geocodeBarFill.style.width = '0%';
  geocodeBarText.textContent = `Locating addresses… 0 / ${myJobs.length}`;
  notFoundBanner.classList.add('hidden');

  geocodeAllRecords((done, total) => {
    geocodeBarFill.style.width = Math.round((done / total) * 100) + '%';
    geocodeBarText.textContent = `Locating addresses… ${done} / ${total}`;
  }).then(({ points, failures }) => {
    geocodedPoints = points;
    geocodeFailures = failures;
    savePoints();
    geocodeBar.classList.add('hidden');
    placeMarkers(getFilteredPoints(), true);
    updateBadge();
    updateStatusBar();
    updateNotFoundBar();
    if (failures.length) showToast(`${failures.length} address${failures.length > 1 ? 'es' : ''} could not be located`);
  });
}

// ── View switching ────────────────────────────
function showMapView() {
  viewHome.classList.add('hidden');
  viewMap.classList.remove('hidden');
  initLeafletMap();
  // Fix Leaflet size after view switch
  setTimeout(() => leafletMap && leafletMap.invalidateSize(), 100);
}

function showHomeView() {
  viewMap.classList.add('hidden');
  viewHome.classList.remove('hidden');
}

// ── Restore from localStorage ─────────────────
function tryRestoreSession() {
  try {
    const saved = localStorage.getItem(RECORDS_KEY);
    if (!saved) return false;
    const records = JSON.parse(saved);
    if (!Array.isArray(records) || !records.length) return false;
    workOrders = records;
    return true;
  } catch (_) { return false; }
}

// ── Event listeners ───────────────────────────
csvFileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (file) loadCSV(file);
  e.target.value = '';
});

csvReloadInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (file) loadCSV(file);
  e.target.value = '';
});

btnLoadNew.addEventListener('click', async () => {
  const csvText = await fetchCSVFromSheets();
  if (csvText) {
    const records = parseCSV(csvText);
    if (records.length && records[0].hasOwnProperty('Street Address')) {
      showToast('Work orders reloaded from Google Sheets');
      applyNewCSV(records, false);
      return;
    }
  }
  csvReloadInput.click(); // fall back to file picker
});

detailClose.addEventListener('click', closeDetailSheet);

// Tap outside detail sheet to close
viewMap.addEventListener('click', (e) => {
  if (sheetJustOpened) return;
  if (!detailSheet.classList.contains('hidden') && !detailSheet.contains(e.target)) {
    closeDetailSheet();
  }
});

engineerFilterSel.addEventListener('change', () => {
  placeMarkers(getFilteredPoints(), false);
  updateBadge();
  updateStatusBar();
});

btnDueToday.addEventListener('click', () => {
  dueTodayActive = !dueTodayActive;
  btnDueToday.classList.toggle('active', dueTodayActive);
  placeMarkers(getFilteredPoints(), false);
  updateBadge();
  updateStatusBar();
});

btnMapStyle.addEventListener('click', e => {
  e.stopPropagation();
  mapStyleMenu.querySelectorAll('[data-style]').forEach(i =>
    i.classList.toggle('active', i.dataset.style === mapStyle));
  mapStyleMenu.classList.toggle('hidden');
});

mapStyleMenu.querySelectorAll('[data-style]').forEach(item => {
  item.addEventListener('click', e => {
    e.stopPropagation();
    const style = item.dataset.style;
    applyTileTheme(style);
    localStorage.setItem(MAP_STYLE_KEY, style);
    mapStyleMenu.classList.add('hidden');
  });
});

document.addEventListener('click', () => mapStyleMenu.classList.add('hidden'));

btnMergeKeep.addEventListener('click', () => {
  mergeModal.classList.add('hidden');
  applyNewCSV(pendingRecords, true);
  pendingRecords = null;
});

btnMergeFresh.addEventListener('click', () => {
  mergeModal.classList.add('hidden');
  applyNewCSV(pendingRecords, false);
  pendingRecords = null;
});

btnFixAddresses.addEventListener('click', showGeocodeFix);
geocodeFixClose.addEventListener('click', () => geocodeFixModal.classList.add('hidden'));
overdueDismiss.addEventListener('click', () => overdueWarning.classList.add('hidden'));

detailNavLink.addEventListener('click', (e) => {
  if (activeRow && isRedLock(activeRow) && isLockEndPast(activeRow)) {
    e.preventDefault();
    const tf = (activeRow['targetfinish'] || '').trim();
    overdueText.textContent = `⚠ Lock end date has passed (${tf})`;
    overdueWarning.classList.remove('hidden');
  }
});

btnComplete.addEventListener('click', () => {
  if (!activeRow) return;
  const wo = (activeRow['Workorder'] || '').trim();
  if (!wo) return;

  if (completions[wo]) {
    // Undo
    delete completions[wo];
    saveCompletions();
    refreshMarkerIcon(wo);
    btnComplete.classList.remove('done');
    btnComplete.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24"
      fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
      <polyline points="20 6 9 17 4 12"/></svg> Complete`;
  } else {
    // Mark complete
    completions[wo] = { date: new Date().toLocaleString('en-CA') };
    saveCompletions();
    refreshMarkerIcon(wo);
    btnComplete.classList.add('done');
    btnComplete.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24"
      fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
      <polyline points="20 6 9 17 4 12"/></svg> Completed`;
    showToast('Work order completed');
  }
  updateStatusBar();
});

// ── Boot ──────────────────────────────────────
function boot() {
  completions = loadCompletions();

  // Hide splash after letter animation sequence completes (~4.3s)
  setTimeout(() => {
    splash.style.transition = 'opacity 0.4s';
    splash.style.opacity = '0';
    setTimeout(() => splash.remove(), 400);

    // Always try Google Sheets first so engineers see fresh data every time
    fetchCSVFromSheets().then(csvText => {
      if (csvText) {
        const records = parseCSV(csvText);
        if (records.length && records[0].hasOwnProperty('Street Address')) {
          showToast('Work orders loaded from Google Sheets');
          applyNewCSV(records, false);
          return;
        }
      }

      // Sheets not configured or fetch failed — fall back to session restore
      const savedEngineer = localStorage.getItem(ENGINEER_KEY) || '';
      const hasSession = tryRestoreSession();

      if (hasSession && savedEngineer) {
        selectedEngineer = savedEngineer;
        showMapView();
        const myJobs = workOrders.filter(r => (r['engineer'] || '').trim() === savedEngineer);
        woCountBadge.textContent = `${myJobs.length} job${myJobs.length !== 1 ? 's' : ''}`;

        // Restore saved points — wait for Leaflet to size itself first
        const savedPoints = loadPoints();
        setTimeout(() => {
          if (savedPoints.length) {
            geocodedPoints = savedPoints;
            placeMarkers(getFilteredPoints(), true);
            updateBadge();
            updateStatusBar();
            updateNotFoundBar();
          } else {
            // No saved points (e.g. crash during geocoding) — re-geocode from cache
            geocodeAllRecords(() => { }).then(({ points, failures }) => {
              geocodedPoints = points;
              geocodeFailures = failures;
              if (points.length) savePoints();
              placeMarkers(getFilteredPoints(), true);
              updateBadge();
              updateStatusBar();
              updateNotFoundBar();
            });
          }
        }, 150);
      } else if (hasSession) {
        // CSV loaded but no engineer saved — show picker
        showEngineerView();
      } else {
        showHomeView();
      }
    });
  }, 4300);
}

boot();

// ── Service worker registration ───────────────
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('./sw.js').catch(() => { });
}
