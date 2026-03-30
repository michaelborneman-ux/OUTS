/* ================================================
   Meter Reader PWA — App Logic
   Home / Bundle Loading Page
   ================================================ */

(function () {
  'use strict';

  const APP_VERSION = 'v7.3';

  // ─── State ────────────────────────────────────────
  let allRecords = [];         // all CSV rows
  let bundles = [];            // grouped by bundle key
  let sentBundles = new Set(); // bundle keys marked as sent (persisted)
  let bundleRates = {};        // bundle key → { est3, est46, est7p } — persisted
  let pendingRecordsForRates = null; // CSV records waiting for rate entry
  let pendingEmailBundle = null; // bundle queued in email provider picker
  let readerName = '';         // current reader name
  let pendingBundle = null;    // bundle waiting for reader name
  let currentBundle = null;    // bundle open in detail view
  let currentRow = null;    // row open in card detail view
  let cardReturnTarget = 'bundle'; // 'bundle' or 'map'
  let reverseDirection = false;    // when true, save auto-advances to previous card
  let pickStartMode = false;       // route optimization: waiting for user to pick start address
  let originalRowsOrder = null;    // saved copy of rows array before optimization
  let routeOptimized = false;      // whether current bundle has been optimized

  // Default rates by area code (used to pre-fill the rate modal)
  const AREA_RATES = {
    '21': { est3: 2.50, est46: 3.00, est7p: 5.50 },
    '45': { est3: 3.00, est46: 4.00, est7p: 6.00 },
    '47': { est3: 3.00, est46: 8.00, est7p: 8.00 },
  };

  // Map state
  let mapInitialized = false;  // lazy init guard
  let leafletMap = null;   // L.Map instance
  let mapMarkers = [];     // array of { marker, row, bundle }
  let geocodePending = false;  // prevents concurrent geocode runs
  let geocodedPoints = [];     // cached results — reused on subsequent map opens
  let geocodeFailures = [];     // records that could not be geocoded
  let geocodeStale = false;  // set by forceGeocodeBundle — forces re-geocode on next map open
  let pendingFixPoint = null; // { row, bundle, marker, lat, lng } — set when Fix Location is tapped
  let userLocationMarker = null;  // L.marker for the user's GPS position
  let gpsAutoStopTimer = null;
  let gpsWatching = false;
  const GPS_TIMEOUT_MS = 5 * 60 * 1000; // 5 minutes
  let bdShowMissed = false;  // bundle list filter: unread cards
  let bdShowSkipped = false;  // bundle list filter: skipped cards
  let backupBarShown = false;  // session flag — show daily backup bar only once
  let autoSaveTimer = null;   // debounce handle for silent CSV auto-save

  // ─── PIN lock ─────────────────────────────────────
  const APP_PIN = '4959';
  const PIN_KEY = 'outs_pin_unlocked';

  function isPinUnlocked() {
    const today = new Date().toLocaleDateString('en-CA');
    return localStorage.getItem(PIN_KEY) === today;
  }

  function unlockPin() {
    localStorage.setItem(PIN_KEY, new Date().toLocaleDateString('en-CA'));
  }

  // ─── DOM refs ─────────────────────────────────────
  const viewSplash = document.getElementById('view-splash');
  const viewPin = document.getElementById('view-pin');
  const viewHome = document.getElementById('view-home');
  const viewBundle = document.getElementById('view-bundle');
  const viewTotals = document.getElementById('view-totals');
  const homeDate = document.getElementById('home-date');
  const homeTime = document.getElementById('home-time');
  const sumBundles = document.getElementById('sum-bundles');
  const sumCards = document.getElementById('sum-cards');
  const sumRead = document.getElementById('sum-read');
  const sumSent = document.getElementById('sum-sent');
  const homeSearch = document.getElementById('home-search');
  const homeReaderName = document.getElementById('home-reader-name');
  const uploadPrompt = document.getElementById('upload-prompt');
  const bundleList = document.getElementById('bundle-list');
  const csvFileInput = document.getElementById('csv-file-input');
  const toast = document.getElementById('toast');

  // Bundle detail
  const bdDate = document.getElementById('bd-date');
  const bdTime = document.getElementById('bd-time');
  const bdRoutes = document.getElementById('bd-routes');
  const bdTitle = document.getElementById('bd-title');
  const bdArea = document.getElementById('bd-area');
  const bdStatusBadge = document.getElementById('bd-status-badge');
  const bdCountLabel = document.getElementById('bd-count-label');
  const bdCountPct = document.getElementById('bd-count-pct');
  const bdProgressFill = document.getElementById('bd-progress-fill');
  const bdTotal = document.getElementById('bd-total');
  const bdEst3r = document.getElementById('bd-est3r');
  const bdEst3 = document.getElementById('bd-est3');
  const bdEst46r = document.getElementById('bd-est46r');
  const bdEst46 = document.getElementById('bd-est46');
  const bdEst7pr = document.getElementById('bd-est7pr');
  const bdEst7p = document.getElementById('bd-est7p');
  const bdReaderName = document.getElementById('bd-reader-name');
  const bdSearch = document.getElementById('bd-search');
  const addressList = document.getElementById('address-list');
  const bdFilterMissedBtn = document.getElementById('bd-filter-missed');
  const bdFilterSkippedBtn = document.getElementById('bd-filter-skipped');
  document.getElementById('bundle-back-btn').addEventListener('click', goHome);
  document.getElementById('bd-back-nav').addEventListener('click', goHome);
  document.getElementById('bd-nav-map').addEventListener('click', () => {
    viewBundle.classList.add('hidden');
    showMapView();
  });
  document.getElementById('totals-back-btn').addEventListener('click', () => {
    viewTotals.classList.add('hidden');
    viewBundle.classList.remove('hidden');
  });
  document.getElementById('totals-print-btn').addEventListener('click', () => window.print());

  // Card detail
  const viewCard = document.getElementById('view-card');
  const cardDtDate = document.getElementById('card-dt-date');
  const cardDtTime = document.getElementById('card-dt-time');
  const cardDtSeq = document.getElementById('card-dt-seq');
  const cardDtAddress = document.getElementById('card-dt-address');
  const cardDtCity = document.getElementById('card-dt-city');
  const cardDtInstrumentRow = document.getElementById('card-dt-instrument-row');
  const cardDtInstrument = document.getElementById('card-dt-instrument');
  const cardDtLocRow = document.getElementById('card-dt-loc-row');
  const cardDtLoc = document.getElementById('card-dt-loc');
  const cardDtMtrSize = document.getElementById('card-dt-mtr-size');
  const cardDtSerial = document.getElementById('card-dt-serial');
  const cardDtEst = document.getElementById('card-dt-est');
  const cardDtSpecWrap = document.getElementById('card-dt-spec-wrap');
  const cardDtSpec = document.getElementById('card-dt-spec');
  const cardDtReading = document.getElementById('card-dt-reading');
  const cardDtSkip = document.getElementById('card-dt-skip');
  const cardDtSkipOther = document.getElementById('card-dt-skip-other');
  const cardDtReadDate = document.getElementById('card-dt-read-date');
  const cardDtComments = document.getElementById('card-dt-comments');
  const cardDtSaveBtn = document.getElementById('card-dt-save-btn');
  const cardNavPos = document.getElementById('card-nav-pos');
  const cardNavComplete = document.getElementById('card-nav-complete');
  const cardPrevBtn = document.getElementById('card-prev-btn');
  const cardNextBtn = document.getElementById('card-next-btn');
  const cardMenuBtn = document.getElementById('card-menu-btn');
  const cardMenuDropdown = document.getElementById('card-menu-dropdown');
  const cardMenuReverseCheck = document.getElementById('card-menu-reverse-check');
  document.getElementById('card-back-btn').addEventListener('click', cardGoBack);
  document.getElementById('card-back-nav').addEventListener('click', () => {
    viewCard.classList.add('hidden');
    viewBundle.classList.add('hidden');
    showMapView();
  });

  document.getElementById('card-bundles-nav').addEventListener('click', () => {
    viewCard.classList.add('hidden');
    viewBundle.classList.remove('hidden');
  });

  // ─── Card burger menu ─────────────────────────────
  function openCardMenu() {
    cardMenuDropdown.classList.remove('hidden');
    cardMenuBtn.classList.add('menu-open');
  }

  function closeCardMenu() {
    cardMenuDropdown.classList.add('hidden');
    cardMenuBtn.classList.remove('menu-open');
  }

  function setReadingLocked(locked) {
    cardDtReading.readOnly = locked;
    cardDtReading.classList.toggle('reading-locked', locked);
  }

  // Use pointerdown on the button so stopPropagation prevents the document
  // handler from firing on the same interaction (avoids open→close→open race).
  cardMenuBtn.addEventListener('pointerdown', (e) => {
    e.stopPropagation();
    cardMenuDropdown.classList.contains('hidden') ? openCardMenu() : closeCardMenu();
  });

  document.addEventListener('pointerdown', (e) => {
    if (!cardMenuDropdown.classList.contains('hidden') &&
      !cardMenuDropdown.contains(e.target)) {
      closeCardMenu();
    }
  });

  document.getElementById('card-menu-reverse').addEventListener('click', () => {
    reverseDirection = !reverseDirection;
    cardMenuReverseCheck.classList.toggle('hidden', !reverseDirection);
    document.getElementById('card-menu-reverse').classList.toggle('active', reverseDirection);
    closeCardMenu();
    showToast(reverseDirection ? 'Direction reversed — saving goes backward' : 'Direction restored — saving goes forward');
  });

  document.getElementById('card-menu-search').addEventListener('click', () => {
    closeCardMenu();
    viewCard.classList.add('hidden');
    viewBundle.classList.remove('hidden');
    setTimeout(() => { bdSearch.focus(); bdSearch.select(); }, 80);
  });

  document.getElementById('card-menu-list').addEventListener('click', () => {
    closeCardMenu();
    viewCard.classList.add('hidden');
    viewBundle.classList.remove('hidden');
  });

  document.getElementById('card-menu-report').addEventListener('click', () => {
    closeCardMenu();
    viewCard.classList.add('hidden');
    showTotalsView(currentBundle);
  });

  document.getElementById('card-menu-delete-reading').addEventListener('click', () => {
    closeCardMenu();
    document.getElementById('del-reading-modal').classList.remove('hidden');
  });

  document.getElementById('del-reading-cancel').addEventListener('click', () => {
    document.getElementById('del-reading-modal').classList.add('hidden');
  });

  document.getElementById('del-reading-confirm').addEventListener('click', () => {
    document.getElementById('del-reading-modal').classList.add('hidden');
    currentRow['READING'] = '';
    const hasSkip = (currentRow['SKIP'] || '').trim() !== '' && currentRow['SKIP'] !== 'Other';
    const hasComments = (currentRow['COMMENTS'] || '').trim() !== '';
    if (!hasSkip && !hasComments) {
      currentRow['READ DATE'] = '';
    }
    cardDtReading.value = '';
    setReadingLocked(false);
    document.getElementById('card-menu-delete-reading').classList.add('hidden');
    updateMapMarkerForRow(currentRow);
    refreshBundleStats(currentBundle);
    saveRecordsBackup();
    renderHome();
    showToast('Reading deleted');
  });

  // Map view DOM refs
  const viewMap = document.getElementById('view-map');
  const mapDate = document.getElementById('map-date');
  const mapTime = document.getElementById('map-time');
  const mapGeocodeBar = document.getElementById('map-geocode-bar');
  const mapGeocodeLabel = document.getElementById('map-geocode-label');
  const mapGeocodeFill = document.getElementById('map-geocode-fill');
  const mapReaderName = document.getElementById('map-reader-name');
  const homeNavMap = document.getElementById('home-nav-map');
  const mapNavBundles = document.getElementById('map-nav-bundles');
  const mapNotFoundBar = document.getElementById('map-notfound-bar');
  const mapNotFoundLabel = document.getElementById('map-notfound-label');
  const geocodeFixModal = document.getElementById('geocode-fix-modal');
  const geocodeFixList = document.getElementById('geocode-fix-list');
  const multiUnitModal = document.getElementById('multi-unit-modal');
  const mapBundleFilter = document.getElementById('map-bundle-filter');

  // Rate modal
  const rateModal = document.getElementById('rate-modal');
  const rateEst3In = document.getElementById('rate-est3');
  const rateEst46In = document.getElementById('rate-est46');
  const rateEst7pIn = document.getElementById('rate-est7p');
  function showRateModal(records) {
    pendingRecordsForRates = records;
    // Detect area from first record to pre-fill defaults
    const firstMruRow = records.find(r => (r['MRU id'] || '').trim());
    const mruArea = firstMruRow
      ? String((firstMruRow['MRU id'] || '').trim()).padStart(6, '0').substring(2, 4)
      : '';
    // Prefer previously stored rates for a matching bundle key, then area defaults
    const newKeys = getNewBundleKeys(records);
    const stored = newKeys.map(k => bundleRates[k]).find(Boolean);
    const defaults = stored || AREA_RATES[mruArea] || {};
    rateEst3In.value = defaults.est3 != null ? defaults.est3 : '';
    rateEst46In.value = defaults.est46 != null ? defaults.est46 : '';
    rateEst7pIn.value = defaults.est7p != null ? defaults.est7p : '';
    rateModal.classList.remove('hidden');
    rateEst3In.focus();
  }

  document.getElementById('rate-modal-confirm').addEventListener('click', () => {
    const rates = {
      est3: parseFloat(rateEst3In.value) || 0,
      est46: parseFloat(rateEst46In.value) || 0,
      est7p: parseFloat(rateEst7pIn.value) || 0,
    };
    rateModal.classList.add('hidden');
    commitLoad(pendingRecordsForRates, rates);
    pendingRecordsForRates = null;
  });

  document.getElementById('rate-modal-cancel').addEventListener('click', () => {
    rateModal.classList.add('hidden');
    pendingRecordsForRates = null;
  });

  // Map tab navigation
  homeNavMap.addEventListener('click', showMapView);
  mapNavBundles.addEventListener('click', () => {
    if (leafletMap) leafletMap.stopLocate();
    clearGpsTimer();
    viewMap.classList.add('hidden');
    viewHome.classList.remove('hidden');
  });

  mapBundleFilter.addEventListener('change', applyBundleFilter);

  document.getElementById('map-geocode-btn').addEventListener('click', () => {
    const key = mapBundleFilter.value;
    const cache = loadGeoCache();
    const targets = key ? bundles.filter(b => b.key === key) : bundles;
    targets.forEach(bundle => {
      bundle.rows.forEach(r => {
        const num = (r['#'] || '').trim();
        const street = (r['STREET'] || '').trim();
        const city = (r['City'] || '').trim();
        if (!street || !city) return;
        const { cleanNum, cleanStreet } = stripUnitInfo(num, street);
        const cacheKey = `${cleanNum} ${cleanStreet},${city}`.trim().toLowerCase();
        delete cache[cacheKey];
      });
    });
    saveGeoCache(cache);
    if (key) {
      geocodedPoints = geocodedPoints.filter(p => p.bundle.key !== key);
      geocodeFailures = geocodeFailures.filter(f => f.bundle.key !== key);
    } else {
      geocodedPoints = [];
      geocodeFailures = [];
    }
    geocodeStale = true;
    showMapView();
  });

  document.getElementById('map-download-btn').addEventListener('click', downloadMapTiles);

  // Not-found bar + fix modal
  document.getElementById('map-notfound-btn').addEventListener('click', showGeocodeFix);
  document.getElementById('geocode-fix-close').addEventListener('click', () => {
    geocodeFixModal.classList.add('hidden');
  });

  document.getElementById('multi-unit-close').addEventListener('click', () => {
    multiUnitModal.classList.add('hidden');
  });

  // ── Route Optimization Helpers ───────────────────────────────────────────

  function getRowCoords(row) {
    const cache = loadGeoCache();
    const mapAddress = (row['Map Address'] || '').trim();
    const num = (row['#'] || '').trim();
    const street = (row['STREET'] || '').trim();
    const city = (row['City'] || '').trim();
    const { cleanNum, cleanStreet } = stripUnitInfo(num, street);
    const key = mapAddress
      ? mapAddress.trim().toLowerCase()
      : `${cleanNum} ${cleanStreet},${city}`.trim().toLowerCase();
    return cache[key] || null;
  }

  function haversineDistance(a, b) {
    const R = 6371000;
    const toRad = x => x * Math.PI / 180;
    const dLat = toRad(b.lat - a.lat);
    const dLng = toRad(b.lng - a.lng);
    const sLat = Math.sin(dLat / 2);
    const sLng = Math.sin(dLng / 2);
    const c = sLat * sLat + Math.cos(toRad(a.lat)) * Math.cos(toRad(b.lat)) * sLng * sLng;
    return R * 2 * Math.atan2(Math.sqrt(c), Math.sqrt(1 - c));
  }

  // Fetch NxN road-time matrix from OSRM Table API.
  // Returns 2D array of travel times (seconds), or null on failure.
  async function getRoadDistanceMatrix(coords) {
    if (!navigator.onLine) return null;
    // OSRM demo server caps at ~100 coordinates
    if (coords.length > 100) return null;
    try {
      const coordStr = coords.map(c => `${c.lng},${c.lat}`).join(';');
      const resp = await fetch(
        `https://router.project-osrm.org/table/v1/driving/${coordStr}?annotations=duration`,
        { headers: { 'User-Agent': 'MeterReaderPWA/1.0' }, signal: AbortSignal.timeout(10000) }
      );
      if (!resp.ok) return null;
      const data = await resp.json();
      if (data.code !== 'Ok' || !data.durations) return null;
      return data.durations;  // matrix[i][j] = seconds from i to j
    } catch {
      return null;
    }
  }

  async function computeOptimizedOrder(rows, startRow) {
    const points = rows.map((r, i) => ({ i, coords: getRowCoords(r) }));
    const withCoords = points.filter(p => p.coords);
    const withoutCoords = points.filter(p => !p.coords);

    if (withCoords.length === 0) return { ordered: rows, usedRoads: false };

    const startOrigIdx = rows.indexOf(startRow);
    let start = withCoords.findIndex(p => p.i === startOrigIdx);
    if (start < 0) start = 0;

    // Try road distances first, fall back to straight-line
    const matrix = await getRoadDistanceMatrix(withCoords.map(p => p.coords));
    const usedRoads = matrix !== null;

    function dist(a, b) {
      if (matrix) {
        const t = matrix[a]?.[b];
        // null/undefined values in OSRM matrix mean unreachable — fall back to haversine
        return (t != null) ? t : haversineDistance(withCoords[a].coords, withCoords[b].coords);
      }
      return haversineDistance(withCoords[a].coords, withCoords[b].coords);
    }

    // Nearest-neighbor
    const visited = new Array(withCoords.length).fill(false);
    const order = [start];
    visited[start] = true;
    for (let step = 1; step < withCoords.length; step++) {
      const last = order[order.length - 1];
      let best = -1, bestDist = Infinity;
      for (let j = 0; j < withCoords.length; j++) {
        if (visited[j]) continue;
        const d = dist(last, j);
        if (d < bestDist) { bestDist = d; best = j; }
      }
      order.push(best);
      visited[best] = true;
    }

    // 2-opt improvement
    let improved = true;
    while (improved) {
      improved = false;
      for (let i = 1; i < order.length - 1; i++) {
        for (let j = i + 1; j < order.length; j++) {
          const a = order[i - 1], b = order[i], c = order[j];
          const d = j + 1 < order.length ? order[j + 1] : -1;
          const distBefore = dist(a, b) + (d >= 0 ? dist(c, d) : 0);
          const distAfter = dist(a, c) + (d >= 0 ? dist(b, d) : 0);
          if (distAfter < distBefore - 0.1) {
            order.splice(i, j - i + 1, ...order.slice(i, j + 1).reverse());
            improved = true;
          }
        }
      }
    }

    const optimized = order.map(idx => rows[withCoords[idx].i]);
    const ungeocoded = withoutCoords.map(p => rows[p.i]);
    return { ordered: [...optimized, ...ungeocoded], usedRoads };
  }

  function enterPickStartMode() {
    pickStartMode = true;
    document.getElementById('bd-optimize-banner').classList.remove('hidden');
    document.getElementById('address-list').classList.add('pick-start-mode');
  }

  function exitPickStartMode() {
    pickStartMode = false;
    document.getElementById('bd-optimize-banner').classList.add('hidden');
    document.getElementById('address-list').classList.remove('pick-start-mode');
  }

  async function applyRouteOptimization(bundle, startRow) {
    exitPickStartMode();
    const rows = bundle.rows;
    if (!routeOptimized) {
      originalRowsOrder = rows.slice();
    }

    // Show loading state while OSRM is queried
    const optimizeBtn = document.getElementById('bd-optimize-btn');
    optimizeBtn.textContent = 'Optimizing...';
    optimizeBtn.disabled = true;

    const { ordered, usedRoads } = await computeOptimizedOrder(rows, startRow);
    ordered.forEach((r, i) => { r['Seq #'] = i + 1; });
    bundle.rows = ordered;
    routeOptimized = true;

    optimizeBtn.textContent = 'Optimize Route';
    optimizeBtn.disabled = false;

    document.getElementById('bd-reset-order-btn').classList.remove('hidden');
    showBundleDetail(bundle, true);
    showToast(usedRoads ? 'Route optimized using road distances.' : 'Offline — route optimized using straight-line distances.');
  }

  function resetRouteOrder() {
    if (!originalRowsOrder) return;
    // Restore original Seq # values
    originalRowsOrder.forEach((r, i) => { r['Seq #'] = i + 1; });
    currentBundle.rows = originalRowsOrder.slice();
    originalRowsOrder = null;
    routeOptimized = false;
    document.getElementById('bd-reset-order-btn').classList.add('hidden');
    showBundleDetail(currentBundle, true);
  }

  // ─────────────────────────────────────────────────────────────────────────

  // Returns expected digit count for a given meter size, or null if unrestricted.
  function getExpectedDigits(mtrSize) {
    const s = (mtrSize || '').trim().toUpperCase();
    if (s === 'UD') return 6;
    if (s === 'UC' || s === 'UB') return 5;
    if (s.startsWith('S')) return 5;
    if (s.startsWith('C') || s.startsWith('D') || s.startsWith('G')) return 4;
    return null;
  }

  function padReading(val, mtrSize) {
    const digits = getExpectedDigits(mtrSize);
    return (digits && val.length > 0 && val.length < digits)
      ? val.padStart(digits, '0')
      : val;
  }

  cardDtReading.addEventListener('input', () => {
    const cleaned = cardDtReading.value.replace(/\D/g, '');
    if (cleaned !== cardDtReading.value) cardDtReading.value = cleaned;
  });

  cardDtReading.addEventListener('blur', () => {
    const val = cardDtReading.value.trim();
    if (!val || !currentRow) return;
    cardDtReading.value = padReading(val, currentRow['MTR SIZE']);
  });

  cardDtSkip.addEventListener('change', () => {
    cardDtSkipOther.classList.toggle('hidden', cardDtSkip.value !== 'Other');
    if (cardDtSkip.value === 'Other') setTimeout(() => cardDtSkipOther.focus(), 30);
  });

  // Reader name modal
  const readerModal = document.getElementById('reader-name-modal');
  const readerInput = document.getElementById('reader-name-input');
  const readerConfirm = document.getElementById('reader-name-confirm');
  const readerCancel = document.getElementById('reader-name-cancel');

  // ─── Clock / Date ─────────────────────────────────
  function updateClock() {
    const now = new Date();
    const t = now.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: true });
    const d = now.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
    homeTime.textContent = t;
    homeDate.textContent = d;
    bdTime.textContent = t;
    bdDate.textContent = d;
    cardDtTime.textContent = t;
    cardDtDate.textContent = d;
    mapTime.textContent = t;
    mapDate.textContent = d;
  }
  updateClock();
  setInterval(updateClock, 10000);

  // ─── Persistence ──────────────────────────────────
  const SENT_KEY = 'mtr_sent_bundles';
  const READER_KEY = 'mtr_reader_name';
  const RECORDS_KEY = 'mtr_records_backup';
  const DELETED_KEY = 'mtr_deleted_bundles';
  const GEOCACHE_KEY = 'mtr_geocache';
  const RATES_KEY = 'mtr_bundle_rates';
  const BACKUP_DATE_KEY = 'mtr_last_backup_date';
  const TWO_WEEKS = 14 * 24 * 60 * 60 * 1000;

  // ─── Auto-backup (File System Access API) ─────────
  const IDB_NAME = 'mtr_fs';
  const IDB_VERSION = 1;
  const IDB_STORE = 'handles';
  const HANDLE_KEY = 'backupDir';
  const AUTO_BACKUP_INDICATOR_KEY = 'mtr_auto_backup_active';

  let deletedBundles = [];  // [{ key, bundleName, rows, deletedAt }]

  function loadSentState() {
    try {
      const raw = localStorage.getItem(SENT_KEY);
      if (raw) sentBundles = new Set(JSON.parse(raw));
    } catch (_) { }
  }

  function saveSentState() {
    try {
      localStorage.setItem(SENT_KEY, JSON.stringify([...sentBundles]));
    } catch (_) { }
  }

  function saveDeletedBundles() {
    try {
      localStorage.setItem(DELETED_KEY, JSON.stringify(deletedBundles));
    } catch (_) { }
  }

  function loadDeletedBundles() {
    try {
      const raw = localStorage.getItem(DELETED_KEY);
      if (!raw) return;
      const all = JSON.parse(raw);
      const cutoff = Date.now() - TWO_WEEKS;
      deletedBundles = all.filter(d => d.deletedAt >= cutoff);
      // Prune expired entries from storage
      if (deletedBundles.length !== all.length) saveDeletedBundles();
    } catch (_) { }
  }

  function saveRecordsBackup() {
    try {
      localStorage.setItem(RECORDS_KEY, JSON.stringify(allRecords));
    } catch (_) { }
    const csv = buildAllRecordsCSV();
    saveToOPFS(csv);
    saveCSVToDisk(csv);
    scheduleAutoSave();
  }

  function saveToOPFS(csv) {
    if (!navigator.storage || !navigator.storage.getDirectory) return;
    navigator.storage.getDirectory().then(function (dir) {
      return dir.getFileHandle('meter-backup.csv', { create: true });
    }).then(function (fh) {
      return fh.createWritable();
    }).then(function (writable) {
      return writable.write(csv).then(function () { return writable.close(); });
    }).catch(function () { /* silent — data already safe in localStorage */ });
  }

  function loadRecordsBackup() {
    try {
      const raw = localStorage.getItem(RECORDS_KEY);
      if (!raw) return false;
      const saved = JSON.parse(raw);
      if (!Array.isArray(saved) || !saved.length) return false;
      allRecords = saved;
      bundles = groupIntoBundles(allRecords);
      return true;
    } catch (_) { return false; }
  }

  // ─── IndexedDB helpers (for FileSystemDirectoryHandle storage) ────
  function openHandleDB() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(IDB_NAME, IDB_VERSION);
      req.onupgradeneeded = (e) => { e.target.result.createObjectStore(IDB_STORE); };
      req.onsuccess = (e) => resolve(e.target.result);
      req.onerror = (e) => reject(e.target.error);
    });
  }

  async function saveDirectoryHandle(handle) {
    try {
      const db = await openHandleDB();
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).put(handle, HANDLE_KEY);
      return new Promise((resolve, reject) => {
        tx.oncomplete = () => { db.close(); resolve(); };
        tx.onerror = (e) => { db.close(); reject(e.target.error); };
      });
    } catch (_) { }
  }

  async function loadDirectoryHandle() {
    try {
      const db = await openHandleDB();
      const tx = db.transaction(IDB_STORE, 'readonly');
      const req = tx.objectStore(IDB_STORE).get(HANDLE_KEY);
      return new Promise((resolve, reject) => {
        req.onsuccess = () => { db.close(); resolve(req.result || null); };
        req.onerror = (e) => { db.close(); reject(e.target.error); };
      });
    } catch (_) { return null; }
  }

  async function clearDirectoryHandle() {
    try {
      const db = await openHandleDB();
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).delete(HANDLE_KEY);
      return new Promise((resolve) => {
        tx.oncomplete = () => { db.close(); resolve(); };
        tx.onerror = () => { db.close(); resolve(); };
      });
    } catch (_) { }
  }

  function saveCSVToDisk(csv) {
    // Background saves always use OPFS — no dialogs, works on all platforms.
    saveToOPFS(csv);
  }

  // ─── Geocoding ────────────────────────────────────
  function loadGeoCache() {
    try { return JSON.parse(localStorage.getItem(GEOCACHE_KEY) || '{}'); } catch (_) { return {}; }
  }

  function saveGeoCache(cache) {
    try { localStorage.setItem(GEOCACHE_KEY, JSON.stringify(cache)); } catch (_) { }
  }

  function loadBundleRates() {
    try { return JSON.parse(localStorage.getItem(RATES_KEY) || '{}'); } catch (_) { return {}; }
  }

  function saveBundleRates() {
    try { localStorage.setItem(RATES_KEY, JSON.stringify(bundleRates)); } catch (_) { }
  }

  // Strips suite/unit/lot qualifiers from a civic number and street name
  // so geocoders can find the base address without being confused by unit info.
  function stripUnitInfo(num, street) {
    // Remove unit/suite/lot/apt suffixes from the house number.
    // Handles: "123-4", "123 UNIT 4", "123 APT B", "123 SUITE 200", "LOT 5"
    const cleanNum = num
      .replace(/\s*(unit|apt|apartment|suite|ste|lot|rm|room)\s*\S*/gi, '')
      .replace(/-\w+$/, '')   // strip trailing dash-unit e.g. "123-4A"
      .trim();

    // Remove unit/suite/lot qualifiers from the street name.
    // Handles: "MAIN ST SUITE 200", "OAK AVE, APT 3", "ELM RD LOT B"
    // Also strips floor/level descriptors: "FLOOR 1-6", "FLOOR MAIN", "MAIN FLR", "BASEMENT"
    const cleanStreet = street
      .replace(/[,\s]*(unit|apt|apartment|suite|ste|lot|rm|room)\s*\S*/gi, '')
      .replace(/[,\s]*(floor\s+(?:[1-6]|main)|main\s+flr|basement)\b/gi, '')
      .trim();

    return { cleanNum, cleanStreet };
  }

  // Returns Promise<{coords, count}>.
  // coords = {lat,lng} or null. count = number of real HTTP requests made.
  // Uses mapAddress directly if provided; otherwise builds from num/street/city with fallback.
  function geocodeAddress(num, street, city, cache, mapAddress) {
    const { cleanNum, cleanStreet } = stripUnitInfo(num, street);
    const key = mapAddress
      ? mapAddress.trim().toLowerCase()
      : `${cleanNum} ${cleanStreet},${city}`.trim().toLowerCase();
    if (cache[key]) return Promise.resolve({ coords: cache[key], count: 0 });

    const doFetch = (q) =>
      fetch(`https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(q)}&format=json&limit=1&countrycodes=ca`,
        { headers: { 'User-Agent': 'MeterReaderPWA/1.0' } })
        .then(r => r.json())
        .then(d => (d && d[0]) ? { lat: parseFloat(d[0].lat), lng: parseFloat(d[0].lon) } : null)
        .catch(() => null);

    if (mapAddress) {
      const stripped = mapAddress.replace(/\b[A-Za-z]\d[A-Za-z]\s?\d[A-Za-z]\d\b/, '').replace(/,?\s*$/, '').trim();
      return doFetch(stripped || mapAddress).then(coords => {
        if (coords) { cache[key] = coords; saveGeoCache(cache); }
        return { coords, count: 1 };
      });
    }

    return doFetch(`${cleanNum} ${cleanStreet}, ${city}, Ontario`).then(coords => {
      if (coords) {
        cache[key] = coords;
        saveGeoCache(cache);
        return { coords, count: 1 };
      }
      // Fallback: street + city only (no house number)
      return doFetch(`${cleanStreet}, ${city}, Ontario`).then(coords2 => {
        if (coords2) { cache[key] = coords2; saveGeoCache(cache); }
        return { coords: coords2, count: 2 };
      });
    });
  }

  // Rate-limited geocode of all records (1 req/sec per Nominatim policy).
  // Calls progressCb(done, total) after each attempt.
  function geocodeAllRecords(progressCb) {
    const cache = loadGeoCache();
    const tasks = [];
    bundles.forEach(bundle => {
      bundle.rows.forEach(row => {
        const num = (row['#'] || '').trim();
        const street = (row['STREET'] || '').trim();
        const city = (row['City'] || '').trim();
        const mapAddress = (row['Map Address'] || '').trim();
        if (!mapAddress && (!street || !city)) return;
        tasks.push({ row, bundle, num, street, city, mapAddress });
      });
    });
    const results = [];
    const failures = [];
    let doneCount = 0;
    const total = tasks.length;

    function processNext(i) {
      if (i >= total) return Promise.resolve({ points: results, failures });
      const { row, bundle, num, street, city, mapAddress } = tasks[i];
      return geocodeAddress(num, street, city, cache, mapAddress).then(({ coords, count }) => {
        doneCount++;
        progressCb(doneCount, total);
        if (coords) results.push({ lat: coords.lat, lng: coords.lng, row, bundle });
        else failures.push({ row, bundle, num, street, city, mapAddress });
        // Wait 1050ms per real HTTP request made (0 for cache hits)
        const delay = count * 1050;
        return new Promise(res => setTimeout(res, delay)).then(() => processNext(i + 1));
      });
    }
    return processNext(0);
  }

  // ─── Map View ─────────────────────────────────────
  function getMarkerColor(row) {
    const reading = (row['READING'] || '').trim();
    const skip = (row['SKIP'] || '').trim();
    if (skip) return '#f59e0b';  // amber — skipped
    if (reading) return '#22c55e';  // green — read
    return '#3b82f6';               // blue — unread
  }

  function makeCircleIcon(color, count) {
    if (count && count > 1) {
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 28 28">` +
        `<circle cx="14" cy="14" r="12" fill="${color}" stroke="#fff" stroke-width="2.5"/>` +
        `<text x="14" y="19" text-anchor="middle" font-size="13" font-weight="bold" fill="#fff" font-family="sans-serif">${count}</text>` +
        `</svg>`;
      return L.divIcon({ html: svg, className: '', iconSize: [28, 28], iconAnchor: [14, 14], popupAnchor: [0, -14] });
    }
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20">` +
      `<circle cx="10" cy="10" r="8" fill="${color}" stroke="#fff" stroke-width="2"/></svg>`;
    return L.divIcon({ html: svg, className: '', iconSize: [20, 20], iconAnchor: [10, 10], popupAnchor: [0, -12] });
  }

  function getGroupColor(rows) {
    if (rows.some(r => getMarkerColor(r) === '#3b82f6')) return '#3b82f6';
    if (rows.some(r => getMarkerColor(r) === '#f59e0b')) return '#f59e0b';
    return '#22c55e';
  }

  // ─── Offline tile download ────────────────────────
  function latLngToTileXY(lat, lng, zoom) {
    const n = 1 << zoom;
    const x = Math.floor((lng + 180) / 360 * n);
    const latRad = lat * Math.PI / 180;
    const y = Math.floor((1 - Math.log(Math.tan(latRad) + 1 / Math.cos(latRad)) / Math.PI) / 2 * n);
    return { x: Math.max(0, Math.min(n - 1, x)), y: Math.max(0, Math.min(n - 1, y)) };
  }

  async function downloadMapTiles() {
    if (!leafletMap) return;
    const label = document.getElementById('map-download-label');
    const btn = document.getElementById('map-download-btn');

    const bounds = leafletMap.getBounds();
    const latPad = (bounds.getNorth() - bounds.getSouth()) * 0.15;
    const lngPad = (bounds.getEast() - bounds.getWest()) * 0.15;
    const N = bounds.getNorth() + latPad, S = bounds.getSouth() - latPad;
    const E = bounds.getEast() + lngPad, W = bounds.getWest() - lngPad;

    const SUB = ['a', 'b', 'c'];
    const urls = [];
    for (let z = 12; z <= 18; z++) {
      const min = latLngToTileXY(N, W, z);
      const max = latLngToTileXY(S, E, z);
      for (let x = min.x; x <= max.x; x++) {
        for (let y = min.y; y <= max.y; y++) {
          urls.push(`https://${SUB[(x + y) % 3]}.tile.openstreetmap.org/${z}/${x}/${y}.png`);
        }
      }
    }

    if (urls.length > 3000) {
      showToast(`Area too large (${urls.length} tiles). Zoom in closer first.`, true);
      return;
    }

    btn.disabled = true;
    let done = 0;
    const total = urls.length;
    const BATCH = 8;

    for (let i = 0; i < urls.length; i += BATCH) {
      await Promise.all(
        urls.slice(i, i + BATCH).map(u => fetch(u, { mode: 'cors' }).catch(() => { }))
      );
      done = Math.min(i + BATCH, total);
      label.textContent = `${done} / ${total}`;
    }

    btn.disabled = false;
    label.textContent = '✓ Map Saved';
    showToast(`${total} tiles cached — map works offline`);
    setTimeout(() => { label.textContent = 'Download Map'; }, 4000);
  }

  function initLeafletMap() {
    if (mapInitialized) return;
    if (typeof L === 'undefined') {
      showToast('Map failed to load — check your internet connection and reload the page.', true);
      return;
    }
    leafletMap = L.map('map-container', { zoomControl: true });
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
      maxZoom: 19
    }).addTo(leafletMap);

    // Custom "Locate Me" control — top-right, below zoom buttons
    const LocateControl = L.Control.extend({
      options: { position: 'topright' },
      onAdd() {
        const btn = L.DomUtil.create('button', 'locate-btn leaflet-bar');
        btn.title = 'Show my location';
        btn.setAttribute('aria-label', 'Show my location');
        btn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18"
            viewBox="0 0 24 24" fill="none" stroke="currentColor"
            stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
          <circle cx="12" cy="12" r="3"/>
          <path d="M12 2v3M12 19v3M2 12h3M19 12h3"/>
        </svg>`;
        L.DomEvent.on(btn, 'click', L.DomEvent.stopPropagation);
        L.DomEvent.on(btn, 'click', L.DomEvent.preventDefault);
        L.DomEvent.on(btn, 'click', startLocating);
        return btn;
      }
    });
    new LocateControl().addTo(leafletMap);
    leafletMap.on('locationfound', onLocationFound);
    leafletMap.on('locationerror', onLocationError);
    leafletMap.on('move', resetGpsTimer);
    leafletMap.on('zoom', resetGpsTimer);
    leafletMap.on('click', resetGpsTimer);

    mapInitialized = true;
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

  function clearGpsTimer() {
    clearTimeout(gpsAutoStopTimer);
    gpsAutoStopTimer = null;
    gpsWatching = false;
  }

  function startLocating() {
    if (!mapInitialized || !leafletMap) return;
    if (!('geolocation' in navigator)) {
      showToast('Geolocation is not supported by this browser.', true);
      return;
    }
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
            fill="#f97316" stroke="#fff" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"
            class="user-location-person">
          <circle cx="12" cy="7" r="4"/>
          <path d="M4 21v-1a8 8 0 0 1 16 0v1"/>
        </svg>`,
        className: '',
        iconSize: [28, 28],
        iconAnchor: [14, 28],
        popupAnchor: [0, -30]
      });
      userLocationMarker = L.marker(e.latlng, { icon, zIndexOffset: 1000 })
        .addTo(leafletMap)
        .bindPopup('<strong>Your Location</strong>');
      leafletMap.setView(e.latlng, 17);
    }
  }

  function onLocationError(e) {
    const msg = e.code === 1
      ? 'Location access denied — check your browser permissions.'
      : 'Location unavailable — GPS signal lost or timed out.';
    showToast(msg, true);
  }

  function clearMapMarkers() {
    mapMarkers.forEach(({ marker }) => marker.remove());
    mapMarkers = [];
  }

  function updateMapMarkerForRow(row) {
    const entry = mapMarkers.find(m => m.row === row);
    if (!entry) return;
    const sharedRows = mapMarkers.filter(m => m.marker === entry.marker).map(m => m.row);
    entry.marker.setIcon(makeCircleIcon(getGroupColor(sharedRows), sharedRows.length > 1 ? sharedRows.length : undefined));
  }

  // Single tap → briefly show popup. Double tap → open data entry (or unit picker for stacked markers).
  function attachMarkerTapHandlers(marker, rows, bundles, lat, lng) {
    let tapTimer = null;
    marker.on('click', () => {
      if (tapTimer) {
        clearTimeout(tapTimer);
        tapTimer = null;
        marker.closePopup();
        if (rows.length === 1) {
          viewMap.classList.add('hidden');
          showBundleDetail(bundles[0]);
          showCardDetail(rows[0], bundles[0], 'map');
        } else {
          showMultiUnitPicker(rows, bundles);
        }
      } else {
        pendingFixPoint = { row: rows[0], bundle: bundles[0], marker, lat, lng };
        marker.openPopup();
        tapTimer = setTimeout(() => { tapTimer = null; }, 300);
        setTimeout(() => marker.closePopup(), 2500);
      }
    });
    // Prevent double-tap from zooming the map
    marker.on('dblclick', (e) => { L.DomEvent.stopPropagation(e); });
  }

  function placeMarkers(points) {
    if (!mapInitialized) return;
    clearMapMarkers();
    const bounds = [];

    // Group overlapping points by lat/lng so stacked addresses get one marker
    const groups = new Map();
    points.forEach(pt => {
      const key = `${pt.lat.toFixed(6)},${pt.lng.toFixed(6)}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(pt);
    });

    groups.forEach(pts => {
      const { lat, lng } = pts[0];
      const rows = pts.map(p => p.row);
      const bundles = pts.map(p => p.bundle);
      const count = pts.length;
      const color = getGroupColor(rows);
      const icon = makeCircleIcon(color, count > 1 ? count : undefined);
      const mapsUrl = `https://www.google.com/maps/dir/?api=1&destination=${lat},${lng}`;
      const addr = [rows[0]['#'], rows[0]['STREET']].filter(Boolean).join(' ');
      const loc = (rows[0]['LOC'] || '').trim();
      let popup;
      if (count === 1) {
        popup = `<strong>${addr}</strong>${loc ? `<br><span style="font-size:0.85em;opacity:0.8">${loc}</span>` : ''}<br>${bundles[0].bundleName || ''}<br><a href="${mapsUrl}" target="_blank" rel="noopener" class="popup-nav-link">&#9654; Navigate</a><br><a href="#" class="popup-nav-link" onclick="window._openFixLocation();return false;">&#9999; Fix Location</a>`;
      } else {
        popup = `<strong>${addr}</strong> <span style="background:#3b82f6;color:#fff;border-radius:4px;padding:1px 6px;font-size:0.78em;font-weight:700;">${count} units</span><br><span style="font-size:0.82em;opacity:0.7">Double-tap to select a unit</span><br><a href="${mapsUrl}" target="_blank" rel="noopener" class="popup-nav-link">&#9654; Navigate</a><br><a href="#" class="popup-nav-link" onclick="window._openFixLocation();return false;">&#9999; Fix Location</a>`;
      }
      const marker = L.marker([lat, lng], { icon })
        .addTo(leafletMap)
        .bindPopup(popup, { autoClose: false, closeOnClick: false });
      attachMarkerTapHandlers(marker, rows, bundles, lat, lng);
      rows.forEach((row, i) => mapMarkers.push({ marker, row, bundle: bundles[i] }));
      bounds.push([lat, lng]);
    });
    populateBundleFilter();
    applyBundleFilter();
  }

  function showMapView() {
    if (!allRecords.length) {
      showToast('Load a CSV file first to use the map.', true);
      return;
    }
    viewHome.classList.add('hidden');
    viewMap.classList.remove('hidden');
    mapReaderName.textContent = readerName || 'Field Reader';

    initLeafletMap();
    if (!mapInitialized) return;  // Leaflet failed to load — bail out

    // Double-rAF: ensures Leaflet measures container after layout is complete
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        if (leafletMap) leafletMap.invalidateSize();
      });
    });

    // Already geocoded and not stale — just refresh marker colors and return
    if ((geocodedPoints.length || geocodeFailures.length) && !geocodeStale) {
      placeMarkers(geocodedPoints);
      updateNotFoundBar();
      populateBundleFilter();
      return;
    }

    if (geocodePending) return;  // geocode in progress — wait for it

    geocodeStale = false;
    geocodePending = true;
    mapGeocodeBar.classList.remove('hidden');
    mapGeocodeFill.style.width = '0%';
    mapGeocodeLabel.textContent = 'Geocoding addresses…';

    geocodeAllRecords((done, total) => {
      const pct = total > 0 ? Math.round((done / total) * 100) : 100;
      mapGeocodeFill.style.width = `${pct}%`;
      mapGeocodeLabel.textContent = `Geocoding… ${done} / ${total}`;
    }).then(({ points, failures }) => {
      geocodePending = false;
      geocodedPoints = points;
      geocodeFailures = failures;
      mapGeocodeBar.classList.add('hidden');
      placeMarkers(points);
      updateNotFoundBar();
      if (!points.length) showToast('No addresses could be geocoded.', true);
      else showToast(`${points.length} mapped${failures.length ? `, ${failures.length} not found` : ''}`);
    }).catch(() => {
      geocodePending = false;
      mapGeocodeBar.classList.add('hidden');
      showToast('Geocoding failed — check your connection and try again.', true);
    });
  }

  function updateNotFoundBar() {
    if (geocodeFailures.length) {
      mapNotFoundBar.classList.remove('hidden');
      mapNotFoundLabel.textContent =
        `${geocodeFailures.length} address${geocodeFailures.length !== 1 ? 'es' : ''} not found`;
    } else {
      mapNotFoundBar.classList.add('hidden');
    }
  }

  // ─── Bundle Filter & Force Geocode ───────────────
  function populateBundleFilter() {
    const prev = mapBundleFilter.value;
    mapBundleFilter.innerHTML = '<option value="">All Bundles</option>';
    bundles.forEach(b => {
      const opt = document.createElement('option');
      opt.value = b.key;
      opt.textContent = b.bundleName;
      mapBundleFilter.appendChild(opt);
    });
    if (prev && bundles.some(b => b.key === prev)) mapBundleFilter.value = prev;
  }

  function applyBundleFilter() {
    const key = mapBundleFilter.value;
    const visibleLatLngs = [];
    mapMarkers.forEach(({ marker, bundle }) => {
      if (!key || bundle.key === key) {
        if (!leafletMap.hasLayer(marker)) marker.addTo(leafletMap);
        visibleLatLngs.push(marker.getLatLng());
      } else {
        if (leafletMap.hasLayer(marker)) marker.remove();
      }
    });
    if (visibleLatLngs.length) {
      leafletMap.fitBounds(visibleLatLngs.map(c => [c.lat, c.lng]), { padding: [32, 32] });
    }
  }


  function addSingleMarker(coords, row, bundle) {
    const icon = makeCircleIcon(getMarkerColor(row));
    const addr = [row['#'], row['STREET']].filter(Boolean).join(' ');
    const loc = (row['LOC'] || '').trim();
    const mapsUrl = `https://www.google.com/maps/dir/?api=1&destination=${coords.lat},${coords.lng}`;
    const popup = `<strong>${addr}</strong>${loc ? `<br><span style="font-size:0.85em;opacity:0.8">${loc}</span>` : ''}<br>${bundle.bundleName || ''}<br><a href="${mapsUrl}" target="_blank" rel="noopener" class="popup-nav-link">&#9654; Navigate</a><br><a href="#" class="popup-nav-link" onclick="window._openFixLocation();return false;">&#9999; Fix Location</a>`;
    const marker = L.marker([coords.lat, coords.lng], { icon })
      .addTo(leafletMap)
      .bindPopup(popup, { autoClose: false, closeOnClick: false });
    attachMarkerTapHandlers(marker, [row], [bundle], coords.lat, coords.lng);
    mapMarkers.push({ marker, row, bundle });
  }

  function showMultiUnitPicker(rows, bundles) {
    const list = document.getElementById('multi-unit-list');
    list.innerHTML = '';
    rows.forEach((row, i) => {
      const bundle = bundles[i];
      const addr = [row['#'], row['STREET']].filter(Boolean).join(' ');
      const loc = (row['LOC'] || '').trim();
      const color = getMarkerColor(row);
      const statusLabel = color === '#22c55e' ? 'Read' : color === '#f59e0b' ? 'Skipped' : 'Unread';
      const item = document.createElement('div');
      item.className = 'recover-item';
      item.innerHTML =
        `<div class="recover-item-info">` +
        `<div class="recover-item-name">${addr}${loc ? ` \u2014 ${loc}` : ''}</div>` +
        `<div class="recover-item-meta">${bundle.bundleName || ''} \u00b7 ${statusLabel}</div>` +
        `</div>` +
        `<button class="recover-item-btn" style="background:${color}">Open</button>`;
      item.querySelector('button').addEventListener('click', () => {
        multiUnitModal.classList.add('hidden');
        viewMap.classList.add('hidden');
        showBundleDetail(bundle);
        showCardDetail(row, bundle, 'map');
      });
      list.appendChild(item);
    });
    multiUnitModal.classList.remove('hidden');
  }

  function showGeocodeFix() {
    geocodeFixList.innerHTML = '';
    geocodeFailures.forEach(f => {
      const defaultQ = f.mapAddress || `${f.num ? f.num + ' ' : ''}${f.street}, ${f.city}, Ontario`;
      const item = document.createElement('div');
      item.className = 'geocode-fix-item';
      item.innerHTML = `
        <div class="geocode-fix-addr">${esc(f.mapAddress || (f.num ? f.num + ' ' : '') + f.street + ', ' + f.city)}</div>
        <div class="geocode-fix-row">
          <input type="text" class="geocode-fix-input" value="${esc(defaultQ)}">
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
        status.className = 'geocode-fix-status';

        const url = `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(q)}&format=json&limit=1&countrycodes=ca`;
        fetch(url, { headers: { 'User-Agent': 'MeterReaderPWA/1.0' } })
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
            // Cache under the original key
            const cache = loadGeoCache();
            const key = f.mapAddress ? f.mapAddress.trim().toLowerCase() : `${f.num} ${f.street},${f.city}`.trim().toLowerCase();
            cache[key] = coords;
            saveGeoCache(cache);
            // Add to geocodedPoints + place marker
            geocodedPoints.push({ lat: coords.lat, lng: coords.lng, row: f.row, bundle: f.bundle });
            geocodeFailures = geocodeFailures.filter(x => x !== f);
            addSingleMarker(coords, f.row, f.bundle);
            updateNotFoundBar();
            // Mark item as resolved
            status.textContent = 'Found!';
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

  // ─── CSV Parser ───────────────────────────────────
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
    if (lines.length === 0) return [];

    const parseLine = (line) => {
      const fields = []; let f = '', q = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          if (q && line[i + 1] === '"') { f += '"'; i++; } else q = !q;
        } else if (ch === ',' && !q) {
          fields.push(f.trim());
          f = '';
        } else {
          f += ch;
        }
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
      // Strip trailing ".0" from the address number — Excel/spreadsheets
      // sometimes export integer columns as floats (e.g. "123.0" → "123")
      if (obj['#'] && /^\d+\.0$/.test(obj['#'])) {
        obj['#'] = obj['#'].replace(/\.0$/, '');
      }
      return obj;
    });
  }

  // ─── Bundle Grouping ──────────────────────────────
  // Groups by Bundle column if present, otherwise by City.
  function groupIntoBundles(records) {
    if (!records.length) return [];

    // Check if a non-empty Bundle column exists
    const hasBundle = records.some(r => (r['Bundle'] || '').trim() !== '');
    const getKey = (r) => hasBundle
      ? ((r['Bundle'] || '').trim() || (r['City'] || '').trim() || 'Unknown')
      : ((r['City'] || '').trim() || 'Unknown');

    const map = new Map();
    records.forEach(r => {
      const key = getKey(r);
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(r);
    });

    return Array.from(map.entries()).map(([key, rows]) => {
      const city = rows[0]['City'] || '';
      const bundleId = rows[0]['BundleID'] || '';

      // Pad MRU ids to 6 digits, then decode CC AA RR
      const padMru = (id) => id.padStart(6, '0');
      const mruIds = [...new Set(rows.map(r => padMru((r['MRU id'] || '').trim())).filter(Boolean))];
      const firstMru = mruIds[0] || '000000';
      const mruCycle = firstMru.substring(0, 2);   // first 2 digits
      const mruArea = firstMru.substring(2, 4);   // middle 2 digits
      // Route numbers = last 2 digits of each unique MRU id
      const routeNums = mruIds.map(id => id.substring(4, 6));

      const isRead = (r) => (r['READING'] || '').trim() !== '' ||
        ((r['SKIP'] || '').trim() !== '' && (r['SKIP'] || '').trim() !== 'Other');
      const read = rows.filter(isRead).length;
      const total = rows.length;
      const estVal = (r) => parseInt(r['# EST'] || r['Estimates'] || '0', 10);
      const est3Rows = rows.filter(r => estVal(r) === 3);
      const est46Rows = rows.filter(r => { const v = estVal(r); return v >= 4 && v <= 6; });
      const est7plusRows = rows.filter(r => estVal(r) >= 7);
      const est3 = est3Rows.length;
      const est46 = est46Rows.length;
      const est7plus = est7plusRows.length;
      const est3Read = est3Rows.filter(isRead).length;
      const est46Read = est46Rows.filter(isRead).length;
      const est7plusRead = est7plusRows.filter(isRead).length;
      return { key, bundleName: key, mruIds, routeNums, mruCycle, mruArea, bundleId, rows, city, total, read, est3, est46, est7plus, est3Read, est46Read, est7plusRead };
    });
  }

  // ─── Bundle Status ────────────────────────────────
  function getBundleStatus(bundle) {
    if (sentBundles.has(bundle.key)) return 'sent';
    const allDone = bundle.rows.length > 0 && bundle.rows.every(r =>
      (r['READING'] || '').trim() !== '' || (r['SKIP'] || '').trim() !== ''
    );
    if (allDone) return 'complete';
    return 'in-progress';
  }

  function statusLabel(status) {
    if (status === 'sent') return 'Email Sent';
    if (status === 'complete') return 'Complete';
    return 'In Progress';
  }

  // ─── Render Dashboard ─────────────────────────────
  function renderHome() {
    const totalCards = bundles.reduce((s, b) => s + b.total, 0);
    const totalRead = bundles.reduce((s, b) => s + b.read, 0);
    const totalSent = bundles.filter(b => sentBundles.has(b.key)).length;

    sumBundles.textContent = bundles.length;
    sumCards.textContent = totalCards;
    sumRead.textContent = totalRead;
    sumSent.textContent = totalSent;

    if (bundles.length === 0) {
      uploadPrompt.classList.remove('hidden');
      bundleList.classList.add('hidden');
      bundleList.innerHTML = '';
      return;
    }

    uploadPrompt.classList.add('hidden');
    bundleList.classList.remove('hidden');
    bundleList.innerHTML = '';

    const frag = document.createDocumentFragment();
    bundles.forEach(bundle => {
      const status = getBundleStatus(bundle);
      const pct = bundle.total > 0
        ? Math.round((bundle.read / bundle.total) * 100) : 0;
      const isSent = status === 'sent';

      const card = document.createElement('div');
      card.className = 'bundle-card';
      card.innerHTML = `
        <div class="bundle-card-strip ${status}"></div>
        <div class="bundle-card-body">
          <div class="bundle-card-top">
            <div>
              <div class="bundle-mru">${bundle.mruIds.length} Route${bundle.mruIds.length !== 1 ? 's' : ''}</div>
              <div class="bundle-name">${esc(bundle.bundleName)}</div>
              <div class="bundle-area">Area ${esc(bundle.mruArea)}</div>
            </div>
            <span class="status-badge ${status}">
              <span class="status-dot"></span>
              ${statusLabel(status)}
            </span>
          </div>

          <div class="bundle-card-count">
            <span class="bundle-count-label">${bundle.read} of ${bundle.total} cards read</span>
            <span class="bundle-count-frac">${pct}%</span>
          </div>
          <div class="bundle-progress-bar">
            <div class="bundle-progress-fill ${status}" style="width:${pct}%"></div>
          </div>

          <div class="bundle-card-footer">
            <div class="bundle-meta-item">
              <span class="bundle-meta-label">Cards</span>
              <span class="bundle-meta-val">${bundle.total}</span>
            </div>
            <div class="bundle-meta-item">
              <span class="bundle-meta-label">3 Est</span>
              <span class="bundle-meta-val"><span class="meta-complete">${bundle.est3Read}</span>/${bundle.est3}</span>
            </div>
            <div class="bundle-meta-item">
              <span class="bundle-meta-label">4-6 Est</span>
              <span class="bundle-meta-val"><span class="meta-complete">${bundle.est46Read}</span>/${bundle.est46}</span>
            </div>
            <div class="bundle-meta-item">
              <span class="bundle-meta-label">7+ Est</span>
              <span class="bundle-meta-val"><span class="meta-complete">${bundle.est7plusRead}</span>/${bundle.est7plus}</span>
            </div>
          </div>

          ${status === 'complete' || isSent ? `
          <button class="bundle-send-btn ${isSent ? 'sent-state' : ''}"
                  data-key="${esc(bundle.key)}"
                  ${isSent ? 'disabled' : ''}>
            ${isSent
            ? `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg> Sent`
            : `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg> Send Bundle`
          }
          </button>` : ''}

          <button class="bundle-email-btn" title="Email this bundle">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
              <rect x="2" y="4" width="20" height="16" rx="2"/>
              <polyline points="2,4 12,13 22,4"/>
            </svg>
          </button>

          <button class="bundle-delete-btn" title="Delete this bundle">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
              <polyline points="3 6 5 6 21 6"/>
              <path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>
              <path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/>
            </svg>
          </button>
        </div>
      `;

      // Send button handler
      const sendBtn = card.querySelector('.bundle-send-btn');
      if (sendBtn && !isSent) {
        sendBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          markAsSent(bundle.key);
        });
      }

      // Email button handler
      card.querySelector('.bundle-email-btn').addEventListener('click', (e) => {
        e.stopPropagation();
        emailBundle(bundle);
      });

      // Delete button handler
      card.querySelector('.bundle-delete-btn').addEventListener('click', (e) => {
        e.stopPropagation();
        showDeleteConfirm(bundle);
      });

      // Searchable blob: all addresses + serials + meter sizes in this bundle
      card.dataset.search = bundle.rows.map(r => [
        (r['#'] || '') + ' ' + (r['STREET'] || ''),
        r['Serial No.'] || '',
        r['MTR SIZE'] || '',
      ].join(' ')).join(' ').toLowerCase();

      // Card tap → open bundle detail
      card.addEventListener('click', () => showBundle(bundle));

      frag.appendChild(card);
    });

    bundleList.appendChild(frag);
    const verEl = document.createElement('div');
    verEl.className = 'app-version';
    verEl.textContent = APP_VERSION;
    bundleList.appendChild(verEl);
    updateRecoverBar();
  }

  // ─── Mark Bundle as Sent ──────────────────────────
  function markAsSent(key) {
    sentBundles.add(key);
    saveSentState();
    renderHome();
    showToast(`Bundle "${key}" marked as sent`);
  }

  // ─── Navigation ───────────────────────────────────
  function goHome() {
    if (pickStartMode) exitPickStartMode();
    if (leafletMap) leafletMap.stopLocate();
    clearGpsTimer();
    viewBundle.classList.add('hidden');
    viewCard.classList.add('hidden');
    viewMap.classList.add('hidden');
    viewHome.classList.remove('hidden');
  }

  function cardGoBack() {
    viewCard.classList.add('hidden');
    if (cardReturnTarget === 'map') {
      viewBundle.classList.add('hidden');  // showBundleDetail may have un-hidden this during save
      viewMap.classList.remove('hidden');
      requestAnimationFrame(() => {
        requestAnimationFrame(() => { if (leafletMap) leafletMap.invalidateSize(); });
      });
    } else {
      viewBundle.classList.remove('hidden');
    }
  }

  function showBundle(bundle) {
    pendingBundle = bundle;
    if (!readerName) {
      readerInput.value = '';
      readerModal.classList.remove('hidden');
      setTimeout(() => readerInput.focus(), 50);
    } else {
      showBundleDetail(bundle);
    }
  }

  // ─── Recompute Bundle Stats After Edits ───────────
  function refreshBundleStats(bundle) {
    const isRead = (r) => (r['READING'] || '').trim() !== '' ||
      ((r['SKIP'] || '').trim() !== '' && (r['SKIP'] || '').trim() !== 'Other');
    const estVal = (r) => parseInt(r['# EST'] || r['Estimates'] || '0', 10);
    const est3Rows = bundle.rows.filter(r => estVal(r) === 3);
    const est46Rows = bundle.rows.filter(r => { const v = estVal(r); return v >= 4 && v <= 6; });
    const est7plusRows = bundle.rows.filter(r => estVal(r) >= 7);
    bundle.read = bundle.rows.filter(isRead).length;
    bundle.est3 = est3Rows.length;
    bundle.est46 = est46Rows.length;
    bundle.est7plus = est7plusRows.length;
    bundle.est3Read = est3Rows.filter(isRead).length;
    bundle.est46Read = est46Rows.filter(isRead).length;
    bundle.est7plusRead = est7plusRows.filter(isRead).length;
  }

  // ─── Card Detail View ─────────────────────────────
  function showCardDetail(row, bundle, returnTarget) {
    currentRow = row;
    currentBundle = bundle;
    if (returnTarget) cardReturnTarget = returnTarget;
    const backLabel = cardReturnTarget === 'map' ? 'Map' : 'List';
    document.getElementById('card-back-label').textContent = backLabel;
    document.getElementById('card-back-nav-label').textContent = 'Map';

    const num = row['#'] || '';
    const street = row['STREET'] || '';
    const address = [num, street].filter(Boolean).join(' ');
    const loc = (row['LOC'] || '').trim();
    const spec = (row['SPEC INSTRUCTIONS'] || '').trim();

    cardDtSeq.textContent = `Seq #${row['Seq #'] || '—'}`;
    cardDtAddress.textContent = address || '—';
    cardDtCity.textContent = row['City'] || '—';
    const instrument = (row['Instrument'] || '').trim();
    cardDtInstrument.textContent = instrument || '—';
    cardDtInstrumentRow.style.display = instrument ? '' : 'none';
    const mtrSize = row['MTR SIZE'] || '';
    cardDtMtrSize.textContent = mtrSize || '—';
    cardDtSerial.textContent = row['Serial No.'] || '—';
    cardDtEst.textContent = row['# EST'] || '0';

    const expectedDigits = getExpectedDigits(mtrSize);
    cardDtReading.maxLength = expectedDigits || 10;
    cardDtReading.placeholder = expectedDigits
      ? `Enter ${expectedDigits}-digit reading…`
      : 'Enter meter reading…';

    // Location — hide row if empty
    cardDtLoc.textContent = loc;
    cardDtLocRow.style.display = loc ? '' : 'none';

    // Special instructions — hide block if empty
    cardDtSpec.textContent = spec;
    cardDtSpecWrap.classList.toggle('hidden', !spec);

    // Pre-fill reading, skip, comments, and date
    cardDtReading.value = (row['READING'] || '').trim();
    const hasReading = cardDtReading.value !== '';
    setReadingLocked(hasReading);
    document.getElementById('card-menu-delete-reading').classList.toggle('hidden', !hasReading);
    const savedSkip = row['SKIP'] || '';
    cardDtSkip.value = savedSkip;
    cardDtSkipOther.value = row['SKIP_OTHER'] || '';
    cardDtSkipOther.classList.toggle('hidden', savedSkip !== 'Other');
    cardDtComments.value = row['COMMENTS'] || '';
    const today = new Date().toLocaleDateString('en-CA');
    const savedDate = (row['READ DATE'] || '').trim();
    const isComplete = (row['READING'] || '').trim() !== '' ||
      ((row['SKIP'] || '').trim() !== '' && (row['SKIP'] || '').trim() !== 'Other') ||
      (row['COMMENTS'] || '').trim() !== '';
    cardDtReadDate.value = savedDate || today;
    cardDtReadDate.readOnly = isComplete;
    cardDtReadDate.classList.toggle('date-locked', isComplete);

    // Prev / next nav
    const rows = bundle.rows;
    const idx = rows.indexOf(row);
    cardNavPos.textContent = `${idx + 1} / ${rows.length}`;
    const complete = rows.filter(r => (r['READING'] || '').trim() || (r['SKIP'] || '').trim()).length;
    cardNavComplete.textContent = `${complete} / ${rows.length} complete`;
    cardPrevBtn.disabled = idx === 0;
    cardNextBtn.disabled = idx === rows.length - 1;

    viewBundle.classList.add('hidden');
    viewCard.classList.remove('hidden');
    setTimeout(() => cardDtReading.focus(), 80);
  }

  cardPrevBtn.addEventListener('click', () => {
    const rows = currentBundle.rows;
    const idx = rows.indexOf(currentRow);
    if (idx > 0) showCardDetail(rows[idx - 1], currentBundle);
  });

  cardNextBtn.addEventListener('click', () => {
    const rows = currentBundle.rows;
    const idx = rows.indexOf(currentRow);
    if (idx < rows.length - 1) showCardDetail(rows[idx + 1], currentBundle);
  });

  // Save reading handler
  cardDtSaveBtn.addEventListener('click', () => {
    const reading = padReading(cardDtReading.value.trim(), currentRow['MTR SIZE']);
    const skip = cardDtSkip.value;
    const date = cardDtReadDate.value;
    const comments = cardDtComments.value.trim();

    if (reading && skip) {
      showToast('Cannot have both a reading and a skip — clear one first.', true);
      return;
    }

    const isComplete = reading !== '' || (skip !== '' && skip !== 'Other') || comments !== '';
    currentRow['READING'] = reading;
    currentRow['READ DATE'] = isComplete ? date : '';
    currentRow['SKIP'] = skip;
    currentRow['SKIP_OTHER'] = skip === 'Other' ? cardDtSkipOther.value.trim() : '';
    currentRow['COMMENTS'] = comments;

    updateMapMarkerForRow(currentRow);  // refresh marker color immediately if on map

    // Lock or unlock the date field to match the new state
    cardDtReadDate.readOnly = isComplete;
    cardDtReadDate.classList.toggle('date-locked', isComplete);

    // Lock the reading input if a reading was saved
    setReadingLocked(reading !== '');
    document.getElementById('card-menu-delete-reading').classList.toggle('hidden', reading === '');

    refreshBundleStats(currentBundle);
    saveRecordsBackup();               // persist every save in case of crash
    renderHome();                      // keeps main page cards in sync

    // Brief "Saved" confirmation, then advance to the next card
    cardDtSaveBtn.textContent = '✓ Saved';
    cardDtSaveBtn.disabled = true;
    playDing();

    const rows = currentBundle.rows;
    const savedIdx = rows.indexOf(currentRow);

    setTimeout(() => {
      cardDtSaveBtn.textContent = 'Save Reading';
      cardDtSaveBtn.disabled = false;
      showBundleDetail(currentBundle);   // re-renders list + header
      const nextIdx = reverseDirection ? savedIdx - 1 : savedIdx + 1;
      if (nextIdx >= 0 && nextIdx < rows.length) {
        showCardDetail(rows[nextIdx], currentBundle);
      }
    }, 600);
  });

  // ─── Reader Name Modal Handlers ───────────────────
  function setReaderName(name) {
    readerName = name;
    homeReaderName.textContent = name || 'Tap to set name';
    bdReaderName.textContent = name;
    try { localStorage.setItem(READER_KEY, name); } catch (_) { }
  }

  readerConfirm.addEventListener('click', () => {
    const name = readerInput.value.trim();
    if (!name) { readerInput.focus(); return; }
    setReaderName(name);
    readerModal.classList.add('hidden');
    if (pendingBundle) showBundleDetail(pendingBundle);
  });

  readerCancel.addEventListener('click', () => {
    readerModal.classList.add('hidden');
    pendingBundle = null;
  });

  readerInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') readerConfirm.click();
  });

  // ─── Tap Name to Edit ─────────────────────────────
  homeReaderName.addEventListener('click', () => {
    readerInput.value = readerName;
    pendingBundle = null;
    readerModal.classList.remove('hidden');
    setTimeout(() => readerInput.focus(), 50);
  });

  // ─── Home Bundle Search ───────────────────────────
  homeSearch.addEventListener('input', () => {
    const q = homeSearch.value.trim().toLowerCase();
    let visible = 0;
    Array.from(bundleList.children).forEach(el => {
      if (el.classList.contains('bd-no-results')) return;
      if (!q) { el.hidden = false; visible++; return; }
      const match = (el.dataset.search || '').includes(q);
      el.hidden = !match;
      if (match) visible++;
    });
    let noRes = bundleList.querySelector('.bd-no-results');
    if (q && visible === 0) {
      if (!noRes) {
        noRes = document.createElement('div');
        noRes.className = 'bd-no-results';
        bundleList.appendChild(noRes);
      }
      noRes.textContent = `No bundles match "${homeSearch.value.trim()}"`;
    } else if (noRes) {
      noRes.remove();
    }
  });

  // ─── Address List Search ──────────────────────────
  function applyBdFilters() {
    const q = bdSearch.value.trim().toLowerCase();
    const filterActive = bdShowMissed || bdShowSkipped;
    let visible = 0;
    Array.from(addressList.children).forEach(el => {
      if (el.classList.contains('bd-no-results')) return;
      const matchesSearch = !q ||
        (el.dataset.address || '').includes(q) ||
        (el.dataset.serial || '').includes(q) ||
        (el.dataset.mtrsize || '').includes(q);
      const matchesFilter = !filterActive ||
        (bdShowMissed && el.dataset.status === 'unread') ||
        (bdShowSkipped && el.dataset.status === 'skip');
      el.hidden = !(matchesSearch && matchesFilter);
      if (!el.hidden) visible++;
    });
    let noRes = addressList.querySelector('.bd-no-results');
    if ((q || filterActive) && visible === 0) {
      if (!noRes) {
        noRes = document.createElement('div');
        noRes.className = 'bd-no-results';
        addressList.appendChild(noRes);
      }
      noRes.textContent = q
        ? `No results for "${bdSearch.value.trim()}"`
        : 'No cards match the selected filter';
    } else if (noRes) {
      noRes.remove();
    }
  }

  bdSearch.addEventListener('input', applyBdFilters);

  document.querySelector('#view-bundle .bd-search-icon').addEventListener('click', () => {
    bdSearch.blur();
  });

  bdFilterMissedBtn.addEventListener('click', () => {
    bdShowMissed = !bdShowMissed;
    bdFilterMissedBtn.classList.toggle('active', bdShowMissed);
    applyBdFilters();
  });

  bdFilterSkippedBtn.addEventListener('click', () => {
    bdShowSkipped = !bdShowSkipped;
    bdFilterSkippedBtn.classList.toggle('active', bdShowSkipped);
    applyBdFilters();
  });

  document.getElementById('bd-optimize-btn').addEventListener('click', () => {
    enterPickStartMode();
  });

  document.getElementById('bd-optimize-cancel-btn').addEventListener('click', () => {
    exitPickStartMode();
  });

  document.getElementById('bd-reset-order-btn').addEventListener('click', () => {
    resetRouteOrder();
  });

  // ─── Totals / Daily Report View ───────────────────
  function showTotalsView(bundle) {
    const rates = bundleRates[bundle.key] || null;
    const fmtRate = (r) => r != null ? `$${r.toFixed(2)}` : '—';
    const today = new Date().toLocaleDateString('en-CA');
    const issuedDate = rates?.issuedDate || today;
    const dueDate = rates?.dueDate || '';
    const routeStr = bundle.mruIds.length
      ? `A${bundle.mruArea} – ${bundle.routeNums.join(', ')}`
      : '—';

    const row = (label, rate, read, issued) => `
      <tr>
        <td>${label}</td>
        <td class="report-rate">${fmtRate(rate)}</td>
        <td class="report-col-center">${read}</td>
        <td class="report-col-center report-divider">/</td>
        <td class="report-col-center">${issued}</td>
      </tr>`;

    document.getElementById('totals-body').innerHTML = `
      <div class="report-title-block">
        <div class="report-company">M.E.T. Utilities Management Ltd.</div>
        <div class="report-main-title">METER READER DAILY REPORT</div>
        <div class="report-sub-title">Outcards Regular Calls</div>
      </div>

      <div class="report-info-grid">
        <div class="report-info-item">
          <span class="report-info-label">Cycle</span>
          <span class="report-info-value">${esc(bundle.mruCycle)}</span>
        </div>
        <div class="report-info-item">
          <span class="report-info-label">Issued / Due Date</span>
          <span class="report-info-value">${issuedDate} / ${dueDate || '—'}</span>
        </div>
        <div class="report-info-item">
          <span class="report-info-label">Route #</span>
          <span class="report-info-value">${esc(routeStr)}</span>
        </div>
        <div class="report-info-item">
          <span class="report-info-label">City / Route</span>
          <span class="report-info-value">${esc(bundle.city || bundle.bundleName)}</span>
        </div>
        <div class="report-info-item">
          <span class="report-info-label">Total Issued</span>
          <span class="report-info-value report-info-big">${bundle.total}</span>
        </div>
        <div class="report-info-item">
          <span class="report-info-label">Total Read</span>
          <span class="report-info-value report-info-big">${bundle.read}</span>
        </div>
      </div>

      <table class="report-table">
        <thead>
          <tr>
            <th>Category</th>
            <th>Rate</th>
            <th class="report-col-center">Read</th>
            <th class="report-col-center"></th>
            <th class="report-col-center">Issued</th>
          </tr>
        </thead>
        <tbody>
          ${row('Est 3', rates ? rates.est3 : null, bundle.est3Read, bundle.est3)}
          ${row('Est 4–6', rates ? rates.est46 : null, bundle.est46Read, bundle.est46)}
          ${row('Est 7+', rates ? rates.est7p : null, bundle.est7plusRead, bundle.est7plus)}
          <tr class="report-total-row">
            <td>Total</td>
            <td></td>
            <td class="report-col-center">${bundle.read}</td>
            <td class="report-col-center report-divider">/</td>
            <td class="report-col-center">${bundle.total}</td>
          </tr>
        </tbody>
      </table>

      <div class="report-reader-section">
        <div class="report-field-row">
          <span class="report-field-label">Meter Reader Name</span>
          <span class="report-field-value">${esc(readerName || '—')}</span>
        </div>
        <div class="report-field-row">
          <span class="report-field-label">Signed (Meter Reader)</span>
          <span class="report-sig-line"></span>
        </div>
      </div>
    `;

    viewBundle.classList.add('hidden');
    viewTotals.classList.remove('hidden');
  }

  // ─── Bundle Detail View ───────────────────────────
  function showBundleDetail(bundle, forcePreserveSearch) {
    pendingBundle = null;
    document.getElementById('bd-report-btn').onclick = () => showTotalsView(bundle);

    // Reset optimization state when switching to a different bundle
    const isNewBundle = bundle !== currentBundle;
    if (isNewBundle) {
      if (pickStartMode) exitPickStartMode();
      routeOptimized = false;
      originalRowsOrder = null;
      document.getElementById('bd-reset-order-btn').classList.add('hidden');
    }
    currentBundle = bundle;

    const status = getBundleStatus(bundle);
    const pct = bundle.total > 0 ? Math.round((bundle.read / bundle.total) * 100) : 0;

    // Header — row 1
    bdStatusBadge.textContent = statusLabel(status);
    bdStatusBadge.className = `status-badge ${status}`;

    // Header — row 2
    bdRoutes.textContent = `${bundle.mruIds.length} Route${bundle.mruIds.length !== 1 ? 's' : ''}`;
    bdTitle.textContent = bundle.bundleName;
    bdArea.textContent = `Area ${bundle.mruArea}`;

    // Progress
    bdCountLabel.textContent = `${bundle.read} of ${bundle.total} cards read`;
    bdCountPct.textContent = `${pct}%`;
    bdProgressFill.style.width = `${pct}%`;
    bdProgressFill.className = `bd-progress-fill ${status}`;

    // Stats footer
    bdTotal.textContent = bundle.total;
    bdEst3r.textContent = bundle.est3Read;
    bdEst3.textContent = bundle.est3;
    bdEst46r.textContent = bundle.est46Read;
    bdEst46.textContent = bundle.est46;
    bdEst7pr.textContent = bundle.est7plusRead;
    bdEst7p.textContent = bundle.est7plus;

    bdReaderName.textContent = readerName;

    // Sort rows by Seq #
    const sorted = [...bundle.rows].sort((a, b) => {
      const sa = parseInt(a['Seq #'] || '0', 10);
      const sb = parseInt(b['Seq #'] || '0', 10);
      return sa - sb;
    });

    const preserveSearch = forcePreserveSearch || !isNewBundle;
    const savedSearch = preserveSearch ? bdSearch.value : '';
    bdSearch.value = '';
    if (!preserveSearch) {
      bdShowMissed = false;
      bdShowSkipped = false;
      bdFilterMissedBtn.classList.remove('active');
      bdFilterSkippedBtn.classList.remove('active');
    }
    addressList.innerHTML = '';
    const frag = document.createDocumentFragment();

    sorted.forEach(row => {
      const seq = row['Seq #'] || '';
      const num = row['#'] || '';
      const street = row['STREET'] || '';
      const address = [num, street].filter(Boolean).join(' ');
      const spec = row['SPEC INSTRUCTIONS'] || '';
      const loc = row['LOC'] || '';
      const mtrSize = row['MTR SIZE'] || '';
      const serial = row['Serial No.'] || '';
      const reading = (row['READING'] || '').trim();
      const skip = (row['SKIP'] || '').trim();
      const readDate = (row['READ DATE'] || '').trim();
      const comment = (row['COMMENTS'] || '').trim();
      const estVal = parseInt(row['# EST'] || '0', 10);

      let estClass = 'est-ok';
      let estLabel = '';
      if (estVal === 3) { estClass = 'est-3'; estLabel = '3 Est'; }
      else if (estVal >= 4 && estVal <= 6) { estClass = 'est-46'; estLabel = `${estVal} Est`; }
      else if (estVal >= 7) { estClass = 'est-7p'; estLabel = `${estVal} Est`; }

      const meterMeta = [
        loc ? `Loc: ${loc}` : '',
        mtrSize ? `Size: ${mtrSize}` : '',
        serial ? `#: ${serial}` : '',
      ].filter(Boolean).join('  ·  ');

      const card = document.createElement('div');
      card.className = `addr-card${skip ? ' addr-skip' : reading ? ' addr-read' : ''}`;
      card.dataset.address = address.toLowerCase();
      card.dataset.serial = serial.toLowerCase();
      card.dataset.mtrsize = mtrSize.toLowerCase();
      card.dataset.status = skip ? 'skip' : reading ? 'read' : 'unread';
      card.innerHTML = `
        <div class="addr-seq">${esc(seq)}</div>
        <div class="addr-info">
          <div class="addr-street">${esc(address) || '—'}</div>
          ${meterMeta ? `<div class="addr-meter">${esc(meterMeta)}</div>` : ''}
          ${spec ? `<div class="addr-spec">${esc(spec)}</div>` : ''}
          ${(readDate || comment) ? `<div class="addr-read-info">${readDate ? `<span class="addr-read-date">${esc(readDate)}</span>` : ''}${comment ? `<span class="addr-comment">${esc(comment)}</span>` : ''}</div>` : ''}
        </div>
        <div class="addr-right">
          ${skip ? `<span class="addr-status-badge addr-skip-badge">Skip</span>` : ''}
          ${!skip && reading ? `<span class="addr-status-badge addr-read-badge">Read</span>` : ''}
          ${estLabel ? `<span class="addr-est-badge ${estClass}">${estLabel}</span>` : ''}
        </div>
      `;
      card.addEventListener('click', () => {
        if (pickStartMode) { applyRouteOptimization(bundle, row); return; }
        showCardDetail(row, bundle, 'bundle');
      });

      frag.appendChild(card);
    });

    addressList.appendChild(frag);

    if (savedSearch) {
      bdSearch.value = savedSearch;
      applyBdFilters();
    }

    viewHome.classList.add('hidden');
    viewBundle.classList.remove('hidden');
  }

  // ─── CSV File Load ────────────────────────────────
  const dupModal = document.getElementById('dup-bundle-modal');
  const dupDesc = document.getElementById('dup-bundle-desc');
  const dupCancelBtn = document.getElementById('dup-bundle-cancel');

  function commitLoad(records, rates) {
    allRecords = [...allRecords, ...records];
    bundles = groupIntoBundles(allRecords);
    geocodedPoints = [];  // reset so map re-geocodes with new addresses
    if (rates) {
      const issuedDate = new Date().toLocaleDateString('en-CA');
      const hasBundle = records.some(r => (r['Bundle'] || '').trim() !== '');
      getNewBundleKeys(records).forEach(key => {
        const rec = records.find(r =>
          (hasBundle
            ? ((r['Bundle'] || '').trim() || (r['City'] || '').trim() || 'Unknown')
            : ((r['City'] || '').trim() || 'Unknown')) === key
        );
        const dueDate = (rec?.['DUE_DATE'] || '').trim();
        bundleRates[key] = { ...rates, issuedDate, dueDate };
      });
      saveBundleRates();
    }
    renderHome();
    saveRecordsBackup();
    showToast(`Added ${records.length} records · ${bundles.length} bundles total`);
  }

  // ─── Delete Bundle Modal ──────────────────────────
  const delModal = document.getElementById('del-bundle-modal');
  const delDesc = document.getElementById('del-bundle-desc');
  const delCancelBtn = document.getElementById('del-bundle-cancel');
  const delConfirmBtn = document.getElementById('del-bundle-confirm');
  let pendingDelete = null;   // bundle queued for deletion

  function showDeleteConfirm(bundle) {
    delDesc.textContent = `Delete "${bundle.bundleName}"? All records and readings will be removed.`;
    pendingDelete = bundle;
    delModal.classList.remove('hidden');
  }

  delCancelBtn.addEventListener('click', () => {
    delModal.classList.add('hidden');
    pendingDelete = null;
  });

  delConfirmBtn.addEventListener('click', () => {
    delModal.classList.add('hidden');
    if (!pendingDelete) return;
    const key = pendingDelete.key;

    // Archive rows before removing
    const removedRows = allRecords.filter(r =>
      ((r['Bundle'] || r['City'] || '').trim() || 'Unknown') === key
    );
    deletedBundles.push({ key, bundleName: pendingDelete.bundleName, rows: removedRows, deletedAt: Date.now() });
    saveDeletedBundles();

    allRecords = allRecords.filter(r =>
      ((r['Bundle'] || r['City'] || '').trim() || 'Unknown') !== key
    );
    sentBundles.delete(key);
    bundles = groupIntoBundles(allRecords);
    saveRecordsBackup();
    saveSentState();
    renderHome();
    updateRecoverBar();
    showToast(`Bundle "${pendingDelete.bundleName}" deleted`);
    pendingDelete = null;
  });

  // ─── Recover Deleted Bundles ──────────────────────
  const recoverModal = document.getElementById('recover-modal');
  const recoverList = document.getElementById('recover-list');
  document.getElementById('recover-modal-close').addEventListener('click', () => {
    recoverModal.classList.add('hidden');
  });

  function updateRecoverBar() {
    // Remove any existing recover bar from bundle-list
    const existing = bundleList.querySelector('.recover-bar');
    if (existing) existing.remove();

    const cutoff = Date.now() - TWO_WEEKS;
    const visible = deletedBundles.filter(d => d.deletedAt >= cutoff);
    if (!visible.length) return;

    const bar = document.createElement('div');
    bar.className = 'recover-bar';
    bar.innerHTML = `
      <button class="recover-btn" id="recover-btn">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
          <polyline points="1 4 1 10 7 10"/>
          <path d="M3.51 15a9 9 0 1 0 .49-4.5"/>
        </svg>
        Recover Deleted Bundles
      </button>`;
    bar.querySelector('.recover-btn').addEventListener('click', openRecoverModal);
    bundleList.appendChild(bar);

    // If no bundles are loaded, renderHome() hides bundleList — un-hide it so the recover bar shows
    if (bundles.length === 0) {
      bundleList.classList.remove('hidden');
    }
  }

  function openRecoverModal() {
    const cutoff = Date.now() - TWO_WEEKS;
    const visible = deletedBundles.filter(d => d.deletedAt >= cutoff);
    recoverList.innerHTML = '';
    visible.forEach(d => {
      const daysLeft = Math.ceil((d.deletedAt + TWO_WEEKS - Date.now()) / (24 * 60 * 60 * 1000));
      const deletedDate = new Date(d.deletedAt).toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
      const item = document.createElement('div');
      item.className = 'recover-item';
      item.innerHTML = `
        <div class="recover-item-info">
          <div class="recover-item-name">${esc(d.bundleName)}</div>
          <div class="recover-item-meta">Deleted ${deletedDate} · ${d.rows.length} records · expires in ${daysLeft}d</div>
        </div>
        <button class="recover-item-btn">Restore</button>
      `;
      item.querySelector('.recover-item-btn').addEventListener('click', () => {
        const currentForBundle = allRecords.filter(r =>
          ((r['Bundle'] || r['City'] || '').trim() || 'Unknown') === d.key
        );
        const alreadyLoaded = currentForBundle.length > 0;

        let restoredRows;
        if (!alreadyLoaded) {
          // Simple restore — no conflict
          restoredRows = d.rows;
        } else {
          // Merge: current wins for actioned records; archived fills in unread ones
          const matchKey = r => (r['Serial No.'] || '').trim()
            || ((r['#'] || '').trim() + '|' + (r['STREET'] || '').trim());
          const isActioned = r => !!(r['READING'] || '').trim() || !!(r['SKIP'] || '').trim();

          const archivedMap = new Map(d.rows.map(r => [matchKey(r), r]));
          const currentMap = new Map(currentForBundle.map(r => [matchKey(r), r]));

          let recovered = 0;

          // Start with current records, filling in archived readings where current is unread
          restoredRows = currentForBundle.map(cur => {
            if (isActioned(cur)) return cur; // current wins
            const arch = archivedMap.get(matchKey(cur));
            if (arch && isActioned(arch)) {
              recovered++;
              return {
                ...cur,
                'READING': arch['READING'] || '',
                'READ DATE': arch['READ DATE'] || '',
                'SKIP': arch['SKIP'] || '',
                'SKIP_OTHER': arch['SKIP_OTHER'] || '',
                'COMMENTS': arch['COMMENTS'] || '',
              };
            }
            return cur;
          });

          // Append archived records that don't exist in current copy
          d.rows.forEach(arch => {
            if (!currentMap.has(matchKey(arch))) {
              restoredRows.push(arch);
              if (isActioned(arch)) recovered++;
            }
          });

          allRecords = allRecords.filter(r =>
            ((r['Bundle'] || r['City'] || '').trim() || 'Unknown') !== d.key
          );
          allRecords = [...allRecords, ...restoredRows];
          bundles = groupIntoBundles(allRecords);
          deletedBundles = deletedBundles.filter(x => x !== d);
          saveDeletedBundles();
          saveRecordsBackup();
          renderHome();
          updateRecoverBar();
          recoverModal.classList.add('hidden');
          showToast(recovered > 0
            ? `Bundle "${d.bundleName}" merged — ${recovered} reading${recovered !== 1 ? 's' : ''} recovered`
            : `Bundle "${d.bundleName}" merged — no new readings to recover`
          );
          return;
        }

        allRecords = allRecords.filter(r =>
          ((r['Bundle'] || r['City'] || '').trim() || 'Unknown') !== d.key
        );
        allRecords = [...allRecords, ...restoredRows];
        bundles = groupIntoBundles(allRecords);
        deletedBundles = deletedBundles.filter(x => x !== d);
        saveDeletedBundles();
        saveRecordsBackup();
        renderHome();
        updateRecoverBar();
        recoverModal.classList.add('hidden');
        showToast(`Bundle "${d.bundleName}" restored`);
      });
      recoverList.appendChild(item);
    });
    recoverModal.classList.remove('hidden');
  }

  dupCancelBtn.addEventListener('click', () => {
    dupModal.classList.add('hidden');
  });

  function getNewBundleKeys(records) {
    const hasBundle = records.some(r => (r['Bundle'] || '').trim() !== '');
    return [...new Set(records.map(r => hasBundle
      ? ((r['Bundle'] || '').trim() || (r['City'] || '').trim() || 'Unknown')
      : ((r['City'] || '').trim() || 'Unknown')
    ))];
  }

  csvFileInput.addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    if (!files.length) return;

    let completed = 0;
    let skipped = 0;
    const validRecordSets = [];

    files.forEach(file => {
      const reader = new FileReader();
      reader.onload = (ev) => {
        const records = parseCSV(ev.target.result);
        if (records.length && records[0].hasOwnProperty('MRU id')) {
          validRecordSets.push(records);
        } else {
          skipped++;
        }

        completed++;
        if (completed < files.length) return;

        // All files parsed — flatten and check for duplicates
        const incoming = validRecordSets.flat();
        if (!incoming.length) {
          showToast(`No valid records found`, true);
          return;
        }

        const existingKeys = new Set(bundles.map(b => b.key));
        const dupKeys = getNewBundleKeys(incoming).filter(k => existingKeys.has(k));

        if (dupKeys.length) {
          const names = dupKeys.join(', ');
          dupDesc.textContent = `The bundle${dupKeys.length > 1 ? 's' : ''} "${names}" ${dupKeys.length > 1 ? 'have' : 'has'} already been loaded. Loading it again would double the cards.`;
          dupModal.classList.remove('hidden');
        } else {
          const first = incoming[0];
          if (first.hasOwnProperty('RATE_3EST') && first.hasOwnProperty('RATE_46EST') && first.hasOwnProperty('RATE_7PEST')) {
            const rates = {
              est3: parseFloat(first['RATE_3EST']) || 0,
              est46: parseFloat(first['RATE_46EST']) || 0,
              est7p: parseFloat(first['RATE_7PEST']) || 0,
            };
            commitLoad(incoming, rates);
          } else {
            showRateModal(incoming);
          }
        }

        if (skipped) showToast(`${skipped} file${skipped > 1 ? 's' : ''} skipped — not a valid meter CSV`, true);
      };
      reader.readAsText(file);
    });

    // Reset so same file(s) can be re-selected if needed
    e.target.value = '';
  });

  // ─── Per-bundle Email ─────────────────────────────
  function buildReportText(bundle) {
    const rates = bundleRates[bundle.key] || null;
    const fmtRate = (r) => r != null ? `$${r.toFixed(2)}` : '—';
    const today = new Date().toLocaleDateString('en-CA');
    const issuedDate = rates?.issuedDate || today;
    const dueDate = rates?.dueDate || '';
    const routeStr = bundle.mruIds.length
      ? `A${bundle.mruArea} – ${bundle.routeNums.join(', ')}`
      : '—';

    const sep = '  ' + '─'.repeat(40);
    const row = (label, rate, read, total) =>
      `  ${label.padEnd(10)} ${fmtRate(rate).padEnd(8)} ${String(read).padStart(4)} / ${String(total)}`;

    return [
      'METER READER DAILY REPORT',
      'M.E.T. Utilities Management Ltd.',
      'Outcards Regular Calls',
      '',
      `Cycle:        ${bundle.mruCycle}`,
      `Issued Date:  ${issuedDate}`,
      `Due Date:     ${dueDate || '—'}`,
      `Route #:      ${routeStr}`,
      `City / Route: ${bundle.city || bundle.bundleName}`,
      `Total Issued: ${bundle.total}`,
      `Total Read:   ${bundle.read}`,
      '',
      '  Category   Rate     Read / Issued',
      sep,
      row('Est 3', rates?.est3 ?? null, bundle.est3Read, bundle.est3),
      row('Est 4–6', rates?.est46 ?? null, bundle.est46Read, bundle.est46),
      row('Est 7+', rates?.est7p ?? null, bundle.est7plusRead, bundle.est7plus),
      sep,
      row('Total', null, bundle.read, bundle.total),
      '',
      `Meter Reader: ${readerName || '—'}`,
      `Signed:       ______________________________`,
    ].join('\r\n');
  }

  function buildBundleCSV(bundle) {
    const rows = bundle.rows;
    if (!rows.length) return '';

    // Collect headers: original keys + ensure SKIP / SKIP_OTHER / COMMENTS included
    const baseHeaders = Object.keys(rows[0]);
    const extra = ['SKIP', 'SKIP_OTHER', 'COMMENTS'].filter(k => !baseHeaders.includes(k));
    const allHeaders = [...baseHeaders, ...extra];
    const FIRST_COLS = ['MTR SIZE', 'Serial No.', 'READ DATE', 'READING', 'COMMENTS'];
    const headers = [...FIRST_COLS, ...allHeaders.filter(h => !FIRST_COLS.includes(h))];
    const HEADER_RENAME = { 'MTR SIZE': 'MATERIAL_SIZE', 'Serial No.': 'MATERIAL_NO', 'READ DATE': 'DATE', 'COMMENTS': 'COMMENT' };

    function csvCell(val) {
      const s = (val == null ? '' : String(val));
      return (s.includes(',') || s.includes('"') || s.includes('\n'))
        ? `"${s.replace(/"/g, '""')}"` : s;
    }

    const lines = [headers.map(h => csvCell(HEADER_RENAME[h] || h)).join(',')];
    for (const row of rows) {
      const skip = (row['SKIP'] || '').trim();
      const skipOther = (row['SKIP_OTHER'] || '').trim();
      const readingVal = skip
        ? (skip === 'Other' && skipOther ? skipOther : skip)
        : (row['READING'] || '');
      lines.push(headers.map(h => {
        if (h === 'READING') return csvCell(readingVal);
        return csvCell(row[h] || '');
      }).join(','));
    }
    return lines.join('\r\n');
  }

  // ─── Daily Backup ─────────────────────────────────
  function buildAllRecordsCSV() {
    if (!allRecords.length) return '';
    const baseHeaders = Object.keys(allRecords[0]);
    const extra = ['SKIP', 'SKIP_OTHER', 'COMMENTS'].filter(k => !baseHeaders.includes(k));
    const headers = [...baseHeaders, ...extra];
    function csvCell(val) {
      const s = val == null ? '' : String(val);
      return (s.includes(',') || s.includes('"') || s.includes('\n'))
        ? `"${s.replace(/"/g, '""')}"` : s;
    }
    const lines = [headers.map(csvCell).join(',')];
    for (const row of allRecords) {
      const skip = (row['SKIP'] || '').trim();
      const skipOther = (row['SKIP_OTHER'] || '').trim();
      const readingVal = skip
        ? (skip === 'Other' && skipOther ? skipOther : skip)
        : (row['READING'] || '');
      lines.push(headers.map(h => {
        if (h === 'READING') return csvCell(readingVal);
        return csvCell(row[h] || '');
      }).join(','));
    }
    return lines.join('\r\n');
  }

  // ─── Auto-backup: silent write to authorized folder ───

  async function writeCSVToFolder(handle) {
    if (!allRecords.length) return;
    const dateStr = new Date().toLocaleDateString('en-CA');
    const csv = buildAllRecordsCSV();
    if (!csv) return;
    try {
      const fileHandle = await handle.getFileHandle(`meter-backup-${dateStr}.csv`, { create: true });
      const writable = await fileHandle.createWritable();
      await writable.write(csv);
      await writable.close();
      try { localStorage.setItem(BACKUP_DATE_KEY, dateStr); } catch (_) { }
      const bar = bundleList.querySelector('.backup-bar');
      if (bar) bar.remove();
    } catch (err) {
      if (err.name === 'NotAllowedError') {
        await clearDirectoryHandle();
        try { localStorage.removeItem(AUTO_BACKUP_INDICATOR_KEY); } catch (_) { }
        checkDailyBackup();
      }
    }
  }

  function scheduleAutoSave() {
    clearTimeout(autoSaveTimer);
    autoSaveTimer = setTimeout(async () => {
      let handle;
      try { handle = await loadDirectoryHandle(); } catch (_) { return; }
      if (!handle) return;
      let perm;
      try { perm = await handle.queryPermission({ mode: 'readwrite' }); } catch (_) { perm = 'denied'; }
      if (perm === 'granted') {
        await writeCSVToFolder(handle);
      } else {
        await clearDirectoryHandle();
        try { localStorage.removeItem(AUTO_BACKUP_INDICATOR_KEY); } catch (_) { }
        updateBackupBarUI();
        checkDailyBackup();
      }
    }, 2000);
  }

  async function chooseBackupFolder() {
    if (!window.showDirectoryPicker) {
      showToast('Auto-backup not supported in this browser');
      return;
    }
    try {
      const handle = await window.showDirectoryPicker({ mode: 'readwrite' });
      await saveDirectoryHandle(handle);
      try { localStorage.setItem(AUTO_BACKUP_INDICATOR_KEY, '1'); } catch (_) { }
      showToast('Auto-backup folder set — backups will save silently');
      updateBackupBarUI();
      await writeCSVToFolder(handle);
    } catch (err) {
      if (err.name !== 'AbortError') showToast('Could not set backup folder');
    }
  }

  function updateBackupBarUI() {
    const existing = bundleList.querySelector('.backup-bar');
    if (existing) existing.remove();
    if (!allRecords.length) return;

    let isAutoActive = false;
    try { isAutoActive = !!localStorage.getItem(AUTO_BACKUP_INDICATOR_KEY); } catch (_) { }

    if (isAutoActive) {
      const bar = document.createElement('div');
      bar.className = 'backup-bar backup-bar--active';
      bar.innerHTML = `
        <span class="backup-active-dot"></span>
        <span class="backup-active-label">Auto-backup active</span>
        <button class="backup-change-btn" title="Change backup folder">Change folder</button>`;
      bar.querySelector('.backup-change-btn').addEventListener('click', chooseBackupFolder);
      bundleList.prepend(bar);
      return;
    }

    // Show manual bar only if not already shown this session and not done today
    if (backupBarShown) return;
    const today = new Date().toLocaleDateString('en-CA');
    try { if (localStorage.getItem(BACKUP_DATE_KEY) === today) return; } catch (_) { }
    backupBarShown = true;
    const bar = document.createElement('div');
    bar.className = 'backup-bar';
    bar.innerHTML = `
      <button class="backup-save-btn">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
          <polyline points="7 10 12 15 17 10"/>
          <line x1="12" y1="15" x2="12" y2="3"/>
        </svg>
        Save Daily Backup
      </button>
      <button class="backup-folder-btn" title="Set auto-backup folder">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
          <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>
        </svg>
        Auto
      </button>
      <button class="backup-dismiss-btn" title="Dismiss">✕</button>`;
    bar.querySelector('.backup-save-btn').addEventListener('click', doBackupDownload);
    bar.querySelector('.backup-folder-btn').addEventListener('click', chooseBackupFolder);
    bar.querySelector('.backup-dismiss-btn').addEventListener('click', () => bar.remove());
    bundleList.prepend(bar);
  }

  function doBackupDownload() {
    const csv = buildAllRecordsCSV();
    if (!csv) return;
    const dateStr = new Date().toLocaleDateString('en-CA');
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
    const a = document.createElement('a');
    a.href = url; a.download = `meter-backup-${dateStr}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    try { localStorage.setItem(BACKUP_DATE_KEY, dateStr); } catch (_) { }
    const bar = bundleList.querySelector('.backup-bar');
    if (bar) bar.remove();
  }

  function checkDailyBackup() {
    updateBackupBarUI();
  }

  // ─── Email Provider Picker ────────────────────────
  const emailProviderModal = document.getElementById('email-provider-modal');

  function emailBundle(bundle) {
    pendingEmailBundle = bundle;
    // Show the Share button only when the browser supports file sharing
    const shareBtn = document.getElementById('email-provider-share');
    const testFile = new File([''], 'test.txt', { type: 'text/plain' });
    shareBtn.classList.toggle('hidden', !(navigator.canShare && navigator.canShare({ files: [testFile] })));
    emailProviderModal.classList.remove('hidden');
  }

  function doEmailWithProvider(provider) {
    const bundle = pendingEmailBundle;
    pendingEmailBundle = null;
    emailProviderModal.classList.add('hidden');
    if (!bundle) return;

    markAsSent(bundle.key);

    const safeName = bundle.bundleName.replace(/\s+/g, '_');
    const dateStr = new Date().toISOString().slice(0, 10);
    const csvContent = buildBundleCSV(bundle);
    const reportText = buildReportText(bundle);
    const csvFile = new File([csvContent], `bundle_${safeName}_${dateStr}.csv`, { type: 'text/csv' });
    const reportFile = new File([reportText], `report_${safeName}_${dateStr}.txt`, { type: 'text/plain' });
    const subject = `Meter Reading — Bundle ${bundle.bundleName} — ${new Date().toLocaleDateString('en-US')}`;
    const recipients = 'rovana.adjodha@metutilities.com,Sue.Vaillancourt@metutilities.com,Brandon.Bain@metutilities.com,mike.borneman@metutilities.com';

    if (provider === 'share') {
      navigator.share({ files: [csvFile, reportFile], title: subject, text: reportText }).catch(() => { });
      return;
    }

    // Download both files before opening the compose window
    [csvFile, reportFile].forEach(f => {
      const url = URL.createObjectURL(f);
      const a = document.createElement('a');
      a.href = url; a.download = f.name;
      a.click();
      URL.revokeObjectURL(url);
    });

    const to = encodeURIComponent(recipients);
    const su = encodeURIComponent(subject);
    const body = encodeURIComponent(reportText + '\n\n(See attached CSV for full reading data.)');

    const urls = {
      gmail: `https://mail.google.com/mail/?view=cm&to=${to}&su=${su}&body=${body}`,
      yahoo: `https://compose.mail.yahoo.com/?to=${to}&subject=${su}&body=${body}`,
      outlook: `https://outlook.live.com/mail/deeplink/compose?to=${to}&subject=${su}&body=${body}`,
    };

    setTimeout(() => {
      if (urls[provider]) {
        window.open(urls[provider], '_blank');
      } else {
        // Default mail client via mailto:
        window.location.href = `mailto:${recipients ? encodeURIComponent(recipients) : ''}?subject=${su}&body=${body}`;
      }
    }, 400);
  }

  document.getElementById('email-provider-cancel').addEventListener('click', () => {
    emailProviderModal.classList.add('hidden');
    pendingEmailBundle = null;
  });
  document.getElementById('email-provider-share').addEventListener('click', () => doEmailWithProvider('share'));
  ['gmail', 'yahoo', 'outlook', 'mailto'].forEach(p =>
    document.getElementById(`epm-${p}`).addEventListener('click', () => doEmailWithProvider(p))
  );

  // ─── PIN unlock handler ───────────────────────────
  const pinInput = document.getElementById('pin-input');
  const pinError = document.getElementById('pin-error');
  const pinUnlockBtn = document.getElementById('pin-unlock-btn');

  function tryUnlockPin() {
    if (pinInput.value === APP_PIN) {
      unlockPin();
      viewPin.classList.add('hidden');
      viewHome.classList.remove('hidden');
      renderHome();
      initAutoBackup();
      updateRecoverBar();
      checkDueDateWarnings();
      if (restored) showToast(`Session restored — ${allRecords.length} records loaded`);
    } else {
      pinError.classList.remove('hidden');
      pinInput.value = '';
      pinInput.focus();
    }
  }

  pinUnlockBtn.addEventListener('click', tryUnlockPin);
  pinInput.addEventListener('keydown', e => { if (e.key === 'Enter') tryUnlockPin(); });

  // ─── Fix Location Modal ───────────────────────────
  const fixLocModal = document.getElementById('fix-location-modal');
  const fixLocLat = document.getElementById('fix-loc-lat');
  const fixLocLng = document.getElementById('fix-loc-lng');

  window._openFixLocation = () => {
    if (!pendingFixPoint) return;
    fixLocLat.value = pendingFixPoint.lat;
    fixLocLng.value = pendingFixPoint.lng;
    pendingFixPoint.marker.closePopup();
    fixLocModal.classList.remove('hidden');
    fixLocLat.focus();
  };

  document.getElementById('fix-loc-cancel').addEventListener('click', () => {
    fixLocModal.classList.add('hidden');
  });

  document.getElementById('fix-loc-confirm').addEventListener('click', () => {
    const lat = parseFloat(fixLocLat.value);
    const lng = parseFloat(fixLocLng.value);
    if (isNaN(lat) || isNaN(lng)) { showToast('Enter valid coordinates.', true); return; }

    const { row, bundle, marker } = pendingFixPoint;
    const mapAddress = (row['Map Address'] || '').trim();
    const num = (row['#'] || '').trim().replace(/\.0$/, '');
    const street = (row['STREET'] || '').trim();
    const city = (row['City'] || '').trim();
    const cacheKey = mapAddress
      ? mapAddress.toLowerCase()
      : `${num} ${street},${city}`.trim().toLowerCase();

    const cache = loadGeoCache();
    cache[cacheKey] = { lat, lng };
    saveGeoCache(cache);

    // Update geocodedPoints in-place
    const pt = geocodedPoints.find(p => p.row === row);
    if (pt) { pt.lat = lat; pt.lng = lng; }

    // Move the marker
    marker.setLatLng([lat, lng]);
    const mapsUrl = `https://www.google.com/maps/dir/?api=1&destination=${lat},${lng}`;
    const addr = [row['#'], row['STREET']].filter(Boolean).join(' ');
    const loc = (row['LOC'] || '').trim();
    marker.setPopupContent(`<strong>${addr}</strong>${loc ? `<br><span style="font-size:0.85em;opacity:0.8">${loc}</span>` : ''}<br>${bundle.bundleName || ''}<br><a href="${mapsUrl}" target="_blank" rel="noopener" class="popup-nav-link">&#9654; Navigate</a><br><a href="#" class="popup-nav-link" onclick="window._openFixLocation();return false;">&#9999; Fix Location</a>`);

    fixLocModal.classList.add('hidden');
    pendingFixPoint = null;
    showToast('Location updated.');
  });

  // ─── Due Date Warning ─────────────────────────────
  const dueWarningModal = document.getElementById('due-warning-modal');
  const dueWarningDesc = document.getElementById('due-warning-desc');
  document.getElementById('due-warning-ok').addEventListener('click', () =>
    dueWarningModal.classList.add('hidden')
  );

  // Parse a YYYY-MM-DD string as local midnight (new Date(str) treats it as UTC).
  function parseDateLocal(str) {
    const [y, m, d] = str.split('-').map(Number);
    return new Date(y, m - 1, d);
  }

  function checkDueDateWarnings() {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);

    const warnings = bundles.filter(b => {
      if (sentBundles.has(b.key)) return false;
      const dueDate = (bundleRates[b.key]?.dueDate || '').trim();
      if (!dueDate) return false;
      const due = parseDateLocal(dueDate);
      return due <= tomorrow;
    });

    if (!warnings.length) return;

    const lines = warnings.map(b => {
      const dueDate = bundleRates[b.key]?.dueDate || '';
      const due = parseDateLocal(dueDate);
      const todayMs = new Date(); todayMs.setHours(0, 0, 0, 0);
      const status = due < todayMs ? 'Overdue' : due.getTime() === todayMs.getTime() ? 'Due today' : 'Due tomorrow';
      return `• ${b.bundleName} — ${status} (${dueDate})`;
    });

    dueWarningDesc.textContent = lines.join('\n');
    dueWarningModal.classList.remove('hidden');
  }

  // ─── Ding sound ───────────────────────────────────
  function playDing() {
    try {
      const ctx = new (window.AudioContext || window.webkitAudioContext)();
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.type = 'sine';
      osc.frequency.setValueAtTime(1047, ctx.currentTime);        // C6
      osc.frequency.setValueAtTime(1319, ctx.currentTime + 0.08); // E6
      gain.gain.setValueAtTime(0.4, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.5);
      osc.start(ctx.currentTime);
      osc.stop(ctx.currentTime + 0.5);
      osc.onended = () => ctx.close();
    } catch (_) { }
  }

  // ─── Toast ────────────────────────────────────────
  function showToast(msg, isError = false) {
    toast.textContent = msg;
    toast.classList.toggle('toast-error', isError);
    toast.classList.remove('hidden');
    toast.classList.add('show');
    clearTimeout(toast._t);
    toast._t = setTimeout(() => {
      toast.classList.remove('show');
      setTimeout(() => toast.classList.add('hidden'), 300);
    }, isError ? 4000 : 3000);
  }

  // ─── Helpers ──────────────────────────────────────
  function esc(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ─── Splash Animation ──────────────────────────────
  function runSplashAnimation(onDone) {
    [
      { id: 'spl-m', delay: 150 },
      { id: 'spl-e', delay: 350 },
      { id: 'spl-t', delay: 550 },
    ].forEach(({ id, delay }) => {
      setTimeout(() => { document.getElementById(id).classList.add('visible'); }, delay);
    });
    // Spin OUTS
    setTimeout(() => {
      const outs = document.getElementById('spl-outs');
      outs.classList.add('spin');
      // After spin (650ms) + 300ms pause → explode
      setTimeout(() => {
        outs.classList.remove('spin');
        void outs.offsetWidth; // flush animation state
        outs.classList.add('explode');
        setTimeout(onDone, 450); // explode (400ms) + 50ms buffer
      }, 650 + 300);
    }, 800);
  }

  // ─── Boot ─────────────────────────────────────────
  async function initAutoBackup() {
    if (!allRecords.length) { checkDailyBackup(); return; }

    let handle = null;
    try { handle = await loadDirectoryHandle(); } catch (_) { }
    if (!handle) { checkDailyBackup(); return; }

    // queryPermission is non-prompting — safe to call outside a user gesture
    let perm;
    try { perm = await handle.queryPermission({ mode: 'readwrite' }); } catch (_) { perm = 'denied'; }

    if (perm !== 'granted') {
      await clearDirectoryHandle();
      try { localStorage.removeItem(AUTO_BACKUP_INDICATOR_KEY); } catch (_) { }
      checkDailyBackup();
      return;
    }

    const today = new Date().toLocaleDateString('en-CA');
    let doneToday = false;
    try { doneToday = localStorage.getItem(BACKUP_DATE_KEY) === today; } catch (_) { }

    if (!doneToday) await writeCSVToFolder(handle);
    updateBackupBarUI();
  }

  function boot() {
    bundleRates = loadBundleRates();
    loadSentState();
    loadDeletedBundles();
    try {
      const saved = localStorage.getItem(READER_KEY);
      if (saved) setReaderName(saved);
    } catch (_) { }
    // Restore any previously saved records
    const restored = loadRecordsBackup();

    // Run splash animation, then transition to home
    runSplashAnimation(() => {
      viewSplash.classList.add('hidden');
      if (!isPinUnlocked()) {
        viewPin.classList.remove('hidden');
        setTimeout(() => pinInput.focus(), 100);
      } else {
        viewHome.classList.remove('hidden');
        renderHome();
        initAutoBackup();
        updateRecoverBar();
        checkDueDateWarnings();
        if (restored) showToast(`Session restored — ${allRecords.length} records loaded`);
      }
    });

    // Register service worker
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('./sw.js').catch(() => { });
    }
  }

  boot();

})();
