// ─── STATE ───────────────────────────────────────────────────────────────────
const state = {
    allData: [],
    displayedData: [],
    isSearching: false,
    theme: localStorage.getItem('theme') || 'light',
    currentLot: null,
    refreshInterval: null,
    lastUpdated: null,
    renderedKeys: new Set()   // track which batch+date keys are already on screen
};

const API_URL = 'https://script.google.com/macros/s/AKfycbyZonxQXddjrdO8WRnCBinrIAS-QyAhCo5a15yYiO4Omz_K8XNm_abnQZJAnIEuOhk-/exec';
const REFRESH_EVERY_MS = 5000; // 5 seconds

// ─── DOM REFERENCES ──────────────────────────────────────────────────────────
const el = {
    app:            document.body,
    menuBtn:        document.getElementById('menuBtn'),
    dropdown:       document.getElementById('dropdownMenu'),
    themeToggle:    document.getElementById('themeToggle'),
    aboutBtn:       document.getElementById('aboutBtn'),
    refreshBtn:     document.getElementById('refreshBtn'),
    searchInput:    document.getElementById('searchInput'),
    clearSearch:    document.getElementById('clearSearch'),
    dataStatus:     document.getElementById('dataStatus'),
    resultCount:    document.getElementById('resultCount'),
    refreshStatus:  document.getElementById('refreshStatus'),
    resultsList:    document.getElementById('resultsList'),
    loader:         document.getElementById('loader'),
    errorState:     document.getElementById('errorState'),
    emptyState:     document.getElementById('emptyState'),
    retryBtn:       document.getElementById('retryBtn'),
    detailsModal:   document.getElementById('detailsModal'),
    closeModalBtn:  document.getElementById('closeModalBtn'),
    shareResultBtn: document.getElementById('shareResultBtn'),
    aboutModal:     document.getElementById('aboutModal'),
    closeAboutBtn:  document.getElementById('closeAboutBtn'),
    repBuyer:       document.getElementById('repBuyer'),
    repDate:        document.getElementById('repDate'),
    repOrder:       document.getElementById('repOrder'),
    repShift:       document.getElementById('repShift'),
    repBatch:       document.getElementById('repBatch'),
    repReport:      document.getElementById('repReport'),
    repColor:       document.getElementById('repColor'),
    repFabType:     document.getElementById('repFabType'),
    repComposition: document.getElementById('repComposition'),
    repQty:         document.getElementById('repQty'),
    repReqGsm:      document.getElementById('repReqGsm'),
    repGsmResult:   document.getElementById('repGsmResult'),
    repReqDia:      document.getElementById('repReqDia'),
    repFinishDia:   document.getElementById('repFinishDia'),
    repLength:      document.getElementById('repLength'),
    repWidth:       document.getElementById('repWidth'),
    repTwisting:    document.getElementById('repTwisting'),
    repRubbingDry:  document.getElementById('repRubbingDry'),
    repRubbingWet:  document.getElementById('repRubbingWet'),
    repPh:          document.getElementById('repPh'),
    repDrying:      document.getElementById('repDrying'),
    repOthers:      document.getElementById('repOthers'),
};

// ─── INIT ────────────────────────────────────────────────────────────────────
function init() {
    applyTheme();
    attachEvents();
    fetchData();
    startAutoRefresh();
}

// ─── AUTO REFRESH ────────────────────────────────────────────────────────────
function startAutoRefresh() {
    if (state.refreshInterval) clearInterval(state.refreshInterval);
    state.refreshInterval = setInterval(() => fetchData(true), REFRESH_EVERY_MS);
}

// ─── EVENTS ──────────────────────────────────────────────────────────────────
function attachEvents() {
    el.menuBtn.addEventListener('click', e => { e.stopPropagation(); el.dropdown.classList.toggle('hidden'); });
    document.addEventListener('click', e => {
        if (!el.menuBtn.contains(e.target) && !el.dropdown.contains(e.target))
            el.dropdown.classList.add('hidden');
    });
    el.themeToggle.addEventListener('click', () => { toggleTheme(); el.dropdown.classList.add('hidden'); });
    el.aboutBtn.addEventListener('click', () => { el.dropdown.classList.add('hidden'); openModal(el.aboutModal); });
    el.refreshBtn.addEventListener('click', () => { el.dropdown.classList.add('hidden'); fetchData(); });
    el.searchInput.addEventListener('input', handleSearch);
    el.clearSearch.addEventListener('click', () => { el.searchInput.value = ''; handleSearch({ target: el.searchInput }); });
    el.closeModalBtn.addEventListener('click', () => closeModal(el.detailsModal));
    el.closeAboutBtn.addEventListener('click', () => closeModal(el.aboutModal));
    el.retryBtn.addEventListener('click', fetchData);
    el.shareResultBtn.addEventListener('click', shareViaWhatsApp);
    window.addEventListener('click', e => { if (e.target.classList.contains('modal-overlay')) closeModal(e.target); });
}

// ─── THEME ───────────────────────────────────────────────────────────────────
function applyTheme() {
    const dark = state.theme === 'dark';
    el.app.classList.toggle('dark-mode', dark);
    el.app.classList.toggle('light-mode', !dark);
    el.themeToggle.innerHTML = dark ? '<i class="fas fa-sun"></i> Light Mode' : '<i class="fas fa-moon"></i> Dark Mode';
}
function toggleTheme() {
    state.theme = state.theme === 'dark' ? 'light' : 'dark';
    localStorage.setItem('theme', state.theme);
    applyTheme();
}

// ─── FETCH ───────────────────────────────────────────────────────────────────
async function fetchData(silent = false) {
    if (!silent) showState('loading');
    try {
        const res  = await fetch(API_URL + '?t=' + Date.now());
        const data = await res.json();
        if (data.ok && data.batches) {
            // tag each row with its original Excel/sheet order so we can
            // preserve serial order (just reversed) instead of sorting by lot number
            state.allData = data.batches.map((item, idx) => ({ ...item, _idx: idx }));
            state.lastUpdated = new Date();
            updateRefreshStatus();
            if (state.isSearching) {
                runSearch(el.searchInput.value.toLowerCase().trim());
            } else {
                processAndDisplay();
            }
        } else throw new Error('bad');
    } catch {
        if (!silent) showState('error');
    }
}

function updateRefreshStatus() {
    const t = state.lastUpdated;
    if (!t) return;
    const hms = [t.getHours(), t.getMinutes(), t.getSeconds()].map(n => String(n).padStart(2,'0')).join(':');
    el.refreshStatus.textContent = 'Live • ' + hms;
}

// ─── DATE PARSE ──────────────────────────────────────────────────────────────
function parseDate(str) {
    if (!str) return new Date(0);
    const p = String(str).split('-');
    if (p.length === 3) return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
    return new Date(0);
}

// ─── PERCENTAGE ──────────────────────────────────────────────────────────────
// Applies to Length, Width AND Twisting — if value is a small decimal (< 1)
// convert to %, otherwise show raw (for large entered numbers).
function toPercent(val) {
    if (val === '' || val === null || val === undefined) return '';
    const str = String(val).trim();
    if (str === '') return '';
    const num = parseFloat(str);
    if (isNaN(num)) return str;  // e.g. "480kg" → show as-is
    if (num !== 0 && Math.abs(num) < 1) return (num * 100).toFixed(1) + '%';
    if (num === 0) return '0.0%';
    // Large numbers → show as-is but append % if it looks like a percentage
    // (Twisting entries like 2.0 stored as 2 → show as "2.0%")
    if (Number.isInteger(num) && Math.abs(num) <= 100) return num.toFixed(1) + '%';
    return str;
}

function hlClass(type, pctStr) {
    const n = parseFloat(pctStr);
    if (isNaN(n)) return '';
    if ((type === 'length' || type === 'width') && (n < -5 || n > 5)) return 'highlight-error';
    if (type === 'twisting' && n > 5) return 'highlight-error';
    return '';
}

// ─── PROCESS & DISPLAY ───────────────────────────────────────────────────────
function processAndDisplay() {
    const unique = [...new Set(state.allData.map(d => d.date))];
    unique.sort((a, b) => parseDate(b) - parseDate(a));
    const top5 = new Set(unique.slice(0, 5));
    state.displayedData = state.allData
        .filter(i => top5.has(i.date))
        .sort((a, b) => {
            const dateDiff = parseDate(b.date) - parseDate(a.date);
            if (dateDiff !== 0) return dateDiff;
            // same date → keep sheet's serial order, just reversed (last row first)
            return b._idx - a._idx;
        });
    el.dataStatus.textContent = 'Showing last 5 days • Live';
    renderList();
}

// ─── SEARCH ──────────────────────────────────────────────────────────────────
function handleSearch(e) {
    const q = e.target.value.toLowerCase().trim();
    if (q) {
        state.isSearching = true;
        el.clearSearch.classList.remove('hidden');
        runSearch(q);
    } else {
        state.isSearching = false;
        el.clearSearch.classList.add('hidden');
        processAndDisplay();
    }
}

function runSearch(q) {
    state.displayedData = state.allData
        .filter(i =>
            String(i.batchNo   || '').toLowerCase().includes(q) ||
            String(i.buyerName || '').toLowerCase().includes(q) ||
            String(i.orderNo   || '').toLowerCase().includes(q) ||
            String(i.atlNumber || '').toLowerCase().includes(q)
        )
        .sort((a, b) => {
            const dateDiff = parseDate(b.date) - parseDate(a.date);
            if (dateDiff !== 0) return dateDiff;
            return b._idx - a._idx;
        });
    el.dataStatus.textContent = `Search: "${q}" (all dates)`;
    renderList();
}

// ─── SMART RENDER — no flicker ────────────────────────────────────────────────
function cardKey(item) {
    return item.batchNo + '|' + item.date;
}

function renderList() {
    el.resultCount.textContent = state.displayedData.length;
    if (!state.displayedData.length) { showState('empty'); return; }

    // Build a map of currently displayed card DOM nodes
    const existingMap = new Map();
    el.resultsList.querySelectorAll('.result-card[data-key]').forEach(card => {
        existingMap.set(card.dataset.key, card);
    });

    const newKeys = new Set(state.displayedData.map(cardKey));

    // Remove cards that are no longer in data
    existingMap.forEach((card, key) => {
        if (!newKeys.has(key)) card.remove();
    });

    // Build ordered list — insert new cards at top with animation, existing stay
    const orderedKeys = state.displayedData.map(cardKey);
    orderedKeys.forEach((key, idx) => {
        if (!existingMap.has(key)) {
            // New entry — create card with animation class
            const item = state.displayedData[idx];
            const card = buildCard(item, true);
            card.dataset.key = key;
            // Insert at correct position
            const refNode = el.resultsList.children[idx] || null;
            el.resultsList.insertBefore(card, refNode);
            // Remove animation class after it completes so future refreshes don't re-trigger
            setTimeout(() => card.classList.remove('new-entry'), 600);
        }
    });

    showState('list');
}

function buildCard(item, isNew = false) {
    const card = document.createElement('div');
    card.className = 'result-card' + (isNew ? ' new-entry' : '');

    const lenStr = toPercent(item.length);
    const widStr = toPercent(item.width);
    const twStr  = toPercent(item.twisting);

    const batchTxt = item.batchNo  || '';
    const orderTxt = item.orderNo  || item.atlNumber || '';
    const titleHtml = `Lot ${batchTxt}${orderTxt ? `<span style="color:var(--text-muted);font-weight:500"> &nbsp;&nbsp; ATL ${orderTxt}</span>` : ''}`;

    card.innerHTML = `
      <div class="card-header">
        <span class="batch-no">${titleHtml}</span>
        <span class="date-badge"><i class="far fa-calendar-alt"></i> ${item.date || ''}</span>
      </div>
      <div class="card-body">
        <div class="data-item">
          <span class="data-label">Buyer</span>
          <span class="data-value">${item.buyerName || '—'}</span>
        </div>
        <div class="data-item">
          <span class="data-label">Length</span>
          <span class="data-value ${hlClass('length', lenStr)}">${lenStr || '—'}</span>
        </div>
        <div class="data-item">
          <span class="data-label">Width</span>
          <span class="data-value ${hlClass('width', widStr)}">${widStr || '—'}</span>
        </div>
        <div class="data-item">
          <span class="data-label">Twisting</span>
          <span class="data-value ${hlClass('twisting', twStr)}">${twStr || '—'}</span>
        </div>
      </div>`;

    card.addEventListener('click', () => openDetails(item));
    return card;
}

// ─── DETAILS MODAL ───────────────────────────────────────────────────────────
// Tries every possible API key name variant (camelCase, lowercase, etc.)
function pick(item, ...keys) {
    for (const k of keys) {
        const v = item[k];
        if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
    }
    return '';
}

function openDetails(item) {
    state.currentLot = item;

    // ── Excel column → possible API JSON keys ──────────────────────────────
    el.repBuyer.textContent       = pick(item, 'buyerName','buyer','Buyer');
    el.repDate.textContent        = pick(item, 'date','Date');
    el.repOrder.textContent       = pick(item, 'orderNo','orderNumber','atlNumber','OrderNo','Order No');
    el.repShift.textContent       = pick(item, 'shift','Shift');
    el.repBatch.textContent       = pick(item, 'batchNo','batchNumber','Batch No','BatchNo');
    el.repReport.textContent      = pick(item, 'reportNumber','reportNo','Report number','reportNum');
    el.repColor.textContent       = pick(item, 'colour','color','Colour','Color');
    el.repFabType.textContent     = pick(item, 'fabType','fabricType','fab.Type','Fab. Type','fabtype');
    el.repComposition.textContent = pick(item, 'composition','Composition');
    el.repQty.textContent         = pick(item, 'qty','quantity','Qty','Quantity');
    el.repReqGsm.textContent      = pick(item, 'rGsm','reqGsm','r_gsm','R. GSM','r.gsm','gsm');
    el.repGsmResult.textContent   = pick(item, 'gsmResult','GSM Result','gsmresult','foundGsm');
    el.repReqDia.textContent      = pick(item, 'reqDia','Req. Dia','reqdia','req_dia');
    el.repFinishDia.textContent   = pick(item, 'fDia','finishDia','F. Dia','fdia','f_dia');
    el.repDrying.textContent      = pick(item, 'drying','Drying');
    el.repPh.textContent          = pick(item, 'ph','pH','phValue','ph_value');
    el.repRubbingDry.textContent  = pick(item, 'dryRubbingResult','Dry Rubbing Result','dryRubbing','dry');
    el.repRubbingWet.textContent  = pick(item, 'wetRubbingResult','Wet Rubbing Result','wetRubbing','wet');

    const others = pick(item, 'others','Others');
    const info   = pick(item, 'info','INFO','Info');
    el.repOthers.textContent = [others, info].filter(Boolean).join(' | ');

    // ── Dimensional values ─────────────────────────────────────────────────
    const lenStr = toPercent(item.length);
    const widStr = toPercent(item.width);
    const twStr  = toPercent(item.twisting);

    el.repLength.textContent   = lenStr || '—';
    el.repWidth.textContent    = widStr || '—';
    el.repTwisting.textContent = twStr  || '—';

    el.repLength.className   = `text-center highlightable ${hlClass('length',   lenStr)}`;
    el.repWidth.className    = `text-center highlightable ${hlClass('width',    widStr)}`;
    el.repTwisting.className = `text-center highlightable ${hlClass('twisting', twStr)}`;

    openModal(el.detailsModal);
}

// ─── WHATSAPP TEXT SHARE ─────────────────────────────────────────────────────
function shareViaWhatsApp() {
    if (!state.currentLot) return;
    const i  = state.currentLot;
    const orderNo = pick(i, 'orderNo','atlNumber','orderNumber');
    const color   = pick(i, 'colour','color');
    const gsm     = pick(i, 'rGsm','reqGsm','gsm');
    const lenStr  = toPercent(i.length);
    const widStr  = toPercent(i.width);
    const twStr   = toPercent(i.twisting);

    let msg = `*Lab Test Result*\n\n`;
    msg += `*Lot / Batch No:* ${i.batchNo || ''}`;
    if (orderNo) msg += ` | Order: ${orderNo}`;
    if (color)   msg += ` | Color: ${color}`;
    if (gsm)     msg += ` | GSM: ${gsm}`;
    msg += `\n*Buyer:* ${i.buyerName || ''}\n`;
    msg += `*Date:* ${i.date || ''}\n\n`;
    msg += `*Dimensional Stability*\n`;
    msg += `Length: ${lenStr  || 'N/A'}\n`;
    msg += `Width: ${widStr   || 'N/A'}\n`;
    msg += `Twisting: ${twStr || 'N/A'}\n`;
    msg += `\n_Generated by Lab Test App_`;

    window.open('https://wa.me/?text=' + encodeURIComponent(msg), '_blank');
}

// ─── MODAL HELPERS ───────────────────────────────────────────────────────────
function openModal(m) {
    m.classList.remove('hidden');
    requestAnimationFrame(() => requestAnimationFrame(() => m.classList.add('active')));
    document.body.style.overflow = 'hidden';
}
function closeModal(m) {
    m.classList.remove('active');
    setTimeout(() => { m.classList.add('hidden'); document.body.style.overflow = ''; }, 280);
}

// ─── STATE DISPLAY ───────────────────────────────────────────────────────────
function showState(s) {
    el.loader.classList.add('hidden');
    el.errorState.classList.add('hidden');
    el.emptyState.classList.add('hidden');
    el.resultsList.classList.add('hidden');
    if (s === 'loading') el.loader.classList.remove('hidden');
    if (s === 'error')   el.errorState.classList.remove('hidden');
    if (s === 'empty')   el.emptyState.classList.remove('hidden');
    if (s === 'list')    el.resultsList.classList.remove('hidden');
}

// ─── START ───────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', init);
