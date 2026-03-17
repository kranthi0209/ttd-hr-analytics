/* ================================================================
   TTD HR Analytics Dashboard – Main Application Script
   ================================================================ */

// ── Global State ──
let RAW_DATA = [];
let SERVICE_DATA = [];
let FILTERED_DATA = [];
let OCCUPIED_DATA = [];
let ACTIVE_EMP_TYPES = new Set(); // filled after data loads; empty = no type filter yet

let dashCharts = {};
let pivotChartInstance = null;
let currentPage = 1;
let sortCol = null;
let sortDir = 'asc';
let visibleColumns = {};
let tableFilteredData = [];

const TABLE_COLUMNS = [
    { key: 'photo', label: 'Photo', show: false, noExport: true },
    { key: 'emp_id', label: 'Emp ID', show: true },
    { key: 'emp_name', label: 'Employee Name', show: true },
    { key: 'designation_name', label: 'Designation', show: true },
    { key: 'hod_name', label: 'Department (HOD)', show: true },
    { key: 'hos_name', label: 'Section Head', show: false },
    { key: 'section_name', label: 'Section', show: false },
    { key: 'employee_type', label: 'Employee Type', show: true },
    { key: 'post_status', label: 'Post Status', show: true },
    { key: 'community', label: 'Community', show: true },
    { key: 'sub_community', label: 'Sub Community', show: false },
    { key: 'caste', label: 'Caste', show: false },
    { key: 'gender', label: 'Gender', show: true },
    { key: 'work_location', label: 'Location', show: true },
    { key: 'dob_fmt', label: 'Date of Birth', show: false },
    { key: 'doj_fmt', label: 'Date of Joining', show: true },
    { key: 'dor_fmt', label: 'Retirement Date', show: true },
    { key: 'type_of_recruitment', label: 'Recruitment Type', show: false },
    { key: 'recruited_post', label: 'Recruited Post', show: false },
    { key: 'native_dist', label: 'Native District', show: false },
    { key: 'local_dist_study', label: 'Local/Study District', show: false },
    { key: 'joined_during', label: 'Joined During', show: false },
    { key: 'present_post_by', label: 'Present Post Posted In', show: false },
    { key: 'working_in_present_place_since_fmt', label: 'In Present Post Since', show: true },
    { key: 'service_years_current_post', label: 'Service (Yrs) Current Post', show: true },
    { key: 'mobile', label: 'Mobile', show: false }
];

// ── Utilities ──
function cleanStr(val) {
    if (val == null) return '';
    return String(val).replace(/_x000D_/g, '').replace(/[\t\r\n]+/g, '').trim();
}

function epochToDate(val) {
    if (!val || val === '' || val === 0) return null;
    const d = new Date(val);
    if (isNaN(d.getTime())) return null;
    return d;
}

function formatDate(d) {
    if (!d) return '—';
    const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    const day = d.getDate();
    const suffix = [11,12,13].includes(day) ? 'th' : day % 10 === 1 ? 'st' : day % 10 === 2 ? 'nd' : day % 10 === 3 ? 'rd' : 'th';
    return `${MONTHS[d.getMonth()]} ${day}${suffix} ${d.getFullYear()}`;
}

function yearsBetween(d1, d2) {
    if (!d1 || !d2) return null;
    return Math.max(0, ((d2 - d1) / (365.25 * 24 * 60 * 60 * 1000)));
}

function formatNum(n) {
    if (n == null) return '—';
    return n.toLocaleString('en-IN');
}

const CHART_COLORS = [
    '#2563eb', '#16a34a', '#dc2626', '#d97706', '#7c3aed',
    '#0d9488', '#ea580c', '#c026d3', '#0284c7', '#65a30d',
    '#db2777', '#ca8a04', '#4f46e5', '#0891b2', '#84cc16',
    '#e11d48', '#9333ea', '#f97316', '#06b6d4', '#22c55e'
];

function getColors(n) {
    const c = [];
    for (let i = 0; i < n; i++) c.push(CHART_COLORS[i % CHART_COLORS.length]);
    return c;
}

// ── Data Loading ──
async function loadData() {
    const msgEl = document.getElementById('loadingMsg');
    try {
        msgEl.textContent = 'Loading employee records...';
        const [empRes, svcRes] = await Promise.all([
            fetch('TTD_Employee_Full_List.json'),
            fetch('Regular_Employees_Service.json')
        ]);
        RAW_DATA = await empRes.json();
        SERVICE_DATA = await svcRes.json();

        msgEl.textContent = 'Processing data...';
        await new Promise(r => setTimeout(r, 50));
        processData();

        msgEl.textContent = 'Building dashboard...';
        await new Promise(r => setTimeout(r, 50));
        initApp();

        document.getElementById('loadingScreen').classList.add('hidden');
    } catch (e) {
        msgEl.textContent = 'Error loading data: ' + e.message;
        console.error(e);
    }
}

function processData() {
    const now = new Date();
    RAW_DATA.forEach(r => {
        // Clean string fields
        for (const k of Object.keys(r)) {
            if (typeof r[k] === 'string') r[k] = cleanStr(r[k]);
        }
        // Parse dates
        r._dob = epochToDate(r.dob);
        r._doj = epochToDate(r.doj_ttd);
        r._dor = epochToDate(r.dor);
        r._presentPostSince = epochToDate(r.working_in_present_place_since);
        r._cadreSince = epochToDate(r.working_in_present_cadre_since);

        r.dob_fmt = formatDate(r._dob);
        r.doj_fmt = formatDate(r._doj);
        r.dor_fmt = formatDate(r._dor);
        r.working_in_present_place_since_fmt = formatDate(r._presentPostSince);

        // Compute derived fields
        r.age = r._dob ? Math.floor(yearsBetween(r._dob, now)) : null;
        r.total_service_years = r._doj ? parseFloat(yearsBetween(r._doj, now).toFixed(1)) : null;
        r.service_years_current_post = r._presentPostSince ? parseFloat(yearsBetween(r._presentPostSince, now).toFixed(1)) : null;
        r.retirement_year = r._dor ? r._dor.getFullYear() : null;

        const _ps = String(r.post_status || '').toLowerCase().trim();
        r.post_status = _ps;          // normalise in-place so filters match correctly
        r.is_occupied = _ps === 'occupied';
        
        // fix work_location
        if (!r.work_location || r.work_location === 'None' || r.work_location === 'null') {
            r.work_location = 'Not Specified';
        }
    });

    // Process service data
    SERVICE_DATA.forEach(s => {
        s._from = epochToDate(s.service_from_date);
        s._to = epochToDate(s.service_to_date);
        if (typeof s.service_office === 'string') s.service_office = cleanStr(s.service_office);
        if (typeof s.designation_name === 'string') s.designation_name = cleanStr(s.designation_name);
    });
}

// ── Initialize App ──
function initApp() {
    setupNavigation();
    initEmployeeTypeSwitches(); // must run before filters so ACTIVE_EMP_TYPES is populated
    updateEmpTypeWidget();      // static widget — built from full RAW_DATA
    populateFilters();
    applyDashboardFilters();
    initDataTable();
    initPivotTable();
    initPivotChart();
    initEmployeeModal();
}

// ── Employee Type Strength Widget ──
const ETW_ACCENTS = [
    '#7b1c00', '#1565c0', '#2e7d32', '#c62828',
    '#6a1b9a', '#e65100', '#00695c', '#ad1457',
    '#4527a0', '#558b2f'
];

function updateEmpTypeWidget() {
    const wrap = document.getElementById('empTypeStrCards');
    if (!wrap) return;

    const REGULAR_TYPE = 'Regular TTD Employees';
    const R = 36, CIRC = +(2 * Math.PI * R).toFixed(2); // SVG ring constants

    // Group full RAW_DATA by employee_type
    const typeMap = {};
    RAW_DATA.forEach(r => {
        const t = r.employee_type || 'Unspecified';
        if (!typeMap[t]) typeMap[t] = { sanctioned: 0, working: 0, vacant: 0 };
        typeMap[t].sanctioned++;
        if (r.post_status === 'occupied') typeMap[t].working++;
        if (r.post_status === 'vacant')   typeMap[t].vacant++;
    });

    // Regular TTD first, then alphabetically
    const types = Object.keys(typeMap).sort((a, b) => {
        if (a === REGULAR_TYPE) return -1;
        if (b === REGULAR_TYPE) return  1;
        return a.localeCompare(b);
    });

    let html = '';
    types.forEach((t, i) => {
        const d        = typeMap[t];
        const isReg    = t === REGULAR_TYPE;
        const isActive = ACTIVE_EMP_TYPES.has(t);
        const accent   = ETW_ACCENTS[i % ETW_ACCENTS.length];
        const dimCls   = isActive ? '' : ' str-card-dim';

        if (isReg) {
            const pct     = d.sanctioned > 0 ? (d.working / d.sanctioned * 100) : 0;
            const pctTxt  = pct.toFixed(1);
            const filled  = +(CIRC * pct / 100).toFixed(2);
            html += `
            <div class="str-card str-featured${dimCls}" style="--sa:${accent}">
                <div class="str-feat-gauge">
                    <svg width="88" height="88" viewBox="0 0 88 88">
                        <circle cx="44" cy="44" r="${R}" fill="none"
                            stroke="rgba(200,134,10,0.15)" stroke-width="8"/>
                        <circle cx="44" cy="44" r="${R}" fill="none"
                            stroke="${accent}" stroke-width="8"
                            stroke-linecap="round"
                            stroke-dasharray="0 ${CIRC}"
                            data-dash="${filled} ${CIRC}"
                            transform="rotate(-90 44 44)"
                            style="transition:stroke-dasharray 1.1s cubic-bezier(.4,0,.2,1)"/>
                        <text x="44" y="41" text-anchor="middle"
                            fill="${accent}" font-size="13" font-weight="700"
                            font-family="JetBrains Mono,monospace">${pctTxt}%</text>
                        <text x="44" y="54" text-anchor="middle"
                            fill="#9a7150" font-size="7.5"
                            font-family="sans-serif" letter-spacing="0.5">OCCUPANCY</text>
                    </svg>
                </div>
                <div class="str-feat-info">
                    <div class="str-feat-name">${t}</div>
                    <div class="str-feat-stats">
                        <div class="str-stat">
                            <span class="str-num str-sanctioned">${formatNum(d.sanctioned)}</span>
                            <span class="str-lbl">Sanctioned</span>
                        </div>
                        <div class="str-stat-sep"></div>
                        <div class="str-stat">
                            <span class="str-num str-working">${formatNum(d.working)}</span>
                            <span class="str-lbl">Working</span>
                        </div>
                        <div class="str-stat-sep"></div>
                        <div class="str-stat">
                            <span class="str-num str-vacant">${formatNum(d.vacant)}</span>
                            <span class="str-lbl">Vacant</span>
                        </div>
                    </div>
                </div>
            </div>`;
        } else {
            html += `
            <div class="str-card str-compact${dimCls}" style="--sa:${accent}">
                <div class="str-compact-dot"></div>
                <div class="str-compact-name">${t}</div>
                <div class="str-compact-num">${formatNum(d.working)}</div>
                <div class="str-compact-lbl">Working</div>
            </div>`;
        }
    });

    wrap.innerHTML = html;

    // Animate SVG rings after paint
    requestAnimationFrame(() => requestAnimationFrame(() => {
        wrap.querySelectorAll('circle[data-dash]').forEach(el => {
            el.setAttribute('stroke-dasharray', el.dataset.dash);
        });
    }));
}

// ── Employee Type Switches ──
function initEmployeeTypeSwitches() {
    const types = [...new Set(RAW_DATA.map(r => r.employee_type || 'Unknown').filter(Boolean))].sort();
    const DEFAULT_ON = 'Regular TTD Employees';

    ACTIVE_EMP_TYPES.clear();
    // Default: turn on "Regular TTD Employees"; if not present fall back to first type
    if (types.includes(DEFAULT_ON)) {
        ACTIVE_EMP_TYPES.add(DEFAULT_ON);
    } else if (types.length) {
        ACTIVE_EMP_TYPES.add(types[0]);
    }

    const container = document.getElementById('empTypeSwitches');
    if (!container) return;
    container.innerHTML = '';

    types.forEach(t => {
        const isOn = ACTIVE_EMP_TYPES.has(t);
        const label = document.createElement('label');
        label.className = 'emp-type-item';
        label.title = t;
        label.innerHTML = `
            <span class="emp-toggle ${isOn ? 'emp-on' : 'emp-off'}" data-etype="${t}"></span>
            <span class="emp-type-name">${t}</span>`;
        const tog = label.querySelector('.emp-toggle');
        tog.addEventListener('click', e => {
            e.preventDefault();
            const etype = tog.dataset.etype;
            if (ACTIVE_EMP_TYPES.has(etype)) {
                ACTIVE_EMP_TYPES.delete(etype);
                tog.classList.replace('emp-on', 'emp-off');
            } else {
                ACTIVE_EMP_TYPES.add(etype);
                tog.classList.replace('emp-off', 'emp-on');
            }
            refreshAllViews();
        });
        container.appendChild(label);
    });
}

function refreshAllViews() {
    updateEmpTypeWidget();
    applyDashboardFilters();
    applyTableFilters();
    // Re-generate pivot table if it has already been built
    const pivotOuter = document.getElementById('pivotTableOuter');
    if (pivotOuter && pivotOuter.style.display !== 'none') {
        generatePivotTable();
    }
    // Re-generate pivot chart if it already exists
    if (pivotChartInstance) {
        generatePivotChart();
    }
}

// ── Navigation ──
function setupNavigation() {
    const navItems = document.querySelectorAll('.nav-item');
    navItems.forEach(item => {
        item.addEventListener('click', () => {
            navItems.forEach(n => n.classList.remove('active'));
            item.classList.add('active');
            document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
            document.getElementById('page-' + item.dataset.page).classList.add('active');
            // Close mobile sidebar
            document.getElementById('sidebar').classList.remove('open');
            document.getElementById('sidebarOverlay').classList.remove('active');
        });
    });

    // Mobile menu
    document.getElementById('mobileMenuBtn').addEventListener('click', () => {
        document.getElementById('sidebar').classList.add('open');
        document.getElementById('sidebarOverlay').classList.add('active');
    });
    document.getElementById('sidebarOverlay').addEventListener('click', () => {
        document.getElementById('sidebar').classList.remove('open');
        document.getElementById('sidebarOverlay').classList.remove('active');
    });
}

// ── Populate Filter Dropdowns ──
function populateFilters() {
    // Dashboard unified filters
    initUnifiedFilters('dash', 'dashSelectFiltersBtn', 'dashSelectFiltersDropdown', 'dashFiltersGrid', DASH_DEFAULT_FIELDS, applyDashboardFilters);

    // Table unified filters
    initUnifiedFilters('table', 'tableSelectFiltersBtn', 'tableSelectFiltersDropdown', 'tableFiltersGrid', TABLE_DEFAULT_FIELDS, applyTableFilters);

    // Pivot Chart category/series dropdowns — all fields
    const catSel = document.getElementById('pcCategory');
    const serSel = document.getElementById('pcSeries');
    ALL_FILTER_FIELDS.forEach(f => {
        catSel.innerHTML += `<option value="${f.key}">${f.label}</option>`;
        serSel.innerHTML += `<option value="${f.key}">${f.label}</option>`;
    });
    // Also add age_group (derived), retirement_year
    ['retirement_year'].forEach(k => {
        const label = k === 'retirement_year' ? 'Retirement Year' : k;
        if (!catSel.querySelector(`option[value="${k}"]`)) catSel.innerHTML += `<option value="${k}">${label}</option>`;
        if (!serSel.querySelector(`option[value="${k}"]`)) serSel.innerHTML += `<option value="${k}">${label}</option>`;
    });

    // Range sliders
    setupRangeSlider('fDashServiceMin', 'fDashServiceMax', 'fDashServiceMinVal', 'fDashServiceMaxVal', applyDashboardFilters);
    setupRangeSlider('fTableServiceMin', 'fTableServiceMax', 'fTableServiceMinVal', 'fTableServiceMaxVal', applyTableFilters);

    // Clear All handlers
    document.getElementById('dashClearFilters').addEventListener('click', () => {
        clearUnifiedFilters('dash', 'dashFiltersGrid', applyDashboardFilters);
        document.getElementById('fDashServiceMin').value = 0;
        document.getElementById('fDashServiceMax').value = 40;
        document.getElementById('fDashServiceMinVal').textContent = '0';
        document.getElementById('fDashServiceMaxVal').textContent = '40';
    });

    document.getElementById('tableClearFilters').addEventListener('click', () => {
        clearUnifiedFilters('table', 'tableFiltersGrid', applyTableFilters);
        document.getElementById('fTableSearch').value = '';
        document.getElementById('fTableServiceMin').value = 0;
        document.getElementById('fTableServiceMax').value = 40;
        document.getElementById('fTableServiceMinVal').textContent = '0';
        document.getElementById('fTableServiceMaxVal').textContent = '40';
        document.getElementById('dtPsOccupied').checked = true;
        document.getElementById('dtPsVacant').checked = false;
        applyTableFilters();
    });

    document.getElementById('dtPsOccupied').addEventListener('change', applyTableFilters);
    document.getElementById('dtPsVacant').addEventListener('change', applyTableFilters);

    document.getElementById('fTableSearch').addEventListener('input', debounce(applyTableFilters, 300));
}

function setupRangeSlider(minId, maxId, minValId, maxValId, callback) {
    const minSlider = document.getElementById(minId);
    const maxSlider = document.getElementById(maxId);
    const minVal = document.getElementById(minValId);
    const maxVal = document.getElementById(maxValId);

    function update() {
        let lo = parseInt(minSlider.value);
        let hi = parseInt(maxSlider.value);
        if (lo > hi) { if (this === minSlider) maxSlider.value = lo; else minSlider.value = hi; }
        minVal.textContent = minSlider.value;
        maxVal.textContent = maxSlider.value;
        callback();
    }
    minSlider.addEventListener('input', update);
    maxSlider.addEventListener('input', update);
}

function debounce(fn, ms) {
    let t;
    return function(...args) { clearTimeout(t); t = setTimeout(() => fn.apply(this, args), ms); };
}

// ── Custom MultiSelect ──
const MULTISELECT_STATE = {};

function initMultiSelect(id, options, onChange) {
    const origSelect = document.getElementById(id);
    if (!origSelect) return;

    MULTISELECT_STATE[id] = new Set();

    const wrapper = document.createElement('div');
    wrapper.className = 'ms-wrapper';
    wrapper.id = 'ms-' + id;

    const trigger = document.createElement('div');
    trigger.className = 'ms-trigger';
    trigger.innerHTML = '<span class="ms-display placeholder">— All —</span><span class="ms-arrow">▾</span>';

    const dropdown = document.createElement('div');
    dropdown.className = 'ms-dropdown';
    dropdown.innerHTML =
        '<div class="ms-search-wrap"><input class="ms-search" placeholder="Search..."></div>' +
        '<div class="ms-actions"><button class="ms-select-all">✔ Select All</button><button class="ms-deselect-all">✕ Clear</button></div>' +
        '<div class="ms-list"></div>';

    wrapper.appendChild(trigger);
    wrapper.appendChild(dropdown);
    origSelect.parentNode.insertBefore(wrapper, origSelect);
    origSelect.style.display = 'none';

    const listEl = dropdown.querySelector('.ms-list');
    const searchEl = dropdown.querySelector('.ms-search');
    const displayEl = trigger.querySelector('.ms-display');

    function buildList(filter) {
        filter = (filter || '').toLowerCase();
        listEl.innerHTML = '';
        options.forEach(opt => {
            if (filter && !String(opt).toLowerCase().includes(filter)) return;
            const item = document.createElement('label');
            const checked = MULTISELECT_STATE[id].has(opt);
            item.className = 'ms-option' + (checked ? ' checked' : '');
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = opt;
            cb.checked = checked;
            cb.addEventListener('change', () => {
                if (cb.checked) { MULTISELECT_STATE[id].add(opt); item.classList.add('checked'); }
                else { MULTISELECT_STATE[id].delete(opt); item.classList.remove('checked'); }
                updateDisplay();
                if (onChange) onChange();
            });
            item.appendChild(cb);
            item.appendChild(document.createTextNode(' ' + opt));
            listEl.appendChild(item);
        });
    }

    function updateDisplay() {
        const sel = [...MULTISELECT_STATE[id]];
        if (sel.length === 0) {
            displayEl.textContent = '— All —';
            displayEl.className = 'ms-display placeholder';
        } else if (sel.length === 1) {
            displayEl.textContent = sel[0];
            displayEl.className = 'ms-display';
        } else {
            displayEl.textContent = sel.length + ' selected';
            displayEl.className = 'ms-display';
        }
    }

    buildList();

    searchEl.addEventListener('input', () => buildList(searchEl.value));

    dropdown.querySelector('.ms-select-all').addEventListener('click', () => {
        options.forEach(o => MULTISELECT_STATE[id].add(o));
        buildList(searchEl.value);
        updateDisplay();
        if (onChange) onChange();
    });

    dropdown.querySelector('.ms-deselect-all').addEventListener('click', () => {
        MULTISELECT_STATE[id].clear();
        buildList(searchEl.value);
        updateDisplay();
        if (onChange) onChange();
    });

    trigger.addEventListener('click', e => {
        e.stopPropagation();
        const isOpen = wrapper.classList.contains('open');
        document.querySelectorAll('.ms-wrapper.open').forEach(w => w.classList.remove('open'));
        if (!isOpen) { wrapper.classList.add('open'); searchEl.value = ''; buildList(); searchEl.focus(); }
    });

    document.addEventListener('click', () => wrapper.classList.remove('open'));
    dropdown.addEventListener('click', e => e.stopPropagation());
}

function getMultiSelectValues(id) {
    return MULTISELECT_STATE[id] ? [...MULTISELECT_STATE[id]] : [];
}

// ── Dynamic Filter Provision ──
const DYNAMIC_FILTER_STATE = {}; // prefix -> { fieldKey -> { enabled, selected } }

function initDynamicFilterProvision(btnId, dropdownId, gridId, onChange, prefix) {
    const btn = document.getElementById(btnId);
    const dropdown = document.getElementById(dropdownId);
    const grid = document.getElementById(gridId);
    if (!btn || !dropdown || !grid) return;

    if (!DYNAMIC_FILTER_STATE[prefix]) DYNAMIC_FILTER_STATE[prefix] = {};

    const ALL_FILTER_FIELDS = [
        { key: 'hod_name', label: 'Department (HOD)' },
        { key: 'designation_name', label: 'Designation' },
        { key: 'community', label: 'Community' },
        { key: 'sub_community', label: 'Sub Community' },
        { key: 'caste', label: 'Caste' },
        { key: 'gender', label: 'Gender' },
        { key: 'work_location', label: 'Work Location' },
        { key: 'type_of_recruitment', label: 'Recruitment Type' },
        { key: 'joined_during', label: 'Joined in TTD During' },
        { key: 'present_post_by', label: 'Present Post Posted In' },
        { key: 'native_dist', label: 'Native District' },
        { key: 'local_dist_study', label: 'Local/Study District' },
        { key: 'post_status', label: 'Post Status' },
        { key: 'retirement_year', label: 'Retirement Year' },
        { key: 'section_name', label: 'Section' },
        { key: 'hos_name', label: 'Section Head' }
    ];

    ALL_FILTER_FIELDS.forEach(f => {
        DYNAMIC_FILTER_STATE[prefix][f.key] = { enabled: false, selected: new Set() };
        const lbl = document.createElement('label');
        lbl.style.cssText = 'display:flex;align-items:center;gap:8px;padding:4px 0;font-size:0.8rem;color:var(--text-secondary);cursor:pointer;user-select:none;';
        lbl.innerHTML = `<input type="checkbox" value="${f.key}"> ${f.label}`;
        const cb = lbl.querySelector('input');
        cb.addEventListener('change', () => {
            DYNAMIC_FILTER_STATE[prefix][f.key].enabled = cb.checked;
            renderDynamicFilterGrid(prefix, grid, onChange);
        });
        dropdown.appendChild(lbl);
    });

    btn.addEventListener('click', e => { e.stopPropagation(); dropdown.classList.toggle('active'); });
    document.addEventListener('click', () => dropdown.classList.remove('active'));
    dropdown.addEventListener('click', e => e.stopPropagation());
}

function renderDynamicFilterGrid(prefix, grid, onChange) {
    grid.innerHTML = '';
    let hasAny = false;
    const state = DYNAMIC_FILTER_STATE[prefix];
    if (!state) return;

    const dropdownId = prefix === 'dashExtra' ? 'dashFilterFieldDropdown' : 'tableFilterFieldDropdown';

    Object.entries(state).forEach(([fk, fState]) => {
        if (!fState.enabled) return;
        hasAny = true;
        const labelEl = document.querySelector(`#${dropdownId} input[value="${fk}"]`);
        const label = labelEl ? labelEl.closest('label').textContent.trim() : fk;
        const vals = [...new Set(getEmpTypeData().filter(r => r.is_occupied).map(r => String(r[fk] || '')).filter(v => v && v !== 'None' && v !== 'null'))].sort();

        const group = document.createElement('div');
        group.className = 'filter-group';
        group.innerHTML = `<label>${label}</label>`;
        grid.appendChild(group);

        initDynamicMultiSelect(`dynf-${prefix}-${fk}`, fk, vals, fState.selected, onChange, group);
    });

    grid.style.display = hasAny ? '' : 'none';
}

function initDynamicMultiSelect(wrapperId, fieldKey, options, selectedSet, onChange, parentEl) {
    const wrapper = document.createElement('div');
    wrapper.className = 'ms-wrapper';
    wrapper.id = 'dynms-' + wrapperId;

    const trigger = document.createElement('div');
    trigger.className = 'ms-trigger';
    trigger.innerHTML = '<span class="ms-display placeholder">— All —</span><span class="ms-arrow">▾</span>';

    const dropdownEl = document.createElement('div');
    dropdownEl.className = 'ms-dropdown';
    dropdownEl.innerHTML =
        '<div class="ms-search-wrap"><input class="ms-search" placeholder="Search..."></div>' +
        '<div class="ms-actions"><button class="ms-select-all">✔ Select All</button><button class="ms-deselect-all">✕ Clear</button></div>' +
        '<div class="ms-list"></div>';

    wrapper.appendChild(trigger);
    wrapper.appendChild(dropdownEl);
    parentEl.appendChild(wrapper);

    const listEl = dropdownEl.querySelector('.ms-list');
    const searchEl = dropdownEl.querySelector('.ms-search');
    const displayEl = trigger.querySelector('.ms-display');

    function buildList(filter) {
        filter = (filter || '').toLowerCase();
        listEl.innerHTML = '';
        options.forEach(opt => {
            if (filter && !String(opt).toLowerCase().includes(filter)) return;
            const item = document.createElement('label');
            const checked = selectedSet.has(opt);
            item.className = 'ms-option' + (checked ? ' checked' : '');
            const cb = document.createElement('input');
            cb.type = 'checkbox'; cb.value = opt; cb.checked = checked;
            cb.addEventListener('change', () => {
                if (cb.checked) { selectedSet.add(opt); item.classList.add('checked'); }
                else { selectedSet.delete(opt); item.classList.remove('checked'); }
                updateDisplay();
                if (onChange) onChange();
            });
            item.appendChild(cb);
            item.appendChild(document.createTextNode(' ' + opt));
            listEl.appendChild(item);
        });
    }

    function updateDisplay() {
        const sel = [...selectedSet];
        if (sel.length === 0) { displayEl.textContent = '— All —'; displayEl.className = 'ms-display placeholder'; }
        else if (sel.length === 1) { displayEl.textContent = sel[0]; displayEl.className = 'ms-display'; }
        else { displayEl.textContent = sel.length + ' selected'; displayEl.className = 'ms-display'; }
    }

    buildList();
    searchEl.addEventListener('input', () => buildList(searchEl.value));
    dropdownEl.querySelector('.ms-select-all').addEventListener('click', () => {
        options.forEach(o => selectedSet.add(o));
        buildList(searchEl.value); updateDisplay(); if (onChange) onChange();
    });
    dropdownEl.querySelector('.ms-deselect-all').addEventListener('click', () => {
        selectedSet.clear();
        buildList(searchEl.value); updateDisplay(); if (onChange) onChange();
    });
    trigger.addEventListener('click', e => {
        e.stopPropagation();
        const isOpen = wrapper.classList.contains('open');
        document.querySelectorAll('.ms-wrapper.open').forEach(w => w.classList.remove('open'));
        if (!isOpen) { wrapper.classList.add('open'); searchEl.value = ''; buildList(); searchEl.focus(); }
    });
    document.addEventListener('click', () => wrapper.classList.remove('open'));
    dropdownEl.addEventListener('click', e => e.stopPropagation());
}

function getDynamicFilterValues(prefix) {
    const state = DYNAMIC_FILTER_STATE[prefix];
    if (!state) return {};
    const result = {};
    Object.entries(state).forEach(([fk, fState]) => {
        if (fState.enabled && fState.selected.size > 0) result[fk] = [...fState.selected];
    });
    return result;
}

function clearMultiSelect(id) {
    if (!MULTISELECT_STATE[id]) return;
    MULTISELECT_STATE[id].clear();
    const wrapper = document.getElementById('ms-' + id);
    if (wrapper) {
        const disp = wrapper.querySelector('.ms-display');
        if (disp) { disp.textContent = '— All —'; disp.className = 'ms-display placeholder'; }
        wrapper.querySelectorAll('input[type="checkbox"]').forEach(cb => { cb.checked = false; });
        wrapper.querySelectorAll('.ms-option').forEach(el => el.classList.remove('checked'));
    }
}

// ──────────────────────────────────────────────────────────────
//   UNIFIED CASCADING FILTER SYSTEM
// ──────────────────────────────────────────────────────────────

const ALL_FILTER_FIELDS = [
    { key: 'hod_name', label: 'Department (HOD)' },
    { key: 'designation_name', label: 'Designation' },
    { key: 'community', label: 'Community' },
    { key: 'sub_community', label: 'Sub Community' },
    { key: 'caste', label: 'Caste' },
    { key: 'gender', label: 'Gender' },
    { key: 'work_location', label: 'Work Location' },
    { key: 'type_of_recruitment', label: 'Recruitment Type' },
    { key: 'joined_during', label: 'Joined in TTD During' },
    { key: 'present_post_by', label: 'Present Post Posted In' },
    { key: 'native_dist', label: 'Native District' },
    { key: 'local_dist_study', label: 'Local/Study District' },
    { key: 'employee_type', label: 'Employee Type' },
    { key: 'post_status', label: 'Post Status' },
    { key: 'retirement_year', label: 'Retirement Year' },
    { key: 'section_name', label: 'Section' },
    { key: 'hos_name', label: 'Section Head' },
    { key: 'recruited_post', label: 'Recruited Post' },
    // Date range fields
    { key: '_doj', label: 'Date of Joining', type: 'date', rawDateKey: '_doj' },
    { key: '_dob', label: 'Date of Birth', type: 'date', rawDateKey: '_dob' },
    { key: '_dor', label: 'Date of Retirement', type: 'date', rawDateKey: '_dor' },
    { key: '_presentPostSince', label: 'In Present Post Since', type: 'date', rawDateKey: '_presentPostSince' }
];

const UNIFIED_FILTERS = {};

const DASH_DEFAULT_FIELDS = ['hod_name', 'community', 'gender', 'work_location', 'type_of_recruitment', 'joined_during', 'present_post_by'];
const TABLE_DEFAULT_FIELDS = ['hod_name', 'community', 'gender', 'work_location', 'designation_name', 'post_status'];
const PC_DEFAULT_FIELDS = ['hod_name', 'community', 'gender', 'work_location'];

function initUnifiedFilters(prefix, btnId, dropdownId, gridId, defaultFields, onChange) {
    UNIFIED_FILTERS[prefix] = {};
    ALL_FILTER_FIELDS.forEach(f => {
        if (f.type === 'date') {
            UNIFIED_FILTERS[prefix][f.key] = { enabled: defaultFields.includes(f.key), type: 'date', from: null, to: null, rawDateKey: f.rawDateKey };
        } else {
            UNIFIED_FILTERS[prefix][f.key] = { enabled: defaultFields.includes(f.key), selected: new Set() };
        }
    });

    const dropdown = document.getElementById(dropdownId);
    if (!dropdown) return;
    ALL_FILTER_FIELDS.forEach(f => {
        const lbl = document.createElement('label');
        lbl.style.cssText = 'display:flex;align-items:center;gap:8px;padding:4px 0;font-size:0.8rem;color:var(--text-secondary);cursor:pointer;user-select:none;';
        const checked = defaultFields.includes(f.key);
        lbl.innerHTML = `<input type="checkbox" value="${f.key}" ${checked ? 'checked' : ''}> ${f.label}`;
        const cb = lbl.querySelector('input');
        cb.addEventListener('change', () => {
            UNIFIED_FILTERS[prefix][f.key].enabled = cb.checked;
            if (!cb.checked) {
                UNIFIED_FILTERS[prefix][f.key].selected.clear();
                if (onChange) onChange();
            }
            renderUnifiedFilterGrid(prefix, gridId, onChange);
        });
        dropdown.appendChild(lbl);
    });

    const btn = document.getElementById(btnId);
    if (btn) {
        btn.addEventListener('click', e => { e.stopPropagation(); dropdown.classList.toggle('active'); });
        document.addEventListener('click', () => dropdown.classList.remove('active'));
        dropdown.addEventListener('click', e => e.stopPropagation());
    }

    renderUnifiedFilterGrid(prefix, gridId, onChange);
}

// Returns RAW_DATA filtered by the active employee-type switches
function getEmpTypeData() {
    if (!ACTIVE_EMP_TYPES.size) return RAW_DATA;
    return RAW_DATA.filter(r => ACTIVE_EMP_TYPES.has(r.employee_type || ''));
}

function getBaseData(prefix) {
    const base = getEmpTypeData();
    if (prefix === 'dash' || prefix === 'pchart') return base.filter(r => r.is_occupied);
    return base;
}

function getCascadedOptions(prefix, forField) {
    const state = UNIFIED_FILTERS[prefix];
    if (!state) return [];
    const baseData = getBaseData(prefix);
    const filtered = baseData.filter(r => {
        for (const [fk, fState] of Object.entries(state)) {
            if (fk === forField || !fState.enabled) continue;
            if (fState.type === 'date') {
                const dv = r[fState.rawDateKey];
                if (fState.from && (!dv || dv < fState.from)) return false;
                if (fState.to && (!dv || dv > fState.to)) return false;
            } else if (fState.selected && fState.selected.size > 0) {
                if (!fState.selected.has(String(r[fk] || ''))) return false;
            }
        }
        return true;
    });
    return [...new Set(filtered.map(r => String(r[forField] || '')).filter(v => v && v !== 'None' && v !== 'null' && v !== 'undefined'))].sort();
}

function renderUnifiedFilterGrid(prefix, gridId, onChange) {
    const grid = document.getElementById(gridId);
    if (!grid) return;
    const state = UNIFIED_FILTERS[prefix];
    grid.innerHTML = '';
    ALL_FILTER_FIELDS.forEach(f => {
        if (!state || !state[f.key] || !state[f.key].enabled) return;
        const group = document.createElement('div');
        group.className = 'filter-group';
        group.id = `ufg-${prefix}-${f.key}`;
        group.innerHTML = `<label>${f.label}</label>`;
        grid.appendChild(group);
        if (f.type === 'date') {
            buildDateRangeFilter(prefix, f.key, group, onChange);
        } else {
            buildUnifiedMultiSelect(prefix, f.key, group, onChange);
        }
    });
}

function buildDateRangeFilter(prefix, fieldKey, parentEl, onChange) {
    const state = UNIFIED_FILTERS[prefix][fieldKey];
    const container = document.createElement('div');
    container.style.cssText = 'display:flex;flex-direction:column;gap:4px;';

    function makeRow(labelText, inputId) {
        const row = document.createElement('div');
        row.style.cssText = 'display:flex;align-items:center;gap:6px;';
        const span = document.createElement('span');
        span.textContent = labelText;
        span.style.cssText = 'font-size:0.7rem;color:var(--text-muted);min-width:26px;';
        const inp = document.createElement('input');
        inp.type = 'date';
        inp.id = inputId;
        inp.className = 'filter-input';
        inp.style.cssText = 'padding:4px 6px;font-size:0.78rem;flex:1;min-width:0;';
        row.appendChild(span);
        row.appendChild(inp);
        container.appendChild(row);
        return inp;
    }

    const fromInp = makeRow('From', `drf-${prefix}-${fieldKey}-from`);
    const toInp   = makeRow('To',   `drf-${prefix}-${fieldKey}-to`);

    if (state.from) fromInp.value = state.from.toISOString().split('T')[0];
    if (state.to)   toInp.value   = state.to.toISOString().split('T')[0];

    fromInp.addEventListener('change', () => {
        state.from = fromInp.value ? new Date(fromInp.value) : null;
        if (onChange) onChange();
    });
    toInp.addEventListener('change', () => {
        state.to = toInp.value ? new Date(toInp.value + 'T23:59:59') : null;
        if (onChange) onChange();
    });
    parentEl.appendChild(container);
}

function buildUnifiedMultiSelect(prefix, fieldKey, parentEl, onChange) {
    const state = UNIFIED_FILTERS[prefix][fieldKey];
    const options = getCascadedOptions(prefix, fieldKey);
    // Remove values no longer in cascaded options
    [...state.selected].forEach(v => { if (!options.includes(v)) state.selected.delete(v); });

    const wrapper = document.createElement('div');
    wrapper.className = 'ms-wrapper';
    wrapper.id = `ums-${prefix}-${fieldKey}`;

    const trigger = document.createElement('div');
    trigger.className = 'ms-trigger';
    trigger.innerHTML = '<span class="ms-display placeholder">— All —</span><span class="ms-arrow">▾</span>';

    const dropdownEl = document.createElement('div');
    dropdownEl.className = 'ms-dropdown';
    dropdownEl.innerHTML =
        '<div class="ms-search-wrap"><input class="ms-search" placeholder="Search..."></div>' +
        '<div class="ms-actions"><button class="ms-select-all">✔ Select All</button><button class="ms-deselect-all">✕ Clear</button></div>' +
        '<div class="ms-list"></div>';

    wrapper.appendChild(trigger);
    wrapper.appendChild(dropdownEl);
    parentEl.appendChild(wrapper);

    const listEl = dropdownEl.querySelector('.ms-list');
    const searchEl = dropdownEl.querySelector('.ms-search');
    const displayEl = trigger.querySelector('.ms-display');

    function buildList(filter) {
        filter = (filter || '').toLowerCase();
        listEl.innerHTML = '';
        options.forEach(opt => {
            if (filter && !opt.toLowerCase().includes(filter)) return;
            const item = document.createElement('label');
            const checked = state.selected.has(opt);
            item.className = 'ms-option' + (checked ? ' checked' : '');
            const cb = document.createElement('input');
            cb.type = 'checkbox'; cb.value = opt; cb.checked = checked;
            cb.addEventListener('change', () => {
                if (cb.checked) { state.selected.add(opt); item.classList.add('checked'); }
                else { state.selected.delete(opt); item.classList.remove('checked'); }
                updateDisplay();
                cascadeRebuild(prefix, fieldKey, onChange);
                if (onChange) onChange();
            });
            item.appendChild(cb);
            item.appendChild(document.createTextNode(' ' + opt));
            listEl.appendChild(item);
        });
    }

    function updateDisplay() {
        const sel = [...state.selected];
        if (sel.length === 0) { displayEl.textContent = '— All —'; displayEl.className = 'ms-display placeholder'; }
        else if (sel.length === 1) { displayEl.textContent = sel[0]; displayEl.className = 'ms-display'; }
        else { displayEl.textContent = sel.length + ' selected'; displayEl.className = 'ms-display'; }
    }

    buildList();
    updateDisplay();
    searchEl.addEventListener('input', () => buildList(searchEl.value));
    dropdownEl.querySelector('.ms-select-all').addEventListener('click', () => {
        options.forEach(o => state.selected.add(o));
        buildList(searchEl.value); updateDisplay();
        cascadeRebuild(prefix, fieldKey, onChange);
        if (onChange) onChange();
    });
    dropdownEl.querySelector('.ms-deselect-all').addEventListener('click', () => {
        state.selected.clear();
        buildList(searchEl.value); updateDisplay();
        cascadeRebuild(prefix, fieldKey, onChange);
        if (onChange) onChange();
    });
    trigger.addEventListener('click', e => {
        e.stopPropagation();
        const isOpen = wrapper.classList.contains('open');
        document.querySelectorAll('.ms-wrapper.open').forEach(w => w.classList.remove('open'));
        if (!isOpen) { wrapper.classList.add('open'); searchEl.value = ''; buildList(); searchEl.focus(); }
    });
    document.addEventListener('click', () => wrapper.classList.remove('open'));
    dropdownEl.addEventListener('click', e => e.stopPropagation());
}

function cascadeRebuild(prefix, changedField, onChange) {
    const state = UNIFIED_FILTERS[prefix];
    if (!state) return;
    ALL_FILTER_FIELDS.forEach(f => {
        if (f.type === 'date') return; // date filters don't cascade via multiselect
        if (f.key === changedField || !state[f.key] || !state[f.key].enabled) return;
        const parentEl = document.getElementById(`ufg-${prefix}-${f.key}`);
        if (!parentEl) return;
        const oldMs = parentEl.querySelector('.ms-wrapper');
        if (oldMs) oldMs.remove();
        buildUnifiedMultiSelect(prefix, f.key, parentEl, onChange);
    });
}

function getUnifiedFilterValues(prefix) {
    const state = UNIFIED_FILTERS[prefix];
    const result = {};
    if (!state) return result;
    ALL_FILTER_FIELDS.forEach(f => {
        if (!state[f.key] || !state[f.key].enabled) return;
        if (f.type === 'date') {
            const { from, to, rawDateKey } = state[f.key];
            if (from || to) result[f.key] = { type: 'date', from, to, rawDateKey };
        } else if (state[f.key].selected.size > 0) {
            result[f.key] = [...state[f.key].selected];
        }
    });
    return result;
}

// Helper: apply unified filter values to a record
function passesFilters(r, filters) {
    for (const [fk, val] of Object.entries(filters)) {
        if (!val) continue;
        if (val.type === 'date') {
            const dv = r[val.rawDateKey];
            if (val.from && (!dv || dv < val.from)) return false;
            if (val.to && (!dv || dv > val.to)) return false;
        } else if (Array.isArray(val) && val.length && !val.includes(String(r[fk] || ''))) {
            return false;
        }
    }
    return true;
}

function clearUnifiedFilters(prefix, gridId, onChange) {
    const state = UNIFIED_FILTERS[prefix];
    if (state) ALL_FILTER_FIELDS.forEach(f => {
        if (!state[f.key]) return;
        if (f.type === 'date') { state[f.key].from = null; state[f.key].to = null; }
        else state[f.key].selected.clear();
    });
    renderUnifiedFilterGrid(prefix, gridId, onChange);
    if (onChange) onChange();
}

// Navigate to data table, copying all active dashboard filters + the clicked field/value.
// postStatusFilter: 'occupied' | 'vacant' | null (null = no occupancy restriction)
// dateRange: { from: Date, to: Date } — when set, applies a date-range filter on fieldKey
function navigateToTableWithFilter(fieldKey, fieldValue, postStatusFilter = 'occupied', dateRange = null) {
    const tableNav = document.querySelector('[data-page="datatable"]');
    if (tableNav) tableNav.click();
    if (!UNIFIED_FILTERS['table'] || !UNIFIED_FILTERS['dash']) return;

    // Clear ALL existing table filters so previous clicks don't linger
    ALL_FILTER_FIELDS.forEach(f => {
        const s = UNIFIED_FILTERS['table'][f.key];
        if (!s) return;
        if (f.type === 'date') { s.from = null; s.to = null; s.enabled = false; }
        else { s.selected = new Set(); s.enabled = false; }
    });

    // Copy active dashboard filters across
    ALL_FILTER_FIELDS.forEach(f => {
        const dS = UNIFIED_FILTERS['dash'][f.key];
        const tS = UNIFIED_FILTERS['table'][f.key];
        if (!dS || !tS) return;
        if (f.type === 'date') {
            tS.from = dS.from; tS.to = dS.to;
            if (dS.from || dS.to) tS.enabled = true;
        } else {
            if (dS.selected && dS.selected.size > 0) {
                tS.enabled = true; tS.selected = new Set(dS.selected);
            }
        }
    });

    // Apply the clicked chart value (categorical or date range)
    if (fieldKey && UNIFIED_FILTERS['table'][fieldKey]) {
        UNIFIED_FILTERS['table'][fieldKey].enabled = true;
        if (dateRange) {
            UNIFIED_FILTERS['table'][fieldKey].from = dateRange.from;
            UNIFIED_FILTERS['table'][fieldKey].to   = dateRange.to;
        } else if (fieldValue !== null) {
            UNIFIED_FILTERS['table'][fieldKey].selected = new Set([String(fieldValue)]);
        }
    }

    // Apply occupancy filter (post_status)
    if (postStatusFilter !== null && UNIFIED_FILTERS['table']['post_status']) {
        UNIFIED_FILTERS['table']['post_status'].enabled = true;
        UNIFIED_FILTERS['table']['post_status'].selected = new Set([postStatusFilter]);
    }

    // Sync select-filters dropdown checkboxes
    const dd = document.getElementById('tableSelectFiltersDropdown');
    if (dd) {
        ALL_FILTER_FIELDS.forEach(f => {
            const cb = dd.querySelector(`input[value="${f.key}"]`);
            if (cb) cb.checked = !!(UNIFIED_FILTERS['table'][f.key] && UNIFIED_FILTERS['table'][f.key].enabled);
        });
    }

    renderUnifiedFilterGrid('table', 'tableFiltersGrid', applyTableFilters);
    applyTableFilters();

    // Scroll data table to top
    setTimeout(() => {
        const mainContent = document.querySelector('.main-content');
        if (mainContent) mainContent.scrollTop = 0;
        window.scrollTo(0, 0);
    }, 120);
}

// ── Filter Logic ──
function getFilteredOccupied(hods, comms, genders, locs, recruits, joineds, postedBys, svcMin, svcMax, extraFilters = {}) {
    return getEmpTypeData().filter(r => {
        if (!r.is_occupied) return false;
        if (hods.length && !hods.includes(r.hod_name)) return false;
        if (comms.length && !comms.includes(r.community)) return false;
        if (genders.length && !genders.includes(r.gender)) return false;
        if (locs.length && !locs.includes(r.work_location)) return false;
        if (recruits.length && !recruits.includes(r.type_of_recruitment)) return false;
        if (joineds.length && !joineds.includes(r.joined_during)) return false;
        if (postedBys.length && !postedBys.includes(r.present_post_by)) return false;
        if (r.service_years_current_post != null) {
            if (r.service_years_current_post < svcMin || r.service_years_current_post > svcMax) return false;
        }
        for (const [fk, vals] of Object.entries(extraFilters)) {
            if (vals.length && !vals.includes(String(r[fk] || ''))) return false;
        }
        return true;
    });
}

// ──────────────────────────────────────────────────────────────
//   DASHBOARD PAGE
// ──────────────────────────────────────────────────────────────

function applyDashboardFilters() {
    const filters = getUnifiedFilterValues('dash');
    const svcMin = parseInt(document.getElementById('fDashServiceMin').value) || 0;
    const svcMax = parseInt(document.getElementById('fDashServiceMax').value) || 40;

    const filteredAll = getEmpTypeData().filter(r => passesFilters(r, filters));

    OCCUPIED_DATA = getEmpTypeData().filter(r => {
        if (!r.is_occupied) return false;
        if (!passesFilters(r, filters)) return false;
        if (r.service_years_current_post != null) {
            if (r.service_years_current_post < svcMin || r.service_years_current_post > svcMax) return false;
        }
        return true;
    });

    updateKPIs(filteredAll, OCCUPIED_DATA);
    updateDashboardCharts(OCCUPIED_DATA, filteredAll);
}

function updateKPIs(allData, occupiedData) {
    const total    = allData.length;
    const occupied = allData.filter(r => r.post_status === 'occupied').length;
    const vacant   = allData.filter(r => r.post_status === 'vacant').length;
    const depts = new Set(allData.map(r => r.hod_name)).size;
    const desigs = new Set(occupiedData.map(r => r.designation_name)).size;
    const avgSvc = occupiedData.length > 0
        ? (occupiedData.reduce((s, r) => s + (r.total_service_years || 0), 0) / occupiedData.length).toFixed(1)
        : 0;

    const _etd = getEmpTypeData();
    const origTotal      = _etd.length;
    const origOccupied   = _etd.filter(r => r.post_status === 'occupied').length;
    const origVacant     = _etd.filter(r => r.post_status === 'vacant').length;
    const origDepts      = new Set(_etd.map(r => r.hod_name)).size;
    const origOccupiedData = _etd.filter(r => r.post_status === 'occupied');
    const origDesigs = new Set(origOccupiedData.map(r => r.designation_name)).size;
    const origAvgSvc = origOccupiedData.length > 0
        ? (origOccupiedData.reduce((s, r) => s + (r.total_service_years || 0), 0) / origOccupiedData.length).toFixed(1)
        : 0;

    const isFiltered = total !== origTotal;

    function setKpi(id, filtVal, origVal) {
        document.getElementById(id).textContent = formatNum(filtVal);
        const origEl = document.getElementById(id + 'Orig');
        if (origEl) {
            if (isFiltered) {
                origEl.textContent = `of ${formatNum(origVal)} total`;
                origEl.style.display = 'block';
            } else {
                origEl.textContent = '';
                origEl.style.display = 'none';
            }
        }
    }

    setKpi('kpiTotal', total, origTotal);
    setKpi('kpiOccupied', occupied, origOccupied);
    setKpi('kpiVacant', vacant, origVacant);
    setKpi('kpiDepts', depts, origDepts);
    setKpi('kpiDesignations', desigs, origDesigs);

    document.getElementById('kpiAvgService').textContent = avgSvc;
    const avgOrigEl = document.getElementById('kpiAvgServiceOrig');
    if (avgOrigEl) {
        if (isFiltered) {
            avgOrigEl.textContent = `of ${origAvgSvc} overall`;
            avgOrigEl.style.display = 'block';
        } else {
            avgOrigEl.textContent = '';
            avgOrigEl.style.display = 'none';
        }
    }
}

function updateDashboardCharts(data, allData) {
    // Destroy existing charts
    Object.values(dashCharts).forEach(c => c.destroy());
    dashCharts = {};

    // 1. Department Strength (stacked horizontal bar — occupied vs vacant)
    dashCharts.dept = createDeptStackedChart('chartDeptStrength', allData, 15);

    // 2. Community (doughnut)
    dashCharts.community = createCountChart('chartCommunity', data, 'community', 'doughnut');

    // 3. Gender (pie)
    dashCharts.gender = createCountChart('chartGender', data, 'gender', 'pie');

    // 4. Recruitment Type (bar)
    dashCharts.recruitment = createCountChart('chartRecruitment', data, 'type_of_recruitment', 'bar');

    // 5. Joined During (bar)
    dashCharts.joined = createCountChart('chartJoined', data, 'joined_during', 'bar');

    // 6. Age Distribution (histogram — clickable)
    const ageBuckets = {};
    data.forEach(r => {
        if (r.age) {
            const bucket = Math.floor(r.age / 5) * 5;
            const label = `${bucket}-${bucket + 4}`;
            ageBuckets[label] = (ageBuckets[label] || 0) + 1;
        }
    });
    const ageSorted = Object.entries(ageBuckets).sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
    const ageLabels = ageSorted.map(e => e[0]);
    const ageValues = ageSorted.map(e => e[1]);
    const ageTotalCount = ageValues.reduce((a, b) => a + b, 0);
    const ageNow = new Date();
    dashCharts.age = new Chart(document.getElementById('chartAge'), {
        type: 'bar',
        data: {
            labels: ageLabels,
            datasets: [{ label: 'Employees', data: ageValues, backgroundColor: CHART_COLORS[0] + '99', borderColor: CHART_COLORS[0], borderWidth: 1, borderRadius: 4 }]
        },
        options: {
            ...chartOptions('Age Distribution'),
            onClick: (_ev, els) => {
                if (!els.length) return;
                const lbl = ageLabels[els[0].index];
                const parts = lbl.split('-');
                const minAge = parseInt(parts[0]), maxAge = parseInt(parts[1]);
                const toDate   = new Date(ageNow.getFullYear() - minAge, ageNow.getMonth(), ageNow.getDate());
                const fromDate = new Date(ageNow.getFullYear() - maxAge - 1, ageNow.getMonth(), ageNow.getDate() + 1);
                navigateToTableWithFilter('_dob', null, 'occupied', { from: fromDate, to: toDate });
            },
            plugins: { ...chartOptions('').plugins, tooltip: { ...chartOptions('').plugins.tooltip,
                callbacks: { label: ctx => ` ${ctx.raw} (${ageTotalCount ? (ctx.raw/ageTotalCount*100).toFixed(1) : 0}%)`, footer: () => '🔍 Click to filter Data Table' }
            }}
        }
    });

    // 7. Retirement Forecast (clickable)
    const now = new Date();
    const retYears = {};
    for (let y = now.getFullYear(); y <= now.getFullYear() + 5; y++) retYears[y] = 0;
    data.forEach(r => {
        if (r.retirement_year && retYears.hasOwnProperty(r.retirement_year)) retYears[r.retirement_year]++;
    });
    const retLabels = Object.keys(retYears);
    const retValues = Object.values(retYears);
    const retTotal = retValues.reduce((a, b) => a + b, 0);
    dashCharts.retirement = new Chart(document.getElementById('chartRetirement'), {
        type: 'bar',
        data: { labels: retLabels, datasets: [{ label: 'Retirements', data: retValues, backgroundColor: CHART_COLORS[2] + '99', borderColor: CHART_COLORS[2], borderWidth: 1, borderRadius: 4 }] },
        options: {
            ...chartOptions('Retirement Forecast'),
            onClick: (_ev, els) => {
                if (!els.length) return;
                navigateToTableWithFilter('retirement_year', retLabels[els[0].index], 'occupied');
            },
            plugins: { ...chartOptions('').plugins, tooltip: { ...chartOptions('').plugins.tooltip,
                callbacks: { label: ctx => ` ${ctx.raw} (${retTotal ? (ctx.raw/retTotal*100).toFixed(1) : 0}%)`, footer: () => '🔍 Click to filter Data Table' }
            }}
        }
    });

    // 8. Top 15 Designations (stacked: occupied vs vacant)
    dashCharts.designations = createDeptStackedChart('chartDesignations', allData, 15, 'designation_name');

    // 9. Posted During (Year posted to present post — clickable)
    const postedYearCounts = {};
    data.forEach(r => {
        if (r._presentPostSince) {
            const yr = r._presentPostSince.getFullYear();
            postedYearCounts[yr] = (postedYearCounts[yr] || 0) + 1;
        }
    });
    const postedYearEntries = Object.entries(postedYearCounts).sort((a, b) => a[0] - b[0]);
    const postedLabels = postedYearEntries.map(e => e[0]);
    const postedValues = postedYearEntries.map(e => e[1]);
    const postedTotal = postedValues.reduce((a, b) => a + b, 0);
    dashCharts.postedDuring = new Chart(document.getElementById('chartPostedDuring'), {
        type: 'bar',
        data: { labels: postedLabels, datasets: [{ label: 'Employees Posted', data: postedValues, backgroundColor: CHART_COLORS[4] + '99', borderColor: CHART_COLORS[4], borderWidth: 1, borderRadius: 4 }] },
        options: {
            ...chartOptions('Posted During (Year)'),
            onClick: (_ev, els) => {
                if (!els.length) return;
                const year = parseInt(postedLabels[els[0].index]);
                navigateToTableWithFilter('_presentPostSince', null, 'occupied', {
                    from: new Date(year, 0, 1),
                    to:   new Date(year, 11, 31, 23, 59, 59)
                });
            },
            plugins: { ...chartOptions('').plugins, tooltip: { ...chartOptions('').plugins.tooltip,
                callbacks: { label: ctx => ` ${ctx.raw} (${postedTotal ? (ctx.raw/postedTotal*100).toFixed(1) : 0}%)`, footer: () => '🔍 Click to filter Data Table' }
            }}
        }
    });

    // 10. Posted By
    dashCharts.postedBy = createCountChart('chartPostedBy', data, 'present_post_by', 'doughnut');
}

// field: the grouping field key (e.g. 'hod_name', 'designation_name')
function createDeptStackedChart(canvasId, allData, topN = 15, field = 'hod_name') {
    const grpMap = {};
    allData.forEach(r => {
        const d = r[field] || 'N/A';
        if (!grpMap[d]) grpMap[d] = { occupied: 0, vacant: 0 };
        if (r.is_occupied) grpMap[d].occupied++;
        else grpMap[d].vacant++;
    });

    let entries = Object.entries(grpMap)
        .map(([name, v]) => ({ name, occupied: v.occupied, vacant: v.vacant, total: v.occupied + v.vacant }))
        .sort((a, b) => b.total - a.total);
    if (topN) entries = entries.slice(0, topN);

    const labels = entries.map(e => e.name);
    const occupiedVals = entries.map(e => e.occupied);
    const vacantVals = entries.map(e => e.vacant);
    const totals = entries.map(e => e.total);
    const base = chartOptions('');

    return new Chart(document.getElementById(canvasId), {
        type: 'bar',
        data: {
            labels,
            datasets: [
                { label: 'Occupied', data: occupiedVals, backgroundColor: '#2e7d32bb', borderColor: '#2e7d32', borderWidth: 1, borderRadius: 2 },
                { label: 'Vacant',   data: vacantVals,   backgroundColor: '#c62828bb', borderColor: '#c62828', borderWidth: 1, borderRadius: 2 }
            ]
        },
        options: {
            ...base,
            indexAxis: 'y',
            scales: {
                x: { stacked: true, ticks: { color: '#6b3a00', font: { size: 10 } }, grid: { color: '#e8b86d' } },
                y: { stacked: true, ticks: { color: '#6b3a00', font: { size: 9 } }, grid: { color: '#e8b86d' } }
            },
            onClick: (_event, elements) => {
                if (!elements.length) return;
                const dsIdx = elements[0].datasetIndex; // 0=Occupied, 1=Vacant
                const postFilter = dsIdx === 1 ? 'vacant' : 'occupied';
                navigateToTableWithFilter(field, labels[elements[0].index], postFilter);
            },
            plugins: {
                ...base.plugins,
                legend: { display: true, position: 'top', labels: { color: '#6b3a00', font: { size: 11 } } },
                tooltip: {
                    ...base.plugins.tooltip,
                    callbacks: {
                        label: (ctx) => {
                            const val = ctx.raw;
                            const total = totals[ctx.dataIndex];
                            const pct = total ? (val / total * 100).toFixed(1) : '0.0';
                            return ` ${ctx.dataset.label}: ${val} (${pct}% of total)`;
                        },
                        footer: () => '🔍 Click to filter Data Table'
                    }
                }
            }
        }
    });
}

function createCountChart(canvasId, data, field, type, topN = null, horizontal = false) {
    const counts = {};
    data.forEach(r => {
        const v = r[field] || 'N/A';
        counts[v] = (counts[v] || 0) + 1;
    });

    let entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    if (topN) entries = entries.slice(0, topN);

    const labels = entries.map(e => e[0]);
    const values = entries.map(e => e[1]);
    const colors = getColors(labels.length);

    const isPie = ['pie', 'doughnut', 'polarArea'].includes(type);
    // Check if this field is filterable in the data table
    const isFilterable = ALL_FILTER_FIELDS.some(f => f.key === field && f.type !== 'date');

    const config = {
        type: type === 'horizontalBar' ? 'bar' : type,
        data: {
            labels,
            datasets: [{
                label: 'Count',
                data: values,
                backgroundColor: isPie ? colors : colors[0] + '99',
                borderColor: isPie ? colors.map(c => c) : colors[0],
                borderWidth: 1,
                borderRadius: isPie ? 0 : 4
            }]
        },
        options: {
            ...chartOptions(''),
            indexAxis: (horizontal || type === 'horizontalBar') ? 'y' : 'x',
            onClick: isFilterable ? (_event, elements) => {
                if (elements.length > 0) {
                    const idx = elements[0].index;
                    navigateToTableWithFilter(field, labels[idx]);
                }
            } : undefined,
            plugins: {
                ...chartOptions('').plugins,
                tooltip: {
                    ...chartOptions('').plugins.tooltip,
                    callbacks: {
                        label: (ctx) => {
                            const val = ctx.raw;
                            const total = values.reduce((a, b) => a + b, 0);
                            const pct = total ? (val / total * 100).toFixed(1) : '0.0';
                            return ` ${val} (${pct}%)`;
                        },
                        ...(isFilterable ? { footer: () => '🔍 Click to filter Data Table' } : {})
                    }
                }
            }
        }
    };

    if (isFilterable) config.options.cursor = 'pointer';

    if (isPie) {
        config.options.plugins.legend = { display: true, position: 'right', labels: { color: '#78716c', font: { size: 11 } } };
    }

    return new Chart(document.getElementById(canvasId), config);
}

function chartOptions(title) {
    return {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            title: { display: !!title, text: title, color: '#1c1917', font: { size: 13, weight: '600' } },
            legend: { display: false },
            tooltip: {
                backgroundColor: '#fff9f0',
                titleColor: '#2d1500',
                bodyColor: '#6b3a00',
                borderColor: '#e8b86d',
                borderWidth: 1,
                cornerRadius: 8,
                padding: 10
            }
        },
        scales: {
            x: { ticks: { color: '#6b3a00', font: { size: 10 } }, grid: { color: '#e8b86d' } },
            y: { ticks: { color: '#6b3a00', font: { size: 10 } }, grid: { color: '#e8b86d' } }
        }
    };
}

function toggleChartFullscreen(btn) {
    const card = btn.closest('.chart-card');
    card.classList.toggle('fullscreen');
    btn.textContent = card.classList.contains('fullscreen') ? '✕' : '⛶';
    // Resize chart inside
    const canvas = card.querySelector('canvas');
    if (canvas) {
        const chartInstance = Chart.getChart(canvas);
        if (chartInstance) setTimeout(() => chartInstance.resize(), 100);
    }
}

// ──────────────────────────────────────────────────────────────
//   DATA TABLE PAGE
// ──────────────────────────────────────────────────────────────

function initDataTable() {
    // Initialize visible columns
    TABLE_COLUMNS.forEach(c => { visibleColumns[c.key] = c.show; });

    // Build column toggle dropdown
    const dropdown = document.getElementById('colToggleDropdown');
    TABLE_COLUMNS.forEach(c => {
        const label = document.createElement('label');
        label.innerHTML = `<input type="checkbox" ${c.show ? 'checked' : ''} data-col="${c.key}"> ${c.label}`;
        label.querySelector('input').addEventListener('change', (e) => {
            visibleColumns[c.key] = e.target.checked;
            renderTable();
        });
        dropdown.appendChild(label);
    });

    document.getElementById('colToggleBtn').addEventListener('click', () => {
        document.getElementById('colToggleDropdown').classList.toggle('active');
    });
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.column-toggle-container')) {
            document.getElementById('colToggleDropdown').classList.remove('active');
        }
    });

    // Pagination
    document.getElementById('pageSize').addEventListener('change', () => { currentPage = 1; renderTable(); });
    document.getElementById('btnFirst').addEventListener('click', () => { currentPage = 1; renderTable(); });
    document.getElementById('btnPrev').addEventListener('click', () => { if (currentPage > 1) { currentPage--; renderTable(); } });
    document.getElementById('btnNext').addEventListener('click', () => {
        const ps = parseInt(document.getElementById('pageSize').value);
        const maxPage = Math.ceil(tableFilteredData.length / ps);
        if (currentPage < maxPage) { currentPage++; renderTable(); }
    });
    document.getElementById('btnLast').addEventListener('click', () => {
        const ps = parseInt(document.getElementById('pageSize').value);
        currentPage = Math.ceil(tableFilteredData.length / ps);
        renderTable();
    });

    applyTableFilters();
}

function applyTableFilters() {
    const search = document.getElementById('fTableSearch').value.toLowerCase();
    const svcMin = parseInt(document.getElementById('fTableServiceMin').value) || 0;
    const svcMax = parseInt(document.getElementById('fTableServiceMax').value) || 40;
    const filters = getUnifiedFilterValues('table');

    const dtPsOccupied = document.getElementById('dtPsOccupied');
    const dtPsVacant   = document.getElementById('dtPsVacant');
    const dtShowOcc    = !dtPsOccupied || dtPsOccupied.checked;
    const dtShowVac    = dtPsVacant && dtPsVacant.checked;

    tableFilteredData = getEmpTypeData().filter(r => {
        if (r.post_status === 'occupied' && !dtShowOcc) return false;
        if (r.post_status === 'vacant'   && !dtShowVac) return false;
        if (!passesFilters(r, filters)) return false;
        if (r.is_occupied && r.service_years_current_post != null) {
            if (r.service_years_current_post < svcMin || r.service_years_current_post > svcMax) return false;
        }
        if (search) {
            const hay = [r.emp_name, r.emp_id, r.designation_name, r.hod_name, r.work_location]
                .map(v => String(v || '').toLowerCase()).join(' ');
            if (!hay.includes(search)) return false;
        }
        return true;
    });

    if (sortCol) {
        tableFilteredData.sort((a, b) => {
            let va = a[sortCol], vb = b[sortCol];
            if (va == null) va = '';
            if (vb == null) vb = '';
            if (typeof va === 'number' && typeof vb === 'number') return sortDir === 'asc' ? va - vb : vb - va;
            return sortDir === 'asc' ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
        });
    }

    currentPage = 1;
    renderTable();
}

function renderTable() {
    const cols = TABLE_COLUMNS.filter(c => visibleColumns[c.key]);
    const hasPhotoCol = cols.some(c => c.key === 'photo');
    const ps = parseInt(document.getElementById('pageSize').value);
    const totalPages = Math.max(1, Math.ceil(tableFilteredData.length / ps));
    if (currentPage > totalPages) currentPage = totalPages;
    const start = (currentPage - 1) * ps;
    const pageData = tableFilteredData.slice(start, start + ps);

    // Toggle photo-row class on table
    const tableEl = document.getElementById('dataTable');
    tableEl.classList.toggle('table-photo-rows', hasPhotoCol);

    // Header
    const thead = document.getElementById('dataTableHead');
    thead.innerHTML = '<tr>' + cols.map(c => {
        if (c.key === 'photo') return `<th class="photo-th">Photo</th>`;
        const icon = sortCol === c.key ? (sortDir === 'asc' ? '▲' : '▼') : '⇅';
        return `<th data-col="${c.key}">${c.label} <span class="sort-icon">${icon}</span></th>`;
    }).join('') + '</tr>';

    thead.querySelectorAll('th[data-col]').forEach(th => {
        th.addEventListener('click', () => {
            const col = th.dataset.col;
            if (sortCol === col) sortDir = sortDir === 'asc' ? 'desc' : 'asc';
            else { sortCol = col; sortDir = 'asc'; }
            applyTableFilters();
        });
    });

    // Body
    const tbody = document.getElementById('dataTableBody');
    tbody.innerHTML = pageData.map(r => {
        return `<tr data-empid="${r.emp_id}">` + cols.map(c => {
            if (c.key === 'photo') {
                return `<td class="photo-td">${_empPhotoHtml(r.emp_id, r.gender, true)}</td>`;
            }
            let val = r[c.key];
            if (val == null || val === '' || val === 'null' || val === 'None') val = '—';
            if (c.key === 'post_status') {
                const cls = r.is_occupied ? 'badge-occupied' : 'badge-vacant';
                val = `<span class="badge ${cls}">${r.is_occupied ? 'Occupied' : 'Vacant'}</span>`;
            }
            if (c.key === 'service_years_current_post' && typeof val === 'number') val = val.toFixed(1);
            return `<td title="${String(r[c.key] || '')}">${val}</td>`;
        }).join('') + '</tr>';
    }).join('');

    tbody.querySelectorAll('tr').forEach(tr => {
        tr.addEventListener('click', () => {
            const empId = tr.dataset.empid;
            const emp = RAW_DATA.find(r => String(r.emp_id) === String(empId));
            if (emp) openEmployeePopup(emp);
        });
    });

    // Info
    document.getElementById('tableRecordCount').textContent = `${formatNum(tableFilteredData.length)} records`;
    document.getElementById('paginationInfo').textContent = `${start + 1}–${Math.min(start + ps, tableFilteredData.length)} of ${tableFilteredData.length}`;
    document.getElementById('pageIndicator').textContent = `Page ${currentPage} of ${totalPages}`;
}

// ── Table Export ──
function exportTableExcel() {
    const cols = TABLE_COLUMNS.filter(c => visibleColumns[c.key] && !c.noExport);
    const wsData = [cols.map(c => c.label)];
    tableFilteredData.forEach(r => {
        wsData.push(cols.map(c => {
            let v = r[c.key];
            if (v == null || v === 'null' || v === 'None') return '';
            return v;
        }));
    });
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'TTD Employees');
    XLSX.writeFile(wb, 'TTD_Employees_Data.xlsx');
}

function exportTablePDF() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const cols = TABLE_COLUMNS.filter(c => visibleColumns[c.key] && !c.noExport);

    doc.setFontSize(14);
    doc.text('TTD Employee Data Report', 14, 15);
    doc.setFontSize(8);
    doc.text(`Generated: ${new Date().toLocaleString()} | Records: ${tableFilteredData.length}`, 14, 21);

    const head = [cols.map(c => c.label)];
    const body = tableFilteredData.slice(0, 500).map(r => cols.map(c => {
        let v = r[c.key];
        if (v == null || v === 'null' || v === 'None') return '';
        return String(v);
    }));

    doc.autoTable ? doc.autoTable({
        head, body, startY: 25,
        styles: { fontSize: 6, halign: 'center', valign: 'middle', lineWidth: 0.2, lineColor: [180, 83, 9] },
        headStyles: { fillColor: [217, 119, 6], halign: 'center', valign: 'middle', fontStyle: 'bold' },
        bodyStyles: { halign: 'center', valign: 'middle' },
        columnStyles: { 0: { halign: 'left' } }
    }) : simpleTablePDF(doc, head[0], body, 25);

    doc.save('TTD_Employees_Data.pdf');
}

function simpleTablePDF(doc, headers, rows, startY) {
    const colWidth = (doc.internal.pageSize.getWidth() - 28) / headers.length;
    let y = startY;
    doc.setFontSize(6);
    doc.setFont(undefined, 'bold');
    headers.forEach((h, i) => doc.text(h, 14 + i * colWidth, y, { maxWidth: colWidth - 2 }));
    y += 5;
    doc.setFont(undefined, 'normal');
    rows.forEach(row => {
        if (y > doc.internal.pageSize.getHeight() - 15) { doc.addPage(); y = 15; }
        row.forEach((v, i) => doc.text(String(v).substring(0, 25), 14 + i * colWidth, y, { maxWidth: colWidth - 2 }));
        y += 4;
    });
}

// ──────────────────────────────────────────────────────────────
//   PIVOT TABLE PAGE
// ──────────────────────────────────────────────────────────────

const PIVOT_FIELDS = [
    { key: 'hod_name', label: 'Department (HOD)' },
    { key: 'hos_name', label: 'Section Head' },
    { key: 'section_name', label: 'Section' },
    { key: 'designation_name', label: 'Designation' },
    { key: 'community', label: 'Community' },
    { key: 'sub_community', label: 'Sub Community' },
    { key: 'caste', label: 'Caste' },
    { key: 'gender', label: 'Gender' },
    { key: 'work_location', label: 'Work Location' },
    { key: 'type_of_recruitment', label: 'Recruitment Type' },
    { key: 'joined_during', label: 'Joined in TTD During' },
    { key: 'present_post_by', label: 'Present Post Posted In' },
    { key: 'native_dist', label: 'Native District' },
    { key: 'local_dist_study', label: 'Local/Study District' },
    { key: 'employee_type', label: 'Employee Type' },
    { key: 'post_status', label: 'Post Status' },
    { key: 'retirement_year', label: 'Retirement Year' },
    { key: 'recruited_post', label: 'Recruited Post' }
];

const pivotState = { rows: [], cols: [] };
const pivotFilterFields = {}; // key -> { label, enabled }
const MULTISELECT_PIVOT_STATE = {};

// Pivot expand/collapse state — key: compound path, false=collapsed, undefined=expanded
let pivotExpandState = {};

// Pivot column expand/collapse state — key: first-col-field-value, false=collapsed, undefined=expanded
let pivotColExpandState = {};

// Module-level storage for renderPivotBody() to use without re-reading DOM
let _pvRowGroups = {};
let _pvDisplayCols = [];
let _pvRowFields = [];
let _pvGrandTotals = {};
let _pvGrandGrand = 0;
let _pvStyles = {};
let _pvValueMode = 'count';
let _pvSortedRowKeys = [];
let _pvVisibleCols = null;  // array of {key, isGroupTotal, subCols, cf1}
let _pvColFields = [];       // column field keys
let _pvHeaderAngle = 0;
let _pvHeaderAlign = 'left';

function initPivotTable() {
    const fieldList = document.getElementById('fieldList');
    PIVOT_FIELDS.forEach(f => {
        const chip = createFieldChip(f.key, f.label);
        fieldList.appendChild(chip);
    });

    // Setup drop zones (only rows and cols)
    document.querySelectorAll('.drop-area').forEach(zone => {
        zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.closest('.drop-zone').classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => { zone.closest('.drop-zone').classList.remove('drag-over'); });
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.closest('.drop-zone').classList.remove('drag-over');
            const fieldKey = e.dataTransfer.getData('text/plain');
            const fieldLabel = e.dataTransfer.getData('text/label');
            const zoneName = zone.dataset.zone;
            addFieldToZone(fieldKey, fieldLabel, zoneName, zone);
        });
    });

    document.getElementById('btnGeneratePivot').addEventListener('click', generatePivotTable);

    // Filter Fields provision
    initPivotFilterProvision();
}

function initPivotFilterProvision() {
    const dropdown = document.getElementById('pivotFilterFieldDropdown');
    const btn = document.getElementById('pivotFilterFieldBtn');
    const grid = document.getElementById('pivotFilterGrid');

    // Build checkbox list
    PIVOT_FIELDS.forEach(f => {
        pivotFilterFields[f.key] = { label: f.label, enabled: false };
        const lbl = document.createElement('label');
        lbl.innerHTML = `<input type="checkbox" data-key="${f.key}"> ${f.label}`;
        lbl.style.cssText = 'display:flex;align-items:center;gap:8px;padding:4px 0;font-size:0.8rem;color:var(--text-secondary);cursor:pointer;user-select:none;';
        const cb = lbl.querySelector('input');
        cb.addEventListener('change', () => {
            pivotFilterFields[f.key].enabled = cb.checked;
            renderPivotFilterGrid();
        });
        dropdown.appendChild(lbl);
    });

    // Toggle dropdown
    btn.addEventListener('click', e => {
        e.stopPropagation();
        dropdown.classList.toggle('active');
    });
    document.addEventListener('click', () => dropdown.classList.remove('active'));
    dropdown.addEventListener('click', e => e.stopPropagation());

    // Clear all filters
    document.getElementById('pivotClearFilters').addEventListener('click', () => {
        Object.keys(MULTISELECT_PIVOT_STATE).forEach(k => {
            MULTISELECT_PIVOT_STATE[k] = new Set();
            const w = document.getElementById('pivms-' + k);
            if (w) {
                const disp = w.querySelector('.ms-display');
                if (disp) { disp.textContent = '— All —'; disp.className = 'ms-display placeholder'; }
                w.querySelectorAll('input[type="checkbox"]').forEach(cb => { cb.checked = false; });
                w.querySelectorAll('.ms-option').forEach(el => el.classList.remove('checked'));
            }
        });
    });
}

function renderPivotFilterGrid() {
    const grid = document.getElementById('pivotFilterGrid');
    grid.innerHTML = '';
    let hasAny = false;
    PIVOT_FIELDS.forEach(f => {
        if (!pivotFilterFields[f.key] || !pivotFilterFields[f.key].enabled) return;
        hasAny = true;
        // Get unique values for this field
        const vals = [...new Set(getEmpTypeData().filter(r => r.is_occupied).map(r => String(r[f.key] || '')).filter(v => v && v !== 'None' && v !== 'null'))].sort();
        const group = document.createElement('div');
        group.className = 'filter-group';
        group.innerHTML = `<label>${f.label}</label>`;

        // Hidden native select for the wrapper target
        const sel = document.createElement('select');
        sel.id = 'pivotf-' + f.key;
        sel.style.display = 'none';
        group.appendChild(sel);
        grid.appendChild(group);

        // Init multiselect using pivot-specific state
        initPivotMultiSelect('pivotf-' + f.key, f.key, vals, group);
    });
    grid.style.display = hasAny ? '' : 'none';
}

function initPivotMultiSelect(wrapperId, fieldKey, options, parentEl) {
    if (!MULTISELECT_PIVOT_STATE[fieldKey]) MULTISELECT_PIVOT_STATE[fieldKey] = new Set();

    const wrapper = document.createElement('div');
    wrapper.className = 'ms-wrapper';
    wrapper.id = 'pivms-' + fieldKey;

    const trigger = document.createElement('div');
    trigger.className = 'ms-trigger';
    trigger.innerHTML = '<span class="ms-display placeholder">— All —</span><span class="ms-arrow">▾</span>';

    const dropdownEl = document.createElement('div');
    dropdownEl.className = 'ms-dropdown';
    dropdownEl.innerHTML =
        '<div class="ms-search-wrap"><input class="ms-search" placeholder="Search..."></div>' +
        '<div class="ms-actions"><button class="ms-select-all">✔ Select All</button><button class="ms-deselect-all">✕ Clear</button></div>' +
        '<div class="ms-list"></div>';

    wrapper.appendChild(trigger);
    wrapper.appendChild(dropdownEl);
    parentEl.appendChild(wrapper);

    const listEl = dropdownEl.querySelector('.ms-list');
    const searchEl = dropdownEl.querySelector('.ms-search');
    const displayEl = trigger.querySelector('.ms-display');

    function buildList(filter) {
        filter = (filter || '').toLowerCase();
        listEl.innerHTML = '';
        options.forEach(opt => {
            if (filter && !String(opt).toLowerCase().includes(filter)) return;
            const item = document.createElement('label');
            const checked = MULTISELECT_PIVOT_STATE[fieldKey].has(opt);
            item.className = 'ms-option' + (checked ? ' checked' : '');
            const cb = document.createElement('input');
            cb.type = 'checkbox'; cb.value = opt; cb.checked = checked;
            cb.addEventListener('change', () => {
                if (cb.checked) { MULTISELECT_PIVOT_STATE[fieldKey].add(opt); item.classList.add('checked'); }
                else { MULTISELECT_PIVOT_STATE[fieldKey].delete(opt); item.classList.remove('checked'); }
                updateDisplay();
            });
            item.appendChild(cb);
            item.appendChild(document.createTextNode(' ' + opt));
            listEl.appendChild(item);
        });
    }

    function updateDisplay() {
        const sel = [...MULTISELECT_PIVOT_STATE[fieldKey]];
        if (sel.length === 0) { displayEl.textContent = '— All —'; displayEl.className = 'ms-display placeholder'; }
        else if (sel.length === 1) { displayEl.textContent = sel[0]; displayEl.className = 'ms-display'; }
        else { displayEl.textContent = sel.length + ' selected'; displayEl.className = 'ms-display'; }
    }

    buildList();
    searchEl.addEventListener('input', () => buildList(searchEl.value));
    dropdownEl.querySelector('.ms-select-all').addEventListener('click', () => {
        options.forEach(o => MULTISELECT_PIVOT_STATE[fieldKey].add(o));
        buildList(searchEl.value); updateDisplay();
    });
    dropdownEl.querySelector('.ms-deselect-all').addEventListener('click', () => {
        MULTISELECT_PIVOT_STATE[fieldKey].clear();
        buildList(searchEl.value); updateDisplay();
    });
    trigger.addEventListener('click', e => {
        e.stopPropagation();
        const isOpen = wrapper.classList.contains('open');
        document.querySelectorAll('.ms-wrapper.open').forEach(w => w.classList.remove('open'));
        if (!isOpen) { wrapper.classList.add('open'); searchEl.value = ''; buildList(); searchEl.focus(); }
    });
    document.addEventListener('click', () => wrapper.classList.remove('open'));
    dropdownEl.addEventListener('click', e => e.stopPropagation());
}

function createFieldChip(key, label) {
    const chip = document.createElement('div');
    chip.className = 'field-chip';
    chip.draggable = true;
    chip.dataset.key = key;
    chip.innerHTML = `${label}<span class="remove-chip" title="Remove">×</span>`;
    chip.addEventListener('dragstart', (e) => {
        e.dataTransfer.setData('text/plain', key);
        e.dataTransfer.setData('text/label', label);
        chip.classList.add('dragging');
    });
    chip.addEventListener('dragend', () => chip.classList.remove('dragging'));
    chip.querySelector('.remove-chip').addEventListener('click', () => {
        const zone = chip.closest('.drop-area');
        if (zone) {
            const zoneName = zone.dataset.zone;
            if (pivotState[zoneName]) pivotState[zoneName] = pivotState[zoneName].filter(f => f.key !== key);
            chip.remove();
            // Add back to field bank
            const bankChip = createFieldChip(key, label);
            document.getElementById('fieldList').appendChild(bankChip);
        }
    });
    return chip;
}

function addFieldToZone(key, label, zoneName, zoneEl) {
    if (!['rows', 'cols'].includes(zoneName)) return;

    // Remove from field bank if present
    const bankChip = document.querySelector(`#fieldList .field-chip[data-key="${key}"]`);
    if (bankChip) bankChip.remove();

    // Check if already in this zone
    if (pivotState[zoneName].find(f => f.key === key)) return;

    // Remove from other zones if present
    for (const z of ['rows', 'cols']) {
        pivotState[z] = pivotState[z].filter(f => f.key !== key);
        const existing = document.querySelector(`.drop-area[data-zone="${z}"] .field-chip[data-key="${key}"]`);
        if (existing) existing.remove();
    }

    pivotState[zoneName].push({ key, label });
    const chip = createFieldChip(key, label);
    chip.classList.add('in-zone');
    zoneEl.appendChild(chip);
}

function generatePivotTable() {
    if (pivotState.rows.length === 0) { alert('Please add at least one Row field.'); return; }

    // Reset expand/collapse state on each fresh Generate
    pivotExpandState = {};
    pivotColExpandState = {};

    // Read styling options
    const colHeaderColor    = (document.getElementById('ptColHeaderColor')     || {}).value || '#7b1c00';
    const rowHeaderColor    = (document.getElementById('ptRowHeaderColor')     || {}).value || '#fef3c7';
    const colHeaderTextClr  = (document.getElementById('ptColHeaderTextColor') || {}).value || '#ffffff';
    const valueAlign        = (document.getElementById('ptValueAlign')         || {}).value || 'center';
    const fontSize          = (document.getElementById('ptFontSize')           || {}).value || '12';
    const fontFamily        = (document.getElementById('ptFont')               || {}).value || 'inherit';
    const fontColor         = (document.getElementById('ptFontColor')          || {}).value || '#1c1917';
    const titleFontSize     = (document.getElementById('ptTitleFontSize')      || {}).value || '16';
    const titleFont         = (document.getElementById('ptTitleFont')          || {}).value || 'inherit';
    const titleColor        = (document.getElementById('ptTitleColor')         || {}).value || '#7b1c00';
    const valueMode         = (document.getElementById('ptValueMode')          || {}).value || 'count';
    const headerAlign       = (document.getElementById('ptHeaderAlign')        || {}).value || 'left';
    const headerAngle       = parseInt((document.getElementById('ptHeaderAngle') || {}).value || '0', 10) || 0;

    // Read Post Status checkboxes (default: Occupied only)
    const ptPsOccupied = document.getElementById('ptPsOccupied');
    const ptPsVacant   = document.getElementById('ptPsVacant');
    const showOccupied = !ptPsOccupied || ptPsOccupied.checked;
    const showVacant   = ptPsVacant && ptPsVacant.checked;
    if (!showOccupied && !showVacant) {
        alert('Please select at least one Post Status (Occupied / Vacant).');
        return;
    }

    const baseData = getEmpTypeData().filter(r => {
        if (r.post_status === 'occupied') return showOccupied;
        if (r.post_status === 'vacant')   return showVacant;
        return false;
    });

    // Apply filters from filter bar
    let data = baseData.filter(r => {
        for (const [fk, vals] of Object.entries(MULTISELECT_PIVOT_STATE)) {
            if (vals.size && !vals.has(String(r[fk] || ''))) return false;
        }
        return true;
    });

    const rowFields = pivotState.rows.map(f => f.key);
    const colFields = pivotState.cols.map(f => f.key); // support multiple col fields

    // Get unique compound column values (cartesian of all col field values)
    let colValues = [];
    if (colFields.length > 0) {
        const colSet = new Set();
        data.forEach(r => {
            const cv = colFields.map(cf => String(r[cf] || 'N/A')).join(' \u00D7 ');
            colSet.add(cv);
        });
        colValues = [...colSet].sort();
    }

    // Build row groups
    const rowGroups = {};
    data.forEach(r => {
        const rowKey = rowFields.map(f => String(r[f] || 'N/A')).join('|||');
        if (!rowGroups[rowKey]) rowGroups[rowKey] = {};

        if (colFields.length > 0) {
            const cv = colFields.map(cf => String(r[cf] || 'N/A')).join(' \u00D7 ');
            rowGroups[rowKey][cv] = (rowGroups[rowKey][cv] || 0) + 1;
        } else {
            rowGroups[rowKey]['Total'] = (rowGroups[rowKey]['Total'] || 0) + 1;
        }
    });

    const sortedRowKeys = Object.keys(rowGroups).sort();
    const displayCols = colFields.length > 0 ? colValues : ['Total'];

    const rowLabel = document.getElementById('pivotRowLabel').value || rowFields.map(f => PIVOT_FIELDS.find(pf => pf.key === f)?.label || f).join(' / ');
    const colLabel = document.getElementById('pivotColLabel').value || colFields.map(f => PIVOT_FIELDS.find(pf => pf.key === f)?.label || f).join(' × ');
    const valLabel = document.getElementById('pivotValueLabel').value || 'Count';

    // Pre-compute grand total for percentage calculations
    let grandGrand = 0;
    const grandTotals = {};
    displayCols.forEach(c => grandTotals[c] = 0);
    sortedRowKeys.forEach(rk => {
        displayCols.forEach(c => {
            const v = rowGroups[rk][c] || 0;
            grandTotals[c] = (grandTotals[c] || 0) + v;
            grandGrand += v;
        });
    });

    // Apply table-level styling
    const thStyle = `background:${colHeaderColor};color:${colHeaderTextClr};font-size:${fontSize}px;font-family:${fontFamily};text-align:${headerAlign};`;
    const tdBaseStyle = `text-align:${valueAlign};font-size:${fontSize}px;font-family:${fontFamily};color:${fontColor};`;
    const rowHdrStyle = `text-align:left;background:${rowHeaderColor};font-size:${fontSize}px;font-family:${fontFamily};color:${fontColor};font-weight:600;`;
    const gtHdrStyle  = `background:${colHeaderColor};color:${colHeaderTextClr};font-size:${fontSize}px;font-family:${fontFamily};text-align:${headerAlign};`;

    // Helper to render a header cell text (with angle support)
    function thContent(text) {
        if (headerAngle > 0) {
            return `<span class="th-text" style="transform:rotate(-${headerAngle}deg)">${text}</span>`;
        }
        return text;
    }
    const angledClass = headerAngle > 0 ? ' class="angled"' : '';
    // Estimate height for angled headers (rough: label length * sin(angle) * fontSize)
    const angledHeightStyle = headerAngle > 0 ? `height:${Math.round(80 * Math.sin(headerAngle * Math.PI / 180) + 24)}px;` : '';

    // Helper: escape a key for use inside onclick="...'"
    function escGKgen(s) { return s.replace(/\\/g, '\\\\').replace(/'/g, "\\'"); }

    // Build visible columns structure (for col expand/collapse)
    function buildVisibleCols(dCols, cFields) {
        if (cFields.length <= 1) {
            return dCols.map(c => ({ key: c, isGroupTotal: false, subCols: null, cf1: c }));
        }
        const cf1Groups = {};
        dCols.forEach(ck => {
            const cf1 = ck.split(' \u00D7 ')[0];
            if (!cf1Groups[cf1]) cf1Groups[cf1] = [];
            cf1Groups[cf1].push(ck);
        });
        const result = [];
        Object.keys(cf1Groups).sort().forEach(cf1 => {
            if (pivotColExpandState[cf1] === false) {
                result.push({ key: cf1 + '__group', isGroupTotal: true, subCols: cf1Groups[cf1], cf1 });
            } else {
                cf1Groups[cf1].forEach(ck => result.push({ key: ck, isGroupTotal: false, subCols: null, cf1 }));
            }
        });
        return result;
    }

    // Build thead HTML (handles single-row and two-row header)
    function buildColHeaderHtml(rFields, visCols, dCols, cFields) {
        if (cFields.length <= 1) {
            // Single-row header
            let hHtml = `<tr style="${thStyle}">`;
            // Row field headers — NO angle
            rFields.forEach((f, i) => {
                const label = rowLabel.split('/')[i]?.trim() || PIVOT_FIELDS.find(pf => pf.key === f)?.label || f;
                hHtml += `<th style="${rowHdrStyle}text-align:left;">${label}</th>`;
            });
            // Value column headers — YES angle
            visCols.forEach(vc => {
                const lbl = vc.isGroupTotal ? ('Total: ' + vc.cf1) : vc.key;
                hHtml += `<th${angledClass} style="${thStyle}${angledHeightStyle}">${thContent(lbl)}</th>`;
            });
            // Grand Total — YES angle
            hHtml += `<th${angledClass} style="${gtHdrStyle}${angledHeightStyle}">${thContent('Grand Total')}</th></tr>`;
            return hHtml;
        }

        // Two-row header for multiple column fields
        const cf1Groups = {};
        dCols.forEach(ck => {
            const cf1 = ck.split(' \u00D7 ')[0];
            if (!cf1Groups[cf1]) cf1Groups[cf1] = [];
            cf1Groups[cf1].push(ck);
        });

        let row1 = `<tr style="${thStyle}">`;
        // Row field spacer cells (rowspan=2) — NO angle
        rFields.forEach((f, i) => {
            const label = rowLabel.split('/')[i]?.trim() || PIVOT_FIELDS.find(pf => pf.key === f)?.label || f;
            row1 += `<th rowspan="2" style="${rowHdrStyle}text-align:left;">${label}</th>`;
        });
        // CF1 group cells — YES angle, clickable
        Object.keys(cf1Groups).sort().forEach(cf1 => {
            const isCollapsed = pivotColExpandState[cf1] === false;
            const colSpan = isCollapsed ? 1 : cf1Groups[cf1].length;
            const icon = isCollapsed ? '&#9658;' : '&#9660;';
            row1 += `<th colspan="${colSpan}"${angledClass} style="${thStyle}${angledHeightStyle}cursor:pointer;" onclick="togglePivotColGroup('${escGKgen(cf1)}')">${thContent(icon + ' ' + cf1)}</th>`;
        });
        // Grand Total (rowspan=2) — YES angle
        row1 += `<th rowspan="2"${angledClass} style="${gtHdrStyle}${angledHeightStyle}">${thContent('Grand Total')}</th>`;
        row1 += '</tr>';

        let row2 = `<tr style="${thStyle}">`;
        visCols.forEach(vc => {
            if (vc.isGroupTotal) {
                row2 += `<th${angledClass} style="${thStyle}${angledHeightStyle}">${thContent('Total')}</th>`;
            } else {
                const subLabel = vc.key.split(' \u00D7 ').slice(1).join(' \u00D7 ');
                row2 += `<th${angledClass} style="${thStyle}${angledHeightStyle}">${thContent(subLabel)}</th>`;
            }
        });
        row2 += '</tr>';

        return row1 + row2;
    }

    // Compute visible cols
    const visibleCols = buildVisibleCols(displayCols, colFields);

    // Build table
    const table = document.getElementById('pivotTable');
    const thead = document.getElementById('pivotTableHead');

    // Build and set header HTML
    thead.innerHTML = buildColHeaderHtml(rowFields, visibleCols, displayCols, colFields);

    // Add sort to pivot table headers
    thead.querySelectorAll('th').forEach((th, idx) => {
        th.addEventListener('click', () => { sortPivotTable(idx, th); });
    });

    // Save to module-level vars for renderPivotBody
    _pvRowGroups     = rowGroups;
    _pvDisplayCols   = displayCols;
    _pvRowFields     = rowFields;
    _pvGrandTotals   = grandTotals;
    _pvGrandGrand    = grandGrand;
    _pvSortedRowKeys = sortedRowKeys;
    _pvValueMode     = valueMode;
    _pvStyles        = { thStyle, tdBaseStyle, rowHdrStyle, gtHdrStyle, colHeaderColor, rowHeaderColor, colHeaderTextClr, fontSize, fontFamily, fontColor };
    _pvVisibleCols   = visibleCols;
    _pvColFields     = colFields;
    _pvHeaderAngle   = headerAngle;
    _pvHeaderAlign   = headerAlign;

    // Render tbody
    renderPivotBody();

    // Apply table-level font / sizing and auto-centre
    table.style.cssText = `display:table;margin:0 auto;font-size:${fontSize}px;font-family:${fontFamily};color:${fontColor};border-collapse:collapse;`;

    // Show outer wrapper
    const outerEl = document.getElementById('pivotTableOuter');
    if (outerEl) outerEl.style.display = 'inline-block';

    // Update report title display above the table
    const titleEl = document.getElementById('pivotReportTitle');
    if (titleEl) {
        const titleText = document.getElementById('pivotTitle').value || 'Pivot Report';
        titleEl.textContent = titleText;
        titleEl.style.cssText = `display:block;text-align:center;padding:10px 0 4px;font-weight:700;font-size:${titleFontSize}px;font-family:${titleFont};color:${titleColor};`;
    }

    document.getElementById('pivotEmpty').style.display = 'none';

    // Initialise drag-resize handles
    initPivotTableResize();
}

function renderPivotBody() {
    const tbody = document.getElementById('pivotTableBody');
    const rowGroups     = _pvRowGroups;
    const displayCols   = _pvDisplayCols;
    const rowFields     = _pvRowFields;
    const grandTotals   = _pvGrandTotals;
    const grandGrand    = _pvGrandGrand;
    const sortedRowKeys = _pvSortedRowKeys;
    const valueMode     = _pvValueMode;
    const visibleCols   = _pvVisibleCols || displayCols.map(c => ({ key: c, isGroupTotal: false, subCols: null, cf1: c }));
    const { tdBaseStyle, rowHdrStyle, rowHeaderColor } = _pvStyles;

    // Format a cell value according to chosen mode
    function fmtCell(val, rowTotal, colTotal) {
        const gt = grandGrand || 1;
        switch (valueMode) {
            case 'pct_total':     return `${(val / gt * 100).toFixed(1)}%`;
            case 'pct_row':       return `${rowTotal ? (val / rowTotal * 100).toFixed(1) : 0}%`;
            case 'pct_col':       return `${colTotal ? (val / colTotal * 100).toFixed(1) : 0}%`;
            case 'count_pct':     return `${val} (${(val / gt * 100).toFixed(1)}%)`;
            case 'count_pct_row': return `${val} (${rowTotal ? (val / rowTotal * 100).toFixed(1) : 0}%)`;
            case 'count_pct_col': return `${val} (${colTotal ? (val / colTotal * 100).toFixed(1) : 0}%)`;
            default:              return val;
        }
    }

    // Format row total cell
    function fmtRowTotal(rt) {
        if (valueMode === 'count') return rt;
        if (valueMode === 'pct_total') return `${(rt / (grandGrand || 1) * 100).toFixed(1)}%`;
        if (valueMode === 'pct_row') return '100%';
        return rt;
    }

    // Escape a group key for use inside onclick="...'"
    function escGK(s) { return s.replace(/\\/g, '\\\\').replace(/'/g, "\\'"); }

    // Aggregate all rowKeys' visible col values; also compute _total from all displayCols
    function aggregate(rowKeys) {
        const agg = {};
        displayCols.forEach(c => { agg[c] = 0; });
        rowKeys.forEach(rk => displayCols.forEach(c => { agg[c] += rowGroups[rk][c] || 0; }));
        agg._total = displayCols.reduce((s, c) => s + agg[c], 0);
        return agg;
    }

    // Build data cells HTML for aggregated values; extraStyle is appended to each td style
    function dataCells(agg, rowTotal, extraStyle) {
        extraStyle = extraStyle || '';
        let h = '';
        visibleCols.forEach(vc => {
            let val = 0;
            if (vc.isGroupTotal) {
                val = (vc.subCols || []).reduce((s, k) => s + (agg[k] || 0), 0);
            } else {
                val = agg[vc.key] || 0;
            }
            const colTotal = vc.isGroupTotal
                ? (vc.subCols || []).reduce((s, k) => s + (grandTotals[k] || 0), 0)
                : (grandTotals[vc.key] || 0);
            h += `<td style="${tdBaseStyle}${extraStyle}">${fmtCell(val, rowTotal, colTotal)}</td>`;
        });
        h += `<td style="${tdBaseStyle}font-weight:700;${extraStyle}">${fmtRowTotal(rowTotal)}</td>`;
        return h;
    }

    // Recursive multi-level row group rendering
    // level 0..rowFields.length-2 are non-leaf (group with expand/collapse + subtotal)
    // level rowFields.length-1 is leaf (one row per rowKey, no expand/collapse)
    function renderGroup(rowKeys, level, path) {
        const rows = [];

        if (level === rowFields.length - 1) {
            // Leaf level: one data row per rowKey
            rowKeys.slice().sort().forEach(rk => {
                const parts = rk.split('|||');
                const leafVal = parts[level];
                const agg = aggregate([rk]);
                let cells = `<td style="${rowHdrStyle}">${leafVal}</td>`;
                cells += dataCells(agg, agg._total, '');
                rows.push(`<tr>${cells}</tr>`);
            });
            return rows;
        }

        // Non-leaf: group by value at this level
        const groups = {};
        rowKeys.forEach(rk => {
            const key = rk.split('|||')[level];
            if (!groups[key]) groups[key] = [];
            groups[key].push(rk);
        });

        Object.keys(groups).sort().forEach(key => {
            const gRowKeys = groups[key];
            const expandKey = [...path, key].join('|||');
            const isCollapsed = pivotExpandState[expandKey] === false;
            const agg = aggregate(gRowKeys);
            const indent = '\u00a0'.repeat(level * 3);

            if (isCollapsed) {
                // One summary row spanning remaining row field columns
                const colspan = rowFields.length - level;
                let cells = `<td colspan="${colspan}" style="${rowHdrStyle}cursor:pointer;user-select:none;" onclick="togglePivotGroup('${escGK(expandKey)}')">${indent}&#9658; ${key}</td>`;
                cells += dataCells(agg, agg._total, '');
                rows.push(`<tr>${cells}</tr>`);
            } else {
                // Expanded: get child rows recursively
                const childRows = renderGroup(gRowKeys, level + 1, [...path, key]);

                // Total rows this group spans = childRows + 1 subtotal row
                const totalSpan = childRows.length + 1;

                // Group header cell with rowspan injected into first child row
                const groupCell = `<td rowspan="${totalSpan}" style="${rowHdrStyle}vertical-align:middle;cursor:pointer;user-select:none;" onclick="togglePivotGroup('${escGK(expandKey)}')">${indent}&#9660; ${key}</td>`;

                if (childRows.length > 0) {
                    childRows[0] = childRows[0].replace(/^<tr>/, '<tr>' + groupCell);
                }
                rows.push(...childRows);

                // Subtotal row: colspan covers from level+1 to rowFields.length-1
                const subColspan = rowFields.length - level - 1;
                let stCells = '';
                if (subColspan > 0) {
                    stCells += `<td colspan="${subColspan}" style="${rowHdrStyle}font-style:italic;font-weight:700;">${indent}\u00a0\u00a0Subtotal: ${key}</td>`;
                } else {
                    stCells += `<td style="${rowHdrStyle}font-style:italic;font-weight:700;">${indent}\u00a0\u00a0Subtotal: ${key}</td>`;
                }
                stCells += dataCells(agg, agg._total, 'font-style:italic;');
                rows.push(`<tr style="background:rgba(200,134,10,0.13)">${stCells}</tr>`);
            }
        });

        return rows;
    }

    let bodyHtml = '';

    if (rowFields.length <= 1) {
        // Single row field: flat rendering (no grouping, no expand/collapse)
        sortedRowKeys.forEach(rk => {
            const parts = rk.split('|||');
            const agg = aggregate([rk]);
            let cells = `<td style="${rowHdrStyle}">${parts[0] || ''}</td>`;
            cells += dataCells(agg, agg._total, '');
            bodyHtml += `<tr>${cells}</tr>`;
        });
    } else {
        // Multi-level: recursive rendering
        const topLevelRows = renderGroup(sortedRowKeys, 0, []);
        bodyHtml = topLevelRows.join('');
    }

    // Grand total row — "Grand Total" spans all row field columns
    let gtRow = `<tr style="background:${rowHeaderColor}">`;
    gtRow += `<td colspan="${rowFields.length}" style="${rowHdrStyle}font-weight:800;">Grand Total</td>`;
    // For grand total data cells, use full displayCols grand totals
    visibleCols.forEach(vc => {
        let val = 0;
        if (vc.isGroupTotal) {
            val = (vc.subCols || []).reduce((s, k) => s + (grandTotals[k] || 0), 0);
        } else {
            val = grandTotals[vc.key] || 0;
        }
        const colTotal = vc.isGroupTotal
            ? (vc.subCols || []).reduce((s, k) => s + (grandTotals[k] || 0), 0)
            : (grandTotals[vc.key] || 0);
        gtRow += `<td style="${tdBaseStyle}font-weight:700;">${fmtCell(val, grandGrand, colTotal)}</td>`;
    });
    gtRow += `<td style="${tdBaseStyle}font-weight:800;">${valueMode === 'count' ? grandGrand : '100%'}</td>`;
    gtRow += '</tr>';
    bodyHtml += gtRow;

    tbody.innerHTML = bodyHtml;
}

window.togglePivotGroup = function(groupKey) {
    // undefined or true = expanded → collapse
    // false = collapsed → expand (delete key)
    if (pivotExpandState[groupKey] === false) {
        delete pivotExpandState[groupKey];
    } else {
        pivotExpandState[groupKey] = false;
    }
    renderPivotBody();
    initPivotTableResize();
};

window.togglePivotColGroup = function(cf1Key) {
    if (pivotColExpandState[cf1Key] === false) {
        delete pivotColExpandState[cf1Key];
    } else {
        pivotColExpandState[cf1Key] = false;
    }
    rebuildPivotColHeader();
    renderPivotBody();
    initPivotTableResize();
};

function rebuildPivotColHeader() {
    const thead = document.getElementById('pivotTableHead');
    if (!thead) return;

    const colHeaderColor   = _pvStyles.colHeaderColor   || '#7b1c00';
    const colHeaderTextClr = _pvStyles.colHeaderTextClr || '#ffffff';
    const rowHeaderColor   = _pvStyles.rowHeaderColor   || '#fef3c7';
    const fontSize         = _pvStyles.fontSize         || '12';
    const fontFamily       = _pvStyles.fontFamily       || 'inherit';
    const fontColor        = _pvStyles.fontColor        || '#1c1917';
    const headerAngle      = _pvHeaderAngle;
    const headerAlign      = _pvHeaderAlign;

    const thStyle    = `background:${colHeaderColor};color:${colHeaderTextClr};font-size:${fontSize}px;font-family:${fontFamily};text-align:${headerAlign};`;
    const rowHdrStyle = `text-align:left;background:${rowHeaderColor};font-size:${fontSize}px;font-family:${fontFamily};color:${fontColor};font-weight:600;`;
    const gtHdrStyle  = `background:${colHeaderColor};color:${colHeaderTextClr};font-size:${fontSize}px;font-family:${fontFamily};text-align:${headerAlign};`;

    function thContent(text) {
        if (headerAngle > 0) {
            return `<span class="th-text" style="transform:rotate(-${headerAngle}deg)">${text}</span>`;
        }
        return text;
    }
    const angledClass = headerAngle > 0 ? ' class="angled"' : '';
    const angledHeightStyle = headerAngle > 0 ? `height:${Math.round(80 * Math.sin(headerAngle * Math.PI / 180) + 24)}px;` : '';

    function escGK(s) { return s.replace(/\\/g, '\\\\').replace(/'/g, "\\'"); }

    const rowLabel = document.getElementById('pivotRowLabel').value ||
        _pvRowFields.map(f => PIVOT_FIELDS.find(pf => pf.key === f)?.label || f).join(' / ');

    const displayCols = _pvDisplayCols;
    const colFields   = _pvColFields;
    const rowFields   = _pvRowFields;

    // Rebuild visibleCols
    if (colFields.length > 1) {
        const cf1Groups = {};
        displayCols.forEach(ck => {
            const cf1 = ck.split(' \u00D7 ')[0];
            if (!cf1Groups[cf1]) cf1Groups[cf1] = [];
            cf1Groups[cf1].push(ck);
        });
        const result = [];
        Object.keys(cf1Groups).sort().forEach(cf1 => {
            if (pivotColExpandState[cf1] === false) {
                result.push({ key: cf1 + '__group', isGroupTotal: true, subCols: cf1Groups[cf1], cf1 });
            } else {
                cf1Groups[cf1].forEach(ck => result.push({ key: ck, isGroupTotal: false, subCols: null, cf1 }));
            }
        });
        _pvVisibleCols = result;
    } else {
        _pvVisibleCols = displayCols.map(c => ({ key: c, isGroupTotal: false, subCols: null, cf1: c }));
    }

    // Build header HTML
    let headerHtml = '';

    if (colFields.length <= 1) {
        headerHtml = `<tr style="${thStyle}">`;
        rowFields.forEach((f, i) => {
            const label = rowLabel.split('/')[i]?.trim() || PIVOT_FIELDS.find(pf => pf.key === f)?.label || f;
            headerHtml += `<th style="${rowHdrStyle}text-align:left;">${label}</th>`;
        });
        _pvVisibleCols.forEach(vc => {
            const lbl = vc.isGroupTotal ? ('Total: ' + vc.cf1) : vc.key;
            headerHtml += `<th${angledClass} style="${thStyle}${angledHeightStyle}">${thContent(lbl)}</th>`;
        });
        headerHtml += `<th${angledClass} style="${gtHdrStyle}${angledHeightStyle}">${thContent('Grand Total')}</th></tr>`;
    } else {
        const cf1Groups = {};
        displayCols.forEach(ck => {
            const cf1 = ck.split(' \u00D7 ')[0];
            if (!cf1Groups[cf1]) cf1Groups[cf1] = [];
            cf1Groups[cf1].push(ck);
        });

        let row1 = `<tr style="${thStyle}">`;
        rowFields.forEach((f, i) => {
            const label = rowLabel.split('/')[i]?.trim() || PIVOT_FIELDS.find(pf => pf.key === f)?.label || f;
            row1 += `<th rowspan="2" style="${rowHdrStyle}text-align:left;">${label}</th>`;
        });
        Object.keys(cf1Groups).sort().forEach(cf1 => {
            const isCollapsed = pivotColExpandState[cf1] === false;
            const colSpan = isCollapsed ? 1 : cf1Groups[cf1].length;
            const icon = isCollapsed ? '&#9658;' : '&#9660;';
            row1 += `<th colspan="${colSpan}"${angledClass} style="${thStyle}${angledHeightStyle}cursor:pointer;" onclick="togglePivotColGroup('${escGK(cf1)}')">${thContent(icon + ' ' + cf1)}</th>`;
        });
        row1 += `<th rowspan="2"${angledClass} style="${gtHdrStyle}${angledHeightStyle}">${thContent('Grand Total')}</th>`;
        row1 += '</tr>';

        let row2 = `<tr style="${thStyle}">`;
        _pvVisibleCols.forEach(vc => {
            if (vc.isGroupTotal) {
                row2 += `<th${angledClass} style="${thStyle}${angledHeightStyle}">${thContent('Total')}</th>`;
            } else {
                const subLabel = vc.key.split(' \u00D7 ').slice(1).join(' \u00D7 ');
                row2 += `<th${angledClass} style="${thStyle}${angledHeightStyle}">${thContent(subLabel)}</th>`;
            }
        });
        row2 += '</tr>';
        headerHtml = row1 + row2;
    }

    thead.innerHTML = headerHtml;
    thead.querySelectorAll('th').forEach((th, idx) => {
        th.addEventListener('click', () => sortPivotTable(idx, th));
    });
}

function initPivotTableResize() {
    const table = document.getElementById('pivotTable');
    if (!table) return;

    // --- Column resize handles ---
    table.querySelectorAll('thead th').forEach(th => {
        // Remove any existing resizer to avoid duplicates
        const existing = th.querySelector('.col-resizer');
        if (existing) existing.remove();

        const resizer = document.createElement('div');
        resizer.className = 'col-resizer';
        th.appendChild(resizer);

        let startX = 0;
        let startW = 0;

        resizer.addEventListener('mousedown', function(e) {
            e.preventDefault();
            e.stopPropagation();
            startX = e.clientX;
            startW = th.offsetWidth;

            function onMouseMove(ev) {
                const newW = Math.max(40, startW + ev.clientX - startX);
                th.style.width = newW + 'px';
                th.style.minWidth = newW + 'px';
                th.style.whiteSpace = 'normal';
                th.style.wordBreak = 'break-word';
            }
            function onMouseUp() {
                document.removeEventListener('mousemove', onMouseMove);
                document.removeEventListener('mouseup', onMouseUp);
            }
            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
        });
    });

    // --- Row resize handles (last td of each tbody tr) ---
    table.querySelectorAll('tbody tr').forEach(tr => {
        const cells = tr.cells;
        if (!cells.length) return;
        const lastTd = cells[cells.length - 1];

        // Remove existing resizer
        const existing = lastTd.querySelector('.row-resizer');
        if (existing) existing.remove();

        // Need position:relative on the cell
        lastTd.style.position = 'relative';

        const resizer = document.createElement('div');
        resizer.className = 'row-resizer';
        lastTd.appendChild(resizer);

        let startY = 0;
        let startH = 0;

        resizer.addEventListener('mousedown', function(e) {
            e.preventDefault();
            e.stopPropagation();
            startY = e.clientY;
            startH = tr.offsetHeight;

            function onMouseMove(ev) {
                const newH = Math.max(20, startH + ev.clientY - startY);
                tr.style.height = newH + 'px';
            }
            function onMouseUp() {
                document.removeEventListener('mousemove', onMouseMove);
                document.removeEventListener('mouseup', onMouseUp);
            }
            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
        });
    });
}

function sortPivotTable(colIdx, thEl) {
    const table = document.getElementById('pivotTable');
    const tbody = table.querySelector('tbody');
    const rows = Array.from(tbody.querySelectorAll('tr'));
    const grandRow = rows.pop(); // Keep grand total at bottom

    const dir = thEl.dataset.sort === 'asc' ? 'desc' : 'asc';
    thEl.dataset.sort = dir;

    rows.sort((a, b) => {
        const aVal = a.cells[colIdx]?.textContent || '';
        const bVal = b.cells[colIdx]?.textContent || '';
        const aNum = parseFloat(aVal);
        const bNum = parseFloat(bVal);
        if (!isNaN(aNum) && !isNaN(bNum)) return dir === 'asc' ? aNum - bNum : bNum - aNum;
        return dir === 'asc' ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
    });

    tbody.innerHTML = '';
    rows.forEach(r => tbody.appendChild(r));
    tbody.appendChild(grandRow);
}

// ── Pivot Table Export ──
function exportPivotExcel() {
    const table = document.getElementById('pivotTable');
    if (table.style.display === 'none') { alert('Generate a pivot table first.'); return; }
    const title = document.getElementById('pivotTitle').value || 'Pivot Report';
    const ws = XLSX.utils.table_to_sheet(table);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Pivot');
    XLSX.writeFile(wb, title.replace(/[^a-zA-Z0-9]/g, '_') + '.xlsx');
}

// Crop a portion of a canvas vertically: from startY, height = heightPx
function _cropCanvas(src, startY, heightPx) {
    heightPx = Math.max(1, Math.round(heightPx));
    const out = document.createElement('canvas');
    out.width  = src.width;
    out.height = heightPx;
    out.getContext('2d').drawImage(src,
        0, Math.round(startY), src.width, heightPx,
        0, 0,                  src.width, heightPx);
    return out;
}

async function exportPivotPDF() {
    const table = document.getElementById('pivotTable');
    if (!table || table.style.display === 'none') { alert('Generate a pivot table first.'); return; }
    const title = document.getElementById('pivotTitle').value || 'Pivot Report';

    const orientEl = document.getElementById('ptPdfOrientation');
    const orientation = (orientEl && orientEl.value === 'portrait') ? 'portrait' : 'landscape';

    const { jsPDF } = window.jspdf;
    const PAGE_W  = orientation === 'landscape' ? 297 : 210;
    const PAGE_H  = orientation === 'landscape' ? 210 : 297;
    const MARGIN  = 10;
    const FTR_H   = 8;
    const AVAIL_W = PAGE_W - 2 * MARGIN;

    // ── Compute dynamic header height based on title wrapping ──
    const TITLE_FONT_SZ = 12;
    const TITLE_LINE_H  = TITLE_FONT_SZ * 0.352778 * 1.4; // pt → mm with line spacing
    // Measure using a temporary doc so we know how many lines the title needs
    const _tmpDoc = new jsPDF({ orientation, unit: 'mm', format: 'a4' });
    _tmpDoc.setFontSize(TITLE_FONT_SZ);
    const titleLines = _tmpDoc.splitTextToSize(title, AVAIL_W - 30); // -30 leaves room for "(Continued)"
    const HDR_H = Math.max(18, 5 + titleLines.length * TITLE_LINE_H + 3);

    const AVAIL_H = PAGE_H - HDR_H - MARGIN - FTR_H;

    const SCALE = 2;

    // ── 1. Measure each tbody row's pixel position relative to the table top BEFORE
    //       html2canvas runs, so we know exact row boundaries for page-break snapping. ──
    const outer = document.getElementById('pivotTableOuter');
    if (outer) { outer.scrollTop = 0; outer.scrollLeft = 0; }

    const tableRect = table.getBoundingClientRect();
    const bodyRows  = Array.from(document.querySelectorAll('#pivotTableBody tr'));
    // rowBounds in canvas pixels (table-relative, scaled by SCALE)
    const rowBounds = bodyRows.map(tr => {
        const r = tr.getBoundingClientRect();
        return {
            top:    Math.round((r.top    - tableRect.top) * SCALE),
            bottom: Math.round((r.bottom - tableRect.top) * SCALE)
        };
    });

    // ── 2. Capture full table + thead separately ──
    const fullCanvas = await html2canvas(table, {
        scale: SCALE, backgroundColor: '#ffffff', useCORS: true, scrollY: 0
    });
    const theadEl = document.getElementById('pivotTableHead');
    const theadCanvas = await html2canvas(theadEl, {
        scale: SCALE, backgroundColor: '#7b1c00', useCORS: true
    });

    const mmPerPx       = AVAIL_W / fullCanvas.width;
    const theadH_mm     = theadCanvas.height * mmPerPx;
    const theadH_px     = theadCanvas.height;
    const bodyPerPage_mm = AVAIL_H - theadH_mm;
    const bodyPerPage_px = Math.round(bodyPerPage_mm / mmPerPx);
    const totalBody_px  = fullCanvas.height - theadH_px;

    // ── 3. Build row-boundary-aware page breaks ──
    // All positions below are in "body pixel space": 0 = first pixel of tbody in the canvas.
    // We scan rowBounds (which are table-relative canvas pixels) and convert via: bodyPx = rb.top - theadH_px
    const pageBreaks = []; // [{ start, end }] in body pixel space
    let cursor = 0;

    while (cursor < totalBody_px) {
        let pageEnd = cursor;

        for (const rb of rowBounds) {
            const rTop    = rb.top    - theadH_px;   // row top in body pixel space
            const rBottom = rb.bottom - theadH_px;   // row bottom in body pixel space
            if (rBottom <= cursor) continue;          // already placed on a previous page
            if (rTop >= cursor + bodyPerPage_px) break; // starts beyond this page capacity

            if (rBottom <= cursor + bodyPerPage_px) {
                pageEnd = rBottom;   // row fits completely — extend page end to include it
            } else {
                break;               // row would be cut — leave it for the next page
            }
        }

        // Safety: if not even one complete row fits (very tall row), force-include that row
        // to prevent an infinite loop.
        if (pageEnd <= cursor) {
            for (const rb of rowBounds) {
                const rBottom = rb.bottom - theadH_px;
                if (rBottom > cursor) { pageEnd = rBottom; break; }
            }
            if (pageEnd <= cursor) pageEnd = Math.min(cursor + bodyPerPage_px, totalBody_px);
        }

        pageEnd = Math.min(pageEnd, totalBody_px);
        pageBreaks.push({ start: cursor, end: pageEnd });
        cursor = pageEnd;
    }

    if (!pageBreaks.length) pageBreaks.push({ start: 0, end: totalBody_px });

    // ── 4. Chrome helper (maroon band + gold line + footer) ──
    const totalPages = pageBreaks.length;
    function drawChrome(doc, pageNum) {
        doc.setFillColor(123, 28, 0);
        doc.rect(0, 0, PAGE_W, HDR_H, 'F');
        // Draw wrapped title lines centred inside the band
        doc.setFontSize(TITLE_FONT_SZ);
        doc.setTextColor(212, 175, 55);
        const firstLineY = 5 + TITLE_LINE_H;
        titleLines.forEach((line, i) => {
            doc.text(line, PAGE_W / 2, firstLineY + i * TITLE_LINE_H, { align: 'center' });
        });
        if (pageNum > 1) {
            doc.setFontSize(8); doc.setTextColor(255, 215, 90);
            doc.text('(Continued)', PAGE_W - MARGIN, firstLineY, { align: 'right' });
        }
        doc.setDrawColor(200, 134, 10);
        doc.setLineWidth(0.4);
        doc.line(MARGIN, HDR_H + 1, PAGE_W - MARGIN, HDR_H + 1);
        doc.setTextColor(150, 80, 0);
        doc.setFontSize(7);
        doc.text('© Muddada Ravi Chandra IAS, EO, TTD', PAGE_W / 2, PAGE_H - 4, { align: 'center' });
        doc.setTextColor(130, 130, 130);
        doc.text(`Page ${pageNum} of ${totalPages}`, PAGE_W - MARGIN, PAGE_H - 4, { align: 'right' });
    }

    // ── 5. Render pages ──
    const doc     = new jsPDF({ orientation, unit: 'mm', format: 'a4' });
    const tableTop = HDR_H + 3; // mm: below the maroon band

    pageBreaks.forEach(({ start, end }, idx) => {
        const pageNum = idx + 1;
        if (pageNum > 1) doc.addPage();
        drawChrome(doc, pageNum);

        if (pageNum === 1) {
            // Page 1: crop from table top (pixel 0) down through tbody end of this page
            const slicePx  = Math.min(theadH_px + end, fullCanvas.height);
            const sliceImg = _cropCanvas(fullCanvas, 0, slicePx);
            doc.addImage(sliceImg.toDataURL('image/png'), 'PNG',
                MARGIN, tableTop, AVAIL_W, slicePx * mmPerPx);
        } else {
            // Pages 2+: repeat thead, then the body slice for this page
            doc.addImage(theadCanvas.toDataURL('image/png'), 'PNG',
                MARGIN, tableTop, AVAIL_W, theadH_mm);
            const slicePx = end - start;
            if (slicePx > 0) {
                const sliceImg = _cropCanvas(fullCanvas, theadH_px + start, slicePx);
                doc.addImage(sliceImg.toDataURL('image/png'), 'PNG',
                    MARGIN, tableTop + theadH_mm, AVAIL_W, slicePx * mmPerPx);
            }
        }
    });

    doc.save(title.replace(/[^a-zA-Z0-9]/g, '_') + '.pdf');
}

// ── Fullscreen toggle ──
function toggleFullscreen() {
    const btn = document.getElementById('btnFullscreen');
    if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen().catch(() => {});
        if (btn) btn.textContent = '✕';
        if (btn) btn.title = 'Exit Fullscreen';
    } else {
        document.exitFullscreen().catch(() => {});
        if (btn) btn.textContent = '⛶';
        if (btn) btn.title = 'Toggle Fullscreen';
    }
}
document.addEventListener('fullscreenchange', () => {
    const btn = document.getElementById('btnFullscreen');
    if (!btn) return;
    if (document.fullscreenElement) { btn.textContent = '✕'; btn.title = 'Exit Fullscreen'; }
    else { btn.textContent = '⛶'; btn.title = 'Toggle Fullscreen'; }
});

// ── Apply BG color to chart container live ──
function applyChartBg(color) {
    const container = document.getElementById('pivotChartContainer');
    if (container) container.style.background = color;
}

// ──────────────────────────────────────────────────────────────
//   PIVOT CHART PAGE
// ──────────────────────────────────────────────────────────────

function initPivotChart() {
    document.getElementById('btnGenerateChart').addEventListener('click', generatePivotChart);

    // Unified filter provision for pivot chart
    initUnifiedFilters('pchart', 'pcSelectFiltersBtn', 'pcSelectFiltersDropdown', 'pcFiltersGrid', PC_DEFAULT_FIELDS, null);

    const clearBtn = document.getElementById('pcClearFilters');
    if (clearBtn) clearBtn.addEventListener('click', () => clearUnifiedFilters('pchart', 'pcFiltersGrid', null));
}

function generatePivotChart() {
    const category = document.getElementById('pcCategory').value;
    const series = document.getElementById('pcSeries').value;
    const agg = document.getElementById('pcAggregation').value;
    const chartType = document.getElementById('pcChartType').value;
    const topN = document.getElementById('pcTopN').value;
    const title        = document.getElementById('pivotChartTitle').value || 'Chart';
    const xLabel       = (document.getElementById('pcXLabel') || {}).value || '';
    const yLabel       = (document.getElementById('pcYLabel') || {}).value || '';
    const chartColor   = (document.getElementById('pcChartColor')    || {}).value || '#2563eb';
    const titleFontSz  = parseInt((document.getElementById('pcTitleFontSize') || {}).value || '16');
    const titleFontFam = (document.getElementById('pcTitleFont')     || {}).value || 'inherit';
    const titleClr     = (document.getElementById('pcTitleColor')    || {}).value || '#1c1917';
    const labelFontSz  = parseInt((document.getElementById('pcLabelFontSize') || {}).value || '11');
    const labelClr     = (document.getElementById('pcLabelColor')    || {}).value || '#6b3a00';

    // Post Status filter
    const pcPsOccupied = document.getElementById('pcPsOccupied');
    const pcPsVacant   = document.getElementById('pcPsVacant');
    const pcShowOcc    = !pcPsOccupied || pcPsOccupied.checked;
    const pcShowVac    = pcPsVacant && pcPsVacant.checked;

    const pcFilters = getUnifiedFilterValues('pchart');
    let data = getEmpTypeData().filter(r => {
        if (!passesFilters(r, pcFilters)) return false;
        if (r.post_status === 'occupied') return pcShowOcc;
        if (r.post_status === 'vacant')   return pcShowVac;
        return false;
    });

    // BG color applied live to container
    const bgColor = (document.getElementById('pcBgColor') || {}).value || '#fefce8';
    applyChartBg(bgColor);

    // Multi-color & data-labels options
    const multiColor      = !!(document.getElementById('pcMultiColor') || {}).checked;
    const showDataLabels  = !!(document.getElementById('pcShowDataLabels') || {}).checked;
    const dataLabelSz     = parseInt((document.getElementById('pcDataLabelSize') || {}).value || '11');
    const dataLabelMode   = (document.getElementById('pcDataLabelMode') || {}).value || 'count';

    if (!category) { alert('Please select a Category field.'); return; }

    if (pivotChartInstance) pivotChartInstance.destroy();

    const canvas = document.getElementById('pivotChartCanvas');
    document.getElementById('pivotChartEmpty').style.display = 'none';
    canvas.style.display = 'block';

    const isPie = ['pie', 'doughnut', 'polarArea', 'radar'].includes(chartType);
    const isHorizontal = chartType === 'horizontalBar' || chartType === 'stackedHBar';
    const isStacked = chartType === 'stackedBar' || chartType === 'stackedHBar';

    // Helper: get label string for a record field, formatting Date objects properly
    const DATE_FIELD_KEYS = new Set(['_doj', '_dob', '_dor', '_presentPostSince']);
    function getLabel(r, field) {
        if (DATE_FIELD_KEYS.has(field)) {
            return r[field] ? formatDate(r[field]) : 'N/A';
        }
        return String(r[field] || 'N/A');
    }

    // Build axis scale config with optional title labels and user font/color settings
    function buildScales(stacked) {
        if (isPie) return {};
        const gridColor = '#e8b86d';
        return {
            x: {
                stacked,
                ticks: { color: labelClr, font: { size: labelFontSz } },
                grid: { color: gridColor },
                title: xLabel ? { display: true, text: xLabel, color: labelClr, font: { size: labelFontSz + 1, weight: '600' } } : { display: false }
            },
            y: {
                stacked,
                ticks: { color: labelClr, font: { size: labelFontSz } },
                grid: { color: gridColor },
                title: yLabel ? { display: true, text: yLabel, color: labelClr, font: { size: labelFontSz + 1, weight: '600' } } : { display: false }
            }
        };
    }

    if (!series) {
        // Simple aggregation
        const counts = {};
        data.forEach(r => {
            const cat = getLabel(r, category);
            if (agg === 'count') counts[cat] = (counts[cat] || 0) + 1;
            else {
                if (!counts[cat]) counts[cat] = new Set();
                counts[cat].add(r.emp_id);
            }
        });

        let entries = Object.entries(counts).map(([k, v]) => [k, agg === 'distinct' ? v.size : v]);
        entries.sort((a, b) => b[1] - a[1]);
        if (topN !== 'all') entries = entries.slice(0, parseInt(topN));

        const labels = entries.map(e => e[0]);
        const values = entries.map(e => e[1]);
        const colors = isPie ? getColors(labels.length) : null;

        const total = values.reduce((s, v) => s + v, 0);
        const dlFormatter = (value) => {
            const pct = total > 0 ? Math.round(value / total * 100) + '%' : '0%';
            if (dataLabelMode === 'pct') return pct;
            if (dataLabelMode === 'both') return value + '\n' + pct;
            return value;
        };
        const dlConfig = showDataLabels
            ? { display: true, color: labelClr, font: { size: dataLabelSz, weight: '600' }, formatter: dlFormatter,
                anchor: isPie ? 'center' : 'end', align: isPie ? 'center' : 'end', clamp: true }
            : { display: false };
        const bgColors = multiColor && !isPie ? getColors(values.length).map(c => c + 'bb') : (isPie ? colors : chartColor + 'bb');
        const bdColors = multiColor && !isPie ? getColors(values.length) : (isPie ? colors : chartColor);

        pivotChartInstance = new Chart(canvas, {
            plugins: [window.ChartDataLabels].filter(Boolean),
            type: (isHorizontal || isStacked) ? 'bar' : chartType,
            data: {
                labels,
                datasets: [{
                    label: title,
                    data: values,
                    backgroundColor: bgColors,
                    borderColor: bdColors,
                    borderWidth: 1,
                    borderRadius: isPie ? 0 : 4
                }]
            },
            options: {
                ...chartOptions(title),
                indexAxis: isHorizontal ? 'y' : 'x',
                scales: buildScales(isStacked),
                plugins: {
                    ...chartOptions(title).plugins,
                    title: { display: true, text: title, color: titleClr, font: { size: titleFontSz, weight: '700', family: titleFontFam } },
                    legend: { display: isPie, position: 'right', labels: { color: labelClr, font: { size: labelFontSz } } },
                    datalabels: dlConfig
                }
            }
        });
    } else {
        // Grouped / stacked
        const catSet = new Set();
        const serSet = new Set();
        const matrix = {};

        data.forEach(r => {
            const cat = getLabel(r, category);
            const ser = getLabel(r, series);
            catSet.add(cat);
            serSet.add(ser);
            if (!matrix[cat]) matrix[cat] = {};
            matrix[cat][ser] = (matrix[cat][ser] || 0) + 1;
        });

        // Sort categories by total descending
        let catEntries = [...catSet].map(c => {
            const total = Object.values(matrix[c] || {}).reduce((s, v) => s + v, 0);
            return [c, total];
        }).sort((a, b) => b[1] - a[1]);
        if (topN !== 'all') catEntries = catEntries.slice(0, parseInt(topN));
        const labels = catEntries.map(e => e[0]);
        const seriesValues = [...serSet].sort();
        const colors = getColors(seriesValues.length);

        const datasets = seriesValues.map((sv, si) => ({
            label: sv,
            data: labels.map(l => (matrix[l] && matrix[l][sv]) || 0),
            backgroundColor: colors[si] + 'bb',
            borderColor: colors[si],
            borderWidth: 1,
            borderRadius: isStacked ? 0 : 3
        }));

        const dlFormatterGrouped = (value) => {
            if (dataLabelMode === 'pct') {
                const ds = datasets.find(d => d.data.includes(value));
                const dsTotal = ds ? ds.data.reduce((s, v) => s + v, 0) : 1;
                return dsTotal > 0 ? Math.round(value / dsTotal * 100) + '%' : '0%';
            }
            if (dataLabelMode === 'both') {
                const ds = datasets.find(d => d.data.includes(value));
                const dsTotal = ds ? ds.data.reduce((s, v) => s + v, 0) : 1;
                const pct = dsTotal > 0 ? Math.round(value / dsTotal * 100) + '%' : '0%';
                return value + '\n' + pct;
            }
            return value;
        };
        const dlConfigGrouped = showDataLabels
            ? { display: true, color: labelClr, font: { size: dataLabelSz, weight: '600' },
                formatter: dlFormatterGrouped, anchor: 'end', align: 'end', clamp: true }
            : { display: false };

        pivotChartInstance = new Chart(canvas, {
            plugins: [window.ChartDataLabels].filter(Boolean),
            type: 'bar',
            data: { labels, datasets },
            options: {
                ...chartOptions(title),
                indexAxis: isHorizontal ? 'y' : 'x',
                scales: buildScales(isStacked),
                plugins: {
                    ...chartOptions(title).plugins,
                    title: { display: true, text: title, color: titleClr, font: { size: titleFontSz, weight: '700', family: titleFontFam } },
                    legend: { display: true, position: 'top', labels: { color: labelClr, font: { size: labelFontSz } } },
                    datalabels: dlConfigGrouped
                }
            }
        });
    }
}

function buildExportCanvas() {
    const srcCanvas = document.getElementById('pivotChartCanvas');
    // A4 landscape at 96 DPI: 297mm × 210mm → 1122 × 794 px
    const A4_W = 1122;
    const A4_H = 794;
    const exp = document.createElement('canvas');
    exp.width = A4_W;
    exp.height = A4_H;
    const ctx = exp.getContext('2d');

    // Background color from picker or default
    const bgEl = document.getElementById('pcBgColor');
    ctx.fillStyle = bgEl ? bgEl.value : '#fefce8';
    ctx.fillRect(0, 0, A4_W, A4_H);

    // Draw chart fitted proportionally with padding
    const pad = 20;
    const availW = A4_W - pad * 2;
    const availH = A4_H - pad * 2;
    const scaleW = availW / srcCanvas.width;
    const scaleH = availH / srcCanvas.height;
    const scale = Math.min(scaleW, scaleH);
    const drawW = srcCanvas.width * scale;
    const drawH = srcCanvas.height * scale;
    const drawX = pad + (availW - drawW) / 2;
    const drawY = pad + (availH - drawH) / 2;
    ctx.drawImage(srcCanvas, drawX, drawY, drawW, drawH);
    return exp;
}

function exportPivotChartImage(format) {
    if (!pivotChartInstance) { alert('Generate a chart first.'); return; }
    const title = document.getElementById('pivotChartTitle').value || 'Chart';
    const exp = buildExportCanvas();
    const link = document.createElement('a');
    link.download = title.replace(/[^a-zA-Z0-9]/g, '_') + '.' + format;
    link.href = exp.toDataURL('image/' + format, 1.0);
    link.click();
}

function exportPivotChartPDF() {
    if (!pivotChartInstance) { alert('Generate a chart first.'); return; }
    const title = document.getElementById('pivotChartTitle').value || 'Chart';
    const exp = buildExportCanvas();
    const imgData = exp.toDataURL('image/jpeg', 0.95);

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

    const pageW = 297;
    const pageH = 210;
    const margin = 12;

    // ── Dynamic header height based on title wrapping ──
    const CHART_TITLE_SZ  = 15;
    const CHART_TITLE_LH  = CHART_TITLE_SZ * 0.352778 * 1.4; // pt → mm with line spacing
    doc.setFontSize(CHART_TITLE_SZ);
    const chartTitleLines = doc.splitTextToSize(title, pageW - 2 * margin);
    const hdrH = Math.max(22, 4 + chartTitleLines.length * CHART_TITLE_LH + 6); // +6 for subtitle

    // Header background
    doc.setFillColor(123, 28, 0);
    doc.rect(0, 0, pageW, hdrH, 'F');

    // Wrapped title lines
    doc.setFontSize(CHART_TITLE_SZ);
    doc.setTextColor(245, 200, 66);
    const chartTitleStartY = 4 + CHART_TITLE_LH;
    chartTitleLines.forEach((line, i) => {
        doc.text(line, pageW / 2, chartTitleStartY + i * CHART_TITLE_LH, { align: 'center' });
    });

    // Subtitle — placed below last title line
    const subtitleY = chartTitleStartY + (chartTitleLines.length - 1) * CHART_TITLE_LH + 5;
    doc.setFontSize(8);
    doc.setTextColor(255, 235, 180);
    doc.text('TTD HR Analytics  |  Generated: ' + new Date().toLocaleString(), pageW / 2, subtitleY, { align: 'center' });

    // Gold separator line
    doc.setDrawColor(200, 134, 10);
    doc.setLineWidth(0.8);
    doc.line(margin, hdrH + 1, pageW - margin, hdrH + 1);

    // Chart image — fit within remaining space
    const imgY = hdrH + 4;
    const availH = pageH - imgY - margin;
    const availW = pageW - 2 * margin;
    const aspect = exp.width / exp.height;
    let drawW = availW;
    let drawH = availW / aspect;
    if (drawH > availH) { drawH = availH; drawW = availH * aspect; }
    const drawX = margin + (availW - drawW) / 2;

    doc.addImage(imgData, 'JPEG', drawX, imgY, drawW, drawH);

    // Footer
    doc.setFontSize(7);
    doc.setTextColor(154, 113, 80);
    doc.text('\u00A9 Muddada Ravi Chandra IAS, EO, TTD', pageW / 2, pageH - 4, { align: 'center' });

    doc.save(title.replace(/[^a-zA-Z0-9]/g, '_') + '_Chart.pdf');
}

// ──────────────────────────────────────────────────────────────
//   SERVICE DURATION HELPERS
// ──────────────────────────────────────────────────────────────

// March 14 2026 is the data cut-off date — treat as "today"
function _isDataCutoff(epochMs) {
    if (!epochMs) return false;
    const d = new Date(epochMs);
    return d.getFullYear() === 2026 && d.getMonth() === 2 && d.getDate() === 14;
}

function _serviceDuration(fromMs, toMs) {
    if (!fromMs) return '—';
    const from = new Date(fromMs);
    const to   = toMs ? new Date(toMs) : new Date();
    let years  = to.getFullYear() - from.getFullYear();
    let months = to.getMonth()    - from.getMonth();
    if (to.getDate() < from.getDate()) months--;
    if (months < 0) { years--; months += 12; }
    if (years < 0)  return '—';
    if (years === 0 && months === 0) return '< 1 Month';
    if (years === 0) return `${months} Month${months !== 1 ? 's' : ''}`;
    if (months === 0) return `${years} Year${years !== 1 ? 's' : ''}`;
    return `${years} Year${years !== 1 ? 's' : ''} ${months} Month${months !== 1 ? 's' : ''}`;
}

// ──────────────────────────────────────────────────────────────
//   EMPLOYEE PHOTO HELPERS
// ──────────────────────────────────────────────────────────────

const _PHOTO_EXTS = ['jpg','JPG','jpeg','JPEG','png','PNG','webp'];

function empPhotoTryNext(img, idx, empId, gender) {
    const next = idx + 1;
    if (next < _PHOTO_EXTS.length) {
        img.onerror = () => empPhotoTryNext(img, next, empId, gender);
        img.src = `Employee_Photos/${empId}.${_PHOTO_EXTS[next]}`;
    } else {
        const g = (gender || '').toLowerCase();
        const isFemale = g === 'female' || g === 'f';
        const bg    = isFemale ? '#fce7f3' : '#dbeafe';
        const fill  = isFemale ? '#be185d' : '#1d4ed8';
        const svgPath = isFemale
            ? '<circle cx="12" cy="7" r="4"/><path d="M6 21v-1a6 6 0 0 1 12 0v1"/><path d="M9 11.5l1.5 4 1.5-3 1.5 3 1.5-4"/>'
            : '<circle cx="12" cy="8" r="4"/><path d="M4 20c0-4 3.6-7 8-7s8 3 8 7"/>';
        const size = img.dataset.thumbSize === '1' ? 36 : 52;
        img.outerHTML = `<div class="emp-photo-placeholder" style="background:${bg}">
            <svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="${fill}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">${svgPath}</svg>
            <span style="color:${fill};font-size:0.7rem;font-weight:700;margin-top:2px">No Photo</span>
        </div>`;
    }
}

function _empPhotoHtml(empId, gender, thumbSize) {
    const thumb = thumbSize ? ' data-thumb-size="1"' : '';
    return `<img src="Employee_Photos/${empId}.${_PHOTO_EXTS[0]}" alt="${empId}"${thumb}
        onerror="empPhotoTryNext(this,0,'${empId}','${gender||''}')"
        style="width:100%;height:100%;object-fit:cover;object-position:top center;display:block">`;
}

// ──────────────────────────────────────────────────────────────
//   EMPLOYEE DETAIL POPUP
// ──────────────────────────────────────────────────────────────

function openEmployeePopup(emp) {
    const modal = document.getElementById('empModalOverlay');
    const title = document.getElementById('empModalTitle');
    const body = document.getElementById('empModalBody');

    title.textContent = 'Employee Details';

    const psClass = emp.is_occupied ? 'occupied' : 'vacant';
    const psLabel = emp.is_occupied ? '✅ Occupied' : '⚠️ Vacant';

    // Photo banner
    let html = `<div class="emp-photo-banner">
        <div class="emp-photo-wrap">${_empPhotoHtml(emp.emp_id, emp.gender)}</div>
        <div class="emp-photo-info">
            <div class="emp-photo-name">${emp.emp_name || '—'}</div>
            <div class="emp-photo-desig">${emp.designation_name || '—'}</div>
            <div class="emp-photo-meta">
                <span class="emp-photo-badge">${emp.emp_id || ''}</span>
                <span class="emp-photo-badge ${psClass}">${psLabel}</span>
                ${emp.gender ? `<span class="emp-photo-badge">${emp.gender}</span>` : ''}
                ${emp.work_location ? `<span class="emp-photo-badge">📍 ${emp.work_location}</span>` : ''}
                ${emp.age ? `<span class="emp-photo-badge">Age: ${emp.age} yrs</span>` : ''}
            </div>
        </div>
    </div>`;

    const fields = [
        { label: 'Employee ID', val: emp.emp_id },
        { label: 'Employee Name', val: emp.emp_name },
        { label: 'Designation', val: emp.designation_name },
        { label: 'Post Status', val: emp.is_occupied ? '✅ Occupied' : '⚠️ Vacant' },
        { label: 'Gender', val: emp.gender },
        { label: 'Age', val: emp.age ? emp.age + ' years' : null },
        { label: 'Department (HOD)', val: emp.hod_name },
        { label: 'Section Head', val: emp.hos_name },
        { label: 'Section', val: emp.section_name },
        { label: 'Work Location', val: emp.work_location },
        { label: 'Community', val: emp.community },
        { label: 'Sub Community', val: emp.sub_community },
        { label: 'Caste', val: emp.caste },
        { label: 'Date of Birth', val: emp.dob_fmt },
        { label: 'Date of Joining', val: emp.doj_fmt },
        { label: 'Date of Retirement', val: emp.dor_fmt },
        { label: 'In Present Post Since', val: emp.working_in_present_place_since_fmt },
        { label: 'Service (Yrs) Current Post', val: emp.service_years_current_post != null ? parseFloat(emp.service_years_current_post).toFixed(1) + ' yrs' : null },
        { label: 'Total Service (Yrs)', val: emp.total_service_years != null ? emp.total_service_years + ' yrs' : null },
        { label: 'Recruitment Type', val: emp.type_of_recruitment },
        { label: 'Recruited Post', val: emp.recruited_post },
        { label: 'Joined During', val: emp.joined_during },
        { label: 'Present Post Posted In', val: emp.present_post_by },
        { label: 'Native District', val: emp.native_dist },
        { label: 'Local/Study District', val: emp.local_dist_study },
        { label: 'Mobile', val: emp.mobile }
    ];

    html += '<div class="emp-details-grid">';
    fields.forEach(f => {
        const v = f.val && f.val !== 'null' && f.val !== 'None' && f.val !== '' ? f.val : '—';
        html += `<div class="emp-detail-item">
            <span class="emp-detail-label">${f.label}</span>
            <span class="emp-detail-value">${v}</span>
        </div>`;
    });
    html += '</div>';

    // Service history
    const services = SERVICE_DATA.filter(s =>
        String(s.employee_code) === String(emp.emp_id)
    );
    services.sort((a, b) => (a.service_from_date || 0) - (b.service_from_date || 0));

    html += `<div class="emp-service-title">📋 Service History (${services.length} records)</div>`;
    if (services.length > 0) {
        html += `<table class="emp-service-table">
            <thead><tr><th>#</th><th>Designation</th><th>Office / Place</th><th>From</th><th>To</th><th>Service</th></tr></thead>
            <tbody>`;
        services.forEach((s, i) => {
            const fromDate   = s.service_from_date ? formatDate(epochToDate(s.service_from_date)) : '—';
            const isPresent  = !s.service_to_date || _isDataCutoff(s.service_to_date);
            const toDate     = isPresent ? '<span class="svc-today">Till Today</span>' : formatDate(epochToDate(s.service_to_date));
            const toMs       = isPresent ? Date.now() : s.service_to_date;
            const duration   = _serviceDuration(s.service_from_date, toMs);
            html += `<tr>
                <td>${i + 1}</td>
                <td>${s.designation_name || '—'}</td>
                <td>${s.service_office || s.place_of_posting || '—'}</td>
                <td>${fromDate}</td>
                <td>${toDate}</td>
                <td><span class="svc-duration">${duration}</span></td>
            </tr>`;
        });
        html += '</tbody></table>';
    } else {
        html += '<p style="color:var(--text-muted);font-size:0.85rem;margin-top:8px">No service records found for this employee.</p>';
    }

    body.innerHTML = html;
    modal.classList.add('active');
}

function initEmployeeModal() {
    const overlay = document.getElementById('empModalOverlay');
    if (!overlay) return;
    document.getElementById('empModalClose').addEventListener('click', () => overlay.classList.remove('active'));
    overlay.addEventListener('click', (e) => {
        if (e.target === overlay) overlay.classList.remove('active');
    });
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') overlay.classList.remove('active');
    });
}

// ──────────────────────────────────────────────────────────────
//   SERVICE RECORDS PAGE
// ──────────────────────────────────────────────────────────────

function initServicePage() {
    document.getElementById('btnServiceSearch').addEventListener('click', searchServiceRecords);
    document.getElementById('fServiceSearch').addEventListener('keypress', (e) => {
        if (e.key === 'Enter') searchServiceRecords();
    });
}

function searchServiceRecords() {
    const query = document.getElementById('fServiceSearch').value.trim().toLowerCase();
    if (!query) return;

    const container = document.getElementById('serviceResults');

    // Find matching employees
    const matches = getEmpTypeData().filter(r => {
        if (!r.is_occupied) return false;
        const id = String(r.emp_id || '').toLowerCase();
        const name = String(r.emp_name || '').toLowerCase();
        return id === query || name.includes(query) || id.includes(query);
    }).slice(0, 20);

    if (matches.length === 0) {
        container.innerHTML = `<div class="pivot-empty-state"><div class="empty-icon">🔍</div><h3>No employees found</h3><p>Try a different ID or name.</p></div>`;
        return;
    }

    let html = '';
    matches.forEach(emp => {
        const services = SERVICE_DATA.filter(s => s.employee_code === emp.emp_id || s.employee_code === parseInt(emp.emp_id));
        services.sort((a, b) => (a.service_from_date || 0) - (b.service_from_date || 0));

        html += `<div class="service-card">
            <div class="service-card-header">
                <div class="service-avatar">${emp.gender === 'Female' ? '👩' : '👨'}</div>
                <div class="service-emp-info">
                    <h3>${emp.emp_name} <small style="color:var(--text-muted)">(ID: ${emp.emp_id})</small></h3>
                    <p>${emp.designation_name} | ${emp.hod_name} | ${emp.work_location}</p>
                    <p>DOJ: ${emp.doj_fmt} | DOR: ${emp.dor_fmt} | Community: ${emp.community} | Gender: ${emp.gender}</p>
                </div>
            </div>
            <h4 style="margin-bottom:12px; color:var(--text-secondary)">Service History (${services.length} records)</h4>
            <div class="service-timeline">`;

        if (services.length > 0) {
            services.forEach(s => {
                html += `<div class="timeline-item">
                    <h4>${s.designation_name || '—'}</h4>
                    <p>${s.service_office || '—'}</p>
                    <div class="timeline-dates">${formatDate(s._from)} → ${formatDate(s._to)}</div>
                </div>`;
            });
        } else {
            html += `<p style="color:var(--text-muted)">No service records found for this employee.</p>`;
        }

        html += `</div></div>`;
    });

    container.innerHTML = html;
}

// ── Start ──
document.addEventListener('DOMContentLoaded', loadData);
