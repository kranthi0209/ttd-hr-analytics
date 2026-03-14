/* ================================================================
   TTD HR Analytics Dashboard – Main Application Script
   ================================================================ */

// ── Global State ──
let RAW_DATA = [];
let SERVICE_DATA = [];
let FILTERED_DATA = [];
let OCCUPIED_DATA = [];

let dashCharts = {};
let pivotChartInstance = null;
let currentPage = 1;
let sortCol = null;
let sortDir = 'asc';
let visibleColumns = {};
let tableFilteredData = [];

const TABLE_COLUMNS = [
    { key: 'emp_id', label: 'Emp ID', show: true },
    { key: 'emp_name', label: 'Employee Name', show: true },
    { key: 'designation_name', label: 'Designation', show: true },
    { key: 'hod_name', label: 'Department (HOD)', show: true },
    { key: 'hos_name', label: 'Section Head', show: false },
    { key: 'section_name', label: 'Section', show: false },
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
            fetch('Regular_Employees.json'),
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

        r.is_occupied = r.post_status === 'occupied';
        
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
    populateFilters();
    applyDashboardFilters();
    initDataTable();
    initPivotTable();
    initPivotChart();
    initEmployeeModal();
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
    });

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
        const vals = [...new Set(RAW_DATA.filter(r => r.is_occupied).map(r => String(r[fk] || '')).filter(v => v && v !== 'None' && v !== 'null'))].sort();

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

function getBaseData(prefix) {
    if (prefix === 'dash' || prefix === 'pchart') return RAW_DATA.filter(r => r.is_occupied);
    return RAW_DATA;
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

// Navigate to data table, copying all active dashboard filters + the clicked field/value
function navigateToTableWithFilter(fieldKey, fieldValue) {
    const tableNav = document.querySelector('[data-page="datatable"]');
    if (tableNav) tableNav.click();
    if (!UNIFIED_FILTERS['table'] || !UNIFIED_FILTERS['dash']) return;

    // Copy all active dashboard filter state into table filter state
    ALL_FILTER_FIELDS.forEach(f => {
        const dashState = UNIFIED_FILTERS['dash'][f.key];
        const tableState = UNIFIED_FILTERS['table'][f.key];
        if (!dashState || !tableState) return;
        if (f.type === 'date') {
            tableState.from = dashState.from;
            tableState.to = dashState.to;
            if (dashState.from || dashState.to) tableState.enabled = true;
        } else {
            if (dashState.selected && dashState.selected.size > 0) {
                tableState.enabled = true;
                tableState.selected = new Set(dashState.selected);
            }
        }
    });

    // Now apply the clicked chart filter on top
    if (UNIFIED_FILTERS['table'][fieldKey]) {
        UNIFIED_FILTERS['table'][fieldKey].enabled = true;
        UNIFIED_FILTERS['table'][fieldKey].selected = new Set([String(fieldValue)]);
    }

    // Sync the dropdown checkboxes
    const dd = document.getElementById('tableSelectFiltersDropdown');
    if (dd) {
        ALL_FILTER_FIELDS.forEach(f => {
            const cb = dd.querySelector(`input[value="${f.key}"]`);
            if (cb && UNIFIED_FILTERS['table'][f.key] && UNIFIED_FILTERS['table'][f.key].enabled) cb.checked = true;
        });
    }

    renderUnifiedFilterGrid('table', 'tableFiltersGrid', applyTableFilters);
    applyTableFilters();
}

// ── Filter Logic ──
function getFilteredOccupied(hods, comms, genders, locs, recruits, joineds, postedBys, svcMin, svcMax, extraFilters = {}) {
    return RAW_DATA.filter(r => {
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

    const filteredAll = RAW_DATA.filter(r => passesFilters(r, filters));

    OCCUPIED_DATA = RAW_DATA.filter(r => {
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
    const total = allData.length;
    const occupied = allData.filter(r => r.is_occupied).length;
    const vacant = allData.filter(r => !r.is_occupied).length;
    const depts = new Set(allData.map(r => r.hod_name)).size;
    const desigs = new Set(occupiedData.map(r => r.designation_name)).size;
    const avgSvc = occupiedData.length > 0
        ? (occupiedData.reduce((s, r) => s + (r.total_service_years || 0), 0) / occupiedData.length).toFixed(1)
        : 0;

    const origTotal = RAW_DATA.length;
    const origOccupied = RAW_DATA.filter(r => r.is_occupied).length;
    const origVacant = RAW_DATA.filter(r => !r.is_occupied).length;
    const origDepts = new Set(RAW_DATA.map(r => r.hod_name)).size;
    const origOccupiedData = RAW_DATA.filter(r => r.is_occupied);
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

    // 1. Department Strength (top 15 horizontal bar)
    dashCharts.dept = createCountChart('chartDeptStrength', data, 'hod_name', 'bar', 15, true);

    // 2. Community (doughnut)
    dashCharts.community = createCountChart('chartCommunity', data, 'community', 'doughnut');

    // 3. Gender (pie)
    dashCharts.gender = createCountChart('chartGender', data, 'gender', 'pie');

    // 4. Recruitment Type (bar)
    dashCharts.recruitment = createCountChart('chartRecruitment', data, 'type_of_recruitment', 'bar');

    // 5. Joined During (bar)
    dashCharts.joined = createCountChart('chartJoined', data, 'joined_during', 'bar');

    // 6. Age Distribution (histogram)
    const ageBuckets = {};
    data.forEach(r => {
        if (r.age) {
            const bucket = Math.floor(r.age / 5) * 5;
            const label = `${bucket}-${bucket + 4}`;
            ageBuckets[label] = (ageBuckets[label] || 0) + 1;
        }
    });
    const ageSorted = Object.entries(ageBuckets).sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
    dashCharts.age = new Chart(document.getElementById('chartAge'), {
        type: 'bar',
        data: {
            labels: ageSorted.map(e => e[0]),
            datasets: [{
                label: 'Employees',
                data: ageSorted.map(e => e[1]),
                backgroundColor: CHART_COLORS[0] + '99',
                borderColor: CHART_COLORS[0],
                borderWidth: 1,
                borderRadius: 4
            }]
        },
        options: chartOptions('Age Distribution')
    });

    // 7. Retirement Forecast
    const now = new Date();
    const retYears = {};
    for (let y = now.getFullYear(); y <= now.getFullYear() + 5; y++) retYears[y] = 0;
    data.forEach(r => {
        if (r.retirement_year && retYears.hasOwnProperty(r.retirement_year)) {
            retYears[r.retirement_year]++;
        }
    });
    dashCharts.retirement = new Chart(document.getElementById('chartRetirement'), {
        type: 'bar',
        data: {
            labels: Object.keys(retYears),
            datasets: [{
                label: 'Retirements',
                data: Object.values(retYears),
                backgroundColor: CHART_COLORS[2] + '99',
                borderColor: CHART_COLORS[2],
                borderWidth: 1,
                borderRadius: 4
            }]
        },
        options: chartOptions('Retirement Forecast')
    });

    // 8. Top 15 Designations
    dashCharts.designations = createCountChart('chartDesignations', data, 'designation_name', 'bar', 15, true);

    // 9. Posted During (Year posted to present post)
    const postedYearCounts = {};
    data.forEach(r => {
        if (r._presentPostSince) {
            const yr = r._presentPostSince.getFullYear();
            postedYearCounts[yr] = (postedYearCounts[yr] || 0) + 1;
        }
    });
    const postedYearEntries = Object.entries(postedYearCounts).sort((a,b) => a[0] - b[0]);
    dashCharts.postedDuring = new Chart(document.getElementById('chartPostedDuring'), {
        type: 'bar',
        data: {
            labels: postedYearEntries.map(e => e[0]),
            datasets: [{
                label: 'Employees Posted',
                data: postedYearEntries.map(e => e[1]),
                backgroundColor: CHART_COLORS[4] + '99',
                borderColor: CHART_COLORS[4],
                borderWidth: 1,
                borderRadius: 4
            }]
        },
        options: chartOptions('Posted During (Year)')
    });

    // 10. Posted By
    dashCharts.postedBy = createCountChart('chartPostedBy', data, 'present_post_by', 'doughnut');
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
                    callbacks: isFilterable ? { footer: () => '🔍 Click to filter Data Table' } : {}
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

    tableFilteredData = RAW_DATA.filter(r => {
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
    const ps = parseInt(document.getElementById('pageSize').value);
    const totalPages = Math.max(1, Math.ceil(tableFilteredData.length / ps));
    if (currentPage > totalPages) currentPage = totalPages;
    const start = (currentPage - 1) * ps;
    const pageData = tableFilteredData.slice(start, start + ps);

    // Header
    const thead = document.getElementById('dataTableHead');
    thead.innerHTML = '<tr>' + cols.map(c => {
        const icon = sortCol === c.key ? (sortDir === 'asc' ? '▲' : '▼') : '⇅';
        return `<th data-col="${c.key}">${c.label} <span class="sort-icon">${icon}</span></th>`;
    }).join('') + '</tr>';

    thead.querySelectorAll('th').forEach(th => {
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
    const cols = TABLE_COLUMNS.filter(c => visibleColumns[c.key]);
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
    const cols = TABLE_COLUMNS.filter(c => visibleColumns[c.key]);

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
    { key: 'post_status', label: 'Post Status' },
    { key: 'retirement_year', label: 'Retirement Year' },
    { key: 'recruited_post', label: 'Recruited Post' }
];

const pivotState = { rows: [], cols: [] };
const pivotFilterFields = {}; // key -> { label, enabled }
const MULTISELECT_PIVOT_STATE = {};

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
        const vals = [...new Set(RAW_DATA.filter(r => r.is_occupied).map(r => String(r[f.key] || '')).filter(v => v && v !== 'None' && v !== 'null'))].sort();
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

    const occupied = RAW_DATA.filter(r => r.is_occupied);
    // Apply filters from filter bar
    let data = occupied.filter(r => {
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

    // Build table
    const table = document.getElementById('pivotTable');
    const thead = document.getElementById('pivotTableHead');
    const tbody = document.getElementById('pivotTableBody');

    const rowLabel = document.getElementById('pivotRowLabel').value || rowFields.map(f => PIVOT_FIELDS.find(pf => pf.key === f)?.label || f).join(' / ');
    const colLabel = document.getElementById('pivotColLabel').value || colFields.map(f => PIVOT_FIELDS.find(pf => pf.key === f)?.label || f).join(' × ');
    const valLabel = document.getElementById('pivotValueLabel').value || 'Count';

    // Header
    let headerHtml = '<tr>';
    rowFields.forEach((f, i) => {
        headerHtml += `<th>${rowLabel.split('/')[i] || PIVOT_FIELDS.find(pf => pf.key === f)?.label || f}</th>`;
    });
    displayCols.forEach(c => { headerHtml += `<th>${c}</th>`; });
    headerHtml += '<th>Grand Total</th></tr>';
    thead.innerHTML = headerHtml;

    // Body
    const grandTotals = {};
    displayCols.forEach(c => grandTotals[c] = 0);
    let grandGrand = 0;

    let bodyHtml = '';
    sortedRowKeys.forEach(rk => {
        const parts = rk.split('|||');
        let rowTotal = 0;
        bodyHtml += '<tr>';
        parts.forEach(p => { bodyHtml += `<td style="text-align:left; font-family:var(--font-main); font-weight:500">${p}</td>`; });
        displayCols.forEach(c => {
            const val = rowGroups[rk][c] || 0;
            rowTotal += val;
            grandTotals[c] = (grandTotals[c] || 0) + val;
            bodyHtml += `<td>${val}</td>`;
        });
        grandGrand += rowTotal;
        bodyHtml += `<td style="font-weight:700">${rowTotal}</td>`;
        bodyHtml += '</tr>';
    });

    // Grand total row
    bodyHtml += '<tr>';
    rowFields.forEach((f, i) => { bodyHtml += `<td style="text-align:left; font-weight:700">${i === 0 ? 'Grand Total' : ''}</td>`; });
    displayCols.forEach(c => { bodyHtml += `<td>${grandTotals[c] || 0}</td>`; });
    bodyHtml += `<td>${grandGrand}</td></tr>`;

    tbody.innerHTML = bodyHtml;

    // Add sort to pivot table headers
    thead.querySelectorAll('th').forEach((th, idx) => {
        th.addEventListener('click', () => {
            sortPivotTable(idx, th);
        });
    });

    document.getElementById('pivotEmpty').style.display = 'none';
    table.style.display = 'table';
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

function exportPivotPDF() {
    const table = document.getElementById('pivotTable');
    if (table.style.display === 'none') { alert('Generate a pivot table first.'); return; }
    const title = document.getElementById('pivotTitle').value || 'Pivot Report';

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

    doc.setFontSize(14);
    doc.text(title, 14, 15);
    doc.setFontSize(8);
    doc.text(`Generated: ${new Date().toLocaleString()}`, 14, 21);

    const rows = table.querySelectorAll('tr');
    const head = [];
    const body = [];
    rows.forEach((row, ri) => {
        const cells = Array.from(row.cells).map(c => c.textContent);
        if (ri === 0) head.push(cells);
        else body.push(cells);
    });

    if (doc.autoTable) {
        doc.autoTable({
            head, body, startY: 25,
            styles: { fontSize: 6, cellPadding: 2, halign: 'center', valign: 'middle', lineWidth: 0.2, lineColor: [180, 83, 9] },
            headStyles: { fillColor: [217, 119, 6], halign: 'center', valign: 'middle', fontStyle: 'bold' },
            bodyStyles: { halign: 'center', valign: 'middle' },
            columnStyles: { 0: { halign: 'left' } }
        });
    } else {
        simpleTablePDF(doc, head[0], body, 25);
    }
    doc.save(title.replace(/[^a-zA-Z0-9]/g, '_') + '.pdf');
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
    const title = document.getElementById('pivotChartTitle').value || 'Chart';
    const xLabel = (document.getElementById('pcXLabel') || {}).value || '';
    const yLabel = (document.getElementById('pcYLabel') || {}).value || '';
    const chartColorEl = document.getElementById('pcChartColor');
    const chartColor = chartColorEl ? chartColorEl.value : '#2563eb';

    // Apply unified filters (no default is_occupied restriction — user controls via post_status filter)
    const pcFilters = getUnifiedFilterValues('pchart');
    let data = RAW_DATA.filter(r => passesFilters(r, pcFilters));

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

    // Build axis scale config with optional title labels
    function buildScales(stacked) {
        if (isPie) return {};
        const axisColor = '#6b3a00';
        const gridColor = '#e8b86d';
        return {
            x: {
                stacked,
                ticks: { color: axisColor, font: { size: 10 } },
                grid: { color: gridColor },
                title: xLabel ? { display: true, text: xLabel, color: axisColor, font: { size: 12, weight: '600' } } : { display: false }
            },
            y: {
                stacked,
                ticks: { color: axisColor, font: { size: 10 } },
                grid: { color: gridColor },
                title: yLabel ? { display: true, text: yLabel, color: axisColor, font: { size: 12, weight: '600' } } : { display: false }
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

        pivotChartInstance = new Chart(canvas, {
            type: (isHorizontal || isStacked) ? 'bar' : chartType,
            data: {
                labels,
                datasets: [{
                    label: title,
                    data: values,
                    backgroundColor: isPie ? colors : chartColor + 'bb',
                    borderColor: isPie ? colors : chartColor,
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
                    title: { display: true, text: title, color: '#1c1917', font: { size: 16, weight: '700' } },
                    legend: { display: isPie, position: 'right', labels: { color: '#78716c' } }
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

        pivotChartInstance = new Chart(canvas, {
            type: 'bar',
            data: { labels, datasets },
            options: {
                ...chartOptions(title),
                indexAxis: isHorizontal ? 'y' : 'x',
                scales: buildScales(isStacked),
                plugins: {
                    ...chartOptions(title).plugins,
                    title: { display: true, text: title, color: '#1c1917', font: { size: 16, weight: '700' } },
                    legend: { display: true, position: 'top', labels: { color: '#78716c' } }
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

    // Header background
    doc.setFillColor(123, 28, 0);
    doc.rect(0, 0, pageW, 22, 'F');

    // Title
    doc.setFontSize(15);
    doc.setTextColor(245, 200, 66);
    doc.text(title, pageW / 2, 10, { align: 'center' });

    // Subtitle
    doc.setFontSize(8);
    doc.setTextColor(255, 235, 180);
    doc.text('TTD HR Analytics  |  Generated: ' + new Date().toLocaleString(), pageW / 2, 17, { align: 'center' });

    // Gold separator line
    doc.setDrawColor(200, 134, 10);
    doc.setLineWidth(0.8);
    doc.line(margin, 23, pageW - margin, 23);

    // Chart image — fit within remaining space
    const imgY = 26;
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
//   EMPLOYEE DETAIL POPUP
// ──────────────────────────────────────────────────────────────

function openEmployeePopup(emp) {
    const modal = document.getElementById('empModalOverlay');
    const title = document.getElementById('empModalTitle');
    const body = document.getElementById('empModalBody');

    title.textContent = (emp.emp_name || '—') + '  —  Employee Details';

    const fields = [
        { label: 'Employee ID', val: emp.emp_id },
        { label: 'Employee Name', val: emp.emp_name },
        { label: 'Designation', val: emp.designation_name },
        { label: 'Department (HOD)', val: emp.hod_name },
        { label: 'Section Head', val: emp.hos_name },
        { label: 'Section', val: emp.section_name },
        { label: 'Post Status', val: emp.is_occupied ? '✅ Occupied' : '⚠️ Vacant' },
        { label: 'Community', val: emp.community },
        { label: 'Sub Community', val: emp.sub_community },
        { label: 'Caste', val: emp.caste },
        { label: 'Gender', val: emp.gender },
        { label: 'Work Location', val: emp.work_location },
        { label: 'Date of Birth', val: emp.dob_fmt },
        { label: 'Age', val: emp.age ? emp.age + ' years' : null },
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

    let html = '<div class="emp-details-grid">';
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
            <thead><tr><th>#</th><th>Designation</th><th>Office / Place</th><th>From</th><th>To</th></tr></thead>
            <tbody>`;
        services.forEach((s, i) => {
            const fromDate = s.service_from_date ? formatDate(epochToDate(s.service_from_date)) : '—';
            const toDate = s.service_to_date ? formatDate(epochToDate(s.service_to_date)) : 'Present';
            html += `<tr>
                <td>${i + 1}</td>
                <td>${s.designation_name || '—'}</td>
                <td>${s.service_office || s.place_of_posting || '—'}</td>
                <td>${fromDate}</td>
                <td>${toDate}</td>
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
    const matches = RAW_DATA.filter(r => {
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
