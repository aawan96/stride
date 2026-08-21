// Mobile nav toggle
const navToggle = document.getElementById('navToggle');
const mobileNav = document.getElementById('mobileNav');
if (navToggle && mobileNav) {
  navToggle.addEventListener('click', () => {
    const open = mobileNav.classList.toggle('open');
    navToggle.setAttribute('aria-expanded', String(open));
  });
  mobileNav.querySelectorAll('a').forEach((a) =>
    a.addEventListener('click', () => {
      mobileNav.classList.remove('open');
      navToggle.setAttribute('aria-expanded', 'false');
    })
  );
}

// Inventory filter
const chips = document.querySelectorAll('.chip');
const cards = document.querySelectorAll('#modelGrid .card');
chips.forEach((chip) => {
  chip.addEventListener('click', () => {
    chips.forEach((c) => c.classList.remove('is-active'));
    chip.classList.add('is-active');
    const filter = chip.dataset.filter;
    cards.forEach((card) => {
      const show = filter === 'all' || card.dataset.category === filter;
      card.classList.toggle('is-hidden', !show);
    });
  });
});

// Header shadow on scroll
const headerMain = document.getElementById('headerMain');
if (headerMain) {
  const onScroll = () => headerMain.classList.toggle('scrolled', window.scrollY > 4);
  onScroll();
  window.addEventListener('scroll', onScroll, { passive: true });
}

// Current year in footer
const yearEl = document.getElementById('year');
if (yearEl) yearEl.textContent = new Date().getFullYear();

/* ---------------- Compare models ---------------- */
/* COMPARE_MODELS:START — generated from product-data.xlsx by scripts/generate_products.py */
const COMPARE_MODELS = {
  'glide-s1': {
    name: 'Stride Glide S1', tag: 'Folding · Entry', price: 'PKR 185,000', img: 'images/q5.png',
    wa: 'https://wa.me/923001234567?text=Hi%20Stride%2C%20I%27m%20interested%20in%20the%20Stride%20Glide%20S1.%20Please%20share%20availability%20and%20final%20price.',
    specs: { 'Range': '15 km', 'Top speed': '6 km/h', 'Max load': '100 kg', 'Kerb weight': '24 kg', 'Battery': 'Li-ion 24V 12Ah', 'Motor': 'Brushless 250W', 'Charge time': '6–8 hrs', 'Foldable': 'Yes (one-fold)', 'Suspension': '—', 'Tyres': 'Solid, puncture-free', 'Controls': 'Joystick', 'Recline': '—', 'Warranty': '1 year' },
  },
  'urban-u2': {
    name: 'Stride Urban U2', tag: 'City commuter', price: 'PKR 235,000', img: 'images/q5.png',
    wa: 'https://wa.me/923001234567?text=Hi%20Stride%2C%20I%27m%20interested%20in%20the%20Stride%20Urban%20U2.%20Please%20share%20availability%20and%20final%20price.',
    specs: { 'Range': '22 km', 'Top speed': '8 km/h', 'Max load': '120 kg', 'Kerb weight': '27 kg', 'Battery': 'Li-ion 24V 20Ah', 'Motor': 'Brushless 300W', 'Charge time': '6–8 hrs', 'Foldable': 'Yes', 'Suspension': 'Front', 'Tyres': 'Anti-slip solid', 'Controls': 'Joystick', 'Recline': '—', 'Warranty': '1 year' },
  },
  'cruise-c3': {
    name: 'Stride Cruise C3', tag: 'Comfort · Mid', price: 'PKR 265,000', img: 'images/q5.png',
    wa: 'https://wa.me/923001234567?text=Hi%20Stride%2C%20I%27m%20interested%20in%20the%20Stride%20Cruise%20C3.%20Please%20share%20availability%20and%20final%20price.',
    specs: { 'Range': '25 km', 'Top speed': '8 km/h', 'Max load': '120 kg', 'Kerb weight': '30 kg', 'Battery': 'Li-ion 24V 20Ah', 'Motor': 'Brushless 300W', 'Charge time': '6–8 hrs', 'Foldable': 'Yes', 'Suspension': '—', 'Tyres': 'PU solid', 'Controls': 'Joystick', 'Recline': '—', 'Warranty': '1 year' },
  },
  'compact-air': {
    name: 'Stride Compact Air', tag: 'Ultra-light · Travel', price: 'PKR 320,000', img: 'images/q5.png',
    wa: 'https://wa.me/923001234567?text=Hi%20Stride%2C%20I%27m%20interested%20in%20the%20Stride%20Compact%20Air.%20Please%20share%20availability%20and%20final%20price.',
    specs: { 'Range': '20 km', 'Top speed': '6 km/h', 'Max load': '110 kg', 'Kerb weight': '18 kg', 'Battery': 'Li-ion (airline-safe)', 'Motor': 'Brushless 250W', 'Charge time': '5–7 hrs', 'Foldable': 'Yes (auto-fold option)', 'Suspension': '—', 'Tyres': 'Solid', 'Controls': 'Joystick', 'Recline': '—', 'Warranty': '1 year' },
  },
  'terra-x': {
    name: 'Stride Terra X', tag: 'All-terrain · Dual motor', price: 'PKR 420,000', img: 'images/q5.png',
    wa: 'https://wa.me/923001234567?text=Hi%20Stride%2C%20I%27m%20interested%20in%20the%20Stride%20Terra%20X.%20Please%20share%20availability%20and%20final%20price.',
    specs: { 'Range': '35 km', 'Top speed': '10 km/h', 'Max load': '150 kg', 'Kerb weight': '42 kg', 'Battery': 'Li-ion 24V 30Ah', 'Motor': 'Dual 2×350W', 'Charge time': '8–10 hrs', 'Foldable': 'No', 'Suspension': 'Full', 'Tyres': 'Pneumatic off-road', 'Controls': 'Joystick', 'Recline': '—', 'Warranty': '1 year' },
  },
  'recline-r5': {
    name: 'Stride Recline R5', tag: 'Full recline · High support', price: 'PKR 495,000', img: 'images/q5.png',
    wa: 'https://wa.me/923001234567?text=Hi%20Stride%2C%20I%27d%20like%20to%20pre-order%20the%20Stride%20Recline%20R5.%20Please%20share%20details%20and%20timeline.',
    specs: { 'Range': '30 km', 'Top speed': '8 km/h', 'Max load': '135 kg', 'Kerb weight': '38 kg', 'Battery': 'Li-ion 24V 25Ah', 'Motor': 'Brushless 350W', 'Charge time': '8–10 hrs', 'Foldable': '—', 'Suspension': '—', 'Tyres': '—', 'Controls': 'Joystick + recline', 'Recline': 'Full + elevating legrests', 'Warranty': '1 year' },
  },
};
/* COMPARE_MODELS:END */

const SPEC_ORDER = ['Range', 'Top speed', 'Max load', 'Kerb weight', 'Battery', 'Motor', 'Charge time', 'Foldable', 'Suspension', 'Tyres', 'Controls', 'Recline', 'Warranty'];
// Which rows have a "best" value, and whether higher or lower wins.
const SPEC_BEST = { 'Price': 'low', 'Range': 'high', 'Top speed': 'high', 'Max load': 'high', 'Kerb weight': 'low' };
const COMPARE_MAX = 4;

const compareState = new Set();

const compareToggles = document.querySelectorAll('[data-compare]');
const tray = document.getElementById('compareTray');
const trayItems = document.getElementById('compareTrayItems');
const compareOpenBtn = document.getElementById('compareOpen');
const compareCountEl = document.getElementById('compareCount');
const compareClearBtn = document.getElementById('compareClear');
const compareModal = document.getElementById('compareModal');
const compareModalBody = document.getElementById('compareModalBody');

function parseNum(str) {
  const m = String(str).match(/-?\d+(\.\d+)?/);
  return m ? parseFloat(m[0]) : null;
}

// Returns the set of model ids that hold the best value for a given spec row.
function bestIdsFor(label, ids) {
  const dir = SPEC_BEST[label];
  if (!dir) return new Set();
  const vals = ids.map((id) => {
    const raw = label === 'Price' ? COMPARE_MODELS[id].price : COMPARE_MODELS[id].specs[label];
    return { id, n: parseNum(raw) };
  }).filter((v) => v.n !== null);
  if (vals.length < 2) return new Set();
  const target = dir === 'high' ? Math.max(...vals.map((v) => v.n)) : Math.min(...vals.map((v) => v.n));
  const winners = vals.filter((v) => v.n === target);
  // Only mark a winner when it's not a tie across every model.
  if (winners.length === vals.length) return new Set();
  return new Set(winners.map((v) => v.id));
}

function updateToggleStates() {
  const atMax = compareState.size >= COMPARE_MAX;
  compareToggles.forEach((btn) => {
    const selected = compareState.has(btn.dataset.compare);
    btn.classList.toggle('is-selected', selected);
    btn.setAttribute('aria-pressed', String(selected));
    btn.disabled = atMax && !selected;
    const label = btn.querySelector('.compare-toggle-label');
    if (label) label.textContent = selected ? 'Comparing' : 'Compare';
  });
}

function renderTray() {
  const ids = [...compareState];
  trayItems.innerHTML = '';
  ids.forEach((id) => {
    const m = COMPARE_MODELS[id];
    const chip = document.createElement('span');
    chip.className = 'compare-chip';
    chip.innerHTML = `<span>${m.name}</span><button type="button" aria-label="Remove ${m.name}">&times;</button>`;
    chip.querySelector('button').addEventListener('click', () => toggleModel(id, false));
    trayItems.appendChild(chip);
  });
  if (ids.length < 2) {
    const hint = document.createElement('span');
    hint.className = 'compare-hint';
    hint.textContent = ids.length === 0 ? 'Select models to compare' : 'Add at least one more to compare';
    trayItems.appendChild(hint);
  }
  compareCountEl.textContent = String(ids.length);
  compareOpenBtn.disabled = ids.length < 2;
  tray.classList.toggle('open', ids.length > 0);
  document.body.classList.toggle('compare-active', ids.length > 0);
  tray.setAttribute('aria-hidden', ids.length > 0 ? 'false' : 'true');
}

function toggleModel(id, force) {
  const willSelect = force !== undefined ? force : !compareState.has(id);
  if (willSelect) {
    if (compareState.size >= COMPARE_MAX) return;
    compareState.add(id);
  } else {
    compareState.delete(id);
  }
  updateToggleStates();
  renderTray();
}

compareToggles.forEach((btn) => {
  btn.addEventListener('click', () => toggleModel(btn.dataset.compare));
});

if (compareClearBtn) {
  compareClearBtn.addEventListener('click', () => {
    compareState.clear();
    updateToggleStates();
    renderTray();
  });
}

// Builds the inner <thead>/<tbody> markup for a comparison table over the given model ids.
function comparisonTableHTML(ids) {
  const rows = ['Price', ...SPEC_ORDER];
  const val = (id, label) => (label === 'Price' ? COMPARE_MODELS[id].price : COMPARE_MODELS[id].specs[label]);

  let head = '<thead><tr><th class="compare-corner" scope="col"><span class="sr-only">Specification</span></th>';
  ids.forEach((id) => {
    const m = COMPARE_MODELS[id];
    head += `<th scope="col" class="compare-model-head">
      <img class="compare-model-img" src="${m.img}" alt="${m.name}">
      <span class="compare-model-name">${m.name}</span>
      <span class="compare-model-tag">${m.tag}</span>
    </th>`;
  });
  head += '</tr></thead>';

  let body = '<tbody>';
  rows.forEach((label) => {
    const best = bestIdsFor(label, ids);
    body += `<tr><th scope="row" class="compare-row-label">${label}</th>`;
    ids.forEach((id) => {
      const isBest = best.has(id);
      body += `<td class="compare-cell${isBest ? ' is-best' : ''}${label === 'Price' ? ' compare-price' : ''}">${val(id, label)}</td>`;
    });
    body += '</tr>';
  });
  body += '</tbody>';
  return head + body;
}

function buildCompareTable() {
  compareModalBody.innerHTML = `<div class="compare-table-scroll"><table class="compare-table">${comparisonTableHTML([...compareState])}</table></div>`;
}

// Full always-on comparison table (below "Why Stride")
const compareAllTable = document.getElementById('compareAllTable');
if (compareAllTable) {
  compareAllTable.innerHTML = comparisonTableHTML(Object.keys(COMPARE_MODELS));
}

let lastFocused = null;
function openCompareModal() {
  if (compareState.size < 2) return;
  buildCompareTable();
  lastFocused = document.activeElement;
  compareModal.classList.add('open');
  compareModal.setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');
  const closeBtn = compareModal.querySelector('#compareModalClose');
  if (closeBtn) closeBtn.focus();
}

function closeCompareModal() {
  compareModal.classList.remove('open');
  compareModal.setAttribute('aria-hidden', 'true');
  document.body.classList.remove('modal-open');
  if (lastFocused && typeof lastFocused.focus === 'function') lastFocused.focus();
}

if (compareOpenBtn) compareOpenBtn.addEventListener('click', openCompareModal);
if (compareModal) {
  compareModal.querySelectorAll('[data-compare-close]').forEach((el) =>
    el.addEventListener('click', closeCompareModal)
  );
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && compareModal.classList.contains('open')) closeCompareModal();
  });
}

// Initialise
if (compareToggles.length) {
  updateToggleStates();
  renderTray();
}
