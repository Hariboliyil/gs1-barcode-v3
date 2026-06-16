/* global XLSX, bwipjs, html2canvas, QRCode */

/* ── State ── */
const state = {
  workbookName: '',
  master: [],
  stock: [],
  index: [],
  currentProduct: null,
  form: {
    productCode: '',
    batch: '',
    expiry: '',
    genericCode: '',
    description: '',
    gtin: '',
    locator: '',
    timestamp: '',
  },
};

/* ── DOM refs ── */
const el = {
  sourceStatus:    document.getElementById('sourceStatus'),
  recordCount:     document.getElementById('recordCount'),
  workbookFile:    document.getElementById('workbookFile'),
  uploadLabel:     document.getElementById('uploadLabel'),
  productCode:     document.getElementById('productCode'),
  batch:           document.getElementById('batch'),
  expiry:          document.getElementById('expiry'),
  applyBtn:        document.getElementById('applyBtn'),
  copyBtn:         document.getElementById('copyBtn'),
  downloadBtn:     document.getElementById('downloadBtn'),
  printBtn:        document.getElementById('printBtn'),
  barcodeCanvas:   document.getElementById('barcodeCanvas'),
  barcodeTextView: document.getElementById('barcodeTextView'),
  renderStatus:    document.getElementById('renderStatus'),
  genericCodeView: document.getElementById('genericCodeView'),
  productCodeView: document.getElementById('productCodeView'),
  descriptionView: document.getElementById('descriptionView'),
  batchView:       document.getElementById('batchView'),
  expiryView:      document.getElementById('expiryView'),
  gtinView:        document.getElementById('gtinView'),
  locatorView:     document.getElementById('locatorView'),
  locatorQr:       document.getElementById('locatorQr'),
  timestampView:   document.getElementById('timestampView'),
  expiryWarning:   document.getElementById('expiryWarning'),
  labelPreview:    document.getElementById('labelPreview'),
};

/* ── Helpers ── */
function text(v)  { return String(v ?? '').trim(); }
function norm(v)  { return text(v).toUpperCase(); }

function optionalText(v) {
  const c = text(v);
  if (!c || /^no reference$/i.test(c) || /^n\/a$/i.test(c)) return '';
  return c;
}

function excelSerialToDate(n) {
  const d = Number(n);
  if (!Number.isFinite(d) || d <= 0) return null;
  return new Date(Date.UTC(1899, 11, 30) + d * 86400000);
}

function parseDate(v) {
  if (!v) return '';
  if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10);
  if (typeof v === 'number') {
    const s = String(Math.trunc(v));
    if (s.length === 8) {
      return new Date(Date.UTC(+s.slice(4), +s.slice(2,4)-1, +s.slice(0,2))).toISOString().slice(0,10);
    }
    const d = excelSerialToDate(v);
    return d ? d.toISOString().slice(0,10) : '';
  }
  const s = text(v);
  if (!s) return '';
  if (/^\d{8}$/.test(s)) {
    return new Date(Date.UTC(+s.slice(4), +s.slice(2,4)-1, +s.slice(0,2))).toISOString().slice(0,10);
  }
  const d = new Date(s);
  return isNaN(d) ? '' : d.toISOString().slice(0,10);
}

function displayDate(v) {
  const iso = parseDate(v);
  if (!iso) return '';
  return new Intl.DateTimeFormat('en-GB', {
    day:'2-digit', month:'short', year:'numeric', timeZone:'UTC'
  }).format(new Date(iso + 'T00:00:00Z'));
}

function gs1Date(v) {
  const iso = parseDate(v);
  return iso ? iso.slice(2,4) + iso.slice(5,7) + iso.slice(8,10) : '';
}

function timestampNow() {
  const d = new Date();
  const p = n => String(n).padStart(2,'0');
  return `${p(d.getDate())}${p(d.getMonth()+1)}${d.getFullYear()}${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}

function setStatus(msg) { el.renderStatus.textContent = msg; }

/* ── Expiry warning ── */
function renderExpiryWarning() {
  const iso = parseDate(state.form.expiry);
  let msg = '';
  if (iso) {
    const exp   = new Date(iso + 'T00:00:00');
    const today = new Date(); today.setHours(0,0,0,0);
    const soon  = new Date(today); soon.setMonth(soon.getMonth()+3);
    if (exp < today) msg = 'Warning: expiry date is already in the past.';
    else if (exp < soon) msg = 'Warning: expiry date is less than 3 months away.';
  }
  el.expiryWarning.textContent = msg;
  return msg;
}

/* ── Workbook ── */
function readSheet(wb, name) {
  const sheet = wb.Sheets[name];
  return sheet ? XLSX.utils.sheet_to_json(sheet, { defval:'', raw:false, cellDates:true }) : [];
}

function buildIndex() {
  state.index = state.master.map(row => {
    const productCode = text(row.PRODUCT || row['Product Code'] || row.Product);
    const shortDesc   = text(row['ITEM SHORT DESCRIPTION'] || row['Item Short Description']);
    const partNo      = text(row['Manufacturer Part Number'] || row.Y);
    const generic     = text(row['GENERIC CODE'] || row['Generic Code'] || row.J);
    const gtin        = text(row['GTIN Number'] || row.AB);
    return {
      productCode,
      row,
      searchText: norm([productCode, shortDesc, partNo, generic, gtin].filter(Boolean).join(' ')),
    };
  });
}

function findProduct(code) {
  const key = norm(code);
  return state.index.find(e => norm(e.productCode) === key) || null;
}

function findStock(code) {
  const key = norm(code);
  return state.stock.filter(r => norm(r['Product Code']) === key)[0] || {};
}

function deriveFields(row, productCode) {
  const stock    = findStock(productCode);
  const partNo   = text(row['Manufacturer Part Number'] || row.Y);
  const shortDesc= text(row['ITEM SHORT DESCRIPTION'] || row['Item Short Description']);
  const gtin     = text(row['GTIN Number'] || row.AB);
  return {
    productCode:  text(productCode),
    genericCode:  text(row['GENERIC CODE'] || row['Generic Code'] || row.J),
    description:  [partNo, shortDesc].filter(Boolean).join('-') || shortDesc,
    gtin,
    locator:      optionalText(stock['Bin Wise'] || row['Bin Wise']),
  };
}

function loadWorkbook(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx','xlsm','xls'].includes(ext)) {
    alert(`Unsupported file type ".${ext}". Please upload an .xlsx, .xlsm or .xls file.`);
    return;
  }
  el.sourceStatus.textContent = `Loading "${file.name}"…`;
  el.recordCount.textContent = 'Reading…';
  setStatus('Reading workbook…');

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = e.target.result;
      if (!data || data.byteLength === 0) throw new Error('File is empty.');
      const wb = XLSX.read(new Uint8Array(data), { type:'array', cellDates:true });
      if (!wb.SheetNames.includes('Master'))
        throw new Error(`Sheet "Master" not found. Found: ${wb.SheetNames.join(', ')}`);
      state.workbookName = file.name;
      state.master = readSheet(wb, 'Master');
      state.stock  = readSheet(wb, 'Tbl_StockList');
      buildIndex();
      el.sourceStatus.textContent = file.name;
      el.recordCount.textContent  = `${state.master.length.toLocaleString()} products`;
      clearForm();
      setStatus(`Loaded — ${state.master.length.toLocaleString()} products`);
    } catch (err) {
      console.error(err);
      el.sourceStatus.textContent = 'Load failed';
      el.recordCount.textContent  = 'No data';
      setStatus('Error: ' + (err.message || err));
      alert('Could not read workbook:\n\n' + (err.message || err));
    }
  };
  reader.onerror = () => {
    const msg = reader.error ? reader.error.message : 'Unknown error';
    el.sourceStatus.textContent = 'File read failed';
    el.recordCount.textContent  = 'No data';
    setStatus('File read error: ' + msg);
    alert('Could not read file:\n\n' + msg);
  };
  reader.readAsArrayBuffer(file);
}

/* ── Form ── */
function updateViews() {
  el.productCode.value = state.form.productCode;
  el.batch.value       = state.form.batch;
  el.expiry.value      = state.form.expiry;

  el.genericCodeView.textContent = state.form.genericCode;
  el.productCodeView.textContent = state.form.productCode;
  el.descriptionView.textContent = state.form.description;
  el.batchView.textContent       = state.form.batch;
  el.expiryView.textContent      = displayDate(state.form.expiry);
  el.gtinView.textContent        = state.form.gtin;
  el.locatorView.textContent     = state.form.locator;
  el.timestampView.textContent   = state.form.timestamp;
}

function fitDescription() {
  const node = el.descriptionView;
  const cs   = window.getComputedStyle(node);
  const max  = parseFloat(cs.getPropertyValue('--desc-max-size')) || 26;
  const min  = parseFloat(cs.getPropertyValue('--desc-min-size')) || 11;
  let size   = max;
  node.style.cssText = `font-size:${size}px;white-space:normal;word-break:break-word;overflow:hidden;`;
  while (size > min && (node.scrollHeight > node.clientHeight+1 || node.scrollWidth > node.clientWidth+1)) {
    size -= 0.5;
    node.style.fontSize = size + 'px';
  }
}

function clearForm() {
  state.currentProduct = null;
  state.form = { productCode:'', batch:'', expiry:'', genericCode:'', description:'', gtin:'', locator:'', timestamp:'' };
  updateViews();
  renderQr();
  renderBarcode();
}

function applyLookup(code) {
  if (!state.index.length) return;
  const rec = findProduct(code);
  if (!rec) {
    state.form = { ...state.form, productCode:text(code), genericCode:'', description:'', gtin:'', locator:'' };
    updateViews(); fitDescription(); renderQr(); renderBarcode();
    setStatus('Product not found in workbook');
    return;
  }
  const f = deriveFields(rec.row, rec.productCode);
  state.form = { ...state.form, ...f };
  updateViews(); fitDescription(); renderQr(); renderBarcode();
  setStatus('Loaded ' + f.productCode);
}

/* ── QR ── */
let _qr = null;
function renderQr() {
  const loc = text(state.form.locator);
  el.locatorQr.innerHTML = '';
  _qr = null;
  if (!loc) return;
  try {
    _qr = new QRCode(el.locatorQr, {
      text: loc, width: 104, height: 104,
      colorDark:'#000', colorLight:'#fff',
      correctLevel: QRCode.CorrectLevel.M,
    });
  } catch(e) { console.warn('QR failed', e); }
}

/* ── Barcode ── */
function buildPayload() {
  const gtin   = text(state.form.gtin);
  const batch  = text(state.form.batch);
  const expiry = gs1Date(state.form.expiry);
  if (!gtin || !batch || !expiry) return '';
  return `01${gtin}17${expiry}10${batch}`;
}

function renderBarcode() {
  const payload = buildPayload();
  el.barcodeTextView.textContent = payload || 'Enter product code, batch number, and expiry date';

  const ctx = el.barcodeCanvas.getContext('2d');
  ctx.clearRect(0, 0, el.barcodeCanvas.width, el.barcodeCanvas.height);

  if (!payload) {
    ctx.fillStyle = '#6b7280';
    ctx.font = '600 18px "Space Grotesk", sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Awaiting input', el.barcodeCanvas.width/2, el.barcodeCanvas.height/2);
    setStatus('Waiting for inputs');
    return;
  }
  try {
    bwipjs.toCanvas(el.barcodeCanvas, {
      bcid:'datamatrix', text:payload, scale:6, padding:2, backgroundcolor:'FFFFFF'
    });
    setStatus('Rendered');
  } catch(err) {
    ctx.fillStyle = '#dc2626';
    ctx.font = '600 14px "IBM Plex Mono", monospace';
    ctx.textAlign = 'center';
    ctx.fillText(String(err.message||err), el.barcodeCanvas.width/2, el.barcodeCanvas.height/2);
    setStatus('Barcode error: ' + (err.message||err));
  }
}

function renderLabel() {
  state.form.timestamp = timestampNow();
  updateViews();
  fitDescription();
  renderExpiryWarning();
  renderQr();
  renderBarcode();
}

/* ── Copy ── */
function copyBarcodeText() {
  const p = buildPayload();
  if (!p) return;
  navigator.clipboard?.writeText(p)
    .then(() => setStatus('Copied'))
    .catch(() => setStatus('Clipboard unavailable'));
}

/* ── Download ── */
async function downloadLabel() {
  const c = await html2canvas(el.labelPreview, { backgroundColor:'#ffffff', scale:2, useCORS:true });
  const a = document.createElement('a');
  a.download = `label-${state.form.productCode || 'out'}-${Date.now()}.png`;
  a.href = c.toDataURL('image/png');
  a.click();
}

/* ── Print ── */
async function printLabel() {
  setStatus('Preparing print…');
  try {
    // Render label to high-res image
    const c = await html2canvas(el.labelPreview, {
      backgroundColor: '#ffffff',
      scale: 4,
      useCORS: true,
      logging: false,
    });
    const dataUrl = c.toDataURL('image/png');

    // Remove previous print area if any
    const prev = document.getElementById('printArea');
    if (prev) prev.remove();

    // Build a div with the image, hidden on screen, visible only in print CSS
    const img = document.createElement('img');
    img.src = dataUrl;

    const div = document.createElement('div');
    div.id = 'printArea';
    div.appendChild(img);
    document.body.appendChild(div);

    // Wait for image to fully load then print
    img.onload = () => {
      window.print();
      setTimeout(() => { div.remove(); setStatus('Rendered'); }, 2000);
    };
    img.onerror = () => { div.remove(); setStatus('Print image failed'); };

  } catch(err) {
    console.error(err);
    setStatus('Print failed: ' + (err.message||err));
  }
}

/* ── Events ── */
function wireEvents() {
  // File upload
  el.workbookFile.addEventListener('change', () => {
    const f = el.workbookFile.files && el.workbookFile.files[0];
    if (f) loadWorkbook(f);
    el.workbookFile.value = '';
  });

  // Drag and drop
  el.uploadLabel.addEventListener('dragover', e => {
    e.preventDefault();
    el.uploadLabel.style.background = 'rgba(45,212,191,0.2)';
  });
  el.uploadLabel.addEventListener('dragleave', () => {
    el.uploadLabel.style.background = '';
  });
  el.uploadLabel.addEventListener('drop', e => {
    e.preventDefault();
    el.uploadLabel.style.background = '';
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) loadWorkbook(f);
  });

  // Inputs
  el.productCode.addEventListener('input', () => {
    state.form.productCode = el.productCode.value.trim();
    state.index.length ? applyLookup(state.form.productCode) : renderLabel();
  });

  el.batch.addEventListener('input', () => {
    state.form.batch = el.batch.value.trim();
    renderLabel();
  });

  const onExpiry = () => {
    state.form.expiry = el.expiry.value;
    const w = renderExpiryWarning();
    if (w) setStatus(w);
    renderLabel();
  };
  el.expiry.addEventListener('input', onExpiry);
  el.expiry.addEventListener('change', onExpiry);

  // Buttons
  el.applyBtn.addEventListener('click', () => {
    if (state.form.productCode && state.index.length) applyLookup(state.form.productCode);
    else renderLabel();
  });
  el.copyBtn.addEventListener('click', copyBarcodeText);
  el.downloadBtn.addEventListener('click', () => void downloadLabel());
  el.printBtn.addEventListener('click', () => void printLabel());
}

/* ── Init ── */
function init() {
  el.sourceStatus.textContent = 'No workbook loaded';
  el.recordCount.textContent  = 'No data';
  clearForm();
  wireEvents();
  renderExpiryWarning();
}

init();
