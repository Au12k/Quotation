// ── CUSTOMER DB ───────────────────────────────────────────────
const CK = 'tkt_v2_customers';
const loadC = () => { try { return JSON.parse(localStorage.getItem(CK)) || []; } catch { return []; } };
const saveC = a => localStorage.setItem(CK, JSON.stringify(a));

function saveCust() {
  const th = g('cust-th').trim();
  if (!th) return alert('กรุณากรอกชื่อบริษัทลูกค้า');
  const c = { id: Date.now(), th, en: g('cust-en'), addr: g('cust-addr'), tax: g('cust-tax'), tel: g('cust-tel'), contact: g('cust-contact'), deliver: g('cust-deliver') };
  const a = loadC(); const i = a.findIndex(x => x.th === th);
  if (i >= 0) a[i] = { ...a[i], ...c }; else a.push(c);
  saveC(a);
  const t = document.getElementById('saved-toast'); t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2500);
}
function clearCust() {
  ['cust-th','cust-en','cust-addr','cust-tax','cust-tel','cust-contact','cust-deliver'].forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
  closeCD(); u();
}
function fillCust(json) {
  const c = JSON.parse(json);
  sv('cust-th', c.th || ''); sv('cust-en', c.en || ''); sv('cust-addr', c.addr || '');
  sv('cust-tax', c.tax || ''); sv('cust-tel', c.tel || '');
  sv('cust-contact', c.contact || ''); sv('cust-deliver', c.deliver || '');
  closeCD(); u();
}
function delCust(id, e) {
  e.stopPropagation();
  if (!confirm('ลบลูกค้านี้?')) return;
  saveC(loadC().filter(c => c.id !== id));
  onCS(document.getElementById('cust-th').value);
}
function onCS(v) {
  const a = loadC(), q = v.trim().toLowerCase();
  const f = q ? a.filter(c => c.th.toLowerCase().includes(q) || (c.en||'').toLowerCase().includes(q)) : a;
  const dd = document.getElementById('cust-dd');
  if (!f.length) { dd.classList.remove('open'); return; }
  dd.innerHTML = f.map(c => {
    const j = encodeURIComponent(JSON.stringify(c));
    return `<div class="cust-opt" onmousedown="fillCust(decodeURIComponent('${j}'))">
      <button class="cust-del-btn" onmousedown="delCust(${c.id},event)">ลบ</button>
      <div class="cust-opt-name">${c.th}</div>
      ${c.en ? `<div class="cust-opt-sub">${c.en}</div>` : ''}
      ${c.contact ? `<div class="cust-opt-sub">📞 ${c.contact}</div>` : ''}
    </div>`;
  }).join('');
  dd.classList.add('open');
}
function closeCD() { document.getElementById('cust-dd').classList.remove('open'); }

// ── HELPERS ────────────────────────────────────────────────────
const $ = id => document.getElementById(id);
// ✅ FIX: g() และ gv() รวมเป็นฟังก์ชันเดียวกัน ลด confusion
const g  = id => $(id)?.value || '';
const gv = id => $(id)?.value || '';
const sv = (id, v) => { const el = $(id); if (el) el.value = v; };
const fmt = n => n.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const setText = (id, v) => { const el = $(id); if (el) el.textContent = v; };

// ── FABRIC DB ──────────────────────────────────────────────────
// color labels with special suffixes
const COLOR_LABELS = {
  "LBN161":  { "19": "19 (Waterproof)" },
  "FP26275": { "19": "19 (Waterproof)" },
  "FP26540": { "19": "19 (Waterproof)" },
};
// รหัสที่เป็น PVC / Leather / EPU ไม่ต้องใส่ waterproof
function isLeather(code) {
  if (!code || !FABRICS[code]) return false;
  const comp = (FABRICS[code].comp || '').toLowerCase();
  return comp.includes('pvc') || comp.includes('leather') || comp.includes('epu') || comp.includes('pu ') || comp.includes('หนัง');
}
function colorLabel(code, c) {
  // ถ้ามี label เฉพาะ ให้ใช้ label นั้น
  if (COLOR_LABELS[code] && COLOR_LABELS[code][c]) return COLOR_LABELS[code][c];
  // ถ้าไม่ใช่หนัง ให้ใส่ (waterproof) อัตโนมัติ
  if (!isLeather(code)) return c + ' (waterproof)';
  return c;
}

const FABRICS = {
  "FP25066": { colors:["1","2","3","4","5","8","16","23","25","26","28","29","32"],         cut:119, roll:110, width:140, weight:330, comp:"100% Polyester" },
  "FP25058": { colors:["4","5","6","7","8","9","11","13","16","17","22","24","26","30"],     cut:114, roll:105, width:140, weight:350, comp:"100% Polyester" },
  "FP31645": { colors:["0","1","2","5","10","11","12","14"],                                 cut:117, roll:108, width:142, weight:300, comp:"100% Polyester" },
  "FP31660": { colors:["2A1","2A4","2A5","2A6","2A26"],                                     cut:149, roll:138, width:142, weight:350, comp:"100% Polyester" },
  "FP44879": { colors:["1","6","8","20","21"],                                               cut:128, roll:118, width:150, weight:320, comp:"100% Polyester" },
  "FP26540": { colors:["3","4","19"],                                                        cut:167, roll:155, width:145, weight:340, comp:"100% Polyester" },
  "FP32097": { colors:["1","5","15","18","21","22","25"],                                    cut:181, roll:168, width:145, weight:420, comp:"94% Polyester 6% Rayon" },
  "FP26275": { colors:["1","2","3","5","16","17","18","19"],                                 cut:151, roll:140, width:145, weight:365, comp:"96% Polyester 4% Nylon" },
  "LBN161":  { colors:["1","14","15","19"],                                                  cut:158, roll:146, width:145, weight:370, comp:"100% Polyester" },
  "FP04055": { colors:["1","4","11","13"],                                                   cut:208, roll:196, width:145, weight:370, comp:"100% Polyester" },
  "LBN275":  { colors:["1","3","4","17","18","22","27","31","32"],                           cut:151, roll:140, width:145, weight:370, comp:"100% Polyester" },
  "LBN2041": { colors:["1","2","4","23","24","25","36","38","42"],                           cut:161, roll:149, width:145, weight:370, comp:"96% Polyester 4% Nylon" },
  "FP21641": { colors:["2","26","86","104","107"],                                           cut:110, roll:102, width:140, weight:360, comp:"100% PVC" },
  "FP24815": { colors:["2","3","6","9","35","42","53","57","72","88","104","107"],           cut:116, roll:108, width:137, weight:350, comp:"100% PVC" },
  "FP44660": { colors:["1","2","3","4","5","6","7","8","9","10"],                            cut:115, roll:107, width:140, weight:360, comp:"100% PVC" },
  "FP32285": { colors:["A18","A30","B04","B07","B09","B12","B15"],                           cut:269, roll:250, width:140, weight:580, comp:"15% EPU 85% Leather" },
  "FP32286": { colors:["01","02","03","05","06","13","15","17","22"],                        cut:269, roll:250, width:140, weight:580, comp:"15% EPU 85% Leather" },
};

// ✅ FIX: โหลด FABRICS จาก localStorage หลังจาก declare แล้ว (แก้ ReferenceError)
(function() {
  try {
    const savedFabric = localStorage.getItem('fabricDB');
    if (savedFabric) Object.assign(FABRICS, JSON.parse(savedFabric));
  } catch(e) { console.warn('fabricDB load error:', e); }
})();

let IC = 0;
document.getElementById('qt-date').valueAsDate = new Date();
document.getElementById('t-delivery').value = 'จัดส่งถึงที่อยู่ในเอกสาร / Door to Door';
u();

// -- LOGO STORAGE (IndexedDB + localStorage fallback) --
const LogoStore = (() => {
  const DB_NAME = 'tkt_db', STORE = 'settings', KEY = 'logo';
  let _db = null;
  function openDB() {
    return new Promise((resolve, reject) => {
      if (_db) return resolve(_db);
      const req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = e => e.target.result.createObjectStore(STORE);
      req.onsuccess = e => { _db = e.target.result; resolve(_db); };
      req.onerror = () => reject(req.error);
    });
  }
  async function save(src) {
    try {
      const db = await openDB();
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).put(src, KEY);
      await new Promise((res, rej) => { tx.oncomplete = res; tx.onerror = rej; });
      return true;
    } catch(e) {}
    try { localStorage.setItem('tkt_logo', src); return true; } catch(e) {}
    return false;
  }
  async function load() {
    try {
      const db = await openDB();
      return await new Promise((resolve, reject) => {
        const tx = db.transaction(STORE, 'readonly');
        const req = tx.objectStore(STORE).get(KEY);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => reject(req.error);
      });
    } catch(e) {}
    try { return localStorage.getItem('tkt_logo'); } catch(e) {}
    return null;
  }
  async function remove() {
    try {
      const db = await openDB();
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).delete(KEY);
    } catch(e) {}
    try { localStorage.removeItem('tkt_logo'); } catch(e) {}
  }
  return { save, load, remove };
})();
function applyLogo(src) {
  $('logo-sb').src = src; $('logo-sb').style.display = 'block'; $('logo-ph').style.display = 'none';
  $('doc-logo').src = src; $('doc-logo').style.display = 'block';
}
async function clearLogo() {
  await LogoStore.remove();
  $('logo-sb').src = ''; $('logo-sb').style.display = 'none';
  $('logo-ph').style.display = 'block';
  $('doc-logo').src = ''; $('doc-logo').style.display = 'none';
}
function loadLogo(ev) {
  const f = ev.target.files[0]; if (!f) return;
  const r = new FileReader();
  r.onload = async e => {
    const src = e.target.result;
    await LogoStore.save(src);
    applyLogo(src);
    const t = $('logo-toast');
    if (t) { t.style.display='block'; setTimeout(()=>t.style.display='none', 2500); }
  };
  r.readAsDataURL(f);
}
(async function() {
  try { const saved = await LogoStore.load(); if (saved) applyLogo(saved); } catch(e) {}
})();

function addItem() {
  IC++;
  const id = IC;
  const wrap = $('items-wrap');
  const div = document.createElement('div');
  div.className = 'item-row'; div.id = `item-${id}`;
  div.innerHTML = `
    <div class="item-hd">รายการที่ ${id}</div>
    <button class="btn-rm" onclick="rmItem(${id})">×</button>
    <div class="form-row">
      <div class="form-group" style="margin:0">
        <label>รหัสผ้า</label>
        <div class="fab-combo" id="fab-combo-${id}">
          <input type="text" id="i${id}-code" placeholder="พิมพ์หรือเลือกรหัส..." autocomplete="off"
            oninput="onFabInput(${id})" onfocus="onFabFocus(${id})" onblur="onFabBlur(${id})"
            onkeydown="onFabKey(event,${id})">
          <span class="fab-combo-arrow">▼</span>
          <div class="fab-dd" id="fab-dd-${id}"></div>
        </div>
        <div class="fab-combo-badge" id="fab-badge-${id}"></div>
      </div>
      <div class="form-group" style="margin:0">
        <label>สี / Color</label>
        <div class="fab-combo" id="col-combo-${id}">
          <input type="text" id="i${id}-color" placeholder="เลือกหรือพิมพ์สี..." autocomplete="off"
            oninput="onColInput(${id})" onfocus="onColFocus(${id})" onblur="onColBlur(${id})"
            onkeydown="onColKey(event,${id})">
          <span class="fab-combo-arrow">▼</span>
          <div class="fab-dd" id="col-dd-${id}"></div>
        </div>
      </div>
    </div>
    <div class="form-row" style="margin-top:6px">
      <div class="form-group" style="margin:0">
        <label>ส่วนประกอบ</label>
        <input type="text" id="i${id}-comp" placeholder="เช่น 100% Polyester" oninput="u()">
      </div>
      <div class="form-group" style="margin:0">
        <label>จำนวน (หลา)</label>
        <input type="number" id="i${id}-qty" placeholder="0" oninput="u()">
      </div>
    </div>
    <div class="form-row3" style="margin-top:6px">
      <div class="form-group" style="margin:0">
        <label>ความกว้าง (cm)</label>
        <input type="number" id="i${id}-width" placeholder="145" oninput="u()">
      </div>
      <div class="form-group" style="margin:0">
        <label>น้ำหนัก (gsm)</label>
        <input type="number" id="i${id}-weight" placeholder="370" oninput="u()">
      </div>
      <div class="form-group" style="margin:0">
        <label>จำนวน (ม้วน)</label>
        <input type="number" id="i${id}-rolls" placeholder="0" oninput="u()">
      </div>
    </div>
    <div class="form-group" style="margin-top:6px;margin-bottom:0">
      <label>ราคาขาย (฿/หลา)</label>
      <input type="number" id="i${id}-price" placeholder="0.00" step="0.01" oninput="u()">
    </div>
    <div id="i${id}-cost" class="cost-badge" style="display:none">ราคาทุน: <strong id="i${id}-cv">—</strong></div>
    <div class="form-group" style="margin-top:6px;margin-bottom:0">
      <label>หมายเหตุรายการนี้</label>
      <input type="text" id="i${id}-rem" placeholder="(Roll / หมายเหตุ)" oninput="u()">
    </div>
  `;
  wrap.appendChild(div); u();
}

function rmItem(id) { $(`item-${id}`)?.remove(); u(); }

// ── COLOR COMBO BOX ────────────────────────────────────────────
function onColInput(id) {
  const val = $(`i${id}-color`).value;
  renderColDD(id, val);
  u();
}
function renderColDD(id, q) {
  const dd = $(`col-dd-${id}`);
  const code = $(`i${id}-code`)?.value?.trim().toUpperCase() || '';
  const colors = (code && FABRICS[code]) ? FABRICS[code].colors : [];
  const filtered = q ? colors.filter(c => c.toLowerCase().includes(q.toLowerCase())) : colors;
  if (filtered.length === 0 && colors.length > 0) {
    dd.innerHTML = `<div class="fab-dd-empty">ไม่พบสีนี้ในระบบ — พิมพ์ต่อเพื่อใช้เป็นสีใหม่</div>`;
  } else if (filtered.length === 0) {
    dd.innerHTML = `<div class="fab-dd-empty">พิมพ์เลขสีได้เลย</div>`;
  } else {
    dd.innerHTML = filtered.map(c => {
      const label = colorLabel(code, c);
      return `<div class="fab-dd-item" onmousedown="selectCol(${id},'${c}')">${label}</div>`;
    }).join('');
  }
  dd.classList.add('open');
}
function selectCol(id, color) {
  const code = $(`i${id}-code`)?.value?.trim().toUpperCase() || '';
  $(`i${id}-color`).value = colorLabel(code, color);
  $(`col-dd-${id}`).classList.remove('open');
  u();
}
function onColFocus(id) {
  renderColDD(id, $(`i${id}-color`).value);
}
function onColBlur(id) {
  setTimeout(() => { const dd = $(`col-dd-${id}`); if(dd) dd.classList.remove('open'); }, 150);
}
function onColKey(e, id) {
  const dd = $(`col-dd-${id}`);
  if (!dd.classList.contains('open')) return;
  const items = dd.querySelectorAll('.fab-dd-item');
  let cur = dd.querySelector('.fab-dd-item.active');
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    const next = cur ? cur.nextElementSibling : items[0];
    if (cur) cur.classList.remove('active');
    if (next && next.classList.contains('fab-dd-item')) next.classList.add('active');
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    const prev = cur ? cur.previousElementSibling : items[items.length-1];
    if (cur) cur.classList.remove('active');
    if (prev && prev.classList.contains('fab-dd-item')) prev.classList.add('active');
  } else if (e.key === 'Enter') {
    e.preventDefault();
    if (cur) { selectCol(id, cur.textContent.split(' ')[0]); }
    else { dd.classList.remove('open'); u(); }
  } else if (e.key === 'Escape') {
    dd.classList.remove('open');
  }
}

// ── FABRIC COMBO BOX ───────────────────────────────────────────
function onFabInput(id) {
  const val = $(`i${id}-code`).value.trim().toUpperCase();
  $(`i${id}-code`).value = $(`i${id}-code`).value.toUpperCase();
  renderFabDD(id, val);
  // ถ้าพิมพ์ตรงกับรหัสในระบบ ให้ auto-fill ทันที
  if (val && FABRICS[val]) {
    onFC(id);
  } else {
    // custom code — ล้าง auto-fill แต่ยังกรอกเองได้
    const sel = $(`i${id}-color`);
    sel.innerHTML = '<option value="">--</option>';
    const badge = $(`fab-badge-${id}`);
    if (val && !FABRICS[val]) {
      badge.textContent = '✏️ รหัสพรีออร์เดอร์ — กรอกข้อมูลเองได้เลย';
      badge.className = 'fab-combo-badge show';
      $(`i${id}-code`).classList.add('custom-code');
    } else {
      badge.className = 'fab-combo-badge';
      $(`i${id}-code`).classList.remove('custom-code');
    }
    const cb = $(`i${id}-cost`); if(cb) cb.style.display = 'none';
  }
  u();
}
function renderFabDD(id, q) {
  const dd = $(`fab-dd-${id}`);
  const keys = Object.keys(FABRICS);
  const filtered = q ? keys.filter(k => k.includes(q)) : keys;
  if (filtered.length === 0) {
    dd.innerHTML = `<div class="fab-dd-empty">ไม่พบรหัสนี้ — พิมพ์ต่อเพื่อใช้เป็นรหัสใหม่</div>`;
  } else {
    dd.innerHTML = filtered.map(k => {
      const f = FABRICS[k];
      return `<div class="fab-dd-item" onmousedown="selectFab(${id},'${k}')">${k}<span>${f.comp || ''} | ${f.width || ''}cm / ${f.weight || ''}gsm</span></div>`;
    }).join('');
  }
  dd.classList.add('open');
}
function selectFab(id, code) {
  $(`i${id}-code`).value = code;
  $(`fab-dd-${id}`).classList.remove('open');
  $(`i${id}-code`).classList.remove('custom-code');
  $(`fab-badge-${id}`).className = 'fab-combo-badge';
  onFC(id);
}
function onFabFocus(id) {
  const val = $(`i${id}-code`).value.trim().toUpperCase();
  renderFabDD(id, val);
}
function onFabBlur(id) {
  setTimeout(() => { const dd = $(`fab-dd-${id}`); if(dd) dd.classList.remove('open'); }, 150);
}
function onFabKey(e, id) {
  const dd = $(`fab-dd-${id}`);
  if (!dd.classList.contains('open')) return;
  const items = dd.querySelectorAll('.fab-dd-item');
  let cur = dd.querySelector('.fab-dd-item.active');
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    const next = cur ? cur.nextElementSibling : items[0];
    if (cur) cur.classList.remove('active');
    if (next && next.classList.contains('fab-dd-item')) next.classList.add('active');
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    const prev = cur ? cur.previousElementSibling : items[items.length-1];
    if (cur) cur.classList.remove('active');
    if (prev && prev.classList.contains('fab-dd-item')) prev.classList.add('active');
  } else if (e.key === 'Enter') {
    e.preventDefault();
    if (cur) { const code = cur.textContent.split(' ')[0]; selectFab(id, code); }
    else { dd.classList.remove('open'); onFC(id); }
  } else if (e.key === 'Escape') {
    dd.classList.remove('open');
  }
}

function onFC(id) {
  const code = $(`i${id}-code`).value.trim().toUpperCase();
  // clear color field and dropdown cache
  $(`i${id}-color`).value = '';
  if (code && FABRICS[code]) {
    const fab = FABRICS[code];
    // auto-fill width / weight / comp
    const wEl = $(`i${id}-width`), gEl = $(`i${id}-weight`), cEl = $(`i${id}-comp`);
    if (wEl && fab.width) wEl.value = fab.width;
    if (gEl && fab.weight) gEl.value = fab.weight;
    if (cEl && fab.comp) cEl.value = fab.comp;
    // badge
    const badge = $(`fab-badge-${id}`);
    if (badge) { badge.textContent = `✓ ${fab.comp} | ${fab.width}cm / ${fab.weight}gsm`; badge.className = 'fab-combo-badge show'; }
    onTC(id);
  } else {
    const cb = $(`i${id}-cost`); if(cb) cb.style.display = 'none';
    if (!code) {
      const wEl = $(`i${id}-width`), gEl = $(`i${id}-weight`), cEl = $(`i${id}-comp`);
      if (wEl) wEl.value = ''; if (gEl) gEl.value = ''; if (cEl) cEl.value = '';
      const badge = $(`fab-badge-${id}`);
      if (badge) badge.className = 'fab-combo-badge';
    }
  }
  u();
}
function onTC(id) {
  const code = $(`i${id}-code`)?.value;
  const cb = $(`i${id}-cost`);
  if (code && FABRICS[code]) {
    $(`i${id}-cv`).textContent = `ตัดหลา: ${FABRICS[code].cut} | ม้วน: ${FABRICS[code].roll} ฿/หลา`;
    cb.style.display = 'inline-block';
  } else cb.style.display = 'none';
  u();
}

// ── COLLECT ITEMS ──────────────────────────────────────────────
function getItems() {
  const rows = [];
  for (let i = 1; i <= IC; i++) {
    if (!$(`item-${i}`)) continue;
    const code  = $(`i${i}-code`)?.value || '';
    const color = $(`i${i}-color`)?.value || '';
    const comp  = $(`i${i}-comp`)?.value || '';
    const width = parseFloat($(`i${i}-width`)?.value) || '';
    const weight= parseFloat($(`i${i}-weight`)?.value) || '';
    const price = parseFloat($(`i${i}-price`)?.value) || 0;
    const qty   = parseFloat($(`i${i}-qty`)?.value) || 0;
    const rolls = parseFloat($(`i${i}-rolls`)?.value) || 0;
    const rem   = $(`i${i}-rem`)?.value || '';
    const amt   = price * qty;
    const colorDisplay = color; // color input already contains the display label
    rows.push({ num: rows.length+1, code, color, colorDisplay, comp, width, weight, price, qty, rolls, amt, rem });
  }
  return rows;
}

// ── UPDATE PREVIEW ─────────────────────────────────────────────
function u() {
  const qtno = gv('qt-no') || '—';
  const qtdate = gv('qt-date');
  let ds = '—';
  if (qtdate) { const d = new Date(qtdate); ds = d.toLocaleDateString('th-TH', { year:'numeric', month:'long', day:'numeric' }); }
  setText('d-qtno', qtno); setText('d-qtdate', ds); setText('d-footqt', `QT: ${qtno}`);
  setText('d-sales', gv('sales-name') || '—');
  setText('d-mgr', gv('manager-name') || '—');
  setText('d-contact', gv('cust-contact') || '—');
  setText('d-tel', gv('cust-tel') || '—');
  setText('d-cth', gv('cust-th') || '—');
  setText('d-cen', gv('cust-en'));
  setText('d-caddr', gv('cust-addr'));
  setText('d-ctax', gv('cust-tax') ? `Tax ID: ${gv('cust-tax')}` : '');
  const deliver = gv('cust-deliver');
  setText('d-cdth', gv('cust-th') || '—');
  setText('d-cdeliver', deliver || gv('cust-addr') || '');
  setText('d-sign-sel', gv('sales-name') || 'บริษัท ไทยคิม เท็กไทล์ จำกัด');
  setText('d-sign-buy', gv('cust-th') || 'ลูกค้า / Customer');
  // Terms
  setText('d-tdeliv', gv('t-delivery') || '—'); setText('d-tlead', gv('t-lead') || '—');
  setText('d-tpay', gv('t-payment') || '—'); setText('d-tvalid', gv('t-valid') || '—');
  setText('d-tdep', gv('t-deposit') || '—');
  // Remark
  const rem = gv('remark');
  const rb = $('d-remarkbox');
  if (rem) { setText('d-remark', rem); rb.style.display = 'block'; } else rb.style.display = 'none';

  // Table
  const items = getItems();
  const tbody = $('d-tbody');
  tbody.innerHTML = '';
  let subtotal = 0;
  if (items.length === 0) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;color:#aaa;padding:16px;font-style:italic">กรุณาเพิ่มรายการผ้า</td></tr>';
  } else {
    items.forEach(item => {
      subtotal += item.amt;
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="c">${item.num}</td>
        <td><span class="code">${item.code || '—'}</span></td>
        <td class="c">${item.colorDisplay ? `<span class="ctag">${item.colorDisplay}</span>` : '—'}</td>
        <td class="c">${item.comp || '—'}</td>
        <td class="c">${item.width || '—'}</td>
        <td class="c">${item.weight || '—'}</td>
        <td class="r">${fmt(item.price)}</td>
        <td class="c">${item.qty || '—'}</td>
        <td class="c">${item.rolls || '—'}</td>
        <td class="r">${fmt(item.amt)}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  // Totals
  const discPct = parseFloat(gv('discount')) || 0;
  const delFee  = parseFloat(gv('delivery-fee')) || 0;
  const vatOn   = gv('vat') === 'yes';
  const discAmt = subtotal * discPct / 100;
  const afterDisc = subtotal - discAmt;
  const vatAmt  = vatOn ? afterDisc * 0.07 : 0;  // VAT จากยอดสินค้าเท่านั้น
  const grand   = afterDisc + delFee + vatAmt;

  setText('d-sub', `${fmt(subtotal)} ฿`);
  if (discPct > 0) {
    $('d-discrow').style.display = 'flex'; $('d-afrow').style.display = 'flex';
    setText('d-discpct', discPct); setText('d-discamt', `-${fmt(discAmt)} ฿`);
    setText('d-afamt', `${fmt(afterDisc)} ฿`);
  } else { $('d-discrow').style.display = 'none'; $('d-afrow').style.display = 'none'; }
  setText('d-delamt', `${fmt(delFee)} ฿`);
  if (vatOn) { $('d-vatrow').style.display = 'flex'; setText('d-vatamt', `${fmt(vatAmt)} ฿`); }
  else $('d-vatrow').style.display = 'none';
  setText('d-grand', `${fmt(grand)} ฿`);

  // ✅ FIX: คำนวณ balance จากตัวแปร grand โดยตรง (ไม่อ่านจาก DOM เพื่อกัน race condition)
  (function() {
    const depStr = gv('t-deposit') || '';
    const depMatch = depStr.match(/([\d.]+)\s*%/);
    let balText = '';
    if (depMatch) {
      const depPct = parseFloat(depMatch[1]);
      if (grand > 0) {
        const depAmt = grand * depPct / 100;
        const bal = grand - depAmt;
        balText = bal.toLocaleString('th-TH', {minimumFractionDigits:2, maximumFractionDigits:2}) + ' ฿';
      }
    } else if (depStr && depStr !== 'N/A' && depStr !== '-') {
      const depNum = parseFloat(depStr.replace(/[^\d.]/g,'')) || 0;
      if (grand > 0 && depNum > 0) {
        const bal = grand - depNum;
        balText = bal.toLocaleString('th-TH', {minimumFractionDigits:2, maximumFractionDigits:2}) + ' ฿';
      }
    } else {
      if (grand > 0) balText = fmt(grand) + ' ฿';
    }
    document.getElementById('t-balance').value = balText;
    setText('d-tbal', balText || '—');
  })();

  // ── สลับบัญชีธนาคารตาม VAT และ Sales ──────────────────────
  const sales = (gv('sales-name') || '').trim().toLowerCase();
  const isThip = sales === 'thip';

  const BANK_VAT = {
    name: 'บจก.ไทย คิม เท็กซ์ไทล์',
    bank: 'กสิกรไทย (K-Bank)',
    no:   '129-1-17288-8',
    bg:   '#EEF4FA', border: '#1F3864', labelColor: '#1F3864'
  };
  const BANK_NO_VAT_THIP = {
    name: 'Mr.Zhang Jian',
    bank: 'ไทยพาณิชย์ (SCB)',
    no:   '414-133416-2',
    bg:   '#FFF8EE', border: '#C8963E', labelColor: '#C8963E'
  };
  const BANK_NO_VAT_OTHER = {
    name: 'Liu Hongwen',
    bank: 'ไทยพาณิชย์ (SCB)',
    no:   '245-215274-8',
    bg:   '#FFF8EE', border: '#C8963E', labelColor: '#C8963E'
  };

  let bk;
  if (vatOn) {
    bk = BANK_VAT;
  } else if (isThip) {
    bk = BANK_NO_VAT_THIP;   // Thip + ไม่มี VAT → Zhang Jian
  } else {
    bk = BANK_NO_VAT_OTHER;  // Sales อื่น + ไม่มี VAT → บัญชีปกติ
  }
  setText('d-bank-name', bk.name);
  setText('d-bank-bank', bk.bank);
  setText('d-bank-no',   bk.no);
  const bankBox = $('d-bankbox');
  if (bankBox) {
    bankBox.style.background = bk.bg;
    bankBox.style.borderColor = bk.border;
    bankBox.querySelector('.blbl').style.color = bk.labelColor;
  }
}
// ─── ดึงข้อมูลลูกค้าจาก CRM (localStorage key: "customers") ──
function importCRMCustomers() {
  try {
    const raw = localStorage.getItem("customers");
    if (!raw) {
      alert("❌ ไม่พบข้อมูลลูกค้าจาก CRM\nกรุณาเปิด CRM ก่อนแล้ว Import โฟลเดอร์ลูกค้าให้เรียบร้อย");
      return;
    }

    const crmCustomers = JSON.parse(raw);
    if (!crmCustomers.length) {
      alert("❌ CRM ยังไม่มีข้อมูลลูกค้า");
      return;
    }

    // แปลงรูปแบบ CRM → Quotation
    // CRM: { id, name, tel, orders[], remark }
    // Quotation: { id, th, en, addr, tax, tel, contact, deliver }
    const qtCusts = loadC(); // โหลดที่มีอยู่ใน Quotation แล้ว
    let added = 0, updated = 0;

    crmCustomers.forEach(c => {
      // ใช้ชื่อจาก CRM เป็น th
      const th = c.name || c.id || "";
      if (!th) return;

      const newEntry = {
        id:      c.id || Date.now(),
        th:      th,
        en:      c.en      || "",
        addr:    c.addr    || "",
        tax:     c.tax     || "",
        tel:     c.tel     || "",
        contact: c.contact || "",
        deliver: c.deliver || "",
      };
      // ✅ field ใหม่จาก export Excel
      if (!newEntry.en      && c.customerEn)      newEntry.en      = c.customerEn;
      if (!newEntry.addr    && c.customerAddr)    newEntry.addr    = c.customerAddr;
      if (!newEntry.tax     && c.customerTax)     newEntry.tax     = c.customerTax;
      if (!newEntry.contact && c.customerContact) newEntry.contact = c.customerContact;
      if (!newEntry.deliver && c.customerShip)    newEntry.deliver = c.customerShip;

      const idx = qtCusts.findIndex(x =>
        x.th.toLowerCase() === th.toLowerCase() ||
        x.id === c.id
      );

      if (idx >= 0) {
        // อัปเดตเฉพาะ field ที่ยังว่างใน Quotation
        if (!qtCusts[idx].tel     && newEntry.tel)     qtCusts[idx].tel     = newEntry.tel;
        if (!qtCusts[idx].en      && newEntry.en)      qtCusts[idx].en      = newEntry.en;
        if (!qtCusts[idx].addr    && newEntry.addr)    qtCusts[idx].addr    = newEntry.addr;
        if (!qtCusts[idx].contact && newEntry.contact) qtCusts[idx].contact = newEntry.contact;
        updated++;
      } else {
        qtCusts.push(newEntry);
        added++;
      }
    });

    saveC(qtCusts);
    alert(`✅ Import ลูกค้าจาก CRM สำเร็จ!\nเพิ่มใหม่: ${added} ราย\nอัปเดต: ${updated} ราย\nรวมทั้งหมด: ${qtCusts.length} ราย`);

    // refresh dropdown ถ้ามีข้อความค้นหาอยู่แล้ว
    const input = document.getElementById('cust-th');
    if (input) onCS(input.value);

  } catch(e) {
    alert("❌ Error: " + e.message);
  }
}
