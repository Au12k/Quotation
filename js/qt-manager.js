// ============================================================
//  qt-manager.js
//  1. Save / Load ใบเสนอราคา (localStorage)
//  2. เลข QT อัตโนมัติ
//  3. ค้นหา / เปิด QT เก่า
//  4. Validation ก่อน Export
// ============================================================

const QT_KEY = "qt_drafts";   // localStorage key สำหรับเก็บ QT ทั้งหมด

// ─── helpers ────────────────────────────────────────────────
function loadDrafts() {
  try { return JSON.parse(localStorage.getItem(QT_KEY)) || []; }
  catch { return []; }
}
function saveDrafts(arr) {
  localStorage.setItem(QT_KEY, JSON.stringify(arr));
}

// ─── 1. เลข QT อัตโนมัติ ────────────────────────────────────
// format: FTSQ + YYMMDD + running 2 หลัก เช่น FTSQ2603260001
function genQTNo() {
  const now   = new Date();
  const yy    = String(now.getFullYear()).slice(2);
  const mm    = String(now.getMonth() + 1).padStart(2, "0");
  const dd    = String(now.getDate()).padStart(2, "0");
  const prefix = `FTSQ${yy}${mm}${dd}`;

  const drafts  = loadDrafts();
  const today   = drafts.filter(d => (d.qtno || "").startsWith(prefix));
  const running = String(today.length + 1).padStart(4, "0");
  return `${prefix}${running}`;
}

function autoQTNo() {
  const el = document.getElementById("qt-no");
  if (!el) return;
  if (el.value.trim()) {
    if (!confirm("มีเลข QT อยู่แล้ว ต้องการสร้างใหม่ไหม?")) return;
  }
  el.value = genQTNo();
  if (typeof u === "function") u();
}

// ─── 2. รวบรวมข้อมูลทั้งใบ ──────────────────────────────────
function collectDraft() {
  const items = typeof getItems === "function" ? getItems() : [];
  return {
    id:          Date.now(),
    savedAt:     new Date().toISOString(),
    qtno:        gv("qt-no"),
    qtdate:      gv("qt-date"),
    salesName:   gv("sales-name"),
    managerName: gv("manager-name"),
    custTh:      gv("cust-th"),
    custEn:      gv("cust-en"),
    custAddr:    gv("cust-addr"),
    custTax:     gv("cust-tax"),
    custTel:     gv("cust-tel"),
    custContact: gv("cust-contact"),
    custDeliver: gv("cust-deliver"),
    discount:    gv("discount"),
    deliveryFee: gv("delivery-fee"),
    vat:         gv("vat"),
    tDelivery:   gv("t-delivery"),
    tLead:       gv("t-lead"),
    tValid:      gv("t-valid"),
    tPayment:    gv("t-payment"),
    tDeposit:    gv("t-deposit"),
    tBalance:    gv("t-balance"),
    remark:      gv("remark"),
    items:       items,
  };
}

// ─── 3. กู้คืนข้อมูลลงฟอร์ม ─────────────────────────────────
function restoreDraft(d) {
  sv("qt-no",        d.qtno        || "");
  sv("qt-date",      d.qtdate      || "");
  sv("sales-name",   d.salesName   || "");
  sv("manager-name", d.managerName || "");
  sv("cust-th",      d.custTh      || "");
  sv("cust-en",      d.custEn      || "");
  sv("cust-addr",    d.custAddr    || "");
  sv("cust-tax",     d.custTax     || "");
  sv("cust-tel",     d.custTel     || "");
  sv("cust-contact", d.custContact || "");
  sv("cust-deliver", d.custDeliver || "");
  sv("discount",     d.discount    || "0");
  sv("delivery-fee", d.deliveryFee || "0");
  sv("remark",       d.remark      || "");
  sv("t-lead",       d.tLead       || "");
  sv("t-valid",      d.tValid      || "");
  sv("t-deposit",    d.tDeposit    || "");

  // VAT & Payment
  const vatEl = document.getElementById("vat");
  if (vatEl && d.vat) vatEl.value = d.vat;
  const payEl = document.getElementById("t-payment");
  if (payEl && d.tPayment) payEl.value = d.tPayment;

  // สร้าง items ใหม่
  const wrap = document.getElementById("items-wrap");
  if (wrap) wrap.innerHTML = "";
  if (d.items && d.items.length) {
    d.items.forEach(item => {
      if (typeof addItem === "function") addItem();
      // หา row ล่าสุดที่เพิ่งสร้าง
      const allRows = document.querySelectorAll("#items-wrap .item-row");
      const row = allRows[allRows.length - 1];
      if (!row) return;
      const rowId = row.id.replace("item-", "");
      sv(`i${rowId}-code`,   item.code   || "");
      sv(`i${rowId}-color`,  item.color  || "");
      sv(`i${rowId}-comp`,   item.comp   || "");
      sv(`i${rowId}-width`,  item.width  || "");
      sv(`i${rowId}-weight`, item.weight || "");
      sv(`i${rowId}-price`,  item.price  || "");
      sv(`i${rowId}-qty`,    item.qty    || "");
      sv(`i${rowId}-rolls`,  item.rolls  || "");
      sv(`i${rowId}-rem`,    item.rem    || "");
    });
  }

  if (typeof u === "function") u();
}

// ─── 4. Save ────────────────────────────────────────────────
function saveQT() {
  try {
    const draft = collectDraft();
    if (!draft.qtno) {
      alert("กรุณากรอกเลขที่ใบเสนอราคาก่อน Save");
      return;
    }

  const drafts = loadDrafts();
  const idx    = drafts.findIndex(d => d.qtno === draft.qtno);

  if (idx >= 0) {
    if (!confirm(`มี QT ${draft.qtno} อยู่แล้ว\nต้องการ overwrite ไหม?`)) return;
    drafts[idx] = { ...drafts[idx], ...draft };
  } else {
    drafts.push(draft);
  }

    saveDrafts(drafts);
    showToast(`💾 บันทึก ${draft.qtno} สำเร็จ`);
    renderQTList();
  } catch(err) {
    alert("❌ Save error: " + err.message);
    console.error("saveQT error:", err);
  }
}

// ─── Auto-save ทุก 30 วินาที ──────────────────────────────────
let autoSaveTimer = null;
function startAutoSave() {
  if (autoSaveTimer) clearInterval(autoSaveTimer);
  autoSaveTimer = setInterval(() => {
    const draft = collectDraft();
    if (!draft.qtno) return;
    const drafts = loadDrafts();
    const idx = drafts.findIndex(d => d.qtno === draft.qtno);
    if (idx >= 0) {
      drafts[idx] = { ...drafts[idx], ...draft, savedAt: new Date().toISOString() };
      saveDrafts(drafts);
      showToast("💾 Auto-saved");
    }
  }, 30000);
}

// ─── 5. Load / ค้นหา QT เก่า ────────────────────────────────
function openQTManager() {
  renderQTList();
  const modal = document.getElementById("qt-modal");
  if (modal) modal.style.display = "flex";
}

function closeQTManager() {
  const modal = document.getElementById("qt-modal");
  if (modal) modal.style.display = "none";
}

function renderQTList(filter) {
  const list = document.getElementById("qt-modal-list");
  if (!list) return;

  let drafts = loadDrafts().slice().reverse(); // ใหม่สุดก่อน
  if (filter) {
    const q = filter.toLowerCase();
    drafts = drafts.filter(d =>
      (d.qtno    || "").toLowerCase().includes(q) ||
      (d.custTh  || "").toLowerCase().includes(q) ||
      (d.qtdate  || "").includes(q)
    );
  }

  if (!drafts.length) {
    list.innerHTML = `<div style="text-align:center;color:#aaa;padding:24px">ไม่พบรายการ</div>`;
    return;
  }

  list.innerHTML = drafts.map(d => {
    const saved = d.savedAt ? new Date(d.savedAt).toLocaleString("th-TH") : "";
    return `
      <div class="qt-list-item">
        <div class="qt-list-info">
          <div class="qt-list-no">${d.qtno || "—"}</div>
          <div class="qt-list-cust">${d.custTh || "—"}</div>
          <div class="qt-list-meta">${d.qtdate || ""} · บันทึก: ${saved}</div>
        </div>
        <div class="qt-list-actions">
          <button class="qt-btn-open" onclick="loadQT('${d.qtno}')">📂 เปิด</button>
          <button class="qt-btn-del"  onclick="deleteQT('${d.qtno}')">🗑</button>
        </div>
      </div>`;
  }).join("");
}

function loadQT(qtno) {
  const draft = loadDrafts().find(d => d.qtno === qtno);
  if (!draft) return alert("ไม่พบ QT นี้");
  if (!confirm(`เปิด QT ${qtno}?\nข้อมูลปัจจุบันในฟอร์มจะถูกแทนที่`)) return;
  restoreDraft(draft);
  closeQTManager();
  startAutoSave();
}

function deleteQT(qtno) {
  if (!confirm(`ลบ QT ${qtno} ถาวร?`)) return;
  const drafts = loadDrafts().filter(d => d.qtno !== qtno);
  saveDrafts(drafts);
  renderQTList(document.getElementById("qt-search")?.value || "");
}

// ─── 6. Validation ก่อน Export ──────────────────────────────
function validateBeforeExport() {
  const errors = [];

  if (!gv("qt-no"))   errors.push("• กรอกเลขที่ใบเสนอราคา");
  if (!gv("qt-date")) errors.push("• กรอกวันที่");
  if (!gv("cust-th")) errors.push("• กรอกชื่อลูกค้า");

  const items = typeof getItems === "function" ? getItems() : [];
  if (!items.length) {
    errors.push("• เพิ่มรายการสินค้าอย่างน้อย 1 รายการ");
  } else {
    items.forEach((item, i) => {
      if (!item.code)          errors.push(`• รายการ ${i+1}: กรอกรหัสผ้า`);
      if (!item.price || item.price <= 0) errors.push(`• รายการ ${i+1}: ราคา/หลา ต้องมากกว่า 0`);
      if (!item.qty   || item.qty   <= 0) errors.push(`• รายการ ${i+1}: จำนวนหลา ต้องมากกว่า 0`);
    });
  }

  if (errors.length) {
    alert("❌ กรุณาแก้ไขก่อน Export:\n\n" + errors.join("\n"));
    return false;
  }
  return true;
}

// ─── 7. Toast notification ──────────────────────────────────
function showToast(msg, duration = 2500) {
  let toast = document.getElementById("qt-toast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "qt-toast";
    toast.style.cssText = `
      position:fixed;bottom:24px;right:24px;
      background:#1a9a58;color:#fff;
      padding:10px 20px;border-radius:10px;
      font-size:0.9rem;font-weight:600;
      z-index:9999;opacity:0;
      transition:opacity 0.3s;
      font-family:'Sarabun',sans-serif;
    `;
    document.body.appendChild(toast);
  }
  toast.textContent = msg;
  toast.style.opacity = "1";
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => toast.style.opacity = "0", duration);
}

// ─── 8. Modal HTML (inject เมื่อ DOM ready) ─────────────────
function injectQTModal() {
  if (document.getElementById("qt-modal")) return;

  const modal = document.createElement("div");
  modal.id = "qt-modal";
  modal.style.cssText = `
    display:none;position:fixed;inset:0;
    background:rgba(0,0,0,0.5);z-index:5000;
    align-items:center;justify-content:center;
  `;
  modal.innerHTML = `
    <div style="background:#fff;border-radius:14px;width:580px;max-width:95vw;max-height:80vh;
                display:flex;flex-direction:column;overflow:hidden;box-shadow:0 8px 32px rgba(0,0,0,0.2);">
      <div style="padding:18px 20px;border-bottom:1px solid #eee;display:flex;align-items:center;justify-content:space-between;">
        <div style="font-size:1.1rem;font-weight:700;color:#1a1a2e;">📂 ใบเสนอราคาที่บันทึกไว้</div>
        <button onclick="closeQTManager()" style="background:none;border:none;font-size:1.4rem;cursor:pointer;color:#999;">×</button>
      </div>
      <div style="padding:12px 16px;border-bottom:1px solid #eee;">
        <input id="qt-search" type="text" placeholder="🔍 ค้นหาเลข QT หรือชื่อลูกค้า..."
          oninput="renderQTList(this.value)"
          style="width:100%;padding:8px 12px;border:1.5px solid #d0d5e8;border-radius:8px;
                 font-size:0.9rem;outline:none;font-family:'Sarabun',sans-serif;">
      </div>
      <div id="qt-modal-list" style="overflow-y:auto;flex:1;padding:8px 0;"></div>
    </div>
  `;
  document.body.appendChild(modal);

  // ปิดเมื่อคลิก backdrop
  modal.addEventListener("click", e => { if (e.target === modal) closeQTManager(); });
}

// ─── Init ────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", () => {
  try {
    injectQTModal();

    // ไม่ auto-generate แล้ว — user พิมพ์เองหรือกดปุ่ม 🔢
    startAutoSave();
    console.log("✅ qt-manager.js loaded OK");
  } catch(err) {
    console.error("❌ qt-manager init error:", err);
  }
});
