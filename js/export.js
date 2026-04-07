// ── EXCEL EXPORT ───────────────────────────────────────────────
function exportExcel() {
  if (typeof validateBeforeExport === "function" && !validateBeforeExport()) return;
  const wb = XLSX.utils.book_new();
  const items = getItems();
  const qtno    = gv('qt-no') || 'QT';
  const discPct = parseFloat(gv('discount')) || 0;
  const delFee  = parseFloat(gv('delivery-fee')) || 0;
  const vatOn   = gv('vat') === 'yes';
  const sub2    = items.reduce((s, i) => s + i.amt, 0);
  const discAmt  = sub2 * discPct / 100;
  const afterDisc = sub2 - discAmt;
  const vatAmt   = vatOn ? afterDisc * 0.07 : 0;  // VAT คำนวณจากยอดสินค้าเท่านั้น ไม่รวมค่าขนส่ง
  const grand    = afterDisc + delFee + vatAmt;

  // ══════════════════════════════════════════════════════════════
  // SHEET 1 : DATA  — format ตาม sample + เพิ่ม color, tel
  // ══════════════════════════════════════════════════════════════
  const dataRows = [
    ['orderId','customer','tel','custEn','addr','taxId','contact','shipAddr',
     'date','code','color','price','qty','rolls','total','vat','grandTotal']
  ];

  items.forEach(item => {
    const itemVat   = vatOn ? item.amt * 0.07 : 0;
    const grandItem = item.amt + itemVat;
    dataRows.push([
      gv('qt-no')       || '',
      gv('cust-th')     || '',
      gv('cust-tel')    || '',
      gv('cust-en')     || '',
      gv('cust-addr')   || '',
      gv('cust-tax')    || '',
      gv('cust-contact')|| '',
      gv('cust-deliver')|| '',
      gv('qt-date')     || '',
      item.code,
      item.color        || '',
      item.price,
      item.qty,
      item.rolls        || 0,
      item.amt,
      vatOn ? 'yes' : 'no',
      grandItem
    ]);
  });

  const wsData = XLSX.utils.aoa_to_sheet(dataRows);
  wsData['!cols'] = [
    {wch:18}, {wch:28}, {wch:14}, {wch:28},  // orderId, customer, tel, custEn
    {wch:30}, {wch:16}, {wch:18}, {wch:30},  // addr, taxId, contact, shipAddr
    {wch:12}, {wch:10}, {wch:20}, {wch:10},  // date, code, color, price
    {wch:8},  {wch:8},  {wch:12}, {wch:5}, {wch:14}  // qty, rolls, total, vat, grandTotal
  ];
  XLSX.utils.book_append_sheet(wb, wsData, 'DATA');

  // ══════════════════════════════════════════════════════════════
  // SHEET 2 : Quotation — ใบเสนอราคาเต็ม
  // ══════════════════════════════════════════════════════════════
  const rows = [];
  rows.push(['THAI KIM TEXTILE Co., Ltd.  |  บริษัท ไทยคิม เท็กไทล์ จำกัด']);
  rows.push(['92/25 ชั้น 11 อาคารสาทรธานี 2 ถ.สาทรเหนือ แขวงสีลม เขตบางรัก กทม. 10500 | Tel: 061-879-8222 | Tax ID: 0105565028514']);
  rows.push(['ใบเสนอราคา / QUOTATION']);
  rows.push([]);
  rows.push(['เลขที่:', gv('qt-no') || '', '', 'วันที่:', gv('qt-date') || '']);
  rows.push(['Sales:', gv('sales-name') || '', '', 'Manager:', gv('manager-name') || '']);
  rows.push([]);
  rows.push(['ผู้ขาย / SELLER', '', '', '', '', 'ผู้ซื้อ / BUYER']);
  rows.push(['THAI KIM TEXTILE Co., Ltd.', '', '', '', '', gv('cust-th') || '']);
  rows.push(['Tax ID: 0105565028514', '', '', '', '', gv('cust-en') || '']);
  rows.push(['Tel: 061-879-8222', '', '', '', '',
    'Tel: ' + (gv('cust-tel') || '') + (gv('cust-tax') ? '  |  Tax ID: ' + gv('cust-tax') : '')
  ]);
  rows.push([]);
  rows.push(['#', 'รหัสผ้า / Item', 'สี / Color', 'ส่วนประกอบ', 'กว้าง(cm)', 'น้ำหนัก(gsm)', 'ราคา/หลา(฿)', 'จำนวน(หลา)', 'จำนวน(ม้วน)', 'รวม(฿)', 'หมายเหตุ']);
  items.forEach(item => rows.push([
    item.num, item.code, item.color, item.comp,
    item.width, item.weight, item.price,
    item.qty, item.rolls || 0, item.amt, item.rem || ''
  ]));
  rows.push([]);
  rows.push(['', '', '', '', '', '', '', '', '', 'ยอดรวม / Subtotal', sub2]);
  if (discPct > 0) {
    rows.push(['', '', '', '', '', '', '', '', '', 'ส่วนลด ' + discPct + '%', -discAmt]);
    rows.push(['', '', '', '', '', '', '', '', '', 'หลังส่วนลด', afterDisc]);
  }
  rows.push(['', '', '', '', '', '', '', '', '', 'ค่าจัดส่ง / Delivery Fee', delFee]);
  if (vatOn) rows.push(['', '', '', '', '', '', '', '', '', 'VAT 7%', vatAmt]);
  rows.push(['', '', '', '', '', '', '', '', '', 'ยอดรวมสุทธิ / GRAND TOTAL', grand]);
  rows.push([]);
  rows.push(['เงื่อนไข / TERMS & CONDITIONS']);
  rows.push(['การจัดส่ง / Delivery:', gv('t-delivery') || '']);
  rows.push(['ระยะเวลาส่ง / Lead Time:', gv('t-lead') || '']);
  rows.push(['เงื่อนไขชำระ / Payment Term:', gv('t-payment') || '']);
  rows.push(['มัดจำ / Deposit:', gv('t-deposit') || '']);
  rows.push(['ยอดคงเหลือ / Balance:', gv('t-balance') || '']);
  rows.push(['อายุเอกสาร / Validity:', gv('t-valid') || '']);
  const rem = gv('remark');
  if (rem) rows.push(['หมายเหตุ:', rem]);
  rows.push([]);
  rows.push(['ข้อมูลธนาคาร / Bank Details']);
  const salesName = (gv('sales-name') || '').trim().toLowerCase();
  const isThip    = salesName === 'thip';

  if (vatOn) {
    rows.push(['ชื่อบัญชี:', 'บจก.ไทย คิม เท็กซ์ไทล์', 'ธนาคาร:', 'กสิกรไทย (K-Bank)', 'เลขบัญชี:', '129-1-17288-8']);
  } else if (isThip) {
    rows.push(['ชื่อบัญชี:', 'Mr.Zhang Jian', 'ธนาคาร:', 'ไทยพาณิชย์ (SCB)', 'เลขบัญชี:', '414-133416-2']);
  } else {
    rows.push(['ชื่อบัญชี:', 'Liu Hongwen', 'ธนาคาร:', 'ไทยพาณิชย์ (SCB)', 'เลขบัญชี:', '245-215274-8']);
  }

  const wsQt = XLSX.utils.aoa_to_sheet(rows);
  wsQt['!cols'] = [{wch:4},{wch:14},{wch:16},{wch:16},{wch:11},{wch:11},{wch:13},{wch:12},{wch:12},{wch:18},{wch:16}];
  XLSX.utils.book_append_sheet(wb, wsQt, 'Quotation');

  XLSX.writeFile(wb,qtno + '.xlsx', { bookType: 'xlsx', type: 'binary' });
}

// ── PDF EXPORT ─────────────────────────────────────────────────
async function exportPDF() {
  if (typeof validateBeforeExport === "function" && !validateBeforeExport()) return;
  const qtno = gv('qt-no') || 'QT';
  const el = document.getElementById('qdoc');
  const btns = document.querySelectorAll('.print-hide');
  btns.forEach(b => b.style.visibility = 'hidden');
  try {
    const prevScroll = el.parentElement.scrollTop;
    el.parentElement.scrollTop = 0;
    await new Promise(r => setTimeout(r, 100));
    const canvas = await html2canvas(el, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      scrollX: 0,
      scrollY: 0,
      windowWidth: el.scrollWidth,
      windowHeight: el.scrollHeight,
      x: 0,
      y: 0,
      width: el.offsetWidth,
      height: el.scrollHeight
    });
    el.parentElement.scrollTop = prevScroll;
    const imgData = canvas.toDataURL('image/jpeg', 0.97);
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const pgW = 210, pgH = 297;
    const imgH = (canvas.height * pgW) / canvas.width;
    let y = 0;
    while (y < imgH) {
      if (y > 0) pdf.addPage();
      pdf.addImage(imgData, 'JPEG', 0, -y, pgW, imgH);
      y += pgH;
    }
    pdf.save('Quotation_' + qtno + '.pdf');
  } finally {
    btns.forEach(b => b.style.visibility = '');
  }
}

// ── PNG EXPORT ─────────────────────────────────────────────────
async function exportPNG() {
  const qtno = gv('qt-no') || 'QT';
  const el = document.getElementById("qdoc");
  u();
  await document.fonts.ready;
  await new Promise(r => setTimeout(r, 200));
  const canvas = await html2canvas(el, {
    scale: 4,
    useCORS: true,
    backgroundColor: "#ffffff"
  });
  const link = document.createElement("a");
  link.download = "Quotation_" + qtno + ".png";
  link.href = canvas.toDataURL("image/png");
  link.click();
}



