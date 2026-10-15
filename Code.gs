const SHEET_NAME = 'Orders';
const MASTER_DATA_SHEET_NAME = 'Master Data';
const FLAT_SHEET_NAME = 'Orders (Readable)';
const DOC_FOLDER_NAME = 'Satija Purchase Docs';
const PURCHASE_FORM_SHEET_ID = '1bzFrGTL-oaGrrjyXgI8WJuKOzgJG2-9jcD8r80QR73w';
const TIMEZONE = 'Asia/Kolkata';

const STEP_HOURS = { PO1: 0, PO2: 24, PO3: 48, PO4: 72, PO5: 96, PO6: 120 };

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const lock = LockService.getScriptLock();
  let result;
  try {
    lock.waitLock(30000);

    let params = {};
    if (e && e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      params = e.parameter;
      if (typeof params.order === 'string') {
        try { params.order = JSON.parse(params.order); } catch (ex) {}
      }
    }

    const action = params.action;
    if (action === 'getOrders') {
      result = { orders: getOrders() };
    } else if (action === 'saveOrder') {
      result = saveOrder(params.order);
    } else if (action === 'uploadDoc') {
      result = uploadDoc(params);
    } else if (action === 'resync') {
      const formResult = syncFromPurchaseForm();
      let backfillUpdated = 0;
      try { const r = backfillFullFields(); backfillUpdated = r.updated || 0; } catch (ex) {}
      result = {
        added: formResult.added,
        totalInForm: formResult.totalInForm,
        statusUpdated: 0,
        oldFmsAvailable: false,
        backfillUpdated: backfillUpdated
      };
    } else if (action === 'dedupe') {
      result = dedupeOrders();
    } else if (action === 'backfill') {
      result = backfillFullFields();
    } else if (action === 'seedHistory') {
      result = seedHistory(params.historyOrders);
    } else {
      result = { error: 'Unknown action: ' + action };
    }

    if (['saveOrder', 'resync', 'dedupe', 'backfill', 'seedHistory'].includes(action)) {
      try { regenerateFlatView(); } catch (ex) {}
    }
    if (['resync', 'backfill'].includes(action)) {
      try { regenerateMasterDataView(); } catch (ex) {}
    }
  } catch (err) {
    result = { error: err.message };
  } finally {
    try { lock.releaseLock(); } catch (ex) {}
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['PO No', 'Mill', 'PO Date', 'CreatedAt', 'ItemsJSON', 'StepsJSON', 'ExtraJSON']);
  }
  return sheet;
}

function getOrders() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const orders = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const extra = safeParse(row[6], {});
    const order = {
      id: String(row[0]).trim(),
      mill: row[1] || '',
      poDate: row[2] instanceof Date ? Utilities.formatDate(row[2], TIMEZONE, 'yyyy-MM-dd') : String(row[2] || ''),
      createdAt: Number(row[3]) || Date.now(),
      items: safeParse(row[4], []),
      steps: safeParse(row[5], {})
    };
    Object.keys(extra).forEach(k => { order[k] = extra[k]; });
    orders.push(order);
  }
  return orders;
}

function saveOrder(order) {
  if (!order || !order.id) return { error: 'Invalid order' };
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const cleanId = String(order.id).trim();
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === cleanId) {
      rowIndex = i + 1;
      break;
    }
  }

  // If transferStatus is Yes, clear planned & actual dates and mark steps as skipped
  if (order.steps && order.steps.PO3 && order.steps.PO3.transferStatus === 'Yes') {
    ['PO4', 'PO5', 'PO6'].forEach(function(k) {
      if (!order.steps[k]) order.steps[k] = {};
      order.steps[k].status = 'skipped';
      order.steps[k].planned = null;
      order.steps[k].actual = null;
    });
  }

  const extra = {
    gstin: order.gstin || '',
    address: order.address || '',
    phone: order.phone || '',
    contactPerson: order.contactPerson || '',
    paymentTerms: order.paymentTerms || '',
    creditDays: order.creditDays || '',
    deliveryDestination: order.deliveryDestination || '',
    deliveryDate: order.deliveryDate || '',
    remarks: order.remarks || '',
    pdfLink: order.pdfLink || '',
    insurance: order.insurance !== undefined ? order.insurance : '',
    freight: order.freight !== undefined ? order.freight : '',
    transportName: order.transportName || '',
    transporterNumber: order.transporterNumber || '',
    vehicleNumber: order.vehicleNumber || '',
    agreedTransportRate: order.agreedTransportRate || '',
    freightPaidByUs: order.freightPaidByUs || '',
    transportPdfLink: order.transportPdfLink || '',
    brand: order.brand || '',
    _legacySharedInvoiceQty: order._legacySharedInvoiceQty || null
  };

  const rowValues = [
    cleanId,
    order.mill || '',
    order.poDate || '',
    order.createdAt || Date.now(),
    JSON.stringify(order.items || []),
    JSON.stringify(order.steps || {}),
    JSON.stringify(extra)
  ];

  if (rowIndex === -1) {
    sheet.appendRow(rowValues);
  } else {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  }
  return { success: true };
}

function syncFromPurchaseForm() {
  const formSS = SpreadsheetApp.openById(PURCHASE_FORM_SHEET_ID);
  const formSheet = formSS.getSheets()[0];
  const data = formSheet.getDataRange().getValues();
  const col = {};
  data[0].forEach((h, i) => { col[String(h).trim()] = i; });

  const existingIds = new Set(getOrders().map(o => String(o.id).trim()));
  const grouped = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const po = row[col['PO No']] ? String(row[col['PO No']]).trim() : '';
    if (!po) continue;

    if (!grouped[po]) {
      const poDateRaw = row[col['PO Date']] || row[col['Date']];
      let poDateStr = '';
      let createdAt = Date.now();
      if (poDateRaw instanceof Date) {
        poDateStr = Utilities.formatDate(poDateRaw, TIMEZONE, 'yyyy-MM-dd');
        createdAt = poDateRaw.getTime();
      } else if (poDateRaw) {
        poDateStr = String(poDateRaw);
      }

      grouped[po] = {
        id: po,
        mill: row[col['Supplier']] || row[col['Brand']] || '',
        poDate: poDateStr,
        createdAt: createdAt,
        items: []
      };
    }

    grouped[po].items.push({
      id: 'i' + (grouped[po].items.length + 1),
      product: row[col['Product Name']] || '',
      orderQty: Number(row[col['Kg']]) || 0,
      unit: 'Kg',
      rate: Number(row[col['Rs/Kg']]) || 0,
      invoiceQty: null
    });
  }

  let added = 0;
  Object.keys(grouped).forEach(po => {
    if (existingIds.has(po)) return;
    const o = grouped[po];
    o.steps = buildDefaultSteps(o.createdAt);
    saveOrder(o);
    added++;
  });
  return { added: added, totalInForm: Object.keys(grouped).length };
}

function buildDefaultSteps(createdAt) {
  const steps = {};
  Object.keys(STEP_HOURS).forEach(key => {
    steps[key] = {
      status: key === 'PO1' ? 'done' : 'pending',
      planned: createdAt + STEP_HOURS[key] * 3600 * 1000,
      actual: key === 'PO1' ? createdAt : null
    };
  });
  return steps;
}

function dedupeOrders() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const bestByPo = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = String(row[0]).trim();
    if (!id) continue;
    let steps = {};
    try { steps = JSON.parse(row[5] || '{}'); } catch (e) {}
    const doneCount = Object.values(steps).filter(s => s && (s.status === 'done' || s.status === 'skipped')).length;
    if (!bestByPo[id] || doneCount > bestByPo[id].doneCount) {
      bestByPo[id] = { row: row, doneCount: doneCount };
    }
  }
  const cleanRows = Object.keys(bestByPo).map(id => bestByPo[id].row);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 7).setValues([['PO No', 'Mill', 'PO Date', 'CreatedAt', 'ItemsJSON', 'StepsJSON', 'ExtraJSON']]);
  if (cleanRows.length) {
    sheet.getRange(2, 1, cleanRows.length, 7).setValues(cleanRows.map(r => {
      while (r.length < 7) r.push('');
      return r.slice(0, 7);
    }));
  }
  return { totalUnique: cleanRows.length };
}

function seedHistory(historyOrders) {
  const existingIds = new Set(getOrders().map(o => String(o.id).trim()));
  let added = 0;
  (historyOrders || []).forEach(o => {
    const id = String(o.id).trim();
    if (existingIds.has(id)) return;
    existingIds.add(id);
    saveOrder(o);
    added++;
  });
  return { added: added };
}

function backfillFullFields() {
  const orders = getOrders();
  const orderMap = {};
  orders.forEach(o => { orderMap[String(o.id).trim()] = o; });

  const formSS = SpreadsheetApp.openById(PURCHASE_FORM_SHEET_ID);
  const formSheet = formSS.getSheets()[0];
  const formData = formSheet.getDataRange().getValues();
  const col = {};
  formData[0].forEach((h, i) => { col[String(h).trim()] = i; });

  const seenPo = {};
  for (let i = 1; i < formData.length; i++) {
    const row = formData[i];
    const po = row[col['PO No']] ? String(row[col['PO No']]).trim() : '';
    if (!po || seenPo[po]) continue;
    seenPo[po] = true;
    const o = orderMap[po];
    if (!o) continue;
    o.gstin = o.gstin || row[col['GSTIN']] || '';
    o.address = o.address || row[col['Address']] || '';
    o.phone = o.phone || row[col['Phone']] || '';
    o.contactPerson = o.contactPerson || row[col['Contact Person']] || '';
    o.paymentTerms = o.paymentTerms || row[col['Payment Terms']] || '';
    o.creditDays = o.creditDays || row[col['Credit Days']] || '';
    o.remarks = o.remarks || row[col['Remarks']] || '';
    o.pdfLink = o.pdfLink || row[col['PDF Link']] || '';
    o.deliveryDestination = o.deliveryDestination || row[col['Delivery Destination']] || '';
  }

  const formRowsByPo = {};
  for (let i = 1; i < formData.length; i++) {
    const row = formData[i];
    const po = row[col['PO No']] ? String(row[col['PO No']]).trim() : '';
    if (!po) continue;
    (formRowsByPo[po] = formRowsByPo[po] || []).push(row);
  }

  const hasVal = v => v !== '' && v !== null && v !== undefined;
  const firstNonBlank = (rows, idx) => rows.map(r => r[idx]).find(hasVal);
  const dateOrRaw = v => v instanceof Date ? Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd') : v;

  Object.keys(orderMap).forEach(po => {
    const o = orderMap[po];
    const rows = formRowsByPo[po];
    if (!rows) return;
    if (!hasVal(o.insurance)) { const v = firstNonBlank(rows, col['Insurance']); if (v !== undefined) o.insurance = v; }
    if (!hasVal(o.freight)) { const v = firstNonBlank(rows, col['Freight']); if (v !== undefined) o.freight = v; }
    if (!hasVal(o.deliveryDate)) { const v = firstNonBlank(rows, col['Delivery Date']); if (v !== undefined) o.deliveryDate = dateOrRaw(v); }
    if (!hasVal(o.transportName)) { const v = firstNonBlank(rows, col['Transport Name']); if (v !== undefined) o.transportName = v; }
    if (!hasVal(o.transporterNumber)) { const v = firstNonBlank(rows, col['Transporter Number']); if (v !== undefined) o.transporterNumber = v; }
    if (!hasVal(o.vehicleNumber)) { const v = firstNonBlank(rows, col['Vehicle Number']); if (v !== undefined) o.vehicleNumber = v; }
    if (!hasVal(o.agreedTransportRate)) { const v = firstNonBlank(rows, col['Agree Rate of Transport']); if (v !== undefined) o.agreedTransportRate = v; }
    if (!hasVal(o.freightPaidByUs)) { const v = firstNonBlank(rows, col['Freight (Paid by Us)']); if (v !== undefined) o.freightPaidByUs = v; }
    if (!hasVal(o.transportPdfLink)) { const v = firstNonBlank(rows, col['Transport PDF Link']); if (v !== undefined) o.transportPdfLink = v; }
  });

  let updated = 0;
  Object.values(orderMap).forEach(o => { saveOrder(o); updated++; });
  return { updated: updated };
}

function regenerateMasterDataView() {
  const formSS = SpreadsheetApp.openById(PURCHASE_FORM_SHEET_ID);
  const formSheet = formSS.getSheets()[0];
  const formData = formSheet.getDataRange().getValues();
  if (!formData.length) return;
  const header = formData[0].map(h => String(h).trim());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MASTER_DATA_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(MASTER_DATA_SHEET_NAME);
  sheet.clearContents();
  const rows = formData.slice(1).filter(row => row.some(v => v !== '' && v !== null)).map(row => row.map(v => v instanceof Date ? Utilities.formatDate(v, TIMEZONE, 'dd/MM/yyyy') : v));
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}

function docUrl(step, label) { return (step && step.docs && step.docs[label] && step.docs[label].url) || ''; }

function regenerateFlatView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(FLAT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(FLAT_SHEET_NAME);
  sheet.clearContents();
  const header = [
    'PO No', 'Mill', 'PO Date', 'Product', 'Order Qty', 'Unit', 'Rate', 'Confirmed Qty',
    'PO1 Status', 'PO1 Planned', 'PO1 Actual',
    'PO2 Status', 'PO2 Planned', 'PO2 Actual', 'Confirmation PDF URL',
    'PO3 Status', 'PO3 Planned', 'PO3 Actual', 'Vehicle Number', 'Driver Number', 'Transporter Name', 'Invoice Number', 'Invoice Date', 'Credit Days (Invoice)', 'Invoice Amount', 'Transfer Status', 'Packing List URL', 'Invoice PDF URL', 'Insurance (Confirmed)', 'Freight (Confirmed)',
    'PO4 Status', 'PO4 Planned', 'PO4 Actual', 'Lead Time', 'Driver Status',
    'PO5 Status', 'PO5 Planned', 'PO5 Actual', 'Godown Slip URL', 'GR Signed Copy URL',
    'PO6 Status', 'PO6 Planned', 'PO6 Actual', 'E-Way Bill Number', 'E-Way Bill URL', 'GR Copy URL',
    'GSTIN', 'Address', 'Phone', 'Contact Person', 'Payment Terms', 'Credit Days (Order)', 'Delivery Destination', 'Delivery Date', 'Remarks', 'PDF Link', 'Insurance (Purchase Form)', 'Freight (Purchase Form)', 'Transport Name (Purchase Form)', 'Transporter Number (Purchase Form)', 'Vehicle Number (Purchase Form)', 'Agreed Transport Rate', 'Freight Paid By Us', 'Transport PDF Link'
  ];
  const asDate = ms => ms ? Utilities.formatDate(new Date(ms), TIMEZONE, 'dd/MM/yyyy HH:mm') : '';
  const rows = [];
  getOrders().forEach(o => {
    const st = o.steps || {};
    const p1 = st.PO1 || {}, p2 = st.PO2 || {}, p3 = st.PO3 || {}, p4 = st.PO4 || {}, p5 = st.PO5 || {}, p6 = st.PO6 || {};
    const isTransferred = p3.transferStatus === 'Yes';
    (o.items || []).forEach(it => {
      rows.push([
        o.id, o.mill, o.poDate, it.product, it.orderQty, it.unit, it.rate, it.invoiceQty,
        p1.status, asDate(p1.planned), asDate(p1.actual),
        p2.status, asDate(p2.planned), asDate(p2.actual), docUrl(p2, 'Confirmation PDF'),
        p3.status, asDate(p3.planned), asDate(p3.actual),
        p3.vehicleNumber, p3.driverNumber, p3.transporterName, p3.invoiceNumber, p3.invoiceDate, p3.invoiceCreditDays, p3.invoiceAmount, p3.transferStatus, docUrl(p3, 'Packing List'), docUrl(p3, 'Invoice PDF'), p3.insurance, p3.freight,
        isTransferred ? 'skipped' : p4.status, isTransferred ? '' : asDate(p4.planned), isTransferred ? '' : asDate(p4.actual), isTransferred ? '' : (p4.leadTime || ''), isTransferred ? '' : (p4.driverStatus || ''),
        isTransferred ? 'skipped' : p5.status, isTransferred ? '' : asDate(p5.planned), isTransferred ? '' : asDate(p5.actual), isTransferred ? '' : docUrl(p5, 'Godown Slip'), isTransferred ? '' : docUrl(p5, 'GR Signed Copy'),
        isTransferred ? 'skipped' : p6.status, isTransferred ? '' : asDate(p6.planned), isTransferred ? '' : asDate(p6.actual), isTransferred ? '' : (p6.ewayBillNumber || ''), isTransferred ? '' : docUrl(p6, 'E-Way Bill'), isTransferred ? '' : docUrl(p6, 'GR Copy'),
        o.gstin, o.address, o.phone, o.contactPerson, o.paymentTerms, o.creditDays, o.deliveryDestination, o.deliveryDate, o.remarks, o.pdfLink, o.insurance, o.freight, o.transportName, o.transporterNumber, o.vehicleNumber, o.agreedTransportRate, o.freightPaidByUs, o.transportPdfLink
      ]);
    });
  });
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}

function uploadDoc(params) {
  const folder = getOrCreateFolder(DOC_FOLDER_NAME);
  const rawBase64 = params.base64.indexOf(',') > -1 ? params.base64.split(',')[1] : params.base64;
  const cleanBase64 = rawBase64.replace(/\s/g, '');
  const bytes = Utilities.base64Decode(cleanBase64);
  const fileName = (params.orderId + '_' + params.stepKey + '_' + params.docLabel + '_' + params.fileName).replace(/[\/\\?%*:|"<>]/g, '-');
  const blob = Utilities.newBlob(bytes, params.mimeType || 'application/octet-stream', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: file.getUrl(), id: file.getId(), name: fileName };
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function safeParse(str, fallback) {
  try { return str ? JSON.parse(str) : fallback; } catch (e) { return fallback; }
}

function setupAutoSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'autoSync') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('autoSync').timeBased().everyMinutes(15).create();
}

function autoSync() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    syncFromPurchaseForm();
    try { backfillFullFields(); } catch (e) {}
    try { regenerateMasterDataView(); } catch (e) {}
    try { regenerateFlatView(); } catch (e) {}
  } catch (e) {
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function resetPO126And127ToPending() {
  const targetPOs = ['PO-SP/26-27/0126', 'PO-SP/26-27/0127'];
  const orders = getOrders();
  let count = 0;
  
  orders.forEach(o => {
    if (targetPOs.includes(String(o.id).trim())) {
      o.steps = buildDefaultSteps(o.createdAt || Date.now());
      saveOrder(o);
      count++;
    }
  });
  
  regenerateFlatView();
  Logger.log('Reset completed for ' + count + ' orders.');
}

function resetFullOrders126And127() {
  const targetPOs = ['PO-SP/26-27/0126', 'PO-SP/26-27/0127'];
  const orders = getOrders();
  let count = 0;

  orders.forEach(o => {
    if (targetPOs.includes(String(o.id).trim())) {
      const createdAt = o.createdAt || Date.now();

      // 1. Reset all items (clear invoice quantities)
      if (Array.isArray(o.items)) {
        o.items.forEach(it => {
          it.invoiceQty = null;
        });
      }

      // 2. Reset all steps completely (no docs, no vehicle/invoice/driver fields)
      o.steps = {
        PO1: { status: 'done', planned: createdAt, actual: createdAt },
        PO2: { status: 'pending', planned: createdAt + (24 * 3600 * 1000), actual: null },
        PO3: { status: 'pending', planned: createdAt + (48 * 3600 * 1000), actual: null },
        PO4: { status: 'pending', planned: createdAt + (72 * 3600 * 1000), actual: null },
        PO5: { status: 'pending', planned: createdAt + (96 * 3600 * 1000), actual: null },
        PO6: { status: 'pending', planned: createdAt + (120 * 3600 * 1000), actual: null }
      };

      // 3. Save clean order
      saveOrder(o);
      count++;
    }
  });

  regenerateFlatView();
  Logger.log('Completely reset ' + count + ' orders to fresh pending state.');
}
