const SHEET_NAME = 'Orders';
const DOC_FOLDER_NAME = 'Satija Purchase Docs';
const PURCHASE_FORM_SHEET_ID = '1bzFrGTL-oaGrrjyXgI8WJuKOzgJG2-9jcD8r80QR73w';
const OLD_FMS_SHEET_ID = '1XrykdT3AGzqeWxyxTjwW-wx_ah9g1vXsz7VaB90pVgQ';
const OLD_FMS_TAB_NAME = 'FMS';
// Column indices (0-based) in the old FMS tab. If the old sheet's layout
// ever changes, these need to be updated to match.
const OLD_FMS_COLS = {
  po: 1,
  PO1: [27, 28], PO2: [32, 33], PO3: [38, 39],
  PO4: [56, 57], PO5: [62, 63], PO6: [69, 70],
  confirmationPdf: 34,
  packingList: 41, vehicleNumber: 42, driverNumber: 43, transporterName: 44,
  invoiceNumber: 47, invoiceDate: 48, invoiceAmount: 50, uploadInvoice: 51, transferStatus: 54,
  leadTime: 55, driverStatus: 58,
  godownSlip: 64, grSignedCopy: 65,
  eWayBill: 74, grCopy: 75
};
const STEP_HOURS = { PO1: 0, PO2: 24, PO3: 48, PO4: 72, PO5: 96, PO6: 120 };

// ---------- ENTRY POINTS ----------

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  let result;
  try {
    const params = e.postData
      ? JSON.parse(e.postData.contents)
      : e.parameter;
    const action = params.action;

    if (action === 'getOrders') {
      result = { orders: getOrders() };
    } else if (action === 'saveOrder') {
      result = saveOrder(params.order);
    } else if (action === 'uploadDoc') {
      result = uploadDoc(params);
    } else if (action === 'resync') {
      const formResult = syncFromPurchaseForm();
      let fmsUpdated = 0;
      let fmsAvailable = true;
      try {
        const fmsResult = syncStatusFromOldFMS();
        fmsUpdated = fmsResult.updated;
      } catch (fmsErr) {
        // Old FMS Sheet not reachable (deleted, moved, or access revoked) -
        // that's fine, the new system keeps working independently on
        // Purchase Form data alone.
        fmsAvailable = false;
      }
      let backfillUpdated = 0;
      try {
        const backfillResult = backfillFullFields();
        backfillUpdated = backfillResult.updated || 0;
      } catch (backfillErr) {
        // Purchase Form / old FMS temporarily unreachable - resync itself still succeeded
      }
      result = {
        added: formResult.added,
        totalInForm: formResult.totalInForm,
        statusUpdated: fmsUpdated,
        oldFmsAvailable: fmsAvailable,
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
    if (['saveOrder','resync','dedupe','backfill','seedHistory'].includes(action)) {
      try { regenerateFlatView(); } catch (flatErr) { /* don't fail the real action over this */ }
    }
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- LIVE SYNC FROM PURCHASE FORM ----------

// Reads the real Purchase Form Sheet directly (by ID) and adds any PO that
// isn't already tracked here. Existing orders (and their step progress /
// invoice qty / uploaded docs) are never touched or overwritten.
function syncFromPurchaseForm() {
  const formSS = SpreadsheetApp.openById(PURCHASE_FORM_SHEET_ID);
  const formSheet = formSS.getSheets()[0];
  const data = formSheet.getDataRange().getValues();
  const header = data[0];
  const col = {};
  header.forEach((h, i) => { col[String(h).trim()] = i; });

  const existingIds = new Set(getOrders().map(o => String(o.id).trim()));
  const grouped = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const po = row[col['PO No']] ? String(row[col['PO No']]).trim() : '';
    if (!po) continue;
    if (!grouped[po]) {
      const poDate = row[col['PO Date']] || row[col['Date']];
      const createdAt = poDate instanceof Date ? poDate.getTime() : Date.now();
      grouped[po] = {
        id: po,
        mill: row[col['Supplier']] || row[col['Brand']] || '',
        poDate: poDate instanceof Date ? poDate.toISOString().slice(0, 10) : String(poDate || ''),
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

// Reads the OLD Purchase FMS Sheet directly and mirrors real step completion
// (and timestamps) into the new Orders tab. This is the ongoing source of
// truth for step status: whatever gets marked done in the old FMS shows up
// here too, every time this runs. It only ever moves a step from pending to
// done (never regresses an already-done step), so it's safe to run repeatedly.
function syncStatusFromOldFMS() {
  const oldSS = SpreadsheetApp.openById(OLD_FMS_SHEET_ID);
  const oldSheet = oldSS.getSheetByName(OLD_FMS_TAB_NAME);
  const data = oldSheet.getDataRange().getValues();

  // Row 1-6 are headers/merged title rows in the old sheet; real data starts at row 7.
  const rows = data.slice(6).filter(r => r[OLD_FMS_COLS.po]);

  const realStepsByPo = {};
  rows.forEach(r => {
    const po = String(r[OLD_FMS_COLS.po]).trim();
    if (!po) return;
    if (!realStepsByPo[po]) realStepsByPo[po] = {};
    Object.keys(STEP_HOURS).forEach(key => {
      const cols = OLD_FMS_COLS[key];
      const plannedCell = r[cols[0]];
      const actualCell = r[cols[1]];
      if (actualCell instanceof Date) {
        const actualMs = actualCell.getTime();
        const existing = realStepsByPo[po][key];
        if (!existing || actualMs > existing.actual) {
          realStepsByPo[po][key] = {
            actual: actualMs,
            planned: plannedCell instanceof Date ? plannedCell.getTime() : null
          };
        }
      }
    });
  });

  const orders = getOrders();
  let updated = 0;
  orders.forEach(o => {
    const real = realStepsByPo[String(o.id).trim()];
    if (!real) return;
    let changed = false;
    Object.keys(real).forEach(key => {
      if (o.steps[key] && o.steps[key].status !== 'done') {
        o.steps[key] = {
          status: 'done',
          planned: real[key].planned || o.steps[key].planned,
          actual: real[key].actual
        };
        changed = true;
      }
    });
    if (changed) {
      saveOrder(o);
      updated++;
    }
  });

  return { updated: updated, totalChecked: orders.length };
}

// ---------- SHEET HELPERS ----------

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['PO No', 'Mill', 'PO Date', 'CreatedAt', 'ItemsJSON', 'StepsJSON', 'ExtraJSON']);
  }
  return sheet;
}

// Reads every order (all POs) as structured JSON for the dashboard.
// ExtraJSON (column 7) carries order-level fields backfilled from the
// Purchase Form (GSTIN, address, insurance, freight, etc.) that don't
// belong inside items/steps.
function getOrders() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const orders = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const extra = safeParse(row[6], {});
    const order = {
      id: row[0],
      mill: row[1],
      poDate: row[2],
      createdAt: row[3],
      items: safeParse(row[4], []),
      steps: safeParse(row[5], {})
    };
    Object.keys(extra).forEach(k => { order[k] = extra[k]; });
    orders.push(order);
  }
  return orders;
}

// Creates a new PO row, or updates an existing one (matched by PO No).
function saveOrder(order) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const cleanId = String(order.id).trim();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === cleanId) { rowIndex = i + 1; break; }
  }
  const extra = {
    gstin: order.gstin, address: order.address, phone: order.phone,
    contactPerson: order.contactPerson, paymentTerms: order.paymentTerms,
    creditDays: order.creditDays, deliveryDestination: order.deliveryDestination,
    deliveryDate: order.deliveryDate, remarks: order.remarks, pdfLink: order.pdfLink,
    insurance: order.insurance, freight: order.freight
  };
  const rowValues = [
    cleanId,
    order.mill,
    order.poDate,
    order.createdAt,
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

// Removes duplicate PO rows (same PO No appearing more than once). Keeps
// whichever copy has the most completed steps, so real progress is never lost.
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
    const doneCount = Object.values(steps).filter(s => s && s.status === 'done').length;
    if (!bestByPo[id] || doneCount > bestByPo[id].doneCount) {
      bestByPo[id] = { row: row, doneCount: doneCount };
    }
  }
  sheet.clear();
  sheet.appendRow(['PO No', 'Mill', 'PO Date', 'CreatedAt', 'ItemsJSON', 'StepsJSON', 'ExtraJSON']);
  const cleanRows = Object.keys(bestByPo).map(id => bestByPo[id].row);
  if (cleanRows.length) {
    sheet.getRange(2, 1, cleanRows.length, cleanRows[0].length).setValues(cleanRows);
  }
  return { totalUnique: cleanRows.length };
}

// Adds any orders from the given historyOrders list that aren't already in
// the Sheet. Existing orders are never touched. Runs as one batch write for
// reliability (avoids many separate network round-trips from the browser).
function seedHistory(historyOrders) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const existingIds = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) existingIds.add(String(data[i][0]).trim());
  }

  const rowsToAdd = [];
  (historyOrders || []).forEach(o => {
    const id = String(o.id).trim();
    if (existingIds.has(id)) return;
    existingIds.add(id);
    rowsToAdd.push([
      id, o.mill, o.poDate, o.createdAt,
      JSON.stringify(o.items || []),
      JSON.stringify(o.steps || {})
    ]);
  });

  if (rowsToAdd.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 6).setValues(rowsToAdd);
  }
  return { added: rowsToAdd.length };
}

// ---------- ONE-TIME BACKFILL (safe: only fills blanks, never overwrites) ----------

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

  // Insurance/Freight can land on any one item-row of the shipment (not
  // necessarily the first), and 0 is a legitimate value (not "blank") - so
  // scan every row for that PO and take whichever one actually has a value.
  // One shipment = one freight/insurance value, no matter how many item rows.
  // Applies to every mill, not just one.
  const formRowsByPo = {};
  for (let i = 1; i < formData.length; i++) {
    const row = formData[i];
    const po = row[col['PO No']] ? String(row[col['PO No']]).trim() : '';
    if (!po) continue;
    (formRowsByPo[po] = formRowsByPo[po] || []).push(row);
  }
  const hasVal = v => v !== '' && v !== null && v !== undefined;
  const firstNonBlank = (rows, idx) => rows.map(r => r[idx]).find(hasVal);
  Object.keys(orderMap).forEach(po => {
    const o = orderMap[po];
    const rows = formRowsByPo[po];
    if (!rows) return;
    if (!hasVal(o.insurance)) {
      const v = firstNonBlank(rows, col['Insurance']);
      if (v !== undefined) o.insurance = v;
    }
    if (!hasVal(o.freight)) {
      const v = firstNonBlank(rows, col['Freight']);
      if (v !== undefined) o.freight = v;
    }
  });

  try {
    const oldSS = SpreadsheetApp.openById(OLD_FMS_SHEET_ID);
    const oldSheet = oldSS.getSheetByName(OLD_FMS_TAB_NAME);
    const fmsRows = oldSheet.getDataRange().getValues().slice(6).filter(r => r[OLD_FMS_COLS.po]);
    const byPo = {};
    fmsRows.forEach(r => {
      const po = String(r[OLD_FMS_COLS.po]).trim();
      (byPo[po] = byPo[po] || []).push(r);
    });
    const firstVal = (rows, idx) => rows.map(r => r[idx]).find(v => v);
    const dateStr = v => v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : (v || '');

    Object.keys(orderMap).forEach(po => {
      const o = orderMap[po];
      const rows = byPo[po];
      if (!rows) return;

      if (o.steps.PO2) {
        const link = firstVal(rows, OLD_FMS_COLS.confirmationPdf);
        o.steps.PO2.docs = o.steps.PO2.docs || {};
        if (link && !o.steps.PO2.docs['Confirmation PDF']) o.steps.PO2.docs['Confirmation PDF'] = { name: 'Confirmation PDF', url: link };
      }
      if (o.steps.PO3) {
        const st = o.steps.PO3;
        st.vehicleNumber = st.vehicleNumber || firstVal(rows, OLD_FMS_COLS.vehicleNumber) || '';
        st.driverNumber = st.driverNumber || firstVal(rows, OLD_FMS_COLS.driverNumber) || '';
        st.transporterName = st.transporterName || firstVal(rows, OLD_FMS_COLS.transporterName) || '';
        st.invoiceNumber = st.invoiceNumber || firstVal(rows, OLD_FMS_COLS.invoiceNumber) || '';
        st.invoiceDate = st.invoiceDate || dateStr(firstVal(rows, OLD_FMS_COLS.invoiceDate));
        st.invoiceAmount = st.invoiceAmount || firstVal(rows, OLD_FMS_COLS.invoiceAmount) || null;
        st.transferStatus = st.transferStatus || firstVal(rows, OLD_FMS_COLS.transferStatus) || '';
        st.docs = st.docs || {};
        const pkl = firstVal(rows, OLD_FMS_COLS.packingList);
        const upl = firstVal(rows, OLD_FMS_COLS.uploadInvoice);
        if (pkl && !st.docs['Packing List']) st.docs['Packing List'] = { name: 'Packing List', url: pkl };
        if (upl && !st.docs['Invoice PDF']) st.docs['Invoice PDF'] = { name: 'Invoice PDF', url: upl };
      }
      if (o.steps.PO4) {
        const st = o.steps.PO4;
        st.leadTime = st.leadTime || firstVal(rows, OLD_FMS_COLS.leadTime) || '';
        st.driverStatus = st.driverStatus || firstVal(rows, OLD_FMS_COLS.driverStatus) || '';
      }
      if (o.steps.PO5) {
        const st = o.steps.PO5;
        st.docs = st.docs || {};
        const gsl = firstVal(rows, OLD_FMS_COLS.godownSlip);
        const grs = firstVal(rows, OLD_FMS_COLS.grSignedCopy);
        if (gsl && !st.docs['Godown Slip']) st.docs['Godown Slip'] = { name: 'Godown Slip', url: gsl };
        if (grs && !st.docs['GR Signed Copy']) st.docs['GR Signed Copy'] = { name: 'GR Signed Copy', url: grs };
      }
      if (o.steps.PO6) {
        const st = o.steps.PO6;
        st.docs = st.docs || {};
        const ewb = firstVal(rows, OLD_FMS_COLS.eWayBill);
        const grc = firstVal(rows, OLD_FMS_COLS.grCopy);
        if (ewb && !st.docs['E-Way Bill']) st.docs['E-Way Bill'] = { name: 'E-Way Bill', url: ewb };
        if (grc && !st.docs['GR Copy']) st.docs['GR Copy'] = { name: 'GR Copy', url: grc };
      }
    });
  } catch (e) {
    // old FMS unreachable - order-level backfill from Purchase Form still applies
  }

  let updated = 0;
  Object.values(orderMap).forEach(o => { saveOrder(o); updated++; });
  return { updated: updated };
}

const FLAT_SHEET_NAME = 'Orders (Readable)';

// Reads a step's uploaded-doc URL by label (e.g. 'Invoice PDF'), if present.
function docUrl(step, label) {
  return (step && step.docs && step.docs[label] && step.docs[label].url) || '';
}

function regenerateFlatView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(FLAT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(FLAT_SHEET_NAME);
  sheet.clear();

  const header = [
    'PO No','Mill','PO Date','Product','Order Qty','Unit','Rate','Confirmed Qty',
    'PO1 Status','PO1 Planned','PO1 Actual',
    'PO2 Status','PO2 Planned','PO2 Actual','Confirmation PDF URL',
    'PO3 Status','PO3 Planned','PO3 Actual','Vehicle Number','Driver Number','Transporter Name',
    'Invoice Number','Invoice Date','Credit Days (Invoice)','Invoice Amount','Transfer Status',
    'Packing List URL','Invoice PDF URL',
    'Insurance (Confirmed)','Freight (Confirmed)',
    'PO4 Status','PO4 Planned','PO4 Actual','Lead Time','Driver Status',
    'PO5 Status','PO5 Planned','PO5 Actual','Godown Slip URL','GR Signed Copy URL',
    'PO6 Status','PO6 Planned','PO6 Actual','E-Way Bill URL','GR Copy URL',
    'GSTIN','Address','Phone','Contact Person','Payment Terms','Credit Days (Order)',
    'Delivery Destination','Delivery Date','Remarks','PDF Link',
    'Insurance (Purchase Form)','Freight (Purchase Form)'
  ];

  const asDate = ms => ms ? Utilities.formatDate(new Date(ms), 'Asia/Kolkata', 'dd/MM/yyyy HH:mm') : '';
  const rows = [];
  getOrders().forEach(o => {
    const st = o.steps || {};
    const p1=st.PO1||{}, p2=st.PO2||{}, p3=st.PO3||{}, p4=st.PO4||{}, p5=st.PO5||{}, p6=st.PO6||{};
    (o.items || []).forEach(it => {
      rows.push([
        o.id, o.mill, o.poDate, it.product, it.orderQty, it.unit, it.rate, it.invoiceQty,
        p1.status, asDate(p1.planned), asDate(p1.actual),
        p2.status, asDate(p2.planned), asDate(p2.actual), docUrl(p2, 'Confirmation PDF'),
        p3.status, asDate(p3.planned), asDate(p3.actual), p3.vehicleNumber, p3.driverNumber, p3.transporterName,
        p3.invoiceNumber, p3.invoiceDate, p3.invoiceCreditDays, p3.invoiceAmount, p3.transferStatus,
        docUrl(p3, 'Packing List'), docUrl(p3, 'Invoice PDF'),
        p3.insurance, p3.freight,
        p4.status, asDate(p4.planned), asDate(p4.actual), p4.leadTime, p4.driverStatus,
        p5.status, asDate(p5.planned), asDate(p5.actual), docUrl(p5, 'Godown Slip'), docUrl(p5, 'GR Signed Copy'),
        p6.status, asDate(p6.planned), asDate(p6.actual), docUrl(p6, 'E-Way Bill'), docUrl(p6, 'GR Copy'),
        o.gstin, o.address, o.phone, o.contactPerson, o.paymentTerms, o.creditDays,
        o.deliveryDestination, o.deliveryDate, o.remarks, o.pdfLink,
        o.insurance, o.freight
      ]);
    });
  });

  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}

// ---------- DRIVE UPLOAD ----------

// Saves an uploaded file (sent as base64) to a dedicated Drive folder and
// returns a shareable link. Called for every doc upload (Confirmation PDF,
// Invoice, Godown Slip, GR Copy, etc.)
function uploadDoc(params) {
  const folder = getOrCreateFolder(DOC_FOLDER_NAME);
  const rawBase64 = params.base64.indexOf(',') > -1
    ? params.base64.split(',')[1]
    : params.base64;
  const bytes = Utilities.base64Decode(rawBase64);
  const fileName = (params.orderId + '_' + params.stepKey + '_' + params.docLabel + '_' + params.fileName)
    .replace(/[\/\\?%*:|"<>]/g, '-');
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

// ---------- UTIL ----------

function safeParse(str, fallback) {
  try {
    return str ? JSON.parse(str) : fallback;
  } catch (e) {
    return fallback;
  }
}

function setupAutoSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoSync') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoSync').timeBased().everyMinutes(15).create();
}

function autoSync() {
  syncFromPurchaseForm();
  try { syncStatusFromOldFMS(); } catch (e) {}
}
