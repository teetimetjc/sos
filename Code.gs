// ============================================================
// S.O.S. Septic — Digital Work Order System
// Apps Script backend: serves the app and handles data I/O
// ============================================================

var SECRET = PropertiesService.getScriptProperties().getProperty('SECRET') || '';
var SHEET_NAME   = 'WorkOrders';
var CALLIN_SHEET = 'CallIns';

// WorkOrders tab columns (A1:AK1, 37 headers):
//   ID | WO Number | Type | Date | Time | Bill To | Job | Phone 1 | Phone 2 |
//   Pumpout Ordered | Technician | Pump Type | Year Built | Water Level | Drainfield |
//   Scum | Sludge | Tank Sound | Compartment | Outlet T | Material | Trap Location |
//   Directions | Schedule | LPO | Next Pump | Special Notes | Line Items JSON | Total |
//   Billing | Terms | Check Num | Payment Amt | Comments | Signed By | Saved At | Client Email

// CallIns tab columns (A1:AE1, 31 headers):
//   ID | CI Number | Rep | Date | Scheduled Date | Scheduled Time |
//   Bill To Name | Bill To Address | Bill To Phone |
//   Job Address | Job Phone | Email |
//   Reason | Reason Detail | Beds | Baths | Year Built | LPO |
//   Lift Station | ATU | Vac Job | Tank Count | Directions |
//   Technicians | Extra Time | Price | Overfull | Field Runback |
//   Payment Type | Billing Type | Saved At

// --------------- Spreadsheet menu ---------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SOS Tools')
    .addItem('Export Selected Row as PDF', 'exportRowAsPdf')
    .addToUi();
}

function exportRowAsPdf() {
  var ui    = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { ui.alert('WorkOrders sheet not found.'); return; }

  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('Please click on a work order row first (not the header row).');
    return;
  }

  var lastCol  = sheet.getLastColumn();
  var headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var values   = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  var raw = {};
  headers.forEach(function(h, i) { raw[h] = values[i]; });

  function fmtVal(v) {
    if (!v && v !== 0) return '';
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    return String(v);
  }

  var lineItems = [];
  try { lineItems = JSON.parse(raw['Line Items JSON'] || '[]'); } catch(e) {}

  var data = {
    type:        fmtVal(raw['Type']),
    date:        fmtVal(raw['Date']),
    time:        fmtVal(raw['Time']),
    billTo:      fmtVal(raw['Bill To']),
    job:         fmtVal(raw['Job']),
    phone1:      fmtVal(raw['Phone 1']),
    phone2:      fmtVal(raw['Phone 2']),
    clientEmail: fmtVal(raw['Client Email']),
    technician:  fmtVal(raw['Technician']),
    pumpType:    fmtVal(raw['Pump Type']),
    yearBuilt:   fmtVal(raw['Year Built']),
    waterLevel:  fmtVal(raw['Water Level']),
    drainfield:  fmtVal(raw['Drainfield']),
    scum:        fmtVal(raw['Scum']),
    sludge:      fmtVal(raw['Sludge']),
    tankSound:   fmtVal(raw['Tank Sound']),
    compartment: fmtVal(raw['Compartment']),
    outletT:     fmtVal(raw['Outlet T']),
    material:    fmtVal(raw['Material']),
    trapLocation:fmtVal(raw['Trap Location']),
    directions:  fmtVal(raw['Directions']),
    lpo:         fmtVal(raw['LPO']),
    nextPump:    fmtVal(raw['Next Pump']),
    specialNotes:fmtVal(raw['Special Notes']),
    lineItems:   lineItems,
    total:       fmtVal(raw['Total']),
    billing:     fmtVal(raw['Billing']),
    terms:       fmtVal(raw['Terms']),
    checkNum:    fmtVal(raw['Check Num']),
    paymentAmt:  fmtVal(raw['Payment Amt']),
    comments:    fmtVal(raw['Comments']),
    signedBy:    fmtVal(raw['Signed By'])
  };

  var woNum = fmtVal(raw['WO Number']) || 'WO-?????';
  var tmpl  = HtmlService.createTemplateFromFile('PdfExport');
  tmpl.rowData = JSON.stringify(data);
  tmpl.woNum   = woNum;
  ui.showModalDialog(tmpl.evaluate().setWidth(480).setHeight(300), 'Export Work Order — ' + woNum);
}

// --------------- HTTP entry points ---------------

function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var action = params.action || '';
  var sheet  = params.sheet  || '';

  if (action === 'list' && sheet === 'callins') {
    return jsonResponse(listCallIns(params.secret));
  }
  if (action === 'list') {
    return jsonResponse(listOrders(params.secret));
  }
  if (action === 'lookup') {
    return jsonResponse(lookupCustomer(params.phone, params.secret));
  }

  // Serve the HTML app
  var tmpl = HtmlService.createTemplateFromFile('App');
  tmpl.secret = SECRET;
  return tmpl.evaluate()
    .setTitle('S.O.S. Septic Work Orders')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    if (payload.secret !== SECRET) {
      return jsonResponse({ ok: false, error: 'Unauthorized' });
    }
    if (payload.sheetType === 'callin') {
      var ciNum = saveCallIn(payload);
      return jsonResponse({ ok: true, ciNumber: ciNum });
    }
    var woNum = saveOrder(payload);
    return jsonResponse({ ok: true, woNumber: woNum });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// --------------- Work order helpers ---------------

function saveOrder(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var woNum = nextWoNumber();
  var id    = Utilities.getUuid();
  var now   = new Date();

  sheet.appendRow([
    id,
    woNum,
    data.type || '',
    data.date || '',
    data.time || '',
    data.billTo || '',
    data.job || '',
    data.phone1 || '',
    data.phone2 || '',
    data.pumpoutOrdered || '',
    data.technician || '',
    data.pumpType || '',
    data.yearBuilt || '',
    data.waterLevel || '',
    data.drainfield || '',
    data.scum || '',
    data.sludge || '',
    data.tankSound || '',
    data.compartment || '',
    data.outletT || '',
    data.material || '',
    data.trapLocation || '',
    data.directions || '',
    data.schedule || '',
    data.lpo || '',
    data.nextPump || '',
    data.specialNotes || '',
    JSON.stringify(data.lineItems || []),
    data.total || '',
    data.billing || '',
    data.terms || '',
    data.checkNum || '',
    data.paymentAmt || '',
    data.comments || '',
    data.signedBy || '',
    now.toISOString(),
    data.clientEmail || ''
  ]);
  return woNum;
}

function listOrders(secret) {
  if (secret !== SECRET) return { ok: false, error: 'Unauthorized' };
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, orders: [] };

  var headers = data[0];
  var orders  = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
    orders.push(obj);
    if (orders.length >= 100) break;
  }
  return { ok: true, orders: orders };
}

function lookupCustomer(phone, secret) {
  if (secret !== SECRET) return { ok: false, error: 'Unauthorized' };
  if (!phone) return { ok: false, error: 'No phone provided' };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, customer: null };

  var headers = data[0];
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    var p1  = String(row[idx['Phone 1']] || '').replace(/\D/g, '');
    var p2  = String(row[idx['Phone 2']] || '').replace(/\D/g, '');
    var q   = phone.replace(/\D/g, '');
    if (q && (p1 === q || p2 === q)) {
      return {
        ok: true,
        customer: {
          billTo:      row[idx['Bill To']],
          job:         row[idx['Job']],
          phone1:      row[idx['Phone 1']],
          phone2:      row[idx['Phone 2']],
          clientEmail: row[idx['Client Email']],
          material:    row[idx['Material']],
          compartment: row[idx['Compartment']],
          outletT:     row[idx['Outlet T']],
          tankSound:   row[idx['Tank Sound']],
          drainfield:  row[idx['Drainfield']],
          trapLocation:row[idx['Trap Location']],
          yearBuilt:   row[idx['Year Built']],
          pumpType:    row[idx['Pump Type']]
        }
      };
    }
  }
  return { ok: true, customer: null };
}

// --------------- Call-in helpers ---------------

function saveCallIn(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CALLIN_SHEET);
  var ciNum = nextCiNumber();
  var id    = Utilities.getUuid();
  var now   = new Date();

  sheet.appendRow([
    id,
    ciNum,
    data.rep || '',
    data.date || '',
    data.scheduledDate || '',
    data.scheduledTime || '',
    data.billToName || '',
    data.billToAddress || '',
    data.billToPhone || '',
    data.jobAddress || '',
    data.jobPhone || '',
    data.email || '',
    data.reason || '',
    data.reasonDetail || '',
    data.beds || '',
    data.baths || '',
    data.yearBuilt || '',
    data.lpo || '',
    data.liftStation ? 'Yes' : '',
    data.atu         ? 'Yes' : '',
    data.vacJob      ? 'Yes' : '',
    data.tankCount || '',
    data.directions || '',
    data.technicians || '',
    data.extraTime || '',
    data.price || '',
    data.overfull    ? 'Yes' : '',
    data.fieldRunback ? 'Yes' : '',
    data.paymentType || '',
    data.billingType || '',
    now.toISOString()
  ]);
  return ciNum;
}

function listCallIns(secret) {
  if (secret !== SECRET) return { ok: false, error: 'Unauthorized' };
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CALLIN_SHEET);
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, callins: [] };

  var headers = data[0];
  var callins = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
    callins.push(obj);
    if (callins.length >= 50) break;
  }
  return { ok: true, callins: callins };
}

// --------------- Number sequencing ---------------

function nextWoNumber() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var last  = parseInt(props.getProperty('LAST_WO') || '0', 10);
    var next  = last + 1;
    props.setProperty('LAST_WO', String(next));
    return 'WO-' + String(next).padStart(5, '0');
  } finally {
    lock.releaseLock();
  }
}

function nextCiNumber() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var last  = parseInt(props.getProperty('LAST_CI') || '0', 10);
    var next  = last + 1;
    props.setProperty('LAST_CI', String(next));
    return 'CI-' + String(next).padStart(5, '0');
  } finally {
    lock.releaseLock();
  }
}

// --------------- Utility ---------------

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
