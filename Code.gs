// ============================================================
// S.O.S. Septic — Digital Work Order System
// Apps Script backend: serves the app and handles data I/O
// ============================================================

var SECRET = PropertiesService.getScriptProperties().getProperty('SECRET') || '';
var SHEET_NAME = 'WorkOrders';

// --------------- HTTP entry points ---------------

function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var action = params.action || '';

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
    var woNum = saveOrder(payload);
    return jsonResponse({ ok: true, woNumber: woNum });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// --------------- Data helpers ---------------

function saveOrder(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var woNum = nextWoNumber();
  var id = Utilities.getUuid();
  var now = new Date();

  var row = [
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
    now.toISOString()
  ];

  sheet.appendRow(row);
  return woNum;
}

function listOrders(secret) {
  if (secret !== SECRET) return { ok: false, error: 'Unauthorized' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, orders: [] };

  var headers = data[0];
  var orders = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    orders.push(obj);
    if (orders.length >= 100) break; // cap at 100 most recent
  }
  return { ok: true, orders: orders };
}

function lookupCustomer(phone, secret) {
  if (secret !== SECRET) return { ok: false, error: 'Unauthorized' };
  if (!phone) return { ok: false, error: 'No phone provided' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, customer: null };

  var headers = data[0];
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  // Search newest-first
  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    var p1 = String(row[idx['Phone 1']] || '').replace(/\D/g, '');
    var p2 = String(row[idx['Phone 2']] || '').replace(/\D/g, '');
    var q  = phone.replace(/\D/g, '');
    if (q && (p1 === q || p2 === q)) {
      return {
        ok: true,
        customer: {
          billTo:      row[idx['Bill To']],
          job:         row[idx['Job']],
          phone1:      row[idx['Phone 1']],
          phone2:      row[idx['Phone 2']],
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

// --------------- WO number sequencing ---------------

function nextWoNumber() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var last = parseInt(props.getProperty('LAST_WO') || '0', 10);
    var next = last + 1;
    props.setProperty('LAST_WO', String(next));
    return 'WO-' + String(next).padStart(5, '0');
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
