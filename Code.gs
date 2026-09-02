// ============================================================
// S.O.S. Septic — Digital Work Order System
// Apps Script backend: serves the app and handles data I/O
// ============================================================

var SECRET = PropertiesService.getScriptProperties().getProperty('SECRET') || '';
var SHEET_NAME    = 'WorkOrders';
var CALLIN_SHEET  = 'CallIns';
var SALARY_SHEET  = 'SalaryResearch';

// WorkOrders tab columns (A1:AJ1, 36 headers):
//   ID | WO Number | Type | Date | Time | Bill To | Job | Phone 1 | Phone 2 |
//   Pumpout Ordered | Technician | Pump Type | Year Built | Water Level | Drainfield |
//   Scum | Sludge | Tank Sound | Compartment | Outlet T | Material | Trap Location |
//   Directions | Schedule | LPO | Next Pump | Special Notes | Line Items JSON | Total |
//   Billing | Terms | Check Num | Payment Amt | Comments | Signed By | Saved At

// CallIns tab columns (A1:AE1, 31 headers):
//   ID | CI Number | Rep | Date | Scheduled Date | Scheduled Time |
//   Bill To Name | Bill To Address | Bill To Phone |
//   Job Address | Job Phone | Email |
//   Reason | Reason Detail | Beds | Baths | Year Built | LPO |
//   Lift Station | ATU | Vac Job | Tank Count | Directions |
//   Technicians | Extra Time | Price | Overfull | Field Runback |
//   Payment Type | Billing Type | Saved At

// --------------- HTTP entry points ---------------

function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var action = params.action || '';
  var sheet  = params.sheet  || '';

  if (action === 'list' && sheet === 'callins') {
    return jsonResponse(listCallIns(params.secret));
  }
  if (action === 'list' && sheet === 'salary') {
    return jsonResponse(listSalaryEntries(params.secret));
  }
  if (action === 'list') {
    return jsonResponse(listOrders(params.secret));
  }
  if (action === 'lookup') {
    return jsonResponse(lookupCustomer(params.phone, params.secret));
  }
  if (action === 'deleteSalary') {
    return jsonResponse(deleteSalaryEntry(params.id, params.secret));
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
    if (payload.sheetType === 'salary') {
      var salId = saveSalaryEntry(payload);
      return jsonResponse({ ok: true, id: salId });
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
    now.toISOString()
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

// --------------- Salary seed (run once from Apps Script editor) ---------------

function seedSalaryData() {
  var entries = [
    // ── My log (user-provided salary + URL) ──
    {
      title:      'Director of Marketing',
      dateLogged: '2026-09-02',
      location:   'Chicago, IL',
      payMin:     135000,
      payMax:     165000,
      url:        'https://jobs.greystar.com/job/chicago/director-of-marketing/35302/92168771344',
      notes:      'DOE',
      source:     'My log',
      status:     'Expired'
    },
    {
      title:      'Director, Data Engineering',
      dateLogged: '2026-09-02',
      location:   'Remote (United States)',
      payMin:     160000,
      payMax:     195000,
      url:        'https://jobs.greystar.com/job/united-states/director-data-engineering/35302/88017300368',
      notes:      '',
      source:     'My log',
      status:     'Expired'
    },
    {
      title:      'Director of External Communications',
      dateLogged: '2026-09-02',
      location:   'Charleston, SC',
      payMin:     125000,
      payMax:     170000,
      url:        'https://jobs.greystar.com/job/charleston/director-external-communications/35302/94068008192',
      notes:      '',
      source:     'My log',
      status:     'Expired'
    },
    // ── Found via search (salary pulled from Google index) ──
    {
      title:      'Director, Procurement Governance & Systems – Operations',
      dateLogged: '2026-09-02',
      location:   'Remote (United States)',
      payMin:     100000,
      payMax:     140000,
      url:        'https://jobs.greystar.com/job/united-states/director-procurement-governance-and-systems-operations/35302/97796775952',
      notes:      'Salary from search index — verify on listing',
      source:     'Search',
      status:     'Active'
    },
    {
      title:      'Director, Web Platform Engineering',
      dateLogged: '2026-09-02',
      location:   'Remote (United States)',
      payMin:     150000,
      payMax:     170000,
      url:        'https://jobs.greystar.com/job/united-states/director-web-platform-engineering/35302/99366834224',
      notes:      'Listing no longer active as of 2026-09-02',
      source:     'Search',
      status:     'Expired'
    },
    {
      title:      'Director, Real Estate – New England Owned Assets',
      dateLogged: '2026-09-02',
      location:   'Boston, MA',
      payMin:     175000,
      payMax:     185000,
      url:        'https://jobs.greystar.com/job/boston/director-real-estate-new-england-owned-assets/35302/96969653872',
      notes:      'Salary from search index — verify on listing',
      source:     'Search',
      status:     'Active'
    },
    {
      title:      'Director, Real Estate',
      dateLogged: '2026-09-02',
      location:   'San Francisco, CA',
      payMin:     175000,
      payMax:     210000,
      url:        'https://jobs.greystar.com/job/san-francisco/director-real-estate/35302/97759545488',
      notes:      'Listing no longer active as of 2026-09-02',
      source:     'Search',
      status:     'Expired'
    },
    // ── My links (user-provided URL, salary from search index) ──
    {
      title:      'Director, Enablement & Engagement',
      dateLogged: '2026-09-02',
      location:   'Remote (United States)',
      payMin:     130000,
      payMax:     150000,
      url:        'https://jobs.greystar.com/job/united-states/director-enablement-and-engagement/35302/98712652128',
      notes:      'Salary from search index — verify on listing',
      source:     'My links',
      status:     'Active'
    },
    {
      title:      'Director, Corporate Compliance',
      dateLogged: '2026-09-02',
      location:   'Charleston, SC',
      payMin:     '',
      payMax:     '',
      url:        'https://jobs.greystar.com/job/charleston/director-corporate-compliance/35302/98049976848',
      notes:      'Salary not found in search — check listing',
      status:     'Active'
    },
    {
      title:      'Director of Client Services',
      dateLogged: '2026-09-02',
      location:   'Orlando, FL',
      payMin:     '',
      payMax:     '',
      url:        'https://jobs.greystar.com/job/orlando/director-of-client-services/35302/100021355728',
      notes:      'Salary not found in search — check listing',
      status:     'Active'
    },
  ];

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SALARY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SALARY_SHEET);
    sheet.appendRow(['ID','Date Logged','Title','Location','Pay Min','Pay Max','URL','Notes','Source','Status','Saved At']);
  }

  var now = new Date().toISOString();
  entries.forEach(function(e) {
    sheet.appendRow([
      Utilities.getUuid(),
      e.dateLogged,
      e.title,
      e.location,
      e.payMin,
      e.payMax,
      e.url,
      e.notes,
      e.source || 'My log',
      e.status,
      now
    ]);
  });

  Logger.log('Seeded ' + entries.length + ' salary entries.');
}

// --------------- Salary research helpers ---------------

function saveSalaryEntry(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SALARY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SALARY_SHEET);
    sheet.appendRow(['ID','Date Logged','Title','Location','Pay Min','Pay Max','URL','Notes','Source','Status','Saved At']);
  }
  var id  = Utilities.getUuid();
  var now = new Date();
  sheet.appendRow([
    id,
    data.dateLogged  || '',
    data.title       || '',
    data.location    || '',
    data.payMin      || '',
    data.payMax      || '',
    data.url         || '',
    data.notes       || '',
    data.source      || 'Manual',
    data.status      || 'Active',
    now.toISOString()
  ]);
  return id;
}

function listSalaryEntries(secret) {
  if (secret !== SECRET) return { ok: false, error: 'Unauthorized' };
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SALARY_SHEET);
  if (!sheet) return { ok: true, entries: [] };
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, entries: [] };

  var headers = data[0];
  var entries = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
    entries.push(obj);
  }
  return { ok: true, entries: entries };
}

function deleteSalaryEntry(id, secret) {
  if (secret !== SECRET) return { ok: false, error: 'Unauthorized' };
  if (!id) return { ok: false, error: 'No ID provided' };
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SALARY_SHEET);
  if (!sheet) return { ok: false, error: 'Sheet not found' };
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Entry not found' };
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
