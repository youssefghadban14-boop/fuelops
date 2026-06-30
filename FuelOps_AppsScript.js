// ═══════════════════════════════════════════════════════════════════════════════
// C.A.T. FuelOps — Google Apps Script Backend v2
// Deploy as Web App: Execute as "Me", Access "Anyone"
// ═══════════════════════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();

const SHEETS = {
  storage:     'Fixed Storage',
  tankers:     'Mobile Tankers',
  suppliers:   'Suppliers',
  fleet:       'Fleet Equipment',
  drivers:     'Drivers',
  users:       'Users & Passwords',
  projects:    'Projects',
  txReceive:   'Tx - Receive',
  txTransfer:  'Tx - Transfer',
  txDistribute:'Tx - Distribute',
  txAdjustment:'Tx - Adjustment',
  locations:   'Locations'
};

// ── ENTRY POINT ───────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const a = e.parameter.action;
    let result;
    if      (a === 'getData')        result = getData(e.parameter.role, e.parameter.userId);
    else if (a === 'getLoginUsers')  result = getLoginUsers();
    else if (a === 'addTx')          result = addTransaction(e.parameter);
    else if (a === 'saveStorage')    result = saveRecord('storage',  e.parameter);
    else if (a === 'saveTanker')     result = saveRecord('tankers',  e.parameter);
    else if (a === 'saveSupplier')   result = saveRecord('suppliers',e.parameter);
    else if (a === 'saveFleet')      result = saveRecord('fleet',    e.parameter);
    else if (a === 'saveDriver')     result = saveRecord('drivers',  e.parameter);
    else if (a === 'saveUser')       result = saveRecord('users',    e.parameter);
    else if (a === 'saveProject')    result = saveRecord('projects', e.parameter);
    else if (a === 'saveLocation')   result = saveRecord('locations', e.parameter);
    else if (a === 'deleteRecord')   result = deleteRecord(e.parameter.sheet, e.parameter.id);
    else if (a === 'saveAdjustment') result = saveAdjustment(e.parameter);
    else result = { error: 'Unknown action: ' + a };
    return out({ ok: true, data: result });
  } catch(err) {
    return out({ ok: false, error: err.message });
  }
}

function out(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ── SAVE ADJUSTMENT ────────────────────────────────────────────────────────────
function saveAdjustment(p) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = SS.getSheetByName('Tx - Adjustment');
  if (!sheet) {
    sheet = SS.insertSheet('Tx - Adjustment');
    sheet.appendRow(['ID','Date/Time','Tank/Tanker','Type','Location','System Stock (L)','Physical Stock (L)','Difference (L)','Reason','Notes','Adjusted By','User ID']);
  }
  // Apply the difference using the same reliable matching used by every other transaction (by Tank Nb./Fleet Nb.)
  const sheetKey = p.tankKind === 'tanker' ? 'tankers' : 'storage';
  adjustStock(sheetKey, p.tankName, +p.diff || 0);
  // Record adjustment transaction
  sheet.appendRow([p.id, p.date, p.tankName, p.tankKind, p.location||'', +p.sysStock||0, +p.physical||0, +p.diff||0, p.reason, p.notes||'', p.userName, p.userId, p.tankId||'']);
  return { ok: true };
}

// ── GET LOGIN USERS (for login dropdown) ──────────────────────────────────────
function getLoginUsers() {
  return readSheet('users').filter(u => u['Status'] !== 'Inactive');
}

// ── GET DATA (role-based) ─────────────────────────────────────────────────────
function getData(role, userId) {
  const r = {};
  if (role === 'admin') {
    r.storage    = readSheet('storage');
    r.tankers    = readSheet('tankers');
    r.suppliers  = readSheet('suppliers');
    r.fleet      = readSheet('fleet');
    r.drivers    = readSheet('drivers');
    r.users      = readSheet('users');
    r.projects   = readSheet('projects');
    r.locations  = readSheet('locations');
    r.txReceive    = readSheet('txReceive');
    r.txTransfer   = readSheet('txTransfer');
    r.txDistribute = readSheet('txDistribute');
  }
  else if (role === 'storekeeper') {
    // Find this user's assigned tank by matching user's "Assigned ID" to Tank Nb.
    const me = readSheet('users').find(u => u['ID'] === userId);
    const myTankNbs = me ? (me['Assigned ID']||'').split(',').map(s=>s.trim()).filter(Boolean) : [];
    r.storage    = readSheet('storage').filter(x => myTankNbs.includes(x['Tank Nb.']));
    r.tankers    = readSheet('tankers');
    r.suppliers  = readSheet('suppliers');
    r.fleet      = readSheet('fleet');
    r.drivers    = readSheet('drivers');
    r.projects   = readSheet('projects');
    r.locations  = readSheet('locations');
    // Return ALL transactions for their assigned tanks, not just their own
    const skTxR = readSheet('txReceive');
    const skTxT = readSheet('txTransfer');
    const skTxD = readSheet('txDistribute');
    const myTankNbs2 = myTankNbs;
    r.txReceive    = skTxR.filter(x => myTankNbs2.includes(x['Destination']) || x['User ID'] === userId);
    r.txTransfer   = skTxT.filter(x => myTankNbs2.some(nb=>x['From']&&x['From'].includes(nb)) || myTankNbs2.some(nb=>x['To']&&x['To'].includes(nb)) || x['User ID'] === userId);
    r.txDistribute = skTxD.filter(x => myTankNbs2.some(nb=>x['From']&&x['From'].includes(nb)) || x['User ID'] === userId);
  }
  else if (role === 'pmv_admin') {
    r.storage    = readSheet('storage');
    r.tankers    = readSheet('tankers');
    r.fleet      = readSheet('fleet');
    r.projects   = readSheet('projects');
    r.locations  = readSheet('locations');
    r.txReceive    = readSheet('txReceive');
    r.txTransfer   = readSheet('txTransfer');
    r.txDistribute = readSheet('txDistribute');
  }
  else if (role === 'tanker_driver') {
    // Find this user's assigned tanker by matching user's "Assigned ID" to Fleet Nb.
    const me = readSheet('users').find(u => u['ID'] === userId);
    const myFleetNb = me ? me['Assigned ID'] : '';
    r.tankers  = readSheet('tankers').filter(x => x['Fleet Nb.'] === myFleetNb);
    r.fleet    = readSheet('fleet');
    r.drivers  = readSheet('drivers');
    r.projects = readSheet('projects');
    const tdTxD = readSheet('txDistribute');
    r.txDistribute = tdTxD.filter(x => x['From'] === myFleetNb || x['User ID'] === userId);
  }
  return r;
}

// ── READ SHEET ────────────────────────────────────────────────────────────────
function readSheet(key) {
  const sheet = SS.getSheetByName(SHEETS[key]);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1)
    .filter(row => row.some(c => c !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { let v = row[i] !== undefined ? String(row[i]) : ''; if(h==='Status' && v) v=v.charAt(0).toUpperCase()+v.slice(1).toLowerCase(); obj[h] = v; });
      return obj;
    });
}

// ── ADD TRANSACTION ───────────────────────────────────────────────────────────
function addTransaction(p) {
  const type = p.type;
  let key, headers, row;

  if (type === 'receive') {
    key = 'txReceive';
    headers = ['ID','Date/Time','Supplier','Destination','Volume (L)','Unit Price','Invoice Ref','Notes','Recorded By','User ID'];
    row = [p.id, p.date, p.from, p.to, +p.volume||0, +p.price||0, p.ref||'', p.notes||'', p.userName, p.userId];
    adjustStock('storage', p.destId, +p.volume||0);
  }
  else if (type === 'transfer') {
    key = 'txTransfer';
    headers = ['ID','Date/Time','From','To','Volume (L)','Notes','Recorded By','User ID'];
    row = [p.id, p.date, p.from, p.to, +p.volume||0, p.notes||'', p.userName, p.userId];
    adjustStock('storage', p.fromId, -(+p.volume||0));
    if (p.destKind === 'tank') adjustStock('storage', p.destId, +p.volume||0);
    else adjustStock('tankers', p.destId, +p.volume||0);
  }
  else if (type === 'distribute') {
    key = 'txDistribute';
    headers = ['ID','Date/Time','From','To Equipment','Fleet No','Volume (L)','Meter Reading','Meter Type','Project','Driver Name','Notes','Recorded By','User ID'];
    row = [p.id, p.date, p.from, p.to, p.fleetNo||'', +p.volume||0, p.meter||'', p.meterType||'', p.project||'', p.driverName||'', p.notes||'', p.userName, p.userId, p.warning||''];
    if (p.srcKind === 'tank') adjustStock('storage', p.srcId, -(+p.volume||0));
    else adjustStock('tankers', p.srcId, -(+p.volume||0));
  }

  const sheet = getOrCreate(key);
  ensureHeaders(sheet, headers);
  sheet.appendRow(row);
  return { written: true };
}

// ── ADJUST STOCK ──────────────────────────────────────────────────────────────
function adjustStock(sheetKey, recordId, delta) {
  const sheet = SS.getSheetByName(SHEETS[sheetKey]);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const h = data[0].map(x => String(x).trim());
  const idCol = sheetKey === 'storage' ? 'Tank Nb.' : 'Fleet Nb.';
  const idIdx = h.indexOf(idCol);
  const stIdx = h.indexOf('Current Stock (L)');
  if (idIdx < 0 || stIdx < 0) return;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(recordId)) {
      const cur = parseFloat(data[i][stIdx]) || 0;
      sheet.getRange(i+1, stIdx+1).setValue(Math.max(0, cur + delta));
      break;
    }
  }
}

// ── SAVE RECORD ───────────────────────────────────────────────────────────────
function saveRecord(type, p) {
  const sheet = getOrCreate(type);
  const colDefs = {
    storage:   { headers: ['ID','Tank Nb.','Location','Capacity (L)','Current Stock (L)','Status'],
                 row: [p.id, p.nb, p.location||'', +p.capacity||0, +p.current||0, p.status||'Active'] },
    tankers:   { headers: ['ID','Fleet Nb.','Plate Nb.','Owner','Make','Model','Year','Location','Capacity (L)','Current Stock (L)','Status'],
                 row: [p.id, p.fleetNb, p.plateNb, p.owner||'', p.make||'', p.model||'', p.year||'', p.location||'', +p.capacity||0, +p.current||0, p.status||'Available'] },
    suppliers: { headers: ['ID','Name','Contact','Phone','Email','Order/Month (L)','Status'],
                 row: [p.id, p.name, p.contact||'', p.phone||'', p.email||'', +p.orderPerMonth||0, p.status||'Active'] },
    fleet:     { headers: ['ID','Type','Category','Fleet No','Make','Model','Capacity','Year','Tank Capacity (L)','Meter Type','Owner','Status'],
                 row: [p.id, p.ftype||'', p.category||'', p.fleetNo||'', p.make||'', p.model||'', p.capacity||'', p.year||'', +p.tankCap||0, p.meterType||'hourmeter', p.owner||'C.A.T.', p.status||'Active'] },
    drivers:   { headers: ['ID','Employee ID','Residency ID','Name','Phone','License Type','Status'],
                 row: [p.id, p.empId||'', p.residencyId||'', p.name, p.phone||'', p.license||'', p.status||'Active'] },
    users:     { headers: ['ID','Employee ID','Residency ID','Full Name','Role','Password','Assigned ID','Status'],
                 row: [p.id, p.empId||'', p.residencyId||'', p.name, p.role, p.password, p.assignedId||'', p.status||'Active'] },
    projects:  { headers: ['ID','Project Code','Project Name','Status'],
                 row: [p.id, p.code||'', p.name, p.status||'Active'] },
    locations: { headers: ['ID','Location Name','Status'],
                 row: [p.id, p.name, p.status||'Active'] }
  };
  const def = colDefs[type];
  ensureHeaders(sheet, def.headers);
  const data = sheet.getDataRange().getValues();
  const h = data[0].map(x => String(x).trim());
  const idIdx = h.indexOf('ID');
  if (idIdx >= 0 && p.id) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(p.id)) {
        sheet.getRange(i+1, 1, 1, def.row.length).setValues([def.row]);
        return { updated: true };
      }
    }
  }
  sheet.appendRow(def.row);
  return { inserted: true };
}

// ── DELETE (soft) ─────────────────────────────────────────────────────────────
function deleteRecord(sheetKey, recordId) {
  const sheet = SS.getSheetByName(SHEETS[sheetKey]);
  if (!sheet) return { error: 'Sheet not found' };
  const data = sheet.getDataRange().getValues();
  const h = data[0].map(x => String(x).trim());
  const idIdx = h.indexOf('ID');
  const stIdx = h.indexOf('Status');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(recordId)) {
      if (stIdx >= 0) sheet.getRange(i+1, stIdx+1).setValue('Inactive');
      return { deleted: true };
    }
  }
  return { error: 'Not found' };
}

// ── HELPERS ───────────────────────────────────────────────────────────────────
function getOrCreate(key) {
  const name = SHEETS[key];
  return SS.getSheetByName(name) || SS.insertSheet(name);
}
function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0 || sheet.getRange(1,1).getValue() === '') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#c31c2a').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
}

// ── SETUP ─────────────────────────────────────────────────────────────────────
function setupSheets() {
  const defs = {
    'Fixed Storage':    ['ID','Tank Nb.','Location','Capacity (L)','Current Stock (L)','Status'],
    'Mobile Tankers':   ['ID','Fleet Nb.','Plate Nb.','Owner','Make','Model','Year','Location','Capacity (L)','Current Stock (L)','Status'],
    'Suppliers':        ['ID','Name','Contact','Phone','Email','Order/Month (L)','Status'],
    'Fleet Equipment':  ['ID','Type','Category','Fleet No','Make','Model','Capacity','Year','Tank Capacity (L)','Meter Type','Owner','Status'],
    'Drivers':          ['ID','Employee ID','Residency ID','Name','Phone','License Type','Status'],
    'Users & Passwords':['ID','Employee ID','Residency ID','Full Name','Role','Password','Assigned ID','Status'],
    'Projects':         ['ID','Project Code','Project Name','Status'],
    'Locations':        ['ID','Location Name','Status'],
    'Tx - Receive':     ['ID','Date/Time','Supplier','Destination','Volume (L)','Unit Price','Invoice Ref','Notes','Recorded By','User ID'],
    'Tx - Transfer':    ['ID','Date/Time','From','To','Volume (L)','Notes','Recorded By','User ID'],
    'Tx - Distribute':  ['ID','Date/Time','From','To Equipment','Fleet No','Volume (L)','Meter Reading','Meter Type','Project','Driver Name','Notes','Recorded By','User ID','Warning'],
    'Tx - Adjustment':  ['ID','Date/Time','Tank/Tanker','Type','Location','System Stock (L)','Physical Stock (L)','Difference (L)','Reason','Notes','Adjusted By','User ID','Tank ID']
  };
  Object.entries(defs).forEach(([name, cols]) => {
    let sheet = SS.getSheetByName(name) || SS.insertSheet(name);
    if (sheet.getLastRow() === 0 || sheet.getRange(1,1).getValue() === '') {
      sheet.getRange(1,1,1,cols.length).setValues([cols]);
      sheet.getRange(1,1,1,cols.length).setFontWeight('bold').setBackground('#c31c2a').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, cols.length);
    }
  });
  // Default admin
  const us = SS.getSheetByName('Users & Passwords');
  if (us.getLastRow() <= 1) {
    us.appendRow(['u_admin','','','Admin','admin','admin123','','Active']);
  }
  // Seed default locations if empty
  const locSheet = SS.getSheetByName('Locations');
  if(locSheet && locSheet.getLastRow()<=1){
    const locs=[['loc_1','ABQIQ','Active'],['loc_2','HALFMOON','Active'],['loc_3','HARADH','Active'],['loc_4','HOUTA','Active'],['loc_5','SHAYBAH','Active'],['loc_6','WASIT','Active']];
    locs.forEach(l=>locSheet.appendRow(l));
  }
  SpreadsheetApp.getUi().alert('✅ C.A.T. FuelOps setup complete! All 11 sheets created.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('FuelOps').addItem('Setup all sheets','setupSheets').addToUi();
}
