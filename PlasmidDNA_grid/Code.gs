/*********************************
 * PlasmidDNA_Grid — Standalone-safe web app
 *
 * Spreadsheet:
 * https://docs.google.com/spreadsheets/your_link_of_interest
 *
 * Sheets:
 * - BOXES
 *   Columns: BoxID, BOX_UID, BoxType, BoxSampleType
 *
 * - PLASMID_DNA
 *   Columns:
 *   Plasmid Name
 *   Gene_HUGO
 *   Addgene
 *   Antibiotics_Resist
 *   Sequence_Verified
 *   BoxID
 *   Position
 *   Position_Key
 *   Plasmid_DNA_UID
 *   ConstructedBy
 *   Stored_Date
 *   Link
 *   Notes
 *
 * PLASMID_DNA[BoxID] stores BOX_UID in raw Google Sheet.
 *
 * Only allows boxes where:
 * BOXES[BoxSampleType] == "Plasmid DNA"
 *
 * Web app URL parameters:
 *   ?boxUid=<BOX_UID>
 *   ?boxId=<human BoxID>
 *   ?pos=A01
 *   ?itemKey=<Plasmid_DNA_UID>
 *
 * UI actions:
 * - place unplaced plasmid DNA into empty well
 * - move plasmid DNA record to another empty well
 * - delete selected plasmid DNA rows
 *
 * Auto-refresh: 2 seconds
 *********************************/


/************ CONFIG ************/
const CONFIG = {
  APP_NAME: 'PlasmidDNA_Grid',

  SPREADSHEET_ID: 'your_link_of_interest',

  BOXES_SHEET: 'BOXES',
  DNA_SHEET: 'PLASMID_DNA',
  ALLOWED_USERS_SHEET: 'GRID_ALLOWED_USERS',

  REQUIRE_ALLOWLIST_FOR_READ: true,

  // Must exactly match BOXES[BoxSampleType]
  REQUIRED_BOXSAMPLETYPE: 'Plasmid DNA',

  BOX_COLS: {
    BOXID: 'BoxID',
    BOXUID: 'BOX_UID',
    BOXTYPE: 'BoxType',
    BOXSAMPLETYPE: 'BoxSampleType'
  },

  DNA_COLS: {
    KEY: 'Plasmid_DNA_UID',
    LABEL: 'Plasmid Name',
    BOXID: 'BoxID',
    POSITION: 'Position',
    POSITION_KEY: 'Position_Key'
  },

  BOXTYPE_SPECS: {
    '9x9': {
      rows: 'ABCDEFGHI'.split(''),
      cols: 9
    },
    '8x12': {
      rows: 'ABCDEFGH'.split(''),
      cols: 12
    }
  },

  CACHE_TTL_SECONDS: 60
};


/************ OPTIONAL SETUP ************/
function setSpreadsheetId_() {
  PropertiesService.getScriptProperties().setProperty(
    'SPREADSHEET_ID',
    CONFIG.SPREADSHEET_ID
  );
}


/************ SPREADSHEET ACCESS ************/
function getSS_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();

  if (active) return active;

  const propId = PropertiesService
    .getScriptProperties()
    .getProperty('SPREADSHEET_ID');

  const id = propId || CONFIG.SPREADSHEET_ID;

  if (!id) {
    throw new Error(
      'SPREADSHEET_ID not set. Set CONFIG.SPREADSHEET_ID or run setSpreadsheetId_().'
    );
  }

  return SpreadsheetApp.openById(id);
}


function readSheet_(name) {
  const ss = getSS_();
  const sh = ss.getSheetByName(name);

  if (!sh) {
    throw new Error(`Sheet not found: ${name}`);
  }

  const values = sh.getDataRange().getValues();
  const header = values[0] || [];

  return {
    sh,
    header,
    values
  };
}


function cleanHeader_(header) {
  return header.map(h => String(h || '').trim());
}


/************ ALLOWLIST ************/
function getCallerEmail_() {
  return String(Session.getActiveUser().getEmail() || '')
    .toLowerCase()
    .trim();
}


function truthy_(v) {
  const s = String(v ?? '')
    .toLowerCase()
    .trim();

  return s === 'true' ||
         s === 'yes' ||
         s === 'y' ||
         s === '1';
}


function isAllowedUser_() {
  const email = getCallerEmail_();

  if (!email) return false;

  const { values } = readSheet_(CONFIG.ALLOWED_USERS_SHEET);

  if (!values || values.length < 2) return false;

  const header = cleanHeader_(values[0]);

  const emailIdx = header.indexOf('Email');
  const activeIdx = header.indexOf('ACTIVE');

  if (emailIdx < 0 || activeIdx < 0) {
    throw new Error(
      `Missing required columns in ${CONFIG.ALLOWED_USERS_SHEET}: Email, ACTIVE`
    );
  }

  for (let r = 1; r < values.length; r++) {
    const rowEmail = String(values[r][emailIdx] || '')
      .toLowerCase()
      .trim();

    if (!rowEmail) continue;

    if (rowEmail === email) {
      return truthy_(values[r][activeIdx]);
    }
  }

  return false;
}


function assertAllowedUser_(opName) {
  if (!isAllowedUser_()) {
    const email = getCallerEmail_() || '(email unavailable)';
    throw new Error(`Not authorized for ${opName || 'operation'}: ${email}`);
  }
}


function assertAllowedUserForReadIfEnabled_() {
  if (!CONFIG.REQUIRE_ALLOWLIST_FOR_READ) return;

  assertAllowedUser_('read');
}


/************ BOX MAPS ************/
function buildBoxMapByUid_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('BOX_MAP_UID');

  if (cached) {
    return JSON.parse(cached);
  }

  const { values } = readSheet_(CONFIG.BOXES_SHEET);

  if (!values || values.length < 2) {
    return {};
  }

  const header = cleanHeader_(values[0]);

  const boxIdIdx = header.indexOf(CONFIG.BOX_COLS.BOXID);
  const boxUidIdx = header.indexOf(CONFIG.BOX_COLS.BOXUID);
  const boxTypeIdx = header.indexOf(CONFIG.BOX_COLS.BOXTYPE);
  const boxSampleIdx = header.indexOf(CONFIG.BOX_COLS.BOXSAMPLETYPE);

  if (
    boxIdIdx < 0 ||
    boxUidIdx < 0 ||
    boxTypeIdx < 0 ||
    boxSampleIdx < 0
  ) {
    throw new Error(
      'BOXES must have columns: BoxID, BOX_UID, BoxType, BoxSampleType'
    );
  }

  const m = {};

  for (let r = 1; r < values.length; r++) {
    const boxId = String(values[r][boxIdIdx] || '').trim();
    const boxUid = String(values[r][boxUidIdx] || '').trim();
    const boxType = String(values[r][boxTypeIdx] || '').trim();
    const boxSampleType = String(values[r][boxSampleIdx] || '').trim();

    if (!boxUid || !boxType) continue;

    m[boxUid] = {
      boxId,
      boxType,
      boxSampleType
    };
  }

  cache.put(
    'BOX_MAP_UID',
    JSON.stringify(m),
    CONFIG.CACHE_TTL_SECONDS
  );

  return m;
}


function buildBoxMapById_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('BOX_MAP_ID');

  if (cached) {
    return JSON.parse(cached);
  }

  const uidMap = buildBoxMapByUid_();
  const m = {};

  for (const uid of Object.keys(uidMap)) {
    const b = uidMap[uid];

    if (b.boxId) {
      m[b.boxId] = uid;
    }
  }

  cache.put(
    'BOX_MAP_ID',
    JSON.stringify(m),
    CONFIG.CACHE_TTL_SECONDS
  );

  return m;
}


function lookupBoxUidByBoxId_(boxId) {
  const m = buildBoxMapById_();

  return m[String(boxId || '').trim()] || '';
}


function getBoxInfoByUid_(boxUid) {
  const uid = String(boxUid || '').trim();

  const m = buildBoxMapByUid_();
  const info = m[uid];

  if (!info) {
    throw new Error(`BOX_UID not found in BOXES: ${boxUid}`);
  }

  return {
    boxUid: uid,
    boxId: info.boxId,
    boxType: info.boxType,
    boxSampleType: info.boxSampleType
  };
}


function assertCorrectSampleType_(boxUid) {
  const b = getBoxInfoByUid_(boxUid);
  const actual = String(b.boxSampleType || '').trim();

  if (actual !== CONFIG.REQUIRED_BOXSAMPLETYPE) {
    throw new Error(
      `Restricted to "${CONFIG.REQUIRED_BOXSAMPLETYPE}" boxes. BoxSampleType="${actual}".`
    );
  }
}


/************ WEB APP ENTRY ************/
function doGet(e) {
  assertAllowedUserForReadIfEnabled_();

  const rawBoxUid = String(e?.parameter?.boxUid || '').trim();
  const rawBoxId = String(e?.parameter?.boxId || '').trim();
  const selectedPos = String(e?.parameter?.pos || '')
    .trim()
    .toUpperCase();

  const itemKey = String(e?.parameter?.itemKey || '').trim();

  let boxUid = rawBoxUid;

  if (!boxUid && rawBoxId) {
    boxUid = lookupBoxUidByBoxId_(rawBoxId);
  }

  if (!boxUid) {
    const tpl = HtmlService.createTemplateFromFile('Index');

    tpl.appName = CONFIG.APP_NAME;
    tpl.mode = 'browser';
    tpl.boxUid = '';
    tpl.boxId = '';
    tpl.boxType = '';
    tpl.selectedPos = selectedPos || '';
    tpl.itemKey = itemKey || '';

    return tpl.evaluate()
      .setTitle(CONFIG.APP_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  assertCorrectSampleType_(boxUid);

  const boxInfo = getBoxInfoByUid_(boxUid);

  const tpl = HtmlService.createTemplateFromFile('Index');

  tpl.appName = CONFIG.APP_NAME;
  tpl.mode = 'grid';
  tpl.boxUid = boxUid;
  tpl.boxId = boxInfo.boxId;
  tpl.boxType = boxInfo.boxType;
  tpl.selectedPos = selectedPos || '';
  tpl.itemKey = itemKey || '';

  return tpl.evaluate()
    .setTitle(`${CONFIG.APP_NAME}: ${boxInfo.boxId}`)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/************ API: LIST BOXES ************/
function apiListBoxes() {
  assertAllowedUserForReadIfEnabled_();

  const map = buildBoxMapByUid_();
  const out = [];

  for (const uid of Object.keys(map)) {
    const b = map[uid];

    if (
      String(b.boxSampleType || '').trim() !==
      CONFIG.REQUIRED_BOXSAMPLETYPE
    ) {
      continue;
    }

    out.push({
      boxUid: uid,
      boxId: b.boxId,
      boxType: b.boxType
    });
  }

  out.sort((a, b) =>
    String(a.boxId).localeCompare(String(b.boxId))
  );

  return out;
}


/************ API: GRID STATE ************/
function apiGetGridState(boxUid) {
  assertAllowedUserForReadIfEnabled_();
  assertCorrectSampleType_(boxUid);

  const boxInfo = getBoxInfoByUid_(boxUid);
  const spec = CONFIG.BOXTYPE_SPECS[String(boxInfo.boxType || '').trim()];

  if (!spec) {
    throw new Error(`Unsupported BoxType: ${boxInfo.boxType}`);
  }

  const cacheKey = `OCC:${boxUid}`;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const occ = getOccupancy_(boxUid);
  const labels = generateLabels_(spec.rows, spec.cols);

  const cells = labels.map(label => ({
    label,
    occupied: occ.has(label),
    itemKey: occ.has(label) ? occ.get(label).key : '',
    labelText: occ.has(label) ? occ.get(label).label : ''
  }));

  const state = {
    boxUid,
    boxId: boxInfo.boxId,
    boxType: boxInfo.boxType,
    cells
  };

  cache.put(
    cacheKey,
    JSON.stringify(state),
    CONFIG.CACHE_TTL_SECONDS
  );

  return state;
}


/************ API: UNPLACED ITEMS ************/
function apiListUnplacedItems(boxUid) {
  assertAllowedUserForReadIfEnabled_();
  assertCorrectSampleType_(boxUid);

  const { header, values } = readSheet_(CONFIG.DNA_SHEET);
  const cleanHeader = cleanHeader_(header);

  const boxIdx = cleanHeader.indexOf(CONFIG.DNA_COLS.BOXID);
  const posIdx = cleanHeader.indexOf(CONFIG.DNA_COLS.POSITION);
  const keyIdx = cleanHeader.indexOf(CONFIG.DNA_COLS.KEY);
  const labelIdx = cleanHeader.indexOf(CONFIG.DNA_COLS.LABEL);

  if (
    boxIdx < 0 ||
    posIdx < 0 ||
    keyIdx < 0 ||
    labelIdx < 0
  ) {
    throw new Error(
      'PLASMID_DNA missing required columns: BoxID, Position, Plasmid_DNA_UID, Plasmid Name'
    );
  }

  const out = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    if (
      String(row[boxIdx] || '').trim() !==
      String(boxUid || '').trim()
    ) {
      continue;
    }

    const pos = String(row[posIdx] || '').trim();

    if (pos) continue;

    const key = String(row[keyIdx] || '').trim();

    if (!key) continue;

    const label = String(row[labelIdx] || '').trim() || key;

    out.push({
      key,
      label
    });
  }

  out.sort((a, b) =>
    String(a.label).localeCompare(String(b.label))
  );

  return out;
}


/************ API: PLACE ITEM ************/
function apiPlaceItem(boxUid, itemKey, targetPos) {
  assertAllowedUser_('place');
  assertCorrectSampleType_(boxUid);

  if (!itemKey) {
    throw new Error('Missing Plasmid_DNA_UID');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const boxInfo = getBoxInfoByUid_(boxUid);

    validatePos_(boxInfo.boxType, targetPos);

    const pos = normPos_(targetPos);
    const occ = getOccupancy_(boxUid);

    if (occ.has(pos)) {
      throw new Error(`Target ${pos} occupied`);
    }

    const rowIndex = findRowIndexByKey_(
      CONFIG.DNA_SHEET,
      CONFIG.DNA_COLS.KEY,
      itemKey
    );

    if (!rowIndex) {
      throw new Error(`Plasmid_DNA_UID not found: ${itemKey}`);
    }

    const { sh, header } = readSheet_(CONFIG.DNA_SHEET);
    const cleanHeader = cleanHeader_(header);

    const boxIdx = cleanHeader.indexOf(CONFIG.DNA_COLS.BOXID);
    const posIdx = cleanHeader.indexOf(CONFIG.DNA_COLS.POSITION);
    const pkIdx = cleanHeader.indexOf(CONFIG.DNA_COLS.POSITION_KEY);

    if (boxIdx < 0 || posIdx < 0 || pkIdx < 0) {
      throw new Error(
        'PLASMID_DNA missing required columns: BoxID, Position, Position_Key'
      );
    }

    const rowVals = sh
      .getRange(rowIndex, 1, 1, sh.getLastColumn())
      .getValues()[0];

    if (
      String(rowVals[boxIdx] || '').trim() !==
      String(boxUid || '').trim()
    ) {
      throw new Error('Record belongs to a different box.');
    }

    if (String(rowVals[posIdx] || '').trim()) {
      throw new Error('Already placed.');
    }

    sh.getRange(rowIndex, posIdx + 1).setValue(pos);
    sh.getRange(rowIndex, pkIdx + 1).setValue(`${boxInfo.boxId}_${pos}`);

    invalidateBoxCache_(boxUid);

    return {
      ok: true
    };
  } finally {
    lock.releaseLock();
  }
}


/************ API: MOVE ITEM ************/
function apiMoveItem(boxUid, fromPos, toPos) {
  assertAllowedUser_('move');
  assertCorrectSampleType_(boxUid);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const boxInfo = getBoxInfoByUid_(boxUid);

    validatePos_(boxInfo.boxType, fromPos);
    validatePos_(boxInfo.boxType, toPos);

    const fromP = normPos_(fromPos);
    const toP = normPos_(toPos);

    const occ = getOccupancy_(boxUid);

    if (!occ.has(fromP)) {
      throw new Error(`Source ${fromP} empty`);
    }

    if (occ.has(toP)) {
      throw new Error(`Target ${toP} occupied`);
    }

    const rowIndex = occ.get(fromP).rowIndex;

    const { sh, header } = readSheet_(CONFIG.DNA_SHEET);
    const cleanHeader = cleanHeader_(header);

    const posIdx = cleanHeader.indexOf(CONFIG.DNA_COLS.POSITION);
    const pkIdx = cleanHeader.indexOf(CONFIG.DNA_COLS.POSITION_KEY);

    if (posIdx < 0 || pkIdx < 0) {
      throw new Error(
        'PLASMID_DNA missing required columns: Position, Position_Key'
      );
    }

    sh.getRange(rowIndex, posIdx + 1).setValue(toP);
    sh.getRange(rowIndex, pkIdx + 1).setValue(`${boxInfo.boxId}_${toP}`);

    invalidateBoxCache_(boxUid);

    return {
      ok: true
    };
  } finally {
    lock.releaseLock();
  }
}


/************ API: DELETE ROWS ************/
function apiDeleteRowsByPositions(boxUid, positions) {
  assertAllowedUser_('delete');
  assertCorrectSampleType_(boxUid);

  if (!Array.isArray(positions) || positions.length === 0) {
    return {
      ok: true,
      deleted: 0
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const boxInfo = getBoxInfoByUid_(boxUid);
    const occ = getOccupancy_(boxUid);

    const toDelete = [];

    for (const p0 of positions) {
      validatePos_(boxInfo.boxType, p0);

      const p = normPos_(p0);

      if (!occ.has(p)) continue;

      toDelete.push(occ.get(p).rowIndex);
    }

    if (!toDelete.length) {
      return {
        ok: true,
        deleted: 0
      };
    }

    toDelete.sort((a, b) => b - a);

    const ss = getSS_();
    const sh = ss.getSheetByName(CONFIG.DNA_SHEET);

    if (!sh) {
      throw new Error(`Sheet not found: ${CONFIG.DNA_SHEET}`);
    }

    for (const r of toDelete) {
      sh.deleteRow(r);
    }

    invalidateBoxCache_(boxUid);

    return {
      ok: true,
      deleted: toDelete.length
    };
  } finally {
    lock.releaseLock();
  }
}


/************ OCCUPANCY ************/
function getOccupancy_(boxUid) {
  const ss = getSS_();
  const sh = ss.getSheetByName(CONFIG.DNA_SHEET);

  if (!sh) {
    throw new Error(`Sheet not found: ${CONFIG.DNA_SHEET}`);
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 2) {
    return new Map();
  }

  const header = sh
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(h => String(h || '').trim());

  const boxIdx = header.indexOf(CONFIG.DNA_COLS.BOXID);
  const posIdx = header.indexOf(CONFIG.DNA_COLS.POSITION);
  const keyIdx = header.indexOf(CONFIG.DNA_COLS.KEY);
  const labelIdx = header.indexOf(CONFIG.DNA_COLS.LABEL);

  if (
    boxIdx < 0 ||
    posIdx < 0 ||
    keyIdx < 0 ||
    labelIdx < 0
  ) {
    throw new Error(
      'PLASMID_DNA missing required columns: BoxID, Position, Plasmid_DNA_UID, Plasmid Name'
    );
  }

  const values = sh
    .getRange(2, 1, lastRow - 1, lastCol)
    .getValues();

  const occ = new Map();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    if (
      String(row[boxIdx] || '').trim() !==
      String(boxUid || '').trim()
    ) {
      continue;
    }

    const pos = String(row[posIdx] || '')
      .trim()
      .toUpperCase();

    if (!pos || pos.length !== 3) continue;

    const key = String(row[keyIdx] || '').trim();

    if (!key) continue;

    const label = String(row[labelIdx] || '').trim() || key;

    if (!occ.has(pos)) {
      occ.set(pos, {
        key,
        label,
        rowIndex: i + 2
      });
    }
  }

  return occ;
}


/************ HELPERS ************/
function invalidateBoxCache_(boxUid) {
  CacheService
    .getScriptCache()
    .remove(`OCC:${boxUid}`);
}


function findRowIndexByKey_(sheetName, keyColName, keyValue) {
  const ss = getSS_();
  const sh = ss.getSheetByName(sheetName);

  if (!sh) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 2) {
    return 0;
  }

  const header = sh
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(h => String(h || '').trim());

  const keyIdx = header.indexOf(keyColName);

  if (keyIdx < 0) {
    throw new Error(`Missing column: ${keyColName}`);
  }

  const keys = sh
    .getRange(2, keyIdx + 1, lastRow - 1, 1)
    .getValues();

  for (let i = 0; i < keys.length; i++) {
    if (
      String(keys[i][0] || '').trim() ===
      String(keyValue || '').trim()
    ) {
      return i + 2;
    }
  }

  return 0;
}


function normPos_(pos) {
  return String(pos || '')
    .trim()
    .toUpperCase();
}


function validatePos_(boxType, pos) {
  const spec = CONFIG.BOXTYPE_SPECS[String(boxType || '').trim()];

  if (!spec) {
    throw new Error(`Unsupported BoxType: ${boxType}`);
  }

  const p = normPos_(pos);

  if (p.length !== 3) {
    throw new Error(`Invalid position format: ${p}. Expected A01.`);
  }

  const row = p.substring(0, 1);
  const col = p.substring(1, 3);

  if (spec.rows.indexOf(row) < 0) {
    throw new Error(`Invalid row ${row} for box type ${boxType}`);
  }

  const colNum = Number(col);

  if (!colNum || colNum < 1 || colNum > spec.cols) {
    throw new Error(`Invalid column ${col} for box type ${boxType}`);
  }
}


function generateLabels_(rows, colsCount) {
  const labels = [];

  for (const r of rows) {
    for (let c = 1; c <= colsCount; c++) {
      labels.push(`${r}${String(c).padStart(2, '0')}`);
    }
  }

  return labels;
}
