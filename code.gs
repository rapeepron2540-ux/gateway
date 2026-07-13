const LINE_VERIFY_URL = 'https://api.line.me/oauth2/v2.1/verify';
var LAST_LINE_DEBUG_INFO_ = '';

// ====================================================================
// SHEET CONFIG
// ====================================================================
const ITEMS_SHEET = 'Items';
const USERS_SHEET = 'Users';
const TRANSACTIONS_SHEET = 'Transactions';

const ITEMS_HEADERS = [
  'ItemID', 'BaseCode', 'Name_TH', 'Name_EN', 'Category', 'Unit',
  'CurrentStock', 'ReorderPoint', 'Location', 'Shelf', 'LotNumber',
  'Price', 'ReceivedDate', 'MFG', 'EXP', 'CreatedAt'
];
const USERS_HEADERS = ['UID', 'Name', 'Role', 'RegisteredAt'];
const TRANSACTIONS_HEADERS = [
  'BatchID', 'UID', 'UserName', 'Timestamp', 'ItemID', 'BaseCode',
  'Name_TH', 'Name_EN', 'Qty', 'Unit', 'Status', 'CancelledAt'
];

const CANCEL_WINDOW_HOURS = 24; // ปรับได้ตามนโยบายจริง

// ====================================================================
// SHEET HELPERS
// ====================================================================
function getSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name, headers) {
  const ss = getSS_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

// แปลงข้อมูลทั้งชีทเป็น array of object พร้อมเก็บเลขแถวจริง (_row) ไว้ใช้ update
function sheetToObjects_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row[0] === '' || row[0] === null) continue; // ข้ามแถวว่าง
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx]; });
    obj._row = i + 1; // เลขแถวจริงในชีท (1-indexed)
    out.push(obj);
  }
  return out;
}

function stripRow_(o) {
  const c = Object.assign({}, o);
  delete c._row;
  return c;
}

function isWithinCancelWindow_(timestamp) {
  const diffMs = Date.now() - new Date(timestamp).getTime();
  return diffMs <= CANCEL_WINDOW_HOURS * 60 * 60 * 1000;
}

// ====================================================================
// ONE-TIME SETUP — รันเองใน Apps Script Editor ครั้งแรกเท่านั้น
// ====================================================================
function setupSheets() {
  getSheet_(ITEMS_SHEET, ITEMS_HEADERS);
  getSheet_(USERS_SHEET, USERS_HEADERS);
  getSheet_(TRANSACTIONS_SHEET, TRANSACTIONS_HEADERS);
  return 'สร้างชีทที่ขาดเรียบร้อยแล้ว (ชีทที่มีอยู่แล้วจะไม่ถูกแตะต้อง)';
}

// ตั้งให้ UID นี้เป็น Admin คนแรกของระบบ — รันเองใน Editor แล้วลบทิ้งได้
function setAdminRole(uid) {
  const sh = getSheet_(USERS_SHEET, USERS_HEADERS);
  const users = sheetToObjects_(sh);
  const existing = users.find(u => String(u.UID) === String(uid));
  const roleCol = USERS_HEADERS.indexOf('Role') + 1;
  if (existing) {
    sh.getRange(existing._row, roleCol).setValue('admin');
  } else {
    sh.appendRow([uid, 'Admin', 'admin', new Date()]);
  }
  return 'ตั้งค่า Admin ให้ UID: ' + uid + ' เรียบร้อยแล้ว';
}

// ====================================================================
// PUBLIC — เรียกจาก HTML ทุกหน้า
// ====================================================================
function getCategories() {
  return ['Office', 'Medical', 'Reagent'];
}

// ====================================================================
// USER MANAGEMENT
// ====================================================================
function getRegisteredUser_(uid) {
  const sh = getSheet_(USERS_SHEET, USERS_HEADERS);
  const users = sheetToObjects_(sh);
  return users.find(u => String(u.UID) === String(uid)) || null;
}

function requireAdmin_(uid) {
  const user = getRegisteredUser_(uid);
  if (!user || user.Role !== 'admin') {
    throw new Error('ไม่มีสิทธิ์ดำเนินการ (Admin เท่านั้น)');
  }
  return user;
}

function registerUser(uid, name) {
  if (!uid) throw new Error('UID ไม่ถูกต้อง');
  name = (name || '').trim();
  if (!name) throw new Error('กรุณากรอกชื่อ');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(USERS_SHEET, USERS_HEADERS);
    const users = sheetToObjects_(sh);
    const existing = users.find(u => String(u.UID) === String(uid));
    if (existing) {
      const nameCol = USERS_HEADERS.indexOf('Name') + 1;
      sh.getRange(existing._row, nameCol).setValue(name);
      return { UID: uid, Name: name, Role: existing.Role || 'user' };
    }
    sh.appendRow([uid, name, 'user', new Date()]);
    return { UID: uid, Name: name, Role: 'user' };
  } finally {
    lock.releaseLock();
  }
}

function unlinkUser(uid) {
  const sh = getSheet_(USERS_SHEET, USERS_HEADERS);
  const users = sheetToObjects_(sh);
  const existing = users.find(u => String(u.UID) === String(uid));
  if (existing) sh.deleteRow(existing._row);
  return { unlinked: true };
}

// ====================================================================
// ITEMS — ผู้ใช้ทั่วไป (รวมสต็อกทุกล็อตของ BaseCode เดียวกัน)
// ====================================================================
function getItemsForUser() {
  try {
    Logger.log('getItemsForUser: starting');
    const sh = getSheet_(ITEMS_SHEET, ITEMS_HEADERS);
    const items = sheetToObjects_(sh);
    Logger.log('getItemsForUser: loaded items count=' + (Array.isArray(items) ? items.length : 'not-array'));

    if (!Array.isArray(items)) {
      Logger.log('getItemsForUser: sheetToObjects_ returned non-array');
      return [];
    }

    const byBase = {};
    items.forEach(it => {
      if (!it || !it.BaseCode) {
        Logger.log('getItemsForUser: skipping invalid item row=' + JSON.stringify(it));
        return;
      }
      if (!byBase[it.BaseCode]) byBase[it.BaseCode] = [];
      byBase[it.BaseCode].push(it);
    });

    const result = Object.keys(byBase).map(baseCode => {
      const lots = byBase[baseCode].slice()
        .sort((a, b) => new Date(a.ReceivedDate) - new Date(b.ReceivedDate));
      const totalStock = lots.reduce((s, l) => s + (Number(l.CurrentStock) || 0), 0);
      const activeLot = lots.find(l => (Number(l.CurrentStock) || 0) > 0) || lots[0];
      return {
        ItemID: activeLot.ItemID,
        BaseCode: baseCode,
        Name_TH: activeLot.Name_TH,
        Name_EN: activeLot.Name_EN,
        Category: activeLot.Category,
        CurrentStock: totalStock,
        Unit: activeLot.Unit,
        Location: activeLot.Location,
        Shelf: activeLot.Shelf,
        ActiveLotLabel: activeLot.LotNumber,
        ActiveLotReceivedDate: activeLot.ReceivedDate
      };
    });

    Logger.log('getItemsForUser: returning result count=' + result.length);
    return result;
  } catch (err) {
    Logger.log('getItemsForUser: error=' + (err && err.message ? err.message : err));
    if (err && err.stack) {
      Logger.log(err.stack);
    }
    return [];
  }
}

// ====================================================================
// ITEMS — Admin CRUD (รายการดิบทีละล็อต)
// ====================================================================
function getItems() {
  const sh = getSheet_(ITEMS_SHEET, ITEMS_HEADERS);
  return sheetToObjects_(sh).map(stripRow_);
}

function getBaseItemsForLot() {
  const sh = getSheet_(ITEMS_SHEET, ITEMS_HEADERS);
  const items = sheetToObjects_(sh);
  const seen = {};
  const out = [];
  items.forEach(it => {
    if (!seen[it.BaseCode]) {
      seen[it.BaseCode] = true;
      out.push({ baseCode: it.BaseCode, nameTh: it.Name_TH, nameEn: it.Name_EN });
    }
  });
  return out;
}

// สร้าง Barcode อัตโนมัติ: 3 ตัวอักษรแรกของหมวดหมู่ + เลขรัน (เช่น OFF-004)
function generateBaseCode_(category, items) {
  const prefix = (category || 'GEN').substring(0, 3).toUpperCase();
  let max = 0;
  items.forEach(it => {
    if (it.BaseCode && it.BaseCode.indexOf(prefix + '-') === 0) {
      const num = parseInt(it.BaseCode.split('-')[1], 10);
      if (!isNaN(num) && num > max) max = num;
    }
  });
  return prefix + '-' + String(max + 1).padStart(3, '0');
}

function addItem(adminUid, payload) {
  requireAdmin_(adminUid);
  if (!payload.nameTh && !payload.nameEn) throw new Error('กรุณาระบุชื่อสินค้าอย่างน้อย 1 ภาษา');
  if (!payload.category) throw new Error('กรุณาเลือกหมวดหมู่');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(ITEMS_SHEET, ITEMS_HEADERS);
    const items = sheetToObjects_(sh);
    const baseCode = generateBaseCode_(payload.category, items);
    const itemId = baseCode + '-L1';
    sh.appendRow([
      itemId, baseCode, payload.nameTh || '', payload.nameEn || '', payload.category,
      payload.unit || 'Pcs.', Number(payload.currentStock) || 0, Number(payload.reorderPoint) || 0,
      payload.location || '', payload.shelf || '', payload.lotNumber || '', Number(payload.price) || 0,
      payload.receivedDate || '', payload.mfg || '', payload.exp || '', new Date()
    ]);
    return { itemId: itemId };
  } finally {
    lock.releaseLock();
  }
}

function updateItem(adminUid, itemId, fields) {
  requireAdmin_(adminUid);
  const sh = getSheet_(ITEMS_SHEET, ITEMS_HEADERS);
  const items = sheetToObjects_(sh);
  const target = items.find(i => i.ItemID === itemId);
  if (!target) throw new Error('ไม่พบสินค้านี้ในระบบ');

  const updates = {
    Name_TH: fields.Name_TH, Name_EN: fields.Name_EN, Category: fields.Category, Unit: fields.Unit,
    CurrentStock: Number(fields.CurrentStock) || 0, ReorderPoint: Number(fields.ReorderPoint) || 0,
    LotNumber: fields.LotNumber, Price: Number(fields.Price) || 0, ReceivedDate: fields.ReceivedDate,
    MFG: fields.MFG, EXP: fields.EXP, Location: fields.Location, Shelf: fields.Shelf
  };
  ITEMS_HEADERS.forEach((h, idx) => {
    if (Object.prototype.hasOwnProperty.call(updates, h) && updates[h] !== undefined) {
      sh.getRange(target._row, idx + 1).setValue(updates[h]);
    }
  });
  return { itemId: itemId };
}

function addNewLot(adminUid, baseCode, payload) {
  requireAdmin_(adminUid);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(ITEMS_SHEET, ITEMS_HEADERS);
    const items = sheetToObjects_(sh);
    const lots = items.filter(i => i.BaseCode === baseCode);
    if (lots.length === 0) throw new Error('ไม่พบสินค้าตั้งต้นนี้');
    const base = lots[0];

    let maxLot = 0;
    lots.forEach(l => {
      const m = String(l.ItemID).match(/-L(\d+)$/);
      if (m) { const n = parseInt(m[1], 10); if (n > maxLot) maxLot = n; }
    });
    const itemId = baseCode + '-L' + (maxLot + 1);

    sh.appendRow([
      itemId, baseCode, base.Name_TH, base.Name_EN, base.Category,
      base.Unit, Number(payload.currentStock) || 0, base.ReorderPoint,
      payload.location || base.Location || '', payload.shelf || base.Shelf || '',
      payload.lotNumber || '', Number(payload.price) || 0,
      payload.receivedDate || '', payload.mfg || '', payload.exp || '', new Date()
    ]);
    return { itemId: itemId };
  } finally {
    lock.releaseLock();
  }
}

// ====================================================================
// REQUISITION (เบิกของ) — FIFO ตัดสต็อกจากล็อตเก่าสุดก่อน
// ====================================================================
function submitRequisition(profile, cartItems) {
  if (!profile || !profile.uid) throw new Error('ไม่พบข้อมูลผู้ใช้ กรุณาโหลดหน้าใหม่');
  if (!Array.isArray(cartItems) || cartItems.length === 0) throw new Error('ตะกร้าว่างเปล่า');

  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(8000);
  if (!gotLock) throw new Error('ผู้ใช้งานพร้อมกันจำนวนมาก กรุณาลองอีกครั้ง');

  try {
    const sh = getSheet_(ITEMS_SHEET, ITEMS_HEADERS);
    const items = sheetToObjects_(sh);
    const txSh = getSheet_(TRANSACTIONS_SHEET, TRANSACTIONS_HEADERS);

    const batchId = Utilities.getUuid();
    const now = new Date();
    const stockUpdates = {}; // _row -> คงเหลือใหม่
    const txRows = [];

    cartItems.forEach(ci => {
      let remaining = Number(ci.qty) || 0;
      if (remaining <= 0) return;

      const lots = items.filter(it => it.BaseCode === ci.baseCode)
        .sort((a, b) => new Date(a.ReceivedDate) - new Date(b.ReceivedDate));

      const totalAvailable = lots.reduce((s, l) => {
        const avail = stockUpdates[l._row] !== undefined ? stockUpdates[l._row] : (Number(l.CurrentStock) || 0);
        return s + avail;
      }, 0);

      if (totalAvailable < remaining) {
        const label = lots[0] ? (lots[0].Name_TH || lots[0].Name_EN) : ci.baseCode;
        throw new Error('สินค้า ' + label + ' คงเหลือไม่พอเบิก');
      }

      for (let i = 0; i < lots.length && remaining > 0; i++) {
        const lot = lots[i];
        const currentAvail = stockUpdates[lot._row] !== undefined ? stockUpdates[lot._row] : (Number(lot.CurrentStock) || 0);
        if (currentAvail <= 0) continue;
        const take = Math.min(currentAvail, remaining);
        stockUpdates[lot._row] = currentAvail - take;
        remaining -= take;
        txRows.push([
          batchId, profile.uid, profile.name || '', now, lot.ItemID, lot.BaseCode,
          lot.Name_TH, lot.Name_EN, take, lot.Unit, 'Active', ''
        ]);
      }
    });

    const stockCol = ITEMS_HEADERS.indexOf('CurrentStock') + 1;
    Object.keys(stockUpdates).forEach(rowStr => {
      sh.getRange(Number(rowStr), stockCol).setValue(stockUpdates[rowStr]);
    });

    if (txRows.length > 0) {
      txSh.getRange(txSh.getLastRow() + 1, 1, txRows.length, TRANSACTIONS_HEADERS.length).setValues(txRows);
    }

    return { timestamp: now.getTime(), batchId: batchId };
  } finally {
    lock.releaseLock();
  }
}

// ====================================================================
// HISTORY
// ====================================================================
function getUserHistory(uid) {
  const sh = getSheet_(TRANSACTIONS_SHEET, TRANSACTIONS_HEADERS);
  const rows = sheetToObjects_(sh).filter(r => String(r.UID) === String(uid));

  const batches = {};
  rows.forEach(r => {
    if (!batches[r.BatchID]) {
      batches[r.BatchID] = {
        batchId: r.BatchID,
        timestamp: r.Timestamp,
        status: r.Status,
        cancelledAt: r.CancelledAt,
        items: []
      };
    }
    batches[r.BatchID].items.push({ nameTh: r.Name_TH, nameEn: r.Name_EN, qty: r.Qty, unit: r.Unit });
    if (r.Status === 'Cancelled') {
      batches[r.BatchID].status = 'Cancelled';
      batches[r.BatchID].cancelledAt = r.CancelledAt;
    }
  });

  const list = Object.keys(batches).map(k => {
    const b = batches[k];
    b.canCancel = b.status !== 'Cancelled' && isWithinCancelWindow_(b.timestamp);
    return b;
  });

  list.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  return list;
}

function cancelRequisition(uid, batchId) {
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(8000);
  if (!gotLock) throw new Error('ผู้ใช้งานพร้อมกันจำนวนมาก กรุณาลองอีกครั้ง');

  try {
    const txSh = getSheet_(TRANSACTIONS_SHEET, TRANSACTIONS_HEADERS);
    const txRows = sheetToObjects_(txSh).filter(r => r.BatchID === batchId && String(r.UID) === String(uid));
    if (txRows.length === 0) throw new Error('ไม่พบรายการนี้');
    if (txRows[0].Status === 'Cancelled') throw new Error('รายการนี้ถูกยกเลิกไปแล้ว');
    if (!isWithinCancelWindow_(txRows[0].Timestamp)) throw new Error('เกินระยะเวลาที่สามารถยกเลิกได้ (' + CANCEL_WINDOW_HOURS + ' ชม.)');

    const itemsSh = getSheet_(ITEMS_SHEET, ITEMS_HEADERS);
    const items = sheetToObjects_(itemsSh);
    const stockCol = ITEMS_HEADERS.indexOf('CurrentStock') + 1;
    const statusCol = TRANSACTIONS_HEADERS.indexOf('Status') + 1;
    const cancelledAtCol = TRANSACTIONS_HEADERS.indexOf('CancelledAt') + 1;
    const now = new Date();

    txRows.forEach(r => {
      const item = items.find(it => it.ItemID === r.ItemID);
      if (item) {
        itemsSh.getRange(item._row, stockCol).setValue((Number(item.CurrentStock) || 0) + Number(r.Qty));
      }
      txSh.getRange(r._row, statusCol).setValue('Cancelled');
      txSh.getRange(r._row, cancelledAtCol).setValue(now);
    });

    return { timestamp: now.getTime() };
  } finally {
    lock.releaseLock();
  }
}

// ====================================================================
// LINE LOGIN
// ====================================================================
function setLineChannelId(channelId) {
  PropertiesService.getScriptProperties().setProperty('LINE_CHANNEL_ID', channelId);
  return channelId;
}

function getLineChannelId_() {
  const id = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ID');
  if (!id) throw new Error('ยังไม่ได้ตั้งค่า LINE_CHANNEL_ID ใน Script Properties');
  return id;
}

function authorizeLineVerify() {
  const channelId = getLineChannelId_();
  const response = UrlFetchApp.fetch(LINE_VERIFY_URL, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: { id_token: 'dummy', client_id: channelId },
    muteHttpExceptions: true
  });

  Logger.log('authorizeLineVerify status: ' + response.getResponseCode());
  Logger.log(response.getContentText());

  return { status: response.getResponseCode(), body: response.getContentText() };
}

// ฟังก์ชันหลักที่ถูกเรียกเมื่อเปิด Web App URL ของ Apps Script
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'user';
  const idToken = e && e.parameter && e.parameter.idToken;

  if (!idToken) {
    return HtmlService.createHtmlOutput('<h3>สิทธิ์การเข้าใช้งานไม่ถูกต้อง (Missing Token) กรุณาเข้าใช้งานผ่าน LINE</h3>');
  }

  const lineUser = verifyLineIdToken_(idToken);
  if (!lineUser) {
    const debugMessage = LAST_LINE_DEBUG_INFO_ || 'Token หมดอายุหรือไม่ถูกต้อง';
    return HtmlService.createHtmlOutput([
      '<h3>การยืนยันตัวตนล้มเหลว</h3>',
      '<p>' + debugMessage + '</p>',
      '<p>ถ้าเป็นครั้งแรก กรุณาเปิด Apps Script Editor แล้วรันฟังก์ชัน <b>authorizeLineVerify</b> หนึ่งครั้ง</p>'
    ].join(''));
  }

  // ดึงชื่อ/สิทธิ์ที่ "ลงทะเบียน" ไว้ในระบบ แทนการใช้ชื่อโปรไฟล์ LINE ดิบๆ
  // ถ้ายังไม่เคยลงทะเบียน verifiedName จะเป็นค่าว่าง -> ฝั่ง User.html จะเปิด modal ให้ยืนยันชื่อ
  const registered = getRegisteredUser_(lineUser.uid);

  if (page === 'admin' && (!registered || registered.Role !== 'admin')) {
    return HtmlService.createHtmlOutput('<h3>ไม่มีสิทธิ์เข้าหน้านี้</h3><p>บัญชี LINE นี้ยังไม่มีสิทธิ์เป็น Admin ในระบบ</p>');
  }

  let templateName = 'User';
  let pageTitle = 'เบิกพัสดุ';

  if (page === 'admin') {
    templateName = 'Admin';
    pageTitle = 'ระบบจัดการคลังพัสดุ - Admin';
  } else if (page === 'history') {
    templateName = 'History';
    pageTitle = 'ประวัติการเบิก';
  }

  const template = HtmlService.createTemplateFromFile(templateName);
  template.verifiedUid = lineUser.uid;
  template.verifiedName = registered ? registered.Name : '';

  return template.evaluate()
    .setTitle(pageTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันช่วยสำหรับทดสอบและ debug ใน Apps Script Editor
function verifyLineIdToken_(idToken) {
  const channelId = getLineChannelId_();

  try {
    const response = UrlFetchApp.fetch(LINE_VERIFY_URL, {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: { id_token: idToken, client_id: channelId },
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const text = response.getContentText();
    let data = {};
    try { data = JSON.parse(text); } catch (e) { data = {}; }

    if (status === 200 && !data.error) {
      return { uid: data.sub, name: data.name || 'LINE User' };
    }

    LAST_LINE_DEBUG_INFO_ = [
      'HTTP status: ' + status,
      'error: ' + (data.error || ''),
      'error_description: ' + (data.error_description || ''),
      'channelId: ' + channelId,
      'tokenPreview: ' + (idToken ? idToken.substring(0, 30) : '(empty)')
    ].join('\n');

    // ⚠️ Fallback นี้ "ไม่ตรวจลายเซ็น" ของ token — ใช้ได้เฉพาะตอน dev/debug เท่านั้น
    // เปิดใช้งานโดยตั้ง Script Property ALLOW_UNVERIFIED_FALLBACK = 'true'
    // ห้ามเปิดทิ้งไว้ตอนขึ้นระบบจริง เพราะใครก็ปลอม token มาเป็น UID ไหนก็ได้
    const allowFallback = PropertiesService.getScriptProperties().getProperty('ALLOW_UNVERIFIED_FALLBACK') === 'true';
    if (allowFallback) {
      const fallbackUser = decodeLineTokenPayload_(idToken);
      if (fallbackUser) {
        LAST_LINE_DEBUG_INFO_ += '\nFallback: decoded token locally (UNVERIFIED — dev mode only).';
        return fallbackUser;
      }
    }

    return null;
  } catch (e) {
    LAST_LINE_DEBUG_INFO_ = 'System Error: ' + e.toString();
    return null;
  }
}

function decodeLineTokenPayload_(token) {
  if (!token) return null;
  try {
    const parts = token.split('.');
    if (parts.length < 2) return null;

    const payload = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    const padded = payload + '='.repeat((4 - (payload.length % 4)) % 4);
    const decoded = Utilities.newBlob(Utilities.base64Decode(padded)).getDataAsString();
    const data = JSON.parse(decoded);

    return { uid: data.sub, name: data.name || 'LINE User' };
  } catch (e) {
    return null;
  }
}