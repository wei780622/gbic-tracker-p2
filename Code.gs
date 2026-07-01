// ── Utility functions ──────────────────────────────────────────────

const RACK_U_COUNT = 48;
const SPREADSHEET_ID = 'PASTE_NEW_SPREADSHEET_ID_HERE';
const GEMINI_API_KEY = 'AIzaSyAHrtB4fdBwxLY9BH3Ml6vOoo2Us_mlLvw';

function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f0fe');
  }
  return sheet;
}

function safeGetData(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return [];
  return sheet.getDataRange().getValues();
}

function getRackList() {
  const sheet = getSheet('機櫃清單');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map(row => row[0]).filter(v => v !== '');
}

function getUNumbers() {
  return Array.from({ length: RACK_U_COUNT }, (_, i) => `${i + 1}U`);
}

function getModelList() {
  const sheet = getSheet('入庫紀錄');
  if (!sheet) return [];
  const data = safeGetData(sheet);
  const seen = {};
  for (let i = 1; i < data.length; i++) {
    const m = String(data[i][1] || '').trim();
    if (m) seen[m] = true;
  }
  return Object.keys(seen).sort();
}

function getGbicModel(barcode) {
  const sheet = getSheet('GBIC清單');
  const data = safeGetData(sheet);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(barcode)) return String(data[i][1]);
  }
  return '';
}

// ── doGet ───────────────────────────────────────────────────────────

function doGet(e) {
  const mode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : 'user';
  const template = HtmlService.createTemplateFromFile('Index');
  template.mode = mode;
  template.scriptUrl = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle('GMI 機房上架系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── doPost ──────────────────────────────────────────────────────────

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: 'Missing request body' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    let result = null;

    if (action === 'record')               result = recordTransaction(params);
    else if (action === 'getOptions')      result = { racks: getRackList(), uNumbers: getUNumbers(), models: getModelList() };
    else if (action === 'getModel')        result = { model: getGbicModel(params.barcode) };
    else if (action === 'queryByRack')     result = queryByRack(params.rack);
    else if (action === 'queryByBarcode')  result = queryByBarcode(params.barcode);
    else if (action === 'queryRackByDate') result = queryRackByDate(params);
    else if (action === 'queryByModel')    result = queryByModel(params);
    else if (action === 'export')          result = exportSnapshot();
    else if (action === 'recordInventoryIn')  result = recordInventoryIn(params);
    else if (action === 'recordInventoryOut') result = recordInventoryOut(params);
    else if (action === 'queryInventoryOut')  result = queryInventoryOut(params);
    else if (action === 'updateInventoryOut') result = updateInventoryOut(params);
    else if (action === 'recordInventoryReturn') result = recordInventoryReturn(params);
    else if (action === 'queryInventoryReturn')  result = queryInventoryReturn(params);
    else if (action === 'updateInventoryReturn') result = updateInventoryReturn(params);
    else if (action === 'uploadPhoto')        result = uploadPhoto(params);
    else if (action === 'queryInventoryIn')   result = queryInventoryIn(params);
    else if (action === 'deleteBatch')        result = deleteBatch(params);
    else if (action === 'deleteRow')          result = deleteRow(params);
    else if (action === 'addToBatch')         result = addToBatch(params);
    else if (action === 'updateRow')          result = updateRow(params);
    else if (action === 'ocrWithGemini')           result = ocrWithGemini(params);
    else if (action === 'updateBatch')             result = updateBatch(params);
    else if (action === 'queryInventorySummary')   result = queryInventorySummary();
    else if (action === 'checkSNInfo')             result = checkSNInfo(params);
    else if (action === 'checkSNsBatch')           result = checkSNsBatch(params);
    else if (action === 'recordBatch')             result = recordBatch(params);
    else result = { error: 'Unknown action: ' + action };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── recordTransaction ───────────────────────────────────────────────

function recordTransaction(params) {
  if (!params.barcode || !params.engineer || !params.rack || !params.uPosition || !params.type) {
    return { success: false, error: '缺少必要欄位' };
  }
  const sheet = getSheet('異動紀錄');
  sheet.appendRow([new Date(), params.type, params.engineer, params.barcode, params.model || '', params.rack, params.uPosition,
                   '', '', '', '', '', params.usage || '']);
  return { success: true };
}

// ── recordBatch ─────────────────────────────────────────────────────

function recordBatch(params) {
  if (!params.items || !params.items.length || !params.engineer || !params.rack || !params.type) {
    return { success: false, error: '缺少必要欄位' };
  }
  const sheet = getSheet('異動紀錄');
  const now = new Date();
  const rows = params.items.map(function(item) {
    return [now, params.type, params.engineer, item.barcode, item.model || '', params.rack, item.uPosition || '',
            item.deviceModel || '', item.mac1 || '', item.mac2 || '', item.extra1 || '', item.extra2 || '',
            item.usage || ''];
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  return { success: true, count: rows.length };
}

// ── recordInventoryIn ───────────────────────────────────────────────

function recordInventoryIn(params) {
  const sheet = getOrCreateSheet('入庫紀錄', ['日期', '庫存型號', '庫存數量', '簽收人', 'S/N', 'P/N', '批次ID']);
  const items = (params.items && params.items.length > 0) ? params.items : [{ sn: '', pn: '' }];
  const date = params.date ? new Date(params.date) : new Date();
  const batchId = Utilities.getUuid();
  items.forEach(function(item) {
    sheet.appendRow([date, params.model || '', params.quantity || '', params.signee || '', item.sn || '', item.pn || '', batchId]);
  });
  return { success: true, count: items.length };
}

// ── recordInventoryOut ──────────────────────────────────────────────

const OUT_HEADERS = ['日期', '領用設備', '領用型號', '領用數量', 'Rack編號', 'CSI出庫人員', '領用人員', '照片1', '照片2', '照片3', '照片4', '照片5', '設備用途'];
const RET_HEADERS = ['日期', '回庫設備', '回庫型號', '回庫數量', 'Rack編號', 'CSI回庫人員', '回庫人員', '照片1', '照片2', '照片3', '照片4', '照片5'];

function recordInventoryOut(params) {
  const sheet = getOrCreateSheet('出庫紀錄', OUT_HEADERS);
  const date   = params.date ? new Date(params.date) : new Date();
  const photos = Array.isArray(params.photos) ? params.photos : [];
  sheet.appendRow([
    date,
    params.deviceType || '',
    params.model      || '',
    params.quantity   || '',
    params.rack       || '',
    params.csiStaff   || '',
    params.recipient  || '',
    photos[0] || '',
    photos[1] || '',
    photos[2] || '',
    photos[3] || '',
    photos[4] || '',
    params.usage      || ''
  ]);
  return { success: true };
}

// ── queryInventoryIn ────────────────────────────────────────────────

function queryInventoryIn(params) {
  const sheet = getOrCreateSheet('入庫紀錄', ['日期', '庫存型號', '庫存數量', '簽收人', 'S/N', 'P/N', '批次ID']);
  const data = safeGetData(sheet);
  const dateFrom = (params && params.dateFrom && params.dateFrom !== '') ? new Date(params.dateFrom) : null;
  const dateTo   = (params && params.dateTo   && params.dateTo   !== '') ? new Date(params.dateTo + 'T23:59:59') : null;

  // Group by batchId (col 6), fallback key = date+model+signee for old data
  const batchMap = {};
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0] ? new Date(data[i][0]) : null;
    if (dateFrom && rowDate && rowDate < dateFrom) continue;
    if (dateTo   && rowDate && rowDate > dateTo)   continue;

    const batchId = String(data[i][6] || '') || (String(data[i][0]) + '|' + String(data[i][1]) + '|' + String(data[i][3]));
    if (!batchMap[batchId]) {
      batchMap[batchId] = {
        batchId:  batchId,
        date:     data[i][0],
        model:    String(data[i][1]),
        quantity: data[i][2],
        signee:   String(data[i][3]),
        items:    []
      };
    }
    batchMap[batchId].items.push({ sn: String(data[i][4] || ''), pn: String(data[i][5] || ''), _row: i + 1 });
  }

  return Object.values(batchMap).sort(function(a, b) { return new Date(b.date) - new Date(a.date); });
}

// ── deleteBatch ─────────────────────────────────────────────────────

function deleteBatch(params) {
  const sheet = getSheet('入庫紀錄');
  if (!sheet) return { success: false, error: '找不到工作表' };
  const rows = (params.rows || []).map(Number).sort(function(a, b) { return b - a; });
  rows.forEach(function(r) { if (r >= 2) sheet.deleteRow(r); });
  return { success: true };
}

// ── queryInventoryOut ───────────────────────────────────────────────

function queryInventoryOut(params) {
  const sheet    = getOrCreateSheet('出庫紀錄', OUT_HEADERS);
  const data     = safeGetData(sheet);
  const dateFrom = (params && params.dateFrom) ? new Date(params.dateFrom) : null;
  const dateTo   = (params && params.dateTo)   ? new Date(params.dateTo + 'T23:59:59') : null;
  const results  = [];
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0] ? new Date(data[i][0]) : null;
    if (dateFrom && rowDate && rowDate < dateFrom) continue;
    if (dateTo   && rowDate && rowDate > dateTo)   continue;
    results.push({
      date:      data[i][0],
      device:    String(data[i][1] || ''),
      model:     String(data[i][2] || ''),
      quantity:  data[i][3],
      rack:      String(data[i][4] || ''),
      csiStaff:  String(data[i][5] || ''),
      recipient: String(data[i][6] || ''),
      photos:    [
        String(data[i][7]  || ''),
        String(data[i][8]  || ''),
        String(data[i][9]  || ''),
        String(data[i][10] || ''),
        String(data[i][11] || '')
      ].filter(Boolean),
      usage:     String(data[i][12] || ''),
      _sheet:    '出庫紀錄',
      _row:      i + 1
    });
  }
  return results.reverse();
}

// ── updateInventoryOut ──────────────────────────────────────────────

function updateInventoryOut(params) {
  const sheet = getSheet('出庫紀錄');
  if (!sheet) return { success: false, error: '找不到工作表' };
  const row = parseInt(params.rowIndex);
  if (!row || row < 2) return { success: false, error: '參數錯誤' };
  const photos = Array.isArray(params.photos) ? params.photos : [];
  const date   = params.date ? new Date(params.date) : new Date();
  sheet.getRange(row, 1, 1, 13).setValues([[
    date,
    params.deviceType || '',
    params.model      || '',
    params.quantity   || '',
    params.rack       || '',
    params.csiStaff   || '',
    params.recipient  || '',
    photos[0] || '',
    photos[1] || '',
    photos[2] || '',
    photos[3] || '',
    photos[4] || '',
    params.usage      || ''
  ]]);
  return { success: true };
}

// ── recordInventoryReturn ────────────────────────────────────────────

function recordInventoryReturn(params) {
  const sheet = getOrCreateSheet('回庫紀錄', RET_HEADERS);
  const date   = params.date ? new Date(params.date) : new Date();
  const photos = Array.isArray(params.photos) ? params.photos : [];
  sheet.appendRow([
    date,
    params.deviceType  || '',
    params.model       || '',
    params.quantity    || '',
    params.rack        || '',
    params.csiStaff    || '',
    params.returnStaff || '',
    photos[0] || '',
    photos[1] || '',
    photos[2] || '',
    photos[3] || '',
    photos[4] || ''
  ]);
  return { success: true };
}

// ── queryInventoryReturn ─────────────────────────────────────────────

function queryInventoryReturn(params) {
  const sheet    = getOrCreateSheet('回庫紀錄', RET_HEADERS);
  const data     = safeGetData(sheet);
  const dateFrom = (params && params.dateFrom) ? new Date(params.dateFrom) : null;
  const dateTo   = (params && params.dateTo)   ? new Date(params.dateTo + 'T23:59:59') : null;
  const results  = [];
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0] ? new Date(data[i][0]) : null;
    if (dateFrom && rowDate && rowDate < dateFrom) continue;
    if (dateTo   && rowDate && rowDate > dateTo)   continue;
    results.push({
      date:        data[i][0],
      device:      String(data[i][1] || ''),
      model:       String(data[i][2] || ''),
      quantity:    data[i][3],
      rack:        String(data[i][4] || ''),
      csiStaff:    String(data[i][5] || ''),
      returnStaff: String(data[i][6] || ''),
      photos:      [
        String(data[i][7]  || ''),
        String(data[i][8]  || ''),
        String(data[i][9]  || ''),
        String(data[i][10] || ''),
        String(data[i][11] || '')
      ].filter(Boolean),
      _sheet: '回庫紀錄',
      _row:   i + 1
    });
  }
  return results.reverse();
}

// ── updateInventoryReturn ────────────────────────────────────────────

function updateInventoryReturn(params) {
  const sheet = getSheet('回庫紀錄');
  if (!sheet) return { success: false, error: '找不到工作表' };
  const row = parseInt(params.rowIndex);
  if (!row || row < 2) return { success: false, error: '參數錯誤' };
  const photos = Array.isArray(params.photos) ? params.photos : [];
  const date   = params.date ? new Date(params.date) : new Date();
  sheet.getRange(row, 1, 1, 12).setValues([[
    date,
    params.deviceType  || '',
    params.model       || '',
    params.quantity    || '',
    params.rack        || '',
    params.csiStaff    || '',
    params.returnStaff || '',
    photos[0] || '',
    photos[1] || '',
    photos[2] || '',
    photos[3] || '',
    photos[4] || ''
  ]]);
  return { success: true };
}

// ── uploadPhoto ─────────────────────────────────────────────────────

function getOrCreatePhotoFolder() {
  const name = 'GBIC-Photos';
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function uploadPhoto(params) {
  if (!params.base64 || !params.mimeType) return { success: false, error: '缺少圖片資料' };
  const folder   = getOrCreatePhotoFolder();
  const decoded  = Utilities.base64Decode(params.base64);
  const blob     = Utilities.newBlob(decoded, params.mimeType, params.filename || 'photo.jpg');
  const file     = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/uc?export=view&id=' + file.getId();
  return { success: true, url: url };
}

// ── deleteRow ───────────────────────────────────────────────────────

function deleteRow(params) {
  const sheetName = params.sheetName;
  const rowIndex  = parseInt(params.rowIndex);
  if (!sheetName || !rowIndex || rowIndex < 2) return { success: false, error: '參數錯誤' };
  const sheet = getSheet(sheetName);
  if (!sheet) return { success: false, error: '找不到工作表' };
  if (sheet.getLastRow() < rowIndex) return { success: false, error: '列不存在' };
  sheet.deleteRow(rowIndex);
  return { success: true };
}

// ── addToBatch ──────────────────────────────────────────────────────

function addToBatch(params) {
  const sheet = getOrCreateSheet('入庫紀錄', ['日期', '庫存型號', '庫存數量', '簽收人', 'S/N', 'P/N', '批次ID']);
  const date = params.date ? new Date(params.date) : new Date();
  sheet.appendRow([date, params.model || '', params.quantity || '', params.signee || '', params.sn || '', params.pn || '', params.batchId || '']);
  return { success: true };
}

// ── updateRow ────────────────────────────────────────────────────────

function updateRow(params) {
  const sheet = getSheet(params.sheetName);
  if (!sheet) return { success: false, error: '找不到工作表' };
  const row = parseInt(params.rowIndex);
  if (!row || row < 2) return { success: false, error: '列不存在' };
  if (params.sheetName === '異動紀錄') {
    if (params.uPosition   !== undefined) sheet.getRange(row, 7).setValue(params.uPosition);
    if (params.deviceModel !== undefined) sheet.getRange(row, 8).setValue(params.deviceModel);
    if (params.mac1        !== undefined) sheet.getRange(row, 9).setValue(params.mac1);
    if (params.mac2        !== undefined) sheet.getRange(row, 10).setValue(params.mac2);
    if (params.extra1      !== undefined) sheet.getRange(row, 11).setValue(params.extra1);
    if (params.extra2      !== undefined) sheet.getRange(row, 12).setValue(params.extra2);
    if (params.usage       !== undefined) sheet.getRange(row, 13).setValue(params.usage);
  } else {
    if (params.sn !== undefined) sheet.getRange(row, 5).setValue(params.sn);
    if (params.pn !== undefined) sheet.getRange(row, 6).setValue(params.pn);
  }
  return { success: true };
}

// ── updateBatch ───────────────────────────────────────────────────────

function updateBatch(params) {
  const sheet = getSheet('入庫紀錄');
  if (!sheet) return { success: false, error: '找不到工作表' };
  const rows = (params.rows || []).map(Number);
  rows.forEach(function(r) {
    if (r < 2) return;
    if (params.model    !== undefined) sheet.getRange(r, 2).setValue(params.model);
    if (params.quantity !== undefined) sheet.getRange(r, 3).setValue(params.quantity);
  });
  return { success: true };
}

// ── ocrWithGemini ────────────────────────────────────────────────────

function ocrWithGemini(params) {
  const payload = {
    contents: [{
      parts: [
        { text: '這是一張設備出貨標籤照片。請找出所有 S/N（序號），每行只輸出一個序號，不要任何其他文字、標點、編號或說明。序號通常是英文字母加數字的組合，例如 MT2606NT02XT8。' },
        { inlineData: { mimeType: params.mimeType || 'image/jpeg', data: params.imageBase64 } }
      ]
    }],
    generationConfig: { temperature: 0, maxOutputTokens: 1024 }
  };

  const resp = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + GEMINI_API_KEY,
    { method: 'POST', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true }
  );

  const json = JSON.parse(resp.getContentText());
  if (json.error) return { success: false, error: json.error.message };

  const text = json.candidates[0].content.parts[0].text;
  const sns = text.split('\n')
    .map(function(s) { return s.trim().replace(/^[-*\d.)]\s*/, '').trim(); })
    .filter(function(s) { return s.length >= 4 && /^[A-Za-z0-9]/.test(s); });

  return { success: true, sns: sns };
}

// ── updateRackList（執行一次即可，之後可刪除）────────────────────────

function updateRackList() {
  const sheet = getOrCreateSheet('機櫃清單', ['機櫃編號']);
  const extraRows = { 'C': 12, 'D': 12, 'I': 12, 'J': 12 };
  const rows = 'ABCDEFGHIJ'.split('').flatMap(function(r) {
    const count = extraRows[r] || 10;
    return Array.from({ length: count }, function(_, i) {
      return [r + String(i + 1).padStart(2, '0')];
    });
  });
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
  sheet.getRange(2, 1, rows.length, 1).setValues(rows);
  Logger.log('機櫃清單已更新：C、D、I、J 各 12 台，其餘各 10 台，共 ' + rows.length + ' 台');
}

// ── buildLatestSnapshot ─────────────────────────────────────────────

function buildLatestSnapshot() {
  const sheet = getSheet('異動紀錄');
  const data = safeGetData(sheet);
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const barcode = String(data[i][3]);
    if (!barcode) continue;
    map[barcode] = {
      timestamp: data[i][0], type: data[i][1], engineer: data[i][2],
      barcode, model: data[i][4], rack: data[i][5], uPosition: data[i][6],
      deviceModel: data[i][7] || '', mac1: data[i][8] || '', mac2: data[i][9] || '',
      extra1: data[i][10] || '', extra2: data[i][11] || '', usage: data[i][12] || '',
      _sheet: '異動紀錄', _row: i + 1
    };
  }
  return Object.values(map);
}

function queryByRack(rack) {
  return buildLatestSnapshot().filter(
    item => item.type === '上架' && String(item.rack) === String(rack)
  );
}

function queryByBarcode(barcode) {
  const sheet = getSheet('異動紀錄');
  const data = safeGetData(sheet);
  const results = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]) === String(barcode)) {
      results.push({
        timestamp: data[i][0], type: data[i][1], engineer: data[i][2],
        barcode: String(data[i][3]), model: data[i][4], rack: data[i][5], uPosition: data[i][6],
        deviceModel: data[i][7] || '', mac1: data[i][8] || '', mac2: data[i][9] || '',
        extra1: data[i][10] || '', extra2: data[i][11] || '', usage: data[i][12] || '',
        _sheet: '異動紀錄', _row: i + 1
      });
    }
  }
  return results.reverse();
}

// ── queryByModel ─────────────────────────────────────────────────────

function queryByModel(params) {
  const model = String(params.model || '').trim();
  if (!model) return { error: '請輸入型號' };
  const kw = model.toLowerCase();
  const tz = Session.getScriptTimeZone();

  function match(v) { return String(v || '').toLowerCase().indexOf(kw) !== -1; }

  const results = [];

  // 入庫（批次去重）
  const inData = safeGetData(getOrCreateSheet('入庫紀錄', ['日期','庫存型號','庫存數量','簽收人','S/N','P/N','批次ID']));
  const batchSeen = {};
  for (let i = 1; i < inData.length; i++) {
    if (!match(inData[i][1])) continue;
    const bid = String(inData[i][6] || '') || (String(inData[i][0]) + '|' + String(inData[i][1]) + '|' + String(inData[i][3]));
    if (batchSeen[bid]) continue;
    batchSeen[bid] = true;
    results.push({ type: '入庫', date: inData[i][0], model: String(inData[i][1] || ''),
      quantity: Number(inData[i][2] || 0), person: String(inData[i][3] || ''), rack: '', usage: '' });
  }

  // 出庫
  const outData = safeGetData(getOrCreateSheet('出庫紀錄', OUT_HEADERS));
  for (let i = 1; i < outData.length; i++) {
    if (!match(outData[i][2])) continue;
    results.push({ type: '出庫', date: outData[i][0], model: String(outData[i][2] || ''),
      quantity: Number(outData[i][3] || 0),
      person: String(outData[i][6] || '') + (outData[i][5] ? '（CSI:' + outData[i][5] + '）' : ''),
      rack: String(outData[i][4] || ''), usage: String(outData[i][12] || '') });
  }

  // 回庫
  const retData = safeGetData(getOrCreateSheet('回庫紀錄', RET_HEADERS));
  for (let i = 1; i < retData.length; i++) {
    if (!match(retData[i][2])) continue;
    results.push({ type: '回庫', date: retData[i][0], model: String(retData[i][2] || ''),
      quantity: Number(retData[i][3] || 0),
      person: String(retData[i][6] || '') + (retData[i][5] ? '（CSI:' + retData[i][5] + '）' : ''),
      rack: String(retData[i][4] || ''), usage: '' });
  }

  // 上架 / 歸還
  const rackSheet = getSheet('異動紀錄');
  const rackData  = rackSheet ? safeGetData(rackSheet) : [];
  for (let i = 1; i < rackData.length; i++) {
    const devModel = String(rackData[i][7] || '').trim();
    const pnModel  = String(rackData[i][4] || '').trim();
    if (!match(devModel) && !match(pnModel)) continue;
    const type = String(rackData[i][1] || '').trim();
    if (type !== '上架' && type !== '歸還') continue;
    results.push({ type: type, date: rackData[i][0], model: devModel || pnModel,
      quantity: 1, barcode: String(rackData[i][3] || ''),
      person: String(rackData[i][2] || ''),
      rack: String(rackData[i][5] || '') + (rackData[i][6] ? ' ' + rackData[i][6] : ''),
      usage: String(rackData[i][12] || '') });
  }

  results.sort(function(a, b) {
    return (b.date ? new Date(b.date).getTime() : 0) - (a.date ? new Date(a.date).getTime() : 0);
  });
  return results;
}

// ── queryRackByDate ──────────────────────────────────────────────────

function queryRackByDate(params) {
  const sheet = getSheet('異動紀錄');
  if (!sheet) return [];
  const data     = safeGetData(sheet);
  const dateFrom = (params && params.dateFrom) ? new Date(params.dateFrom) : null;
  const dateTo   = (params && params.dateTo)   ? new Date(params.dateTo + 'T23:59:59') : null;
  const results  = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').trim() !== '上架') continue;
    const rowDate = data[i][0] ? new Date(data[i][0]) : null;
    if (dateFrom && rowDate && rowDate < dateFrom) continue;
    if (dateTo   && rowDate && rowDate > dateTo)   continue;
    results.push({
      timestamp:   data[i][0],
      engineer:    String(data[i][2] || ''),
      barcode:     String(data[i][3] || ''),
      model:       String(data[i][4] || ''),
      rack:        String(data[i][5] || ''),
      uPosition:   String(data[i][6] || ''),
      deviceModel: String(data[i][7] || ''),
      mac1:        String(data[i][8] || ''),
      mac2:        String(data[i][9] || ''),
      extra1:      String(data[i][10] || ''),
      extra2:      String(data[i][11] || ''),
      usage:       String(data[i][12] || ''),
      _sheet:      '異動紀錄',
      _row:        i + 1
    });
  }
  return results.reverse();
}

function exportSnapshot() {
  return buildLatestSnapshot();
}

// ── queryInventorySummary ────────────────────────────────────────────

function queryInventorySummary() {
  const inSheet  = getOrCreateSheet('入庫紀錄', ['日期', '庫存型號', '庫存數量', '簽收人', 'S/N', 'P/N', '批次ID']);
  const outSheet = getOrCreateSheet('出庫紀錄', ['日期', '領用設備', '領用型號', '領用數量', 'Rack編號', 'CSI出庫人員', '領用人員']);
  const retSheet = getOrCreateSheet('回庫紀錄', RET_HEADERS);
  const inData   = safeGetData(inSheet);
  const outData  = safeGetData(outSheet);
  const retData  = safeGetData(retSheet);

  const tz = Session.getScriptTimeZone();
  function fmtD(d) { return Utilities.formatDate(new Date(d), tz, 'M/d'); }

  // Sum 入庫 quantity per batch, collect transaction list
  const batchSeen = {};
  const modelMap  = {};
  for (let i = 1; i < inData.length; i++) {
    const model   = String(inData[i][1] || '').trim();
    const qty     = Number(inData[i][2] || 0);
    const batchId = String(inData[i][6] || '') ||
                    (String(inData[i][0]) + '|' + model + '|' + String(inData[i][3]));
    if (!modelMap[model]) modelMap[model] = { in: 0, out: 0, ret: 0, rack: 0, inTx: [], outTx: [], retTx: [], rackTx: [] };
    if (!batchSeen[batchId]) {
      batchSeen[batchId] = true;
      modelMap[model].in += qty;
      const ds = inData[i][0] ? fmtD(inData[i][0]) : '?';
      modelMap[model].inTx.push({ ds: ds, qty: qty, ts: inData[i][0] ? new Date(inData[i][0]).getTime() : 0 });
    }
  }

  // Sum 出庫 quantity per model, collect transaction list + purposes
  for (let i = 1; i < outData.length; i++) {
    const model = String(outData[i][2] || '').trim();
    const qty   = Number(outData[i][3] || 0);
    if (!modelMap[model]) modelMap[model] = { in: 0, out: 0, ret: 0, rack: 0, inTx: [], outTx: [], retTx: [], rackTx: [], purposes: new Set() };
    if (!modelMap[model].purposes) modelMap[model].purposes = new Set();
    modelMap[model].out += qty;
    const ds = outData[i][0] ? fmtD(outData[i][0]) : '?';
    modelMap[model].outTx.push({ ds: ds, qty: qty, ts: outData[i][0] ? new Date(outData[i][0]).getTime() : 0 });
    const usage = String(outData[i][12] || '').trim();
    if (usage) modelMap[model].purposes.add(usage);
  }

  // Sum 回庫 quantity per model, collect transaction list
  for (let i = 1; i < retData.length; i++) {
    const model = String(retData[i][2] || '').trim();
    const qty   = Number(retData[i][3] || 0);
    if (!modelMap[model]) modelMap[model] = { in: 0, out: 0, ret: 0, rack: 0, inTx: [], outTx: [], retTx: [], rackTx: [] };
    modelMap[model].ret += qty;
    const ds = retData[i][0] ? fmtD(retData[i][0]) : '?';
    modelMap[model].retTx.push({ ds: ds, qty: qty, ts: retData[i][0] ? new Date(retData[i][0]).getTime() : 0 });
  }

  // 用 snapshot 計算「目前實際上架中」的數量（已歸還的不計入）
  const snapshot = buildLatestSnapshot();
  const rackModelMap = {}; // model -> { count, dateMap: { ds -> { qty, ts } } }
  for (let si = 0; si < snapshot.length; si++) {
    const item = snapshot[si];
    if (item.type !== '上架') continue;
    const model = (String(item.deviceModel || '').trim()) || (String(item.model || '').trim());
    if (!model) continue;
    if (!rackModelMap[model]) rackModelMap[model] = { count: 0, dateMap: {} };
    rackModelMap[model].count += 1;
    const ts = item.timestamp ? new Date(item.timestamp).getTime() : 0;
    const ds = item.timestamp ? fmtD(item.timestamp) : '?';
    if (!rackModelMap[model].dateMap[ds]) rackModelMap[model].dateMap[ds] = { qty: 0, ts: ts };
    rackModelMap[model].dateMap[ds].qty += 1;
    // 收集上架用途
    const usage = String(item.usage || '').trim();
    if (usage) {
      if (!modelMap[model]) modelMap[model] = { in: 0, out: 0, ret: 0, rack: 0, inTx: [], outTx: [], retTx: [], rackTx: [], purposes: new Set() };
      if (!modelMap[model].purposes) modelMap[model].purposes = new Set();
      modelMap[model].purposes.add(usage);
    }
  }
  for (const model in rackModelMap) {
    if (!modelMap[model]) modelMap[model] = { in: 0, out: 0, ret: 0, rack: 0, inTx: [], outTx: [], retTx: [], rackTx: [], purposes: new Set() };
    modelMap[model].rack = rackModelMap[model].count;
    modelMap[model].rackTx = Object.keys(rackModelMap[model].dateMap).map(function(ds) {
      return { ds: ds, qty: rackModelMap[model].dateMap[ds].qty, ts: rackModelMap[model].dateMap[ds].ts };
    });
  }

  return Object.keys(modelMap).map(function(model) {
    const m = modelMap[model];
    m.inTx.sort(function(a,b){return b.ts-a.ts;});
    m.retTx.sort(function(a,b){return b.ts-a.ts;});
    // 合併出庫 + 上架，按時間排序
    const usedArr = m.outTx.map(function(t){ return { ds: t.ds, qty: t.qty, ts: t.ts, type: '出' }; })
      .concat((m.rackTx || []).map(function(t){ return { ds: t.ds, qty: t.qty, ts: t.ts, type: '架' }; }))
      .sort(function(a,b){ return b.ts - a.ts; });
    return {
      model:      model,
      purposes:   m.purposes ? Array.from(m.purposes) : [],
      totalIn:    m.in,
      totalUsed:  m.out + (m.rack || 0),
      totalRet:   m.ret,
      current:    m.in - m.out + m.ret - (m.rack || 0),
      inTx:   m.inTx.map(function(t){ return t.ds + ' ×' + t.qty; }),
      usedTx: usedArr.map(function(t){ return t.ds + ' ×' + t.qty + '(' + t.type + ')'; }),
      retTx:  m.retTx.map(function(t){ return t.ds + ' ×' + t.qty; })
    };
  }).sort(function(a, b) { return String(a.model).localeCompare(String(b.model)); });
}

// ── checkSNInfo ──────────────────────────────────────────────────────

function checkSNInfo(params) {
  const sn = String(params.sn || '').trim();
  if (!sn) return { sn: sn, inRecord: null, rackRecord: null };

  const inSheet    = getSheet('入庫紀錄');
  const transSheet = getSheet('異動紀錄');
  let inRecord = null, rackRecord = null;

  if (inSheet) {
    const data = safeGetData(inSheet);
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][4] || '').trim() === sn) {
        inRecord = { date: data[i][0], model: String(data[i][1]) };
        break;
      }
    }
  }

  if (transSheet) {
    const data = safeGetData(transSheet);
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][3] || '').trim() === sn) {
        rackRecord = { type: String(data[i][1]), rack: String(data[i][5]), date: data[i][0] };
        break;
      }
    }
  }

  return { sn: sn, inRecord: inRecord, rackRecord: rackRecord };
}

// ── checkSNsBatch ────────────────────────────────────────────────────

function checkSNsBatch(params) {
  const sns = (params.sns || []).map(String);
  if (!sns.length) return [];

  const inSheet    = getSheet('入庫紀錄');
  const transSheet = getSheet('異動紀錄');
  const inMap = {}, rackMap = {};

  if (inSheet) {
    const data = safeGetData(inSheet);
    for (let i = 1; i < data.length; i++) {
      const sn = String(data[i][4] || '').trim();
      if (sn && !inMap[sn]) inMap[sn] = { date: data[i][0], model: String(data[i][1]) };
    }
  }

  if (transSheet) {
    const data = safeGetData(transSheet);
    for (let i = 1; i < data.length; i++) {
      const sn = String(data[i][3] || '').trim();
      if (sn) rackMap[sn] = { type: String(data[i][1]), rack: String(data[i][5]), date: data[i][0] };
    }
  }

  return sns.map(function(sn) {
    return { sn: sn, inRecord: inMap[sn] || null, rackRecord: rackMap[sn] || null };
  });
}
