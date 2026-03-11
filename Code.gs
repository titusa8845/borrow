/**
 * 羽毛球借球系統 — Google Apps Script 後端
 *
 * 使用方式：
 * 1. 建立 Google 試算表，建立以下 5 個 Sheet（工作表）：
 *    - 「班級」：A欄=班級名稱
 *    - 「學生」：A欄=班級名稱, B欄=座號, C欄=姓名
 *    - 「器材」：A欄=器材名稱, B欄=單位, C欄=總數量, D欄=已借出, E欄=單次上限
 *    - 「借用紀錄」：A欄=時間戳, B欄=班級, C欄=座號, D欄=姓名, E欄=器材, F欄=數量, G欄=狀態
 *    - 「設定」：A1=ADMIN_PASSWORD, B1=你的管理密碼
 * 2. 在「器材」Sheet 預填資料，例如：羽毛球|顆|50|0|5 和 球拍|支|20|0|2
 * 3. 開啟 Apps Script（擴充功能 > Apps Script），貼上此檔案內容
 * 4. 部署 > 新增部署 > 網頁應用程式 > 存取權限選「所有人」
 * 5. 複製部署的網址，貼到 index.html 的 API_URL 變數
 */

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  const params = e ? e.parameter : {};
  try {
    const result = dispatch(params);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function dispatch(p) {
  switch (p.action) {
    // === 公開 API ===
    case 'getClasses':    return { ok: true, data: getClasses() };
    case 'getStudents':   return { ok: true, data: getStudents(p['class']) };
    case 'getEquipment':  return { ok: true, data: getEquipment() };
    case 'borrow':        return doBorrow(p);

    // === 管理 API（需密碼）===
    case 'verifyPassword': return { ok: verifyPw(p.pw) };
    case 'getRecords':     return auth(p, () => ({ ok: true, data: getRecords(p.filter) }));
    case 'returnItem':     return auth(p, () => doReturn(p.row));
    case 'addClass':       return auth(p, () => doAddClass(p.name));
    case 'deleteClass':    return auth(p, () => doDeleteClass(p.name));
    case 'addStudent':     return auth(p, () => doAddStudent(p['class'], p.seat, p.name));
    case 'importStudents': return auth(p, () => doImportStudents(p.data));
    case 'deleteStudent':  return auth(p, () => doDeleteStudent(p['class'], p.seat));
    case 'updateEquipment': return auth(p, () => doUpdateEquipment(p.name, p.total, p.max));
    case 'addEquipment':   return auth(p, () => doAddEquipment(p.name, p.unit, p.total, p.max));
    case 'getSettings':    return auth(p, () => ({ ok: true, data: getSettings() }));
    case 'setSetting':     return auth(p, () => doSetSetting(p.key, p.value));
    case 'resetInventory': return auth(p, () => doResetInventory());

    default: return { ok: false, error: '未知的 action' };
  }
}

// ========== 密碼驗證 ==========
function verifyPw(pw) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  const password = sheet.getRange('B1').getValue().toString();
  return pw === password;
}

function auth(p, fn) {
  if (!verifyPw(p.pw)) return { ok: false, error: '密碼錯誤' };
  return fn();
}

// ========== 設定讀寫 ==========
function getSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (const row of data) {
    if (row[0]) settings[row[0]] = row[1];
  }
  return settings;
}

function doSetSetting(key, value) {
  if (!key) return { ok: false, error: '缺少 key' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex(row => row[0] === key);
  if (rowIdx !== -1) {
    sheet.getRange(rowIdx + 1, 2).setValue(value);
  } else {
    sheet.appendRow([key, value]);
  }
  return { ok: true, message: '設定已更新' };
}

function isReturnEnabled() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  const data = sheet.getDataRange().getValues();
  const row = data.find(r => r[0] === 'RETURN_ENABLED');
  return row ? row[1].toString().toUpperCase() === 'TRUE' : false;
}

// ========== 公開 API 實作 ==========
function getClasses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('班級');
  const data = sheet.getDataRange().getValues();
  return data.map(row => row[0]).filter(v => v !== '');
}

function getStudents(className) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('學生');
  const data = sheet.getDataRange().getValues();
  return data
    .filter(row => row[0] === className)
    .map(row => ({ seat: row[1], name: row[2] }))
    .sort((a, b) => a.seat - b.seat);
}

function getEquipment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('器材');
  const data = sheet.getDataRange().getValues();
  const returnOn = isReturnEnabled();

  // 只有歸還功能開啟時才統計今日借出量
  const todayBorrowed = returnOn ? getTodayBorrowed(ss) : {};

  return data.slice(1).filter(row => row[0] !== '').map(row => {
    const name = row[0];
    const total = Number(row[2]);
    const borrowed = todayBorrowed[name] || 0;
    return {
      name,
      unit: row[1],
      total,
      borrowed,
      available: returnOn ? (total - borrowed) : total,  // 歸還關閉時不限制
      maxPerBorrow: Number(row[4]),
      returnEnabled: returnOn
    };
  });
}

// 計算今天各器材的借出總量
function getTodayBorrowed(ss) {
  const recSheet = ss.getSheetByName('借用紀錄');
  const recData = recSheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
  const result = {};
  for (const row of recData) {
    if (!row[0]) continue;
    const recDate = Utilities.formatDate(new Date(row[0]), 'Asia/Taipei', 'yyyy/MM/dd');
    if (recDate === today && row[6] === '借出中') {
      result[row[4]] = (result[row[4]] || 0) + Number(row[5]);
    }
  }
  return result;
}

// ========== 借用 ==========
function doBorrow(p) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { ok: false, error: '系統忙碌中，請稍後再試' };
  }

  try {
    // items 格式: "羽毛球:3,球拍:1"
    const items = p.items.split(',').map(s => {
      const [name, qty] = s.split(':');
      return { name: name.trim(), qty: parseInt(qty) };
    }).filter(item => item.qty > 0);

    if (items.length === 0) return { ok: false, error: '請至少選擇一項器材' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eqSheet = ss.getSheetByName('器材');
    const eqData = eqSheet.getDataRange().getValues();
    const recSheet = ss.getSheetByName('借用紀錄');
    const now = new Date();

    const returnOn = isReturnEnabled();
    const todayBorrowed = returnOn ? getTodayBorrowed(ss) : {};

    // 檢查庫存
    for (const item of items) {
      const rowIdx = eqData.findIndex(row => row[0] === item.name);
      if (rowIdx === -1) return { ok: false, error: `找不到器材：${item.name}` };
      const maxPer = Number(eqData[rowIdx][4]);
      if (item.qty > maxPer) return { ok: false, error: `${item.name} 單次上限為 ${maxPer}` };
      if (returnOn) {
        const total = Number(eqData[rowIdx][2]);
        const borrowed = todayBorrowed[item.name] || 0;
        const available = total - borrowed;
        if (item.qty > available) return { ok: false, error: `${item.name} 今日庫存不足（剩餘 ${available}）` };
      }
    }

    // 寫入紀錄（不再更新器材表的已借出欄）
    for (const item of items) {
      recSheet.appendRow([now, p['class'], p.seat, p.name, item.name, item.qty, '借出中']);
    }

    return { ok: true, message: '借用成功！' };
  } finally {
    lock.releaseLock();
  }
}

// ========== 歸還 ==========
function doReturn(rowNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recSheet = ss.getSheetByName('借用紀錄');
  const row = parseInt(rowNum);
  const status = recSheet.getRange(row, 7).getValue();
  if (status !== '借出中') return { ok: false, error: '此紀錄已歸還' };

  const eqName = recSheet.getRange(row, 5).getValue();
  const qty = Number(recSheet.getRange(row, 6).getValue());

  // 標記已歸還
  recSheet.getRange(row, 7).setValue('已歸還');

  return { ok: true, message: '歸還成功！' };
}

// ========== 紀錄查詢 ==========
function getRecords(filter) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('借用紀錄');
  const data = sheet.getDataRange().getValues();
  const records = data.map((row, idx) => ({
    row: idx + 1,
    time: row[0] ? Utilities.formatDate(new Date(row[0]), 'Asia/Taipei', 'yyyy/MM/dd HH:mm') : '',
    className: row[1],
    seat: row[2],
    name: row[3],
    equipment: row[4],
    qty: row[5],
    status: row[6]
  })).filter(r => r.time !== '');

  if (filter === 'active') return records.filter(r => r.status === '借出中');
  if (filter && filter.match(/^\d{4}\/\d{2}\/\d{2}$/)) {
    return records.filter(r => r.time.startsWith(filter)).reverse();
  }
  return records.reverse();
}

// ========== 班級管理 ==========
function doAddClass(name) {
  if (!name) return { ok: false, error: '請輸入班級名稱' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('班級');
  const existing = sheet.getDataRange().getValues().flat();
  if (existing.includes(name)) return { ok: false, error: '班級已存在' };
  sheet.appendRow([name]);
  return { ok: true, message: '班級新增成功' };
}

function doDeleteClass(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('班級');
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex(row => row[0] === name);
  if (rowIdx === -1) return { ok: false, error: '找不到此班級' };
  sheet.deleteRow(rowIdx + 1);
  return { ok: true, message: '班級已刪除' };
}

// ========== 學生管理 ==========
function doAddStudent(className, seat, name) {
  if (!className || !seat || !name) return { ok: false, error: '請填寫完整資料' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('學生');
  sheet.appendRow([className, parseInt(seat), name]);
  return { ok: true, message: '學生新增成功' };
}

function doDeleteStudent(className, seat) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('學生');
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex(row => row[0] === className && row[1] == seat);
  if (rowIdx === -1) return { ok: false, error: '找不到此學生' };
  sheet.deleteRow(rowIdx + 1);
  return { ok: true, message: '學生已刪除' };
}

// ========== 學生批量匯入 ==========
function doImportStudents(jsonStr) {
  try {
    const students = JSON.parse(jsonStr);
    if (!Array.isArray(students)) return { ok: false, error: '資料格式錯誤，需為陣列' };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('學生');
    const classSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('班級');
    const existingClasses = classSheet.getDataRange().getValues().flat().filter(v => v !== '');
    let count = 0;
    const newClasses = new Set();
    for (const s of students) {
      const cls = s['班級'] || s['class'] || '';
      const seat = s['座號'] || s['seat'] || '';
      const name = s['姓名'] || s['name'] || '';
      if (!cls || !seat || !name) continue;
      sheet.appendRow([cls, parseInt(seat), name]);
      if (!existingClasses.includes(cls) && !newClasses.has(cls)) {
        classSheet.appendRow([cls]);
        newClasses.add(cls);
      }
      count++;
    }
    return { ok: true, message: `成功匯入 ${count} 位學生` + (newClasses.size > 0 ? `，自動新增 ${newClasses.size} 個班級` : '') };
  } catch (e) {
    return { ok: false, error: 'JSON 格式錯誤：' + e.message };
  }
}

// ========== 器材管理 ==========
function doUpdateEquipment(name, total, max) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('器材');
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex(row => row[0] === name);
  if (rowIdx === -1) return { ok: false, error: '找不到此器材' };
  if (total !== undefined) sheet.getRange(rowIdx + 1, 3).setValue(parseInt(total));
  if (max !== undefined) sheet.getRange(rowIdx + 1, 5).setValue(parseInt(max));
  return { ok: true, message: '器材更新成功' };
}

function doAddEquipment(name, unit, total, max) {
  if (!name || !unit) return { ok: false, error: '請填寫完整資料' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('器材');
  sheet.appendRow([name, unit, parseInt(total) || 0, 0, parseInt(max) || 1]);
  return { ok: true, message: '器材新增成功' };
}

// ========== 重設庫存（將今日所有「借出中」標記為已歸還）==========
function doResetInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recSheet = ss.getSheetByName('借用紀錄');
  const data = recSheet.getDataRange().getValues();
  let count = 0;
  for (let i = 0; i < data.length; i++) {
    if (data[i][6] === '借出中') {
      recSheet.getRange(i + 1, 7).setValue('已歸還');
      count++;
    }
  }
  return { ok: true, message: `已重設 ${count} 筆借用紀錄為已歸還` };
}
