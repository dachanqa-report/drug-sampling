/**
 * 屏東廠藥品品質抽查表 (Rev.08) — Apps Script 後端
 * 綁定在 Google Sheet 上（擴充功能 → Apps Script），部署為「網頁應用程式」。
 *
 * 部署前請先修改下面兩組密碼。
 */
const PASSWORD = '請改成填表密碼';        // 手機填表用
const ADMIN_PASSWORD = '請改成管理密碼';  // 上傳 Excel 覆蓋資料用
const SHEET_NAME = '屏東廠藥品品質抽查資料'; // 獨立專案時自動建立的試算表名稱

/** 取得資料試算表：綁定在試算表上就用該試算表；獨立專案則第一次自動建立並記住 ID。 */
function ss_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const act = SpreadsheetApp.getActiveSpreadsheet();
  if (act) return act;
  const created = SpreadsheetApp.create(SHEET_NAME);
  props.setProperty('SHEET_ID', created.getId());
  return created;
}

const SHEETS = {
  POOL: 'Sheet3',          // 客戶自備藥母體
  CHECKED_CUST: 'Sheet1',  // 已查過客戶藥
  CHECKED_COMP: '工作表1',  // 已查過公司藥
  LAB: 'Sheet5',           // 內部檢驗判定
  MASTER: '公司藥主檔',      // 公司藥主檔（藥名、藥廠廠商、是否啟用）
  RECORDS: '抽查紀錄'        // 每週週報
};
const UPLOADABLE = [SHEETS.POOL, SHEETS.CHECKED_CUST, SHEETS.CHECKED_COMP, SHEETS.LAB, SHEETS.MASTER];
const MASTER_HEADER = ['藥名', '藥廠廠商', '是否啟用'];
const RECORD_HEADER = ['抽查日期', '週次代碼', '不合格數', '總數', '內容JSON', '更新時間'];
const BACKUP_KEEP = 3;

/* ---------------- 入口 ---------------- */

function doGet() {
  return json_({ ok: true, msg: '藥品品質抽查表 API 運作中' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    const admin = req.action === 'upload';
    const pw = String(req.pw || '');
    if (admin ? pw !== ADMIN_PASSWORD : (pw !== PASSWORD && pw !== ADMIN_PASSWORD)) {
      return json_({ ok: false, error: admin ? '管理密碼錯誤' : '密碼錯誤' });
    }
    switch (req.action) {
      case 'ping': return json_({ ok: true });
      case 'getData': return json_({ ok: true, data: getData_() });
      case 'getReport': return json_({ ok: true, report: getReport_(req.date) });
      case 'saveReport': return json_(withLock_(() => saveReport_(req.report)));
      case 'deleteReport': return json_(withLock_(() => deleteReport_(req.date)));
      case 'upload': return json_(withLock_(() => upload_(req.sheets || {})));
      default: return json_({ ok: false, error: '未知的 action：' + req.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------------- 初始化（第一次手動執行一次） ---------------- */

function setup() {
  const ss = ss_();
  [SHEETS.POOL, SHEETS.CHECKED_CUST, SHEETS.CHECKED_COMP, SHEETS.LAB].forEach(n => {
    if (!ss.getSheetByName(n)) ss.insertSheet(n);
  });
  let m = ss.getSheetByName(SHEETS.MASTER);
  if (!m) { m = ss.insertSheet(SHEETS.MASTER); }
  if (m.getLastRow() === 0) m.getRange(1, 1, 1, 3).setValues([MASTER_HEADER]).setFontWeight('bold');
  let r = ss.getSheetByName(SHEETS.RECORDS);
  if (!r) { r = ss.insertSheet(SHEETS.RECORDS); }
  if (r.getLastRow() === 0) r.getRange(1, 1, 1, RECORD_HEADER.length).setValues([RECORD_HEADER]).setFontWeight('bold');
  Logger.log('資料試算表：' + ss.getUrl());
  return ss.getUrl();
}

/* ---------------- 讀取 ---------------- */

function values_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastRow() === 0) return [];
  return sh.getRange(1, 1, sh.getLastRow(), Math.max(sh.getLastColumn(), 1)).getDisplayValues();
}

function colIndex_(header, names, fallback) {
  for (const n of names) {
    const i = header.findIndex(h => String(h).trim() === n);
    if (i >= 0) return i;
  }
  return fallback;
}

function getData_() {
  // Sheet3：客戶自備藥母體
  const pv = values_(SHEETS.POOL);
  const pool = [];
  if (pv.length > 1) {
    const h = pv[0];
    const iCust = colIndex_(h, ['客戶姓名'], 0);
    const iName = colIndex_(h, ['藥名'], 2);
    const iQty = colIndex_(h, ['庫存量'], 3);
    const iExp = colIndex_(h, ['有效日期'], 4);
    const iSup = colIndex_(h, ['供應商'], 12);
    for (let r = 1; r < pv.length; r++) {
      const row = pv[r];
      const name = String(row[iName] || '').trim();
      if (!name) continue;
      pool.push({
        name: name,
        cust: String(row[iCust] || '').trim(),
        qty: parseFloat(String(row[iQty]).replace(/,/g, '')) || 0,
        exp: String(row[iExp] || '').trim(),
        sup: String(row[iSup] || '').trim()
      });
    }
  }
  // Sheet1：已查過客戶藥（B=藥名, C=客戶）
  const checkedCust = values_(SHEETS.CHECKED_CUST)
    .map(r => [String(r[1] || '').trim(), String(r[2] || '').trim()])
    .filter(p => p[0]);
  // 工作表1：已查過公司藥（B=藥名）
  const checkedComp = values_(SHEETS.CHECKED_COMP)
    .map(r => String(r[1] || '').trim()).filter(Boolean);
  // 公司藥主檔
  const mv = values_(SHEETS.MASTER);
  const master = [];
  for (let r = 1; r < mv.length; r++) {
    const name = String(mv[r][0] || '').trim();
    if (!name) continue;
    const act = String(mv[r][2] || '').trim();
    master.push({ name: name, maker: String(mv[r][1] || '').trim(), active: !/^(N|否|停用|0|FALSE)$/i.test(act) });
  }
  // Sheet5：內部檢驗（品名、判定、判定日期、檢驗項次）
  const lv = values_(SHEETS.LAB);
  const lab = [];
  if (lv.length > 1) {
    const h = lv[0];
    const iN = colIndex_(h, ['品名'], 2), iJ = colIndex_(h, ['判定'], 4);
    const iD = colIndex_(h, ['判定日期'], 5), iI = colIndex_(h, ['檢驗項次'], 6);
    const iM = colIndex_(h, ['廠商'], 3);
    for (let r = 1; r < lv.length; r++) {
      const n = String(lv[r][iN] || '').trim();
      if (!n) continue;
      lab.push([n, String(lv[r][iJ] || '').trim(), String(lv[r][iD] || '').trim(), String(lv[r][iI] || '').trim(), String(lv[r][iM] || '').trim()]);
    }
  }
  // 抽查紀錄摘要
  const rv = values_(SHEETS.RECORDS);
  const records = [];
  for (let r = 1; r < rv.length; r++) {
    if (!rv[r][0]) continue;
    records.push({ date: rv[r][0], code: rv[r][1], fail: Number(rv[r][2]) || 0, total: Number(rv[r][3]) || 0, updated: rv[r][5] });
  }
  return { pool: pool, checkedCust: checkedCust, checkedComp: checkedComp, master: master, lab: lab, records: records };
}

function findRecordRow_(sh, date) {
  if (sh.getLastRow() < 2) return -1;
  const dates = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
  for (let i = 0; i < dates.length; i++) if (dates[i][0] === date) return i + 2;
  return -1;
}

function getReport_(date) {
  const sh = ss_().getSheetByName(SHEETS.RECORDS);
  if (!sh) return null;
  const row = findRecordRow_(sh, String(date));
  if (row < 0) return null;
  return JSON.parse(sh.getRange(row, 5).getValue());
}

/* ---------------- 寫入 ---------------- */

function saveReport_(rep) {
  if (!rep || !/^\d{4}-\d{2}-\d{2}$/.test(rep.date)) return { ok: false, error: '抽查日期格式錯誤' };
  setup();
  const ss = ss_();
  const sh = ss.getSheetByName(SHEETS.RECORDS);
  const items = rep.items || [];
  const fail = items.filter(it => it.verdict === '不合格').length;
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  const row = [rep.date, rep.code || '', fail, items.length, JSON.stringify(rep), now];
  sh.getRange(1, 5, sh.getMaxRows(), 1).setNumberFormat('@');
  const r = findRecordRow_(sh, rep.date);
  sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');
  if (r > 0) sh.getRange(r, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);

  // 已查過清單：只附加尚未存在的品項
  const norm = s => String(s || '').replace(/[\s\-－—_()（）]/g, '').toLowerCase();
  const cust = ss.getSheetByName(SHEETS.CHECKED_CUST);
  const have = new Set(values_(SHEETS.CHECKED_CUST).map(v => norm(v[1]) + '|' + norm(v[2])));
  let addC = 0, addP = 0;
  items.filter(it => it.type === 'cust' && it.name).forEach(it => {
    const k = norm(it.name) + '|' + norm(it.cust);
    if (have.has(k)) return;
    have.add(k);
    appendText_(cust, [cust.getLastRow() === 0 ? '客戶藥' : '', it.name, it.cust, '', it.appearance, it.exp]);
    addC++;
  });
  const comp = ss.getSheetByName(SHEETS.CHECKED_COMP);
  const haveP = new Set(values_(SHEETS.CHECKED_COMP).map(v => norm(v[1])));
  items.filter(it => it.type === 'comp' && it.name).forEach(it => {
    const k = norm(it.name);
    if (haveP.has(k)) return;
    haveP.add(k);
    appendText_(comp, [comp.getLastRow() === 0 ? '公司藥' : '', it.name, '', it.cust, it.appearance, it.exp]);
    addP++;
  });
  return { ok: true, fail: fail, total: items.length, addedCust: addC, addedComp: addP, updated: now };
}

function appendText_(sh, row) {
  const r = sh.getLastRow() + 1;
  const rg = sh.getRange(r, 1, 1, row.length);
  rg.setNumberFormat('@');
  rg.setValues([row.map(v => v == null ? '' : String(v))]);
}

function deleteReport_(date) {
  const sh = ss_().getSheetByName(SHEETS.RECORDS);
  const r = sh ? findRecordRow_(sh, String(date)) : -1;
  if (r < 0) return { ok: false, error: '找不到這一天的週報' };
  sh.deleteRow(r);
  return { ok: true };
}

/* ---------------- 上傳 Excel 覆蓋 ---------------- */

function upload_(sheets) {
  const ss = ss_();
  const stamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmm');
  const done = [];
  UPLOADABLE.forEach(name => {
    const rows = sheets[name];
    if (!rows || !rows.length) return;
    const width = rows.reduce((m, r) => Math.max(m, r.length), 1);
    const data = rows.map(r => {
      const a = r.map(v => v == null ? '' : String(v));
      while (a.length < width) a.push('');
      return a;
    });
    let sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 0) {
      const bak = sh.copyTo(ss).setName('備份_' + name + '_' + stamp);
      bak.hideSheet();
      pruneBackups_(ss, name);
    }
    if (!sh) sh = ss.insertSheet(name);
    sh.clear();
    if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
    if (sh.getMaxRows() < data.length) sh.insertRowsAfter(sh.getMaxRows(), data.length - sh.getMaxRows());
    const rg = sh.getRange(1, 1, data.length, width);
    rg.setNumberFormat('@');
    rg.setValues(data);
    done.push(name + '（' + data.length + ' 列）');
  });
  setup();
  // 公司藥主檔是空的，就用工作表1 的品項建立初稿
  const m = ss.getSheetByName(SHEETS.MASTER);
  let seeded = 0;
  if (m.getLastRow() <= 1) {
    const seen = new Set();
    const rows = [];
    values_(SHEETS.CHECKED_COMP).forEach(v => {
      const n = String(v[1] || '').trim();
      if (!n || seen.has(n)) return;
      seen.add(n);
      rows.push([n, String(v[3] || v[2] || '').trim(), 'Y']);
    });
    if (rows.length) {
      const rg = m.getRange(2, 1, rows.length, 3);
      rg.setNumberFormat('@');
      rg.setValues(rows);
      seeded = rows.length;
    }
  }
  if (!done.length) return { ok: false, error: '檔案裡沒有可更新的分頁（' + UPLOADABLE.join('、') + '）' };
  return { ok: true, updated: done, seededMaster: seeded };
}

function pruneBackups_(ss, name) {
  const prefix = '備份_' + name + '_';
  const baks = ss.getSheets().filter(s => s.getName().indexOf(prefix) === 0)
    .sort((a, b) => a.getName() < b.getName() ? 1 : -1);
  baks.slice(BACKUP_KEEP).forEach(s => ss.deleteSheet(s));
}
