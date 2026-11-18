/* 純邏輯：比對、日期、判定、Excel XML 填寫。瀏覽器與 Node 共用。 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Core = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ---------- 文字 ---------- */
  function toHalf(s) {
    return String(s == null ? '' : s).replace(/[！-～]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/　/g, ' ');
  }
  function norm(s) {
    return toHalf(s).replace(/[\s\-—_()（）\[\]【】.,，、%]/g, '').toLowerCase();
  }
  function bigrams(s) {
    const a = [];
    for (let i = 0; i < s.length - 1; i++) a.push(s.slice(i, i + 2));
    if (s.length === 1) a.push(s);
    return a;
  }
  function similarity(a, b) {
    a = norm(a); b = norm(b);
    if (!a || !b) return 0;
    if (a === b) return 1;
    let bonus = (a.indexOf(b) >= 0 || b.indexOf(a) >= 0) ? 0.3 : 0;
    const A = bigrams(a), B = bigrams(b).slice();
    let hit = 0;
    A.forEach(x => { const i = B.indexOf(x); if (i >= 0) { hit++; B.splice(i, 1); } });
    // 同長度、只差一字（例如 救甘辛 / 救肝辛）
    if (a.length === b.length && a.length >= 2) {
      let diff = 0;
      for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) diff++;
      if (diff === 1) bonus = Math.max(bonus, 0.5);
    }
    return Math.min(0.99, (2 * hit) / (bigrams(a).length + bigrams(b).length) + bonus);
  }

  /* ---------- 日期 ---------- */
  const pad = n => (n < 10 ? '0' : '') + n;
  function parseISO(s) { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ''); return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null; }
  function todayISO() { const d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function isoWeek(iso) {
    const d = parseISO(iso);
    const day = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - day);
    const y = d.getUTCFullYear();
    const w = Math.ceil(((d - Date.UTC(y, 0, 1)) / 86400000 + 1) / 7);
    return { year: y, week: w };
  }
  function reportCode(iso) { const w = isoWeek(iso); return String(w.year - 1911) + pad(w.week); }
  function rocDateText(iso) { const d = parseISO(iso); return (d.getUTCFullYear() - 1911) + '年' + (d.getUTCMonth() + 1) + '月' + d.getUTCDate() + '日'; }
  // 各種效期寫法 → YYYY-MM-DD（無法解析則回傳 ''）
  function expToISO(s) {
    s = toHalf(s).trim();
    let m = /^(\d{4})(\d{2})(\d{2})$/.exec(s) || /^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})$/.exec(s);
    if (m) return m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]);
    m = /^(\d{2,3})[.\/-](\d{1,2})[.\/-](\d{1,2})$/.exec(s); // 民國
    if (m) return (+m[1] + 1911) + '-' + pad(+m[2]) + '-' + pad(+m[3]);
    return '';
  }
  function expText(iso) { return iso ? iso.replace(/-/g, '.') : ''; }
  // 字串比較即可（YYYY-MM-DD）；注意 2027.04.31 這類不存在日期仍照字串比
  function isExpired(expISO, dateISO) { return !!expISO && expISO < dateISO; }

  /* ---------- Sheet5 檢驗比對 ---------- */
  function buildLabIndex(lab) {
    const map = new Map();
    lab.forEach(r => {
      const key = norm(r[0]);
      if (!key) return;
      const rec = { name: r[0], judge: r[1], date: r[2], items: r[3], maker: r[4] || '' };
      const cur = map.get(key);
      // 有判定的優先；同樣有判定就取判定日期最新
      const score = x => (x.judge ? '1' : '0') + String(x.date || '').padStart(8, '0');
      if (!cur || score(rec) > score(cur)) map.set(key, rec);
    });
    return { map: map, names: Array.from(map.values()) };
  }
  function judgeOK(j) { return /^ok$/i.test(String(j || '').trim()) || String(j).trim() === '合格'; }
  function lookupLab(idx, name) {
    const rec = idx.map.get(norm(name));
    if (rec) return { found: true, rec: rec, pass: rec.judge ? judgeOK(rec.judge) : null };
    const cands = idx.names
      .map(r => ({ rec: r, s: similarity(name, r.name) }))
      .filter(x => x.s >= 0.5)
      .sort((a, b) => b.s - a.s)
      .slice(0, 5)
      .map(x => x.rec);
    return { found: false, rec: null, pass: null, candidates: cands };
  }

  /* ---------- 判定 ---------- */
  // item: {type, appearanceOK, exp(ISO), lab:'是'|'否'|''}
  function autoVerdict(item, dateISO) {
    const reasons = [];
    if (item.appearanceOK === false) reasons.push('外觀異常');
    if (isExpired(item.exp, dateISO)) reasons.push('效期已過');
    if (item.type === 'cust' && item.lab !== '是') reasons.push('檢驗報告不合格或查無');
    return { verdict: reasons.length ? '不合格' : '合格', reasons: reasons };
  }

  /* ---------- 當月週次統計 ---------- */
  // records: [{date, fail, total}]；current: {date, fail, total}
  function monthWeeks(records, current) {
    const ym = current.date.slice(0, 7);
    const list = records.filter(r => String(r.date).slice(0, 7) === ym && r.date !== current.date)
      .map(r => ({ date: r.date, fail: +r.fail || 0, total: +r.total || 0 }));
    list.push({ date: current.date, fail: current.fail, total: current.total });
    list.sort((a, b) => (a.date < b.date ? -1 : 1));
    const weeks = list.slice(0, 5);
    const sum = weeks.reduce((s, w) => ({ fail: s.fail + w.fail, total: s.total + w.total }), { fail: 0, total: 0 });
    return { weeks: weeks, fail: sum.fail, total: sum.total, overflow: list.length > 5 };
  }

  /* ---------- Excel XML ---------- */
  function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  // value: string → 文字；number → 數值；{f, v} → 公式+快取值；null/'' → 空白（保留格式）
  function setCells(xml, cells) {
    Object.keys(cells).forEach(ref => {
      const re = new RegExp('<c r="' + ref + '"((?:\\s+[a-zA-Z:]+="[^"]*")*)\\s*(?:/>|>[\\s\\S]*?</c>)');
      const m = re.exec(xml);
      if (!m) throw new Error('範本找不到儲存格 ' + ref);
      const sm = /\ss="(\d+)"/.exec(m[1]);
      const st = sm ? ' s="' + sm[1] + '"' : '';
      const v = cells[ref];
      let out;
      if (v === null || v === undefined || v === '') out = '<c r="' + ref + '"' + st + '/>';
      else if (typeof v === 'number') out = '<c r="' + ref + '"' + st + '><v>' + v + '</v></c>';
      else if (typeof v === 'object') {
        const isStr = typeof v.v === 'string';
        out = '<c r="' + ref + '"' + st + (isStr ? ' t="str"' : '') + '><f>' + esc(v.f) + '</f>' + (v.v === '' || v.v == null ? '' : '<v>' + esc(v.v) + '</v>') + '</c>';
      } else out = '<c r="' + ref + '"' + st + ' t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
      xml = xml.slice(0, m.index) + out + xml.slice(m.index + m[0].length);
    });
    return xml;
  }

  // report: {date, items:[10]} ；month: monthWeeks 結果
  function reportCells(report, month) {
    const a = {};
    a.B3 = '抽查日期： ' + rocDateText(report.date);
    const cust = report.items.filter(i => i.type === 'cust');
    const comp = report.items.filter(i => i.type === 'comp');
    const rows = [];
    for (let i = 0; i < 5; i++) rows.push(cust[i] || null);
    for (let i = 0; i < 5; i++) rows.push(comp[i] || null);
    rows.forEach((it, k) => {
      const r = 6 + k;
      if (!it) { ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(c => { a[c + r] = ''; }); return; }
      a['B' + r] = it.name || '';
      a['C' + r] = it.cust || '';
      a['D' + r] = '';
      a['E' + r] = it.appearanceOK === false ? '外觀異常' : '外觀正常';
      a['F' + r] = expText(it.exp);
      a['G' + r] = it.lab || '';
      a['H' + r] = it.verdict || '';
      const parts = [];
      if (it.type === 'comp') parts.push('公司藥');
      if (it.appearanceOK === false && it.appearanceNote) parts.push(it.appearanceNote);
      if (it.remark) parts.push(it.remark);
      a['I' + r] = parts.join(' ');
    });
    const items = report.items.filter(Boolean);
    const fail = items.filter(i => i.verdict === '不合格').length;
    const total = items.length;
    a.C18 = fail;
    a.C20 = total;
    a.E19 = { f: 'IF(C20=0,0,C18/C20)', v: total ? fail / total : 0 };
    a.H18 = { f: '不合格率!K3', v: month.fail };
    a.H20 = { f: '不合格率!K4', v: month.total };
    a.I19 = { f: 'IF(H20=0,0,H18/H20)', v: month.total ? month.fail / month.total : 0 };
    return a;
  }

  function rateCells(report, month) {
    const a = {};
    const items = report.items.filter(Boolean);
    const fail = items.filter(i => i.verdict === '不合格').length;
    const total = items.length;
    a.C3 = { f: '抽查表!C18', v: fail };
    a.C4 = { f: '抽查表!C20', v: total };
    a.C5 = { f: 'IF(C4=0,0,C3/C4)', v: total ? fail / total : 0 };
    ['F', 'G', 'H', 'I', 'J'].forEach((c, i) => {
      const w = month.weeks[i];
      if (w) {
        a[c + '3'] = w.fail;
        a[c + '4'] = w.total;
        a[c + '5'] = { f: 'IF(' + c + '4=0,0,' + c + '3/' + c + '4)', v: w.total ? w.fail / w.total : 0 };
      } else { a[c + '3'] = ''; a[c + '4'] = ''; a[c + '5'] = ''; }
    });
    a.K3 = { f: 'SUM(F3:J3)', v: month.fail };
    a.K4 = { f: 'SUM(F4:J4)', v: month.total };
    a.K5 = { f: 'IF(K4=0,0,K3/K4)', v: month.total ? month.fail / month.total : 0 };
    return a;
  }

  // files: {path: string} 只需傳入要改的 XML 文字；回傳修改後的 map 與要刪除的路徑
  function patchWorkbook(files, report, month) {
    const out = {};
    out['xl/worksheets/sheet1.xml'] = setCells(files['xl/worksheets/sheet1.xml'], reportCells(report, month));
    out['xl/worksheets/sheet2.xml'] = setCells(files['xl/worksheets/sheet2.xml'], rateCells(report, month));
    out['xl/workbook.xml'] = files['xl/workbook.xml'].replace(/<calcPr([^>]*?)\s*\/>/, (m, attrs) =>
      '<calcPr' + attrs.replace(/\s*fullCalcOnLoad="[^"]*"/, '') + ' fullCalcOnLoad="1"/>');
    out['xl/_rels/workbook.xml.rels'] = files['xl/_rels/workbook.xml.rels'].replace(/<Relationship [^>]*Target="calcChain\.xml"\/>/, '');
    out['[Content_Types].xml'] = files['[Content_Types].xml'].replace(/<Override [^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, '');
    return { files: out, remove: ['xl/calcChain.xml'] };
  }

  function fileName(iso) { return '屏東廠藥品品質抽查表(Rev.08)' + reportCode(iso) + '.xlsx'; }

  return {
    norm, similarity, todayISO, isoWeek, reportCode, rocDateText, expToISO, expText, isExpired,
    buildLabIndex, lookupLab, judgeOK, autoVerdict, monthWeeks, setCells, reportCells, rateCells, patchWorkbook, fileName
  };
});
