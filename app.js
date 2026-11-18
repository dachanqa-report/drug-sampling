/* 屏東廠藥品品質抽查表 — 前端 */
(function () {
  'use strict';
  const C = window.Core;
  const $ = s => document.querySelector(s);
  const LS = { api: 'dq_api', pw: 'dq_pw', draft: 'dq_draft' };
  const N_CUST = 5, N_COMP = 5;

  const cfg = {
    api: localStorage.getItem(LS.api) || (window.APP_CONFIG && window.APP_CONFIG.API_URL) || '',
    pw: localStorage.getItem(LS.pw) || ''
  };
  let data = null;         // 後端資料
  let labIdx = { map: new Map(), names: [] };
  let checkedCust = new Set(), checkedComp = new Set();
  let report = newReport(C.todayISO());

  /* ---------------- 工具 ---------------- */
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
  function toast(msg, ms) {
    const t = $('#toast'); t.textContent = msg; t.classList.remove('hidden');
    clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.add('hidden'), ms || 2600);
  }
  function status(msg, isErr) {
    const s = $('#status');
    if (!msg) { s.classList.add('hidden'); return; }
    s.innerHTML = msg; s.className = 'status' + (isErr ? ' err' : '');
  }
  const ckey = (name, cust) => C.norm(name) + '|' + C.norm(cust);
  function shuffle(a) { for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; }

  async function api(action, extra, pwOverride) {
    if (!cfg.api) throw new Error('尚未設定 API 網址，請按右上角「設定」');
    const res = await fetch(cfg.api, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(Object.assign({ action: action, pw: pwOverride != null ? pwOverride : cfg.pw }, extra || {}))
    });
    const j = await res.json();
    if (!j.ok) throw new Error(j.error || '伺服器錯誤');
    return j;
  }

  /* ---------------- 週報資料結構 ---------------- */
  function blankItem(type) {
    return { type: type, name: '', cust: '', qty: null, sup: '', appearanceOK: true, appearanceNote: '', exp: '',
      labRec: null, labCands: [], labState: '', lab: '', labManual: false, verdict: '合格', verdictManual: false, remark: '' };
  }
  function newReport(date) {
    const items = [];
    for (let i = 0; i < N_CUST; i++) items.push(blankItem('cust'));
    for (let i = 0; i < N_COMP; i++) items.push(blankItem('comp'));
    return { date: date, items: items };
  }
  function saveDraft() { try { localStorage.setItem(LS.draft, JSON.stringify(report)); } catch (e) { /* ignore */ } }

  // 依藥名到 Sheet5 比對檢驗報告
  function resolveLab(it) {
    it.labRec = null; it.labCands = []; it.labManual = false;
    if (!it.name) { it.labState = ''; it.lab = ''; return; }
    const r = C.lookupLab(labIdx, it.name);
    if (r.found) {
      it.labRec = r.rec; it.labState = 'found';
      it.lab = r.pass === true ? '是' : (it.type === 'cust' ? '否' : (r.pass === false ? '否' : ''));
    } else if (r.candidates && r.candidates.length) {
      it.labCands = r.candidates; it.labState = 'pending'; it.lab = '';
    } else {
      it.labState = 'none'; it.lab = it.type === 'cust' ? '否' : '';
    }
  }
  function pickCandidate(it, rec) {
    it.labRec = rec; it.labCands = []; it.labState = 'similar'; it.labManual = false;
    it.lab = C.judgeOK(rec.judge) ? '是' : (it.type === 'cust' ? '否' : (rec.judge ? '否' : ''));
  }
  function recompute(it) {
    const a = C.autoVerdict(it, report.date);
    it.autoVerdict = a.verdict; it.reasons = a.reasons;
    if (!it.verdictManual) it.verdict = a.verdict;
    else if (it.verdict === a.verdict) it.verdictManual = false;
  }
  function setItem(idx, src) {
    const it = blankItem(report.items[idx].type);
    Object.assign(it, src);
    resolveLab(it); recompute(it);
    report.items[idx] = it;
  }

  /* ---------------- 載入資料 ---------------- */
  async function loadData() {
    status('資料載入中…');
    try {
      const j = await api('getData');
      data = j.data;
      labIdx = C.buildLabIndex(data.lab || []);
      checkedCust = new Set(data.checkedCust.map(p => ckey(p[0], p[1])));
      checkedComp = new Set(data.checkedComp.map(n => C.norm(n)));
      status('');
      const notes = [];
      if (!data.pool.length) notes.push('Sheet3（客戶藥母體）是空的');
      if (!data.lab.length) notes.push('Sheet5（檢驗判定）是空的');
      if (!data.master.length) notes.push('公司藥主檔是空的');
      if (notes.length) status(notes.join('、') + '。請按右上角「上傳」匯入 Excel。');
      // 重新比對草稿
      report.items.forEach(it => { if (it.name && !it.labManual) { const keep = { lab: it.lab, labState: it.labState, labRec: it.labRec }; resolveLab(it); if (keep.labState === 'similar') Object.assign(it, keep); } recompute(it); });
      renderAll();
    } catch (e) {
      data = null;
      status('無法載入資料：' + esc(e.message), true);
      renderAll();
    }
  }

  /* ---------------- 畫面：日期 ---------------- */
  function renderDate() {
    $('#date').value = report.date;
    $('#weekChip').textContent = C.reportCode(report.date) + ' 週';
    const rec = data && data.records.find(r => r.date === report.date);
    $('#dateHint').textContent = C.rocDateText(report.date) + (rec ? ' · 這一天已有週報（最後更新 ' + rec.updated + '）' : '');
  }

  /* ---------------- 畫面：品項卡片 ---------------- */
  function seg(name, opts, cur, idx) {
    return '<div class="seg">' + opts.map(o => '<button type="button" data-act="' + name + '" data-i="' + idx + '" data-v="' + esc(o.v) + '" class="' +
      (cur === o.v ? 'on ' + (o.cls || '') : '') + '">' + esc(o.t) + '</button>').join('') + '</div>';
  }
  function labHtml(it, i) {
    let info = '';
    if (!it.name) return '<span class="lab none">選好藥品後自動比對 Sheet5</span>';
    const r = it.labRec;
    const recTxt = r ? esc(r.name) + '｜' + esc(r.judge || '無判定') + (r.date ? '｜' + esc(r.date) : '') + (r.items ? '<br>' + esc(r.items) : '') : '';
    if (it.labState === 'found') info = '<div class="lab ' + (C.judgeOK(r.judge) ? 'ok' : 'bad') + '">Sheet5：' + recTxt + '</div>';
    else if (it.labState === 'similar') info = '<div class="lab warn">近似比對（已確認）：' + recTxt + '</div>';
    else if (it.labState === 'pending') {
      info = '<div class="lab warn">Sheet5 查無「' + esc(it.name) + '」，是不是下面其中一個？' +
        '<div class="chips">' + it.labCands.map((c, k) => '<button type="button" class="chip" data-act="cand" data-i="' + i + '" data-k="' + k + '">' +
          esc(c.name) + '（' + esc(c.judge || '無判定') + '）</button>').join('') +
        '<button type="button" class="chip" data-act="nocand" data-i="' + i + '">都不是</button></div></div>';
    } else if (it.labState === 'none') info = '<div class="lab none">Sheet5 查無檢驗紀錄' + (it.type === 'cust' ? '，視為「否」' : '，留白') + '</div>';
    const opts = it.type === 'cust' ? [{ v: '是', t: '是', cls: 'okc' }, { v: '否', t: '否', cls: 'badc' }]
      : [{ v: '是', t: '是', cls: 'okc' }, { v: '否', t: '否', cls: 'badc' }, { v: '', t: '留白' }];
    const segHtml = it.labState === 'pending' ? '' : seg('lab', opts, it.lab, i) + (it.labManual ? ' <span class="manual">手動</span>' : '');
    return '<div class="stack">' + segHtml + info + '</div>';
  }
  // 下拉選單：未查過 / ✓已查過 / 過期或無庫存（客戶藥），本週其他欄已選的不能再選
  function selectHtml(it, i) {
    if (!data) return '';
    const used = usedKeys(i);
    const isCust = it.type === 'cust';
    const curKey = it.name ? (isCust ? ckey(it.name, it.cust) : C.norm(it.name)) : '';
    let found = false;
    const groups = isCust ? { a: [], b: [], c: [] } : { a: [], b: [] };
    const list = isCust ? data.pool : data.master;
    list.forEach((p, k) => {
      if (!isCust && !p.active) return;
      const key = isCust ? ckey(p.name, p.cust) : C.norm(p.name);
      const done = isCust ? checkedCust.has(key) : checkedComp.has(key);
      const sel = key === curKey && !found;
      if (sel) found = true;
      const inuse = used.has(key) && !sel;
      let bad = '';
      if (isCust) {
        const e = C.expToISO(p.exp);
        if (!e || e < report.date) bad = '過期';
        else if (!(p.qty > 0)) bad = '無庫存';
      }
      const label = (done ? '✓ ' : '') + p.name + '｜' + (isCust ? (p.cust || '—') + '｜效期 ' + (C.expText(C.expToISO(p.exp)) || p.exp || '—') : (p.maker || '—')) +
        (bad ? '（' + bad + '）' : '') + (inuse ? '（本週已選）' : '');
      const opt = '<option value="' + k + '"' + (sel ? ' selected' : '') + (inuse ? ' disabled' : '') + '>' + esc(label) + '</option>';
      (bad ? groups.c : done ? groups.b : groups.a).push(opt);
    });
    const g = (t, a) => a.length ? '<optgroup label="' + t + '（' + a.length + '）">' + a.join('') + '</optgroup>' : '';
    return '<div class="pick-row"><select class="pick-sel" data-act="sel" data-i="' + i + '">' +
      '<option value=""' + (it.name ? '' : ' selected') + '>— 下拉選擇' + (isCust ? '客戶藥' : '公司藥') + '（✓ = 已查過）—</option>' +
      (it.name && !found ? '<option value="" selected>' + esc(it.name + '（手動輸入）') + '</option>' : '') +
      g('未查過', groups.a) + g('✓ 已查過', groups.b) + (isCust ? g('過期或無庫存', groups.c) : '') + '</select></div>';
  }
  function itemHtml(it, i) {
    const no = it.type === 'cust' ? i + 1 : i - N_CUST + 1;
    const isBad = it.name && it.verdict === '不合格';
    let sub = '';
    if (it.name) {
      sub = it.type === 'cust'
        ? '客戶：' + esc(it.cust || '—') + (it.qty != null ? ' · 庫存 ' + esc(it.qty) : '') + (it.sup ? ' · ' + esc(it.sup) : '') +
          (checkedCust.has(ckey(it.name, it.cust)) ? ' · <b>✓已查過</b>' : '')
        : '藥廠廠商：' + esc(it.cust || '—') + (checkedComp.has(C.norm(it.name)) ? ' · <b>✓已查過</b>' : '');
    }
    const custField = it.type === 'comp'
      ? '<div class="field"><label>藥廠廠商</label><input type="text" data-act="cust" data-i="' + i + '" value="' + esc(it.cust) + '"></div>'
      : '<div class="field"><label>客戶</label><input type="text" data-act="cust" data-i="' + i + '" value="' + esc(it.cust) + '"></div>';
    return '<div class="item' + (isBad ? ' bad' : '') + '" id="item' + i + '">' +
      '<div class="item-top"><div class="item-no">' + no + '</div><div class="item-main">' +
      (it.name ? '<div class="item-name">' + esc(it.name) + '</div><div class="item-sub">' + sub + '</div>'
        : '<div class="item-name empty">尚未選擇藥品</div>') +
      '</div><div class="item-btns"><button type="button" class="mini" data-act="rand" data-i="' + i + '">隨機換</button>' +
      '<button type="button" class="mini" data-act="pick" data-i="' + i + '">搜尋</button></div></div>' +
      selectHtml(it, i) +
      (it.name ? (
        custField +
        '<div class="field"><label>外觀</label><div>' + seg('app', [{ v: 'ok', t: '外觀正常', cls: 'okc' }, { v: 'bad', t: '異常', cls: 'badc' }], it.appearanceOK ? 'ok' : 'bad', i) + '</div></div>' +
        (it.appearanceOK ? '' : '<div class="field"><label>異常說明</label><input type="text" data-act="appnote" data-i="' + i + '" value="' + esc(it.appearanceNote) + '" placeholder="必填，例如：結塊、破袋"></div>') +
        '<div class="field"><label>效期</label><input type="date" data-act="exp" data-i="' + i + '" value="' + esc(it.exp) + '">' + '</div>' +
        (it.type === 'comp' && !it.exp ? '<div class="field"><label></label><span class="manual">查核後請輸入有效日期</span></div>' : '') +
        (C.isExpired(it.exp, report.date) ? '<div class="field"><label></label><span class="reason">效期已過抽查日期</span></div>' : '') +
        '<div class="field"><label>檢驗報告</label>' + labHtml(it, i) + '</div>' +
        '<div class="field"><label>判定</label><div>' + seg('verdict', [{ v: '合格', t: '合格', cls: 'okc' }, { v: '不合格', t: '不合格', cls: 'badc' }], it.verdict, i) +
          (it.verdictManual ? ' <span class="manual">手動修改（自動判定：' + esc(it.autoVerdict) + '）</span>' : '') +
          (!it.verdictManual && it.reasons && it.reasons.length ? '<div class="reason">' + esc(it.reasons.join('、')) + '</div>' : '') + '</div></div>' +
        '<div class="field"><label>備註</label><input type="text" data-act="remark" data-i="' + i + '" value="' + esc(it.remark) + '" placeholder="' + (it.verdictManual ? '手動改判定，請寫原因（必填）' : '選填') + '"></div>'
      ) : '') + '</div>';
  }
  function renderItems() {
    $('#custList').innerHTML = report.items.slice(0, N_CUST).map((it, k) => itemHtml(it, k)).join('');
    $('#compList').innerHTML = report.items.slice(N_CUST).map((it, k) => itemHtml(it, k + N_CUST)).join('');
  }
  function renderItem(i) {
    const el = document.getElementById('item' + i);
    if (!el || !el.parentNode) return renderItems();
    const tmp = document.createElement('div'); tmp.innerHTML = itemHtml(report.items[i], i);
    try { el.replaceWith(tmp.firstChild); } catch (e) { renderItems(); }
  }

  /* ---------------- 畫面：彙總 ---------------- */
  function validate() {
    const errs = [];
    report.items.forEach((it, i) => {
      const label = (it.type === 'cust' ? '客戶藥 ' + (i + 1) : '公司藥 ' + (i - N_CUST + 1));
      if (!it.name) { errs.push(label + '：尚未選擇藥品'); return; }
      if (!it.exp) errs.push(label + '：請輸入效期');
      if (!it.appearanceOK && !it.appearanceNote.trim()) errs.push(label + '：外觀異常請填說明');
      if (it.labState === 'pending') errs.push(label + '：請確認 Sheet5 近似品名');
      if (it.verdictManual && !it.remark.trim()) errs.push(label + '：手動修改判定，請在備註寫原因');
    });
    return errs;
  }
  function stats() {
    const filled = report.items.filter(it => it.name);
    const fail = filled.filter(it => it.verdict === '不合格').length;
    const total = report.items.length;
    const month = C.monthWeeks(data ? data.records : [], { date: report.date, fail: fail, total: total });
    return { fail: fail, total: total, month: month };
  }
  const pct = (a, b) => (b ? (a / b * 100).toFixed(2) : '0.00') + '%';
  function renderSummary() {
    const s = stats();
    const rows = report.items.map((it, i) => '<tr><td>' + (it.type === 'cust' ? '客' : '公') + (it.type === 'cust' ? i + 1 : i - N_CUST + 1) + '</td><td>' +
      (it.name ? esc(it.name) : '<span class="badge grey">未選</span>') + '</td><td>' + esc(it.cust || '') + '</td><td class="c">' + esc(it.lab || '') + '</td><td class="c">' +
      (it.name ? '<span class="badge ' + (it.verdict === '不合格' ? 'bad' : 'ok') + '">' + esc(it.verdict) + '</span>' : '') + '</td></tr>').join('');
    const weekTxt = s.month.weeks.map((w, k) => '第' + '一二三四五'[k] + '週（' + w.date.slice(5).replace('-', '/') + '）' + w.fail + '/' + w.total).join('、');
    const errs = validate();
    $('#summary').innerHTML =
      '<table class="sum"><thead><tr><th></th><th>藥品</th><th>客戶／廠商</th><th class="c">檢驗</th><th class="c">判定</th></tr></thead><tbody>' + rows + '</tbody></table>' +
      '<div class="rates"><div class="rate"><div class="k">本週不合格率</div><div class="v">' + pct(s.fail, s.total) + '</div><div class="s">' + s.fail + ' / ' + s.total + '</div></div>' +
      '<div class="rate"><div class="k">本月累計不合格率</div><div class="v">' + pct(s.month.fail, s.month.total) + '</div><div class="s">' + s.month.fail + ' / ' + s.month.total + '</div></div></div>' +
      '<div class="weeks">' + esc(weekTxt) + (s.month.overflow ? '（本月超過 5 次抽查，只計前 5 次）' : '') + '</div>' +
      (errs.length ? '<div class="errors">送出前還需要處理：<ul>' + errs.map(e => '<li>' + esc(e) + '</li>').join('') + '</ul></div>' : '');
    $('#btnSave').disabled = !!errs.length || !data;
  }
  function renderAll() { renderDate(); renderItems(); renderSummary(); }
  function changed(i) { if (i != null) { recompute(report.items[i]); renderItem(i); } renderSummary(); saveDraft(); }

  /* ---------------- 隨機抽 ---------------- */
  function usedKeys(except) {
    const s = new Set();
    report.items.forEach((it, k) => { if (k !== except && it.name) s.add(it.type === 'cust' ? ckey(it.name, it.cust) : C.norm(it.name)); });
    return s;
  }
  function eligibleCust(except) {
    const used = usedKeys(except);
    return data.pool.filter(p => {
      const e = C.expToISO(p.exp);
      return p.qty > 0 && e && e >= report.date && !checkedCust.has(ckey(p.name, p.cust)) && !used.has(ckey(p.name, p.cust));
    });
  }
  function eligibleComp(except) {
    const used = usedKeys(except);
    const act = data.master.filter(m => m.active && !used.has(C.norm(m.name)));
    const fresh = act.filter(m => !checkedComp.has(C.norm(m.name)));
    return { list: fresh.length ? fresh : act, allChecked: !fresh.length };
  }
  function custFromPool(p) { return { name: p.name, cust: p.cust, qty: p.qty, sup: p.sup, exp: C.expToISO(p.exp) }; }
  function compFromMaster(m) { return { name: m.name, cust: m.maker, exp: '' }; }

  function randomOne(i) {
    if (!data) return toast('資料尚未載入');
    const it = report.items[i];
    if (it.type === 'cust') {
      const el = eligibleCust(i);
      if (!el.length) return toast('沒有符合條件（有效、有庫存、未查過）的客戶藥');
      setItem(i, custFromPool(el[Math.floor(Math.random() * el.length)]));
    } else {
      const r = eligibleComp(i);
      if (!r.list.length) return toast('公司藥主檔沒有可用的藥品');
      if (r.allChecked) toast('主檔的公司藥都查過了，從全部啟用品項中抽');
      setItem(i, compFromMaster(r.list[Math.floor(Math.random() * r.list.length)]));
    }
    renderItems(); changed();
  }
  function randomAll(type) {
    if (!data) return toast('資料尚未載入');
    const idxs = report.items.map((it, k) => k).filter(k => report.items[k].type === type);
    if (idxs.some(k => report.items[k].name) && !confirm('會覆蓋目前已選的 ' + (type === 'cust' ? '客戶藥' : '公司藥') + '，確定嗎？')) return;
    idxs.forEach(k => { report.items[k] = blankItem(type); });
    if (type === 'cust') {
      const el = shuffle(eligibleCust(-1).slice());
      // 盡量不同藥名
      const picked = [], names = new Set();
      el.forEach(p => { if (picked.length < idxs.length && !names.has(C.norm(p.name))) { picked.push(p); names.add(C.norm(p.name)); } });
      el.forEach(p => { if (picked.length < idxs.length && picked.indexOf(p) < 0) picked.push(p); });
      picked.forEach((p, k) => setItem(idxs[k], custFromPool(p)));
      if (picked.length < idxs.length) toast('符合條件的客戶藥只有 ' + picked.length + ' 筆');
    } else {
      const r = eligibleComp(-1);
      const picked = shuffle(r.list.slice()).slice(0, idxs.length);
      picked.forEach((m, k) => setItem(idxs[k], compFromMaster(m)));
      if (r.allChecked) toast('主檔的公司藥都查過了，從全部啟用品項中抽');
      else if (picked.length < idxs.length) toast('未查過的公司藥只有 ' + picked.length + ' 筆');
    }
    renderItems(); changed();
  }

  /* ---------------- 選擇器 ---------------- */
  function openModal(title, html) { $('#modalTitle').textContent = title; $('#modalBody').innerHTML = html; $('#modal').classList.remove('hidden'); document.body.style.overflow = 'hidden'; }
  function closeModal() { $('#modal').classList.add('hidden'); document.body.style.overflow = ''; }

  function openPicker(i) {
    if (!data) return toast('資料尚未載入');
    const it = report.items[i];
    const isCust = it.type === 'cust';
    openModal((isCust ? '選擇客戶藥 ' + (i + 1) : '選擇公司藥 ' + (i - N_CUST + 1)),
      '<input class="search" id="pkQ" placeholder="搜尋藥名' + (isCust ? '或客戶' : '或廠商') + '" autocomplete="off">' +
      '<div class="filters">' + (isCust ? '<label><input type="checkbox" id="pkValid" checked> 只看有效、有庫存</label>' : '<label><input type="checkbox" id="pkValid" checked> 只看啟用</label>') +
      '<label><input type="checkbox" id="pkHide"> 隱藏已查過</label></div><div id="pkList"></div>' +
      '<details style="margin-top:12px"><summary>找不到？手動輸入</summary><div class="form" style="margin-top:8px">' +
      '<label>藥名</label><input id="pkName"><label>' + (isCust ? '客戶' : '藥廠廠商') + '</label><input id="pkCust">' +
      '<button class="btn primary" id="pkManual" type="button">使用手動輸入</button></div></details>');
    const used = usedKeys(i);
    const draw = () => {
      const q = C.norm($('#pkQ').value);
      const valid = $('#pkValid').checked, hide = $('#pkHide').checked;
      let rows;
      if (isCust) {
        rows = data.pool.map((p, k) => {
          const e = C.expToISO(p.exp);
          return { k: k, p: p, done: checkedCust.has(ckey(p.name, p.cust)), inuse: used.has(ckey(p.name, p.cust)), expired: !e || e < report.date, empty: !(p.qty > 0) };
        }).filter(r => (!q || C.norm(r.p.name + r.p.cust).indexOf(q) >= 0) && (!valid || (!r.expired && !r.empty)) && (!hide || !r.done));
      } else {
        rows = data.master.map((m, k) => ({ k: k, p: m, done: checkedComp.has(C.norm(m.name)), inuse: used.has(C.norm(m.name)) }))
          .filter(r => (!q || C.norm(r.p.name + r.p.maker).indexOf(q) >= 0) && (!valid || r.p.active) && (!hide || !r.done));
      }
      rows.sort((a, b) => (a.done - b.done) || (a.inuse - b.inuse));
      const shown = rows.slice(0, 200);
      $('#pkList').innerHTML = (shown.length ? '' : '<p class="hint">沒有符合的品項</p>') + shown.map(r => {
        const p = r.p;
        const d = isCust ? esc(p.cust) + ' · 效期 ' + esc(C.expText(C.expToISO(p.exp)) || p.exp || '—') + ' · 庫存 ' + esc(p.qty) : esc(p.maker || '—') + (p.active ? '' : ' · 停用');
        return '<div class="opt' + (r.done ? ' done' : '') + (r.inuse ? ' inuse' : '') + '" data-k="' + r.k + '"><div class="on"><div class="n">' + esc(p.name) + '</div><div class="d">' + d + '</div></div>' +
          (r.done ? '<span class="tag done">✓已查</span>' : '') + (r.expired ? '<span class="tag exp">過期</span>' : '') + (r.inuse ? '<span class="tag">本週已選</span>' : '') + '</div>';
      }).join('') + (rows.length > 200 ? '<p class="hint">只顯示前 200 筆，請用搜尋縮小範圍</p>' : '');
    };
    draw();
    $('#pkQ').addEventListener('input', draw);
    $('#pkValid').addEventListener('change', draw);
    $('#pkHide').addEventListener('change', draw);
    $('#pkList').addEventListener('click', e => {
      const o = e.target.closest('.opt'); if (!o || o.classList.contains('inuse')) return;
      const k = +o.dataset.k;
      setItem(i, isCust ? custFromPool(data.pool[k]) : compFromMaster(data.master[k]));
      closeModal(); renderItems(); changed();
    });
    $('#pkManual').addEventListener('click', () => {
      const name = $('#pkName').value.trim(); if (!name) return toast('請輸入藥名');
      setItem(i, { name: name, cust: $('#pkCust').value.trim(), exp: '' });
      closeModal(); renderItem(i); changed();
    });
    setTimeout(() => $('#pkQ').focus(), 50);
  }

  /* ---------------- 卡片事件 ---------------- */
  function onClick(e) {
    const b = e.target.closest('[data-act]'); if (!b || b.tagName === 'INPUT') return;
    const i = +b.dataset.i, it = report.items[i], act = b.dataset.act;
    if (act === 'pick') return openPicker(i);
    if (act === 'rand') return randomOne(i);
    if (act === 'app') { it.appearanceOK = b.dataset.v === 'ok'; }
    else if (act === 'lab') { it.lab = b.dataset.v; it.labManual = true; }
    else if (act === 'verdict') { it.verdict = b.dataset.v; it.verdictManual = it.verdict !== C.autoVerdict(it, report.date).verdict; }
    else if (act === 'cand') { pickCandidate(it, it.labCands[+b.dataset.k]); }
    else if (act === 'nocand') { it.labCands = []; it.labState = 'none'; it.lab = it.type === 'cust' ? '否' : ''; }
    else return;
    changed(i);
  }
  function onChange(e) {
    const t = e.target; if (!t.dataset || !t.dataset.act) return;
    const i = +t.dataset.i, it = report.items[i], act = t.dataset.act;
    if (act === 'sel') {
      if (t.value === '') return;
      const k = +t.value;
      setItem(i, it.type === 'cust' ? custFromPool(data.pool[k]) : compFromMaster(data.master[k]));
      renderItems(); return changed();
    }
    if (act === 'exp') { it.exp = t.value; return changed(i); }
    if (act === 'cust') it.cust = t.value.trim();
    else if (act === 'appnote') it.appearanceNote = t.value.trim();
    else if (act === 'remark') it.remark = t.value.trim();
    else return;
    // 文字欄位不重繪卡片，避免點按鈕時被 blur 重繪吃掉
    renderSummary(); saveDraft();
  }

  /* ---------------- Excel ---------------- */
  async function buildExcel(rep) {
    const res = await fetch('template.xlsx', { cache: 'no-cache' });
    if (!res.ok) throw new Error('讀不到 Excel 範本');
    const zip = await JSZip.loadAsync(await res.arrayBuffer());
    const paths = ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'];
    const files = {};
    for (const p of paths) files[p] = await zip.file(p).async('string');
    const fail = rep.items.filter(it => it.name && it.verdict === '不合格').length;
    const month = C.monthWeeks(data ? data.records : [], { date: rep.date, fail: fail, total: rep.items.length });
    const out = C.patchWorkbook(files, rep, month);
    Object.keys(out.files).forEach(p => zip.file(p, out.files[p]));
    out.remove.forEach(p => zip.remove(p));
    const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = C.fileName(rep.date);
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }
  function serial(rep) {
    return {
      date: rep.date, code: C.reportCode(rep.date),
      items: rep.items.map(it => ({
        type: it.type, name: it.name, cust: it.cust, qty: it.qty, sup: it.sup, appearanceOK: it.appearanceOK,
        appearance: it.appearanceOK ? '外觀正常' : '外觀異常', appearanceNote: it.appearanceNote, exp: it.exp, expText: C.expText(it.exp),
        labRec: it.labRec, labState: it.labState === 'pending' ? 'none' : it.labState, lab: it.lab, labManual: it.labManual,
        verdict: it.verdict, verdictManual: it.verdictManual, autoVerdict: it.autoVerdict, remark: it.remark
      }))
    };
  }
  function loadReport(rep) {
    report = { date: rep.date, items: rep.items.map(x => Object.assign(blankItem(x.type), x, { labCands: [] })) };
    report.items.forEach(recompute);
    renderAll(); saveDraft();
  }

  async function onSave() {
    const errs = validate(); if (errs.length) return toast('還有 ' + errs.length + ' 項需要處理');
    const btn = $('#btnSave'); btn.disabled = true; btn.textContent = '送出中…';
    try {
      const rep = serial(report);
      const r = await api('saveReport', { report: rep });
      const rec = { date: rep.date, code: rep.code, fail: r.fail, total: r.total, updated: r.updated };
      data.records = data.records.filter(x => x.date !== rep.date).concat([rec]);
      rep.items.forEach(it => { if (it.type === 'cust') checkedCust.add(ckey(it.name, it.cust)); else checkedComp.add(C.norm(it.name)); });
      await buildExcel(rep);
      toast('已送出（新增已查過：客戶藥 ' + r.addedCust + '、公司藥 ' + r.addedComp + '），Excel 下載中', 4000);
      renderAll();
    } catch (e) { toast('送出失敗：' + e.message, 5000); }
    btn.textContent = '送出並下載 Excel'; renderSummary();
  }
  async function onDownload() {
    if (!report.items.some(it => it.name)) return toast('還沒有任何品項');
    try { await buildExcel(serial(report)); toast('Excel 下載中'); } catch (e) { toast('下載失敗：' + e.message, 5000); }
  }

  /* ---------------- 歷史 ---------------- */
  function openHistory() {
    if (!data) return toast('資料尚未載入');
    const recs = data.records.slice().sort((a, b) => (a.date < b.date ? 1 : -1));
    openModal('歷史週報', recs.length ? recs.map(r => '<div class="list-h"><div><div><b>' + esc(C.rocDateText(r.date)) + '</b> · ' + esc(r.code) + ' 週</div>' +
      '<div class="hint">不合格 ' + r.fail + ' / ' + r.total + ' · 更新 ' + esc(r.updated) + '</div></div><div class="item-btns">' +
      '<button class="mini" data-h="load" data-d="' + esc(r.date) + '">載入</button><button class="mini" data-h="dl" data-d="' + esc(r.date) + '">Excel</button>' +
      '<button class="mini" data-h="del" data-d="' + esc(r.date) + '">刪除</button></div></div>').join('') : '<p class="hint">還沒有週報紀錄</p>');
    $('#modalBody').onclick = async e => {
      const b = e.target.closest('[data-h]'); if (!b) return;
      const d = b.dataset.d;
      try {
        if (b.dataset.h === 'del') {
          if (!confirm('刪除 ' + C.rocDateText(d) + ' 的週報？（已查過清單不會被刪）')) return;
          await api('deleteReport', { date: d });
          data.records = data.records.filter(x => x.date !== d);
          openHistory(); renderAll(); return;
        }
        b.disabled = true;
        const j = await api('getReport', { date: d });
        if (!j.report) throw new Error('找不到週報內容');
        if (b.dataset.h === 'load') { closeModal(); loadReport(j.report); toast('已載入 ' + C.rocDateText(d) + ' 週報，修改後再送出即可覆蓋'); }
        else { await buildExcel(j.report); toast('Excel 下載中'); b.disabled = false; }
      } catch (err) { toast(err.message, 5000); b.disabled = false; }
    };
  }

  /* ---------------- 上傳 Excel ---------------- */
  const UPLOADABLE = ['Sheet3', 'Sheet1', '工作表1', 'Sheet5', '公司藥主檔'];
  function openUpload() {
    openModal('上傳 Excel 更新資料',
      '<div class="form"><p class="hint" style="margin:0">上傳和 11547 相同格式的 Excel。檔案中名為 ' + UPLOADABLE.join('、') +
      ' 的分頁，會<b>整張覆蓋</b> Google Sheet 上對應的分頁（覆蓋前自動備份）。沒有的分頁不動。</p>' +
      '<input type="file" id="upFile" accept=".xlsx,.xls,.xlsm"><div id="upInfo" class="hint"></div>' +
      '<label>管理密碼</label><input type="password" id="upPw" autocomplete="off">' +
      '<button class="btn primary big" id="upGo" disabled>確認上傳覆蓋</button></div>');
    let sheets = null;
    $('#upFile').addEventListener('change', async e => {
      const f = e.target.files[0]; sheets = null; $('#upGo').disabled = true;
      if (!f) return;
      $('#upInfo').textContent = '讀取中…';
      try {
        const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
        sheets = {};
        const found = [];
        UPLOADABLE.forEach(n => {
          if (!wb.Sheets[n]) return;
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: '', blankrows: false });
          // 去掉尾端空白欄
          let w = 0; rows.forEach(r => { for (let c = r.length - 1; c >= 0; c--) if (String(r[c]).trim() !== '') { w = Math.max(w, c + 1); break; } });
          sheets[n] = rows.map(r => r.slice(0, w).map(v => String(v)));
          found.push(n + '（' + rows.length + ' 列）');
        });
        if (!found.length) { $('#upInfo').textContent = '檔案中找不到可更新的分頁。'; sheets = null; return; }
        $('#upInfo').innerHTML = '將更新：' + esc(found.join('、'));
        $('#upGo').disabled = false;
      } catch (err) { $('#upInfo').textContent = '讀取失敗：' + err.message; }
    });
    $('#upGo').addEventListener('click', async () => {
      const pw = $('#upPw').value; if (!pw) return toast('請輸入管理密碼');
      $('#upGo').disabled = true; $('#upGo').textContent = '上傳中…';
      try {
        const r = await api('upload', { sheets: sheets }, pw);
        closeModal();
        toast('已更新：' + r.updated.join('、') + (r.seededMaster ? '；公司藥主檔已用工作表1 建立 ' + r.seededMaster + ' 筆' : ''), 6000);
        await loadData();
      } catch (err) { toast('上傳失敗：' + err.message, 6000); $('#upGo').disabled = false; $('#upGo').textContent = '確認上傳覆蓋'; }
    });
  }

  /* ---------------- 設定 ---------------- */
  function openSettings() {
    openModal('設定',
      '<div class="form"><label>Apps Script 網頁應用程式網址</label><input id="stApi" placeholder="https://script.google.com/macros/s/…/exec" value="' + esc(cfg.api) + '">' +
      '<label>填表密碼</label><input type="password" id="stPw" value="' + esc(cfg.pw) + '">' +
      '<button class="btn primary big" id="stSave">儲存並連線</button><p class="hint">設定只存在這支手機的瀏覽器。</p></div>');
    $('#stSave').addEventListener('click', async () => {
      cfg.api = $('#stApi').value.trim(); cfg.pw = $('#stPw').value;
      localStorage.setItem(LS.api, cfg.api); localStorage.setItem(LS.pw, cfg.pw);
      try { await api('ping'); closeModal(); toast('連線成功'); loadData(); }
      catch (e) { toast('連線失敗：' + e.message, 5000); }
    });
  }

  /* ---------------- 啟動 ---------------- */
  function init() {
    try {
      const d = JSON.parse(localStorage.getItem(LS.draft) || 'null');
      if (d && d.items && d.items.length === N_CUST + N_COMP) report = { date: d.date, items: d.items.map(x => Object.assign(blankItem(x.type), x)) };
    } catch (e) { /* ignore */ }
    $('#main').addEventListener('click', onClick);
    $('#main').addEventListener('change', onChange);
    $('#main').addEventListener('input', e => { if (e.target.dataset && e.target.dataset.act && e.target.type === 'text') onChange(e); });
    $('#date').addEventListener('change', async e => {
      if (!e.target.value) return;
      report.date = e.target.value;
      report.items.forEach(recompute);
      renderAll(); saveDraft();
      const rec = data && data.records.find(r => r.date === report.date);
      if (rec && confirm('這一天已有週報，要載入來修改嗎？')) {
        try { const j = await api('getReport', { date: report.date }); if (j.report) loadReport(j.report); } catch (err) { toast(err.message); }
      }
    });
    $('#btnRandCust').addEventListener('click', () => randomAll('cust'));
    $('#btnRandComp').addEventListener('click', () => randomAll('comp'));
    $('#btnSave').addEventListener('click', onSave);
    $('#btnDownload').addEventListener('click', onDownload);
    $('#btnHistory').addEventListener('click', openHistory);
    $('#btnUpload').addEventListener('click', openUpload);
    $('#btnSettings').addEventListener('click', openSettings);
    $('#modalClose').addEventListener('click', closeModal);
    $('#modal').addEventListener('click', e => { if (e.target.id === 'modal') closeModal(); });
    renderAll();
    if (!cfg.api || !cfg.pw) { status('第一次使用：請按右上角「設定」填入 API 網址和密碼。'); openSettings(); }
    else loadData();
  }
  init();
})();
