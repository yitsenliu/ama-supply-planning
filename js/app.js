(function () {
  'use strict';
  const $ = (s, el = document) => el.querySelector(s);
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const nf = new Intl.NumberFormat('zh-TW', { maximumFractionDigits: 2 });
  const isNum = v => typeof v === 'number' && Number.isFinite(v);
  const n = v => isNum(v) ? nf.format(v) : '<span class="num-null">—</span>';
  const ntext = v => isNum(v) ? nf.format(v) : '—';
  const PAGES = { overview: '總覽與會前重點', plan: '本月產銷', forecast: '次月預估', capacity: '產能排程', bom: 'BOM 齊套', erp: 'ERP 物料查詢', cost: '成本分析', logic: '資料邏輯／ERP 串接', sheets: '原始工作表' };
  const COLS = [
    ['demandExport', '國外需求'], ['demandDomestic', '國內需求'], ['totalDemand', '總需求'], ['salesWeek', '本週實銷'], ['salesTotal', '總實銷'],
    ['stockNorth', '北環倉'], ['stockSouth', '南環倉'], ['stockExport', '外銷倉'], ['stockTotal', '總庫存'],
    ['remainingDemand', '月需求餘量'], ['plannedProduction', '預計生產'], ['outsourced', '託外生產'], ['gap', '產銷差異']
  ];
  const state = { data: null, source: 'encrypted', fileName: '', ui: { plan: {}, forecast: {}, bom: {}, erp: {}, cost: {}, sheets: {}, capacity: 'plan' } };

  // ---------- 共用 ----------
  function toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove('show'), 2600); }
  function download(name, text, type = 'application/json') {
    const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([text], { type })); a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  const csvCell = v => { const s = String(v ?? ''); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  function fmtDate(iso) { return iso ? iso.replace(/^(\d{4})-(\d{2})-(\d{2}).*$/, '$1/$2/$3') : '—'; }
  function gapCell(v, max) {
    if (!isNum(v)) return n(v);
    const w = max ? Math.min(100, Math.abs(v) / max * 100) : 0;
    const bar = v < 0 ? `<span class="neg" style="width:${w}%"></span><span></span>` : `<span></span><span class="pos" style="width:${w}%"></span>`;
    return `<div class="gapbar-cell" title="產銷差異 ${nf.format(v)} 台"><span>${nf.format(v)}</span><span class="gapbar" aria-hidden="true">${bar}</span></div>`;
  }
  function select(id, label, options, value) {
    return `<label>${esc(label)}<select id="${id}">${options.map(([v, t]) => `<option value="${esc(v)}"${String(v) === String(value) ? ' selected' : ''}>${esc(t)}</option>`).join('')}</select></label>`;
  }
  const lineLabel = l => [l.name, l.sub, l.option].filter(Boolean).join('／');
  const shortName = s => s.replace(/-BOM庫存表$|BOM庫存表$/, '');

  // ---------- 解鎖與資料載入 ----------
  async function unlock(ev) {
    ev.preventDefault();
    const pw = $('#gatePassword').value; const msg = $('#gateMsg'); const btn = $('#gateSubmit');
    msg.className = 'gate-msg is-busy'; msg.textContent = '解密中…'; btn.disabled = true;
    try {
      const res = await fetch('data/data.enc.json', { cache: 'no-store' });
      if (!res.ok) throw new Error('讀不到資料檔（HTTP ' + res.status + '）。');
      const env = await res.json();
      const data = await window.AMACrypto.decrypt(env, pw);
      $('#gatePassword').value = '';
      load(data, 'encrypted', data.source && data.source.file);
    } catch (e) {
      msg.className = 'gate-msg'; msg.textContent = e.code === 'BAD_PASSWORD' ? '密碼不正確，請再試一次。' : (e.message || '無法開啟資料。');
    } finally { btn.disabled = false; }
  }
  function loadScript(src) {
    return new Promise((ok, fail) => { if (window.XLSX) return ok(); const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = () => fail(new Error('無法載入 Excel 解析元件。')); document.head.appendChild(s); });
  }
  async function importExcel(file) {
    if (!file) return;
    if (file.size > 30 * 1024 * 1024) { toast('檔案超過 30 MB，請確認是否為正確的產銷檔。'); return; }
    toast('正在本機解析 ' + file.name + '…');
    try {
      await loadScript('vendor/xlsx.full.min.js');
      const wb = window.XLSX.read(await file.arrayBuffer(), { type: 'array', cellFormula: true });
      const data = window.AMAParser.parseWorkbook(wb, { file: file.name, parsedAt: new Date().toISOString() });
      if (!data.plan) throw new Error('找不到「國工風扇產銷表」工作表，請確認檔案。');
      load(data, 'local', file.name);
      toast('已解析 ' + file.name + '（僅在此瀏覽器）');
    } catch (e) { toast(e.message || '解析失敗'); }
  }
  function load(data, source, fileName) {
    state.data = data; state.source = source; state.fileName = fileName || '';
    $('#gate').hidden = true; $('#app').hidden = false;
    const p = data.plan;
    $('#sourceLine').textContent = `來源：${data.source.file || fileName || '未命名'}｜產銷表日期 ${fmtDate(p && p.date)}｜解析於 ${fmtDate(data.source.parsedAt)}`;
    $('#sourceFoot').innerHTML = `資料：${esc(data.source.file || '')}<br>工作表 ${data.sheets.length} 張・ERP 品項 ${data.erp.items.length} 筆`;
    const b = $('#banner');
    if (source === 'local') { b.hidden = false; b.textContent = `目前顯示本機匯入的「${fileName}」。檔案沒有上傳；重新整理或鎖定後會回到網站上的加密版本。`; }
    else b.hidden = true;
    route();
  }
  function lock() { state.data = null; $('#view').innerHTML = ''; $('#app').hidden = true; $('#gate').hidden = false; $('#gateMsg').textContent = ''; $('#gatePassword').focus(); }

  // ---------- 路由 ----------
  function route() {
    if (!state.data) return;
    const page = (location.hash || '#overview').slice(1).split('?')[0];
    const key = PAGES[page] ? page : 'overview';
    document.querySelectorAll('#nav a').forEach(a => { const on = a.dataset.page === key; a.classList.toggle('is-active', on); if (on) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current'); });
    $('#pageTitle').textContent = PAGES[key];
    document.title = PAGES[key] + '｜AMA 產銷協調中心';
    render(key);
  }
  function render(key) {
    const view = $('#view');
    try { view.innerHTML = ({ overview, plan: () => planPage('plan'), forecast: () => planPage('forecast'), capacity, bom, erp, cost, logic, sheets })[key](); }
    catch (e) { console.error(e); view.innerHTML = `<div class="panel"><h2>此頁無法顯示</h2><p class="panel-desc">${esc(e.message)}</p></div>`; }
    bind(key);
  }

  // ---------- 衍生資料 ----------
  function derived() {
    const d = state.data; if (d._derived) return d._derived;
    const P = d.plan || { skus: [], totals: {}, pairs: [] };
    const gaps = P.skus.filter(s => isNum(s.values.gap) && s.values.gap < 0).sort((a, b) => a.values.gap - b.values.gap);
    const bomShort = []; const lineFlags = new Map();
    for (const b of d.bom) for (const l of b.lines) {
      if (isNum(l.diff) && l.diff < 0) bomShort.push({ b, l });
      for (const f of l.flags) { if (!lineFlags.has(f)) lineFlags.set(f, []); lineFlags.get(f).push({ b, l }); }
    }
    const pairFlags = P.pairs.flatMap(p => p.flags.map(f => ({ p, f })));
    const usedBy = new Map();
    for (const b of d.bom) for (const l of b.lines) if (l.code) { if (!usedBy.has(l.code)) usedBy.set(l.code, new Set()); usedBy.get(l.code).add(shortName(b.sheet)); }
    return (d._derived = { P, gaps, bomShort, lineFlags, pairFlags, usedBy });
  }

  // ---------- 總覽 ----------
  function agendaItems() {
    const d = state.data; const { P, gaps, bomShort, lineFlags, pairFlags } = derived(); const items = [];
    for (const s of gaps) {
      const v = s.values;
      items.push({ tag: 'short', label: '成品缺口', text: `${s.spec} 本月不足 ${nf.format(-v.gap)} 台`, note: `月需求餘量 ${ntext(v.remainingDemand)}・總庫存 ${ntext(v.stockTotal)}・預計生產 ${ntext(v.plannedProduction)}・託外 ${ntext(v.outsourced)}（${s.series}／${s.type}）` });
    }
    for (const { b, l } of bomShort) {
      const needQty = -l.diff * (l.qty || 1); const cover = isNum(l.unreceived) && l.unreceived >= needQty;
      items.push({ tag: 'short', label: 'BOM 缺料', text: `${shortName(b.sheet)}：${lineLabel(l)} 差 ${nf.format(-l.diff)} 台份`, note: `庫存 ${ntext(l.stock)}＋在途 ${ntext(l.transit)}，需求 ${ntext(l.demand)} 台；未進貨 ${ntext(l.unreceived)}${cover ? '，到貨後可補足' : ''}${l.flags.length ? '。⚑ 此列需求公式有疑點：' + l.flags.join('；') : ''}` });
    }
    for (const { p, f } of pairFlags) items.push({ tag: 'warn', label: '公式疑點', text: f, note: `產銷表 ${p.cell}「${p.label}」= ${p.formula || ''}；影響引用此小計的 BOM 需求。` });
    for (const [f, hits] of lineFlags) {
      if (/缺少括號/.test(f)) continue;
      const sheets = [...new Set(hits.map(h => shortName(h.b.sheet)))];
      items.push({ tag: 'warn', label: '公式疑點', text: `${lineLabel(hits[0].l)}：${f}`, note: `出現在 ${sheets.join('、')}（${hits.length} 列）` });
    }
    for (const i of d.issues.filter(i => i.level === 'warn' && !/庫存\/台數」為錯誤值/.test(i.text))) items.push({ tag: 'info', label: '資料檢查', text: i.text, note: i.sheet });
    const bomDates = [...new Set(d.bom.map(b => b.stockDate).filter(Boolean))].join('、');
    const trDates = [...new Set(d.bom.map(b => b.transitDate).filter(Boolean))].join('、');
    items.push({ tag: 'info', label: '資料時間', text: `產銷表 ${fmtDate(P.date)}；BOM 庫存取自 ERP ${bomDates || '未標示'}；在途／未進貨 ${trDates || '未標示'}`, note: '日期不一致時，缺料判斷可能已過時。' });
    return items;
  }
  function overview() {
    const d = state.data; const { P, gaps, bomShort, lineFlags, pairFlags } = derived(); const t = P.totals || {}; const F = d.forecast;
    const gapSum = gaps.reduce((a, s) => a - s.values.gap, 0);
    const flagCount = pairFlags.length + [...lineFlags.keys()].filter(f => !/缺少括號/.test(f)).length;
    const kpis = [
      ['本月總需求', t.totalDemand, '台', `國外 ${ntext(t.demandExport)}・國內 ${ntext(t.demandDomestic)}`],
      ['本月總實銷', t.salesTotal, '台', `本週 ${ntext(t.salesWeek)}・上周累 ${ntext(t.salesPrev)}`],
      ['成品總庫存', t.stockTotal, '台', `北環倉 ${ntext(t.stockNorth)}・南環倉 ${ntext(t.stockSouth)}・外銷倉 ${ntext(t.stockExport)}`],
      ['成品缺口型號', gaps.length, '款', `缺口合計 ${nf.format(gapSum)} 台`, gaps.length > 0],
      ['BOM 缺料項', bomShort.length, '項', '僅 AH、AE 吊扇表有需求台數', bomShort.length > 0],
      ['公式疑點', flagCount, '類', '詳見「資料邏輯」頁', flagCount > 0],
      ['次月預計生產', F ? F.totals.plannedProduction : null, '台', F ? `託外 ${ntext(F.totals.outsourced)} 台` : '無次月預估表']
    ];
    const items = agendaItems();
    return `
      <div class="grid-kpi">${kpis.map(([l, v, u, f, a]) => `<div class="kpi${a ? ' is-alert' : ''}"><div class="kpi-label">${esc(l)}</div><div class="kpi-value">${isNum(v) ? nf.format(v) : '—'}<small>${esc(u)}</small></div><div class="kpi-foot">${esc(f)}</div></div>`).join('')}</div>
      <section class="panel" aria-labelledby="agendaTitle">
        <div class="filters" style="justify-content:space-between;align-items:center;margin-bottom:6px">
          <div><h2 id="agendaTitle">會前重點</h2><p class="panel-desc">依目前資料以固定規則整理：成品缺口 → BOM 缺料 → 公式疑點 → 資料檢查。決策仍由業務、採購、生產確認。</p></div>
          <button class="btn btn-secondary" type="button" id="copyAgenda">複製會前摘要</button>
        </div>
        <ol class="agenda">${items.map(i => `<li><span class="tag tag-${i.tag}">${esc(i.label)}</span><div>${esc(i.text)}<small>${esc(i.note || '')}</small></div></li>`).join('')}</ol>
      </section>
      <div class="two-col">
        <section class="panel"><h2>產品族群需求小計</h2><p class="panel-desc">來自產銷表下方小計（BOM 需求用量由這些小計帶入）。</p>
          <div class="table-wrap"><table class="t"><thead><tr><th class="l">族群</th><th>台數</th></tr></thead><tbody>${P.familyTotals.map(f => `<tr><td class="l">${esc(f.label)}</td><td>${n(f.value)}</td></tr>`).join('')}</tbody></table></div>
        </section>
        <section class="panel"><h2>各 BOM 表狀態</h2><p class="panel-desc">ERP 對帳：BOM 表的庫存、在途、未進貨與 ERP 工作表以品項編碼比對。</p>
          <div class="table-wrap"><table class="t"><thead><tr><th class="l">BOM 表</th><th>缺料</th><th>疑點列</th><th class="l">ERP 對帳</th></tr></thead><tbody>${d.bom.map(b => `<tr><td class="l"><a href="#bom" data-bom="${esc(b.sheet)}">${esc(shortName(b.sheet))}</a></td><td class="${b.shortageCount ? 'neg' : ''}">${b.hasDemand ? b.shortageCount : '<span class="muted small">未填需求</span>'}</td><td>${b.flagCount}</td><td class="l small">${b.erpCheck.lines ? (b.erpCheck.mismatch ? `<span class="chip warn">${b.erpCheck.mismatch} 筆不一致</span>` : `<span class="chip ok">${b.erpCheck.lines} 筆一致</span>`) : '<span class="chip">無品項編碼欄</span>'}</td></tr>`).join('')}</tbody></table></div>
        </section>
      </div>`;
  }
  function agendaText() {
    const d = state.data; const P = d.plan; const t = P.totals;
    const lines = [`【AMA 產銷協調會前摘要】產銷表 ${fmtDate(P.date)}`, `本月總需求 ${ntext(t.totalDemand)} 台（國外 ${ntext(t.demandExport)}／國內 ${ntext(t.demandDomestic)}），成品總庫存 ${ntext(t.stockTotal)} 台，總實銷 ${ntext(t.salesTotal)} 台。`, ''];
    agendaItems().forEach((i, k) => lines.push(`${k + 1}. [${i.label}] ${i.text}${i.note ? '（' + i.note + '）' : ''}`));
    lines.push('', `資料來源：${d.source.file}；網站依 Excel 快取值整理，未連線 ERP。`);
    return lines.join('\n');
  }

  // ---------- 本月產銷／次月預估 ----------
  function planPage(kind) {
    const P = state.data[kind];
    if (!P) return `<div class="panel"><h2>找不到${kind === 'plan' ? '產銷表' : '次月預估表'}</h2></div>`;
    const ui = state.ui[kind]; if (ui.active === undefined) ui.active = true;
    const keys = COLS.filter(([k]) => P.rowLabels.some(r => r.key === k));
    const seriesOpts = [['', '全部系列'], ...[...new Set(P.skus.map(s => s.series))].map(s => [s, s])];
    let rows = P.skus.filter(s => (!ui.series || s.series === ui.series) && (!ui.short || (isNum(s.values.gap) && s.values.gap < 0)) && (!ui.active || keys.some(([k]) => isNum(s.values[k]) && s.values[k] !== 0)));
    const max = Math.max(1, ...rows.map(s => Math.abs(s.values.gap || 0)));
    let body = ''; let grp = '';
    for (const s of rows) {
      const g = `${s.series}｜${s.type}`;
      if (g !== grp) { grp = g; body += `<tr class="group"><td colspan="${keys.length + 2}">${esc(g)}</td></tr>`; }
      const short = isNum(s.values.gap) && s.values.gap < 0;
      body += `<tr class="${short ? 'is-short' : ''}"><td class="l"><strong>${esc(s.spec)}</strong></td><td class="l small muted">${esc(s.bomFamily || '')}</td>${keys.map(([k]) => k === 'gap' ? `<td>${gapCell(s.values.gap, max)}</td>` : `<td>${n(s.values[k])}</td>`).join('')}</tr>`;
    }
    const tot = Object.fromEntries(keys.map(([k]) => [k, rows.reduce((a, s) => a + (s.values[k] || 0), 0)]));
    const issues = state.data.issues.filter(i => i.sheet === P.sheet);
    const dateWarn = kind === 'forecast' && issues.find(i => /日期/.test(i.text));
    return `
      <section class="panel">
        <h2>${kind === 'plan' ? '成品產銷表' : '次月預估'}｜表頭日期 ${fmtDate(P.date)}</h2>
        <p class="panel-desc">月需求餘量 = 總需求 ÷ 工天數（${ntext(P.workDays)}）× 剩餘工天（${ntext(P.remainingDays)}）；產銷差異 = 總庫存 + 預計生產 + 託外生產 − 月需求餘量。單位：台。</p>
        ${dateWarn ? `<p class="banner">${esc(dateWarn.text)}</p>` : ''}
        <div class="filters">
          ${select(kind + 'Series', '產品系列', seriesOpts, ui.series || '')}
          <label class="check"><input type="checkbox" id="${kind}Active"${ui.active ? ' checked' : ''}> 只看有數量的型號</label>
          <label class="check"><input type="checkbox" id="${kind}Short"${ui.short ? ' checked' : ''}> 只看缺口</label>
          <span class="muted small">顯示 ${rows.length}／${P.skus.length} 款</span>
        </div>
        <div class="legend"><span><i style="background:var(--short)"></i>不足（負值）</span><span><i style="background:var(--bar-pos)"></i>有餘（正值）</span><span>長條長度依目前清單最大差異等比例</span></div>
        <div class="table-wrap"><table class="t">
          <thead><tr><th class="l">型號／規格</th><th class="l">對應 BOM</th>${keys.map(([, l]) => `<th>${esc(l)}</th>`).join('')}</tr></thead>
          <tbody>${body || `<tr><td class="l" colspan="${keys.length + 2}">沒有符合條件的型號。</td></tr>`}</tbody>
          <tfoot><tr><td class="l">合計（目前清單）</td><td></td>${keys.map(([k]) => `<td>${nf.format(tot[k])}</td>`).join('')}</tr></tfoot>
        </table></div>
        ${P.totalsSheet && isNum(P.totalsSheet.gap) ? `<p class="small muted">原表合計欄：總需求 ${ntext(P.totalsSheet.totalDemand)}、月需求餘量 ${ntext(P.totalsSheet.remainingDemand)}、預計生產 ${ntext(P.totalsSheet.plannedProduction)}、託外 ${ntext(P.totalsSheet.outsourced)}、產銷差異 ${ntext(P.totalsSheet.gap)}。</p>` : ''}
      </section>
      ${P.pairs.length ? `<section class="panel"><h2>族群小計與公式展開</h2><p class="panel-desc">這些小計是 BOM「當月需求用量」的來源。網站讀取 Excel 公式，列出每個小計實際加總了哪些型號。</p>
        <div class="table-wrap"><table class="t"><thead><tr><th class="l">小計</th><th>台數</th><th class="l">儲存格</th><th class="l">實際加總的型號</th><th class="l">疑點</th></tr></thead><tbody>
        ${P.pairs.map(p => `<tr class="${p.flags.length ? 'is-flag' : ''}"><td class="l"><strong>${esc(p.label)}</strong></td><td>${n(p.value)}</td><td class="l small"><span class="formula">${esc(p.cell)} = ${esc(p.formula || '')}</span></td><td class="l sku-list wrap">${esc(p.skus.join('、'))}</td><td class="l wrap">${p.flags.map(f => `<span class="flag">${esc(f)}</span>`).join('')}</td></tr>`).join('')}
        </tbody></table></div></section>` : ''}`;
  }

  // ---------- 產能 ----------
  function capacity() {
    const which = state.ui.capacity; const P = state.data[which] || state.data.plan; const C = P && P.capacity;
    const seg = `<div class="seg" role="group" aria-label="期間"><button type="button" data-cap="plan" aria-pressed="${which === 'plan'}">本月產銷表</button><button type="button" data-cap="forecast" aria-pressed="${which === 'forecast'}">次月預估</button></div>`;
    if (!C) return `<div class="panel">${seg}<p>此表沒有產能區塊。</p></div>`;
    const sites = C.schedules;
    return `<section class="panel"><div class="filters" style="justify-content:space-between"><div><h2>製程工時與排程量</h2><p class="panel-desc">取自${which === 'plan' ? '產銷表' : '次月預估'}右側產能區。日工時 ${ntext(C.dailyMinutes)} 分鐘。</p></div>${seg}</div>
      <div class="grid-kpi">${C.sites.map(s => `<div class="kpi"><div class="kpi-label">${esc(s.name)}目前總工時</div><div class="kpi-value">${n(s.totalMinutes)}<small>分</small></div><div class="kpi-foot">${isNum(s.totalMinutes) && isNum(C.dailyMinutes) ? `約 ${nf.format(s.totalMinutes / C.dailyMinutes)} 個工作天 × ${nf.format(C.dailyMinutes)} 分` : ''}</div></div>`).join('')}</div>
      <div class="table-wrap"><table class="t"><thead><tr><th class="l">製程</th><th class="l">工時代號</th><th>工時（分）</th><th>組／人</th>${sites.map(s => `<th>${esc(s)}排程量</th>`).join('')}</tr></thead><tbody>
      ${C.processes.map(p => `<tr><td class="l">${esc(p.process)}</td><td class="l">${esc(p.code)}</td><td>${n(p.minutes)}</td><td>${n(p.perPerson)}</td>${sites.map(s => `<td>${n(p.schedule[s])}</td>`).join('')}</tr>`).join('')}
      </tbody></table></div>
      <h3>現場備註</h3><ul class="plain">${C.sites.filter(s => s.note).map(s => `<li><strong>${esc(s.name)}：</strong>${esc(s.note)}</li>`).join('') || '<li class="muted">無</li>'}</ul>
      <p class="small muted">排程量未標示單位（台、組或分鐘），看板不換算產能負荷；請生產單位確認口徑後再加入負荷計算。</p></section>`;
  }

  // ---------- BOM ----------
  function bom() {
    const d = state.data; const ui = state.ui.bom;
    if (!d.bom.length) return '<div class="panel"><h2>沒有 BOM 庫存表</h2></div>';
    const B = d.bom.find(b => b.sheet === ui.sheet) || d.bom[0]; ui.sheet = B.sheet;
    const q = (ui.q || '').trim().toLowerCase();
    let rows = B.lines.filter(l => (!ui.only || (isNum(l.diff) && l.diff < 0) || l.flags.length) && (!q || (lineLabel(l) + ' ' + l.code).toLowerCase().includes(q)));
    let last = '';
    const body = rows.map(l => {
      const short = isNum(l.diff) && l.diff < 0; const showName = l.name !== last; last = l.name;
      const cover = short && isNum(l.unreceived) && l.unreceived >= -l.diff * (l.qty || 1);
      const erp = l.erp ? (l.erp.stock && l.erp.transit && l.erp.unreceived ? '' : '<span class="chip warn">與 ERP 不一致</span>') : (l.code ? '<span class="chip warn">ERP 查無</span>' : '');
      return `<tr class="${short ? 'is-short' : l.flags.length ? 'is-flag' : ''}">
        <td class="l wrap">${showName ? `<strong>${esc(l.name)}</strong>` : ''}${l.sub || l.option ? `<div class="small">${esc([l.sub, l.option].filter(Boolean).join('／'))}</div>` : ''}</td>
        <td class="l small">${esc(l.code)}${erp}</td><td>${n(l.qty)}</td><td>${n(l.need)}</td><td>${n(l.stock)}</td><td>${n(l.transit)}</td>
        <td>${l.errors.includes('perUnit') ? '<span class="err">錯誤值</span>' : n(l.perUnit)}</td><td>${n(l.demand)}</td><td class="${short ? 'neg' : isNum(l.diff) ? 'pos' : ''}">${n(l.diff)}${cover ? '<div class="small muted">未進貨可補</div>' : ''}</td><td>${n(l.unreceived)}</td>
        <td class="l wrap">${l.demandSkus.length ? `<details class="inline"><summary>${l.demandSkus.length} 款型號</summary><div class="sku-list">${esc(l.demandSkus.join('、'))}</div></details>` : '<span class="muted small">—</span>'}${l.flags.map(f => `<span class="flag">${esc(f)}</span>`).join('')}${l.note ? `<div class="small muted">${esc(l.note)}</div>` : ''}</td></tr>`;
    }).join('');
    const ec = B.erpCheck;
    return `<section class="panel">
      <div class="filters">
        ${select('bomSheet', 'BOM 表', d.bom.map(b => [b.sheet, shortName(b.sheet) + (b.shortageCount ? `（缺 ${b.shortageCount}）` : '')]), B.sheet)}
        <label>搜尋料件或編碼<input type="search" id="bomQ" value="${esc(ui.q || '')}" placeholder="例如：軸承、FN99"></label>
        <label class="check"><input type="checkbox" id="bomOnly"${ui.only ? ' checked' : ''}> 只看缺料或疑點</label>
      </div>
      <div class="chips" style="margin-bottom:10px">
        <span class="chip">庫存日 ${esc(B.stockDate || '未標示')}</span><span class="chip">在途 ${esc(B.transitDate || '未標示')}</span><span class="chip">未進貨 ${esc(B.unreceivedDate || '未標示')}</span>
        ${ec.lines ? `<span class="chip ${ec.mismatch ? 'warn' : 'ok'}">ERP 對帳 ${ec.lines - ec.mismatch}／${ec.lines} 一致</span>` : '<span class="chip">此表無品項編碼，無法與 ERP 對帳</span>'}
        ${B.hasDemand ? `<span class="chip ${B.shortageCount ? 'short' : 'ok'}">缺料 ${B.shortageCount} 項</span>` : '<span class="chip warn">此表未填需求台數與差異</span>'}
        <span class="chip ${B.flagCount ? 'warn' : ''}">需求公式疑點 ${B.flagCount} 列</span>
      </div>
      <p class="panel-desc">可支應台數 =（庫存 + 在途）÷ BOM量；差異 = 可支應台數 − 需求台數。<strong>未進貨不計入可支應台數</strong>。「當月需求用量」在原表欄名為「當月安存」，實際公式是 BOM量 × 對應型號的總需求。</p>
      <div class="table-wrap"><table class="t"><thead><tr><th class="l">料件</th><th class="l">品項編碼</th><th>BOM量</th><th>當月需求用量</th><th>庫存</th><th>在途</th><th>可支應台數</th><th>需求台數</th><th>差異</th><th>未進貨</th><th class="l">需求來源／疑點</th></tr></thead>
      <tbody>${body || '<tr><td class="l" colspan="11">沒有符合條件的料件。</td></tr>'}</tbody></table></div></section>`;
  }

  // ---------- ERP ----------
  function erp() {
    const d = state.data; const ui = state.ui.erp; const { usedBy } = derived();
    const q = (ui.q || '').trim().toLowerCase(); const f = ui.f || 'all';
    let rows = d.erp.items.filter(i => (!q || (i.code + ' ' + i.name).toLowerCase().includes(q)) &&
      (f === 'all' || (f === 'transit' && i.transit) || (f === 'unreceived' && i.unreceived) || (f === 'bom' && usedBy.has(i.code)) || (f === 'zero' && usedBy.has(i.code) && !i.stock)));
    const total = rows.length; rows = rows.slice(0, 400);
    return `<section class="panel"><h2>ERP 物料（庫存＋在途＋未進貨，以品項編碼合併）</h2>
      <p class="panel-desc">來源：ERP庫存 ${d.erp.counts.stock} 筆、ERP在途 ${d.erp.counts.transit} 筆、ERP未進貨 ${d.erp.counts.unreceived} 筆。${d.erp.duplicates.length ? `有 ${d.erp.duplicates.length} 個編碼重複，已加總。` : '未發現重複編碼。'}</p>
      <div class="filters">
        <label>搜尋編碼或品名<input type="search" id="erpQ" value="${esc(ui.q || '')}" placeholder="例如：磁石、FN99B"></label>
        ${select('erpF', '篩選', [['all', '全部'], ['transit', '有在途'], ['unreceived', '有未進貨'], ['bom', 'BOM 有使用'], ['zero', 'BOM 使用但庫存為 0']], f)}
        <span class="muted small">符合 ${total} 筆${total > 400 ? '，顯示前 400 筆' : ''}</span>
      </div>
      <div class="table-wrap"><table class="t"><thead><tr><th class="l">品項編碼</th><th class="l">品項名稱</th><th>庫存</th><th>在途</th><th>未進貨</th><th class="l">BOM 使用</th><th class="l">備註</th></tr></thead><tbody>
      ${rows.map(i => `<tr><td class="l">${esc(i.code)}</td><td class="l wrap">${esc(i.name)}</td><td>${n(i.stock)}</td><td>${n(i.transit)}</td><td>${n(i.unreceived)}</td><td class="l sku-list wrap">${esc([...(usedBy.get(i.code) || [])].join('、'))}</td><td class="l small wrap">${esc(i.notes.join('；'))}</td></tr>`).join('') || '<tr><td class="l" colspan="7">沒有符合條件的品項。</td></tr>'}
      </tbody></table></div></section>`;
  }

  // ---------- 成本 ----------
  function cost() {
    const d = state.data; const ui = state.ui.cost; const C = d.cost;
    if (!C.products.length) return '<div class="panel"><h2>沒有成本明細表</h2></div>';
    const P = C.products.find(p => p.sheet === ui.sheet) || C.products[0]; ui.sheet = P.sheet;
    const sum = C.summary.map(s => { const p = C.products.find(x => !x.isTarget && x.product.replace(/-燈/, '') === s.name.replace(/-燈/, '')); return { ...s, sheetTotal: p ? p.total : null, sheet: p && p.sheet }; });
    const total = P.total || P.lines.reduce((a, l) => a + (l.subtotal || 0), 0);
    const top = Math.max(1, ...P.lines.map(l => l.subtotal || 0));
    const otherTables = Object.keys(d.tables).filter(k => !/^ERP/.test(k));
    return `<section class="panel"><h2>成品成本總表</h2><p class="panel-desc">總表單價與各明細表「總計」並列；差額可能是加工、管銷或版本差異，請財務／採購確認口徑。</p>
      <div class="table-wrap"><table class="t"><thead><tr><th class="l">品名</th><th>總表單價</th><th>明細表總計</th><th>差額</th><th class="l">備註</th></tr></thead><tbody>
      ${sum.map(s => `<tr><td class="l">${s.sheet ? `<a href="#cost" data-cost="${esc(s.sheet)}">${esc(s.name)}</a>` : esc(s.name)}</td><td>${n(s.price)}</td><td>${n(s.sheetTotal)}</td><td class="${isNum(s.price) && isNum(s.sheetTotal) && Math.abs(s.price - s.sheetTotal) > 1 ? 'neg' : ''}">${isNum(s.price) && isNum(s.sheetTotal) ? nf.format(s.price - s.sheetTotal) : '—'}</td><td class="l small">${esc(s.note)}</td></tr>`).join('')}
      </tbody></table></div></section>
      <section class="panel"><div class="filters">${select('costSheet', '成本明細', C.products.map(p => [p.sheet, p.isTarget ? p.product + '（降成本目標）' : p.product]), P.sheet)}<span class="muted small">明細總計 ${ntext(P.total)} 元</span></div>
      <p class="panel-desc">黑／銀、匝數等「選項擇一」的列會同時列出；占比長條以單列小計相對最大項呈現。</p>
      <div class="table-wrap"><table class="t"><thead><tr><th class="l">料件</th><th>BOM量</th><th>單價</th><th>小計</th><th>占總計</th><th>材重</th><th class="l">廠商</th>${P.isTarget ? '<th>更降成本金額</th><th class="l">新廠商</th><th>降幅</th><th class="l">原由</th>' : ''}</tr></thead><tbody>
      ${P.lines.map(l => `<tr><td class="l wrap"><strong>${esc(l.name)}</strong>${l.sub || l.option ? `<div class="small">${esc([l.sub, l.option].filter(x => x && x !== l.name).join('／'))}</div>` : ''}</td><td>${n(l.qty)}</td><td>${n(l.price)}</td><td><div class="share"><span>${n(l.subtotal)}</span><b style="width:${isNum(l.subtotal) ? Math.max(1, l.subtotal / top * 70) : 0}px" aria-hidden="true"></b></div></td><td>${isNum(l.subtotal) && total ? nf.format(l.subtotal / total * 100) + '%' : '—'}</td><td>${n(l.weight)}</td><td class="l">${esc(l.vendor)}</td>${P.isTarget ? `<td>${n(l.targetPrice)}</td><td class="l">${esc(l.newVendor)}</td><td>${isNum(l.cutPct) ? nf.format(l.cutPct * 100) + '%' : '—'}</td><td class="l wrap small">${esc(l.reason)}</td>` : ''}</tr>`).join('')}
      </tbody><tfoot><tr><td class="l">總計</td><td></td><td></td><td>${n(P.total)}</td><td></td><td></td><td></td>${P.isTarget ? '<td></td><td></td><td></td><td></td>' : ''}</tr></tfoot></table></div></section>
      <section class="panel"><h2>子件與外購價目表</h2><p class="panel-desc">以下工作表以原格式提供，可在「原始工作表」頁查看。</p><div class="chips">${otherTables.map(k => `<a class="chip" href="#sheets" data-sheet="${esc(k)}">${esc(k)}</a>`).join('')}</div></section>`;
  }

  // ---------- 資料邏輯 ----------
  function logic() {
    const d = state.data; const { P, lineFlags, pairFlags } = derived();
    const lf = [...lineFlags.entries()];
    const verifyGap = P.skus.filter(s => isNum(s.values.gap)).length;
    const gapIssues = d.issues.filter(i => /產銷差異/.test(i.text)).length;
    const ex = P.skus.find(s => (s.values.totalDemand || 0) > 0 && (s.values.stockTotal || 0) > (s.values.totalDemand || 0));
    const fcSpecs = new Set(((d.forecast && d.forecast.skus) || []).map(s => s.spec)); const plSpecs = new Set(P.skus.map(s => s.spec));
    const onlyPlan = d.forecast ? [...plSpecs].filter(s => !fcSpecs.has(s)) : []; const onlyFc = [...fcSpecs].filter(s => !plSpecs.has(s));
    return `
    <section class="panel"><h2>資料怎麼流動</h2><p class="panel-desc">這份 Excel 由採購從 ERP 匯出整理。網站把它拆成三層：ERP 原始層（以品項編碼為主鍵）→ BOM 齊套層 → 成品產銷層。日後串接 ERP，只需替換最左邊的來源。</p>
      <div class="flow">
        <div class="node src"><h4>① ERP 匯出</h4><p>ERP庫存（品項編碼、品名、合計）<br>ERP在途（品項編碼、數量）<br>ERP未進貨（品項編碼、未進貨數量）</p></div>
        <div class="node src"><h4>② 業務／生產填報</h4><p>產銷表：各型號國外／國內需求、安存值、實銷、北環倉／南環倉／外銷倉成品庫存、預計生產、託外</p></div>
        <div class="node"><h4>③ 族群小計</h4><p>產銷表下方：AH總數、AH吊扇黑、有燈、40/79匝…<br>以 SUM 指定型號欄</p></div>
        <div class="node"><h4>④ BOM 庫存表（8 張）</h4><p>當月需求用量 = BOM量 × ③小計<br>庫存／在途／未進貨 = 以品項編碼 VLOOKUP ①</p></div>
        <div class="node out"><h4>⑤ 判斷</h4><p>成品：產銷差異<br>料件：可支應台數、差異<br>成本表、產能表另列</p></div>
      </div></section>

    <section class="panel"><h2>欄位與公式對照</h2>
      <div class="table-wrap"><table class="t"><thead><tr><th class="l">欄位</th><th class="l">原表公式（以 C 欄型號為例）</th><th class="l">說明</th></tr></thead><tbody>
        <tr><td class="l">需求總計</td><td class="l"><span class="formula">C8 = C6 + C7</span></td><td class="l wrap">國外 + 國內</td></tr>
        <tr><td class="l">總需求</td><td class="l"><span class="formula">C10 = C8 + C9</span></td><td class="l wrap">需求總計 + 安存值（目前安存值皆空白）</td></tr>
        <tr><td class="l">總實銷</td><td class="l"><span class="formula">C13 = C11 + C12</span></td><td class="l wrap">上周累 + 本週</td></tr>
        <tr><td class="l">總庫存</td><td class="l"><span class="formula">C17 = C14 + C15 + C16</span></td><td class="l wrap">北環倉 + 南環倉 + 外銷倉（成品）</td></tr>
        <tr><td class="l">月需求餘量</td><td class="l"><span class="formula">C19 = C10 / $B$21 × $B$18</span></td><td class="l wrap">總需求 ÷ 工天數（${ntext(P.workDays)}）× 剩餘工天（${ntext(P.remainingDays)}）</td></tr>
        <tr><td class="l">產銷差異</td><td class="l"><span class="formula">C24 = C17 + C20 + C23 − C19</span></td><td class="l wrap">總庫存 + 預計生產 + 託外 − 月需求餘量；本檔 ${verifyGap} 款驗算${gapIssues ? `有 ${gapIssues} 款不一致` : '全部一致'}</td></tr>
        <tr><td class="l">BOM 當月需求用量</td><td class="l"><span class="formula">F = E × 產銷表!小計</span></td><td class="l wrap">原欄名「當月安存」，實為需求用量，不是安全存量</td></tr>
        <tr><td class="l">BOM 庫存／在途／未進貨</td><td class="l"><span class="formula">VLOOKUP(L, ERP庫存!A:C, 3, FALSE)</span></td><td class="l wrap">L 欄為品項編碼；查無時顯示 #N/A</td></tr>
        <tr><td class="l">庫存/台數</td><td class="l"><span class="formula">I = (G + H) / E</span></td><td class="l wrap">（庫存 + 在途）÷ BOM量；未進貨不計入。吊管列為 G + H，未除以 BOM量</td></tr>
        <tr><td class="l">需求台數／差異</td><td class="l"><span class="formula">J = F / E；K = I − J</span></td><td class="l wrap">只有 AH、AE 吊扇兩張表有這兩欄</td></tr>
      </tbody></table></div></section>

    <section class="panel"><h2>本檔發現的公式疑點</h2><p class="panel-desc">網站逐格讀取 Excel 公式，追溯每個需求數字實際加總了哪些型號，再依料件的尺寸、顏色、有無燈比對。以下需要採購確認後修正原表。</p>
      <div class="table-wrap"><table class="t"><thead><tr><th class="l">位置</th><th class="l">疑點</th><th class="l">實際加總的型號</th></tr></thead><tbody>
      ${pairFlags.map(({ p, f }) => `<tr class="is-flag"><td class="l">產銷表 ${esc(p.cell)}</td><td class="l wrap">${esc(f)}<div class="small muted"><span class="formula">${esc(p.formula || '')}</span></div></td><td class="l sku-list wrap">${esc(p.skus.join('、'))}</td></tr>`).join('')}
      ${lf.map(([f, hits]) => `<tr class="is-flag"><td class="l wrap">${esc([...new Set(hits.map(h => shortName(h.b.sheet) + ' 第' + h.l.row + '列'))].join('、'))}</td><td class="l wrap"><strong>${esc(lineLabel(hits[0].l))}</strong>：${esc(f)}${hits[0].l.demandFormula ? `<div class="small muted"><span class="formula">${esc(hits[0].l.demandFormula)}</span></div>` : ''}</td><td class="l sku-list wrap">${esc(hits[0].l.demandSkus.join('、'))}</td></tr>`).join('')}
      </tbody></table></div>
      <div class="filters" style="margin-top:12px"><button class="btn btn-secondary" type="button" id="dlMap">下載「型號 → BOM 料件」需求對照表（CSV）</button><span class="small muted">可作為串接 ERP 時的對照表初稿，修正後取代手寫儲存格引用。</span></div>
    </section>

    <section class="panel"><h2>需要確認的口徑</h2><ol class="steps">
      <li><strong>零件需求未扣成品庫存。</strong>BOM 需求用量取自「總需求」${ex ? `，例如 ${esc(ex.spec)} 需求 ${ntext(ex.values.totalDemand)} 台、成品庫存 ${ntext(ex.values.stockTotal)} 台，仍計入 ${ntext(ex.values.totalDemand)} 台份零件` : ''}。若零件需求應依「需生產量 = max(0, 月需求餘量 − 成品庫存)」或「預計生產數」計算，數字會不同。</li>
      <li><strong>「當月安存」欄名。</strong>實際是當月需求用量；真正的安全存量（產銷表「安存值」列）目前皆空白。</li>
      <li><strong>未進貨不計入可支應台數。</strong>原表只用庫存 + 在途，未進貨另列；網站沿用並標示「未進貨可補」。</li>
      <li><strong>只有 AH、AE 吊扇兩張 BOM 表有需求台數與差異。</strong>立扇、側掛、坦克、水簾表只能看可支應台數。</li>
      <li><strong>次月預估沒有 BOM 展開。</strong>預計生產 ${ntext(d.forecast && d.forecast.totals.plannedProduction)} 台、託外 ${ntext(d.forecast && d.forecast.totals.outsourced)} 台尚未換算成料件需求。${onlyPlan.length || onlyFc.length ? `兩表型號不一致：只在本月表 ${esc(onlyPlan.join('、') || '無')}；只在次月表 ${esc(onlyFc.join('、') || '無')}，需先統一型號主檔。` : ''}</li>
      <li><strong>型號命名。</strong>產銷表以規格文字當欄位，沒有 ERP 成品品項編碼；規格誤植會讓族群小計歸錯（見資料檢查紀錄）。</li>
      <li><strong>產能排程量單位</strong>未標示；工時與排程量的關係待生產確認。</li>
    </ol></section>

    <section class="panel"><h2>串接 ERP 原始資料的建議步驟</h2><ol class="steps">
      <li><strong>建立型號主檔：</strong>每個成品型號對應 ERP 成品品項編碼、系列、尺寸、顏色、有無燈、馬達匝數。產銷表與次月預估共用同一份主檔。</li>
      <li><strong>以「型號 → 料件」對照表取代手寫公式：</strong>欄位為成品品項編碼、料件品項編碼、BOM量、選項條件（顏色／尺寸／燈）。可先用本頁下載的 CSV 修正，最終以 ERP 的 BOM 主檔為準。</li>
      <li><strong>固定 ERP 匯出格式：</strong>庫存（品項編碼、倉別、數量、單位、快照時間）、在途與未進貨（採購單號、品項編碼、數量、預計到貨日、供應商）、銷售訂單、工單與生產入庫。每份檔案保留匯出時間。</li>
      <li><strong>網站改讀原始匯出：</strong>需求用量 = Σ 型號需求 × BOM量，由程式計算；不再依賴儲存格位置，插入新型號欄也不會錯位。</li>
      <li><strong>再加入安全存量、週／月預測與待辦：</strong>沿用 2026-10-19 規劃：成品為主視圖，半成品、供應商備貨與成本依權限開放。</li>
    </ol></section>

    <section class="panel"><h2>資料檢查紀錄</h2><div class="table-wrap"><table class="t"><thead><tr><th class="l">等級</th><th class="l">工作表</th><th class="l">內容</th></tr></thead><tbody>
      ${d.issues.map(i => `<tr><td class="l"><span class="chip ${i.level === 'warn' ? 'warn' : i.level === 'error' ? 'short' : 'info'}">${i.level === 'warn' ? '注意' : i.level === 'error' ? '錯誤' : '提示'}</span></td><td class="l">${esc(i.sheet)}</td><td class="l wrap">${esc(i.text)}</td></tr>`).join('')}
    </tbody></table></div></section>`;
  }
  function mappingCsv() {
    const head = ['BOM表', '列', '料件', '下階細項', '選項', '品項編碼', 'BOM量', '原表需求用量', '需求來源型號', '疑點', '原公式'];
    const rows = [head];
    for (const b of state.data.bom) for (const l of b.lines) rows.push([shortName(b.sheet), l.row, l.name, l.sub, l.option, l.code, l.qty ?? '', l.need ?? '', l.demandSkus.join('、'), l.flags.join('；'), l.demandFormula ? '=' + l.demandFormula : '']);
    return '﻿' + rows.map(r => r.map(csvCell).join(',')).join('\n');
  }

  // ---------- 原始工作表 ----------
  function colName(i) { return window.AMAParser.colName(i); }
  function sheets() {
    const d = state.data; const ui = state.ui.sheets; const name = d.raw[ui.sheet] ? ui.sheet : d.sheets[0]; ui.sheet = name;
    const g = d.raw[name]; const rows = g.rows.slice(0, 700);
    return `<section class="panel"><div class="filters">${select('sheetSel', '工作表', d.sheets.map(s => [s, s]), name)}<span class="muted small">${g.rows.length} 列 × ${g.cols} 欄${g.rows.length > 700 ? '，顯示前 700 列' : ''}。顯示 Excel 儲存的計算結果；「電控BOM」的供應商聯絡資料已移除。</span></div>
      <div class="table-wrap"><table class="t sheet-grid"><thead><tr><th class="rh"></th>${Array.from({ length: g.cols }, (_, i) => `<th>${colName(i)}</th>`).join('')}</tr></thead><tbody>
      ${rows.map((r, i) => `<tr><td class="rh">${i + 1}</td>${Array.from({ length: g.cols }, (_, c) => { const v = r[c]; return `<td title="${esc(v ?? '')}">${v === null || v === undefined ? '' : isNum(v) ? nf.format(v) : esc(v)}</td>`; }).join('')}</tr>`).join('')}
      </tbody></table></div></section>`;
  }

  // ---------- 事件 ----------
  function bind(key) {
    const re = () => render(key);
    const on = (id, ev, fn) => { const el = $('#' + id); if (el) el.addEventListener(ev, fn); };
    if (key === 'overview') on('copyAgenda', 'click', async () => { try { await navigator.clipboard.writeText(agendaText()); toast('已複製會前摘要'); } catch { download('會前摘要.txt', agendaText(), 'text/plain'); } });
    if (key === 'plan' || key === 'forecast') {
      const ui = state.ui[key];
      on(key + 'Series', 'change', e => { ui.series = e.target.value; re(); });
      on(key + 'Active', 'change', e => { ui.active = e.target.checked; re(); });
      on(key + 'Short', 'change', e => { ui.short = e.target.checked; re(); });
    }
    if (key === 'capacity') document.querySelectorAll('[data-cap]').forEach(b => b.addEventListener('click', () => { state.ui.capacity = b.dataset.cap; re(); }));
    if (key === 'bom') {
      const ui = state.ui.bom;
      on('bomSheet', 'change', e => { ui.sheet = e.target.value; re(); });
      on('bomOnly', 'change', e => { ui.only = e.target.checked; re(); });
      on('bomQ', 'input', debounce(e => { ui.q = e.target.value; re(); $('#bomQ').focus(); const v = $('#bomQ'); v.setSelectionRange(v.value.length, v.value.length); }));
    }
    if (key === 'erp') {
      const ui = state.ui.erp;
      on('erpF', 'change', e => { ui.f = e.target.value; re(); });
      on('erpQ', 'input', debounce(e => { ui.q = e.target.value; re(); const v = $('#erpQ'); v.focus(); v.setSelectionRange(v.value.length, v.value.length); }));
    }
    if (key === 'cost') on('costSheet', 'change', e => { state.ui.cost.sheet = e.target.value; re(); });
    if (key === 'sheets') on('sheetSel', 'change', e => { state.ui.sheets.sheet = e.target.value; re(); });
    if (key === 'logic') on('dlMap', 'click', () => download('型號-BOM料件需求對照表.csv', mappingCsv(), 'text/csv;charset=utf-8'));
  }
  function debounce(fn, ms = 220) { let t; return e => { clearTimeout(t); t = setTimeout(() => fn(e), ms); }; }

  document.addEventListener('click', e => {
    const a = e.target.closest('[data-action],[data-bom],[data-cost],[data-sheet]'); if (!a) return;
    if (a.dataset.action === 'pick-excel') { $('#excelFile').value = ''; $('#excelFile').click(); }
    else if (a.dataset.action === 'download-json' && state.data) { const { _derived, ...plain } = state.data; download(`ama-supply-${(state.data.plan && state.data.plan.date) || 'data'}.json`, JSON.stringify(plain, null, 1)); }
    else if (a.dataset.action === 'lock') lock();
    else if (a.dataset.bom) state.ui.bom.sheet = a.dataset.bom;
    else if (a.dataset.cost) { state.ui.cost.sheet = a.dataset.cost; if (location.hash === '#cost') { e.preventDefault(); render('cost'); } }
    else if (a.dataset.sheet) state.ui.sheets.sheet = a.dataset.sheet;
  });
  document.addEventListener('DOMContentLoaded', () => {
    $('#gateForm').addEventListener('submit', unlock);
    $('#excelFile').addEventListener('change', e => importExcel(e.target.files[0]));
    window.addEventListener('hashchange', () => { route(); $('#main').focus({ preventScroll: true }); window.scrollTo(0, 0); });
    $('#gatePassword').focus();
  });
})();
