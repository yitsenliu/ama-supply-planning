/*
 * AMA 產銷 Excel 解析器（瀏覽器與 Node 共用）。
 * 輸入：SheetJS workbook（{SheetNames, Sheets}）；輸出：看板使用的 JSON。
 * 只讀取儲存格的快取值，不重算公式；讀不到或錯誤值會保留並列入資料檢查。
 */
(function (root) {
  'use strict';

  const SCHEMA = 'ama-supply-gh/1';
  // 電控BOM 的 L、M、N 欄為供應商聯絡帳號、地址與聯絡人，不放進網站。
  const REDACT_COLUMNS = { '電控BOM': [11, 12, 13] };

  const PLAN_ROWS = {
    '需求/國外': 'demandExport', '需求/國內': 'demandDomestic', '需求/總計': 'demandTotal',
    '安存值': 'safety', '總需求': 'totalDemand',
    '實銷/上周累': 'salesPrev', '實銷/本周': 'salesWeek', '總實銷': 'salesTotal',
    '庫存/北環倉': 'stockNorth', '庫存/南環倉': 'stockSouth', '庫存/外銷倉': 'stockExport', '總庫存': 'stockTotal',
    '月需求餘量': 'remainingDemand', '預計生產數': 'plannedProduction', '託外生產數': 'outsourced', '產銷差異數': 'gap'
  };

  function colName(i) { let s = ''; i += 1; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  function addr(r, c) { return colName(c) + (r + 1); }
  function decodeRef(ref) {
    const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref || 'A1');
    const ci = s => [...s].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    if (!m) return { r1: 0, c1: 0, r2: 0, c2: 0 };
    return { r1: +m[2] - 1, c1: ci(m[1]), r2: (m[4] ? +m[4] : +m[2]) - 1, c2: ci(m[3] || m[1]) };
  }
  const isNum = v => typeof v === 'number' && Number.isFinite(v);
  const isText = v => typeof v === 'string' && v.trim() !== '';
  const isErr = v => typeof v === 'string' && /^#(N\/A|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!)/.test(v);
  const round = (v, d = 4) => isNum(v) ? Math.round(v * 10 ** d) / 10 ** d : v;
  const clean = v => typeof v === 'string' ? v.replace(/\r/g, '').trim() : v;
  const flat = v => typeof v === 'string' ? v.replace(/\s*\n\s*/g, ' ').trim() : v;

  function sheetAccess(ws, name) {
    const ref = decodeRef(ws['!ref']);
    const redact = new Set(REDACT_COLUMNS[name] || []);
    const mergeMap = new Map();
    for (const m of ws['!merges'] || []) {
      for (let r = m.s.r; r <= m.e.r; r++) for (let c = m.s.c; c <= m.e.c; c++) mergeMap.set(r + ',' + c, m.s);
    }
    function raw(r, c) {
      if (redact.has(c)) return null;
      const cell = ws[addr(r, c)];
      if (!cell || cell.v === undefined || cell.v === null) return null;
      if (cell.t === 'e') return cell.w || '#ERR';
      if (cell.t === 'n') return round(cell.v, 6);
      if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
      if (cell.t === 'd') return String(cell.w || cell.v);
      const s = clean(String(cell.v));
      return s === '' ? null : s;
    }
    function merged(r, c) {
      const s = mergeMap.get(r + ',' + c);
      return s ? raw(s.r, s.c) : raw(r, c);
    }
    function isMergedChild(r, c) { const s = mergeMap.get(r + ',' + c); return !!s && (s.r !== r || s.c !== c); }
    function formula(r, c) { if (redact.has(c)) return null; const cell = ws[addr(r, c)]; return cell && cell.f ? String(cell.f) : null; }
    return { ref, raw, merged, isMergedChild, formula, rows: ref.r2 + 1, cols: ref.c2 + 1 };
  }

  function findCells(S, test) {
    const out = [];
    for (let r = 0; r < S.rows; r++) for (let c = 0; c < S.cols; c++) { const v = S.raw(r, c); if (v !== null && test(v)) out.push({ r, c, v }); }
    return out;
  }
  function firstRight(S, r, c, test) { for (let x = c + 1; x < S.cols; x++) { const v = S.raw(r, x); if (v !== null && test(v)) return { c: x, v }; } return null; }

  function specLabel(s) {
    if (!isText(s)) return '';
    return s.split('\n').map(x => x.trim()).filter(Boolean).reduce((a, l) => a + (l.startsWith('-') ? l : (a ? ' ' : '') + l), '');
  }

  // ---------- 產銷表／次月預估 ----------
  function parsePlan(S, name, issues, ctx = {}) {
    const out = { sheet: name, date: null, skus: [], totals: {}, totalsSheet: {}, familyTotals: [], pairs: [], capacity: null, remainingDays: null, workDays: null };
    if (S.raw(0, 0) && /日期/.test(S.raw(0, 0))) {
      const y = S.raw(0, 1), m = S.raw(0, 2), d = S.raw(0, 3);
      if (isNum(y) && isNum(m)) out.date = `${y}-${String(m).padStart(2, '0')}-${String(isNum(d) ? d : 1).padStart(2, '0')}`;
    }
    const find = re => { for (let r = 0; r < Math.min(S.rows, 12); r++) if (isText(S.raw(r, 0)) && re.test(S.raw(r, 0))) return r; return -1; };
    const seriesRow = find(/產品系列/), typeRow = find(/產品種類/), specRow = find(/規格/);
    if (specRow < 0) { issues.push({ level: 'error', sheet: name, text: '找不到「規格」列，無法解析產品欄位。' }); return out; }
    const skuCols = [];
    for (let c = 2; c < S.cols; c++) { if (S.raw(specRow, c) === null) break; skuCols.push(c); }
    const lastSku = skuCols[skuCols.length - 1] || 2;

    // 列標籤
    const rowKeys = []; let bomRow = -1; let blockStart = S.rows;
    for (let r = specRow + 1; r < S.rows; r++) {
      const a = S.merged(r, 0), b = S.raw(r, 1);
      if (isText(a) && /總數|以上合計/.test(a)) { blockStart = r; break; }
      const label = isText(a) ? (isText(b) && !S.isMergedChild(r, 1) && b !== a ? `${a}/${b}` : a) : null;
      if (!isText(a) && isNum(b)) { out.remainingDays = b; if (skuCols.some(c => isText(S.merged(r, c)))) bomRow = r; continue; }
      if (label === '工天數' && isNum(b)) out.workDays = b;
      if (!label) continue;
      const key = PLAN_ROWS[label];
      if (key) rowKeys.push({ r, key, label });
      else if (label !== '工天數') rowKeys.push({ r, key: 'x:' + label, label });
    }
    const totalCol = (() => { const hit = findCells(S, v => isText(v) && /^(總預估|總數)$/.test(v)).find(x => x.c > lastSku); return hit ? hit.c : lastSku + 1; })();
    let prevSeries = '', prevType = '';
    for (const c of skuCols) {
      const series = flat(S.merged(seriesRow, c)) || prevSeries, type = flat(S.merged(typeRow, c)) || prevType;
      prevSeries = series; prevType = type;
      const rawSpec = S.raw(specRow, c);
      const values = {};
      for (const { r, key } of rowKeys) {
        const v = S.raw(r, c);
        if (isErr(v)) issues.push({ level: 'warn', sheet: name, text: `${addr(r, c)} 為 ${v}` });
        values[key] = isNum(v) ? v : null;
      }
      out.skus.push({ col: colName(c), series, type, spec: specLabel(rawSpec), bomFamily: bomRow >= 0 ? flat(S.merged(bomRow, c)) || null : null, values });
    }
    for (const { r, key } of rowKeys) {
      const v = S.raw(r, totalCol); if (isNum(v)) out.totalsSheet[key] = v;
      out.totals[key] = round(out.skus.reduce((a, s) => a + (s.values[key] || 0), 0), 2);
    }
    out.rowLabels = rowKeys.map(({ key, label }) => ({ key, label }));
    for (const sku of out.skus) if (ambiguous(sku.spec)) issues.push({ level: 'warn', sheet: name, text: `${sku.col}5 規格「${sku.spec}」同時含「無」與「燈」，顏色欄疑似誤植，公式群組可能因此歸錯。` });
    const tr = rowKeys.find(x => x.key === 'totalDemand');
    ctx.S = S; ctx.name = name; ctx.totalRow = tr ? tr.r : -1;
    ctx.skuByCol = new Map(skuCols.map((c, i) => [c, out.skus[i]]));
    // 驗算產銷差異 = 總庫存 + 預計生產 + 託外 − 月需求餘量
    for (const s of out.skus) {
      const v = s.values; if (v.gap === null && v.remainingDemand === null) continue;
      const calc = round((v.stockTotal || 0) + (v.plannedProduction || 0) + (v.outsourced || 0) - (v.remainingDemand || 0), 2);
      s.gapCalc = calc;
      if (isNum(v.gap) && Math.abs(calc - v.gap) > 0.05) issues.push({ level: 'warn', sheet: name, text: `${s.spec}（${s.col} 欄）產銷差異 ${v.gap} 與「庫存＋預計生產＋託外−月需求餘量」${calc} 不一致。` });
    }
    // 族群小計與成對標籤
    for (let r = blockStart; r < S.rows; r++) {
      const a = S.raw(r, 0); const cv = S.raw(r, 2);
      if (isText(a) && isNum(cv)) out.familyTotals.push({ label: a, value: cv });
      for (let c = 3; c <= lastSku; c++) {
        const t = S.raw(r, c), below = r + 1 < S.rows ? S.raw(r + 1, c) : null;
        if (isText(t) && isNum(below)) {
          const pair = { row: r + 1, cell: addr(r + 1, c), label: flat(t), value: below };
          const ex = expandPlanCell(ctx, r + 1, c);
          pair.skus = ex.skus.map(s => s.spec); pair.formula = S.formula(r + 1, c);
          pair.flags = [...ex.flags, ...attrFlags(pair.label, ex.skus, out.skus)];
          out.pairs.push(pair);
        }
      }
    }
    out.capacity = parseCapacity(S);
    return out;
  }

  // ---------- 公式展開：把 BOM 需求追溯到產銷表的產品欄 ----------
  function decodeCell(a) { const m = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(a); return m ? decodeRef(m[1] + m[2]) : null; }
  function refsIn(f) {
    const out = []; let m; const re = /(?:(?:'([^']+)'|([A-Za-z0-9_㐀-鿿\-.]+))!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?/g;
    while ((m = re.exec(f))) {
      const sheet = m[1] || m[2] || null; const a = decodeCell(m[3]); const b = m[4] ? decodeCell(m[4]) : a;
      if (!a) continue; out.push({ sheet, r1: a.r1, c1: a.c1, r2: b.r1, c2: b.c1, text: m[0] });
    }
    return out;
  }
  function expandPlanCell(ctx, r, c, depth = 0, seen = new Set()) {
    const res = { skus: [], flags: [] };
    if (!ctx.S || depth > 6) return res;
    const key = r + ',' + c; if (seen.has(key)) return res; seen.add(key);
    if (r === ctx.totalRow) {
      const sku = ctx.skuByCol.get(c);
      if (sku) res.skus.push(sku); else res.flags.push(`引用產銷表 ${addr(r, c)}，不是產品欄`);
      return res;
    }
    const f = ctx.S.formula(r, c);
    if (!f) {
      const v = ctx.S.raw(r, c);
      res.flags.push(v === null ? `引用產銷表空白儲存格 ${addr(r, c)}（需求恆為 0）` : `引用產銷表 ${addr(r, c)} 的固定值 ${v}`);
      return res;
    }
    for (const ref of refsIn(f)) {
      if (ref.sheet && ref.sheet !== ctx.name) continue;
      for (let rr = ref.r1; rr <= ref.r2; rr++) for (let cc = ref.c1; cc <= ref.c2; cc++) {
        const sub = expandPlanCell(ctx, rr, cc, depth + 1, seen); res.skus.push(...sub.skus); res.flags.push(...sub.flags);
      }
    }
    return res;
  }
  function skuSize(spec) {
    let m = /^(?:AH|AHS|AES|TW|TM|CM)-(\d+(?:\.\d+)?)/.exec(spec); if (m) return +m[1];
    return null; // AE 吊扇型號與葉片尺寸命名不一致（如 AE-360 用 4M 葉片），不做尺寸比對
  }
  const skuModel = spec => (/^([A-Z]+)-/.exec(spec) || [])[1] || '';
  const hasLight = spec => /-燈$|-燈-|燈$/.test(spec);
  const ambiguous = spec => /-無/.test(spec) && /-燈/.test(spec);
  const colorOf = spec => /-黑/.test(spec) ? '黑' : /-銀/.test(spec) ? '銀' : '';
  function attrFlags(label, skus, all) {
    const flags = []; if (!skus.length) return flags;
    const models = new Set(skus.map(s => skuModel(s.spec)).filter(m => m === 'AH' || m === 'AE'));
    if (models.size !== 1) return flags;
    const model = [...models][0]; const pool = all.filter(s => skuModel(s.spec) === model && s.type === skus[0].type && !ambiguous(s.spec));
    const inSet = new Set(skus); skus = skus.filter(s => !ambiguous(s.spec));
    const check = (want, desc) => {
      const missing = pool.filter(s => want(s.spec) && !inSet.has(s)).map(s => s.spec);
      const extra = skus.filter(s => !want(s.spec)).map(s => s.spec);
      if (missing.length) flags.push(`「${label}」未納入 ${missing.join('、')}（${desc}）`);
      if (extra.length) flags.push(`「${label}」納入了不符條件的 ${extra.join('、')}（${desc}）`);
    };
    if (/有燈/.test(label)) check(hasLight, '有燈款');
    else if (/黑$/.test(label)) check(s => colorOf(s) === '黑', '黑色款');
    else if (/銀$/.test(label)) check(s => colorOf(s) === '銀', '銀色款');
    return flags;
  }
  function lineFlags(line, skus, sheetTitle = '') {
    const flags = []; if (!skus.length) return flags;
    const text = [line.name, line.sub, line.option].join(' ');
    const sizeText = line.option || line.sub || '';
    const sm = /\d\s+\d+(?:\.\d+)?\s*米/.test(sizeText) ? null : /(\d+(?:\.\d+)?)\s*米/.exec(sizeText);
    if (/坦克|側掛|水簾/.test(sheetTitle) && /葉片/.test(line.name)) {
      const bad = [...new Set(skus.filter(s => /^(AE|AH)-/.test(s.spec)).map(s => s.spec))];
      if (bad.length) flags.push(`坦克／側掛／水簾扇的葉片需求取自吊扇型號 ${bad.join('、')}`);
    }
    if (sm) {
      const want = +sm[1]; const bad = [...new Set(skus.filter(s => { const z = skuSize(s.spec); return z !== null && Math.abs(z - want) > 0.01; }).map(s => s.spec))];
      if (bad.length) flags.push(`料件為 ${sm[1]} 米，但需求取自 ${bad.join('、')}`);
    }
    if (/坎燈|崁燈/.test(line.name) && !/花盤|底蓋/.test(line.name)) {
      const bad = [...new Set(skus.filter(s => !hasLight(s.spec)).map(s => s.spec))];
      if (bad.length) flags.push(`燈具需求含無燈款 ${bad.join('、')}`);
    }
    const color = /^(黑|銀)$/.test(line.sub) ? line.sub : /^(黑|銀)$/.test(line.option) ? line.option : '';
    if (color) {
      const bad = [...new Set(skus.filter(s => colorOf(s.spec) && colorOf(s.spec) !== color).map(s => s.spec))];
      if (bad.length) flags.push(`料件為${color}色，但需求含 ${bad.join('、')}`);
    }
    return flags;
  }

  function parseCapacity(S) {
    const anchor = findCells(S, v => v === '工時/分')[0];
    if (!anchor) return null;
    const { r, c } = anchor; const processes = [];
    for (let x = c + 1; x < S.cols; x++) {
      const code = S.raw(r + 2, x); if (!isText(code)) break;
      processes.push({ col: colName(x), c: x, process: flat(S.merged(r - 1, x)) || '', code, minutes: S.raw(r, x), perPerson: S.raw(r + 1, x), schedule: {} });
    }
    const sites = [];
    for (const hit of findCells(S, v => isText(v) && /目前總工時$/.test(v))) {
      const val = firstRight(S, hit.r, hit.c, isNum); const note = firstRight(S, hit.r, hit.c, v => isText(v) && v.length > 12);
      sites.push({ name: hit.v.replace('目前總工時', ''), totalMinutes: val ? val.v : null, note: note ? flat(note.v) : '' });
    }
    const daily = findCells(S, v => v === '日工時')[0];
    const dailyMinutes = daily ? (firstRight(S, daily.r, daily.c, isNum) || {}).v ?? null : null;
    const schedules = [];
    for (const hit of findCells(S, v => isText(v) && /排程量$/.test(v))) {
      const site = hit.v.replace('排程量', ''); schedules.push(site);
      for (const p of processes) { const v = S.raw(hit.r, p.c); if (isNum(v)) p.schedule[site] = v; }
    }
    processes.forEach(p => delete p.c);
    return { processes, sites, dailyMinutes, schedules };
  }

  // ---------- 一般表頭表格 ----------
  function parseTable(S, name) {
    let hr = -1;
    for (let r = 0; r < Math.min(S.rows, 6) && hr < 0; r++) for (let c = 0; c < S.cols; c++) { const v = S.raw(r, c); if (isText(v) && /^(品名|品項編碼|生產品項編碼|規格|項次)$/.test(v)) { hr = r; break; } }
    if (hr < 0) return null;
    const title = (() => { for (let r = 0; r < hr; r++) for (let c = 0; c < S.cols; c++) { const v = S.raw(r, c); if (isText(v) && v !== '返回' && v.length > 3) return flat(v); } return name; })();
    const cols = []; const seen = {};
    for (let c = 0; c < S.cols; c++) {
      let h = S.raw(hr, c); if (h !== null) h = flat(String(h));
      if (!isText(h)) continue;
      seen[h] = (seen[h] || 0) + 1;
      cols.push({ c, col: colName(c), label: seen[h] > 1 ? `${h}#${seen[h]}` : h });
    }
    // 無表頭但有編碼的欄
    for (let c = 0; c < S.cols; c++) {
      if (cols.some(x => x.c === c)) continue;
      let codes = 0, any = 0;
      for (let r = hr + 1; r < S.rows; r++) { const v = S.raw(r, c); if (v !== null) { any++; if (isText(v) && /^[0-9A-Z]{8,12}$/.test(v)) codes++; } }
      if (codes >= 3 && codes >= any * 0.6) cols.push({ c, col: colName(c), label: '品項編碼' });
    }
    cols.sort((a, b) => a.c - b.c);
    const rows = [];
    for (let r = hr + 1; r < S.rows; r++) {
      if (!cols.some(x => S.raw(r, x.c) !== null)) continue;
      const row = { _row: r + 1 };
      for (const x of cols) { const v = S.merged(r, x.c); row[x.label] = v; }
      row._total = cols.some(x => isText(S.raw(r, x.c)) && /^(總計|合計)/.test(S.raw(r, x.c)));
      rows.push(row);
    }
    return { sheet: name, title, headerRow: hr + 1, columns: cols.map(({ col, label }) => ({ col, label })), rows };
  }

  // ---------- BOM 庫存表 ----------
  function parseBom(T, issues, S, ctx, erpMap) {
    const L = T.columns.map(c => c.label);
    const pick = test => L.find(test) || null;
    const k = {
      name: pick(l => l === '品名'), sub: pick(l => l === '下階細項'), option: pick(l => l === '選項擇一'),
      qty: pick(l => l === 'BOM量'), need: pick(l => /安存/.test(l)),
      stock: pick(l => /^庫存(?!\/)/.test(l)), transit: pick(l => /^在途/.test(l)),
      perUnit: pick(l => l === '庫存/台數'), demand: pick(l => /需求台數/.test(l)), diff: pick(l => l === '差異'),
      unreceived: pick(l => /^未進貨/.test(l)), code: pick(l => l === '品項編碼'), note: null
    };
    const notes = L.filter(l => /^備註/.test(l));
    if (!k.perUnit && notes.length > 1) {
      k.perUnit = notes[0];
      issues.push({ level: 'info', sheet: T.sheet, text: '「庫存/台數」欄表頭寫成「備註」，依數值判讀為庫存/台數。' });
    }
    k.note = notes.find(n => n !== k.perUnit) || null;
    const needCol = (T.columns.find(c => c.label === k.need) || {}).col;
    const needC = needCol ? [...needCol].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1 : -1;
    const dateOf = l => (l && /(\d{1,2}\/\d{1,2})/.exec(l) || [])[1] || null;
    const num = v => isNum(v) ? v : null;
    const lines = []; const naRows = [];
    for (const r of T.rows) {
      const qty = num(r[k.qty]);
      const perUnit = r[k.perUnit];
      if (qty === null && !isNum(perUnit)) continue;
      const name = flat(r[k.name]) || ''; let sub = flat(r[k.sub]) || ''; let option = flat(r[k.option]) || '';
      if (option === sub) option = ''; if (sub === name) sub = '';
      if (isErr(perUnit)) naRows.push(r._row);
      const line = {
        row: r._row, name, sub, option, qty, need: num(r[k.need]), stock: num(r[k.stock]), transit: num(r[k.transit]), perUnit: num(perUnit),
        demand: k.demand ? num(r[k.demand]) : null, diff: k.diff ? num(r[k.diff]) : null,
        unreceived: k.unreceived ? num(r[k.unreceived]) : null, code: k.code && isText(r[k.code]) ? r[k.code].trim() : '',
        note: k.note && isText(r[k.note]) ? flat(r[k.note]) : '', errors: ['stock', 'transit', 'perUnit', 'diff', 'unreceived'].filter(f => isErr(r[k[f]])),
        demandSkus: [], demandFormula: null, flags: []
      };
      // 需求公式追溯
      if (S && ctx && needC >= 0) {
        const f = S.formula(r._row - 1, needC); line.demandFormula = f;
        if (f) {
          const skus = []; const flags = [];
          const walk = (formula, depth) => {
            for (const ref of refsIn(formula)) {
              if (ref.sheet && ref.sheet === ctx.name) {
                for (let rr = ref.r1; rr <= ref.r2; rr++) for (let cc = ref.c1; cc <= ref.c2; cc++) { const ex = expandPlanCell(ctx, rr, cc); skus.push(...ex.skus); flags.push(...ex.flags); }
              } else if (!ref.sheet && ref.c1 === needC && ref.r1 !== r._row - 1 && depth < 4) {
                const f2 = S.formula(ref.r1, needC); if (f2) walk(f2, depth + 1);
              }
            }
          };
          walk(f, 0);
          if (/^E\d+\*[^(]/.test(f) && /\+/.test(f)) flags.push('公式缺少括號：只有第一項乘上 BOM量（BOM量為 1 時結果不變）');
          line.demandSkus = [...new Set(skus.map(s => s.spec))];
          line.flags = [...new Set([...flags, ...lineFlags(line, skus, T.title)])];
        }
      }
      // ERP 對帳
      if (erpMap && line.code) {
        const e = erpMap.get(line.code); const same = (a, b) => Math.abs((a || 0) - (b || 0)) < 0.01;
        line.erp = e ? { stock: same(line.stock, e.stock), transit: same(line.transit, e.transit), unreceived: same(line.unreceived, e.unreceived) } : null;
      }
      lines.push(line);
    }
    if (naRows.length) issues.push({ level: 'warn', sheet: T.sheet, text: `第 ${naRows.join('、')} 列「庫存/台數」為錯誤值（多為查無品項編碼）。` });
    const hasDemand = !!k.demand;
    const shortages = lines.filter(l => isNum(l.diff) && l.diff < 0);
    const tight = lines.filter(l => isNum(l.diff) && (l.demand || 0) > 0).sort((a, b) => a.diff - b.diff).slice(0, 5).map(l => l.row);
    const coded = lines.filter(l => l.code && l.erp);
    const erpCheck = { lines: coded.length, mismatch: coded.filter(l => !(l.erp.stock && l.erp.transit && l.erp.unreceived)).length, notFound: lines.filter(l => l.code && !l.erp).length, noCode: lines.filter(l => !l.code).length };
    return {
      sheet: T.sheet, title: T.title, stockDate: dateOf(k.stock), transitDate: dateOf(k.transit), unreceivedDate: dateOf(k.unreceived),
      hasDemand, lines, shortageCount: shortages.length, tightRows: tight, flagCount: lines.filter(l => l.flags.length).length, erpCheck
    };
  }

  // ---------- 成本表 ----------
  function parseCost(T) {
    const g = (r, re) => { const key = T.columns.map(c => c.label).find(l => re.test(l)); return key ? r[key] : null; };
    const lines = []; let total = null;
    for (const r of T.rows) {
      if (r._total) { const v = g(r, /^整組小計$/); if (isNum(v)) total = v; continue; }
      const qty = g(r, /^BOM量$/), price = g(r, /^單價$/), sub = g(r, /^整組小計$/);
      if (!isNum(qty) && !isNum(price) && !isNum(sub)) continue;
      lines.push({
        row: r._row, name: flat(g(r, /^品名$/)) || '', sub: flat(g(r, /^下階細項$/)) || '', option: flat(g(r, /^選項擇一$/)) || '',
        qty: isNum(qty) ? qty : null, price: isNum(price) ? price : null, subtotal: isNum(sub) ? sub : null,
        weight: isNum(g(r, /^材重/)) ? g(r, /^材重/) : null, vendor: flat(g(r, /^廠商$/)) || '',
        targetPrice: isNum(g(r, /更降成本金額/)) ? g(r, /更降成本金額/) : null, newVendor: flat(g(r, /new廠商/)) || '',
        cutPct: isNum(g(r, /降%/)) ? g(r, /降%/) : null, reason: flat(g(r, /原由/)) || ''
      });
    }
    const product = T.title.replace(/-?成本明細表$/, '');
    return { sheet: T.sheet, title: T.title, product, total, lines, isTarget: /降目標/.test(T.sheet) };
  }

  // ---------- 原始工作表 ----------
  function grid(S) {
    const rows = []; let maxC = 0;
    for (let r = 0; r < S.rows; r++) {
      const row = []; for (let c = 0; c < S.cols; c++) row.push(S.raw(r, c));
      while (row.length && row[row.length - 1] === null) row.pop();
      maxC = Math.max(maxC, row.length); rows.push(row);
    }
    while (rows.length && !rows[rows.length - 1].length) rows.pop();
    return { rows, cols: maxC };
  }

  function parseWorkbook(wb, meta = {}) {
    const issues = [];
    const sheets = wb.SheetNames.map(n => n);
    const acc = {}; for (const n of sheets) acc[n] = sheetAccess(wb.Sheets[n], n);
    const pick = re => sheets.find(n => re.test(n.trim()));
    const out = { schema: SCHEMA, source: { file: meta.file || '', parsedAt: meta.parsedAt || new Date().toISOString(), sheetCount: sheets.length }, sheets: sheets.map(n => n.trim()), plan: null, forecast: null, bom: [], cost: { summary: [], products: [] }, erp: { items: [], counts: {}, duplicates: [] }, tables: {}, raw: {}, issues };

    const planName = pick(/產銷表$/), fcName = pick(/次月預估$/);
    const planCtx = {};
    if (planName) out.plan = parsePlan(acc[planName], planName.trim(), issues, planCtx); else issues.push({ level: 'error', sheet: '', text: '找不到「國工風扇產銷表」工作表。' });
    if (fcName) out.forecast = parsePlan(acc[fcName], fcName.trim(), issues);
    if (out.plan && out.forecast && out.plan.date && out.forecast.date && out.forecast.date < out.plan.date)
      issues.push({ level: 'warn', sheet: fcName.trim(), text: `次月預估表頭日期為 ${out.forecast.date}，早於產銷表 ${out.plan.date}，日期欄可能未更新。` });

    const bomTables = [];
    for (const n of sheets) {
      const S = acc[n]; const key = n.trim();
      out.raw[key] = grid(S);
      if (n === planName || n === fcName) continue;
      const T = parseTable(S, key);
      if (!T) continue;
      if (/BOM庫存表$/.test(key)) bomTables.push([T, S]);
      else if (/成品成本表$|成品降目標$/.test(key)) out.cost.products.push(parseCost(T));
      else if (/成本總表$/.test(key)) out.cost.summary = T.rows.filter(r => isText(r['品名'])).map(r => ({ name: r['品名'], price: isNum(r['單價']) ? r['單價'] : null, note: r['備註'] || '' }));
      else out.tables[key] = T;
    }
    // 成本總表與明細表總計比對
    for (const s of out.cost.summary) {
      const p = out.cost.products.find(p => !p.isTarget && p.product.replace(/-燈/, '') === s.name.replace(/-燈/, ''));
      if (p && isNum(p.total) && isNum(s.price) && Math.abs(p.total - s.price) > 1) issues.push({ level: 'info', sheet: '國工風扇各成品成本總表', text: `${s.name} 總表單價 ${round(s.price, 2)}，明細表總計 ${round(p.total, 2)}，差 ${round(s.price - p.total, 2)}（可能另含加工或管銷，請確認口徑）。` });
    }

    // ERP 物料合併
    const items = new Map();
    const touch = (code, name) => { if (!items.has(code)) items.set(code, { code, name: name || '', stock: null, transit: null, unreceived: null, notes: [] }); const it = items.get(code); if (!it.name && name) it.name = name; return it; };
    const erp = (re, field, qtyLabel) => {
      const n = sheets.find(x => re.test(x.trim())); if (!n) return 0;
      const T = out.tables[n.trim()]; if (!T) return 0; const seen = new Map(); let count = 0;
      const qk = T.columns.map(c => c.label).find(l => qtyLabel.test(l));
      for (const r of T.rows) {
        const code = r['品項編碼']; if (!isText(code)) continue; count++;
        const it = touch(code.trim(), flat(r['品項名稱'])); const q = isNum(r[qk]) ? r[qk] : 0;
        seen.set(code, (seen.get(code) || 0) + 1);
        it[field] = (it[field] || 0) + q;
        const note = T.columns.map(c => c.label).filter(l => !/品項|數量|合計/.test(l)).map(l => r[l]).find(isText);
        if (note) it.notes.push(flat(note));
      }
      for (const [code, n2] of seen) if (n2 > 1) out.erp.duplicates.push({ sheet: n.trim(), code, times: n2 });
      return count;
    };
    out.erp.counts.stock = erp(/^ERP庫存$/, 'stock', /合計|數量/);
    out.erp.counts.transit = erp(/^ERP在途$/, 'transit', /數量/);
    out.erp.counts.unreceived = erp(/^ERP未進貨$/, 'unreceived', /未進貨數量|數量/);
    out.erp.items = [...items.values()].sort((a, b) => a.code.localeCompare(b.code));
    if (out.erp.duplicates.length) issues.push({ level: 'info', sheet: 'ERP未進貨', text: `${out.erp.duplicates.length} 個品項編碼重複出現，看板已合併加總。` });

    // BOM 齊套（需 ERP 物料與產銷表公式）
    const erpMap = new Map(out.erp.items.map(i => [i.code, i]));
    for (const [T, S] of bomTables) out.bom.push(parseBom(T, issues, S, planCtx, erpMap));
    for (const b of out.bom) if (b.erpCheck.mismatch) issues.push({ level: 'warn', sheet: b.sheet, text: `${b.erpCheck.mismatch} 筆料件的庫存／在途／未進貨與 ERP 工作表不一致。` });

    // 各表錯誤值統計
    for (const n of sheets) {
      const g = out.raw[n.trim()]; let e = 0; g.rows.forEach(r => r.forEach(v => { if (isErr(v)) e++; }));
      if (e) issues.push({ level: 'info', sheet: n.trim(), text: `含 ${e} 個錯誤值（#N/A、#DIV/0! 等），多為 VLOOKUP 查無編碼或分母為 0。` });
    }
    return out;
  }

  const api = { parseWorkbook, SCHEMA, colName };
  root.AMAParser = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window === 'undefined' ? globalThis : window);
