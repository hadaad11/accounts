// Turns the Excel workbooks (read with SheetJS) into the dashboard's tab data.
// Shared by the page (live sync) and build/build.mjs (snapshot).
(function (root) {
  const TOTAL_WORDS = new Set(['total', 'الإجمالي', 'الاجمالي', 'إجمالي', 'اجمالي', 'المجموع']);
  const isNum = v => typeof v === 'number' && isFinite(v);
  const colIdx = s => [...s].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const REF_RE = /\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/g;
  const key = (r, c) => r * 100000 + c;

  // ---------------------------------------------------------- formulas
  // =A1*B1, =E4-(E8+E12), =SUM(C3:C8), =80000+15000 ...
  function parseFormula(f) {
    const s = String(f).replace(/^=/, '');
    let i = 0;
    const peek = () => { while (s[i] === ' ') i++; return s[i]; };
    function primary() {
      peek();
      const rest = s.slice(i);
      let m;
      if ((m = /^SUM\(/i.exec(rest))) {
        i += m[0].length;
        const args = [];
        if (peek() !== ')') { args.push(expr()); while (peek() === ',') { i++; args.push(expr()); } }
        if (peek() !== ')') throw 0;
        i++;
        return {t: 'sum', args};
      }
      if ((m = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/.exec(rest))) {
        i += m[0].length;
        return m[3] ? {t: 'rng', r1: +m[2], c1: colIdx(m[1]), r2: +m[4], c2: colIdx(m[3])} : {t: 'ref', r: +m[2], c: colIdx(m[1])};
      }
      if ((m = /^\d+(\.\d+)?/.exec(rest))) { i += m[0].length; return {t: 'num', v: +m[0]}; }
      if (rest[0] === '(') { i++; const e = expr(); if (peek() !== ')') throw 0; i++; return e; }
      if (rest[0] === '-') { i++; return {t: 'neg', a: primary()}; }
      if (rest[0] === '+') { i++; return primary(); }
      throw 0;
    }
    function term() { let a = primary(); while (peek() === '*' || peek() === '/') { const op = s[i++]; a = {t: 'bin', op, a, b: primary()}; } return a; }
    function expr() { let a = term(); while (peek() === '+' || peek() === '-') { const op = s[i++]; a = {t: 'bin', op, a, b: term()}; } return a; }
    try { const e = expr(); peek(); return i === s.length ? e : null; } catch (_) { return null; }
  }

  const signedTerms = (n, sign = 1) => n.t === 'bin' && (n.op === '+' || n.op === '-')
    ? signedTerms(n.a, sign).concat(signedTerms(n.b, n.op === '-' ? -sign : sign)) : [[sign, n]];

  function constTerms(f) {
    const tree = f && parseFormula(f);
    if (!tree) return null;
    const terms = signedTerms(tree);
    return terms.every(([, n]) => n.t === 'num') ? terms.map(([s, n]) => s * n.v) : null;
  }

  // ---------------------------------------------------------- sheet model
  function Sheet(ws, name) {
    this.name = name;
    this.cells = new Map();
    for (const k of Object.keys(ws)) {
      const m = /^([A-Z]+)(\d+)$/.exec(k);
      if (!m) continue;
      const x = ws[k], f = x.f ? '=' + x.f : null;
      let v = x.t === 'e' ? x.w : x.t === 'z' ? undefined : x.v;
      if (typeof v === 'string') { v = v.trim(); if (!v) continue; }
      if (v === undefined || v === null) { if (!f) continue; v = null; }
      this.cells.set(key(+m[2], colIdx(m[1])), {r: +m[2], c: colIdx(m[1]), v, f});
    }
    for (const cell of this.cells.values()) if (cell.f && cell.v === null) cell.v = this.evaluate(cell.f);
  }
  Sheet.prototype.value = function (r, c) {
    const cell = this.cells.get(key(r, c));
    if (!cell) return 0;
    if (cell.v === null && cell.f) cell.v = this.evaluate(cell.f);
    return isNum(cell.v) ? cell.v : 0;
  };
  Sheet.prototype.evaluate = function (f) {
    const tree = parseFormula(f);
    if (!tree) return null;
    const ev = n => {
      switch (n.t) {
        case 'num': return n.v;
        case 'ref': return this.value(n.r, n.c);
        case 'neg': return -ev(n.a);
        case 'rng': { let s = 0; for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) s += this.value(r, c); return s; }
        case 'sum': return n.args.reduce((a, x) => a + ev(x), 0);
        case 'bin': { const a = ev(n.a), b = ev(n.b); return n.op === '+' ? a + b : n.op === '-' ? a - b : n.op === '*' ? a * b : b ? a / b : NaN; }
      }
    };
    const v = ev(tree);
    return isNum(v) ? v : null;
  };
  Sheet.prototype.rowNums = function () { return [...new Set([...this.cells.values()].map(x => x.r))].sort((a, b) => a - b); };
  Sheet.prototype.row = function (r) { return [...this.cells.values()].filter(x => x.r === r).sort((a, b) => a.c - b.c); };
  Sheet.prototype.blocks = function () {
    const out = [];
    for (const r of this.rowNums()) {
      const last = out[out.length - 1];
      if (last && r === last[last.length - 1] + 1) last.push(r); else out.push([r]);
    }
    return out;
  };
  // each row -> first text cell as label, last number cell as value
  Sheet.prototype.pairs = function () {
    const out = [];
    for (const r of this.rowNums()) {
      const cells = this.row(r);
      const lbl = cells.find(x => typeof x.v === 'string');
      const nums = cells.filter(x => isNum(x.v));
      if (nums.length) out.push({label: lbl ? lbl.v : null, value: nums[nums.length - 1].v, cell: nums[nums.length - 1]});
    }
    return out;
  };

  // ---------------------------------------------------------- receivables workbook
  const isTotal = l => typeof l === 'string' && TOTAL_WORDS.has(l.trim().toLowerCase());

  function parseCompany(sh) {
    const rows = sh.pairs().map(p => ({...p, label: !p.label || p.label.toLowerCase() === sh.name.toLowerCase() ? null : p.label}));
    let seq = 0;
    const label = l => l || `دفعة ${++seq}`;
    const ti = rows.findIndex(x => isTotal(x.label));
    if (ti >= 0) {
      return [rows[ti].value, {
        kind: 'items',
        items: rows.slice(0, ti).filter(x => x.value).map(x => ({label: label(x.label), value: x.value})),
        notes: rows.slice(ti + 1).filter(x => x.label).map(x => ({label: x.label, value: x.value})),
      }];
    }
    if (rows.length === 3 && !rows.some(x => x.label) && Math.abs(rows[0].value - rows[1].value - rows[2].value) < 0.5) {
      const [due, paid, rest] = rows;
      const dueItems = constTerms(due.cell.f), paidItems = constTerms(paid.cell.f);
      return [rest.value, {
        kind: 'contract', items: [], notes: [], due: due.value, paid: paid.value,
        due_items: (dueItems && dueItems.length ? dueItems : [due.value]).map((v, i) => ({label: `بند ${i + 1}`, value: v})),
        paid_items: (paidItems && paidItems.length ? paidItems : [paid.value]).map((v, i) => ({label: `دفعة ${i + 1}`, value: v})),
      }];
    }
    const items = rows.filter(x => x.value).map(x => ({label: label(x.label), value: x.value}));
    return [items.reduce((s, i) => s + i.value, 0), {kind: 'items', items, notes: []}];
  }

  function parseReceivables(sheets) {
    const people = [], companies = [], debts = [];
    for (const sh of sheets) {
      if (/دين|ديون|التزام/.test(sh.name)) {
        sh.pairs().filter(p => p.label && !isTotal(p.label)).forEach(p => debts.push({name: p.label, value: p.value}));
      } else if (/[أا]فراد/.test(sh.name)) {
        sh.pairs().filter(p => p.label && !isTotal(p.label)).forEach(p => people.push({name: p.label, type: 'person', value: p.value}));
      } else if (sh.cells.size) {
        const [value, detail] = parseCompany(sh);
        companies.push({name: sh.name, type: 'company', value, detail});
      }
    }
    const receivables = people.concat(companies);
    const tr = receivables.reduce((s, r) => s + r.value, 0), td = debts.reduce((s, d) => s + d.value, 0);
    return {receivables, debts, total_receivable: tr, total_debt: td, net: tr - td};
  }

  // ---------------------------------------------------------- generic workbook
  const byCol = cells => new Map(cells.map(x => [x.c, x]));

  // two rows: labels over numbers, one number is a formula over the others
  function asCalc(sh, blk) {
    if (blk.length !== 2) return null;
    const head = byCol(sh.row(blk[0])), vals = byCol(sh.row(blk[1]));
    const numc = [...vals.values()].filter(x => isNum(x.v)).map(x => x.c);
    const fcs = numc.filter(c => vals.get(c).f);
    if (fcs.length !== 1 || numc.length < 3 || !numc.every(c => head.has(c) && typeof head.get(c).v === 'string')) return null;
    const fc = fcs[0], body = vals.get(fc).f;
    const refs = [...body.matchAll(REF_RE)].map(m => [+m[2], colIdx(m[1])]);
    if (!refs.length || refs.some(([r, c]) => r !== blk[1] || !numc.includes(c))) return null;
    const merged = new Map([...vals, ...head]);
    const t = [...merged.keys()].sort((a, b) => a - b).find(c => !numc.includes(c) && typeof merged.get(c).v === 'string');
    const parts = numc.filter(c => c !== fc), cols = parts.concat([fc]);
    return {
      title: t === undefined ? null : merged.get(t).v,
      columns: cols.map(c => head.get(c).v), rows: [cols.map(c => ({n: vals.get(c).v}))],
      kind: 'calc', op: body.includes('*') && !body.includes('+') ? 'x' : '+',
      total: vals.get(fc).v, total_label: head.get(fc).v,
      parts: parts.map(c => ({label: head.get(c).v, value: vals.get(c).v})),
      _at: key(blk[1], fc),
    };
  }

  // one row: a label and a formula over calc totals -> a KPI with a waterfall
  function asResult(sh, blk, regions) {
    if (blk.length !== 1) return null;
    const cells = sh.row(blk[0]);
    const texts = cells.filter(x => typeof x.v === 'string'), forms = cells.filter(x => x.f && isNum(x.v));
    if (texts.length !== 1 || forms.length !== 1 || cells.length !== 2) return null;
    const kpi = {label: texts[0].v, value: forms[0].v};
    const tree = parseFormula(forms[0].f);
    const totals = new Map(regions.map((r, i) => [r._at, i]).filter(([at]) => at !== undefined));
    const terms = tree ? signedTerms(tree) : [];
    const at = n => n.t === 'ref' ? key(n.r, n.c) : null;
    if (terms.length && terms.every(([, n]) => totals.has(at(n)))) {
      kpi.steps = terms.map(([s, n]) => { const i = totals.get(at(n)), r = regions[i];
        return {label: r.title || r.total_label, region: i, value: r.total, sign: s}; });
    }
    return kpi;
  }

  function asTable(sh, blk, title) {
    const cols = [...new Set([...sh.cells.values()].filter(x => blk.includes(x.r)).map(x => x.c))].sort((a, b) => a - b);
    const first = byCol(sh.row(blk[0]));
    const hasHead = blk.length > 1 && [...first.values()].every(x => typeof x.v === 'string');
    const body = hasHead ? blk.slice(1) : blk;
    const defaults = ['البند', 'التفاصيل'];
    const columns = cols.map((c, i) => hasHead && first.has(c) ? first.get(c).v : defaults.length ? defaults.shift() : `عمود ${i + 1}`);
    const rows = body.map(r => { const cells = byCol(sh.row(r));
      return cols.map(c => !cells.has(c) ? null : isNum(cells.get(c).v) ? {n: cells.get(c).v} : {s: String(cells.get(c).v)}); });
    const emph = cols.map((c, i) => body.some(r => (sh.cells.get(key(r, c)) || {}).f) ? i : -1).filter(i => i >= 0);
    const reg = {title, columns, rows, kind: 'table', emph};
    if (emph.length) reg.totals = Object.fromEntries(emph.map(i => [String(i), rows.reduce((s, x) => s + (x[i] && 'n' in x[i] ? x[i].n : 0), 0)]));
    return reg;
  }

  function parseGeneric(sheets) {
    const out = [];
    for (const sh of sheets) {
      if (!sh.cells.size) continue;
      const regions = [], kpis = [];
      let title = null;
      for (const blk of sh.blocks()) {
        const calc = asCalc(sh, blk);
        if (calc) { regions.push(calc); continue; }
        const kpi = asResult(sh, blk, regions);
        if (kpi) { kpis.push(kpi); continue; }
        const one = sh.row(blk[0]);
        if (blk.length === 1 && one.length === 1 && typeof one[0].v === 'string') { title = one[0].v; continue; }
        regions.push(asTable(sh, blk, title));
        title = null;
      }
      if (!kpis.length) {
        for (const r of regions) if (r.kind === 'table' && r.emph.length) {
          const i = r.emph[r.emph.length - 1];
          kpis.push({label: `إجمالي ${r.columns[i]}`, value: r.totals[String(i)]});
        }
      }
      regions.forEach(r => delete r._at);
      out.push({name: sh.name, regions, kpis});
    }
    return {sheets: out};
  }

  // wb: SheetJS workbook read with {sheetStubs: true} (keeps formulas that have no saved result); tab: {kind, rename}
  function parseWorkbook(wb, tab) {
    const rename = tab.rename || {};
    const sheets = wb.SheetNames.map(n => new Sheet(wb.Sheets[n], rename[n.trim()] || n.trim()));
    return tab.kind === 'receivables' ? parseReceivables(sheets) : parseGeneric(sheets);
  }

  // key-order-independent JSON, for "did the data change?"
  const stable = v => Array.isArray(v) ? '[' + v.map(stable).join(',') + ']'
    : v && typeof v === 'object' ? '{' + Object.keys(v).sort().map(k => JSON.stringify(k) + ':' + stable(v[k])).join(',') + '}'
    : JSON.stringify(v);

  const api = {parseWorkbook, stable};
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.AccParse = api;
})(typeof self !== 'undefined' ? self : this);
