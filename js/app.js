/*
 * app.js — หน้าเว็บกรอกแบบรายงานผล IT Audit (ก.ล.ต.) รอบปี 2569
 *
 * ข้อมูลทั้งหมดอยู่ในหน่วยความจำของแท็บนี้เท่านั้น:
 *   - ไม่ใช้ localStorage / sessionStorage / IndexedDB / cookie
 *   - ไม่มีการเรียก network (บังคับด้วย Content-Security-Policy ใน index.html)
 *   - ผู้ใช้เก็บงานเองด้วยการ "บันทึกงาน (JSON)" และ "Export xlsx" ลงเครื่อง
 */
'use strict';

(() => {
  const T = TEMPLATE;
  const SCHEMA_VERSION = 1;

  // ---------------- state ----------------
  let wb = null;          // workbook ต้นฉบับ (อยู่ในหน่วยความจำ)
  let tplFileName = '';
  let tpl = null;         // โครงแบบฟอร์มที่อ่านจากไฟล์
  let state = freshState();
  let dirty = false;
  let tab = 'start';
  let ctlSheet = null;
  const filters = { rla: 'auto', status: 'all', q: '' };

  function freshState() {
    return { cells: {}, findings: [], notes: {} };
  }

  // ---------------- DOM helpers ----------------
  function h(tag, props, ...kids) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(props || {})) {
      if (v === null || v === undefined || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k === 'text') el.textContent = v;
      else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
      else if (k === 'value') el.value = v;
      else if (['checked', 'disabled', 'open', 'selected'].includes(k)) el[k] = !!v;
      else el.setAttribute(k, v === true ? '' : v);
    }
    for (const kid of kids.flat(Infinity)) {
      if (kid === null || kid === undefined || kid === false) continue;
      el.append(kid instanceof Node ? kid : document.createTextNode(String(kid)));
    }
    return el;
  }
  const $ = sel => document.querySelector(sel);
  const clean = s => (s === null || s === undefined) ? '' : String(s).replace(/ /g, ' ').trim();

  function toast(msg, kind = 'info') {
    const el = h('div', { class: 'toast ' + kind, text: msg });
    $('#toasts').append(el);
    setTimeout(() => el.remove(), kind === 'error' ? 8000 : 4000);
  }

  function markDirty() {
    dirty = true;
    updateHeader();
  }

  // ---------------- value access ----------------
  function getVal(sheet, ref) {
    return (state.cells[sheet] && state.cells[sheet][ref]) || '';
  }
  function setVal(sheet, ref, v) {
    if (!state.cells[sheet]) state.cells[sheet] = {};
    if (v === '' || v === null || v === undefined) delete state.cells[sheet][ref];
    else state.cells[sheet][ref] = v;
    markDirty();
  }

  // ---------------- template parsing ----------------
  async function parseTemplate(workbook) {
    const t = { fields: {}, lists: {}, controls: [], controlsById: {}, ctlSheets: [] };
    const val = (sheet, ref) => XL.value(workbook, sheet, ref);
    const validations = {};
    async function optionsFor(sheet, ref) {
      if (!validations[sheet]) validations[sheet] = await XL.listValidations(workbook, sheet);
      const v = XL.findValidation(validations[sheet], ref);
      return v ? XL.resolveList(workbook, v.formula) : [];
    }
    function reg(sheet, ref, def) {
      if (!t.fields[sheet]) t.fields[sheet] = {};
      t.fields[sheet][ref] = def;
    }

    for (const s of [T.basic.sheet, T.systems.sheet, T.scope.sheet, T.findings.sheet, T.controls.rawSheet]) {
      if (!workbook.sheets[s]) throw new Error(`ไม่พบ sheet "${s}" — ไฟล์นี้อาจไม่ใช่แบบรายงานผล IT Audit ของ ก.ล.ต. รอบปี 2569`);
    }

    t.version = clean(await val(T.versionCell.sheet, T.versionCell.ref));
    t.title = clean(await val(T.titleCell.sheet, T.titleCell.ref));

    // --- 1. Basic info ---
    const B = T.basic;
    t.basicGroups = [];
    for (const g of B.groups) {
      const grp = { title: g.title, note: g.noteRef ? clean(await val(B.sheet, g.noteRef)) : '', foot: g.footRef ? clean(await val(B.sheet, g.footRef)) : '', fields: [] };
      for (const f of g.fields) {
        const def = {
          ref: f.ref, type: f.type, wide: !!f.wide,
          label: f.label || clean(await val(B.sheet, f.labelRef)),
          hint: f.hintRef ? clean(await val(B.sheet, f.hintRef)) : '',
        };
        if (f.type === 'list' || f.type === 'datalist') def.options = await optionsFor(B.sheet, f.ref);
        reg(B.sheet, f.ref, def);
        grp.fields.push(def);
      }
      t.basicGroups.push(grp);
    }
    const A = B.auditors;
    t.auditorCols = [];
    for (const c of A.cols) {
      const def = { col: c.col, type: c.type, label: clean(await val(B.sheet, c.col + A.headerRow)) };
      if (c.type === 'list') def.options = await optionsFor(B.sheet, c.col + A.rows[0]);
      t.auditorCols.push(def);
      for (const r of A.rows) reg(B.sheet, c.col + r, { type: c.type });
    }
    t.auditorExamples = [];
    for (const r of A.exampleRows) t.auditorExamples.push(await Promise.all(A.cols.map(c => val(B.sheet, c.col + r).then(clean))));

    // --- 2. Systems ---
    const SY = T.systems;
    t.sysNotes = [];
    for (const r of SY.noteRefs) { const v = clean(await val(SY.sheet, r)); if (v) t.sysNotes.push(v); }
    t.sysGroups = [];
    for (const g of SY.groups) {
      const limit = clean(await val(SY.sheet, 'D' + g.header)) || clean(await val(SY.sheet, 'C' + g.header));
      const grp = {
        title: clean(await val(SY.sheet, 'A' + g.header)),
        limit,
        desc: g.desc ? clean(await val(SY.sheet, 'A' + g.desc)) : '',
        example: clean(await val(SY.sheet, 'B' + (g.rows[0] - 1))),
        rows: g.rows, other: g.other,
        options: await optionsFor(SY.sheet, 'B' + g.rows[0]),
      };
      for (const r of g.rows) {
        reg(SY.sheet, 'B' + r, { type: grp.options.length ? 'list' : 'text' });
        if (g.other) reg(SY.sheet, 'C' + r, { type: 'text' });
      }
      t.sysGroups.push(grp);
    }

    // --- 3. Scope ---
    const SC = T.scope;
    t.scopeNote = clean(await val(SC.sheet, SC.noteRef));
    t.scopeCols = [];
    for (const c of SC.cols) {
      const def = { col: c.col, type: c.type, label: clean(await val(SC.sheet, c.col + SC.headerRow)) };
      if (c.type === 'list') def.options = await optionsFor(SC.sheet, c.col + SC.rows[0]);
      t.scopeCols.push(def);
      for (const r of SC.rows) reg(SC.sheet, c.col + r, { type: c.type });
    }
    t.scopeExamples = [];
    for (const r of SC.exampleRows) t.scopeExamples.push(await Promise.all(SC.cols.map(c => val(SC.sheet, c.col + r).then(clean))));

    // --- RawData: CategoryID, AdvControl ---
    const raw = await XL.cells(workbook, T.controls.rawSheet);
    const rawInfo = {};
    for (const [ref, c] of raw) {
      const m = /^A(\d+)$/.exec(ref);
      if (!m || typeof c.v !== 'string') continue;
      const r = m[1];
      const cat = raw.get('E' + r), adv = raw.get('F' + r);
      rawInfo[c.v.trim()] = { cat: cat ? String(cat.v) : '', adv: adv ? adv.v === 'Y' : false };
    }

    // --- 4. Controls (D1–D3) ---
    const C = T.controls;
    for (const name of workbook.sheetOrder.filter(n => C.sheetPattern.test(n))) {
      const map = await XL.cells(workbook, name);
      const get = ref => { const c = map.get(ref); return c ? c.v : null; };
      let maxRow = 0;
      for (const ref of map.keys()) { const r = XL.splitRef(ref).row; if (r > maxRow) maxRow = r; }
      const sh = { name, short: name.replace(/^4\.\s*/, ''), title: clean(get('F2')), items: [], controls: [] };
      for (let r = C.firstRow; r <= maxRow; r++) {
        const id = clean(get(C.col.id + r));
        if (id && C.idPattern.test(id)) {
          const type = clean(get(C.col.type + r));
          const ctl = {
            id, sheet: name, row: r, cell: C.col.result + r,
            appl: { H: clean(get(C.col.H + r)).toLowerCase() === 'x', M: clean(get(C.col.M + r)).toLowerCase() === 'x', L: clean(get(C.col.L + r)).toLowerCase() === 'x', S: clean(get(C.col.S + r)).toLowerCase() === 'x' },
            sor: clean(get(C.col.sor + r)), np: clean(get(C.col.np + r)), ref: clean(get(C.col.ref + r)),
            text: String(get(C.col.text + r) || '').trim(),
            type: /maturity/i.test(type) ? 'Maturity' : (type ? 'YPN' : ''),
            mat: C.col.mat.map(col => String(get(col + r) || '').trim()),
            cat: rawInfo[id] ? rawInfo[id].cat : '',
            adv: rawInfo[id] ? rawInfo[id].adv : false,
          };
          // แถวที่ไม่มี "ประเภทการวัดผล" = แบบฟอร์มล็อกค่าไว้ (เช่น อ้างอิงข้ออื่น) → แสดงอย่างเดียว
          ctl.fixed = ctl.type ? null : (clean(get(C.col.result + r)) || 'N/A');
          sh.items.push({ kind: 'c', ctl });
          sh.controls.push(ctl);
          t.controls.push(ctl);
          t.controlsById[id] = ctl;
        } else if (!id) {
          const f = clean(get(C.col.sor + r));
          if (f && r > 2 && !/^[\d.()]+$/.test(f)) sh.items.push({ kind: 'h', level: /^ส่วนที่/.test(f) ? 2 : 3, text: f });
        }
      }
      t.ctlSheets.push(sh);
    }

    // --- 5. Finding ---
    const F = T.findings;
    t.findNote = clean(await val(F.sheet, F.noteRef));
    t.findCols = [];
    for (const c of F.cols) {
      const def = { col: c.col, key: c.key, type: c.type, label: clean(await val(F.sheet, c.col + F.headerRow)) };
      if (c.type === 'list') def.options = await optionsFor(F.sheet, c.col + F.rows[0]);
      t.findCols.push(def);
    }
    t.findColByKey = Object.fromEntries(t.findCols.map(c => [c.key, c]));
    t.findExamples = [];
    for (const r of F.exampleRows) {
      const ex = {};
      for (const c of F.cols) ex[c.key] = await val(F.sheet, c.col + r);
      if (typeof ex.H === 'number') ex.H = XL.serialToISO(ex.H).slice(0, 7);
      t.findExamples.push(ex);
    }
    return t;
  }

  /** อ่านค่าที่กรอกอยู่แล้วในไฟล์ (กรณีโหลดไฟล์ที่เคย export ไว้) */
  async function readExistingValues(workbook, t) {
    const st = freshState();
    for (const [sheet, defs] of Object.entries(t.fields)) {
      for (const [ref, def] of Object.entries(defs)) {
        let v = await XL.value(workbook, sheet, ref);
        if (v === null || v === undefined || v === '') continue;
        if (def.type === 'date' && typeof v === 'number') v = XL.serialToISO(v);
        else if (def.type === 'taxid' && typeof v === 'number') v = String(Math.round(v)).padStart(13, '0');
        else v = String(v);
        if (!st.cells[sheet]) st.cells[sheet] = {};
        st.cells[sheet][ref] = v;
      }
    }
    for (const c of t.controls) {
      if (c.fixed) continue;
      const v = clean(await XL.value(workbook, c.sheet, c.cell));
      if (v) { if (!st.cells[c.sheet]) st.cells[c.sheet] = {}; st.cells[c.sheet][c.cell] = v; }
    }
    const F = T.findings;
    for (const r of F.rows) {
      const f = newFinding();
      let any = false;
      for (const c of F.cols) {
        let v = await XL.value(workbook, F.sheet, c.col + r);
        if (v === null || v === '') continue;
        if (c.type === 'month' && typeof v === 'number') v = XL.serialToISO(v).slice(0, 7);
        f[c.key] = String(v);
        any = true;
      }
      if (any) st.findings.push(f);
    }
    return st;
  }

  // ---------------- control helpers ----------------
  function ctlValue(c) { return c.fixed || getVal(c.sheet, c.cell); }
  function rlaCol() {
    if (filters.rla === 'all') return null;
    if (filters.rla !== 'auto') return filters.rla;
    return T.rlaToCol[getVal(T.basic.sheet, 'C26')] || null;
  }
  function isApplicable(c, col = rlaCol()) { return !col || c.appl[col]; }
  function resultClass(v) {
    if (['Yes', 'M3', 'M4', 'M5'].includes(v)) return 'ok';
    if (['Partial', 'M2'].includes(v)) return 'part';
    if (['No', 'M1'].includes(v)) return 'gap';
    if (v === 'N/A') return 'na';
    return '';
  }
  const isGap = v => ['No', 'Partial', 'M1', 'M2'].includes(v);
  function sheetProgress(sh, col = rlaCol()) {
    let total = 0, done = 0;
    for (const c of sh.controls) {
      if (c.fixed || !isApplicable(c, col)) continue;
      total++;
      if (ctlValue(c)) done++;
    }
    return { total, done };
  }
  function domainFor(id) {
    const opts = tpl.findColByKey.B.options || [];
    const prefix = id.startsWith('D1') ? 'D1.' : id.startsWith('D2') ? 'D2.' : id.startsWith('D3') ? 'D3.' : '';
    return opts.find(o => o.startsWith(prefix)) || '';
  }
  function topicFor(id) {
    // ลำดับหัวข้อใน dropdown ของแบบฟอร์ม: 1.1, 1.2, 1.3, 2.1 … 2.12, 3 (รวม 16 หัวข้อ)
    const opts = tpl.findColByKey.D.options || [];
    if (opts.length !== 16) return '';
    const c = tpl.controlsById[id];
    if (!c) return '';
    if (c.sheet.includes('D1')) return opts[{ '1.1': 0, '1.2': 1, '1.3': 2 }[c.cat]] || '';
    const m = /D2\.(\d+)/.exec(c.sheet);
    if (m) return opts[2 + parseInt(m[1], 10)] || '';
    if (c.sheet.startsWith('D3')) return opts[15];
    return '';
  }
  function parseIds(s) { return clean(s).split(/[,;\n]+/).map(x => x.trim().replace(/\s+/g, ' ').replace(/\s*#\s*/, ' #')).filter(Boolean); }
  function findingsForControl(id) {
    return state.findings.map((f, i) => [f, i]).filter(([f]) => parseIds(f.C).includes(id));
  }

  // ---------------- findings ----------------
  let uidSeq = 1;
  function newFinding() {
    return { uid: 'f' + Date.now().toString(36) + (uidSeq++), B: '', C: '', D: '', E: '', F: '', G: '', H: '', I: '', status: 'Open', closed: '' };
  }
  function thisMonth() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  }
  function isOverdue(f) { return f.status !== 'Closed' && f.H && f.H < thisMonth(); }

  // ---------------- header ----------------
  function updateHeader() {
    const info = $('#tplInfo');
    info.textContent = tpl ? `แบบฟอร์ม: ${tpl.version || '(ไม่ระบุรอบ)'} · ${tplFileName}` : 'ยังไม่ได้โหลดแบบฟอร์ม';
    const d = $('#dirtyInfo');
    d.textContent = dirty ? '● มีการแก้ไขที่ยังไม่ได้บันทึกงาน (JSON)' : (tpl ? '✓ บันทึกงานแล้ว / ยังไม่มีการแก้ไข' : '');
    d.className = dirty ? 'dirty' : 'clean';
    for (const id of ['#btnSaveJson', '#btnExport', '#btnLoadJson']) $(id).disabled = !tpl;
    document.querySelectorAll('#tabs button').forEach(b => { b.disabled = !tpl && b.dataset.tab !== 'start'; b.classList.toggle('active', b.dataset.tab === tab); });
  }

  function go(newTab) {
    tab = newTab;
    updateHeader();
    render();
    window.scrollTo(0, 0);
  }

  // ---------------- render dispatcher ----------------
  function render() {
    const main = $('#main');
    main.replaceChildren();
    const fn = { start: renderStart, basic: renderBasic, systems: renderSystems, scope: renderScope, controls: renderControls, findings: renderFindings, check: renderCheck }[tab];
    main.append(fn());
  }

  // ---------------- input builders ----------------
  function selectEl(options, value, onChange, blank = '— เลือก —') {
    const sel = h('select', { onchange: e => onChange(e.target.value) },
      blank === null ? null : h('option', { value: '' }, blank),
      (options || []).map(o => h('option', { value: o }, clean(o))));
    if (value && !(options || []).includes(value)) sel.append(h('option', { value }, value + ' (ค่าที่ไม่อยู่ในรายการ)'));
    sel.value = value || '';
    return sel;
  }
  let dlSeq = 0;
  function inputFor(type, value, onChange, opts = {}) {
    switch (type) {
      case 'list': return selectEl(opts.options, value, onChange);
      case 'datalist': {
        const id = 'dl' + (dlSeq++);
        return h('span', { class: 'dl' },
          h('input', { type: 'text', list: id, value: value || '', oninput: e => onChange(e.target.value), placeholder: opts.placeholder || 'พิมพ์เพื่อค้นหา หรือพิมพ์ชื่อเอง' }),
          h('datalist', { id }, (opts.options || []).map(o => h('option', { value: o }))));
      }
      case 'textarea': return h('textarea', { rows: opts.rows || 3, oninput: e => onChange(e.target.value), placeholder: opts.placeholder || '' }, value || '');
      case 'date': return h('input', { type: 'date', value: value || '', onchange: e => onChange(e.target.value) });
      case 'month': return h('input', { type: 'month', value: value || '', onchange: e => onChange(e.target.value), placeholder: 'YYYY-MM' });
      case 'taxid': return h('input', { type: 'text', inputmode: 'numeric', maxlength: 13, value: value || '', placeholder: '13 หลัก', oninput: e => { e.target.value = e.target.value.replace(/\D/g, ''); onChange(e.target.value); } });
      case 'email': return h('input', { type: 'email', value: value || '', oninput: e => onChange(e.target.value) });
      default: return h('input', { type: 'text', value: value || '', oninput: e => onChange(e.target.value), placeholder: opts.placeholder || '' });
    }
  }
  function cellInput(sheet, ref, type, opts = {}) {
    return inputFor(type, getVal(sheet, ref), v => { setVal(sheet, ref, v); if (opts.after) opts.after(v); }, opts);
  }
  function cellTag(sheet, ref) { return h('span', { class: 'cellref', title: 'ตำแหน่งเซลล์ในไฟล์ ก.ล.ต.' }, `${sheet}!${ref}`); }

  // ---------------- tab: start ----------------
  function renderStart() {
    const fileInput = h('input', { type: 'file', accept: '.xlsx', id: 'tplFile', onchange: e => e.target.files[0] && loadTemplateFile(e.target.files[0]) });
    return h('section', { class: 'page' },
      h('h2', {}, 'เริ่มต้นใช้งาน'),
      h('div', { class: 'card privacy' },
        h('h3', {}, '🔒 ข้อมูลไม่ออกจากเครื่องของคุณ'),
        h('ul', {},
          h('li', {}, 'หน้าเว็บนี้ทำงานในเบราว์เซอร์ทั้งหมด ไม่มี server และไม่มีการส่งข้อมูลไปที่ใด (บังคับด้วย Content-Security-Policy: connect-src \'none\')'),
          h('li', {}, 'ไม่มีการบันทึกข้อมูลในเบราว์เซอร์ (ไม่ใช้ localStorage / IndexedDB / cookie) — ปิดแท็บแล้วข้อมูลหายทันที'),
          h('li', {}, 'เก็บงานด้วยปุ่ม "บันทึกงาน (JSON)" และ "Export xlsx" — ไฟล์ทั้งสองมีข้อมูลการตรวจสอบ ควรเก็บตามนโยบายข้อมูลลับของบริษัท'))),
      h('div', { class: 'card' },
        h('h3', {}, 'ขั้นที่ 1 — โหลดไฟล์แบบรายงานของ ก.ล.ต. (.xlsx)'),
        h('p', { class: 'muted' }, 'ใช้ไฟล์ "แบบรายงานผลการตรวจสอบด้าน IT" รอบปี 2569 — จะเป็นไฟล์เปล่า หรือไฟล์ที่เคย Export จากหน้านี้ก็ได้ (ค่าที่กรอกไว้ในไฟล์จะถูกดึงขึ้นมาให้)'),
        fileInput,
        tpl ? h('p', { class: 'ok-text' }, `✓ โหลดแล้ว: ${tplFileName} — ${tpl.controls.length} controls ใน ${tpl.ctlSheets.length} sheets`) : null),
      h('div', { class: 'card' + (tpl ? '' : ' disabled') },
        h('h3', {}, 'ขั้นที่ 2 (ถ้ามี) — โหลดงานที่บันทึกไว้ (JSON)'),
        h('p', { class: 'muted' }, 'ไฟล์ JSON เก็บทั้งค่าที่กรอกและบันทึกส่วนตัว (working note, evidence ref, สถานะ finding) ซึ่งไม่ได้อยู่ในไฟล์ xlsx'),
        h('button', { class: 'btn', disabled: !tpl, onclick: () => $('#jsonFile').click() }, 'เลือกไฟล์ JSON…')),
      h('div', { class: 'card' },
        h('h3', {}, 'ลำดับการใช้งาน'),
        h('ol', {},
          h('li', {}, 'โหลดแบบฟอร์ม → กรอกตามแท็บ 1–5'),
          h('li', {}, 'กด "บันทึกงาน (JSON)" เป็นระยะ เพื่อทำต่อภายหลัง'),
          h('li', {}, 'แท็บ "ตรวจความครบถ้วน" ช่วยหาจุดที่ยังขาดก่อน Export'),
          h('li', {}, 'กด "Export xlsx" ได้ไฟล์แบบฟอร์ม ก.ล.ต. ที่กรอกแล้ว (sheet สรุปผลและกราฟคำนวณใหม่เมื่อเปิดด้วย Excel)'))));
  }

  // ---------------- tab: basic ----------------
  function renderBasic() {
    const S = T.basic.sheet;
    const issuesBox = h('div', { class: 'issues', id: 'basicIssues' });
    const page = h('section', { class: 'page' }, h('h2', {}, '1. ข้อมูลพื้นฐานเกี่ยวกับการตรวจสอบ'), issuesBox);
    for (const g of tpl.basicGroups) {
      page.append(h('div', { class: 'card' },
        h('h3', {}, g.title),
        g.note ? h('p', { class: 'muted' }, g.note) : null,
        h('div', { class: 'grid' }, g.fields.map(f =>
          h('label', { class: 'field' + (f.wide ? ' wide' : '') },
            h('span', { class: 'lbl' }, f.label, ' ', cellTag(S, f.ref)),
            cellInput(S, f.ref, f.type, { options: f.options, after: refreshBasicIssues }),
            f.hint ? h('small', { class: 'hint' }, f.hint) : null))),
        g.foot ? h('p', { class: 'hint pre' }, g.foot) : null));
    }
    page.append(renderAuditors());
    setTimeout(refreshBasicIssues);
    return page;
  }
  function refreshBasicIssues() {
    const box = $('#basicIssues');
    if (!box) return;
    const list = validate().filter(i => i.tab === 'basic');
    box.replaceChildren(...(list.length ? [h('strong', {}, 'สิ่งที่ต้องตรวจในแท็บนี้'), h('ul', {}, list.map(i => h('li', { class: i.level }, i.msg)))] : [h('span', { class: 'ok-text' }, '✓ ข้อมูลในแท็บนี้ผ่านเกณฑ์ตรวจเบื้องต้น')]));
  }
  function visibleRowCount(sheet, rows, cols, min = 3) {
    let last = -1;
    rows.forEach((r, i) => { if (cols.some(c => getVal(sheet, c + r))) last = i; });
    return Math.min(rows.length, Math.max(min, last + 2));
  }
  function renderAuditors() {
    const S = T.basic.sheet, A = T.basic.auditors;
    const colsLetters = tpl.auditorCols.map(c => c.col);
    let shown = visibleRowCount(S, A.rows, colsLetters);
    const tbody = h('tbody');
    const addBtn = h('button', { class: 'btn small', onclick: () => { if (shown < A.rows.length) { shown++; draw(); } } }, '+ เพิ่มแถว');
    function draw() {
      tbody.replaceChildren(...A.rows.slice(0, shown).map((r, i) =>
        h('tr', {}, h('td', { class: 'num' }, i + 1),
          tpl.auditorCols.map(c => h('td', {}, cellInput(S, c.col + r, c.type, { options: c.options, after: refreshBasicIssues }))))));
      addBtn.disabled = shown >= A.rows.length;
    }
    draw();
    return h('div', { class: 'card' },
      h('h3', {}, `รายชื่อผู้ตรวจสอบ (สูงสุด ${A.rows.length} คน)`),
      h('div', { class: 'tablewrap' },
        h('table', { class: 'grid-table' },
          h('thead', {}, h('tr', {}, h('th', {}, '#'), tpl.auditorCols.map(c => h('th', {}, c.label)))),
          h('tbody', { class: 'examples' }, tpl.auditorExamples.map(ex => h('tr', {}, h('td', { class: 'num' }, 'ตัวอย่าง'), ex.map(v => h('td', {}, v))))),
          tbody)),
      addBtn);
  }

  // ---------------- tab: systems ----------------
  function renderSystems() {
    const S = T.systems.sheet;
    return h('section', { class: 'page' },
      h('h2', {}, '2. ระบบ IT ที่ใช้ในการประกอบธุรกิจ'),
      tpl.sysNotes.length ? h('div', { class: 'card muted pre' }, tpl.sysNotes.join('\n')) : null,
      tpl.sysGroups.map(g => h('div', { class: 'card' },
        h('h3', {}, g.title, g.limit ? h('span', { class: 'muted small' }, ' ' + g.limit) : null),
        g.desc ? h('p', { class: 'muted' }, g.desc) : null,
        g.example ? h('p', { class: 'hint' }, 'ตัวอย่าง: ' + g.example) : null,
        h('table', { class: 'grid-table' },
          h('thead', {}, h('tr', {}, h('th', {}, '#'), h('th', {}, 'Vendor / Product'), g.other ? h('th', {}, 'อื่น ๆ (หากไม่มีใน dropdown)') : null)),
          h('tbody', {}, g.rows.map((r, i) => h('tr', {},
            h('td', { class: 'num' }, i + 1),
            h('td', {}, g.options.length ? cellInput(S, 'B' + r, 'list', { options: g.options }) : cellInput(S, 'B' + r, 'text', { placeholder: 'พิมพ์ชื่อผู้ให้บริการ' })),
            g.other ? h('td', {}, cellInput(S, 'C' + r, 'text')) : null)))))));
  }

  // ---------------- tab: scope ----------------
  function renderScope() {
    const S = T.scope.sheet, SC = T.scope;
    const letters = tpl.scopeCols.map(c => c.col);
    let shown = visibleRowCount(S, SC.rows, letters);
    const tbody = h('tbody');
    const addBtn = h('button', { class: 'btn small', onclick: () => { if (shown < SC.rows.length) { shown++; draw(); } } }, '+ เพิ่มระบบ');
    function draw() {
      tbody.replaceChildren(...SC.rows.slice(0, shown).map((r, i) => h('tr', {},
        h('td', { class: 'num' }, i + 1),
        tpl.scopeCols.map(c => h('td', {}, cellInput(S, c.col + r, c.type, { options: c.options, rows: 2 }))))));
      addBtn.disabled = shown >= SC.rows.length;
    }
    draw();
    return h('section', { class: 'page' },
      h('h2', {}, '3. ระบบ IT ที่อยู่ในขอบเขตของการตรวจสอบ'),
      tpl.scopeNote ? h('div', { class: 'card muted pre' }, tpl.scopeNote) : null,
      h('div', { class: 'card' },
        h('div', { class: 'tablewrap' },
          h('table', { class: 'grid-table scope' },
            h('thead', {}, h('tr', {}, h('th', {}, '#'), tpl.scopeCols.map(c => h('th', {}, c.label)))),
            h('tbody', { class: 'examples' }, tpl.scopeExamples.map(ex => h('tr', {}, h('td', { class: 'num' }, 'ตัวอย่าง'), ex.map(v => h('td', {}, v))))),
            tbody)),
        addBtn));
  }

  // ---------------- tab: controls ----------------
  function renderControls() {
    if (!ctlSheet) ctlSheet = tpl.ctlSheets[0].name;
    const sh = tpl.ctlSheets.find(s => s.name === ctlSheet);
    const page = h('section', { class: 'page controls-page' });
    const nav = h('nav', { class: 'subnav', id: 'ctlNav' });
    drawCtlNav(nav);
    const col = rlaCol();
    const rlaLabel = getVal(T.basic.sheet, 'C26');
    const filterBar = h('div', { class: 'filterbar' },
      h('label', {}, 'ระดับ RLA ',
        selectEl(['auto', 'H', 'M', 'L', 'S', 'all'], filters.rla, v => { filters.rla = v || 'auto'; render(); }, null)),
      h('span', { class: 'muted small' }, filters.rla === 'auto' ? (rlaLabel ? `(ตามแท็บ 1: ${rlaLabel})` : '(ยังไม่เลือก RLA ในแท็บ 1 → แสดงทั้งหมด)') : ''),
      h('label', {}, 'สถานะ ',
        selectEl(['all', 'todo', 'gap', 'done'], filters.status, v => { filters.status = v || 'all'; render(); }, null)),
      h('input', { type: 'search', placeholder: 'ค้นหา ID / ข้อความ', value: filters.q, oninput: e => { filters.q = e.target.value; drawList(); } }),
      h('button', { class: 'btn small', title: 'ใส่ N/A ให้ข้อที่ไม่อยู่ในระดับ RLA ของบริษัท (เฉพาะข้อที่ยังว่าง ทุก sheet)', onclick: fillNaForNonApplicable }, 'ใส่ N/A ข้อนอกระดับ RLA'));
    // ป้ายตัวเลือกที่อ่านง่าย
    const lbl = { auto: 'ตามแท็บ 1', H: 'สูง (H)', M: 'ปานกลาง (M)', L: 'ต่ำ (L)', S: 'ขนาดเล็ก (S)', all: 'แสดงทั้งหมด', todo: 'ยังไม่ประเมิน', gap: 'No / Partial / M1–M2', done: 'ประเมินแล้ว' };
    filterBar.querySelectorAll('option').forEach(o => { if (lbl[o.value]) o.textContent = lbl[o.value]; });

    const list = h('div', { class: 'ctl-list', id: 'ctlList' });
    function drawList() {
      const q = filters.q.trim().toLowerCase();
      const out = [h('h2', {}, sh.short + ' — ' + sh.title)];
      let lastHeader = null, shownCount = 0;
      for (const it of sh.items) {
        if (it.kind === 'h') { lastHeader = it; continue; }
        const c = it.ctl, v = ctlValue(c);
        if (!isApplicable(c, col)) continue;
        if (filters.status === 'todo' && (v || c.fixed)) continue;
        if (filters.status === 'gap' && !isGap(v)) continue;
        if (filters.status === 'done' && !v) continue;
        if (q && !(c.id.toLowerCase().includes(q) || c.text.toLowerCase().includes(q) || (state.notes[c.id] && JSON.stringify(state.notes[c.id]).toLowerCase().includes(q)))) continue;
        if (lastHeader) { out.push(h('h' + (lastHeader.level + 1), { class: 'sec-h' }, lastHeader.text)); lastHeader = null; }
        out.push(controlCard(c));
        shownCount++;
      }
      if (!shownCount) out.push(h('p', { class: 'muted' }, 'ไม่มี control ที่ตรงกับตัวกรอง'));
      list.replaceChildren(...out);
    }
    drawList();
    page.append(nav, h('div', { class: 'ctl-main' }, filterBar, list));
    return page;
  }
  function drawCtlNav(nav = $('#ctlNav')) {
    if (!nav) return;
    nav.replaceChildren(...tpl.ctlSheets.map(sh => {
      const p = sheetProgress(sh);
      const gaps = sh.controls.filter(c => isGap(ctlValue(c))).length;
      return h('button', { class: (sh.name === ctlSheet ? 'active ' : '') + (p.total && p.done === p.total ? 'complete' : ''), onclick: () => { ctlSheet = sh.name; render(); } },
        h('span', { class: 'nm' }, sh.short),
        h('span', { class: 'pg' }, `${p.done}/${p.total}`),
        gaps ? h('span', { class: 'gapcount', title: 'จำนวนข้อ No/Partial/M1/M2' }, gaps) : null);
    }));
  }
  function controlCard(c) {
    const card = h('article', { class: 'ctl', id: 'ctl-' + c.id.replace(/\W+/g, '_') });
    function paint() {
      const v = ctlValue(c);
      card.className = 'ctl ' + resultClass(v) + (isApplicable(c) ? '' : ' notappl');
      const opts = c.type === 'Maturity' ? ['M1', 'M2', 'M3', 'M4', 'M5', 'N/A'] : ['Yes', 'Partial', 'No', 'N/A'];
      const note = state.notes[c.id] || {};
      const linked = findingsForControl(c.id);
      const badges = ['H', 'M', 'L', 'S'].map(k => h('span', { class: 'badge' + (c.appl[k] ? ' on' : ''), title: { H: 'RLA สูง', M: 'RLA ปานกลาง', L: 'RLA ต่ำ', S: 'ผู้ประกอบธุรกิจขนาดเล็ก (Cyber Hygiene)' }[k] }, k));
      card.replaceChildren(...[
        h('header', {},
          h('span', { class: 'cid' }, c.id),
          h('span', { class: 'badges' }, badges, c.adv ? h('span', { class: 'badge adv on', title: 'Advanced control (สำหรับผู้ประกอบธุรกิจที่มีความเสี่ยงสูง)' }, 'ADV') : null),
          h('span', { class: 'refs' }, `สธ. ${c.sor || '-'}${c.np ? ' · นป. #' + c.np : ''}`, ' ', cellTag(c.sheet, c.cell))),
        h('p', { class: 'ctext pre' }, c.text),
        c.fixed
          ? h('p', { class: 'fixed' }, `ผลการประเมิน: ${c.fixed} — แบบฟอร์มกำหนดไว้ (ไม่ต้องกรอก)`)
          : h('div', { class: 'results' },
            h('span', { class: 'lbl' }, c.type === 'Maturity' ? 'Maturity:' : 'ผลการประเมิน:'),
            opts.map(o => h('button', { class: 'res ' + resultClass(o) + (v === o ? ' sel' : ''), onclick: () => { setVal(c.sheet, c.cell, v === o ? '' : o); paint(); drawCtlNav(); } }, o)),
            v ? h('button', { class: 'linkbtn', onclick: () => { setVal(c.sheet, c.cell, ''); paint(); drawCtlNav(); } }, 'ล้าง') : null),
        (c.type === 'Maturity' && /^M[1-5]$/.test(v)) ? h('p', { class: 'matdesc' }, h('strong', {}, v + ': '), c.mat[+v[1] - 1]) : null,
        c.type === 'Maturity' ? h('details', {}, h('summary', {}, 'ดูคำอธิบาย Maturity ทั้ง 5 ระดับ'),
          h('ol', { class: 'matlist' }, c.mat.map((m, i) => h('li', { class: v === 'M' + (i + 1) ? 'cur' : '' }, h('strong', {}, `M${i + 1}: `), m)))) : null,
        h('details', { class: 'notes', open: !!(note.note || note.evidence) },
          h('summary', {}, 'บันทึกของผู้ตรวจ ', h('span', { class: 'muted small' }, '(อยู่ใน JSON เท่านั้น ไม่ถูก Export ลงไฟล์ ก.ล.ต.)')),
          h('label', { class: 'field wide' }, h('span', { class: 'lbl' }, 'Working note'),
            h('textarea', { rows: 2, oninput: e => setNote(c.id, 'note', e.target.value) }, note.note || '')),
          h('label', { class: 'field wide' }, h('span', { class: 'lbl' }, 'Evidence ref (ค้นด้วย Ctrl+F ได้)'),
            h('input', { type: 'text', value: note.evidence || '', placeholder: 'เช่น WP-07 IT · ไฟล์ Access Review Q3.xlsx · sheet "Leaver"', oninput: e => setNote(c.id, 'evidence', e.target.value) }))),
        h('footer', {},
          linked.length ? h('span', { class: 'linked' }, 'Finding: ', linked.map(([f, i]) => h('button', { class: 'linkbtn', onclick: () => gotoFinding(f.uid) }, `#${i + 1}${f.F ? ' (' + f.F + ')' : ''}`))) : null,
          !c.fixed ? h('button', { class: 'btn small', onclick: () => createFindingFrom(c) }, '+ สร้าง Finding จากข้อนี้') : null),
      ].filter(Boolean));
    }
    paint();
    return card;
  }
  function setNote(id, key, v) {
    if (!state.notes[id]) state.notes[id] = {};
    if (v) state.notes[id][key] = v; else delete state.notes[id][key];
    if (!Object.keys(state.notes[id]).length) delete state.notes[id];
    markDirty();
  }
  function fillNaForNonApplicable() {
    const col = T.rlaToCol[getVal(T.basic.sheet, 'C26')];
    if (!col) { toast('กรุณาเลือกระดับ RLA ในแท็บ 1 ก่อน', 'error'); return; }
    const targets = tpl.controls.filter(c => !c.fixed && !c.appl[col] && !getVal(c.sheet, c.cell));
    if (!targets.length) { toast('ไม่มีข้อที่ต้องใส่ N/A'); return; }
    if (!confirm(`ใส่ "N/A" ให้ ${targets.length} ข้อที่ไม่อยู่ในระดับ RLA "${getVal(T.basic.sheet, 'C26')}" และยังไม่ได้ประเมิน?\n(ข้อที่ประเมินไว้แล้วจะไม่ถูกเปลี่ยน)`)) return;
    targets.forEach(c => setVal(c.sheet, c.cell, 'N/A'));
    toast(`ใส่ N/A แล้ว ${targets.length} ข้อ`);
    render();
  }

  // ---------------- tab: findings ----------------
  function createFindingFrom(c) {
    if (state.findings.length >= T.findings.rows.length) { toast(`แบบฟอร์มรองรับสูงสุด ${T.findings.rows.length} ข้อ`, 'error'); return; }
    const f = newFinding();
    f.C = c.id; f.B = domainFor(c.id); f.D = topicFor(c.id);
    state.findings.push(f);
    markDirty();
    gotoFinding(f.uid);
  }
  function gotoFinding(uid) {
    tab = 'findings'; updateHeader(); render();
    const el = document.getElementById('fd-' + uid);
    if (el) { el.scrollIntoView({ behavior: 'smooth', block: 'start' }); el.classList.add('flash'); setTimeout(() => el.classList.remove('flash'), 1500); }
  }
  function findingCounts() {
    const c = { H: 0, M: 0, L: 0, open: 0, closed: 0, overdue: 0 };
    for (const f of state.findings) {
      if (c[f.F] !== undefined) c[f.F]++;
      if (f.status === 'Closed') c.closed++; else c.open++;
      if (isOverdue(f)) c.overdue++;
    }
    return c;
  }
  function renderFindings() {
    const listEl = h('div', { class: 'find-list' });
    const summary = h('div', { class: 'card summary' });
    function drawSummary() {
      const c = findingCounts();
      summary.replaceChildren(
        h('div', { class: 'stats' },
          stat('ทั้งหมด', state.findings.length), stat('High', c.H, 'gap'), stat('Medium', c.M, 'part'), stat('Low', c.L, 'ok'),
          stat('Open', c.open), stat('Closed', c.closed, 'na'), stat('เลยกำหนด', c.overdue, c.overdue ? 'gap' : '')),
        h('p', { class: 'hint' },
          `หมายเหตุ: ช่อง ${T.findings.sheet}!E5 และ E6 (จำนวนข้อระดับปานกลาง/ต่ำ) เป็นช่องที่แบบฟอร์มให้กรอกเอง — ตามที่ตกลงไว้ แอปจะไม่เขียนให้ หลัง Export ให้กรอกเอง: ปานกลาง = ${c.M}, ต่ำ = ${c.L} (E4 ระดับสูง คำนวณด้วยสูตรของแบบฟอร์ม)`));
    }
    function draw() {
      drawSummary();
      if (!state.findings.length) { listEl.replaceChildren(h('p', { class: 'muted' }, 'ยังไม่มี Finding — กด "+ เพิ่ม Finding" หรือสร้างจากหน้า Controls')); return; }
      listEl.replaceChildren(...state.findings.map((f, i) => findingCard(f, i, draw, drawSummary)));
    }
    draw();
    return h('section', { class: 'page' },
      h('h2', {}, '5. สรุปประเด็นข้อตรวจพบ/ข้อบกพร่อง'),
      tpl.findNote ? h('p', { class: 'muted' }, tpl.findNote) : null,
      summary,
      h('details', { class: 'card' }, h('summary', {}, 'ดูตัวอย่างจากแบบฟอร์ม'),
        tpl.findExamples.map(ex => h('div', { class: 'example' }, tpl.findCols.map(c => ex[c.key] ? h('p', {}, h('strong', {}, clean(c.label).split('\n')[0] + ': '), String(ex[c.key])) : null)))),
      listEl,
      h('button', { class: 'btn', onclick: () => {
        if (state.findings.length >= T.findings.rows.length) { toast(`แบบฟอร์มรองรับสูงสุด ${T.findings.rows.length} ข้อ`, 'error'); return; }
        const f = newFinding(); state.findings.push(f); markDirty(); draw();
        document.getElementById('fd-' + f.uid).scrollIntoView({ behavior: 'smooth' });
      } }, '+ เพิ่ม Finding'));
  }
  function stat(label, n, cls = '') { return h('div', { class: 'stat ' + cls }, h('span', { class: 'n' }, n), h('span', { class: 'l' }, label)); }
  function findingCard(f, i, redraw, redrawSummary) {
    const C = tpl.findColByKey;
    const idWarn = h('small', { class: 'warn-text' });
    function checkIds() {
      const bad = parseIds(f.C).filter(id => !tpl.controlsById[id]);
      idWarn.textContent = bad.length ? `ไม่พบ control ID: ${bad.join(', ')}` : '';
    }
    const set = (k, v) => { f[k] = v; markDirty(); };
    const row = T.findings.rows[i];
    const card = h('article', { class: 'finding' + (isOverdue(f) ? ' overdue' : ''), id: 'fd-' + f.uid },
      h('header', {},
        h('span', { class: 'cid' }, `Finding #${i + 1}`),
        h('span', { class: 'muted small' }, `→ ${T.findings.sheet} แถว ${row}`),
        h('span', { class: 'spacer' }),
        h('button', { class: 'linkbtn', disabled: i === 0, onclick: () => { [state.findings[i - 1], state.findings[i]] = [state.findings[i], state.findings[i - 1]]; markDirty(); redraw(); } }, '↑'),
        h('button', { class: 'linkbtn', disabled: i === state.findings.length - 1, onclick: () => { [state.findings[i + 1], state.findings[i]] = [state.findings[i], state.findings[i + 1]]; markDirty(); redraw(); } }, '↓'),
        h('button', { class: 'linkbtn danger', onclick: () => { if (confirm(`ลบ Finding #${i + 1}?`)) { state.findings.splice(i, 1); markDirty(); redraw(); } } }, 'ลบ')),
      h('div', { class: 'grid' },
        h('label', { class: 'field' }, h('span', { class: 'lbl' }, label(C.C)),
          inputFor('text', f.C, v => { set('C', v); checkIds(); }, { placeholder: 'เช่น D2.11 #13, D2.11 #14' }),
          idWarn,
          h('button', { class: 'linkbtn small', onclick: () => {
            const first = parseIds(f.C)[0];
            if (!first || !tpl.controlsById[first]) { toast('กรอก control ID ก่อน', 'error'); return; }
            f.B = domainFor(first); f.D = topicFor(first) || f.D; markDirty(); redraw();
          } }, 'เติมหมวด/หัวข้อจาก control ID')),
        h('label', { class: 'field' }, h('span', { class: 'lbl' }, label(C.B)), selectEl(C.B.options, f.B, v => set('B', v))),
        h('label', { class: 'field wide' }, h('span', { class: 'lbl' }, label(C.D)), selectEl(C.D.options, f.D, v => set('D', v))),
        h('label', { class: 'field wide' }, h('span', { class: 'lbl' }, label(C.E)), inputFor('textarea', f.E, v => set('E', v), { rows: 3 })),
        h('label', { class: 'field' }, h('span', { class: 'lbl' }, label(C.F)), selectEl(C.F.options, f.F, v => { set('F', v); redrawSummary(); })),
        h('label', { class: 'field' }, h('span', { class: 'lbl' }, label(C.H)), inputFor('month', f.H, v => { set('H', v); card.classList.toggle('overdue', isOverdue(f)); redrawSummary(); })),
        h('label', { class: 'field wide' }, h('span', { class: 'lbl' }, label(C.G)), inputFor('textarea', f.G, v => set('G', v), { rows: 3 })),
        h('label', { class: 'field wide' }, h('span', { class: 'lbl' }, label(C.I)), inputFor('text', f.I, v => set('I', v)))),
      h('div', { class: 'private' },
        h('span', { class: 'muted small' }, 'ติดตามผล (อยู่ใน JSON เท่านั้น):'),
        h('label', {}, 'สถานะ ', selectEl(['Open', 'Closed'], f.status, v => { set('status', v || 'Open'); card.classList.toggle('overdue', isOverdue(f)); redrawSummary(); }, null)),
        h('label', {}, 'วันที่ปิดจริง ', inputFor('date', f.closed, v => set('closed', v))),
        isOverdue(f) ? h('span', { class: 'badge gap on' }, 'เลยกำหนด') : null));
    checkIds();
    return card;
  }
  function label(col) { return clean(col.label).split('\n')[0] || col.col; }

  // ---------------- validation ----------------
  function validate() {
    const out = [];
    const add = (level, tabName, msg) => out.push({ level, tab: tabName, msg });
    const B = T.basic.sheet, g = r => getVal(B, r);
    const DR = T.dateRules;
    if (!g('C5')) add('error', 'basic', 'ยังไม่ได้กรอกชื่อบริษัท (C5)');
    if (!/^\d{12,13}$/.test(g('C7'))) add('error', 'basic', 'เลขทะเบียนนิติบุคคล (C7) ต้องเป็นตัวเลข 12–13 หลัก');
    if (!g('C9') || !g('F9') || !g('H9')) add('warn', 'basic', 'ข้อมูลผู้ประสานงาน (ชื่อ/เบอร์/อีเมล) ยังไม่ครบ');
    const auditorTypes = (tpl.fields[B].C16 || {}).options || [];
    if (!g('C16')) add('error', 'basic', 'ยังไม่ได้เลือกผู้ดำเนินการตรวจสอบ (C16)');
    else if (g('C16') !== auditorTypes[0] && !g('C18')) add('error', 'basic', 'กรณีใช้ผู้เชี่ยวชาญภายนอก/Co-Sourcing ต้องระบุชื่อบริษัทผู้ตรวจสอบ (C18)');
    if (!g('C23') || !g('E23')) add('error', 'basic', 'ยังไม่ได้กรอกวันที่ดำเนินการตรวจสอบให้ครบ (C23, E23)');
    if (g('C23') && g('C23') < DR.min) add('error', 'basic', `วันเริ่มตรวจ (C23) ต้องไม่ก่อน ${DR.min}`);
    if (g('E23') && (g('E23') < DR.min || g('E23') > DR.auditEndMax)) add('error', 'basic', `วันสุดท้ายของการตรวจ (E23) ต้องอยู่ระหว่าง ${DR.min} ถึง ${DR.auditEndMax}`);
    if (g('C23') && g('E23') && g('E23') < g('C23')) add('error', 'basic', 'วันสุดท้ายของการตรวจ (E23) อยู่ก่อนวันเริ่มตรวจ (C23)');
    for (const r of ['I16', 'I18']) if (g(r) && g(r) < DR.min) add('error', 'basic', `วันที่เสนอคณะกรรมการ (${r}) ต้องไม่ก่อน ${DR.min}`);
    if (!g('I16') && !g('I18')) add('warn', 'basic', 'ยังไม่ระบุวันที่เสนอผลต่อคณะกรรมการ (I16) หรือ Audit Committee (I18)');
    const boardMax = [g('I16'), g('I18')].filter(Boolean).sort().pop();
    if (boardMax && g('E23') && g('E23') > boardMax) add('warn', 'basic', 'วันที่เสนอต่อคณะกรรมการ/AC อยู่ก่อนวันสุดท้ายของการตรวจ (แบบฟอร์มจะขึ้นข้อความเตือนที่ H21)');
    if (!g('C26')) add('error', 'basic', 'ยังไม่ได้เลือกระดับความเสี่ยง RLA (C26)');
    const A = T.basic.auditors, ac = tpl.auditorCols.map(c => c.col);
    const filled = A.rows.filter(r => ac.some(c => g(c + r)));
    if (!filled.length) add('error', 'basic', 'ยังไม่มีรายชื่อผู้ตรวจสอบ');
    for (const r of filled) {
      if (!g('B' + r)) add('warn', 'basic', `รายชื่อผู้ตรวจสอบแถว ${r} ยังไม่มีชื่อ`);
      if (!g('D' + r)) add('warn', 'basic', `รายชื่อผู้ตรวจสอบแถว ${r} ยังไม่ระบุตำแหน่ง`);
    }
    const leaderOpt = ((tpl.auditorCols.find(c => c.col === 'D') || {}).options || [])[0];
    if (filled.length && leaderOpt && !filled.some(r => g('D' + r) === leaderOpt)) add('warn', 'basic', `ยังไม่มีผู้ตรวจในตำแหน่ง "${leaderOpt}"`);

    // scope
    const SC = T.scope;
    const scopeRows = SC.rows.filter(r => SC.cols.some(c => getVal(SC.sheet, c.col + r)));
    if (!scopeRows.length) add('warn', 'scope', 'ยังไม่ได้ระบุระบบที่อยู่ในขอบเขตการตรวจสอบ');
    for (const r of scopeRows) {
      const gv = c => getVal(SC.sheet, c + r);
      if (!gv('B') || !gv('C') || !gv('E') || !gv('F')) add('warn', 'scope', `ระบบแถว ${r}: ข้อมูลยังไม่ครบ (ลักษณะ/ประเภท/ชื่อผลิตภัณฑ์/รูปแบบการใช้งาน)`);
      if (/อื่น/.test(gv('C')) && !gv('D')) add('warn', 'scope', `ระบบแถว ${r}: เลือก "ระบบอื่น ๆ" แต่ยังไม่ระบุประเภทระบบ (คอลัมน์ D)`);
    }

    // controls
    const col = T.rlaToCol[g('C26')] || null;
    let blankOutside = 0;
    for (const sh of tpl.ctlSheets) {
      const p = sheetProgress(sh, col);
      if (p.done < p.total) add('warn', 'controls', `${sh.short}: ยังไม่ประเมิน ${p.total - p.done} จาก ${p.total} ข้อ${col ? ' (ในระดับ RLA)' : ''}`);
      if (col) blankOutside += sh.controls.filter(c => !c.fixed && !c.appl[col] && !getVal(c.sheet, c.cell)).length;
    }
    if (blankOutside) add('info', 'controls', `มี ${blankOutside} ข้อนอกระดับ RLA ที่ยังว่าง — sheet ของแบบฟอร์มจะขึ้น "ยังไม่ได้บันทึกผล" สำหรับข้อเหล่านี้ (ใช้ปุ่ม "ใส่ N/A ข้อนอกระดับ RLA" ได้ถ้าต้องการ)`);

    // consistency
    const referenced = new Set(state.findings.flatMap(f => parseIds(f.C)));
    const gapNoFinding = tpl.controls.filter(c => isGap(ctlValue(c)) && !referenced.has(c.id));
    if (gapNoFinding.length) add('info', 'controls', `ข้อที่ประเมินเป็น No/Partial/M1/M2 แต่ยังไม่มี Finding อ้างถึง (${gapNoFinding.length}): ${gapNoFinding.slice(0, 15).map(c => c.id).join(', ')}${gapNoFinding.length > 15 ? ' …' : ''}`);

    // findings
    state.findings.forEach((f, i) => {
      const n = `Finding #${i + 1}`;
      const ids = parseIds(f.C);
      const miss = [];
      if (!ids.length) miss.push('control ID');
      if (!f.B) miss.push('หมวด');
      if (!f.D) miss.push('หัวข้อ');
      if (!f.E) miss.push('ข้อตรวจพบ');
      if (!f.F) miss.push('ระดับความสำคัญ');
      if (!f.G) miss.push('แผนแก้ไข');
      if (!f.H) miss.push('วันที่แล้วเสร็จ');
      if (miss.length) add('error', 'findings', `${n}: ยังไม่ได้กรอก ${miss.join(', ')}`);
      const bad = ids.filter(id => !tpl.controlsById[id]);
      if (bad.length) add('error', 'findings', `${n}: ไม่พบ control ID ${bad.join(', ')}`);
      const okRated = ids.filter(id => tpl.controlsById[id] && ['Yes', 'M3', 'M4', 'M5'].includes(ctlValue(tpl.controlsById[id])));
      if (okRated.length) add('warn', 'findings', `${n}: อ้างถึง control ที่ประเมินผ่าน (${okRated.map(id => id + '=' + ctlValue(tpl.controlsById[id])).join(', ')}) — ตรวจความสอดคล้อง`);
      if (f.B && ids[0] && tpl.controlsById[ids[0]] && f.B !== domainFor(ids[0])) add('warn', 'findings', `${n}: หมวด "${f.B}" ไม่ตรงกับ control ${ids[0]}`);
    });
    if (state.findings.length) {
      const c = findingCounts();
      add('info', 'findings', `หลัง Export ต้องกรอก ${T.findings.sheet}!E5 = ${c.M} (ปานกลาง) และ E6 = ${c.L} (ต่ำ) เอง`);
    }
    return out;
  }

  // ---------------- tab: check ----------------
  function renderCheck() {
    const issues = validate();
    const tabsName = { basic: '1. ข้อมูลทั่วไป', systems: '2. ระบบ IT', scope: '3. ขอบเขต', controls: '4. Controls', findings: '5. Finding' };
    const levelName = { error: 'ต้องแก้', warn: 'ควรตรวจ', info: 'ข้อสังเกต' };
    const col = T.rlaToCol[getVal(T.basic.sheet, 'C26')] || null;
    const dist = { ok: 0, part: 0, gap: 0, na: 0, blank: 0 };
    const rows = tpl.ctlSheets.map(sh => {
      const d = { ok: 0, part: 0, gap: 0, na: 0, blank: 0 };
      for (const c of sh.controls) {
        if (!isApplicable(c, col)) continue;
        const k = resultClass(ctlValue(c)) || 'blank';
        d[k]++; dist[k]++;
      }
      const p = sheetProgress(sh, col);
      return h('tr', {}, h('td', {}, sh.short), h('td', {}, sh.title), h('td', { class: 'num' }, `${p.done}/${p.total}`),
        h('td', { class: 'num ok' }, d.ok), h('td', { class: 'num part' }, d.part), h('td', { class: 'num gap' }, d.gap), h('td', { class: 'num na' }, d.na), h('td', { class: 'num' }, d.blank));
    });
    return h('section', { class: 'page' },
      h('h2', {}, 'ตรวจความครบถ้วนก่อน Export'),
      h('div', { class: 'card' },
        h('p', {}, `ระดับ RLA: ${getVal(T.basic.sheet, 'C26') || '(ยังไม่เลือก — นับทุกข้อ)'}`),
        h('div', { class: 'tablewrap' }, h('table', { class: 'grid-table' },
          h('thead', {}, h('tr', {}, ['Sheet', 'หมวด', 'ประเมินแล้ว', 'Yes/M3+', 'Partial/M2', 'No/M1', 'N/A', 'ว่าง'].map(x => h('th', {}, x)))),
          h('tbody', {}, rows),
          h('tfoot', {}, h('tr', {}, h('td', {}, 'รวม'), h('td'), h('td'), h('td', { class: 'num ok' }, dist.ok), h('td', { class: 'num part' }, dist.part), h('td', { class: 'num gap' }, dist.gap), h('td', { class: 'num na' }, dist.na), h('td', { class: 'num' }, dist.blank))))),
        h('p', { class: 'hint' }, 'คะแนนและ Compliance % ตามสูตรของ ก.ล.ต. ดูได้ใน sheet "6. Result" หลังเปิดไฟล์ที่ Export ด้วย Excel (รายงานสรุปในหน้านี้จะทำใน Phase 2)')),
      h('div', { class: 'card' },
        h('h3', {}, `รายการตรวจ (${issues.filter(i => i.level === 'error').length} ต้องแก้ · ${issues.filter(i => i.level === 'warn').length} ควรตรวจ · ${issues.filter(i => i.level === 'info').length} ข้อสังเกต)`),
        issues.length ? h('ul', { class: 'issues-list' }, issues.map(i => h('li', { class: i.level },
          h('span', { class: 'lvl' }, levelName[i.level]), ' ',
          h('button', { class: 'linkbtn', onclick: () => go(i.tab) }, tabsName[i.tab]), ' — ', i.msg)))
          : h('p', { class: 'ok-text' }, '✓ ไม่พบประเด็น')),
      h('div', { class: 'card' },
        h('button', { class: 'btn primary', onclick: exportXlsx }, 'Export xlsx'), ' ',
        h('button', { class: 'btn', onclick: saveJson }, 'บันทึกงาน (JSON)'),
        h('p', { class: 'hint' }, 'Export ได้แม้ยังมีประเด็นค้าง (ไฟล์ทำงานระหว่างทาง)')));
  }

  // ---------------- file I/O ----------------
  function stamp() {
    const d = new Date(), p = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
  }
  function download(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = h('a', { href: url, download: name });
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  async function loadTemplateFile(file) {
    if (tpl && (dirty || hasData()) && !confirm('โหลดแบบฟอร์มใหม่จะแทนที่ข้อมูลที่กรอกอยู่ทั้งหมด (ข้อมูลจะถูกอ่านจากไฟล์ใหม่) — ดำเนินการต่อ?')) return;
    try {
      const buf = await file.arrayBuffer();
      const newWb = await XL.open(buf);
      const newTpl = await parseTemplate(newWb);
      const st = await readExistingValues(newWb, newTpl);
      wb = newWb; tpl = newTpl; tplFileName = file.name; state = st; dirty = false; ctlSheet = null;
      if (tpl.version !== T.formYearLabel) toast(`คำเตือน: ไฟล์นี้ระบุ "${tpl.version || 'ไม่ระบุรอบ'}" แต่เครื่องมือนี้ทำไว้สำหรับ "${T.formYearLabel}" — ผังเซลล์อาจไม่ตรง`, 'error');
      const filledCount = Object.values(st.cells).reduce((n, o) => n + Object.keys(o).length, 0);
      toast(`โหลดแบบฟอร์มแล้ว — ${tpl.controls.length} controls${filledCount ? `, พบค่าที่กรอกไว้ ${filledCount} ช่อง, Finding ${st.findings.length} ข้อ` : ''}`);
      go('start');
    } catch (err) {
      console.error(err);
      toast('อ่านไฟล์ไม่สำเร็จ: ' + err.message, 'error');
    }
  }
  function hasData() { return Object.keys(state.cells).length || state.findings.length || Object.keys(state.notes).length; }

  function saveJson() {
    if (!tpl) return;
    const payload = {
      app: T.appId, schemaVersion: SCHEMA_VERSION, formVersion: tpl.version,
      savedAt: new Date().toISOString(),
      note: 'ไฟล์นี้มีข้อมูลการตรวจสอบ — เก็บตามนโยบายข้อมูลลับของบริษัท',
      cells: state.cells, findings: state.findings, notes: state.notes,
    };
    download(new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' }), `ITAudit2569_work_${stamp()}.json`);
    dirty = false; updateHeader();
    toast('บันทึกงาน (JSON) แล้ว');
  }

  async function loadJsonFile(file) {
    try {
      const data = JSON.parse(await file.text());
      if (data.app !== T.appId) throw new Error('ไฟล์นี้ไม่ได้มาจากเครื่องมือนี้ (app id ไม่ตรง)');
      if (hasData() && !confirm('โหลดงานจาก JSON จะแทนที่ข้อมูลปัจจุบันทั้งหมด — ดำเนินการต่อ?')) return;
      if (data.formVersion && tpl.version && data.formVersion !== tpl.version) toast(`คำเตือน: JSON บันทึกจากแบบฟอร์ม "${data.formVersion}" แต่ไฟล์ที่โหลดอยู่คือ "${tpl.version}"`, 'error');
      state = { cells: data.cells || {}, findings: (data.findings || []).map(f => Object.assign(newFinding(), f)), notes: data.notes || {} };
      dirty = false;
      toast(`โหลดงานแล้ว (บันทึกเมื่อ ${data.savedAt ? new Date(data.savedAt).toLocaleString('th-TH') : '-'})`);
      render(); updateHeader();
    } catch (err) {
      toast('โหลด JSON ไม่สำเร็จ: ' + err.message, 'error');
    }
  }

  function buildUpdates() {
    const ups = {};
    const put = (sheet, ref, v) => { if (!ups[sheet]) ups[sheet] = new Map(); ups[sheet].set(ref, v); };
    for (const [sheet, defs] of Object.entries(tpl.fields)) {
      for (const [ref, def] of Object.entries(defs)) {
        const raw = getVal(sheet, ref);
        let v = raw === '' ? null : raw;
        if (v !== null && def.type === 'date') v = XL.isoToSerial(v);
        else if (v !== null && def.type === 'taxid' && /^\d{12,13}$/.test(v)) v = Number(v);
        put(sheet, ref, v);
      }
    }
    for (const c of tpl.controls) if (!c.fixed) put(c.sheet, c.cell, getVal(c.sheet, c.cell) || null);
    const F = T.findings;
    F.rows.forEach((r, i) => {
      const f = state.findings[i];
      for (const c of F.cols) {
        let v = f ? (f[c.key] || null) : null;
        if (v !== null && c.type === 'month') v = XL.isoToSerial(v);
        if (v !== null && c.type === 'ids') v = parseIds(v).join(', ');
        put(F.sheet, c.col + r, v);
      }
    });
    return ups;
  }

  async function exportXlsx() {
    if (!wb) return;
    if (state.findings.length > T.findings.rows.length) { toast(`Finding เกิน ${T.findings.rows.length} ข้อ`, 'error'); return; }
    try {
      const { blob, skipped } = await XL.build(wb, buildUpdates());
      download(blob, `ITAudit2569_filled_${stamp()}.xlsx`);
      if (skipped.length) toast('ข้ามเซลล์ที่มีสูตร: ' + skipped.join(', '), 'error');
      toast('Export xlsx แล้ว — เปิดด้วย Excel เพื่อให้ sheet สรุปผลคำนวณใหม่ (บันทึกโน้ต/สถานะ finding อยู่ใน JSON เท่านั้น)');
    } catch (err) {
      console.error(err);
      toast('Export ไม่สำเร็จ: ' + err.message, 'error');
    }
  }

  // ---------------- init ----------------
  function init() {
    document.querySelectorAll('#tabs button').forEach(b => b.addEventListener('click', () => go(b.dataset.tab)));
    $('#btnLoadJson').addEventListener('click', () => $('#jsonFile').click());
    $('#btnSaveJson').addEventListener('click', saveJson);
    $('#btnExport').addEventListener('click', exportXlsx);
    $('#jsonFile').addEventListener('change', e => { if (e.target.files[0]) loadJsonFile(e.target.files[0]); e.target.value = ''; });
    window.addEventListener('beforeunload', e => { if (dirty) { e.preventDefault(); e.returnValue = ''; } });
    updateHeader();
    render();
  }

  // เปิดให้ชุดทดสอบเรียกได้ (ไม่มีผลต่อการทำงานปกติ)
  window.__app = { get state() { return state; }, get tpl() { return tpl; }, validate, buildUpdates };
  document.addEventListener('DOMContentLoaded', init);
})();
