/*
 * reports.js — รายงาน Phase 2 (พิมพ์เป็น PDF ผ่าน Print ของเบราว์เซอร์)
 *   summary  = 1-page Summary (A4 แนวตั้ง)
 *   detail   = Detailed Summary (A4 แนวนอน)
 *   findings = Finding คงค้าง (A4 แนวนอน)
 * ตัวเลขทั้งหมดมาจาก Scoring.compute() ซึ่งใช้ logic เดียวกับ sheet "6. Result"
 */
'use strict';

const Reports = (() => {
  const KINDS = {
    summary: { title: '1-page Summary', orientation: 'portrait', file: 'Summary' },
    detail: { title: 'Detailed Summary', orientation: 'landscape', file: 'Detailed' },
    findings: { title: 'Finding คงค้าง', orientation: 'landscape', file: 'OpenFindings' },
  };
  const DAY = 86400000;

  // ---------- format helpers ----------
  const num = (x, d = 2) => (x === null || x === undefined || x === '') ? '–' : (x === 'N/A' ? 'N/A' : Number(x).toFixed(d));
  const pct = x => (x === null || x === undefined || x === '' || x === 'N/A') ? '–' : (x * 100).toFixed(1) + '%';
  const dmy = iso => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); return m ? `${+m[3]}/${+m[2]}/${m[1]}` : '–'; };
  const my = ym => { const m = /^(\d{4})-(\d{2})$/.exec(ym || ''); return m ? `${m[2]}/${m[1]}` : '–'; };
  const trunc = (s, n) => { s = String(s || '').replace(/\s+/g, ' ').trim(); return s.length > n ? s.slice(0, n - 1) + '…' : s; };
  const ratingOrder = { H: 0, M: 1, L: 2 };
  const nb = id => String(id).replace(/ #/g, '\u00a0#'); // ไม่ให้ตัดบรรทัดกลาง control ID

  function aging(f, asOfISO) {
    if (!/^\d{4}-\d{2}$/.test(f.H || '')) return { code: 'none', label: 'ไม่ระบุกำหนด', days: null, due: null };
    const [y, m] = f.H.split('-').map(Number);
    const due = Date.UTC(y, m, 0);                // วันสุดท้ายของเดือนเป้าหมาย
    const asOf = Date.UTC(...asOfISO.split('-').map((v, i) => i === 1 ? v - 1 : +v));
    const days = Math.round((due - asOf) / DAY);
    if (days < 0) return { code: 'overdue', label: `เลยกำหนด ${-days} วัน`, days, due };
    if (days <= 30) return { code: 'soon', label: 'ภายใน 30 วัน', days, due };
    return { code: 'ok', label: 'ตามแผน', days, due };
  }

  /**
   * @param {string} kind
   * @param ctx {h, T, tpl, state, getVal, ctlValue, parseIds, settings:{asOf, includeNotes}, score}
   */
  function build(kind, ctx) {
    const meta = KINDS[kind];
    const el = { summary: summary, detail: detail, findings: openFindings }[kind](ctx);
    const stamp = ctx.settings.asOf.replace(/-/g, '');
    return { el, title: meta.title, orientation: meta.orientation, fileTitle: `ITAudit2569_${meta.file}_${stamp}` };
  }

  // ---------- shared blocks ----------
  function header(ctx, title) {
    const { h, T, getVal } = ctx;
    const B = T.basic.sheet, g = r => getVal(B, r);
    const auditor = g('C16') + (g('C18') ? ` (${g('C18')})` : '');
    const board = [g('I16') ? `คณะกรรมการบริษัท ${dmy(g('I16'))}` : '', g('I18') ? `Audit Committee ${dmy(g('I18'))}` : ''].filter(Boolean).join(' · ') || '–';
    return h('header', { class: 'rep-head' },
      h('div', { class: 'rep-conf' }, 'ข้อมูลลับ — ใช้ภายในเท่านั้น'),
      h('h1', {}, title),
      h('p', { class: 'rep-sub' }, 'แบบรายงานผลการตรวจสอบด้านเทคโนโลยีสารสนเทศ (ก.ล.ต.) รอบปี 2569'),
      h('table', { class: 'rep-meta' }, h('tbody', {},
        h('tr', {}, h('th', {}, 'บริษัท'), h('td', {}, g('C5') || '–'), h('th', {}, 'ระดับ RLA'), h('td', {}, g('C26') || '–')),
        h('tr', {}, h('th', {}, 'ระยะเวลาตรวจสอบ'), h('td', {}, `${dmy(g('C23'))} – ${dmy(g('E23'))}`), h('th', {}, 'ผู้ดำเนินการตรวจสอบ'), h('td', {}, auditor || '–')),
        h('tr', {}, h('th', {}, 'เสนอผลต่อ'), h('td', {}, board), h('th', {}, 'ข้อมูล ณ วันที่'), h('td', {}, dmy(ctx.settings.asOf))))));
  }
  function blankWarning(ctx) {
    const n = ctx.score.blankScored;
    if (!n) return null;
    return ctx.h('div', { class: 'rep-warn' },
      `⚠ คะแนนยังไม่สมบูรณ์ — มี ${n} ข้อที่ยังไม่ได้ประเมิน ซึ่งสูตรของแบบฟอร์มนับเป็น 0 คะแนน (ตัวเลขด้านล่างจะเพิ่มขึ้นเมื่อประเมินครบ)`);
  }
  /** ผู้ประกอบธุรกิจขนาดเล็ก: แบบฟอร์มสรุปผลใน sheet 6.1 (Cyber Hygiene) ซึ่งรายงานนี้ไม่ได้คำนวณ */
  function smallWarning(ctx) {
    if (!ctx.score.isSmall) return null;
    return ctx.h('div', { class: 'rep-warn strong' },
      '⚠ RLA = ผู้ประกอบธุรกิจขนาดเล็ก: แบบฟอร์มสรุปผลของกลุ่มนี้ใน sheet "6.1 Result (Cyber Hygiene)" (เฉพาะ control ที่ติดเครื่องหมาย S) — คะแนนในรายงานนี้คำนวณตาม sheet "6. Result" (Basic controls ทั้งหมด) จึงไม่ใช่ผลของกลุ่มนี้ ให้ใช้ตัวเลขจาก sheet 6.1 ในไฟล์ Excel ที่ Export');
  }
  const ADV_NA_NOTE = 'Advanced controls ใช้กับผู้ประกอบธุรกิจที่ RLA สูงเท่านั้น';
  function findingStats(ctx) {
    const c = { H: 0, M: 0, L: 0, total: 0, open: 0, closed: 0, overdue: 0 };
    for (const f of ctx.state.findings) {
      c.total++;
      if (c[f.F] !== undefined) c[f.F]++;
      if (f.status === 'Closed') c.closed++;
      else { c.open++; if (aging(f, ctx.settings.asOf).code === 'overdue') c.overdue++; }
    }
    return c;
  }
  function label(ctx, row) {
    const l = ctx.tpl.resultLabels[row] || {};
    return { key: l.key || '', en: l.en || '', th: l.th || '' };
  }
  function ragText(r) { return { red: 'ต่ำกว่า 50%', amber: '50–79%', green: '80% ขึ้นไป' }[r]; }

  function nistChart(ctx) {
    const { score } = ctx;
    const NS = 'http://www.w3.org/2000/svg';
    const W = 330, rowH = 24, left = 72, barW = 200;
    const svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('viewBox', `0 0 ${W} ${score.nist.length * rowH + 18}`);
    svg.setAttribute('class', 'rep-nist');
    const mk = (tag, attrs, text) => { const e = document.createElementNS(NS, tag); for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v); if (text !== undefined) e.textContent = text; svg.append(e); return e; };
    for (const p of [0, 0.5, 0.8, 1]) {
      mk('line', { x1: left + barW * p, x2: left + barW * p, y1: 0, y2: score.nist.length * rowH, stroke: '#d5dae0', 'stroke-dasharray': p === 0 || p === 1 ? '' : '3 3' });
      mk('text', { x: left + barW * p, y: score.nist.length * rowH + 13, 'font-size': 9, 'text-anchor': 'middle', fill: '#6a7380' }, (p * 100) + '%');
    }
    score.nist.forEach((n, i) => {
      const y = i * rowH;
      mk('text', { x: left - 6, y: y + 15, 'font-size': 11, 'text-anchor': 'end', fill: '#1d232b' }, n.label);
      const v = n.value === null ? 0 : Math.max(0, Math.min(1, n.value));
      mk('rect', { x: left, y: y + 5, width: barW, height: 13, fill: '#eef1f4', rx: 2 });
      mk('rect', { x: left, y: y + 5, width: barW * v, height: 13, fill: '#1f6f9f', rx: 2 });
      mk('text', { x: left + barW + 6, y: y + 15, 'font-size': 10, fill: '#1d232b' }, n.value === null ? '–' : pct(n.value));
    });
    return svg;
  }

  // ---------- ① 1-page Summary ----------
  function summary(ctx) {
    const { h, score, state } = ctx;
    const t = score.total;
    const fs = findingStats(ctx);
    const assessed = t.D - t.blank;
    const domRows = score.rows.filter(r => r.level === 1);
    const kpi = (lbl, val, sub, cls = '') => h('div', { class: 'kpi ' + cls }, h('div', { class: 'kpi-l' }, lbl), h('div', { class: 'kpi-v' }, val), sub ? h('div', { class: 'kpi-s' }, sub) : null);
    const top = state.findings.map((f, i) => ({ f, i }))
      .filter(x => x.f.status !== 'Closed')
      .sort((a, b) => (ratingOrder[a.f.F] ?? 9) - (ratingOrder[b.f.F] ?? 9) || String(a.f.H || '9999').localeCompare(String(b.f.H || '9999')))
      .slice(0, 5);
    const advTotal = t.adv.N;
    return h('section', { class: 'report portrait' },
      header(ctx, 'สรุปผลการตรวจสอบด้าน IT — 1-page Summary'),
      smallWarning(ctx),
      blankWarning(ctx),
      h('div', { class: 'kpis' },
        kpi('ความคืบหน้าการประเมิน (Basic)', `${assessed}/${t.D}`, `${t.D ? Math.round(assessed / t.D * 100) : 0}% ของ Basic controls`),
        kpi('คะแนนการปฏิบัติตามเกณฑ์', `${num(t.K)} / 3`, 'Basic controls'),
        kpi('คะแนนรวม (รวมคะแนนพิเศษ M4/M5)', `${num(t.M)} / 5`, ''),
        score.advApplies
          ? kpi('Advanced controls compliance', pct(t.adv.S), advTotal ? `${ragText(t.rag)} · ${t.adv.R}/${t.adv.N - t.adv.O} ข้อ` : 'ไม่มีข้อ advanced', 'rag-' + t.rag)
          : kpi('Advanced controls compliance', 'ไม่เข้าข่าย', 'เฉพาะ RLA สูง', 'kpi-na'),
        kpi('Finding', String(fs.total), `H ${fs.H} · M ${fs.M} · L ${fs.L} · Open ${fs.open}`)),
      h('h2', {}, 'ผลราย Domain (Basic controls)'),
      h('table', { class: 'rep-table' },
        h('thead', {}, h('tr', {}, ['Domain', 'Controls', 'ประเมินแล้ว', 'N/A', 'No/M1', 'Partial/M2', 'Yes/M3', 'M4', 'M5', 'คะแนน (3)', 'รวม (5)', 'Advanced %'].map(x => h('th', {}, x)))),
        h('tbody', {}, domRows.map(r => {
          const l = label(ctx, r.row);
          return h('tr', {}, h('td', {}, `D${l.key || r.key} ${l.en}`), h('td', { class: 'n' }, r.D), h('td', { class: 'n' }, r.D - r.blank), h('td', { class: 'n' }, r.E), h('td', { class: 'n' }, r.F), h('td', { class: 'n' }, r.G), h('td', { class: 'n' }, r.H), h('td', { class: 'n' }, r.I), h('td', { class: 'n' }, r.J), h('td', { class: 'n b' }, num(r.K)), h('td', { class: 'n b' }, num(r.M)), h('td', { class: 'n' }, r.adv && score.advApplies ? pct(r.adv.S) : '–'));
        })),
        h('tfoot', {}, h('tr', {}, h('td', {}, 'รวม'), h('td', { class: 'n' }, t.D), h('td', { class: 'n' }, assessed), h('td', { class: 'n' }, t.E), h('td', { class: 'n' }, t.F), h('td', { class: 'n' }, t.G), h('td', { class: 'n' }, t.H), h('td', { class: 'n' }, t.I), h('td', { class: 'n' }, t.J), h('td', { class: 'n b' }, num(t.K)), h('td', { class: 'n b' }, num(t.M)), h('td', { class: 'n' }, score.advApplies ? pct(t.adv.S) : '–')))),
      score.advApplies ? null : h('p', { class: 'rep-note' }, ADV_NA_NOTE + ' — คอลัมน์ Advanced % จึงไม่แสดงผลสำหรับ RLA นี้'),
      h('div', { class: 'two-col' },
        h('div', {},
          h('h2', {}, 'NIST CSF 2.0 (ตามระดับ RLA)'),
          nistChart(ctx),
          score.rlaQuirk ? h('p', { class: 'rep-note' }, 'หมายเหตุ: สูตรของแบบฟอร์ม (RawData!AB276:AF276) เทียบ RLA กับ "ต่ำ**" กรณี RLA ต่ำ จึงแสดงค่าของกลุ่มผู้ประกอบธุรกิจขนาดเล็ก') : null,
          score.rcFalse ? h('p', { class: 'rep-note' }, 'หมายเหตุ: สูตร Recover ของแบบฟอร์ม (RawData!AG276) ให้ค่า FALSE (แสดงเป็น 0%) สำหรับ RLA นี้') : null),
        h('div', {},
          h('h2', {}, 'สถานะ Finding'),
          h('table', { class: 'rep-table small' },
            h('thead', {}, h('tr', {}, h('th', {}, ''), h('th', {}, 'H'), h('th', {}, 'M'), h('th', {}, 'L'), h('th', {}, 'รวม'))),
            h('tbody', {}, ['Open', 'Closed'].map(st => {
              const list = state.findings.filter(f => (f.status === 'Closed') === (st === 'Closed'));
              const c = k => list.filter(f => f.F === k).length;
              return h('tr', {}, h('td', {}, st), h('td', { class: 'n' }, c('H')), h('td', { class: 'n' }, c('M')), h('td', { class: 'n' }, c('L')), h('td', { class: 'n b' }, list.length));
            }))),
          h('p', { class: 'rep-note' }, `เลยกำหนด (ณ ${dmy(ctx.settings.asOf)}): ${fs.overdue} ข้อ`))),
      h('h2', {}, 'Finding ที่ต้องติดตาม (เรียงตามระดับความสำคัญ สูงสุด 5 ข้อ)'),
      top.length ? h('table', { class: 'rep-table' },
        h('thead', {}, h('tr', {}, ['#', 'Control', 'ข้อตรวจพบ', 'ระดับ', 'กำหนดเสร็จ', 'สถานะ'].map(x => h('th', {}, x)))),
        h('tbody', {}, top.map(({ f, i }) => {
          const a = aging(f, ctx.settings.asOf);
          return h('tr', {}, h('td', { class: 'n' }, i + 1), h('td', { class: 'nowrap' }, ctx.parseIds(f.C).map(nb).join(', ')), h('td', {}, trunc(f.E, 150)), h('td', { class: 'n rating-' + f.F }, f.F || '–'), h('td', { class: 'n' }, my(f.H)), h('td', { class: 'age-' + a.code }, a.label));
        })))
        : h('p', { class: 'rep-note' }, 'ไม่มี Finding ที่ยังเปิดอยู่'),
      h('p', { class: 'rep-foot' }, 'วิธีคำนวณตาม sheet "6. Result" ของแบบฟอร์ม: คะแนนการปฏิบัติตามเกณฑ์ = (No×1 + Partial×2 + Yes×3 + M4×3 + M5×3) ÷ (จำนวนข้อ − N/A) · คะแนนพิเศษ = (M4×1 + M5×2) ÷ จำนวนข้อ Maturity ที่ไม่ใช่ N/A · Advanced compliance = Yes ÷ (จำนวนข้อ − N/A)'));
  }

  // ---------- ② Detailed Summary ----------
  function groupKeyOf(c) {
    const sh = c.sheet.replace(/^4\.\s*/, '');
    if (sh === 'D1') return c.cat;                 // 1.1 / 1.2 / 1.3
    if (sh === 'D3') return '3';
    if (sh === 'D2.2' || sh === 'D2.8') return c.cat; // 2.2.x / 2.8.x
    return sh.slice(1);                              // D2.x → 2.x
  }
  function detail(ctx) {
    const { h, score, tpl, state } = ctx;
    const linked = id => state.findings.map((f, i) => [f, i]).filter(([f]) => ctx.parseIds(f.C).includes(id)).map(([, i]) => '#' + (i + 1));
    const isGap = v => ['No', 'Partial', 'M1', 'M2'].includes(v);
    const rowsEl = score.rows.map(r => {
      const l = label(ctx, r.row);
      const a = r.adv;
      return h('tr', { class: 'lv' + r.level },
        h('td', { class: 'nowrap' }, l.key || r.key), h('td', {}, l.en, l.th ? h('div', { class: 'th' }, l.th) : null),
        h('td', { class: 'n' }, r.D), h('td', { class: 'n' }, r.E), h('td', { class: 'n' }, r.F), h('td', { class: 'n' }, r.G), h('td', { class: 'n' }, r.H), h('td', { class: 'n' }, r.I), h('td', { class: 'n' }, r.J), h('td', { class: 'n' }, r.blank),
        h('td', { class: 'n b' }, num(r.K)), h('td', { class: 'n' }, r.L === null ? '–' : num(r.L)), h('td', { class: 'n b' }, num(r.M)),
        h('td', { class: 'n sep' }, a ? a.N : '–'), h('td', { class: 'n' }, a ? a.O : '–'), h('td', { class: 'n' }, a ? a.P : '–'), h('td', { class: 'n' }, a ? a.Q : '–'), h('td', { class: 'n' }, a ? a.R : '–'), h('td', { class: 'n b' }, a ? pct(a.S) : '–'));
    });
    const t = score.total;
    const table = h('table', { class: 'rep-table detail' },
      h('thead', {},
        h('tr', {}, h('th', { colspan: 2 }, ''), h('th', { colspan: 11, class: 'grp' }, 'Basic controls'), h('th', { colspan: 6, class: 'grp sep' }, 'Advanced controls (ผปก. เสี่ยงสูง)')),
        h('tr', {}, ['#', 'หัวข้อ', 'Controls', 'N/A', 'No/M1', 'Partial/M2', 'Yes/M3', 'M4', 'M5', 'ว่าง', 'คะแนน (3)', 'พิเศษ (2)', 'รวม (5)'].map(x => h('th', {}, x)),
          ['Controls', 'N/A', 'No', 'Partial', 'Yes', 'Compliance'].map((x, i) => h('th', { class: i === 0 ? 'sep' : '' }, x)))),
      // แถว Total อยู่ใน tbody (ถ้าใช้ tfoot Chrome จะพิมพ์ซ้ำทุกหน้า)
      h('tbody', {}, rowsEl, h('tr', { class: 'total' }, h('td', {}, ''), h('td', {}, 'Total'),
        h('td', { class: 'n' }, t.D), h('td', { class: 'n' }, t.E), h('td', { class: 'n' }, t.F), h('td', { class: 'n' }, t.G), h('td', { class: 'n' }, t.H), h('td', { class: 'n' }, t.I), h('td', { class: 'n' }, t.J), h('td', { class: 'n' }, t.blank),
        h('td', { class: 'n b' }, num(t.K)), h('td', { class: 'n' }, '–'), h('td', { class: 'n b' }, num(t.M)),
        h('td', { class: 'n sep' }, t.adv.N), h('td', { class: 'n' }, t.adv.O), h('td', { class: 'n' }, t.adv.P), h('td', { class: 'n' }, t.adv.Q), h('td', { class: 'n' }, t.adv.R), h('td', { class: 'n b' + (score.advApplies ? ' rag-' + t.rag : '') }, pct(t.adv.S)))));

    // gap list ราย sub-domain
    const groups = new Map();
    for (const c of tpl.controls) {
      const k = groupKeyOf(c);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(c);
    }
    const labelByKey = {};
    for (const r of score.rows) { const l = label(ctx, r.row); labelByKey[l.key || r.key] = l; }
    const gapSections = [];
    for (const [k, list] of groups) {
      const gaps = list.filter(c => isGap(ctx.ctlValue(c)));
      const blanks = list.filter(c => !c.fixed && !ctx.ctlValue(c));
      if (!gaps.length && !blanks.length) continue;
      const l = labelByKey[k] || { en: '', th: '' };
      gapSections.push(h('div', { class: 'gap-sec' },
        h('h3', {}, `${k} ${l.en}`, l.th ? ` — ${l.th}` : ''),
        gaps.length ? h('table', { class: 'rep-table gaps' },
          h('thead', {}, h('tr', {}, h('th', { class: 'c-id' }, 'Control'), h('th', { class: 'c-res' }, 'ผล'), h('th', {}, 'การควบคุมที่ควรมี'), h('th', { class: 'c-fd' }, 'Finding'), ctx.settings.includeNotes ? h('th', { class: 'c-note' }, 'บันทึกผู้ตรวจ / Evidence') : null)),
          h('tbody', {}, gaps.map(c => {
            const n = state.notes[c.id] || {};
            return h('tr', {},
              h('td', { class: 'id' }, nb(c.id), c.adv ? h('span', { class: 'adv' }, ' ADV') : null),
              h('td', { class: 'n res-' + ctx.ctlValue(c) }, ctx.ctlValue(c)),
              h('td', {}, trunc(c.text, 260)),
              h('td', { class: 'nowrap' }, linked(c.id).join(', ') || '–'),
              ctx.settings.includeNotes ? h('td', {}, [n.note, n.evidence ? `Evidence: ${n.evidence}` : ''].filter(Boolean).join(' · ') || '–') : null);
          }))) : null,
        blanks.length ? h('p', { class: 'rep-note' }, `ยังไม่ประเมิน ${blanks.length} ข้อ: ${blanks.map(c => c.id).join(', ')}`) : null));
    }

    let notesAppendix = null;
    if (ctx.settings.includeNotes) {
      const withNotes = tpl.controls.filter(c => state.notes[c.id] && (state.notes[c.id].note || state.notes[c.id].evidence));
      notesAppendix = h('div', { class: 'gap-sec' },
        h('h2', {}, 'ภาคผนวก: บันทึกผู้ตรวจทั้งหมด'),
        withNotes.length ? h('table', { class: 'rep-table gaps' },
          h('thead', {}, h('tr', {}, h('th', { class: 'c-id' }, 'Control'), h('th', { class: 'c-res' }, 'ผล'), h('th', {}, 'Working note'), h('th', { class: 'c-note' }, 'Evidence ref'))),
          h('tbody', {}, withNotes.map(c => h('tr', {}, h('td', { class: 'id' }, nb(c.id)), h('td', { class: 'n' }, ctx.ctlValue(c) || '–'), h('td', {}, state.notes[c.id].note || '–'), h('td', {}, state.notes[c.id].evidence || '–')))))
          : h('p', { class: 'rep-note' }, 'ไม่มีบันทึก'));
    }

    return h('section', { class: 'report landscape' },
      header(ctx, 'สรุปผลการตรวจสอบด้าน IT — Detailed Summary'),
      smallWarning(ctx),
      blankWarning(ctx),
      h('h2', {}, 'ผลราย Domain / Sub-domain (ตาม sheet "6. Result")'),
      table,
      score.advApplies
        ? h('p', { class: 'rep-note' }, `เกณฑ์ Advanced compliance: ต่ำกว่า 50% / 50–79% / 80% ขึ้นไป · ผลรวมปัจจุบัน: ${pct(t.adv.S)} (${ragText(t.rag)})`)
        : h('p', { class: 'rep-note' }, ADV_NA_NOTE + ' — ตัวเลขฝั่ง Advanced ในตารางแสดงตามสูตรของ sheet "6. Result" และไม่ใช้ประเมินผลสำหรับ RLA นี้'),
      h('h2', { class: 'break-before' }, 'รายการที่ไม่เป็นไปตามเกณฑ์ (No / Partial / M1 / M2) และข้อที่ยังไม่ประเมิน'),
      gapSections.length ? gapSections : h('p', { class: 'rep-note' }, 'ไม่มีข้อที่ไม่เป็นไปตามเกณฑ์'),
      notesAppendix);
  }

  // ---------- ③ Finding คงค้าง ----------
  function openFindings(ctx) {
    const { h, state } = ctx;
    const asOf = ctx.settings.asOf;
    const open = state.findings.map((f, i) => ({ f, i, a: aging(f, asOf) })).filter(x => x.f.status !== 'Closed');
    open.sort((x, y) => (x.a.due ?? Infinity) - (y.a.due ?? Infinity) || (ratingOrder[x.f.F] ?? 9) - (ratingOrder[y.f.F] ?? 9));
    const closed = state.findings.filter(f => f.status === 'Closed').length;
    const codes = [['overdue', 'เลยกำหนด'], ['soon', 'ภายใน 30 วัน'], ['ok', 'ตามแผน'], ['none', 'ไม่ระบุกำหนด']];
    const cnt = (r, code) => open.filter(x => (r === '*' || x.f.F === r) && (code === '*' || x.a.code === code)).length;
    return h('section', { class: 'report landscape' },
      header(ctx, 'รายงาน Finding คงค้าง (Open Findings)'),
      h('p', { class: 'rep-note' }, `ขอบเขต: ข้อตรวจพบ/ข้อบกพร่องของรอบปี 2569 ที่สถานะยังเป็น Open ณ วันที่ ${dmy(asOf)} · ปิดแล้ว ${closed} ข้อ (ไม่แสดงในรายงานนี้) · "กำหนดเสร็จ" นับถึงวันสุดท้ายของเดือนเป้าหมาย`),
      h('table', { class: 'rep-table small matrix' },
        h('thead', {}, h('tr', {}, h('th', {}, 'ระดับ'), codes.map(([, l]) => h('th', {}, l)), h('th', {}, 'รวม'))),
        h('tbody', {}, ['H', 'M', 'L'].map(r => h('tr', {}, h('td', { class: 'rating-' + r }, r), codes.map(([c]) => h('td', { class: 'n' + (c === 'overdue' && cnt(r, c) ? ' age-overdue' : '') }, cnt(r, c))), h('td', { class: 'n b' }, cnt(r, '*'))))),
        h('tfoot', {}, h('tr', {}, h('td', {}, 'รวม'), codes.map(([c]) => h('td', { class: 'n' }, cnt('*', c))), h('td', { class: 'n b' }, open.length)))),
      open.length ? h('table', { class: 'rep-table findings' },
        h('thead', {}, h('tr', {}, ['#', 'Control', 'หมวด / หัวข้อ', 'ข้อตรวจพบ', 'ระดับ', 'แผนการแก้ไข', 'กำหนดเสร็จ', 'สถานะเวลา', 'หมายเหตุ'].map(x => h('th', {}, x)))),
        h('tbody', {}, open.map(({ f, i, a }) => h('tr', {},
          h('td', { class: 'n' }, i + 1),
          h('td', { class: 'nowrap' }, ctx.parseIds(f.C).map(nb).join(',\n')),
          h('td', {}, f.B || '–', f.D ? h('div', { class: 'th' }, f.D) : null),
          h('td', { class: 'pre' }, f.E || '–'),
          h('td', { class: 'n rating-' + f.F }, f.F || '–'),
          h('td', { class: 'pre' }, f.G || '–'),
          h('td', { class: 'n' }, my(f.H)),
          h('td', { class: 'age-' + a.code }, a.label),
          h('td', {}, f.I || '')))))
        : h('p', { class: 'rep-note' }, 'ไม่มี Finding ที่ยังเปิดอยู่'));
  }

  return { build, KINDS, aging };
})();
