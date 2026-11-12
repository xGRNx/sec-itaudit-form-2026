/*
 * scoring.js — คำนวณผลตาม logic ของ sheet "6. Result" และ "RawData" ในแบบฟอร์ม ก.ล.ต. รอบปี 2569
 *
 * ตั้งใจทำให้ตัวเลข "ตรงกับไฟล์ Excel" ทุกจุด รวมถึงพฤติกรรมเฉพาะของสูตรต้นฉบับ เช่น
 *  - ข้อที่ยังว่างถูกนับในตัวหาร (ได้ 0 คะแนน)
 *  - คะแนนพิเศษ (L) ของ sub-domain ใน D1 ใช้ตัวหารระดับ domain
 *  - กราฟ NIST เทียบ RLA กับ "ต่ำ**" (RLA ต่ำ จึงได้ค่าของกลุ่มขนาดเล็ก)
 * แต่ละแถวอ้างอิงเลขแถวใน sheet "6. Result" (ฟิลด์ row) เพื่อให้ตรวจย้อนได้
 */
'use strict';

const Scoring = (() => {
  // crit: [field, value] โดย field = dom (RawData D) | sheet (RawData B) | cat (RawData E)
  // basic.sum / adv.sum = แถวที่สูตรต้นฉบับใช้ SUM() ของแถวลูก
  const ROWS = [
    { row: 6, key: '1', level: 1, basic: { D: ['dom', 'D1'], sum: [7, 8, 9] }, L: ['dom', 'D1'], adv: { N: ['dom', 'D1'], sum: [7, 8, 9] } },
    { row: 7, key: '1.1', level: 2, basic: { crit: ['cat', '1.1'] }, L: ['dom', 'D1'], adv: null },
    { row: 8, key: '1.2', level: 2, basic: { crit: ['cat', '1.2'] }, L: ['dom', 'D1'], adv: { crit: ['cat', '1.2'] } },
    { row: 9, key: '1.3', level: 2, basic: { crit: ['cat', '1.3'] }, L: ['dom', 'D1'], adv: { crit: ['cat', '1.3'] } },
    { row: 10, key: '2', level: 1, basic: { D: ['dom', 'D2'], sum: [11, 12, 15, 16, 17, 18, 19, 20, 32, 33, 34, 35] }, L: ['dom', 'D2'], adv: { N: ['dom', 'D2'], sum: [11, 12, 15, 16, 17, 18, 19, 20, 32, 33, 34, 35] } },
    { row: 11, key: '2.1', level: 2, basic: { crit: ['sheet', 'D2.1'] }, L: ['sheet', 'D2.1'], adv: null },
    { row: 12, key: '2.2', level: 2, basic: { sum: [13, 14], Dsum: true }, L: ['sheet', 'D2.2'], adv: { same: 14 } },
    { row: 13, key: '2.2.1', level: 3, basic: { crit: ['cat', '2.2.1'] }, L: ['cat', '2.2.1'], adv: null },
    { row: 14, key: '2.2.2', level: 3, basic: { crit: ['cat', '2.2.2'] }, L: ['cat', '2.2.2'], adv: { crit: ['cat', '2.2.2'] } },
    { row: 15, key: '2.3', level: 2, basic: { crit: ['cat', '2.3'] }, L: ['sheet', 'D2.3'], adv: null },
    { row: 16, key: '2.4', level: 2, basic: { crit: ['cat', '2.4'] }, L: ['sheet', 'D2.4'], adv: null },
    { row: 17, key: '2.5', level: 2, basic: { crit: ['cat', '2.5'] }, L: ['sheet', 'D2.5'], adv: { crit: ['cat', '2.5'] } },
    { row: 18, key: '2.6', level: 2, basic: { crit: ['cat', '2.6'] }, L: ['sheet', 'D2.6'], adv: null },
    { row: 19, key: '2.7', level: 2, basic: { crit: ['cat', '2.7'] }, L: ['sheet', 'D2.7'], adv: null },
    { row: 20, key: '2.8', level: 2, basic: { sum: [21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31], Dsum: true }, L: ['sheet', 'D2.8'], adv: { sum: [21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31], Nsum: true } },
    { row: 21, key: '2.8.1', level: 3, basic: { crit: ['cat', '2.8.1'] }, L: ['cat', '2.8.1'], adv: null },
    { row: 22, key: '2.8.2', level: 3, basic: { crit: ['cat', '2.8.2'] }, L: ['cat', '2.8.2'], adv: { crit: ['cat', '2.8.2'] } },
    { row: 23, key: '2.8.3', level: 3, basic: { crit: ['cat', '2.8.3'] }, L: ['cat', '2.8.3'], adv: null },
    { row: 24, key: '2.8.4', level: 3, basic: { crit: ['cat', '2.8.4'] }, L: ['cat', '2.8.4'], adv: { crit: ['cat', '2.8.4'] } },
    { row: 25, key: '2.8.5', level: 3, basic: { crit: ['cat', '2.8.5'] }, L: ['cat', '2.8.5'], adv: { crit: ['cat', '2.8.5'] } },
    { row: 26, key: '2.8.6', level: 3, basic: { crit: ['cat', '2.8.6'] }, L: ['cat', '2.8.6'], adv: null },
    { row: 27, key: '2.8.7', level: 3, basic: { crit: ['cat', '2.8.7'] }, L: ['cat', '2.8.7'], adv: null },
    { row: 28, key: '2.8.8', level: 3, basic: { crit: ['cat', '2.8.8'] }, L: ['cat', '2.8.8'], adv: { crit: ['cat', '2.8.8'] } },
    { row: 29, key: '2.8.9', level: 3, basic: { crit: ['cat', '2.8.9'] }, L: ['cat', '2.8.9'], adv: null },
    { row: 30, key: '2.8.10', level: 3, basic: { crit: ['cat', '2.8.10'] }, L: ['cat', '2.8.10'], adv: { crit: ['cat', '2.8.10'] } },
    { row: 31, key: '2.8.11', level: 3, basic: { crit: ['cat', '2.8.11'] }, L: ['cat', '2.8.11'], adv: { crit: ['cat', '2.8.11'] } },
    { row: 32, key: '2.9', level: 2, basic: { crit: ['cat', '2.9'] }, L: ['sheet', 'D2.9'], adv: { crit: ['cat', '2.9'] } },
    { row: 33, key: '2.10', level: 2, basic: { crit: ['sheet', 'D2.10'] }, L: ['sheet', 'D2.10'], adv: { crit: ['sheet', 'D2.10'] } },
    { row: 34, key: '2.11', level: 2, basic: { crit: ['cat', '2.11'] }, L: ['sheet', 'D2.11'], adv: { crit: ['cat', '2.11'] } },
    { row: 35, key: '2.12', level: 2, basic: { crit: ['cat', '2.12'] }, L: ['sheet', 'D2.12'], adv: { crit: ['cat', '2.12'] } },
    { row: 36, key: '3', level: 1, basic: { crit: ['cat', '3'] }, L: null, adv: null, sBlankIfZero: true },
  ];
  const DOMAIN_ROWS = [6, 10, 36];

  const NIST = [
    { key: 'GV', label: 'Govern' }, { key: 'ID', label: 'Identify' }, { key: 'PR', label: 'Protect' },
    { key: 'DE', label: 'Detect' }, { key: 'RS', label: 'Respond' }, { key: 'RC', label: 'Recover' },
  ];

  const bucket = v => {
    if (v === 'N/A') return 'E';
    if (v === 'No' || v === 'M1') return 'F';
    if (v === 'Partial' || v === 'M2') return 'G';
    if (v === 'Yes' || v === 'M3') return 'H';
    if (v === 'M4') return 'I';
    if (v === 'M5') return 'J';
    return null; // ว่าง
  };
  const qScore = v => {
    if (v === 'No' || v === 'M1') return 1;
    if (v === 'Partial' || v === 'M2') return 2;
    if (['Yes', 'M3', 'M4', 'M5'].includes(v)) return 3;
    return null;
  };

  /**
   * @param {Array} controls  [{id, domain, sheet, cat, adv, type, appl:{H,M,L,S}, nist}]
   * @param {(c)=>string} valueOf  ผลประเมินของ control ('' = ว่าง)
   * @param {string} rla  ค่าใน 1. Basic info!C26
   */
  function compute(controls, valueOf, rla) {
    const items = controls.map(c => ({ c, v: valueOf(c) || '' }));
    const fieldOf = (c, f) => f === 'dom' ? c.domain : f === 'sheet' ? c.sheet.replace(/^4\.\s*/, '') : c.cat;
    const match = (c, crit) => fieldOf(c, crit[0]) === crit[1];

    const byRow = {};
    // ทำจากแถวลูกขึ้นไปแถวแม่ (แถวที่มี sum ต้องคำนวณทีหลัง)
    const rank = sp => sp.basic.sum ? 10 - sp.level : 0; // แถวลูกก่อน แล้ว sum ระดับ 2 แล้วค่อยระดับ 1
    const order = [...ROWS].sort((a, b) => rank(a) - rank(b));
    for (const spec of order) {
      const r = { row: spec.row, key: spec.key, level: spec.level, D: 0, E: 0, F: 0, G: 0, H: 0, I: 0, J: 0 };
      const basicItems = crit => items.filter(x => !x.c.adv && match(x.c, crit));
      if (spec.basic.crit) {
        const list = basicItems(spec.basic.crit);
        r.D = list.length;
        for (const x of list) { const b = bucket(x.v); if (b) r[b]++; }
        r.blank = list.filter(x => !bucket(x.v)).length;
      } else {
        const kids = spec.basic.sum.map(n => byRow[n]);
        for (const k of 'EFGHIJ') r[k] = kids.reduce((s, x) => s + x[k], 0);
        r.D = spec.basic.Dsum ? kids.reduce((s, x) => s + x.D, 0) : basicItems(spec.basic.D).length;
        r.blank = kids.reduce((s, x) => s + x.blank, 0);
      }
      r.K = (r.D - r.E) === 0 ? 'N/A' : (r.F + 2 * r.G + 3 * r.H + 3 * r.I + 3 * r.J) / (r.D - r.E);
      if (spec.L) {
        const mat = items.filter(x => !x.c.adv && match(x.c, spec.L) && x.c.type === 'Maturity');
        const den = mat.length - mat.filter(x => x.v === 'N/A').length;
        r.L = den === 0 ? 0 : (r.I + 2 * r.J) / den;
        r.M = r.K === 'N/A' ? 'N/A' : r.K + r.L;
      } else {
        r.L = null;
        r.M = r.K;
      }
      // advanced controls
      if (spec.adv) {
        let a;
        if (spec.adv.same) a = { ...byRow[spec.adv.same].advRaw };
        else if (spec.adv.crit) {
          const list = items.filter(x => x.c.adv && match(x.c, spec.adv.crit));
          a = { N: list.length, O: 0, P: 0, Q: 0, R: 0 };
          for (const x of list) { if (x.v === 'N/A') a.O++; else if (x.v === 'No') a.P++; else if (x.v === 'Partial') a.Q++; else if (x.v === 'Yes') a.R++; }
        } else {
          const kids = spec.adv.sum.map(n => byRow[n].advRaw || { N: 0, O: 0, P: 0, Q: 0, R: 0 });
          a = { O: 0, P: 0, Q: 0, R: 0 };
          for (const k of 'OPQR') a[k] = kids.reduce((s, x) => s + x[k], 0);
          a.N = spec.adv.Nsum ? kids.reduce((s, x) => s + x.N, 0) : items.filter(x => x.c.adv && match(x.c, spec.adv.N)).length;
        }
        a.S = (a.N - a.O) === 0 ? 0 : a.R / (a.N - a.O);
        r.advRaw = a;
        r.adv = a;
      } else {
        r.advRaw = { N: 0, O: 0, P: 0, Q: 0, R: 0 };
        r.adv = null;
      }
      byRow[spec.row] = r;
    }

    // แถว Total (6. Result แถว 37)
    const doms = DOMAIN_ROWS.map(n => byRow[n]);
    const t = { row: 37, key: 'Total' };
    for (const k of 'DEFGHIJ') t[k] = doms.reduce((s, x) => s + x[k], 0);
    t.blank = doms.reduce((s, x) => s + x.blank, 0);
    t.K = (t.D - t.E) === 0 ? 'N/A' : (t.F + 2 * t.G + 3 * t.H + 3 * t.I + 3 * t.J) / (t.D - t.E);
    const matAll = items.filter(x => x.c.type === 'Maturity' && x.v !== 'N/A').length;
    t.M = (t.K === 'N/A' || matAll === 0) ? 'N/A' : t.K + (t.I + 2 * t.J) / matAll;
    const advN = doms.reduce((s, x) => s + x.advRaw.N, 0), advO = doms.reduce((s, x) => s + x.advRaw.O, 0), advR = doms.reduce((s, x) => s + x.advRaw.R, 0);
    t.adv = { N: advN, O: advO, R: advR, P: doms.reduce((s, x) => s + x.advRaw.P, 0), Q: doms.reduce((s, x) => s + x.advRaw.Q, 0) };
    t.adv.S = (advN - advO) === 0 ? 0 : advR / (advN - advO);
    t.rag = t.adv.S < 0.5 ? 'red' : t.adv.S < 0.8 ? 'amber' : 'green';

    // NIST CSF 2.0 (RawData AB276:AG276)
    const rlaCol = rla === 'สูง' ? 'H' : rla === 'ปานกลาง' ? 'M' : rla === 'ต่ำ**' ? 'L' : 'S';
    // RawData!AG276 (Recover) เทียบกับ "ผู้ประกอบธุรกิจขนาดเล็ก**" และไม่มีค่า else
    // → RLA ต่ำ / ขนาดเล็ก / ว่าง ได้ FALSE (Excel แสดงเป็น 0)
    const rcFalse = !['สูง', 'ปานกลาง'].includes(rla);
    const nist = NIST.map(fn => {
      if (fn.key === 'RC' && rcFalse) return { ...fn, value: 0, count: 0, formFalse: true };
      const token = `(${fn.key})`;
      const inFn = c => c.nistFlags ? !!c.nistFlags[fn.key] : (c.nist || '').toUpperCase().includes(token);
      const list = items.filter(x => x.c.appl[rlaCol] && inFn(x.c));
      const num = list.reduce((s, x) => s + (qScore(x.v) || 0), 0) / 3;
      const den = list.length - list.filter(x => x.v === 'N/A').length;
      return { ...fn, value: den === 0 ? null : num / den, count: list.length };
    });

    const allBlank = items.filter(x => !x.c.fixed && !x.v).length;
    // --- ข้อมูลประกอบการแสดงผล (ไม่ใช่ส่วนของสูตร 6. Result) ---
    // Advanced controls เข้าข่ายเฉพาะระดับ RLA ที่มีข้อ advanced ติดเครื่องหมาย x (ในแบบฟอร์มปี 2569 คือ RLA สูง)
    const trueCol = { 'สูง': 'H', 'ปานกลาง': 'M', 'ต่ำ': 'L', 'ผู้ประกอบธุรกิจขนาดเล็ก': 'S' }[rla] || null;
    const advApplies = !trueCol || controls.some(c => c.adv && c.appl[trueCol]);
    const advBlank = items.filter(x => x.c.adv && !x.c.fixed && !x.v).length;
    const blankScored = t.blank + (advApplies ? advBlank : 0);
    return {
      rows: ROWS.map(s => byRow[s.row]), total: t, nist, nistBasis: rlaCol, rlaQuirk: rla === 'ต่ำ', rcFalse, blank: allBlank,
      advApplies, isSmall: trueCol === 'S', blankScored,
    };
  }

  return { compute, ROWS, NIST };
})();

if (typeof module !== 'undefined') module.exports = Scoring;
