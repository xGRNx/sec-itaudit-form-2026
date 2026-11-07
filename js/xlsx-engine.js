/*
 * xlsx-engine.js — อ่าน/เขียนไฟล์ .xlsx แบบ "แก้เฉพาะเซลล์" (surgical cell injection)
 *
 * หลักการ:
 *  - อ่าน: แตกไฟล์ด้วย JSZip แล้ว parse XML ของแต่ละ sheet เพื่อดึงค่าเซลล์
 *  - เขียน: แก้ข้อความ XML ของ sheet เฉพาะเซลล์ที่ต้องกรอก ส่วนอื่นของไฟล์
 *    (กราฟ, สูตร, array formula, data validation, การป้องกัน sheet) ไม่ถูกแตะเลย
 *  - ตั้ง fullCalcOnLoad="1" ใน workbook.xml ให้ Excel คำนวณสูตรใหม่ทั้งหมดตอนเปิดไฟล์
 *
 * ทั้งหมดทำงานในเบราว์เซอร์ ไม่มีการส่งข้อมูลออกนอกเครื่อง
 */
'use strict';

const XL = (() => {
  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

  // ---------- helpers ----------
  function colToNum(col) {
    let n = 0;
    for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  function numToCol(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function splitRef(ref) {
    const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref);
    if (!m) throw new Error('Bad cell ref: ' + ref);
    return { col: m[1], row: parseInt(m[2], 10), c: colToNum(m[1]) };
  }
  function expandRange(range) {
    // 'A1:B3' หรือ 'A1' → รายการ ref
    const [a, b] = range.split(':');
    const s = splitRef(a), e = splitRef(b || a);
    const out = [];
    for (let r = s.row; r <= e.row; r++)
      for (let c = s.c; c <= e.c; c++) out.push(numToCol(c) + r);
    return out;
  }
  function inRange(ref, range) {
    const [a, b] = range.split(':');
    const s = splitRef(a), e = splitRef(b || a), x = splitRef(ref);
    return x.row >= s.row && x.row <= e.row && x.c >= s.c && x.c <= e.c;
  }
  function escXml(s) {
    return String(s)
      // ตัด control character ที่ XML ไม่รองรับ (ยกเว้น tab/newline)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function serialToISO(serial) {
    const d = new Date(EXCEL_EPOCH + Math.round(serial) * 86400000);
    return d.toISOString().slice(0, 10);
  }
  function isoToSerial(iso) {
    const m = /^(\d{4})-(\d{2})(?:-(\d{2}))?$/.exec(iso);
    if (!m) return null;
    return Math.round((Date.UTC(+m[1], +m[2] - 1, m[3] ? +m[3] : 1) - EXCEL_EPOCH) / 86400000);
  }

  // ---------- open / read ----------
  async function open(arrayBuffer) {
    const zip = await JSZip.loadAsync(arrayBuffer, { createFolders: false });
    const wbFile = zip.file('xl/workbook.xml');
    if (!wbFile) throw new Error('ไม่ใช่ไฟล์ Excel (.xlsx) ที่ถูกต้อง');
    const parser = new DOMParser();
    const wbDoc = parser.parseFromString(await wbFile.async('string'), 'application/xml');
    const relDoc = parser.parseFromString(await zip.file('xl/_rels/workbook.xml.rels').async('string'), 'application/xml');
    const rels = {};
    for (const r of relDoc.getElementsByTagName('Relationship')) rels[r.getAttribute('Id')] = r.getAttribute('Target');
    const sheets = {};
    const sheetOrder = [];
    for (const s of wbDoc.getElementsByTagNameNS(NS_MAIN, 'sheet')) {
      const target = rels[s.getAttributeNS(NS_REL, 'id')];
      const path = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      const name = s.getAttribute('name');
      sheets[name] = path;
      sheetOrder.push(name);
    }
    const sst = [];
    const sstFile = zip.file('xl/sharedStrings.xml');
    if (sstFile) {
      const doc = parser.parseFromString(await sstFile.async('string'), 'application/xml');
      for (const si of doc.getElementsByTagNameNS(NS_MAIN, 'si')) sst.push(textOfSi(si));
    }
    const meta = {};
    const coreFile = zip.file('docProps/core.xml');
    if (coreFile) {
      const core = await coreFile.async('string');
      for (const [k, tag] of [['created', 'dcterms:created'], ['modified', 'dcterms:modified'], ['creator', 'dc:creator'], ['lastModifiedBy', 'cp:lastModifiedBy']]) {
        const m = new RegExp(`<${tag}[^>]*>([^<]*)</${tag}>`).exec(core);
        if (m && m[1]) meta[k] = decodeXml(m[1]);
      }
    }
    return { zip, sheets, sheetOrder, sst, meta, _cells: {}, _xml: {} };
  }

  function textOfSi(si) {
    // รวมข้อความจาก <t> และ <r><t> แต่ข้าม <rPh> (phonetic)
    let s = '';
    for (const ch of si.childNodes) {
      if (ch.localName === 't') s += ch.textContent;
      else if (ch.localName === 'r') {
        for (const t of ch.childNodes) if (t.localName === 't') s += t.textContent;
      }
    }
    return s;
  }

  async function sheetXml(wb, name) {
    if (wb._xml[name] !== undefined) return wb._xml[name];
    const path = wb.sheets[name];
    if (!path) throw new Error('ไม่พบ sheet: ' + name);
    wb._xml[name] = await wb.zip.file(path).async('string');
    return wb._xml[name];
  }

  /** คืน Map ref → {v, f, ftext} ของ sheet (v = ค่าที่ cache ไว้, f = มีสูตรหรือไม่, ftext = ข้อความสูตร) */
  async function cells(wb, name) {
    if (wb._cells[name]) return wb._cells[name];
    const doc = new DOMParser().parseFromString(await sheetXml(wb, name), 'application/xml');
    const map = new Map();
    const sharedMaster = {}, sharedUsers = [];
    for (const c of doc.getElementsByTagNameNS(NS_MAIN, 'c')) {
      const ref = c.getAttribute('r');
      const t = c.getAttribute('t');
      let v = null, f = false, ftext = '';
      for (const ch of c.childNodes) {
        if (ch.localName === 'f') {
          f = true;
          ftext = ch.textContent;
          // shared formula: แถวแรกเก็บข้อความสูตร แถวอื่นอ้างด้วย si
          if (ch.getAttribute('t') === 'shared') {
            const si = ch.getAttribute('si');
            if (ftext) sharedMaster[si] = ftext; else sharedUsers.push([ref, si]);
          }
        }
        else if (ch.localName === 'v') {
          const raw = ch.textContent;
          if (t === 's') v = wb.sst[parseInt(raw, 10)];
          else if (t === 'str' || t === 'e') v = raw;
          else if (t === 'b') v = raw === '1';
          else v = raw === '' ? null : Number(raw);
        } else if (ch.localName === 'is') v = textOfSi(ch);
      }
      map.set(ref, { v, f, ftext });
    }
    for (const [ref, si] of sharedUsers) { const c = map.get(ref); if (c && sharedMaster[si]) c.ftext = sharedMaster[si]; }
    wb._cells[name] = map;
    return map;
  }

  async function value(wb, sheet, ref) {
    const m = await cells(wb, sheet);
    const c = m.get(ref);
    return c ? c.v : null;
  }

  /** ดึงรายการ data validation แบบ list ทั้งแบบปกติและแบบ x14 extension */
  async function listValidations(wb, sheet) {
    const xml = await sheetXml(wb, sheet);
    const out = [];
    const reStd = /<dataValidation\b([^>]*)>([\s\S]*?)<\/dataValidation>/g;
    let m;
    while ((m = reStd.exec(xml))) {
      if (!/type="list"/.test(m[1])) continue;
      const sq = /sqref="([^"]+)"/.exec(m[1]);
      const f1 = /<formula1>([\s\S]*?)<\/formula1>/.exec(m[2]);
      if (sq && f1) out.push({ sqref: sq[1].split(/\s+/), formula: decodeXml(f1[1]) });
    }
    const reX = /<x14:dataValidation\b([^>]*)>([\s\S]*?)<\/x14:dataValidation>/g;
    while ((m = reX.exec(xml))) {
      if (!/type="list"/.test(m[1])) continue;
      const f = /<xm:f>([\s\S]*?)<\/xm:f>/.exec(m[2]);
      const sq = /<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(m[2]);
      if (f && sq) out.push({ sqref: sq[1].trim().split(/\s+/), formula: decodeXml(f[1]) });
    }
    return out;
  }
  function decodeXml(s) {
    return s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  }

  /** แปลงสูตรของ list validation เป็นรายการตัวเลือก */
  async function resolveList(wb, formula) {
    formula = formula.trim();
    if (/^".*"$/.test(formula)) return formula.slice(1, -1).split(',').map(s => s.trim()).filter(Boolean);
    // ANCHORARRAY('Company List'!$C$2) = คอลัมน์ C เป็น dynamic FILTER ของคอลัมน์ A
    // จึงอ่านรายชื่อเต็มจากคอลัมน์ A ของ sheet เดียวกันแทน
    let m = /ANCHORARRAY\('?([^'!]+)'?!\$?[A-Z]+\$?(\d+)\)/.exec(formula);
    if (m) {
      const map = await cells(wb, m[1]);
      const start = parseInt(m[2], 10);
      const rowsA = [];
      for (const [ref, c] of map) {
        const mm = /^A(\d+)$/.exec(ref);
        if (mm && +mm[1] >= start && c.v !== null && String(c.v).trim() !== '') rowsA.push([+mm[1], String(c.v).trim()]);
      }
      rowsA.sort((a, b) => a[0] - b[0]);
      return [...new Set(rowsA.map(x => x[1]))];
    }
    m = /^'?([^'!]+)'?!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/.exec(formula);
    if (m) {
      const map = await cells(wb, m[1]);
      const vals = [];
      for (const ref of expandRange(m[2].replace(/\$/g, ''))) {
        const c = map.get(ref);
        if (c && c.v !== null && String(c.v).trim() !== '') vals.push(String(c.v));
      }
      return vals;
    }
    return [];
  }

  /** หา list ตัวเลือกของเซลล์จาก validations ที่ดึงไว้ */
  function findValidation(validations, ref) {
    for (const v of validations) for (const r of v.sqref) if (inRange(ref, r)) return v;
    return null;
  }

  // ---------- write ----------
  function cellXml(ref, styleAttr, val) {
    if (val === null || val === undefined || val === '') return `<c r="${ref}"${styleAttr}/>`;
    if (typeof val === 'number' && isFinite(val)) return `<c r="${ref}"${styleAttr}><v>${val}</v></c>`;
    return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${escXml(val)}</t></is></c>`;
  }

  /**
   * แก้ค่าเซลล์ใน XML ของ sheet
   * @param {string} xml
   * @param {Map<string, string|number|null>} updates ref → ค่า (null/'' = ล้างค่า)
   * @returns {{xml: string, skipped: string[]}} skipped = เซลล์ที่มีสูตรอยู่ (ไม่แตะเพื่อความปลอดภัย)
   */
  function setCells(xml, updates) {
    const skipped = [];
    const sdOpen = /<sheetData\s*\/>|<sheetData\b[^>]*>/.exec(xml);
    if (!sdOpen) throw new Error('sheetData not found');
    let head, inner, tail;
    if (sdOpen[0].endsWith('/>')) {
      head = xml.slice(0, sdOpen.index) + '<sheetData>';
      inner = '';
      tail = '</sheetData>' + xml.slice(sdOpen.index + sdOpen[0].length);
    } else {
      const start = sdOpen.index + sdOpen[0].length;
      const end = xml.indexOf('</sheetData>', start);
      head = xml.slice(0, start); inner = xml.slice(start, end); tail = xml.slice(end);
    }

    // แยกเป็นแถว
    const rows = [];
    const reRow = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g;
    let m;
    while ((m = reRow.exec(inner))) {
      const r = parseInt(/\br="(\d+)"/.exec(m[1])[1], 10);
      rows.push({ r, attrs: m[1], body: m[3] || '', dirty: false });
    }

    // จัดกลุ่ม update ตามแถว
    const byRow = new Map();
    for (const [ref, val] of updates) {
      const { row } = splitRef(ref);
      if (!byRow.has(row)) byRow.set(row, []);
      byRow.get(row).push([ref, val]);
    }

    for (const [rowNum, list] of byRow) {
      let row = rows.find(x => x.r === rowNum);
      if (!row) {
        row = { r: rowNum, attrs: ` r="${rowNum}"`, body: '', dirty: true };
        const idx = rows.findIndex(x => x.r > rowNum);
        if (idx === -1) rows.push(row); else rows.splice(idx, 0, row);
      }
      // แยกเซลล์ในแถว
      const cellsArr = [];
      const reC = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
      let cm;
      while ((cm = reC.exec(row.body))) {
        const ref = /\br="([A-Z]+\d+)"/.exec(cm[1])[1];
        cellsArr.push({ ref, c: splitRef(ref).c, xml: cm[0], attrs: cm[1], content: cm[3] || '' });
      }
      let inserted = false;
      for (const [ref, val] of list) {
        const existing = cellsArr.find(x => x.ref === ref);
        if (existing) {
          if (/<f[\s>\/]/.test(existing.content)) { skipped.push(ref); continue; }
          const s = /\bs="(\d+)"/.exec(existing.attrs);
          existing.xml = cellXml(ref, s ? ` s="${s[1]}"` : '', val);
        } else {
          if (val === null || val === undefined || val === '') continue; // ไม่มีเซลล์ และไม่มีค่า = ไม่ต้องทำอะไร
          const c = splitRef(ref).c;
          const node = { ref, c, xml: cellXml(ref, '', val) };
          const idx = cellsArr.findIndex(x => x.c > c);
          if (idx === -1) cellsArr.push(node); else cellsArr.splice(idx, 0, node);
          inserted = true;
        }
        row.dirty = true;
      }
      if (row.dirty) {
        row.body = cellsArr.map(x => x.xml).join('');
        // spans เป็นแค่ hint — ลบทิ้งเมื่อมีการเพิ่มเซลล์ใหม่เพื่อไม่ให้ค่าขัดกัน
        if (inserted) row.attrs = row.attrs.replace(/\s+spans="[^"]*"/, '');
      }
    }

    const newInner = rows.map(x =>
      x.body ? `<row${x.attrs}>${x.body}</row>` : `<row${x.attrs.replace(/\s*\/$/, '')}/>`).join('');
    return { xml: head + newInner + tail, skipped };
  }

  function forceFullCalc(workbookXml) {
    const m = /<calcPr\b[^>]*?\/?>/.exec(workbookXml);
    if (m) {
      let tag = m[0].replace(/\s+fullCalcOnLoad="[^"]*"/, '');
      tag = tag.replace(/^<calcPr/, '<calcPr fullCalcOnLoad="1"');
      return workbookXml.slice(0, m.index) + tag + workbookXml.slice(m.index + m[0].length);
    }
    const anchor = workbookXml.includes('</definedNames>') ? '</definedNames>' : '</sheets>';
    return workbookXml.replace(anchor, anchor + '<calcPr fullCalcOnLoad="1"/>');
  }

  /**
   * สร้างไฟล์ใหม่จาก workbook ต้นฉบับ + ค่าที่จะเขียน
   * @param wb ผลจาก open()
   * @param {Object<string, Map<string, any>>} updatesBySheet
   * @returns {Promise<{blob: Blob, skipped: string[]}>}
   */
  async function build(wb, updatesBySheet) {
    const out = await JSZip.loadAsync(await wb.zip.generateAsync({ type: 'arraybuffer' }), { createFolders: false });
    const skipped = [];
    for (const [sheet, updates] of Object.entries(updatesBySheet)) {
      if (!updates.size) continue;
      const path = wb.sheets[sheet];
      const xml = await out.file(path).async('string');
      const res = setCells(xml, updates);
      res.skipped.forEach(r => skipped.push(`${sheet}!${r}`));
      out.file(path, res.xml, { createFolders: false });
    }
    const wbXml = await out.file('xl/workbook.xml').async('string');
    out.file('xl/workbook.xml', forceFullCalc(wbXml), { createFolders: false });
    const blob = await out.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    return { blob, skipped };
  }

  return {
    open, cells, value, listValidations, resolveList, findValidation, build,
    setCells, forceFullCalc, serialToISO, isoToSerial, splitRef, expandRange, colToNum, numToCol,
  };
})();

if (typeof module !== 'undefined') module.exports = XL;
