/*
 * template-2569.js — ผังเซลล์ของ "แบบรายงานผล IT Audit รอบปี 2569 (ค.ศ. 2026)" ของ ก.ล.ต.
 *
 * ไฟล์นี้เก็บเฉพาะ "ตำแหน่งเซลล์" ที่ต้องกรอก ข้อความของแบบฟอร์ม (ชื่อหัวข้อ, control,
 * คำอธิบาย Maturity, ตัวเลือก dropdown) อ่านจากไฟล์ที่ผู้ใช้โหลดเข้ามาเองทุกครั้ง
 * ถ้า ก.ล.ต. ปรับแบบฟอร์มในปีถัดไป ให้ทำ repo ใหม่ (-2027) แล้วแก้ไฟล์นี้
 */
'use strict';

const TEMPLATE = {
  appId: 'sec-itaudit-form-2026',
  formYearLabel: 'รอบปี 2569',
  versionCell: { sheet: '1. Basic info', ref: 'K2' },
  titleCell: { sheet: '1. Basic info', ref: 'A3' },

  basic: {
    sheet: '1. Basic info',
    groups: [
      {
        title: 'ข้อมูลบริษัท',
        fields: [
          { ref: 'C5', labelRef: 'A5', type: 'datalist', hintRef: 'E5', wide: true },
          { ref: 'C7', labelRef: 'A7', type: 'taxid', hintRef: 'E7' },
          { ref: 'C9', labelRef: 'A9', type: 'text' },
          { ref: 'F9', labelRef: 'E9', type: 'text' },
          { ref: 'H9', labelRef: 'G9', type: 'email' },
        ],
      },
      {
        title: 'ผู้ดำเนินการตรวจสอบ',
        fields: [
          { ref: 'C16', labelRef: 'A16', type: 'list', hintRef: 'E16' },
          { ref: 'C18', labelRef: 'A18', type: 'text', wide: true },
        ],
      },
      {
        title: 'วันที่ดำเนินการตรวจสอบ',
        fields: [
          { ref: 'C23', label: 'ระหว่างวันที่ (วันเริ่มตรวจ)', type: 'date' },
          { ref: 'E23', label: 'ถึงวันที่ (วันสุดท้ายของการตรวจ)', type: 'date' },
        ],
      },
      {
        title: 'การรายงานผลการตรวจสอบและแผนการปรับปรุงแก้ไข',
        noteRef: 'H14',
        fields: [
          { ref: 'I16', labelRef: 'H16', type: 'date' },
          { ref: 'I18', labelRef: 'H18', type: 'date' },
        ],
        footRef: 'H19',
      },
      {
        title: 'ระดับความเสี่ยงจากการประเมิน RLA',
        fields: [{ ref: 'C26', labelRef: 'A26', type: 'list', wide: true }],
        footRef: 'F26',
      },
    ],
    auditors: {
      headerRow: 29,
      exampleRows: [30, 31],
      rows: range(32, 51),
      cols: [
        { col: 'B', type: 'text' },
        { col: 'D', type: 'list' },
        { col: 'F', type: 'text' },
        { col: 'G', type: 'text' },
        { col: 'H', type: 'text' },
        { col: 'I', type: 'text' },
        { col: 'J', type: 'text' },
        { col: 'K', type: 'email' },
        { col: 'L', type: 'text' },
      ],
    },
  },

  systems: {
    sheet: '2. Systems',
    noteRefs: ['A2', 'B3', 'B4'],
    // header = แถวชื่อหมวด, desc = แถวคำอธิบาย (ถ้ามี), rows = แถวที่กรอก, other = มีคอลัมน์ C "อื่น ๆ"
    groups: [
      { header: 6, rows: [9, 10, 11], other: true },
      { header: 13, rows: [16, 17, 18], other: true },
      { header: 20, rows: [23, 24, 25], other: true },
      { header: 27, rows: [30, 31, 32], other: true },
      { header: 34, desc: 35, rows: [38, 39, 40], other: true },
      { header: 42, desc: 43, rows: [46, 47, 48], other: true },
      { header: 50, rows: [53, 54, 55], other: true },
      { header: 57, rows: [60, 61, 62], other: true },
      { header: 64, desc: 65, rows: [68, 69, 70], other: false },
    ],
  },

  scope: {
    sheet: '3. Scope',
    noteRef: 'A2',
    headerRow: 3,
    exampleRows: [4, 5, 6, 7],
    rows: range(8, 37),
    cols: [
      { col: 'B', type: 'list' },
      { col: 'C', type: 'list' },
      { col: 'D', type: 'text' },
      { col: 'E', type: 'text' },
      { col: 'F', type: 'list' },
      { col: 'G', type: 'textarea' },
    ],
  },

  controls: {
    sheetPattern: /^(4\. )?D\d/,
    firstRow: 3,
    // คอลัมน์ในแต่ละ sheet D1–D3
    col: { id: 'A', H: 'B', M: 'C', L: 'D', S: 'E', sor: 'F', np: 'G', ref: 'H', text: 'I', result: 'J', type: 'L', mat: ['M', 'N', 'O', 'P', 'Q'] },
    idPattern: /^D[\d.]+ #\d+$/,
    rawSheet: 'RawData', // A = control ID, E = CategoryID, F = AdvControl (Y/N)
  },

  findings: {
    sheet: '5. Finding',
    headerRow: 10,
    exampleRows: [11, 12],
    rows: range(13, 112),
    noteRef: 'A8',
    // E5/E6 = ช่องกรอกจำนวนข้อระดับปานกลาง/ต่ำ (แบบฟอร์มให้กรอกเอง) → แอปนับและเขียนให้
    // E4 (ระดับสูง) เป็นสูตร COUNTIF ของแบบฟอร์ม ไม่ต้องเขียน
    countRefs: { M: 'E5', L: 'E6' },
    cols: [
      { col: 'B', key: 'B', type: 'list' },
      { col: 'C', key: 'C', type: 'ids' },
      { col: 'D', key: 'D', type: 'list' },
      { col: 'E', key: 'E', type: 'textarea' },
      { col: 'F', key: 'F', type: 'list' },
      { col: 'G', key: 'G', type: 'textarea' },
      { col: 'H', key: 'H', type: 'month' },
      { col: 'I', key: 'I', type: 'text' },
    ],
  },

  result: { sheet: '6. Result', firstRow: 6, lastRow: 36 },

  rlaToCol: { 'สูง': 'H', 'ปานกลาง': 'M', 'ต่ำ': 'L', 'ผู้ประกอบธุรกิจขนาดเล็ก': 'S' },

  // กติกาวันที่ตาม data validation ในแบบฟอร์ม
  dateRules: { min: '2026-01-01', auditEndMax: '2027-04-01' },
};

function range(a, b) { const r = []; for (let i = a; i <= b; i++) r.push(i); return r; }
