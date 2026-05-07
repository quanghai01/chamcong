/**
 * process_chamcong.js
 * 
 * Đọc dữ liệu từ dauvao.xls (XML Spreadsheet 2003)
 * Nếu có file dauvao.xlsx (danh sách chuẩn), tự động khớp tên chuẩn để sửa lỗi font.
 * Chỉ ghi chú các ngày "ko chấm lúc đi/về".
 */

const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// ─── CONFIG ──────────────────────────────────────────────────────────────────
const INPUT_XLS = path.join(__dirname, 'dauvao.xlsx');
// Bỏ qua SAMPLE_XLSX vì đã gộp
const OUTPUT_FILE = path.join(__dirname, 'ketqua.xlsx');

const FULL_HOURS = 7;   // >= FULL_HOURS → FULL
const INVALID_HOUR = 18;  // Vào > giờ này → coi như không check-in
// ─────────────────────────────────────────────────────────────────────────────

// ─── HELPERS ─────────────────────────────────────────────────────────────────
const parseHour = (str) => {
  if (!str) return null;
  const m = String(str).trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? parseInt(m[1]) + parseInt(m[2]) / 60 : null;
};

const parseDate = (raw) => {
  let d;
  if (raw instanceof Date) d = new Date(raw);
  else if (typeof raw === 'string' && raw.includes('T')) { d = new Date(raw); d.setDate(d.getDate() + 1); }
  else if (typeof raw === 'number') { const p = XLSX.SSF.parse_date_code(raw); d = new Date(p.y, p.m - 1, p.d); }
  else return null;
  return isNaN(d.getTime()) ? null : d;
};

// Thuật toán match tên bị lỗi font (fuzzy match)
const removeAccents = (str) => str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D");
const matchName = (garbled, cleanMap) => {
  if (!garbled) return null;
  const cleanNames = Object.keys(cleanMap);
  if (garbled.includes("?")) {
    const regexStr = garbled.replace(/\?/g, '.');
    const regex = new RegExp("^" + regexStr + "$", "i");
    for (const clean of cleanNames) {
      if (regex.test(removeAccents(clean))) return cleanMap[clean];
    }
  }
  const gAscii = garbled.replace(/[^\x00-\x7F]/g, "").toLowerCase();
  let bestMatch = null, maxScore = -1;
  for (const clean of cleanNames) {
    const cAscii = clean.replace(/[^\x00-\x7F]/g, "").toLowerCase();
    let score = 0;
    for (let i = 0; i < gAscii.length; i++) if (cAscii.includes(gAscii[i])) score++;
    if (score > maxScore && score > gAscii.length * 0.5) { maxScore = score; bestMatch = cleanMap[clean]; }
  }
  return bestMatch;
};

// Không cần đọc danh sách chuẩn nữa vì người dùng tự lưu file
let useSample = false;

// ─── 1. ĐỌC DỮ LIỆU TỪ DAUVAO.XLS ──────────────────────────────────────────
console.log('📂 Đọc file dữ liệu:', INPUT_XLS);
const wb = XLSX.readFile(INPUT_XLS);
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

const empData = {};
const empOrder = [];
const monthCount = {};

// Thu thập dữ liệu
rows.forEach((row, idx) => {
  if (idx < 4) return;
  const maNV = String(row[2] || '').trim();
  let name = String(row[3] || '').trim();
  let dept = String(row[4] || '').trim();
  let title = String(row[5] || '').trim();

  if (!maNV || isNaN(Number(maNV))) return;

  const date = parseDate(row[0]);
  if (!date) return;

  const tKey = `${date.getFullYear()}-${date.getMonth() + 1}`;
  monthCount[tKey] = (monthCount[tKey] || 0) + 1;

  if (!empData[maNV]) {
    // Nếu có file mẫu, tự động match tên để fix lỗi font
    if (useSample) {
      const cleanInfo = matchName(name, cleanEmpMap);
      if (cleanInfo) {
        name = cleanInfo.name;
        dept = cleanInfo.dept;
        title = cleanInfo.title;
      }
    }
    empData[maNV] = { name, dept, title, days: {} };
    empOrder.push(maNV);
  }

  // Xử lý giờ vào/ra
  let v1 = String(row[6] || '').trim();
  let r1 = String(row[7] || '').trim();
  let v2 = String(row[8] || '').trim();
  let r2 = String(row[9] || '').trim();

  // Đảo punch nếu chỉ có v1 mà v1 >= 12h (quên chấm đi, máy đẩy sang cột vào)
  const v1h = parseHour(v1);
  if (v1h !== null && v1h >= 12 && !r1) {
    r1 = v1;
    v1 = '';
  }

  // Quá 18h coi như không check-in
  if (parseHour(v1) >= INVALID_HOUR) v1 = '';
  if (parseHour(v2) >= INVALID_HOUR) v2 = '';
  const totalH = parseFloat(row[10]) || 0;

  const hasVao = !!(v1 || v2);
  const hasRa = !!(r1 || r2);

  let mark = '';
  let note = '';

  const day = date.getDate();
  const month = date.getMonth() + 1;

  if (!hasVao && hasRa) {
    note = `Ngày ${day}.${month} ko chấm lúc đi`;
  } else if (hasVao && !hasRa) {
    note = `Ngày ${day}.${month} ko chấm lúc về`;
  }

  if (totalH >= FULL_HOURS) {
    mark = '+';
  } else if (totalH > 0 || hasVao || hasRa) {
    mark = '-';
  }

  const existing = empData[maNV].days[day];
  if (!existing) {
    empData[maNV].days[day] = { mark, note };
  } else {
    if (mark === '+' && existing.mark === '-') existing.mark = '+';
    if (note && (!existing.note || !existing.note.includes(note))) {
      existing.note = (existing.note ? existing.note + '\n' : '') + note;
    }
  }
});

// Xác định tháng chính
const [targetYear, targetMonth] = Object.entries(monthCount)
  .sort((a, b) => b[1] - a[1])[0][0]
  .split('-').map(Number);
const daysInMonth = new Date(targetYear, targetMonth, 0).getDate();

console.log(`📅 Tháng: ${targetMonth}/${targetYear}  |  Nhân viên: ${empOrder.length}`);

// ─── 2. TẠO FILE EXCEL ───────────────────────────────────────────────────────
const wbOut = new ExcelJS.Workbook();
const sh = wbOut.addWorksheet(`BCC T${targetMonth}`, {
  views: [{ state: 'frozen', xSplit: 4, ySplit: 5 }]
});

const C = { hBg: 'FF2563EB', hFg: 'FFFFFFFF', sunBg: 'FFFFF3CD', sunFg: 'FFB45309', plusBg: 'FFD1FAE5', minBg: 'FFFEE2E2', dFg: 'FF1E293B', white: 'FFFFFFFF' };
const bdr = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

const sc = (r, c, val, o = {}) => {
  const cell = sh.getCell(r, c);
  cell.value = val || '';
  cell.font = { bold: !!o.bold, size: o.size || 9, name: 'Arial' };
  if (o.bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: o.bg } };
  cell.alignment = { horizontal: o.align || 'center', vertical: 'middle', wrapText: !!o.wrap };
  cell.border = bdr;
};

// Header
sh.mergeCells(1, 1, 1, 5 + daysInMonth + 1);
sc(1, 1, `BẢNG CHẤM CÔNG THÁNG ${targetMonth}/${targetYear}`, { bold: true, size: 12 });

sc(3, 1, 'STT', { bg: C.hBg, fg: C.hFg, bold: true });
sc(3, 2, 'Họ và Tên', { bg: C.hBg, fg: C.hFg, bold: true });
sc(3, 3, 'Phòng ban', { bg: C.hBg, fg: C.hFg, bold: true });
sc(3, 4, 'Chức vụ', { bg: C.hBg, fg: C.hFg, bold: true });

for (let d = 1; d <= daysInMonth; d++) {
  const wd = ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'][new Date(targetYear, targetMonth - 1, d).getDay()];
  const isSun = wd === 'CN';
  sc(3, 4 + d, wd, { bg: isSun ? C.sunBg : C.hBg, fg: isSun ? C.sunFg : C.hFg, bold: true });
  sc(4, 4 + d, d, { bg: isSun ? C.sunBg : C.hBg, fg: isSun ? C.sunFg : C.hFg, bold: true });
}
sc(3, 4 + daysInMonth + 1, 'Công', { bg: C.hBg, fg: C.hFg, bold: true });
sc(3, 4 + daysInMonth + 2, 'Ghi chú', { bg: C.hBg, fg: C.hFg, bold: true });

// Merge header columns
[1, 2, 3, 4, 4 + daysInMonth + 1, 4 + daysInMonth + 2].forEach(c => {
  try { sh.mergeCells(3, c, 4, c); } catch (e) { }
});

// Data
let curR = 5;
empOrder.forEach((maNV, idx) => {
  const emp = empData[maNV];
  sc(curR, 1, idx + 1);
  sc(curR, 2, emp.name, { align: 'left' });
  sc(curR, 3, emp.dept, { align: 'left' });
  sc(curR, 4, emp.title, { align: 'left' });

  let totalCong = 0;
  let notes = [];

  for (let d = 1; d <= daysInMonth; d++) {
    const isSun = new Date(targetYear, targetMonth - 1, d).getDay() === 0;
    const dayData = emp.days[d];
    const mark = dayData ? dayData.mark : '';

    if (dayData && dayData.note) notes.push(dayData.note);

    // FIX: Bỏ qua không tính công vào Chủ Nhật
    if (!isSun) {
      if (mark === '+') totalCong += 1;
      else if (mark === '-') totalCong += 0.5;
    }

    sc(curR, 4 + d, isSun ? '' : mark, {
      bg: isSun ? C.sunBg : (mark === '+' ? C.plusBg : (mark === '-' ? C.minBg : C.white))
    });
  }

  sc(curR, 4 + daysInMonth + 1, totalCong, { bold: true });
  sc(curR, 4 + daysInMonth + 2, notes.join('\n'), { align: 'left', wrap: true, size: 8 });

  if (notes.length > 0) sh.getRow(curR).height = Math.max(15, notes.length * 12);
  curR++;
});

// Columns width
sh.getColumn(1).width = 5;
sh.getColumn(2).width = 25;
sh.getColumn(3).width = 15;
sh.getColumn(4).width = 15;
for (let d = 1; d <= daysInMonth; d++) sh.getColumn(4 + d).width = 3.5;
sh.getColumn(4 + daysInMonth + 1).width = 8;
sh.getColumn(4 + daysInMonth + 2).width = 40;

wbOut.xlsx.writeFile(OUTPUT_FILE).then(() => {
  console.log(`✅ Đã xuất: ${OUTPUT_FILE}`);
}).catch(e => console.error('❌ Lỗi:', e));
