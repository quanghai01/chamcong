const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

const FULL_HOURS = 7;
const INVALID_HOUR = 18;

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

async function processAttendance(inputBuffer) {
  const wb = XLSX.read(inputBuffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  const empData = {};
  const empOrder = [];
  const monthCount = {};

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
      empData[maNV] = { name, dept, title, days: {} };
      empOrder.push(maNV);
    }

    let v1 = String(row[6] || '').trim();
    let r1 = String(row[7] || '').trim();
    let v2 = String(row[8] || '').trim();
    let r2 = String(row[9] || '').trim();

    const v1h = parseHour(v1);
    if (v1h !== null && v1h >= 12 && !r1) {
      r1 = v1;
      v1 = '';
    }

    if (parseHour(v1) >= INVALID_HOUR) v1 = '';
    if (parseHour(v2) >= INVALID_HOUR) v2 = '';
    const totalH = parseFloat(row[10]) || 0;

    const hasVao = !!(v1 || v2);
    const hasRa  = !!(r1 || r2);

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

  const topMonthKey = Object.entries(monthCount).sort((a, b) => b[1] - a[1])[0];
  if (!topMonthKey) throw new Error("Không tìm thấy dữ liệu chấm công hợp lệ trong file");
  
  const [targetYear, targetMonth] = topMonthKey[0].split('-').map(Number);
  const daysInMonth = new Date(targetYear, targetMonth, 0).getDate();

  const wbOut = new ExcelJS.Workbook();
  const sh = wbOut.addWorksheet(`BCC T${targetMonth}`, {
    views: [{ state: 'frozen', xSplit: 4, ySplit: 5 }]
  });

  const C = { hBg: 'FF2563EB', hFg: 'FFFFFFFF', sunBg: 'FFFFF3CD', sunFg: 'FFB45309', plusBg: 'FFD1FAE5', minBg: 'FFFEE2E2', dFg: 'FF1E293B', white: 'FFFFFFFF' };
  const bdr = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };

  const sc = (r, c, val, o = {}) => {
    const cell = sh.getCell(r, c);
    cell.value = val || '';
    cell.font = { bold: !!o.bold, size: o.size || 9, name: 'Arial' };
    if (o.bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: o.bg } };
    cell.alignment = { horizontal: o.align || 'center', vertical: 'middle', wrapText: !!o.wrap };
    cell.border = bdr;
  };

  sh.mergeCells(1, 1, 1, 5 + daysInMonth + 1);
  sc(1, 1, `BẢNG CHẤM CÔNG THÁNG ${targetMonth}/${targetYear}`, { bold: true, size: 12 });

  sc(3, 1, 'STT', { bg: C.hBg, fg: C.hFg, bold: true });
  sc(3, 2, 'Họ và Tên', { bg: C.hBg, fg: C.hFg, bold: true });
  sc(3, 3, 'Phòng ban', { bg: C.hBg, fg: C.hFg, bold: true });
  sc(3, 4, 'Chức vụ', { bg: C.hBg, fg: C.hFg, bold: true });

  for (let d = 1; d <= daysInMonth; d++) {
    const wd = ['CN','T2','T3','T4','T5','T6','T7'][new Date(targetYear, targetMonth - 1, d).getDay()];
    const isSun = wd === 'CN';
    sc(3, 4 + d, wd, { bg: isSun ? C.sunBg : C.hBg, fg: isSun ? C.sunFg : C.hFg, bold: true });
    sc(4, 4 + d, d, { bg: isSun ? C.sunBg : C.hBg, fg: isSun ? C.sunFg : C.hFg, bold: true });
  }
  sc(3, 4 + daysInMonth + 1, 'Công', { bg: C.hBg, fg: C.hFg, bold: true });
  sc(3, 4 + daysInMonth + 2, 'Ghi chú', { bg: C.hBg, fg: C.hFg, bold: true });

  [1, 2, 3, 4, 4 + daysInMonth + 1, 4 + daysInMonth + 2].forEach(c => {
    try { sh.mergeCells(3, c, 4, c); } catch (e) {}
  });

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

  sh.getColumn(1).width = 5;
  sh.getColumn(2).width = 25;
  sh.getColumn(3).width = 15;
  sh.getColumn(4).width = 15;
  for (let d = 1; d <= daysInMonth; d++) sh.getColumn(4 + d).width = 3.5;
  sh.getColumn(4 + daysInMonth + 1).width = 8;
  sh.getColumn(4 + daysInMonth + 2).width = 40;

  return await wbOut.xlsx.writeBuffer();
}

module.exports = { processAttendance };
