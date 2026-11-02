'use strict';

/**
 * Excel 读取（.xls / .xlsx 统一走 SheetJS）
 *
 * 为什么用 SheetJS：真实样例里 8/13 是 .xls，exceljs 只支持 .xlsx。
 * 读取时优先取单元格的「显示文本」(cell.w)，因为业务上人看到的就是这个值；
 * 金额的原始浮点尾巴（1125.94199999999）也因此在展示层被格式化掉。
 */

const XLSX = require('xlsx');

const MAX_ROWS = 20000;
const MAX_COLS = 80;

function cellText(cell) {
  if (!cell) return '';
  if (cell.w !== undefined && cell.w !== null && cell.w !== '') return String(cell.w);
  if (cell.v === undefined || cell.v === null) return '';
  if (cell.t === 'd' && cell.v instanceof Date) {
    try {
      return XLSX.SSF.format('yyyy-mm-dd', cell.v);
    } catch {
      return cell.v.toISOString().slice(0, 10);
    }
  }
  return String(cell.v);
}

/**
 * 读取工作簿 → 每个 sheet 的文本网格（0-based 数组，行列号与 Excel 绝对坐标一致）
 * 返回 { sheets: [{ name, rows, rowCount, colCount, empty }], truncatedRows }
 */
function loadGrid(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: true, cellNF: true, cellStyles: false });
  let truncatedRows = 0;

  const sheets = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const ref = ws && ws['!ref'];
    if (!ref) return { name, rows: [], rowCount: 0, colCount: 0, empty: true };

    const range = XLSX.utils.decode_range(ref);
    const lastRow = range.e.r;
    const lastCol = Math.min(range.e.c, MAX_COLS - 1);
    const rowCount = Math.min(lastRow + 1, MAX_ROWS);
    if (lastRow + 1 > MAX_ROWS) truncatedRows += lastRow + 1 - MAX_ROWS;

    const rows = [];
    for (let r = 0; r < rowCount; r++) {
      const row = new Array(lastCol + 1).fill('');
      for (let c = 0; c <= lastCol; c++) {
        row[c] = cellText(ws[XLSX.utils.encode_cell({ r, c })]);
      }
      rows.push(row);
    }

    const empty = !rows.some((row) => row.some((t) => String(t).trim() !== ''));
    return { name, rows, rowCount, colCount: lastCol + 1, empty };
  });

  return { sheets, truncatedRows };
}

module.exports = { loadGrid, cellText, MAX_ROWS, MAX_COLS };
