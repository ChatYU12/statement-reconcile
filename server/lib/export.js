'use strict';

/**
 * 差异表导出（xlsx）。用 exceljs 而不是 SheetJS，因为要写样式、列宽、冻结窗格。
 * v3 口径：只核金额总额。明细 8 列，不再有数量和结论列；
 * 金额差为负时标红，金额未知（单据无金额列）写「—」而不是 0.00。
 */

const ExcelJS = require('exceljs');

const COLUMNS = [
  { header: '产品编码', key: 'code', width: 18 },
  { header: '产品名称', key: 'displayName', width: 42 },
  { header: '编码2', key: 'code2', width: 14 },
  { header: '对账单金额', key: 'statementAmount', width: 14, numFmt: '#,##0.00' },
  { header: '发货单金额', key: 'shipmentAmount', width: 14, numFmt: '#,##0.00' },
  { header: '金额差', key: 'amountDiff', width: 12, numFmt: '#,##0.00' },
  { header: '匹配依据', key: 'matchByLabel', width: 14 },
  { header: '备注', key: 'noteText', width: 62 },
];

const MATCH_BY_LABEL = { code: '编码1', code2: '编码2', name: '名称相同', fuzzy: '名称相似（模糊）' };

const CATEGORY_LABEL = {
  amount_diff: '金额不一致',
  statement_only: '仅对账单有',
  shipment_only: '仅发货单有',
  unknown_amount: '金额未知',
  match: '一致',
};

function toNumber(str) {
  const n = Number(str);
  return Number.isFinite(n) ? n : 0;
}

async function buildDiffWorkbook({ summary, rows }) {
  const wb = new ExcelJS.Workbook();
  wb.creator = '对账单核对工具';
  wb.created = new Date();

  const ws = wb.addWorksheet('差异表', {
    pageSetup: { orientation: 'landscape', fitToPage: true },
  });

  // ---- 说明区 ----
  const checked = summary.shipmentFiles.filter((f) => f.amountChecked);
  const excluded = summary.shipmentFiles.filter((f) => !f.amountChecked);
  const checkedText = checked.length
    ? `参与金额核对（${checked.length} 张）：${checked.map((f) => `${f.fileName} ${f.amount} 元`).join('、')}`
    : '参与金额核对（0 张）：没有单据提供金额数据';
  const excludedText = excluded.length
    ? `未参与金额核对（${excluded.length} 张）：${excluded.map((f) => f.fileName).join('、')}`
    : '';

  const total = summary.total || {};
  const diffText =
    total.amountDiff === '0.00'
      ? '差额 0.00 元'
      : `差额 ${total.amountDiff} 元（发货单${String(total.amountDiff).startsWith('-') ? '少' : '多'}）`;

  const info = [
    `对账单：${summary.statementFile}　应收增加合计 ${total.statementAmount} 元`,
    checkedText,
    excludedText,
    `核对时间：${new Date().toLocaleString('zh-CN', { hour12: false })}　金额口径：应收增加 ↔ 发行实洋，按分比较不允许误差`,
    `总额：对账单 ${total.statementAmount} 元，发货单 ${total.shipmentAmount} 元，${diffText} —— ${total.conclusion}`,
  ].filter(Boolean);

  info.forEach((text, i) => {
    const row = ws.getRow(i + 1);
    row.getCell(1).value = text;
    row.getCell(1).font = { size: i === 0 ? 12 : 10, bold: i <= 1, color: { argb: 'FF333333' } };
    ws.mergeCells(i + 1, 1, i + 1, COLUMNS.length);
    row.height = 18;
  });

  // ---- 列名行 ----
  const headerRowIndex = info.length + 1;
  ws.views = [{ state: 'frozen', ySplit: headerRowIndex }];
  const headerRow = ws.getRow(headerRowIndex);
  COLUMNS.forEach((col, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = col.header;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1677FF' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    ws.getColumn(i + 1).width = col.width;
  });
  headerRow.height = 22;

  // ---- 数据 ----
  rows.forEach((r) => {
    const known = r.amountKnown !== false && r.shipmentAmount !== null && r.shipmentAmount !== undefined;
    const noteText =
      (r.note || []).join('；') +
      (r.sources && r.sources.length
        ? `｜来源：${r.sources.map((s) => `${s.fileName} 第${s.rows.join(',')}行`).join('；')}`
        : '');

    const values = [
      r.code,
      r.displayName,
      r.code2,
      toNumber(r.statementAmount),
      known ? toNumber(r.shipmentAmount) : '—',
      known ? toNumber(r.amountDiff) : '—',
      r.matchBy ? MATCH_BY_LABEL[r.matchBy] || r.matchBy : '',
      `${CATEGORY_LABEL[r.category] || r.category}｜${noteText}`,
    ];
    const row = ws.addRow(values);
    COLUMNS.forEach((col, i) => {
      const cell = row.getCell(i + 1);
      if (col.numFmt && typeof cell.value === 'number') cell.numFmt = col.numFmt;
      cell.alignment = { vertical: 'top', wrapText: col.key === 'noteText' };
      cell.font = { size: 10 };
    });

    if (!known) {
      row.getCell(5).alignment = { vertical: 'top', horizontal: 'right' };
      row.getCell(6).alignment = { vertical: 'top', horizontal: 'right' };
    }
    if (known && String(r.amountDiff).startsWith('-')) {
      row.getCell(6).font = { size: 10, color: { argb: 'FFD93026' }, bold: true };
    }
    const color =
      r.category === 'amount_diff' ? 'FFD93026' : r.category === 'match' ? 'FF1A9E55' : 'FFB26A00';
    row.getCell(8).font = { size: 10, color: { argb: color } };
  });

  // ---- 未参与核对 / 解析提示 ----
  if (excluded.length) {
    ws.addRow([]);
    const title = ws.addRow(['未参与金额核对的文件']);
    title.getCell(1).font = { bold: true, color: { argb: 'FFB26A00' } };
    for (const f of excluded) {
      const row = ws.addRow([`${f.fileName}：没有金额列，其 ${f.rows} 行仅用于产品匹配，金额未计入总额`]);
      ws.mergeCells(row.number, 1, row.number, COLUMNS.length);
      row.getCell(1).font = { size: 9, color: { argb: 'FF666666' } };
    }
  }

  if (summary.warnings && summary.warnings.length) {
    ws.addRow([]);
    const title = ws.addRow(['解析提示']);
    title.getCell(1).font = { bold: true, color: { argb: 'FFB26A00' } };
    for (const w of summary.warnings.slice(0, 30)) {
      const row = ws.addRow([w]);
      ws.mergeCells(row.number, 1, row.number, COLUMNS.length);
      row.getCell(1).font = { size: 9, color: { argb: 'FF666666' } };
    }
  }

  ws.autoFilter = { from: { row: headerRowIndex, column: 1 }, to: { row: headerRowIndex, column: COLUMNS.length } };

  const buffer = await wb.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

module.exports = { buildDiffWorkbook, COLUMNS, MATCH_BY_LABEL, CATEGORY_LABEL };
