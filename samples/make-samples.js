'use strict';

/**
 * 生成验收/回归用的合成样例。真实样例（客户的对账单）不能进仓库，
 * 这里造一套覆盖 v3 口径（只核金额总额）的表：
 *
 *   验收-*        : 多个候选金额列并存（标准金额/应收增加、码洋/发行实洋），验证识别优先级；
 *                   同时含 金额差 0.01 / 仅对账单 / 仅发货单 / 名称写法不一致
 *   一致-*        : 总额完全一致 → 页面应显示「总额一致」
 *   无金额-*      : 没有金额列的发货单（真实 xls 单的形态），应不参与金额核对
 *   兼容-xls-*    : 老格式 .xls（表头不在第一行）
 *
 * 用法：npm run samples
 */

const path = require('path');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

const outDir = __dirname;

async function writeXlsx(fileName, { titleLines = [], header, rows, total = null, sheetName = 'Sheet1', moneyCols = [] }) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName);
  const width = header.length;

  let r = 1;
  for (const line of titleLines) {
    ws.getCell(r, 1).value = line;
    ws.mergeCells(r, 1, r, width);
    r++;
  }

  header.forEach((h, i) => {
    const cell = ws.getCell(r, i + 1);
    cell.value = h;
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFEFEF' } };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FF999999' } } };
    ws.getColumn(i + 1).width = i === 1 ? 30 : 14;
  });
  ws.getRow(r).height = 20;
  r++;

  for (const row of rows) {
    row.forEach((v, i) => {
      const cell = ws.getCell(r, i + 1);
      cell.value = v;
      if (typeof v === 'number') {
        cell.numFmt = moneyCols.includes(i + 1) ? '#,##0.00' : '#,##0.####';
      }
    });
    r++;
  }

  if (total) {
    total.forEach((v, i) => {
      const cell = ws.getCell(r, i + 1);
      cell.value = v;
      cell.font = { bold: true };
      if (typeof v === 'number') cell.numFmt = moneyCols.includes(i + 1) ? '#,##0.00' : '#,##0.####';
    });
    r++;
  }

  ws.getCell(r, 1).value = '采购签字：';
  ws.getCell(r + 1, 1).value = '制单：';

  const filePath = path.join(outDir, fileName);
  await wb.xlsx.writeFile(filePath);
  console.log('生成', path.basename(filePath));
  return filePath;
}

function writeXls(fileName, { titleLines = [], header, rows }) {
  const aoa = [...titleLines.map((t) => [t]), header, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  const filePath = path.join(outDir, fileName);
  XLSX.writeFile(wb, filePath, { bookType: 'biff8' });
  console.log('生成', path.basename(filePath));
  return filePath;
}

// 对账单：码洋列（标准金额）与实洋列（应收增加）并存，实洋 = 码洋 × 0.69
const STATEMENT_HEADER = ['产品编码', '产品名称', '数量', '标准金额', '发行折扣', '应收增加', '总收款'];
// 订书单：码洋 / 发行实洋 / 入账实洋 三列并存，应当选中「发行实洋」
const SHIPMENT_HEADER = ['客户编码', '产品编码', '产品名称', '数量', '码洋', '发行折扣', '发行实洋', '入账折扣', '入账实洋'];
// xls 老格式发货单：只有定价和数量，没有金额列
const NO_AMOUNT_HEADER = ['序号', 'ISBN号', '图书名称', '版别', '定价', '数量', '备注'];

function stmtRow(code, name, qty, listAmount, rate = 0.69) {
  const real = Math.round(listAmount * rate * 100) / 100;
  return [code, name, qty, listAmount, `${rate}`, real, ''];
}

function shipRow(customer, code, name, qty, listAmount, rate = 0.69) {
  const real = Math.round(listAmount * rate * 100) / 100;
  return [customer, code, name, qty, listAmount, `${rate}`, real, '0.69', real];
}

async function main() {
  // ========== 场景 1：有差异 ==========
  await writeXlsx('验收-对账单.xlsx', {
    titleLines: ['兰州市域航科技文化有限公司　应收对账单', '账期：2026-09　单位：元'],
    header: STATEMENT_HEADER,
    moneyCols: [4, 6],
    rows: [
      stmtRow('P001', '儿童保温杯 500ml', 10, 1000.0),
      stmtRow('P002', '不锈钢餐盘', 5, 250.0),
      stmtRow('P003', '折叠雨伞', 3, 100.0),
      stmtRow('P004', '仅对账单有的产品', 2, 200.0),
      stmtRow('', 'Ａ-100 型号', 4, 400.0),
    ],
    total: ['合计', '', 24, 1950.0, '', 1345.5, ''],
    sheetName: '对账单',
  });

  await writeXlsx('验收-发货单-1.xlsx', {
    titleLines: ['兰州域航订单', '发货单位：文智达　流向：天水卫校　2026-09'],
    header: SHIPMENT_HEADER,
    moneyCols: [5, 7, 9],
    rows: [
      shipRow('03436', 'P001', '儿童保温杯 500ml', 4, 400.0),
      shipRow('03436', 'P002', '不锈钢餐盘', 3, 150.0),
      shipRow('03436', 'P003', '折叠雨伞', 3, 100.01),
      shipRow('03436', '', 'A-100 型号', 2, 200.0),
    ],
    total: ['03436', '', '', 12, 850.01, '', 586.51, '', 586.51],
  });

  await writeXlsx('验收-发货单-2.xlsx', {
    titleLines: ['兰州域航订单', '发货单位：文智达　流向：天水卫校　2026-09'],
    header: SHIPMENT_HEADER,
    moneyCols: [5, 7, 9],
    rows: [
      shipRow('03436', 'P001', '儿童保温杯 500ml', 6, 600.0),
      shipRow('03436', 'P002', '不锈钢餐盘', 3, 150.0),
      shipRow('03436', '', 'A-100型号', 2, 200.0),
      shipRow('03436', 'P006', '仅发货单有的产品', 1, 50.0),
    ],
    total: ['03436', '', '', 12, 1000.0, '', 690.0, '', 690.0],
  });

  // ========== 场景 2：总额完全一致 ==========
  await writeXlsx('一致-对账单.xlsx', {
    titleLines: ['兰州市域航科技文化有限公司　应收对账单', '账期：2026-09　单位：元'],
    header: STATEMENT_HEADER,
    moneyCols: [4, 6],
    rows: [stmtRow('P001', '儿童保温杯 500ml', 10, 1000.0), stmtRow('P002', '不锈钢餐盘', 5, 250.0)],
    total: ['合计', '', 15, 1250.0, '', 862.5, ''],
    sheetName: '对账单',
  });

  await writeXlsx('一致-发货单-1.xlsx', {
    titleLines: ['兰州域航订单'],
    header: SHIPMENT_HEADER,
    moneyCols: [5, 7, 9],
    rows: [
      shipRow('03436', 'P001', '儿童保温杯 500ml', 4, 400.0),
      shipRow('03436', 'P002', '不锈钢餐盘', 2, 100.0),
    ],
  });

  await writeXlsx('一致-发货单-2.xlsx', {
    titleLines: ['兰州域航订单'],
    header: SHIPMENT_HEADER,
    moneyCols: [5, 7, 9],
    rows: [
      shipRow('03436', 'P001', '儿童保温杯 500ml', 6, 600.0),
      shipRow('03436', 'P002', '不锈钢餐盘', 3, 150.0),
    ],
  });

  // ========== 场景 3：没有金额列的发货单（对应真实 xls 单） ==========
  await writeXlsx('无金额-发货单.xlsx', {
    titleLines: ['兰州域航订单', '发货单位：文智达　流向：天水卫校'],
    header: NO_AMOUNT_HEADER,
    rows: [
      [1, 'P001', '儿童保温杯 500ml', '文智达', 100.0, 10, ''],
      [2, 'P002', '不锈钢餐盘', '文智达', 50.0, 5, ''],
      // P004 只在这张没有金额列的单里出现 → 应归为「金额未知」
      [3, 'P004', '仅对账单有的产品', '文智达', 100.0, 2, ''],
    ],
  });

  // ========== 场景 4：.xls 老格式 ==========
  writeXls('兼容-xls-对账单.xls', {
    titleLines: ['兰州市域航科技文化有限公司　应收对账单'],
    header: STATEMENT_HEADER,
    rows: [stmtRow('X001', '老格式测试产品', 2, 88.8), stmtRow('X002', '老格式第二个产品', 1, 11.11)],
  });

  writeXls('兼容-xls-发货单.xls', {
    titleLines: ['兰州域航订单'],
    header: SHIPMENT_HEADER,
    rows: [shipRow('03436', 'X001', '老格式测试产品', 2, 88.8), shipRow('03436', 'X002', '老格式第二个产品', 1, 11.11)],
  });

  console.log('\n样例已生成到', outDir);
}

if (require.main === module) {
  main().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}

module.exports = { writeXlsx, writeXls };
