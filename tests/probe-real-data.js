'use strict';

/**
 * 开发期诊断脚本：打印真实样例的探测结果与核对结论，方便人工核对字段映射。
 * 用法：node tests/probe-real-data.js [样例目录]
 */

const path = require('path');
const fs = require('fs');
const { loadGrid } = require('../server/lib/sheet');
const { detectHeader } = require('../server/lib/detect');
const { extractRows } = require('../server/lib/extract');
const { reconcile } = require('../server/lib/engine');

const dir = process.argv[2] || '/Users/chatyu/Downloads/陇南卫校 天水卫校 定西 兰州域航';

if (!fs.existsSync(dir)) {
  console.error(`样例目录不存在：${dir}\n用法：node tests/probe-real-data.js <样例目录>`);
  process.exit(2);
}

function parse(fileName) {
  const filePath = path.join(dir, fileName);
  const grid = loadGrid(filePath);
  const sheet = grid.sheets.find((s) => !s.empty);
  if (!sheet) throw new Error(`没有非空 sheet: ${fileName}`);
  const info = detectHeader(sheet.name, sheet.rows, null);
  const data = extractRows({
    rows: sheet.rows,
    headerRow: info.headerRow,
    headerSpan: info.headerSpan,
    columns: info.columns,
  });
  return { fileName, sheet: sheet.name, info, data, skippedSheets: grid.sheets.filter((s) => s.empty).map((s) => s.name) };
}

function printParsed(p) {
  console.log(`\n=== ${p.fileName}  [sheet=${p.sheet}]`);
  console.log('  表头行', p.info.headerRow, '列映射', JSON.stringify(p.info.columns));
  console.log('  表头文字', JSON.stringify(p.info.columnLabels));
  if (p.info.warnings.length) console.log('  探测警告', p.info.warnings);
  console.log(`  数据行 ${p.data.items.length}，跳过 ${p.data.skipped.length}，金额口径 ${p.data.amountMode}`);
  if (p.data.skipped.length) console.log('  跳过明细', JSON.stringify(p.data.skipped.slice(0, 6)));
  if (p.data.warnings.length) console.log('  解析警告', p.data.warnings.slice(0, 6));
  console.log('  前 5 行:', p.data.items.slice(0, 5).map((i) => `${i.code || i.code2 || '-'}/${(i.name || '').slice(0, 12)} qty=${i.qty} amt=${i.amount}`));
}

const files = fs.readdirSync(dir).filter((f) => /\.(xlsx|xls)$/i.test(f)).sort();
const parsed = new Map();
for (const f of files) {
  try {
    const p = parse(f);
    parsed.set(f, p);
    printParsed(p);
  } catch (e) {
    console.log(`\n=== ${f}\n  解析失败: ${e.message}`);
  }
}

const stmtName = '客户应收往来流水帐_29834域航.xls';
const stmt = parsed.get(stmtName);
if (!stmt) {
  console.log('\n没找到对账单样例，跳过核对');
  process.exit(0);
}

for (const shipName of ['天水卫校康养（文智达）7.24.xls', '订书-兰州域航（天水卫校） 杨敏兰7.24.xlsx']) {
  const ship = parsed.get(shipName);
  if (!ship) continue;
  const { summary, rows, suspects } = reconcile({
    statement: { fileId: 's', fileName: stmt.fileName, ...stmt.data },
    shipments: [{ fileId: 'p', fileName: ship.fileName, ...ship.data }],
  });
  console.log(`\n######## 核对：${stmtName}  ×  ${shipName}`);
  console.log(
    `  产品 ${summary.productCount}，匹配 ${summary.matchedCount}，数值差 ${summary.valueDiffCount}，仅对账单 ${summary.statementOnlyCount}，仅发货单 ${summary.shipmentOnlyCount}`,
  );
  console.log('  匹配依据', JSON.stringify(summary.matchStats));
  for (const r of rows.slice(0, 12)) {
    console.log(
      `  [${r.category}] ${r.code || r.code2 || '-'} ${(r.displayName || '').slice(0, 20)}  对账 ${r.statementQty}/${r.statementAmount}  发货 ${r.shipmentQty}/${r.shipmentAmount}  差 ${r.qtyDiff}/${r.amountDiff}  ${r.conclusion} (${r.matchBy || '-'})`,
    );
  }
  if (suspects.length) console.log('  疑似', suspects.slice(0, 5));
}
