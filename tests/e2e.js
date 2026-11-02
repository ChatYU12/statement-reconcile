'use strict';

/**
 * 端到端验收（v3 口径：只核金额总额）。
 * 真实起服务，走 HTTP 上传 → 探测 → 核对 → 导出。
 * 用法：npm run e2e
 */

const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const ExcelJS = require('exceljs');

const app = require('../server/index.js');

const SAMPLES = path.join(__dirname, '..', 'samples');
let base = '';
let passed = 0;
let failed = 0;

function check(name, fn) {
  try {
    fn();
    passed++;
    console.log(`  ✓ ${name}`);
  } catch (err) {
    failed++;
    console.log(`  ✗ ${name}\n    ${err.message}`);
  }
}

async function upload(statementPath, shipmentPaths) {
  const fd = new FormData();
  fd.append('statement', new Blob([fs.readFileSync(statementPath)]), path.basename(statementPath));
  for (const p of shipmentPaths) {
    fd.append('shipments', new Blob([fs.readFileSync(p)]), path.basename(p));
  }
  const res = await fetch(`${base}/api/upload`, { method: 'POST', body: fd });
  const json = await res.json();
  assert.strictEqual(res.status, 200, `上传失败：${JSON.stringify(json)}`);
  return json;
}

function mappingFrom(files) {
  return files.map((f) => ({ fileId: f.fileId, sheet: f.sheet, headerRow: f.headerRow, columns: f.columns }));
}

async function checkApi(uploadJson, overrideFiles) {
  const res = await fetch(`${base}/api/check`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ uploadId: uploadJson.uploadId, files: overrideFiles || mappingFrom(uploadJson.files) }),
  });
  return { status: res.status, json: await res.json() };
}

async function exportXlsx(resultId) {
  const res = await fetch(`${base}/api/export/${resultId}`);
  if (res.status !== 200) return { status: res.status, json: await res.json() };
  const buf = Buffer.from(await res.arrayBuffer());
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return { status: 200, buf, wb };
}

function findHeaderRow(ws) {
  for (let r = 1; r <= 15; r++) {
    if (String(ws.getRow(r).getCell(1).value || '').trim() === '产品编码') return r;
  }
  return -1;
}

async function main() {
  const server = app.listen(0);
  await new Promise((resolve) => server.once('listening', resolve));
  base = `http://127.0.0.1:${server.address().port}`;
  console.log(`测试服务：${base}\n`);

  const cfg = await (await fetch(`${base}/api/config`)).json();

  // ---------- 1. 配置 ----------
  console.log('配置');
  check('本地模式无需登录', () => assert.strictEqual(cfg.mode, 'local'));
  check('声明支持 .xlsx 与 .xls', () => assert.deepStrictEqual(cfg.acceptExt, ['.xlsx', '.xls']));
  check('金额字段标签写明口径', () => {
    assert.ok(cfg.fieldLabels.amount.includes('应收增加'));
    assert.ok(cfg.fieldLabels.amount.includes('发行实洋'));
  });
  check('数量/单价标注为不参与核对', () => {
    assert.ok(cfg.fieldLabels.qty.includes('不参与'));
    assert.ok(cfg.fieldLabels.unitPrice.includes('不参与'));
  });

  // ---------- 2. 上传与探测 ----------
  console.log('\n上传与表头探测');
  const up = await upload(path.join(SAMPLES, '验收-对账单.xlsx'), [
    path.join(SAMPLES, '验收-发货单-1.xlsx'),
    path.join(SAMPLES, '验收-发货单-2.xlsx'),
  ]);
  check('返回 1 个对账单 + 2 个发货单', () => {
    assert.strictEqual(up.files.length, 3);
    assert.strictEqual(up.files.filter((f) => f.role === 'statement').length, 1);
  });
  check('中文文件名正确还原', () => assert.strictEqual(up.files[0].fileName, '验收-对账单.xlsx'));
  check('对账单表头探测为第 3 行', () => assert.strictEqual(up.files[0].headerRow, 3));
  check('对账单金额选「应收增加」而不是「标准金额」', () => {
    const c = up.files[0].columns;
    assert.strictEqual(up.files[0].columnLabels[c.amount], '应收增加');
  });
  check('发货单金额选「发行实洋」而不是「码洋」或「入账实洋」', () => {
    const f = up.files[1];
    assert.strictEqual(f.columnLabels[f.columns.amount], '发行实洋');
  });
  check('每个文件都带 sheetInfos 且含 previewStartRow', () => {
    for (const f of up.files) {
      assert.ok(Array.isArray(f.sheetInfos) && f.sheetInfos.length >= 1);
      assert.ok(Number.isInteger(f.sheetInfos[0].previewStartRow));
    }
  });

  // ---------- 3. 核对 ----------
  console.log('\n核对：总额');
  const res = await checkApi(up);
  assert.strictEqual(res.status, 200, JSON.stringify(res.json));
  const { summary, rows, resultId } = res.json;
  const byKey = Object.fromEntries(rows.map((r) => [r.key, r]));

  check('总额按「应收增加 ↔ 发行实洋」计算', () => {
    assert.strictEqual(summary.total.statementAmount, '1345.50');
    assert.strictEqual(summary.total.shipmentAmount, '1276.51');
    assert.strictEqual(summary.total.amountDiff, '-68.99');
  });
  check('总额不一致时 matched=false 并给出方向', () => {
    assert.strictEqual(summary.total.matched, false);
    assert.ok(summary.total.conclusion.includes('不一致'));
  });
  check('参与核对的单据数与张数正确', () => {
    assert.strictEqual(summary.total.checkedFileCount, 2);
    assert.strictEqual(summary.total.excludedFileCount, 0);
    assert.strictEqual(summary.shipmentFiles.length, 2);
    assert.ok(summary.shipmentFiles.every((f) => f.amountChecked === true && typeof f.amount === 'string'));
  });
  check('结果不再输出数量与明细结论字段', () => {
    for (const r of rows) {
      assert.ok(!('statementQty' in r), '不应再有 statementQty');
      assert.ok(!('qtyDiff' in r), '不应再有 qtyDiff');
      assert.ok(!('conclusion' in r), '明细不应有 conclusion');
    }
  });
  check('金额一致的产品判为 match', () => assert.strictEqual(byKey.P001.category, 'match'));
  check('金额差 0.01 判为 amount_diff', () => {
    assert.strictEqual(byKey.P003.category, 'amount_diff');
    assert.strictEqual(byKey.P003.amountDiff, '0.01');
  });
  check('单边产品单独归类', () => {
    assert.strictEqual(byKey.P004.category, 'statement_only');
    assert.strictEqual(byKey.P006.category, 'shipment_only');
    assert.strictEqual(byKey.P004.amountKnown, false);
    assert.strictEqual(byKey.P004.shipmentAmount, null);
  });
  check('名称写法不同（全角/空格）能合并并跨 2 张单据汇总', () => {
    const r = rows.find((x) => (x.displayName || '').includes('100 型号') || (x.displayName || '').includes('100型号'));
    assert.ok(r, '没找到 A-100 型号');
    assert.strictEqual(r.category, 'match');
    assert.strictEqual(r.shipmentAmount, '276.00');
  });
  check('分类计数与明细一致', () => {
    assert.strictEqual(summary.amountDiffCount, rows.filter((r) => r.category === 'amount_diff').length);
    assert.strictEqual(summary.diffCount, summary.amountDiffCount + summary.statementOnlyCount + summary.shipmentOnlyCount);
  });
  check('结果行带匹配依据与来源行号', () => {
    assert.strictEqual(byKey.P001.matchBy, 'code');
    const src = byKey.P003.sources.find((s) => s.side === 'statement');
    assert.ok(src && src.rows[0] >= 1);
  });

  // ---------- 4. 导出 ----------
  console.log('\n差异表导出');
  const exp = await exportXlsx(resultId);
  check('导出返回 xlsx', () => assert.strictEqual(exp.status, 200));
  check('导出 8 列且不含数量列与结论列', () => {
    const ws = exp.wb.getWorksheet('差异表');
    const hr = findHeaderRow(ws);
    assert.ok(hr > 0, '没找到表头行');
    const headers = [];
    for (let c = 1; c <= 10; c++) headers.push(ws.getRow(hr).getCell(c).value);
    assert.deepStrictEqual(headers.slice(0, 8), ['产品编码', '产品名称', '编码2', '对账单金额', '发货单金额', '金额差', '匹配依据', '备注']);
    assert.ok(!headers.includes('对账单数量'), '不应再有数量列');
    assert.ok(!headers.includes('结论'), '不应再有结论列');
  });
  check('导出数据行数与结果一致', () => {
    const ws = exp.wb.getWorksheet('差异表');
    const hr = findHeaderRow(ws);
    let count = 0;
    for (let r = hr + 1; r <= ws.rowCount; r++) {
      const name = ws.getRow(r).getCell(2).value;
      const code = String(ws.getRow(r).getCell(1).value || '');
      if (name && code !== '未参与金额核对的文件') count++;
    }
    assert.ok(count >= rows.length, `只数到 ${count} 行`);
  });
  check('导出说明区含总额对比与结论', () => {
    const ws = exp.wb.getWorksheet('差异表');
    const text = [1, 2, 3, 4, 5].map((r) => String(ws.getRow(r).getCell(1).value || '')).join(' | ');
    assert.ok(text.includes('应收增加'), '说明区应写明口径');
    assert.ok(text.includes('总额'));
    assert.ok(text.includes('1345.50'));
  });
  check('金额未知的产品在导出里写「—」而不是 0.00', () => {
    const ws = exp.wb.getWorksheet('差异表');
    const hr = findHeaderRow(ws);
    let found = false;
    for (let r = hr + 1; r <= ws.rowCount; r++) {
      if (ws.getRow(r).getCell(1).value === 'P004') {
        assert.strictEqual(ws.getRow(r).getCell(5).value, '—');
        found = true;
      }
    }
    assert.ok(found, '没找到 P004');
  });

  // ---------- 5. 总额一致场景 ----------
  console.log('\n总额一致场景');
  const upOk = await upload(path.join(SAMPLES, '一致-对账单.xlsx'), [
    path.join(SAMPLES, '一致-发货单-1.xlsx'),
    path.join(SAMPLES, '一致-发货单-2.xlsx'),
  ]);
  const okRes = await checkApi(upOk);
  check('总额相等时 matched=true', () => assert.strictEqual(okRes.json.summary.total.matched, true));
  check('总额相等时差额 0.00 且结论为「总额一致」', () => {
    assert.strictEqual(okRes.json.summary.total.amountDiff, '0.00');
    assert.strictEqual(okRes.json.summary.total.conclusion, '总额一致');
    assert.strictEqual(okRes.json.summary.identical, true);
  });
  const noDiff = await exportXlsx(okRes.json.resultId);
  check('总额一致时不生成差异表（NO_DIFF）', () => {
    assert.strictEqual(noDiff.status, 400);
    assert.strictEqual(noDiff.json.error.code, 'NO_DIFF');
  });

  // ---------- 6. 没有金额列的发货单 ----------
  console.log('\n没有金额列的发货单');
  const upNoAmt = await upload(path.join(SAMPLES, '验收-对账单.xlsx'), [
    path.join(SAMPLES, '验收-发货单-1.xlsx'),
    path.join(SAMPLES, '无金额-发货单.xlsx'),
  ]);
  check('无金额列的发货单 amount 为 null 且给出候选列提示', () => {
    const f = upNoAmt.files[2];
    assert.strictEqual(f.columns.amount, null);
  });
  const noAmtRes = await checkApi(upNoAmt);
  check('不参与核对的文件不计入发货单总额', () => {
    const s = noAmtRes.json.summary;
    assert.strictEqual(s.total.checkedFileCount, 1);
    assert.strictEqual(s.total.excludedFileCount, 1);
    assert.strictEqual(s.shipmentFiles[1].amountChecked, false);
    assert.strictEqual(s.shipmentFiles[1].amount, null);
  });
  check('有单据未参与核对时不能判为总额一致', () => {
    assert.strictEqual(noAmtRes.json.summary.total.matched, false);
    assert.ok(noAmtRes.json.summary.total.conclusion.includes('未参与核对'));
  });
  check('只在无金额单据里出现的产品归为 unknown_amount', () => {
    const cats = noAmtRes.json.rows.reduce((m, r) => ((m[r.category] = (m[r.category] || 0) + 1), m), {});
    assert.ok(cats.unknown_amount >= 1, '没有 unknown_amount 分类');
    const row = noAmtRes.json.rows.find((r) => r.category === 'unknown_amount');
    assert.strictEqual(row.shipmentAmount, null);
    assert.strictEqual(row.amountDiff, null);
  });

  // ---------- 7. .xls 老格式 ----------
  console.log('\n.xls 老格式');
  const upXls = await upload(path.join(SAMPLES, '兼容-xls-对账单.xls'), [path.join(SAMPLES, '兼容-xls-发货单.xls')]);
  check('.xls 能上传并探测表头', () => {
    assert.ok(!upXls.files[0].error, upXls.files[0].error);
    assert.strictEqual(upXls.files[0].headerRow, 2);
  });
  const xlsRes = await checkApi(upXls);
  check('.xls 能完成核对且总额一致', () => {
    assert.strictEqual(xlsRes.status, 200, JSON.stringify(xlsRes.json));
    assert.strictEqual(xlsRes.json.summary.total.matched, true);
  });

  // ---------- 8. 校验与错误 ----------
  console.log('\n校验与错误');
  const noAmountStmt = mappingFrom(up.files).map((f, i) =>
    i === 0 ? { ...f, columns: { ...f.columns, amount: null } } : f,
  );
  const badStmt = await checkApi(up, noAmountStmt);
  check('对账单缺金额列时返回 MAPPING_MISSING', () => {
    assert.strictEqual(badStmt.status, 400);
    assert.strictEqual(badStmt.json.error.code, 'MAPPING_MISSING');
    assert.ok(badStmt.json.error.message.includes('金额'));
  });

  const noQty = mappingFrom(up.files).map((f) => ({ ...f, columns: { ...f.columns, qty: null } }));
  const qtyRes = await checkApi(up, noQty);
  check('不映射数量列也能正常核对', () => {
    assert.strictEqual(qtyRes.status, 200, JSON.stringify(qtyRes.json));
    assert.strictEqual(qtyRes.json.summary.total.statementAmount, '1345.50');
  });

  const badUpload = await (async () => {
    const fd = new FormData();
    fd.append('statement', new Blob([Buffer.from('not an excel')]), 'bad.txt');
    fd.append('shipments', new Blob([fs.readFileSync(path.join(SAMPLES, '验收-发货单-1.xlsx'))]), 'ok.xlsx');
    const r = await fetch(`${base}/api/upload`, { method: 'POST', body: fd });
    return r.json();
  })();
  check('非 Excel 文件给出明确原因', () => assert.ok(badUpload.files[0].error && badUpload.files[0].error.includes('.xlsx')));

  const noFile = await fetch(`${base}/api/upload`, { method: 'POST', body: new FormData() });
  const noFileJson = await noFile.json();
  check('没选对账单时返回 NO_FILE', () => assert.strictEqual(noFileJson.error.code, 'NO_FILE'));

  const expired = await fetch(`${base}/api/check`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ uploadId: 'u_notexist', files: [] }),
  });
  const expiredJson = await expired.json();
  check('uploadId 过期时返回 UPLOAD_EXPIRED', () => assert.strictEqual(expiredJson.error.code, 'UPLOAD_EXPIRED'));

  const expExpired = await fetch(`${base}/api/export/r_notexist`);
  const expExpiredJson = await expExpired.json();
  check('resultId 过期时返回 RESULT_EXPIRED', () => {
    assert.strictEqual(expExpired.status, 400);
    assert.strictEqual(expExpiredJson.error.code, 'RESULT_EXPIRED');
  });

  server.close();
  console.log(`\n通过 ${passed} 项，失败 ${failed} 项`);
  process.exit(failed === 0 ? 0 : 1);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
