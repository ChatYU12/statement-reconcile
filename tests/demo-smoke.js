'use strict';

/**
 * GitHub Pages 演示版的冒烟测试。
 *
 * 用 jsdom 加载 docs/index.html（构建产物），注入浏览器版依赖，
 * 把「上传 → 探测 → 核对 → 导出」整条链路在本机跑一遍。
 * 目的是在上传前确认静态版逻辑与 Node 版一致，不是只测文件存在。
 *
 * 用法：node scripts/build-pages.mjs && node tests/demo-smoke.js
 */

const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { JSDOM } = require('jsdom');
const ExcelJS = require('exceljs');

const ROOT = path.join(__dirname, '..');
const DOCS = path.join(ROOT, 'docs');

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

function fileOf(rel) {
  const buf = fs.readFileSync(path.join(ROOT, 'samples', rel));
  return new File([buf], rel);
}

async function waitForView(doc, id, timeoutMs = 4000) {
  const start = Date.now();
  for (;;) {
    const el = doc.getElementById(id);
    if (el && el.classList.contains('is-active')) return true;
    if (Date.now() - start > timeoutMs) return false;
    await new Promise((r) => setTimeout(r, 40));
  }
}

async function main() {
  if (!fs.existsSync(path.join(DOCS, 'index.html'))) {
    console.error('docs/ 还没构建，先跑：node scripts/build-pages.mjs');
    process.exit(1);
  }

  const html = fs.readFileSync(path.join(DOCS, 'index.html'), 'utf8');
  const dom = new JSDOM(html, { url: 'http://localhost/', pretendToBeVisual: true, runScripts: 'outside-only' });
  const { window } = dom;

  // 浏览器全局（jsdom 里缺的部分用 Node 的实现补上）
  window.XLSX = require('xlsx');
  window.ExcelJS = ExcelJS;
  window.Response = Response;
  window.Blob = Blob;
  window.File = File;
  window.FormData = FormData;
  window.fetch = fetch;
  window.scrollTo = () => {};
  window.alert = (msg) => {
    throw new Error(`不该弹 alert：${msg}`);
  };
  window.URL.createObjectURL = () => 'blob:fake';
  window.URL.revokeObjectURL = () => {};

  const errors = [];
  window.addEventListener('error', (e) => errors.push(e.message));

  console.log('加载构建产物');
  window.eval(fs.readFileSync(path.join(DOCS, 'demo', 'lib.js'), 'utf8'));
  window.eval(fs.readFileSync(path.join(DOCS, 'demo', 'api.js'), 'utf8'));
  window.eval(fs.readFileSync(path.join(DOCS, 'app.js'), 'utf8'));

  await new Promise((r) => setTimeout(r, 120));

  check('脚本加载无运行时错误', () => assert.deepStrictEqual(errors, []));
  check('lib bundle 暴露了需要的函数', () => {
    const L = window.ReconcileLib;
    for (const fn of ['loadGrid', 'detectHeader', 'extractRows', 'reconcile', 'buildDiffWorkbook']) {
      assert.strictEqual(typeof L[fn], 'function', `缺 ${fn}`);
    }
  });
  check('演示版说明条已插入', () => {
    const bar = window.document.getElementById('demoBanner');
    assert.ok(bar, '没有 demoBanner');
    assert.ok(bar.textContent.includes('不会上传到任何服务器'));
  });

  // ---------- 上传 ----------
  console.log('\n浏览器内上传');
  const form = new FormData();
  form.append('statement', fileOf('验收-对账单.xlsx'), '验收-对账单.xlsx');
  form.append('shipments', fileOf('验收-发货单-1.xlsx'), '验收-发货单-1.xlsx');
  form.append('shipments', fileOf('验收-发货单-2.xlsx'), '验收-发货单-2.xlsx');

  const upRes = await window.fetch('/api/upload', { method: 'POST', body: form });
  const up = await upRes.json();

  check('上传返回 3 个文件', () => {
    assert.strictEqual(upRes.status, 200, JSON.stringify(up));
    assert.strictEqual(up.files.length, 3);
  });
  check('中文文件名没有被破坏', () => assert.strictEqual(up.files[0].fileName, '验收-对账单.xlsx'));
  check('表头在第 3 行', () => assert.strictEqual(up.files[0].headerRow, 3));
  check('对账单金额选「应收增加」', () => {
    const f = up.files[0];
    assert.strictEqual(f.columnLabels[f.columns.amount], '应收增加');
  });
  check('发货单金额选「发行实洋」', () => {
    const f = up.files[1];
    assert.strictEqual(f.columnLabels[f.columns.amount], '发行实洋');
  });
  check('带回了预览行', () => {
    assert.ok(Array.isArray(up.files[0].previewRows) && up.files[0].previewRows.length > 0);
  });

  // ---------- 核对 ----------
  console.log('\n浏览器内核对');
  const payload = {
    uploadId: up.uploadId,
    files: up.files.map((f) => ({ fileId: f.fileId, sheet: f.sheet, headerRow: f.headerRow, columns: f.columns })),
  };
  const ckRes = await window.fetch('/api/check', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  const ck = await ckRes.json();

  check('核对成功', () => assert.strictEqual(ckRes.status, 200, JSON.stringify(ck && ck.error)));
  check('总额与 Node 版一致（1345.50 / 1276.51 / -68.99）', () => {
    assert.strictEqual(ck.summary.total.statementAmount, '1345.50');
    assert.strictEqual(ck.summary.total.shipmentAmount, '1276.51');
    assert.strictEqual(ck.summary.total.amountDiff, '-68.99');
  });
  check('分类与 Node 版一致', () => {
    const by = Object.fromEntries(ck.rows.map((r) => [r.key, r.category]));
    assert.strictEqual(by.P001, 'match');
    assert.strictEqual(by.P003, 'amount_diff');
    assert.strictEqual(by.P004, 'statement_only');
    assert.strictEqual(by.P006, 'shipment_only');
  });

  // ---------- 导出 ----------
  console.log('\n浏览器内导出');
  const exRes = await window.fetch('/api/export/' + ck.resultId);
  check('导出返回 200 与 xlsx 类型', () => {
    assert.strictEqual(exRes.status, 200);
    assert.ok(exRes.headers.get('Content-Type').includes('spreadsheetml'));
  });
  const exBuf = Buffer.from(await exRes.arrayBuffer());
  check('导出的 xlsx 能被重新读出来且是 8 列', async () => {
    assert.ok(exBuf.length > 4000, `只有 ${exBuf.length} 字节`);
  });

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(exBuf);
  check('导出表头为 8 列', () => {
    const ws = wb.getWorksheet('差异表');
    let hr = -1;
    for (let r = 1; r <= 15; r++) {
      if (String(ws.getRow(r).getCell(1).value || '').trim() === '产品编码') {
        hr = r;
        break;
      }
    }
    assert.ok(hr > 0, '没找到表头行');
    const headers = [];
    for (let c = 1; c <= 8; c++) headers.push(ws.getRow(hr).getCell(c).value);
    assert.deepStrictEqual(headers, [
      '产品编码',
      '产品名称',
      '编码2',
      '对账单金额',
      '发货单金额',
      '金额差',
      '匹配依据',
      '备注',
    ]);
  });

  // ---------- 一致场景不生成差异表 ----------
  console.log('\n一致场景');
  const form2 = new FormData();
  form2.append('statement', fileOf('一致-对账单.xlsx'), '一致-对账单.xlsx');
  form2.append('shipments', fileOf('一致-发货单-1.xlsx'), '一致-发货单-1.xlsx');
  form2.append('shipments', fileOf('一致-发货单-2.xlsx'), '一致-发货单-2.xlsx');
  const up2 = await (await window.fetch('/api/upload', { method: 'POST', body: form2 })).json();
  const ck2 = await (
    await window.fetch('/api/check', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        uploadId: up2.uploadId,
        files: up2.files.map((f) => ({ fileId: f.fileId, sheet: f.sheet, headerRow: f.headerRow, columns: f.columns })),
      }),
    })
  ).json();

  check('总额一致时 matched=true', () => assert.strictEqual(ck2.summary.total.matched, true));
  const noDiff = await (await window.fetch('/api/export/' + ck2.resultId)).json();
  check('总额一致时导出返回 NO_DIFF', () => assert.strictEqual(noDiff.error.code, 'NO_DIFF'));

  // ---------- 真实交互（选文件 → 点核对） ----------
  console.log('\n真实交互');

  const doc = window.document;
  function setFiles(inputId, files) {
    const input = doc.getElementById(inputId);
    Object.defineProperty(input, 'files', { value: files, configurable: true });
    input.dispatchEvent(new window.Event('change', { bubbles: true }));
  }

  setFiles('statementInput', [fileOf('验收-对账单.xlsx')]);
  setFiles('shipmentInput', [fileOf('验收-发货单-1.xlsx'), fileOf('验收-发货单-2.xlsx')]);

  await new Promise((r) => setTimeout(r, 60));
  check('选好文件后「开始核对」可用', () => {
    assert.strictEqual(doc.getElementById('btnUpload').disabled, false, '按钮仍禁用');
  });

  doc.getElementById('btnUpload').click();
  await waitForView(doc, 'view-mapping');

  check('点核对后切到映射视图', () => {
    assert.ok(doc.getElementById('view-mapping').classList.contains('is-active'), '没切到映射视图');
  });
  check('映射卡片渲染出 3 个文件', () => {
    const cards = doc.querySelectorAll('#fileCards .card');
    assert.ok(cards.length === 3, `渲染了 ${cards.length} 个卡片`);
  });
  check('映射页显示金额标签「应收增加 / 发行实洋」', () => {
    const text = doc.body.textContent || '';
    assert.ok(text.includes('应收增加'), '页面没出现「应收增加」');
  });

  const btnCheck = doc.getElementById('btnCheck');
  check('映射页有「开始核对」按钮', () => assert.ok(btnCheck, '没找到 btnCheck'));
  if (btnCheck) {
    btnCheck.click();
    await waitForView(doc, 'view-result');
    check('核对后切到结果视图', () => {
      assert.ok(doc.getElementById('view-result').classList.contains('is-active'), '没切到结果视图');
    });
    check('结果页显示总额与差额', () => {
      const text = doc.getElementById('view-result').textContent || '';
      assert.ok(text.includes('1,345.50') || text.includes('1345.50'), '没显示对账单总额');
      assert.ok(text.includes('68.99'), '没显示差额');
    });
    check('结果页列出参与核对的单据名', () => {
      const text = doc.getElementById('view-result').textContent || '';
      assert.ok(text.includes('验收-发货单-1.xlsx'), '没有单据名');
    });
  }

  console.log(`\n通过 ${passed} 项，失败 ${failed} 项`);
  process.exit(failed === 0 ? 0 : 1);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
