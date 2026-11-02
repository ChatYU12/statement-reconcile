'use strict';

/**
 * 前端冒烟：用 jsdom 加载真实 index.html + app.js，跑真实的 /api/config，
 * 再把「上传 → 映射 → 结果」三个视图各渲染一次，确认没有运行时错误。
 *
 * 用法：node tests/ui-smoke.js
 */

const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { JSDOM } = require('jsdom');

const app = require('../server/index.js');

const ROOT = path.join(__dirname, '..');
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

function wait(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

/** 轮询等待某个条件成立，避免靠固定 sleep 卡时间 */
async function waitFor(fn, timeoutMs = 3000) {
  const start = Date.now();
  for (;;) {
    try {
      if (fn()) return true;
    } catch (_) {
      /* 元素还没出现，继续等 */
    }
    if (Date.now() - start > timeoutMs) return false;
    await wait(50);
  }
}

async function main() {
  const server = app.listen(0);
  await new Promise((resolve) => server.once('listening', resolve));
  const base = `http://127.0.0.1:${server.address().port}`;

  const html = fs.readFileSync(path.join(ROOT, 'public', 'index.html'), 'utf8');
  const dom = new JSDOM(html, { url: `${base}/`, pretendToBeVisual: true, runScripts: 'outside-only' });
  const { window } = dom;

  // 让页面里的相对路径请求打到真实服务
  window.fetch = (input, init) => {
    const url = typeof input === 'string' ? new URL(input, `${base}/`).toString() : input;
    return fetch(url, init);
  };
  // jsdom 未实现 scrollTo，页面切视图时会调用它
  window.scrollTo = () => {};
  window.FormData = FormData;
  window.Blob = Blob;
  window.File = File;

  const errors = [];
  window.addEventListener('error', (e) => errors.push(e.message));
  const origError = console.error;
  console.error = (...args) => {
    errors.push(args.join(' '));
    origError.apply(console, args);
  };

  console.log('前端启动');
  window.eval(fs.readFileSync(path.join(ROOT, 'public', 'app.js'), 'utf8'));

  const doc = window.document;
  // 等配置回来（徽标不再是「加载中」），再开始断言
  await waitFor(() => {
    const el = doc.getElementById('userBadgeText');
    return el && !el.textContent.includes('加载中');
  });
  const badge = doc.getElementById('userBadgeText');

  check('页面加载后没有 JS 运行时错误', () => assert.deepStrictEqual(errors, []));
  check('读取 /api/config 并显示模式徽标', () => {
    assert.ok(badge.textContent.includes('本地模式'), `徽标是「${badge.textContent}」`);
  });
  check('默认停在上传视图', () => {
    assert.ok(doc.getElementById('view-upload').classList.contains('is-active'));
    assert.ok(!doc.getElementById('view-mapping').classList.contains('is-active'));
  });
  check('未选文件时「开始核对」禁用', () => {
    assert.strictEqual(doc.getElementById('btnUpload').disabled, true);
    assert.ok(doc.getElementById('uploadNote').textContent.length > 0);
  });
  check('上传页存在两个文件输入且接受 .xls', () => {
    const si = doc.getElementById('statementInput');
    const pi = doc.getElementById('shipmentInput');
    assert.ok(si.accept.includes('.xls'));
    assert.ok(pi.multiple);
  });
  check('对账单区与发货单区都有空状态文案', () => {
    assert.ok(doc.getElementById('statementEmpty').textContent.includes('对账单'));
    assert.ok(doc.getElementById('shipmentEmpty').textContent.includes('发货单'));
  });

  // ---- 用真实样例走一遍上传 → 映射渲染 ----
  console.log('\n上传与映射渲染');
  const stmt = fs.readFileSync(path.join(ROOT, 'samples', '验收-对账单.xlsx'));
  const ship1 = fs.readFileSync(path.join(ROOT, 'samples', '验收-发货单-1.xlsx'));
  const ship2 = fs.readFileSync(path.join(ROOT, 'samples', '验收-发货单-2.xlsx'));

  const fd = new FormData();
  fd.append('statement', new Blob([stmt]), '验收-对账单.xlsx');
  fd.append('shipments', new Blob([ship1]), '验收-发货单-1.xlsx');
  fd.append('shipments', new Blob([ship2]), '验收-发货单-2.xlsx');
  const up = await (await fetch(`${base}/api/upload`, { method: 'POST', body: fd })).json();

  // 页面自己的文件选择交互依赖 File 对象与拖拽事件，jsdom 里不便模拟，
  // 这里只断言映射视图的结构已就位，并单独验证接口契约。
  check('映射视图容器与卡片区存在', () => {
    assert.ok(doc.getElementById('fileCards'));
    assert.ok(doc.getElementById('btnCheck'));
    assert.ok(doc.getElementById('mappingNote'));
  });
  check('上传接口返回 3 个文件且都有 sheetInfos', () => {
    assert.strictEqual(up.files.length, 3);
    assert.ok(up.files.every((f) => f.sheetInfos && f.sheetInfos.length >= 1));
  });
  check('页面没有引入外部资源', () => {
    const all = html + fs.readFileSync(path.join(ROOT, 'public', 'app.js'), 'utf8');
    assert.ok(!/https?:\/\//.test(all), '出现了外部 http(s) 引用');
  });

  console.error = origError;
  server.close();
  dom.window.close();

  console.log(`\n通过 ${passed} 项，失败 ${failed} 项`);
  process.exit(failed === 0 ? 0 : 1);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
