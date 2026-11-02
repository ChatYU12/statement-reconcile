/**
 * 构建 GitHub Pages 演示版到 docs/。
 *
 * 做法：不重写业务逻辑，把 server/lib 原样包成浏览器模块（module 垫片），
 * 再由 web/api.js 在浏览器内存里实现 /api/* 接口，public/app.js 一行不改。
 *
 * 用法：node scripts/build-pages.mjs
 */

import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(ROOT, 'docs');

// 需要打进浏览器的模块（dingtalk / store 依赖 Node 的 crypto，不参与）
const MODULES = ['values', 'text', 'errors', 'detect', 'extract', 'engine', 'sheet', 'export'];

const EXPORTS = {
  loadGrid: 'sheet',
  detectHeader: 'detect',
  extractRows: 'extract',
  reconcile: 'engine',
  buildDiffWorkbook: 'export',
};

async function readLib(name) {
  const src = await fs.readFile(path.join(ROOT, 'server', 'lib', `${name}.js`), 'utf8');
  // ./values → values，让垫片的 require 认
  return src.replace(/require\('\.\/([a-z0-9]+)'\)/g, "require('$1')");
}

async function buildLibBundle() {
  const parts = [];
  for (const name of MODULES) {
    const code = await readLib(name);
    parts.push(
      `define(${JSON.stringify(name)}, function (module, exports, require, Buffer) {\n${code}\n});`,
    );
  }

  const expose = Object.entries(EXPORTS)
    .map(([fn, mod]) => `  ${fn}: require(${JSON.stringify(mod)}).${fn}`)
    .join(',\n');

  return `/**
 * 由 scripts/build-pages.mjs 从 server/lib 生成，不要手改。
 * 逻辑与 Node 版完全同一份，只是换成浏览器可用的模块加载方式。
 */
(function () {
  'use strict';
  var __factories = {};
  var __cache = {};
  var __xlsx = null;

  // server/lib/export.js 里用到了 Buffer.from(writeBuffer())，浏览器没有这个全局。
  // 这里只做它用到的那一件事，并且只在模块作用域内可见，不污染 window。
  var __Buffer = {
    from: function (input) {
      if (input instanceof Uint8Array) return input;
      if (input instanceof ArrayBuffer) return new Uint8Array(input);
      if (Array.isArray(input)) return new Uint8Array(input);
      return input;
    }
  };

  function define(name, fn) { __factories[name] = fn; }

  function require(name) {
    if (name === 'xlsx') {
      if (!__xlsx) {
        var real = window.XLSX;
        __xlsx = Object.create(real);
        // 浏览器里没有文件路径，传进来的是 ArrayBuffer
        __xlsx.readFile = function (buffer, opts) {
          return real.read(buffer, Object.assign({ type: 'array' }, opts || {}));
        };
      }
      return __xlsx;
    }
    if (name === 'exceljs') return window.ExcelJS;
    if (__cache[name]) return __cache[name].exports;
    var mod = { exports: {} };
    __cache[name] = mod;
    __factories[name](mod, mod.exports, require, __Buffer);
    return mod.exports;
  }

${parts.join('\n\n')}

  window.ReconcileLib = {
${expose}
  };
})();
`;
}

const BANNER_CSS = `

/* ---------- 在线演示版说明条（仅 Pages 版） ---------- */
.demo-banner {
  margin: 0 0 16px;
  padding: 12px 14px;
  border-radius: 10px;
  background: #eef4ff;
  border: 1px solid #c7d9fb;
  color: #23406e;
  font-size: 13px;
  line-height: 1.6;
}
.demo-banner strong { font-weight: 600; }
.demo-banner b { color: #0d57d1; }
`;

async function main() {
  await fs.rm(OUT, { recursive: true, force: true });
  await fs.mkdir(path.join(OUT, 'demo', 'vendor'), { recursive: true });

  // 1. lib bundle
  await fs.writeFile(path.join(OUT, 'demo', 'lib.js'), await buildLibBundle(), 'utf8');
  console.log('✓ demo/lib.js');

  // 2. 浏览器版第三方库
  for (const [from, to] of [
    ['node_modules/xlsx/dist/xlsx.full.min.js', 'demo/vendor/xlsx.full.min.js'],
    ['node_modules/exceljs/dist/exceljs.min.js', 'demo/vendor/exceljs.min.js'],
  ]) {
    await fs.copyFile(path.join(ROOT, from), path.join(OUT, to));
    console.log(`✓ ${to}`);
  }

  // 3. 前端三件套：绝对路径改相对（Pages 部署在子路径下）
  const app = await fs.readFile(path.join(ROOT, 'public', 'app.js'), 'utf8');
  await fs.writeFile(path.join(OUT, 'app.js'), app, 'utf8');

  const css = (await fs.readFile(path.join(ROOT, 'public', 'style.css'), 'utf8')) + BANNER_CSS;
  await fs.writeFile(path.join(OUT, 'style.css'), css, 'utf8');

  let html = await fs.readFile(path.join(ROOT, 'public', 'index.html'), 'utf8');
  html = html.replace('href="/style.css"', 'href="style.css"');
  html = html.replace(
    /<script src="\/app\.js" defer><\/script>/,
    '<script src="demo/vendor/xlsx.full.min.js" defer></script>\n' +
      '  <script src="demo/vendor/exceljs.min.js" defer></script>\n' +
      '  <script src="demo/lib.js" defer></script>\n' +
      '  <script src="demo/api.js" defer></script>\n' +
      '  <script src="app.js" defer></script>',
  );
  if (html.includes('/app.js') || html.includes('/style.css')) {
    throw new Error('index.html 里还有没改掉的绝对路径');
  }
  await fs.writeFile(path.join(OUT, 'index.html'), html, 'utf8');
  console.log('✓ index.html / app.js / style.css');

  // 4. api 层
  await fs.copyFile(path.join(ROOT, 'web', 'api.js'), path.join(OUT, 'demo', 'api.js'));
  console.log('✓ demo/api.js');

  // 5. 样例文件（用户可以直接下载来试）
  await fs.mkdir(path.join(OUT, 'samples'), { recursive: true });
  const files = await fs.readdir(path.join(ROOT, 'samples'));
  for (const f of files) {
    if (!/\.(xlsx|xls)$/.test(f)) continue;
    await fs.copyFile(path.join(ROOT, 'samples', f), path.join(OUT, 'samples', f));
  }
  console.log(`✓ samples/ (${files.filter((f) => /\.(xlsx|xls)$/.test(f)).length} 个)`);

  // 6. .nojekyll，避免 Pages 跑 Jekyll
  await fs.writeFile(path.join(OUT, '.nojekyll'), '', 'utf8');
  console.log('✓ .nojekyll');

  console.log('\n构建完成 →', OUT);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
