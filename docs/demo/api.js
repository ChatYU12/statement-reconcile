/**
 * 浏览器内 API 层（GitHub Pages 演示版用）。
 *
 * 这个文件让 public/app.js 一行不改就能在纯静态环境里跑：
 * 它拦截发往 /api/* 的 fetch，用同一套 server/lib 逻辑（打到 window.ReconcileLib）
 * 在浏览器内存里实现 config / upload / sheet / check / export 五个接口。
 *
 * 与 Node 版的行为差别（都是有意为之，页面顶部有说明条）：
 *   - 文件不上传服务器，全部在本机内存里解析，关掉页面即清空
 *   - 钉钉免登不可用（没有后端换取 userId），固定为「演示模式」
 *   - 导出走 Blob 下载，不经过服务端
 */
(function (global) {
  'use strict';

  var Lib = global.ReconcileLib;
  var doc = global.document;

  var MAX_FILE_MB = 20;
  var MAX_SHIPMENTS = 30;

  var uploads = new Map();
  var results = new Map();
  var idSeq = 0;

  function newId(prefix) {
    idSeq += 1;
    return prefix + Date.now().toString(36) + idSeq.toString(36);
  }

  function fail(code, message) {
    return { ok: false, error: { code: code, message: message } };
  }

  function json(data, status) {
    return new Response(JSON.stringify(data), {
      status: status || 200,
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
    });
  }

  // ---------------- 文件解析（对应 server/routes/upload.js） ----------------

  function sniff(buffer) {
    if (!buffer || buffer.byteLength < 8) return null;
    var b = new Uint8Array(buffer);
    if (b[0] === 0x50 && b[1] === 0x4b && b[2] === 0x03 && b[3] === 0x04) return 'xlsx';
    if (b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0) return 'xls';
    return null;
  }

  function extOf(name) {
    var i = String(name || '').lastIndexOf('.');
    return i < 0 ? '' : String(name).slice(i).toLowerCase();
  }

  function analyzeBuffer(buffer, fileName) {
    var grid = Lib.loadGrid(buffer);
    var sheetInfos = [];
    for (var i = 0; i < grid.sheets.length; i++) {
      var sheet = grid.sheets[i];
      if (sheet.empty) continue;
      sheetInfos.push(Lib.detectHeader(sheet.name, sheet.rows, null));
    }
    if (sheetInfos.length === 0) {
      throw new Error('文件「' + fileName + '」里没有找到任何有数据的表');
    }
    var primary = sheetInfos[0];
    return {
      sheet: primary.sheet,
      sheets: sheetInfos.map(function (s) { return s.sheet; }),
      sheetInfos: sheetInfos,
      headerRow: primary.headerRow,
      columns: primary.columns,
      columnLabels: primary.columnLabels,
      previewRows: primary.previewRows,
      previewStartRow: primary.previewStartRow,
      warnings: [],
      error: null,
    };
  }

  async function handleUpload(form) {
    var statements = form.getAll('statement');
    var shipments = form.getAll('shipments');

    if (!statements.length) return fail('NO_FILE', '请先选择对账单文件');
    if (!shipments.length) return fail('NO_FILE', '请至少选择一个发货单文件');
    if (shipments.length > MAX_SHIPMENTS) {
      return fail('BAD_REQUEST', '发货单最多 ' + MAX_SHIPMENTS + ' 个，本次收到 ' + shipments.length + ' 个');
    }

    var all = [{ role: 'statement', file: statements[0] }];
    shipments.forEach(function (f) { all.push({ role: 'shipment', file: f }); });

    var uploadId = newId('u_');
    var warnings = [];
    var files = [];
    var stored = [];

    for (var i = 0; i < all.length; i++) {
      var role = all[i].role;
      var file = all[i].file;
      var fileId = 'f_' + (i + 1);
      var entry = {
        fileId: fileId,
        role: role,
        fileName: file.name,
        size: file.size,
        sheet: null,
        sheets: [],
        sheetInfos: [],
        warnings: [],
        error: null,
      };

      var ext = extOf(file.name);
      if (ext !== '.xlsx' && ext !== '.xls') {
        entry.error = '不支持的格式「' + (ext || '无扩展名') + '」，请上传 .xlsx 或 .xls';
        files.push(entry);
        continue;
      }
      if (file.size > MAX_FILE_MB * 1024 * 1024) {
        entry.error = '文件超过 ' + MAX_FILE_MB + 'MB，请拆分后再传';
        files.push(entry);
        continue;
      }

      var buffer;
      try {
        buffer = await file.arrayBuffer();
      } catch (e) {
        entry.error = '读取文件失败：' + e.message;
        files.push(entry);
        continue;
      }

      if (!sniff(buffer)) {
        entry.error = '文件内容不是有效的 Excel（可能已损坏或实际是 CSV），请用 Excel 另存为 .xlsx 后重试';
        files.push(entry);
        continue;
      }

      try {
        var info = analyzeBuffer(buffer, file.name);
        Object.assign(entry, info);
        entry.warnings = info.sheetInfos
          .map(function (s) { return s.warnings || []; })
          .reduce(function (a, b) { return a.concat(b); }, [])
          .slice(0, 6);
      } catch (e) {
        entry.error = e && e.message ? e.message : String(e);
      }

      if (!entry.error && entry.sheets.length > 1) {
        warnings.push('文件「' + file.name + '」有 ' + entry.sheets.length + ' 个工作表，当前使用「' + entry.sheet + '」');
      }

      stored.push({ fileId: fileId, role: role, fileName: file.name, buffer: buffer, error: entry.error });
      files.push(entry);
    }

    uploads.set(uploadId, { files: stored, createdAt: Date.now() });
    return { ok: true, uploadId: uploadId, files: files, warnings: warnings };
  }

  // ---------------- 切换工作表 / 改表头行（对应 /api/sheet） ----------------

  function handleSheet(body) {
    var batch = uploads.get(body && body.uploadId);
    if (!batch) return fail('UPLOAD_EXPIRED', '上传会话已过期，请重新上传文件');

    var target = batch.files.filter(function (f) { return f.fileId === body.fileId; })[0];
    if (!target) return fail('BAD_REQUEST', '未找到对应的文件，请重新上传');

    var sheetName = String(body.sheet || '').trim();
    var headerRow = null;
    if (body.headerRow !== undefined && body.headerRow !== null && body.headerRow !== '') {
      headerRow = Number(body.headerRow);
      if (!Number.isInteger(headerRow) || headerRow < 1 || headerRow > 5000) {
        return fail('BAD_REQUEST', '表头行必须是 1~5000 的整数');
      }
    }

    var grid = Lib.loadGrid(target.buffer);
    var sheet = grid.sheets.filter(function (s) { return s.name === sheetName; })[0];
    if (!sheet) {
      return fail('BAD_REQUEST', '工作表「' + sheetName + '」不存在或为空');
    }

    var info = Lib.detectHeader(sheet.name, sheet.rows, headerRow);
    var payload = { ok: true };
    Object.keys(info).forEach(function (k) { payload[k] = info[k]; });
    return payload;
  }

  // ---------------- 核对（对应 /api/check） ----------------

  var COLUMN_KEYS = ['name', 'code', 'code2', 'publisher', 'qty', 'amount', 'unitPrice'];
  var COLUMN_LABELS = {
    name: '产品名称',
    code: '编码1（书号/ISBN）',
    code2: '编码2（货号/内部号）',
    publisher: '出版社',
    qty: '数量（不参与核对）',
    amount: '金额（应收增加 / 发行实洋）',
    unitPrice: '单价（不参与核对）',
  };

  function normalizeMapping(fileName, columns, role) {
    if (!columns || typeof columns !== 'object') {
      throw new Error('文件「' + fileName + '」缺少列映射');
    }
    var normalized = {};
    for (var i = 0; i < COLUMN_KEYS.length; i++) {
      var key = COLUMN_KEYS[i];
      var raw = columns[key];
      if (raw === null || raw === undefined || raw === '') {
        normalized[key] = null;
        continue;
      }
      var num = Number(raw);
      if (!Number.isInteger(num) || num < 1 || num > 200) {
        throw new Error('文件「' + fileName + '」的「' + COLUMN_LABELS[key] + '」列号不合法：' + raw);
      }
      normalized[key] = num;
    }
    if (!normalized.amount && role === 'statement') {
      throw new Error('对账单「' + fileName + '」没有指定金额列（应收增加），无法核对总额');
    }
    if (!normalized.name && !normalized.code && !normalized.code2) {
      throw new Error('文件「' + fileName + '」没有指定产品名称或任一编码列（否则无法匹配产品）');
    }
    return normalized;
  }

  function handleCheck(body) {
    var uploadId = body && body.uploadId;
    if (!uploadId || typeof uploadId !== 'string') {
      return fail('BAD_REQUEST', '缺少 uploadId，请重新上传文件');
    }
    var batch = uploads.get(uploadId);
    if (!batch) return fail('UPLOAD_EXPIRED', '上传会话已过期（超过 2 小时），请重新上传文件');

    var files = body.files;
    if (!Array.isArray(files) || files.length === 0) {
      return fail('BAD_REQUEST', '缺少文件列映射配置');
    }

    var byId = new Map();
    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      if (!f || typeof f.fileId !== 'string') return fail('BAD_REQUEST', '列映射缺少 fileId');
      if (byId.has(f.fileId)) return fail('BAD_REQUEST', '文件 ' + f.fileId + ' 的列映射重复');
      byId.set(f.fileId, f);
    }
    if (byId.size !== batch.files.length) {
      return fail('BAD_REQUEST', '有 ' + batch.files.length + ' 个文件，只收到 ' + byId.size + ' 个文件的列映射');
    }

    var extracted = [];
    try {
      for (var j = 0; j < batch.files.length; j++) {
        var bf = batch.files[j];
        var conf = byId.get(bf.fileId);
        if (bf.error) throw new Error('文件「' + bf.fileName + '」无法解析：' + bf.error);
        var columns = normalizeMapping(bf.fileName, conf.columns, bf.role);
        var headerRow = Number(conf.headerRow);
        if (!Number.isInteger(headerRow) || headerRow < 1 || headerRow > 5000) {
          throw new Error('文件「' + bf.fileName + '」的表头行不合法：' + conf.headerRow);
        }

        var grid = Lib.loadGrid(bf.buffer);
        var sheet = grid.sheets.filter(function (s) { return s.name === conf.sheet; })[0];
        if (!sheet || sheet.empty) {
          throw new Error('文件「' + bf.fileName + '」的工作表「' + conf.sheet + '」不存在或为空');
        }

        var auto = Lib.detectHeader(sheet.name, sheet.rows, null);
        var headerSpan = auto.headerRow === headerRow && auto.headerSpan === 2 ? 2 : 1;

        var data = Lib.extractRows({ rows: sheet.rows, headerRow: headerRow, headerSpan: headerSpan, columns: columns });

        extracted.push({
          fileId: bf.fileId,
          fileName: bf.fileName,
          role: bf.role,
          items: data.items,
          skipped: data.skipped,
          warnings: data.warnings,
          amountMode: data.amountMode,
          amountChecked: data.amountChecked,
        });
      }
    } catch (e) {
      return fail('MAPPING_MISSING', e && e.message ? e.message : String(e));
    }

    var statement = extracted.filter(function (e) { return e.role === 'statement'; })[0];
    var shipments = extracted.filter(function (e) { return e.role === 'shipment'; });
    if (!statement) return fail('BAD_REQUEST', '缺少对账单文件');
    if (!shipments.length) return fail('BAD_REQUEST', '缺少发货单文件');

    var out = Lib.reconcile({ statement: statement, shipments: shipments });
    var resultId = newId('r_');
    results.set(resultId, { summary: out.summary, rows: out.rows, suspects: out.suspects });

    return { ok: true, resultId: resultId, summary: out.summary, rows: out.rows, suspects: out.suspects };
  }

  // ---------------- 导出（对应 /api/export/:id） ----------------

  async function buildExportBlob(resultId) {
    var result = results.get(resultId);
    if (!result) return { error: fail('RESULT_EXPIRED', '核对结果已过期，请重新核对') };
    if (result.summary.total && result.summary.total.matched) {
      return { error: fail('NO_DIFF', '本次核对总额一致，不生成差异表') };
    }
    var buffer = await Lib.buildDiffWorkbook(result);
    var d = new Date();
    var p = function (n) { return String(n).padStart(2, '0'); };
    var stamp = '' + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '-' + p(d.getHours()) + p(d.getMinutes());
    return {
      blob: new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      fileName: '差异表-' + stamp + '.xlsx',
    };
  }

  function saveBlob(blob, fileName) {
    var url = URL.createObjectURL(blob);
    var a = doc.createElement('a');
    a.href = url;
    a.download = fileName;
    doc.body.appendChild(a);
    a.click();
    doc.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  }

  // ---------------- fetch 拦截 ----------------

  function pathOf(input) {
    var url = typeof input === 'string' ? input : (input && input.url) || '';
    try {
      var u = new URL(url, global.location.href);
      return u.pathname;
    } catch (e) {
      return url;
    }
  }

  function readBody(init) {
    if (!init || init.body === undefined || init.body === null) return null;
    var body = init.body;
    if (typeof body === 'string') {
      try { return JSON.parse(body); } catch (e) { return null; }
    }
    if (typeof FormData !== 'undefined' && body instanceof FormData) return body;
    return null;
  }

  var origFetch = global.fetch;

  global.fetch = function (input, init) {
    var path = pathOf(input);

    if (path.indexOf('/api/') !== 0) {
      return origFetch.apply(global, arguments);
    }

    var options = init || {};
    var method = String(options.method || 'GET').toUpperCase();

    if (path === '/api/config') {
      return Promise.resolve(json({
        ok: true,
        mode: 'local',
        corpId: '',
        user: { id: 'local', name: '演示模式' },
        limits: { maxFileSizeMB: MAX_FILE_MB, maxShipments: MAX_SHIPMENTS },
        acceptExt: ['.xlsx', '.xls'],
        fieldLabels: {
          name: '产品名称',
          code: '编码1（书号/ISBN）',
          code2: '编码2（货号/内部号）',
          publisher: '出版社',
          qty: '数量（不参与核对）',
          amount: '金额（应收增加 / 发行实洋）',
          unitPrice: '单价（不参与核对）',
        },
      }));
    }

    if (path === '/api/auth/dingtalk') {
      return Promise.resolve(json({ ok: true, user: { id: 'local', name: '演示模式' }, mode: 'local' }));
    }

    if (path === '/api/upload' && method === 'POST') {
      var form = readBody(options);
      if (!form) return Promise.resolve(json(fail('BAD_REQUEST', '请求格式不正确'), 400));
      return handleUpload(form).then(function (data) {
        return json(data, data.ok ? 200 : 400);
      });
    }

    if (path === '/api/sheet' && method === 'POST') {
      var sheetBody = readBody(options);
      var sheetOut = handleSheet(sheetBody || {});
      return Promise.resolve(json(sheetOut, sheetOut.ok ? 200 : 400));
    }

    if (path === '/api/check' && method === 'POST') {
      var checkOut = handleCheck(readBody(options) || {});
      return Promise.resolve(json(checkOut, checkOut.ok ? 200 : 400));
    }

    if (path.indexOf('/api/export/') === 0) {
      var resultId = path.slice('/api/export/'.length);
      if (method === 'HEAD') {
        var exists = results.has(resultId);
        return Promise.resolve(new Response(null, { status: exists ? 200 : 400 }));
      }
      return buildExportBlob(resultId).then(function (out) {
        if (out.error) return json(out.error, 400);
        return new Response(out.blob, {
          status: 200,
          headers: {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Length': String(out.blob.size),
          },
        });
      });
    }

    return Promise.resolve(json(fail('NOT_FOUND', '演示版不支持该接口：' + path), 404));
  };

  // 结果页的「下载」按钮是 <a href="/api/export/…">，点击会真的跳转。
  // 在捕获阶段拦下来，改成 Blob 下载。
  doc.addEventListener('click', function (event) {
    var node = event.target;
    while (node && node !== doc) {
      if (node.tagName === 'A') {
        var href = node.getAttribute('href') || '';
        if (href.indexOf('/api/export/') === 0) {
          event.preventDefault();
          event.stopPropagation();
          var rid = href.slice('/api/export/'.length);
          buildExportBlob(rid).then(function (out) {
            if (out.error) {
              global.alert(out.error.error.message);
              return;
            }
            saveBlob(out.blob, out.fileName);
          });
          return;
        }
      }
      node = node.parentNode;
    }
  }, true);

  // 演示版说明条：让人一眼知道文件没有上传到任何服务器
  function mountBanner() {
    var view = doc.getElementById('view-upload');
    if (!view || doc.getElementById('demoBanner')) return;
    var bar = doc.createElement('div');
    bar.id = 'demoBanner';
    bar.className = 'demo-banner';
    bar.innerHTML =
      '<strong>在线演示版</strong>：文件只在你自己的浏览器里解析，<b>不会上传到任何服务器</b>，关掉页面即清空。' +
      '（正式版是钉钉工作台里的 Node 服务，功能相同）';
    view.insertBefore(bar, view.firstChild);
  }

  if (doc.readyState === 'loading') {
    doc.addEventListener('DOMContentLoaded', mountBanner);
  } else {
    mountBanner();
  }

  global.DemoApi = { uploads: uploads, results: results };
})(window);
