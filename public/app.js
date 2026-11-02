/* =========================================================
   对账单核对 · 前端交互（契约 v2）
   纯原生 JS，无构建、无框架、无外部请求。

   金额 / 数量一律按后端返回的字符串展示，前端不做任何加减或浮点运算；
   千分位、负数判定、计数判断均为纯字符串处理。
   ========================================================= */

(function () {
  'use strict';

  /* ==================== 常量 ==================== */

  var DEFAULT_LIMITS = { maxFileSizeMB: 20, maxShipments: 30 };
  var DEFAULT_ACCEPT_EXT = ['.xlsx', '.xls'];
  var MAX_HEADER_ROW = 200;
  var STEP_ORDER = ['upload', 'mapping', 'result'];
  var HEADER_REFRESH_DELAY = 450;

  var DEFAULT_FIELD_LABELS = {
    name: '产品名称',
    code: '编码1（书号/ISBN）',
    code2: '编码2（货号/内部号）',
    publisher: '出版社',
    qty: '数量（不参与核对）',
    amount: '金额（应收增加 / 发行实洋）',
    unitPrice: '单价（不参与核对）'
  };

  /* 7 个列位，顺序即卡片里的展示顺序 */
  var FIELD_DEFS = [
    { key: 'name' },
    { key: 'code' },
    { key: 'code2' },
    { key: 'publisher' },
    { key: 'qty' },
    { key: 'amount' },
    { key: 'unitPrice' }
  ];

  var MATCH_BY_META = {
    code: { label: '编码1', tone: 'blue' },
    code2: { label: '编码2', tone: 'blue' },
    name: { label: '名称', tone: 'gray' },
    fuzzy: { label: '模糊匹配', tone: 'orange' }
  };

  var MATCH_STAT_DEFS = [
    { key: 'code', label: '编码1' },
    { key: 'code2', label: '编码2' },
    { key: 'name', label: '名称' },
    { key: 'fuzzy', label: '模糊' }
  ];

  var GROUP_META = [
    {
      category: 'amount_diff',
      tone: 'red',
      title: '金额不一致',
      desc: '该产品两边金额存在差异，需逐条核对。',
      collapsible: false
    },
    {
      category: 'statement_only',
      tone: 'orange',
      title: '仅对账单有',
      desc: '对账单里出现，但发货单汇总里没有。',
      collapsible: false
    },
    {
      category: 'shipment_only',
      tone: 'orange',
      title: '仅发货单有',
      desc: '发货单里出现，但对账单里没有。',
      collapsible: false
    },
    {
      category: 'unknown_amount',
      tone: 'blue',
      title: '金额未知（单据无金额列）',
      desc: '该产品在发货单侧有记录，但相关单据都没有金额列，无法判断金额。',
      collapsible: false
    },
    {
      category: 'match',
      tone: 'green',
      title: '金额一致',
      desc: '两边金额完全一致，默认折叠。',
      collapsible: true
    }
  ];

  /* ==================== 状态 ==================== */

  var state = {
    config: null,
    limits: { maxFileSizeMB: DEFAULT_LIMITS.maxFileSizeMB, maxShipments: DEFAULT_LIMITS.maxShipments },
    acceptExt: DEFAULT_ACCEPT_EXT.slice(),
    fieldLabels: Object.assign({}, DEFAULT_FIELD_LABELS),
    mode: 'local',
    user: null,
    view: 'upload',

    statementFile: null,
    shipmentFiles: [],

    uploadId: null,
    lastSignature: null,
    files: [],
    serverWarnings: [],

    result: null,
    resultId: null,

    busy: false
  };

  /* ==================== 小工具 ==================== */

  function $(id) { return document.getElementById(id); }

  function mk(tag, className, text) {
    var n = document.createElement(tag);
    if (className) n.className = className;
    if (text !== undefined && text !== null) n.textContent = String(text);
    return n;
  }

  function clear(n) {
    if (!n) return;
    while (n.firstChild) n.removeChild(n.firstChild);
  }

  function setText(n, text) {
    if (n) n.textContent = text === null || text === undefined ? '' : String(text);
  }

  /** 空值统一显示为破折号 */
  function displayOrDash(v) {
    if (v === null || v === undefined) return '—';
    var s = String(v).trim();
    return s === '' ? '—' : s;
  }

  function fieldLabel(key) {
    return state.fieldLabels[key] || DEFAULT_FIELD_LABELS[key] || key;
  }

  /** 金额 / 数量：纯字符串加千分位，不做任何数值运算 */
  function groupThousands(value) {
    var raw = String(value === null || value === undefined ? '' : value).trim();
    if (!raw) return '';

    var negative = false;
    if (raw.charAt(0) === '-') { negative = true; raw = raw.slice(1); }
    else if (raw.charAt(0) === '+') { raw = raw.slice(1); }

    var dotAt = raw.indexOf('.');
    var intPart = dotAt === -1 ? raw : raw.slice(0, dotAt);
    var fracPart = dotAt === -1 ? '' : raw.slice(dotAt);

    if (!/^[0-9]+$/.test(intPart)) return (negative ? '-' : '') + raw;

    var grouped = '';
    var counter = 0;
    for (var i = intPart.length - 1; i >= 0; i--) {
      grouped = intPart.charAt(i) + grouped;
      counter++;
      if (counter % 3 === 0 && i > 0) grouped = ',' + grouped;
    }
    return (negative ? '-' : '') + grouped + fracPart;
  }

  function moneyText(v) {
    if (v === null || v === undefined || v === '') return '—';
    return groupThousands(v);
  }

  function countText(v) {
    if (v === null || v === undefined || v === '') return '—';
    return groupThousands(v);
  }

  /** 纯字符串判断负数（-0.00 视为非负数） */
  function isNegativeText(v) {
    var s = String(v === null || v === undefined ? '' : v).trim();
    if (s.charAt(0) !== '-') return false;
    for (var i = 1; i < s.length; i++) {
      var ch = s.charAt(i);
      if (ch >= '1' && ch <= '9') return true;
    }
    return false;
  }

  /** 纯字符串判断计数值是否大于 0 */
  function isPositiveCount(v) {
    var s = String(v === null || v === undefined ? '' : v);
    return /[1-9]/.test(s);
  }

  function formatBytes(size) {
    var bytes = Number(size);
    if (!isFinite(bytes) || bytes <= 0) return '0 KB';
    if (bytes < 1048576) {
      var kb = Math.round(bytes / 1024);
      if (kb < 1) kb = 1;
      return String(kb) + ' KB';
    }
    var mb = (Math.round((bytes / 1048576) * 10) / 10).toFixed(1);
    return String(mb) + ' MB';
  }

  function normalizeInt(value, min, max) {
    var n = Number(value);
    if (!isFinite(n)) return null;
    if (Math.floor(n) !== n) return null;
    if (n < min || n > max) return null;
    return n;
  }

  function normalizeColumn(value) {
    return normalizeInt(value, 1, 100000);
  }

  function limits() {
    return state.limits || DEFAULT_LIMITS;
  }

  function roleLabel(role) {
    return role === 'statement' ? '对账单' : '发货单';
  }

  /* ==================== 统一请求 ==================== */

  function request(url, options) {
    var opts = Object.assign({ credentials: 'same-origin', cache: 'no-store' }, options || {});
    return fetch(url, opts).then(function (res) {
      return res.json().catch(function () { return null; }).then(function (data) {
        if (!res.ok || !data || data.ok !== true) {
          var message = data && data.error && data.error.message
            ? data.error.message
            : '请求失败（HTTP ' + String(res.status) + '）';
          var err = new Error(message);
          err.code = data && data.error && data.error.code ? data.error.code : 'HTTP_' + String(res.status);
          err.detail = data && data.error ? data.error.detail : null;
          err.status = res.status;
          throw err;
        }
        return data;
      });
    });
  }

  /* ==================== Toast ==================== */

  var TOAST_LIMIT = 4;

  function toast(message, tone) {
    var host = $('toastHost');
    if (!host || !message) return;

    var type = tone || 'info';
    var box = mk('div', 'toast toast--' + type);
    box.setAttribute('role', type === 'error' ? 'alert' : 'status');

    var text = mk('p', 'toast__text', message);
    var close = mk('button', 'toast__close', '关闭');
    close.type = 'button';
    close.setAttribute('aria-label', '关闭提示');

    var removed = false;
    var timer = null;

    function remove() {
      if (removed) return;
      removed = true;
      if (timer) clearTimeout(timer);
      box.classList.add('is-out');
      setTimeout(function () {
        if (box.parentNode) box.parentNode.removeChild(box);
      }, 190);
    }

    close.addEventListener('click', remove);
    box.appendChild(text);
    box.appendChild(close);

    while (host.children.length >= TOAST_LIMIT) {
      host.removeChild(host.firstChild);
    }

    host.appendChild(box);
    timer = setTimeout(remove, 3500);
  }

  /* ==================== 顶部提示条 ==================== */

  function showAuthNotice(message) {
    var box = $('authNotice');
    if (!box) return;
    setText($('authNoticeText'), message);
    box.hidden = false;
  }

  function hideAuthNotice() {
    var box = $('authNotice');
    if (box) box.hidden = true;
  }

  function showMappingError(message) {
    var box = $('mappingError');
    if (!box) return;
    setText($('mappingErrorText'), message);
    box.hidden = !message;
  }

  function hideMappingError() {
    var box = $('mappingError');
    if (box) box.hidden = true;
  }

  /* ==================== 视图切换 ==================== */

  function updateSteps(view) {
    var activeIndex = STEP_ORDER.indexOf(view);
    var items = document.querySelectorAll('#steps .step');
    Array.prototype.forEach.call(items, function (li, index) {
      var isActive = index === activeIndex;
      li.classList.toggle('is-active', isActive);
      li.classList.toggle('is-done', index < activeIndex);
      if (isActive) li.setAttribute('aria-current', 'step');
      else li.removeAttribute('aria-current');
    });
  }

  function showView(view, opts) {
    var options = opts || {};
    state.view = view;

    STEP_ORDER.forEach(function (name) {
      var section = $('view-' + name);
      if (!section) return;
      var active = name === view;
      section.classList.toggle('is-active', active);
      if (active) section.removeAttribute('aria-hidden');
      else section.setAttribute('aria-hidden', 'true');
    });

    updateSteps(view);

    try { window.scrollTo(0, 0); } catch (e) { /* 忽略 */ }

    if (options.focus === false) return;
    var title = $('view-' + view) ? $('view-' + view).querySelector('.page-title') : null;
    if (title && typeof title.focus === 'function') {
      try { title.focus({ preventScroll: true }); } catch (e) {
        try { title.focus(); } catch (e2) { /* 忽略 */ }
      }
    }
  }

  /* ==================== 用户徽标 ==================== */

  function setUser(user, mode) {
    state.user = user || null;
    state.mode = mode || 'local';

    var name = user && user.name ? String(user.name) : (state.mode === 'dingtalk' ? '钉钉用户' : '本地模式');
    var badge = $('userBadge');
    var label = name + (state.mode === 'dingtalk' ? ' · 钉钉' : ' · 本地');

    setText($('userBadgeText'), label);
    if (badge) {
      badge.classList.toggle('is-dingtalk', state.mode === 'dingtalk');
      badge.classList.toggle('is-local', state.mode !== 'dingtalk');
      badge.title = label;
    }
  }

  /* ==================== 文件签名 / 选择 ==================== */

  function fileSignature(file) {
    if (!file) return '';
    return [file.name || '', String(file.size), String(file.lastModified || 0)].join('|');
  }

  function collectionSignature() {
    var parts = [];
    parts.push(state.statementFile ? 'S:' + fileSignature(state.statementFile) : 'S:');
    state.shipmentFiles.forEach(function (f) {
      parts.push('D:' + fileSignature(f));
    });
    return parts.join(';;');
  }

  function uploadKey(role, name, size) {
    return role + '|' + String(name || '') + '|' + String(size || '');
  }

  function acceptText() {
    var list = state.acceptExt && state.acceptExt.length ? state.acceptExt : DEFAULT_ACCEPT_EXT;
    return list.join(' / ');
  }

  function isAcceptedName(name) {
    var lower = String(name || '').toLowerCase();
    var list = state.acceptExt && state.acceptExt.length ? state.acceptExt : DEFAULT_ACCEPT_EXT;
    var accepted = false;
    list.forEach(function (ext) {
      var suffix = String(ext || '').toLowerCase();
      if (suffix && lower.length > suffix.length && lower.slice(-suffix.length) === suffix) accepted = true;
    });
    return accepted;
  }

  function pickValidFiles(fileList) {
    var valid = [];
    var maxBytes = Number(limits().maxFileSizeMB) * 1024 * 1024;
    var maxMB = limits().maxFileSizeMB;

    Array.prototype.forEach.call(fileList || [], function (file) {
      if (!file) return;
      var name = file.name ? String(file.name) : '';
      if (!isAcceptedName(name)) {
        toast('「' + (name || '未命名文件') + '」格式不支持：仅支持 ' + acceptText(), 'error');
        return;
      }
      if (file.size <= 0) {
        toast('「' + name + '」是空文件，请换一份。', 'error');
        return;
      }
      if (file.size > maxBytes) {
        toast('「' + name + '」超过 ' + String(maxMB) + 'MB，请压缩或拆分后再上传。', 'error');
        return;
      }
      valid.push(file);
    });

    return valid;
  }

  /* ==================== 上传页渲染 ==================== */

  function buildFileRow(file, actionName, metaText, extraLabel) {
    var li = mk('li', 'filerow');

    var ext = mk('span', 'filerow__ext', /\.xlsx$/i.test(file.name || '') ? 'XLSX' : 'XLS');
    ext.setAttribute('aria-hidden', 'true');

    var body = mk('div', 'filerow__body');
    body.appendChild(mk('p', 'filerow__name', file.name));
    body.appendChild(mk('p', 'filerow__meta', metaText));

    var btn = mk('button', 'iconbtn', '移除');
    btn.type = 'button';
    btn.setAttribute('data-action', actionName);
    btn.setAttribute('aria-label', '移除' + (extraLabel || '') + file.name);

    li.appendChild(ext);
    li.appendChild(body);
    li.appendChild(btn);
    return li;
  }

  function renderPicked() {
    var statementList = $('statementList');
    var shipmentList = $('shipmentList');
    if (!statementList || !shipmentList) return;

    clear(statementList);
    clear(shipmentList);

    if (state.statementFile) {
      statementList.appendChild(buildFileRow(state.statementFile, 'remove-statement', formatBytes(state.statementFile.size), '对账单 '));
      $('statementEmpty').hidden = true;
    } else {
      $('statementEmpty').hidden = false;
    }

    state.shipmentFiles.forEach(function (file, index) {
      var li = buildFileRow(file, 'remove-shipment', formatBytes(file.size), '发货单 ');
      li.querySelector('[data-action="remove-shipment"]').setAttribute('data-index', String(index));
      shipmentList.appendChild(li);
    });

    var count = state.shipmentFiles.length;
    $('shipmentEmpty').hidden = count > 0;
    setText($('shipmentCountHint'), '已选 ' + String(count) + ' 张，最多 ' + String(limits().maxShipments) + ' 张');

    updateUploadReady();
  }

  function uploadBlockers() {
    var out = [];
    if (!state.statementFile) out.push('请选择对账单');
    if (!state.shipmentFiles.length) out.push('请至少选择 1 份发货单');
    return out;
  }

  function updateUploadReady() {
    var btn = $('btnUpload');
    var note = $('uploadNote');
    if (!btn || !note) return;
    if (state.busy) return;

    var blockers = uploadBlockers();
    btn.disabled = blockers.length > 0;
    if (blockers.length) {
      note.textContent = blockers.join('，');
      note.classList.remove('is-ready');
    } else {
      note.textContent = '已就绪：1 份对账单 + ' + String(state.shipmentFiles.length) + ' 份发货单';
      note.classList.add('is-ready');
    }
  }

  function setUploadBusy(busy, label) {
    state.busy = busy;
    document.body.classList.toggle('is-busy', busy);

    var spinner = $('uploadSpinner');
    var skeleton = $('uploadSkeleton');
    var pickStatement = $('statementPick');
    var pickShipment = $('shipmentPick');

    if (spinner) spinner.hidden = !busy;
    if (skeleton) skeleton.hidden = !busy;
    if (pickStatement) pickStatement.disabled = busy;
    if (pickShipment) pickShipment.disabled = busy;

    setText($('uploadBtnLabel'), busy ? (label || '正在解析…') : '开始核对');

    var btn = $('btnUpload');
    if (btn) btn.disabled = busy ? true : uploadBlockers().length > 0;
  }

  /* ==================== 上传 / 解析 ==================== */

  function handleStatementFiles(fileList) {
    if (state.busy) return;
    var valid = pickValidFiles(fileList);
    if (!valid.length) return;
    if (valid.length > 1) toast('对账单只保留 1 份，已取第一个文件。', 'warn');
    state.statementFile = valid[0];
    renderPicked();
  }

  function handleShipmentFiles(fileList) {
    if (state.busy) return;
    var valid = pickValidFiles(fileList);
    if (!valid.length) return;

    var max = Number(limits().maxShipments);
    var added = 0;
    var duplicated = 0;
    var overflow = false;

    valid.forEach(function (file) {
      var sig = fileSignature(file);
      var exists = state.shipmentFiles.some(function (item) {
        return fileSignature(item) === sig;
      });
      if (exists) { duplicated++; return; }
      if (state.shipmentFiles.length >= max) { overflow = true; return; }
      state.shipmentFiles.push(file);
      added++;
    });

    if (duplicated > 0) toast('已跳过 ' + String(duplicated) + ' 个重复选择的文件。', 'warn');
    if (overflow) toast('发货单最多 ' + String(max) + ' 张，多余的文件未加入。', 'warn');

    if (added > 0) renderPicked();
  }

  /* ==================== sheet 探测结果 ==================== */

  /** 文件级与 sheet 级提示合并去重，两边都不能被对方覆盖 */
  function mergeWarnings(a, b) {
    var out = [];
    [a, b].forEach(function (list) {
      if (!Array.isArray(list)) return;
      list.forEach(function (text) {
        var str = String(text);
        if (out.indexOf(str) === -1) out.push(str);
      });
    });
    return out;
  }

  function normalizeSheetInfo(raw) {
    var source = raw || {};
    return {
      name: source.name || source.sheet || '',
      rowCount: source.rowCount,
      headerRow: source.headerRow,
      columns: source.columns,
      columnLabels: source.columnLabels && typeof source.columnLabels === 'object' ? source.columnLabels : {},
      previewRows: Array.isArray(source.previewRows) ? source.previewRows : [],
      previewStartRow: source.previewStartRow,
      warnings: Array.isArray(source.warnings) ? source.warnings.slice() : []
    };
  }

  /** 把当前 file 的活动 sheet 状态写回 sheetInfos 缓存，切走再切回可复原 */
  function cacheSheetInfo(file) {
    if (!file.sheet) return;
    var snapshot = {
      name: file.sheet,
      rowCount: file.rowCount,
      headerRow: file.headerRow,
      columns: Object.assign({}, file.columns),
      columnLabels: Object.assign({}, file.columnLabels),
      previewRows: file.previewRows,
      previewStartRow: file.previewStartRow,
      warnings: (file.sheetWarnings || []).slice()
    };

    var replaced = false;
    file.sheetInfos = (file.sheetInfos || []).map(function (info) {
      if (info && info.name === snapshot.name) { replaced = true; return snapshot; }
      return info;
    });
    if (!replaced) file.sheetInfos.push(snapshot);
  }

  function findSheetInfo(file, name) {
    var found = null;
    (file.sheetInfos || []).forEach(function (info) {
      if (!found && info && info.name === name) found = info;
    });
    return found;
  }

  function applySheetInfo(file, info) {
    if (!info) return;

    file.sheet = info.name || file.sheet;
    file.rowCount = info.rowCount;
    file.headerRow = normalizeInt(info.headerRow, 1, MAX_HEADER_ROW);
    if (file.headerRow === null) file.headerRow = 1;
    file.columns = normalizeColumns(info.columns);
    file.columnLabels = info.columnLabels && typeof info.columnLabels === 'object' ? info.columnLabels : {};
    file.previewRows = Array.isArray(info.previewRows) ? info.previewRows : [];
    var startRow = normalizeInt(info.previewStartRow, 1, 1000000);
    file.previewStartRow = startRow === null ? 1 : startRow;
    file.sheetWarnings = Array.isArray(info.warnings) ? info.warnings.slice() : [];
    file.warnings = mergeWarnings(file.sheetWarnings, file.fileWarnings);

    cacheSheetInfo(file);
  }

  function applyUploadResponse(data, signature) {
    var previous = {};
    state.files.forEach(function (file) {
      previous[file.key] = file;
    });

    state.uploadId = data.uploadId;
    state.serverWarnings = Array.isArray(data.warnings) ? data.warnings.slice() : [];

    var rawFiles = Array.isArray(data.files) ? data.files : [];

    state.files = rawFiles.map(function (raw) {
      var role = raw.role === 'statement' ? 'statement' : 'shipment';
      var sheetList = Array.isArray(raw.sheets) && raw.sheets.length
        ? raw.sheets.slice()
        : [raw.sheet || 'Sheet1'];

      var infos = Array.isArray(raw.sheetInfos)
        ? raw.sheetInfos.map(normalizeSheetInfo).filter(function (info) { return !!info.name; })
        : [];

      // 兼容没有 sheetInfos 的实现：把顶层字段当作唯一工作表
      if (!infos.length) {
        infos = [normalizeSheetInfo({
          name: raw.sheet || sheetList[0],
          rowCount: raw.rowCount,
          headerRow: raw.headerRow,
          columns: raw.columns,
          columnLabels: raw.columnLabels,
          previewRows: raw.previewRows,
          previewStartRow: raw.previewStartRow || 1,
          warnings: raw.warnings
        })];
      }

      var active = null;
      infos.forEach(function (info) {
        if (!active && info.name === raw.sheet) active = info;
      });
      if (!active) active = infos[0];

      var file = {
        fileId: raw.fileId,
        role: role,
        fileName: raw.fileName || '未命名文件',
        size: Number(raw.size) || 0,
        key: uploadKey(role, raw.fileName, raw.size),
        sheets: sheetList,
        sheetInfos: infos,
        sheet: active.name || sheetList[0],
        rowCount: active.rowCount,
        headerRow: normalizeInt(active.headerRow, 1, MAX_HEADER_ROW) || 1,
        columns: normalizeColumns(active.columns),
        columnLabels: active.columnLabels || {},
        previewRows: active.previewRows || [],
        previewStartRow: normalizeInt(active.previewStartRow, 1, 1000000) || 1,
        sheetWarnings: Array.isArray(active.warnings) ? active.warnings.slice() : [],
        fileWarnings: Array.isArray(raw.warnings) ? raw.warnings.slice() : [],
        warnings: [],
        error: raw.error ? String(raw.error) : null,
        ui: null
      };
      file.warnings = mergeWarnings(file.sheetWarnings, file.fileWarnings);

      // 同一文件重传时保留用户此前调整过的映射
      var inherited = previous[file.key];
      if (inherited && !file.error) {
        var inheritedInfo = findSheetInfo(inherited, file.sheet);
        if (inheritedInfo) applySheetInfo(file, inheritedInfo);
      }

      return file;
    });

    state.lastSignature = signature;
  }

  function startUpload() {
    if (state.busy) return;
    var blockers = uploadBlockers();
    if (blockers.length) {
      toast(blockers.join('，'), 'error');
      return;
    }

    var signature = collectionSignature();
    if (state.uploadId && state.lastSignature === signature) {
      // 文件未变化：复用已解析结果，保留用户改过的映射
      enterMapping();
      return;
    }

    setUploadBusy(true, '正在解析…');

    var form = new FormData();
    form.append('statement', state.statementFile, state.statementFile.name);
    state.shipmentFiles.forEach(function (file) {
      form.append('shipments', file, file.name);
    });

    request('/api/upload', { method: 'POST', body: form })
      .then(function (data) {
        applyUploadResponse(data, signature);
        setUploadBusy(false);
        enterMapping();
      })
      .catch(function (err) {
        setUploadBusy(false);
        updateUploadReady();
        toast(err.message || '文件解析失败，请重试。', 'error');
      });
  }

  /* ==================== 映射：有效性 ==================== */

  function normalizeColumns(raw) {
    var out = { name: null, code: null, code2: null, publisher: null, qty: null, amount: null, unitPrice: null };
    if (!raw || typeof raw !== 'object') return out;
    FIELD_DEFS.forEach(function (def) {
      out[def.key] = normalizeColumn(raw[def.key]);
    });
    return out;
  }

  function fileIssues(file) {
    var out = [];

    if (file.error) {
      out.push('该文件解析失败：' + file.error + '。请返回上传页移除后重新上传。');
      return out;
    }

    if (!file.sheet) out.push('请选择工作表。');
    if (file.headerRow === null || file.headerRow === undefined) {
      out.push('表头行需为 1–' + String(MAX_HEADER_ROW) + ' 的整数。');
    }

    var c = file.columns || {};

    // 阻断一：对账单必须有金额列，否则无法核对总额
    if (file.role === 'statement' && !c.amount) {
      out.push('对账单必须指定「' + fieldLabel('amount') + '」列，否则无法核对总额。');
    }

    // 阻断二：必须有产品标识，否则无法定位到具体产品
    if (!c.name && !c.code && !c.code2) {
      out.push('「' + fieldLabel('name') + '」「' + fieldLabel('code') + '」「' + fieldLabel('code2') +
        '」至少指定一个，否则无法定位产品。');
    }

    return out;
  }

  /** 非阻断提示：发货单没有金额列时该文件不参与金额核对 */
  function fileNotes(file) {
    if (file.error) return [];
    var c = file.columns || {};
    if (file.role === 'shipment' && !c.amount) {
      return ['该文件没有金额列，将不参与金额核对。'];
    }
    return [];
  }

  function isBlocking(file) {
    if (fileIssues(file).length > 0) return true;
    return !!(file.ui && file.ui.loading);
  }

  function hasBlockingIssues() {
    var blocked = false;
    state.files.forEach(function (file) {
      if (isBlocking(file)) blocked = true;
    });
    return blocked;
  }

  function updateMappingState() {
    var btn = $('btnCheck');
    var note = $('mappingNote');
    if (!btn || !note) return;
    if (state.busy) return;

    var invalidCount = 0;
    var loadingCount = 0;
    state.files.forEach(function (file) {
      if (fileIssues(file).length > 0) invalidCount++;
      if (file.ui && file.ui.loading) loadingCount++;
    });

    btn.disabled = invalidCount > 0 || loadingCount > 0 || state.files.length === 0;

    if (loadingCount > 0) {
      note.textContent = '正在重新读取表头，请稍候…';
      note.classList.remove('is-ready');
    } else if (invalidCount > 0) {
      note.textContent = '还有 ' + String(invalidCount) + ' 个文件需要确认，标红的卡片处理后可开始比对';
      note.classList.remove('is-ready');
    } else {
      note.textContent = '已确认 ' + String(state.files.length) + ' 个文件的表头与字段映射';
      note.classList.add('is-ready');
    }
  }

  function setCheckBusy(busy) {
    state.busy = busy;
    document.body.classList.toggle('is-busy', busy);

    var spinner = $('checkSpinner');
    if (spinner) spinner.hidden = !busy;
    setText($('checkBtnLabel'), busy ? '正在比对…' : '开始比对');

    var btn = $('btnCheck');
    if (btn) btn.disabled = busy ? true : hasBlockingIssues();
  }

  /* ==================== 映射：重新探测工作表 / 表头行 ==================== */

  function updateCardBusy(file) {
    var ui = file.ui;
    if (!ui || !ui.card) return;

    var loading = !!ui.loading;
    ui.card.classList.toggle('is-loading', loading);

    if (ui.loadingBox) {
      ui.loadingBox.hidden = !loading;
      if (loading) ui.loadingBox.textContent = ui.loadingText || '正在重新读取表头…';
    }

    (ui.controls || []).forEach(function (control) {
      control.disabled = !!file.error || loading;
    });

    updateMappingState();
  }

  function refreshSheet(file, body, focusKey, extra) {
    var options = extra || {};
    var ui = file.ui;
    if (!ui) return;

    ui.seq = (ui.seq || 0) + 1;
    var seq = ui.seq;
    ui.loading = true;
    updateCardBusy(file);

    request('/api/sheet', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    })
      .then(function (data) {
        if (file.ui !== ui || ui.seq !== seq) return;
        ui.loading = false;
        applySheetInfo(file, data);
        remountCard(file, focusKey);
      })
      .catch(function (err) {
        if (file.ui !== ui || ui.seq !== seq) return;
        ui.loading = false;
        if (typeof options.onError === 'function') options.onError();
        remountCard(file, null);
        toast('「' + file.fileName + '」' + err.message, 'error');
      });
  }

  /* ==================== 映射：渲染 ==================== */

  function columnSelectOptions(file, currentValue) {
    var labels = file.columnLabels || {};
    var keys = Object.keys(labels).filter(function (key) {
      return /^[0-9]+$/.test(key);
    }).sort(function (a, b) {
      return Number(a) - Number(b);
    });

    var options = [];

    var empty = document.createElement('option');
    empty.value = '';
    empty.textContent = '（不使用）';
    options.push(empty);

    keys.forEach(function (key) {
      var opt = document.createElement('option');
      opt.value = key;
      opt.textContent = '第 ' + key + ' 列 · ' + String(labels[key]);
      options.push(opt);
    });

    if (currentValue && keys.indexOf(String(currentValue)) === -1) {
      var extra = document.createElement('option');
      extra.value = String(currentValue);
      extra.textContent = '第 ' + String(currentValue) + ' 列 · （表头未知）';
      options.push(extra);
    }

    return options;
  }

  function fieldTagMeta(key, role) {
    if (key === 'amount') {
      return role === 'statement'
        ? { text: '必填', tone: 'req' }
        : { text: '可留空', tone: 'opt' };
    }
    if (key === 'name' || key === 'code' || key === 'code2') {
      return { text: '三选一', tone: 'alt' };
    }
    if (key === 'publisher') {
      return { text: '可留空', tone: 'opt' };
    }
    // 数量 / 单价：标签文案本身已含「不参与核对」，不再重复标注
    return null;
  }

  function metaLine(file) {
    var parts = [];
    if (file.error) {
      parts.push('解析失败');
    } else {
      parts.push('数据 ' + displayOrDash(file.rowCount) + ' 行');
      parts.push(String(file.sheets.length) + ' 个工作表');
    }
    parts.push(formatBytes(file.size));
    return parts.join(' · ');
  }

  function buildPreview(file, ui) {
    var section = mk('div', 'preview');

    var startRow = normalizeInt(file.previewStartRow, 1, 1000000) || 1;

    var head = mk('div', 'preview__head');
    head.appendChild(mk('span', 'preview__title', '原始数据预览'));
    head.appendChild(mk('span', 'preview__hint', '第 ' + String(startRow) + ' 行起，表头行已高亮'));
    section.appendChild(head);

    var labels = file.columnLabels || {};
    var rows = Array.isArray(file.previewRows) ? file.previewRows : [];

    var maxCols = 0;
    rows.forEach(function (row) {
      if (row && row.length > maxCols) maxCols = row.length;
    });
    Object.keys(labels).forEach(function (key) {
      if (!/^[0-9]+$/.test(key)) return;
      var n = Number(key);
      if (isFinite(n) && n > maxCols) maxCols = n;
    });

    if (!maxCols) {
      var emptyBox = mk('div', 'preview-empty');
      emptyBox.textContent = file.error
        ? '该文件无法解析，没有可预览的数据。'
        : '没有可预览的数据，请检查表头行是否填写正确。';
      section.appendChild(emptyBox);
      return section;
    }

    var wrap = mk('div', 'tablewrap tablewrap--preview');
    var table = mk('table', 'table table--preview');

    var thead = mk('thead');
    var headRow = mk('tr');
    var corner = mk('th', 'col-sticky', '行');
    corner.setAttribute('scope', 'col');
    headRow.appendChild(corner);

    for (var c = 1; c <= maxCols; c++) {
      var th = mk('th');
      th.setAttribute('scope', 'col');
      var label = labels[String(c)];
      th.textContent = label ? String(c) + ' · ' + String(label) : '第 ' + String(c) + ' 列';
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    table.appendChild(thead);

    var tbody = mk('tbody');
    rows.forEach(function (row, index) {
      var excelRow = startRow + index;
      var tr = mk('tr');
      if (file.headerRow && excelRow === file.headerRow) tr.classList.add('is-header-row');

      var rowHead = mk('th', 'col-sticky', String(excelRow));
      rowHead.setAttribute('scope', 'row');
      tr.appendChild(rowHead);

      for (var col = 0; col < maxCols; col++) {
        var cell = mk('td');
        var value = row ? row[col] : '';
        cell.textContent = value === null || value === undefined ? '' : String(value);
        if (cell.textContent) cell.title = cell.textContent;
        tr.appendChild(cell);
      }
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    wrap.appendChild(table);
    section.appendChild(wrap);

    var lastRow = startRow + rows.length - 1;
    if (file.headerRow && (file.headerRow < startRow || file.headerRow > lastRow)) {
      var tip = mk('p', 'hint');
      tip.textContent = '当前表头行（第 ' + String(file.headerRow) + ' 行）不在预览范围内，请确认填写是否正确。';
      section.appendChild(tip);
    }

    ui.previewBody = tbody;
    ui.previewStartRow = startRow;
    return section;
  }

  function buildFileCard(file) {
    var prevUi = file.ui || {};
    var ui = file.ui = {
      card: null,
      previewBody: null,
      previewStartRow: 1,
      controls: [],
      focusMap: {},
      sheetSelect: null,
      headerInput: null,
      colSelects: {},
      sheetHint: null,
      issuesBox: null,
      loadingBox: null,
      loading: !!prevUi.loading,
      loadingText: prevUi.loadingText || '',
      headerTimer: prevUi.headerTimer || null,
      seq: prevUi.seq || 0
    };

    var card = mk('article', 'card file-card');
    card.setAttribute('aria-labelledby', 'fc-' + String(file.fileId));
    ui.card = card;

    var head = mk('div', 'file-card__head');
    var titleBox = mk('div', 'file-card__title');

    var name = mk('h3', 'file-card__name', file.fileName);
    name.id = 'fc-' + String(file.fileId);

    titleBox.appendChild(name);
    titleBox.appendChild(mk('p', 'file-card__meta', metaLine(file)));

    var tag = mk('span', 'tag ' + (file.role === 'statement' ? 'tag--blue' : 'tag--gray'), roleLabel(file.role));
    head.appendChild(titleBox);
    head.appendChild(tag);
    card.appendChild(head);

    var loadingBox = mk('p', 'card-loading', ui.loadingText || '正在重新读取表头…');
    loadingBox.hidden = !ui.loading;
    ui.loadingBox = loadingBox;
    card.appendChild(loadingBox);

    var grid = mk('div', 'field-grid');

    // 工作表
    var sheetField = mk('div', 'field');
    var sheetLabel = mk('label', 'field__label', '工作表');
    sheetLabel.setAttribute('for', 'sheet-' + String(file.fileId));
    var sheetSelect = mk('select', 'control');
    sheetSelect.id = 'sheet-' + String(file.fileId);
    var sheetNames = file.sheets.slice();
    if (file.sheet && sheetNames.indexOf(file.sheet) === -1) sheetNames.push(file.sheet);
    sheetNames.forEach(function (sheetName) {
      var opt = document.createElement('option');
      opt.value = sheetName;
      opt.textContent = sheetName;
      sheetSelect.appendChild(opt);
    });
    sheetSelect.value = file.sheet;
    sheetField.appendChild(sheetLabel);
    sheetField.appendChild(sheetSelect);
    grid.appendChild(sheetField);
    ui.sheetSelect = sheetSelect;
    ui.focusMap.sheet = sheetSelect;

    // 表头行
    var headerField = mk('div', 'field');
    var headerLabel = mk('label', 'field__label', '表头所在行');
    headerLabel.setAttribute('for', 'header-' + String(file.fileId));
    var headerInput = mk('input', 'control');
    headerInput.type = 'number';
    headerInput.min = '1';
    headerInput.max = String(MAX_HEADER_ROW);
    headerInput.step = '1';
    headerInput.inputMode = 'numeric';
    headerInput.id = 'header-' + String(file.fileId);
    headerInput.value = file.headerRow === null || file.headerRow === undefined ? '' : String(file.headerRow);
    headerField.appendChild(headerLabel);
    headerField.appendChild(headerInput);
    grid.appendChild(headerField);
    ui.headerInput = headerInput;
    ui.focusMap.header = headerInput;

    // 7 个列下拉
    FIELD_DEFS.forEach(function (def) {
      var field = mk('div', 'field');
      var tagMeta = fieldTagMeta(def.key, file.role);

      // 数量与单价不参与核对，标签已注明，视觉上再淡一档
      if (def.key === 'qty' || def.key === 'unitPrice') field.classList.add('field--deemph');

      var labelEl = mk('label', 'field__label', fieldLabel(def.key));
      if (tagMeta) {
        labelEl.appendChild(mk('span', 'field__tag field__tag--' + tagMeta.tone, tagMeta.text));
      }

      var selectId = 'col-' + def.key + '-' + String(file.fileId);
      labelEl.setAttribute('for', selectId);

      var select = mk('select', 'control');
      select.id = selectId;
      columnSelectOptions(file, file.columns[def.key]).forEach(function (opt) {
        select.appendChild(opt);
      });
      select.value = file.columns[def.key] ? String(file.columns[def.key]) : '';

      select.addEventListener('change', function () {
        file.columns[def.key] = select.value ? Number(select.value) : null;
        cacheSheetInfo(file);
        updateCardState(file);
      });

      ui.colSelects[def.key] = select;
      ui.focusMap['col:' + def.key] = select;
      ui.controls.push(select);

      field.appendChild(labelEl);
      field.appendChild(select);
      grid.appendChild(field);
    });

    ui.controls.push(sheetSelect);
    ui.controls.push(headerInput);

    card.appendChild(grid);

    // 切换工作表后的提示
    var sheetHint = mk('p', 'inline-hint', '已切换到该工作表，列映射按该表的表头自动刷新；如识别有偏差请手动调整。');
    sheetHint.hidden = true;
    ui.sheetHint = sheetHint;
    card.appendChild(sheetHint);

    // 解析失败说明
    if (file.error) {
      var errorBox = mk('p', 'issues');
      errorBox.textContent = '该文件解析失败：' + file.error + '。请返回上传页移除后重新上传。';
      card.appendChild(errorBox);
    }

    // 文件级 warnings
    if (file.warnings && file.warnings.length) {
      var warnList = mk('ul', 'warns');
      file.warnings.forEach(function (text) {
        warnList.appendChild(mk('li', '', text));
      });
      card.appendChild(warnList);
    }

    // 非阻断提示（如发货单缺金额列）
    var noteBox = mk('ul', 'fieldnotes');
    noteBox.hidden = true;
    ui.noteBox = noteBox;
    card.appendChild(noteBox);

    // 校验问题
    var issuesBox = mk('ul', 'issues');
    issuesBox.hidden = true;
    ui.issuesBox = issuesBox;
    card.appendChild(issuesBox);

    // 预览
    card.appendChild(buildPreview(file, ui));

    // ---- 事件 ----

    sheetSelect.addEventListener('change', function () {
      if (file.error) return;
      var next = sheetSelect.value;
      if (!next || next === file.sheet) return;

      var previous = file.sheet;
      var cached = findSheetInfo(file, next);
      file.sheet = next;

      if (cached) {
        applySheetInfo(file, cached);
        sheetHint.hidden = false;
        remountCard(file, 'sheet');
        return;
      }

      sheetHint.hidden = true;
      ui.loadingText = '正在读取「' + next + '」的表头…';
      refreshSheet(file, {
        uploadId: state.uploadId,
        fileId: file.fileId,
        sheet: next
      }, 'sheet', {
        onError: function () { file.sheet = previous; }
      });
    });

    headerInput.addEventListener('input', function () {
      if (file.error) return;

      var value = normalizeInt(headerInput.value, 1, MAX_HEADER_ROW);
      file.headerRow = value;
      highlightPreview(file);
      updateCardState(file);

      if (ui.headerTimer) {
        clearTimeout(ui.headerTimer);
        ui.headerTimer = null;
      }
      if (value === null) return;

      var sheetAtRequest = file.sheet;
      ui.headerTimer = setTimeout(function () {
        ui.headerTimer = null;
        if (file.sheet !== sheetAtRequest) return;
        if (ui.loading) return;

        ui.loadingText = '正在按第 ' + String(value) + ' 行重新读取表头…';
        refreshSheet(file, {
          uploadId: state.uploadId,
          fileId: file.fileId,
          sheet: file.sheet,
          headerRow: value
        }, 'header');
      }, HEADER_REFRESH_DELAY);
    });

    if (file.error) {
      ui.controls.forEach(function (control) {
        control.disabled = true;
      });
    }

    updateCardBusy(file);
    updateCardState(file);
    return card;
  }

  function highlightPreview(file) {
    var ui = file.ui;
    if (!ui || !ui.previewBody) return;

    var startRow = ui.previewStartRow || 1;
    Array.prototype.forEach.call(ui.previewBody.children, function (tr, index) {
      tr.classList.toggle('is-header-row', !!file.headerRow && startRow + index === file.headerRow);
    });
  }

  function updateCardState(file) {
    var ui = file.ui;
    if (!ui || !ui.card) return;

    var issues = fileIssues(file);
    ui.card.classList.toggle('is-invalid', issues.length > 0 && !file.error);
    ui.card.classList.toggle('is-error', !!file.error);

    var noteBox = ui.noteBox;
    if (noteBox) {
      clear(noteBox);
      var notes = fileNotes(file);
      if (notes.length) {
        notes.forEach(function (text) {
          noteBox.appendChild(mk('li', '', text));
        });
        noteBox.hidden = false;
      } else {
        noteBox.hidden = true;
      }
    }

    var box = ui.issuesBox;
    if (box) {
      clear(box);
      if (issues.length && !file.error) {
        issues.forEach(function (text) {
          box.appendChild(mk('li', '', text));
        });
        box.hidden = false;
      } else {
        box.hidden = true;
      }
    }

    updateMappingState();
  }

  function remountCard(file, focusKey) {
    var old = file.ui ? file.ui.card : null;
    var fresh = buildFileCard(file);

    if (old && old.parentNode && typeof old.parentNode.replaceChild === 'function') {
      old.parentNode.replaceChild(fresh, old);
    } else {
      renderFileCards();
      return;
    }

    if (focusKey) {
      var target = file.ui.focusMap ? file.ui.focusMap[focusKey] : null;
      if (target && typeof target.focus === 'function') {
        try { target.focus(); } catch (e) { /* 忽略 */ }
      }
    }
  }

  function renderServerWarnings() {
    var host = $('mappingWarnings');
    if (!host) return;
    clear(host);

    var list = state.serverWarnings || [];
    if (!list.length) return;

    var box = mk('div', 'notice notice--warn');
    var body = mk('div', 'notice__list');
    body.appendChild(mk('p', 'notice__title', '服务端解析提示'));
    list.forEach(function (text) {
      body.appendChild(mk('li', '', text));
    });
    box.appendChild(body);
    host.appendChild(box);
  }

  function renderFileCards() {
    var host = $('fileCards');
    if (!host) return;
    clear(host);

    var ordered = state.files.slice().sort(function (a, b) {
      if (a.role === b.role) return 0;
      return a.role === 'statement' ? -1 : 1;
    });

    ordered.forEach(function (file) {
      host.appendChild(buildFileCard(file));
    });

    updateMappingState();
  }

  function enterMapping() {
    hideMappingError();
    renderServerWarnings();
    renderFileCards();
    showView('mapping');
  }

  /* ==================== 比对 ==================== */

  function buildCheckPayload() {
    return {
      uploadId: state.uploadId,
      files: state.files.map(function (file) {
        var columns = {};
        FIELD_DEFS.forEach(function (def) {
          columns[def.key] = file.columns[def.key] ? file.columns[def.key] : null;
        });
        return {
          fileId: file.fileId,
          sheet: file.sheet,
          headerRow: file.headerRow,
          columns: columns
        };
      })
    };
  }

  function startCheck() {
    if (state.busy) return;

    if (hasBlockingIssues()) {
      showMappingError('还有文件未确认完整，请先处理标红的卡片。');
      return;
    }

    hideMappingError();
    setCheckBusy(true);

    request('/api/check', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(buildCheckPayload())
    })
      .then(function (data) {
        setCheckBusy(false);
        state.result = data;
        state.resultId = data.resultId || null;
        renderResult(data);
        showView('result');
      })
      .catch(function (err) {
        setCheckBusy(false);

        if (err.code === 'UPLOAD_EXPIRED') {
          state.uploadId = null;
          state.lastSignature = null;
          toast('本次上传已过期，请重新选择文件后再核对。', 'error');
          showView('upload');
          return;
        }

        if (err.code === 'MAPPING_MISSING') {
          showMappingError(err.message);
        } else {
          showMappingError('比对失败：' + err.message);
        }
        toast(err.message || '比对失败，请重试。', 'error');
      });
  }

  /* ==================== 结果：渲染 ==================== */

  function statNode(value, label, danger) {
    var box = mk('div', 'stat' + (danger ? ' stat--diff' : ''));
    box.appendChild(mk('span', 'stat__value', value));
    box.appendChild(mk('span', 'stat__label', label));
    return box;
  }

  function renderStats(summary) {
    var host = $('resultStats');
    if (!host) return;
    clear(host);

    host.appendChild(statNode(countText(summary.productCount), '产品（个）', false));
    host.appendChild(statNode(
      countText(summary.amountDiffCount),
      '金额不一致（个）',
      isPositiveCount(summary.amountDiffCount)
    ));
    host.appendChild(statNode(
      countText(summary.unknownAmountCount),
      '金额未知（个）',
      isPositiveCount(summary.unknownAmountCount)
    ));
  }

  function renderMatchBar(summary) {
    var host = $('matchBar');
    if (!host) return;
    clear(host);

    var stats = summary && summary.matchStats;
    if (!stats || typeof stats !== 'object') {
      host.hidden = true;
      return;
    }

    host.appendChild(mk('span', 'matchbar__label', '匹配依据'));

    var shown = 0;
    MATCH_STAT_DEFS.forEach(function (item) {
      var value = stats[item.key];
      if (value === undefined || value === null) return;
      shown++;

      var chip = mk('span', 'chip chip--' + item.key);
      if (!isPositiveCount(value)) chip.classList.add('chip--zero');
      chip.appendChild(mk('span', 'chip__k', item.label));
      chip.appendChild(mk('span', 'chip__v', countText(value)));
      host.appendChild(chip);
    });

    host.hidden = shown === 0;
  }

  function numCell(value) {
    var td = mk('td', 'num');
    td.textContent = moneyText(value);
    if (isNegativeText(value)) td.classList.add('diff-neg');
    return td;
  }

  /** 发货单侧金额：amountKnown 为 false 时必须是「—」，不得显示 0.00 */
  function amountCell(value, amountKnown) {
    if (amountKnown === false) {
      return mk('td', 'num num--unknown', '—');
    }
    return numCell(value);
  }

  function noteText(note) {
    if (Array.isArray(note)) return note.join('；');
    if (note === null || note === undefined) return '';
    return String(note);
  }

  function buildSources(sources) {
    if (!Array.isArray(sources) || !sources.length) return null;

    var details = mk('details', 'srcs');
    details.appendChild(mk('summary', 'srcs__head', '来源 ' + String(sources.length) + ' 处'));

    var ul = mk('ul', 'srcs__list');
    sources.forEach(function (src) {
      var line = String((src && src.fileName) || '未知文件');
      var rows = src && Array.isArray(src.rows) ? src.rows : [];
      if (rows.length) {
        line += ' · 第 ' + rows.map(String).join('、') + ' 行';
      }
      ul.appendChild(mk('li', '', line));
    });
    details.appendChild(ul);
    return details;
  }

  function matchByTag(matchBy) {
    var meta = MATCH_BY_META[matchBy];
    if (!meta) {
      return mk('span', 'mchip mchip--single', '单边');
    }
    return mk('span', 'mchip mchip--' + meta.tone, meta.label);
  }

  function buildResultTable(rows, tone) {
    var wrap = mk('div', 'tablewrap tablewrap--result');
    var table = mk('table', 'table table--result');

    var headers = [
      ['产品编码', 'col-sticky'],
      ['产品名称', ''],
      ['编码2', ''],
      ['对账单金额', 'num'],
      ['发货单金额', 'num'],
      ['金额差', 'num'],
      ['匹配依据', ''],
      ['备注', '']
    ];

    var thead = mk('thead');
    var headRow = mk('tr');
    headers.forEach(function (item) {
      var th = mk('th', item[1], item[0]);
      th.setAttribute('scope', 'col');
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);

    var tbody = mk('tbody');
    rows.forEach(function (row) {
      var tr = mk('tr');

      var codeText = displayOrDash(row.code);
      var codeCell = mk('th', 'col-sticky cell-code', codeText);
      codeCell.setAttribute('scope', 'row');
      if (codeText !== '—') codeCell.title = codeText;
      tr.appendChild(codeCell);

      var nameText = displayOrDash(row.displayName || row.key);
      var nameCell = mk('td', 'cell-name', nameText);
      if (nameText !== '—') nameCell.title = nameText;
      tr.appendChild(nameCell);

      tr.appendChild(mk('td', '', displayOrDash(row.code2)));

      // 对账单侧金额总是已知
      tr.appendChild(numCell(row.statementAmount));
      // 发货单侧金额与差额：amountKnown=false 时显示「—」
      tr.appendChild(amountCell(row.shipmentAmount, row.amountKnown));
      tr.appendChild(amountCell(row.amountDiff, row.amountKnown));

      var matchCell = mk('td');
      matchCell.appendChild(matchByTag(row.matchBy));
      tr.appendChild(matchCell);

      var noteCell = mk('td', 'cell-note');
      noteCell.appendChild(mk('p', 'note-text', noteText(row.note) || '—'));
      var sources = buildSources(row.sources);
      if (sources) noteCell.appendChild(sources);
      tr.appendChild(noteCell);

      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrap.appendChild(table);
    return wrap;
  }

  function buildGroup(meta, rows) {
    var body;
    var wrapper;

    if (meta.collapsible) {
      wrapper = mk('details', 'group');
      var summary = mk('summary', 'group__head group__head--toggle');

      var dot = mk('span', 'dot dot--' + meta.tone);
      dot.setAttribute('aria-hidden', 'true');
      summary.appendChild(dot);
      summary.appendChild(mk('h3', 'group__title', meta.title));
      summary.appendChild(mk('span', 'group__count', String(rows.length) + ' 条'));
      summary.appendChild(mk('span', 'group__toggle', ''));

      wrapper.appendChild(summary);
      wrapper.appendChild(mk('p', 'group__desc', meta.desc));

      body = mk('div', 'group__body group__body--flush');
      if (rows.length) body.appendChild(buildResultTable(rows, meta.tone));
      else body.appendChild(mk('p', 'group__empty', '本组无记录。'));
      wrapper.appendChild(body);
      return wrapper;
    }

    wrapper = mk('div', 'group');
    var head = mk('div', 'group__head');
    var marker = mk('span', 'dot dot--' + meta.tone);
    marker.setAttribute('aria-hidden', 'true');
    head.appendChild(marker);
    head.appendChild(mk('h3', 'group__title', meta.title));
    head.appendChild(mk('span', 'group__count', String(rows.length) + ' 条'));
    wrapper.appendChild(head);
    wrapper.appendChild(mk('p', 'group__desc', meta.desc));

    if (rows.length) {
      body = mk('div', 'group__body group__body--flush');
      body.appendChild(buildResultTable(rows, meta.tone));
      wrapper.appendChild(body);
    } else {
      wrapper.appendChild(mk('p', 'group__empty', '本组无记录。'));
    }

    return wrapper;
  }

  function renderGroups(rows) {
    var host = $('resultGroups');
    if (!host) return;
    clear(host);

    var buckets = {
      amount_diff: [],
      statement_only: [],
      shipment_only: [],
      unknown_amount: [],
      match: []
    };
    (rows || []).forEach(function (row) {
      var key = row && buckets[row.category] ? row.category : 'amount_diff';
      buckets[key].push(row);
    });

    GROUP_META.forEach(function (meta) {
      host.appendChild(buildGroup(meta, buckets[meta.category] || []));
    });
  }

  function buildShipmentBlock(opts) {
    var card = mk('section', 'card block block--files');
    if (opts.muted) card.classList.add('block--muted');
    card.setAttribute('aria-label', opts.title);

    var head = mk('div', 'block__head');
    head.appendChild(mk('h3', 'block__title', opts.title));
    head.appendChild(mk('span', 'block__hint', String(opts.items.length) + ' 张'));
    card.appendChild(head);

    if (!opts.items.length) {
      card.appendChild(mk('p', 'empty', opts.empty || '没有相关单据。'));
      return card;
    }

    var ul = mk('ul', 'filelist--result');
    opts.items.forEach(function (item) {
      var li = mk('li');
      var excluded = item && item.amountChecked === false;
      if (excluded) li.classList.add('is-excluded');

      li.appendChild(mk('span', 'listname', item.fileName || '未命名文件'));

      var parts = [displayOrDash(item.rows) + ' 行'];
      if (excluded) parts.push('未参与核对');
      else parts.push(moneyText(item.amount) + ' 元');

      var metaEl = mk('span', 'listmeta', parts.join(' · '));
      if (item.amountMode) metaEl.title = String(item.amountMode);
      li.appendChild(metaEl);
      ul.appendChild(li);
    });
    card.appendChild(ul);

    if (opts.hint) card.appendChild(mk('p', 'hint', opts.hint));

    return card;
  }

  function buildListNotice(title, items, tone) {
    var box = mk('div', 'notice notice--' + tone);
    var body = mk('div', 'notice__list');
    body.appendChild(mk('p', 'notice__title', title));
    items.forEach(function (text) {
      body.appendChild(mk('li', '', text));
    });
    box.appendChild(body);
    return box;
  }

  function renderNotices(summary, data) {
    var host = $('resultNotices');
    if (!host) return;
    clear(host);

    // 参与 / 未参与金额核对 —— 业务方判断「总额对不上是不是漏传了订书单」的唯一线索
    var files = Array.isArray(summary.shipmentFiles) ? summary.shipmentFiles : [];
    var checked = [];
    var excluded = [];
    files.forEach(function (item) {
      if (!item) return;
      if (item.amountChecked === false) excluded.push(item);
      else checked.push(item);
    });

    host.appendChild(buildShipmentBlock({
      title: '参与金额核对',
      items: checked,
      muted: false,
      empty: '没有发货单参与金额核对，请检查是否漏传了本次要核对的单据。',
      hint: '请确认没有漏传或多传，归属错误系统无法自动判断。'
    }));

    if (excluded.length) {
      host.appendChild(buildShipmentBlock({
        title: '未参与金额核对',
        items: excluded,
        muted: true,
        empty: '',
        hint: '这些单据没有金额列，其产品在明细中记为「金额未知」，无法判断总额差异是否由它们造成。'
      }));
    }

    var merged = Array.isArray(summary.nameMerged) ? summary.nameMerged : [];
    if (merged.length) {
      var mergedLines = merged.map(function (item) {
        var aliases = Array.isArray(item.aliases) ? item.aliases : [];
        var shown = aliases.slice(0, 4);
        var text = String(item.key || '') + '：' + shown.map(String).join(' / ');
        if (aliases.length > shown.length) {
          text += ' 等 ' + String(aliases.length) + ' 种写法';
        }
        return text;
      });
      host.appendChild(buildListNotice(
        '名称标准化合并 ' + String(merged.length) + ' 组',
        mergedLines,
        'info'
      ));
    }

    // 契约里 suspects 是响应顶层字段，这里兼容放在 summary 的实现
    var suspects = Array.isArray(data && data.suspects)
      ? data.suspects
      : (Array.isArray(summary.suspects) ? summary.suspects : []);
    if (suspects.length) {
      var suspectLines = suspects.map(function (item) {
        var reason = item.reason ? String(item.reason) : '名称高度相似';
        return String(item.left || '') + ' ↔ ' + String(item.right || '') + '：' + reason;
      });
      host.appendChild(buildListNotice(
        '疑似同一产品 ' + String(suspects.length) + ' 组',
        suspectLines,
        'warn'
      ));
    }

    var warnings = Array.isArray(summary.warnings) ? summary.warnings : [];
    if (warnings.length) {
      host.appendChild(buildListNotice('解析提示', warnings.map(String), 'warn'));
    }
  }

  /** 纯字符串判断金额是否为 0（"0.00" / "-0.00" / "+0" 都算 0） */
  function isZeroAmount(value) {
    var text = String(value === null || value === undefined ? '' : value).replace(/^[+-]/, '');
    return !/[1-9]/.test(text);
  }

  /** 总额卡片：结果页第一视觉层级，结论只取 summary.total */
  function renderTotal(summary) {
    var total = summary && summary.total;
    var card = $('totalCard');
    if (!card) return;

    if (!total || typeof total !== 'object') {
      card.hidden = true;
      return;
    }

    setText($('totalStatement'), moneyText(total.statementAmount));
    setText($('totalShipment'), moneyText(total.shipmentAmount));
    setText($('totalDiff'), moneyText(total.amountDiff));

    var noteParts = [];
    if (total.checkedFileCount !== undefined && total.checkedFileCount !== null) {
      noteParts.push('参与核对 ' + countText(total.checkedFileCount) + ' 张');
    }
    if (isPositiveCount(total.excludedFileCount)) {
      noteParts.push('未参与 ' + countText(total.excludedFileCount) + ' 张');
    }
    setText($('totalFilesNote'), noteParts.join(' · '));

    var matched = total.matched === true;
    var diffRaw = total.amountDiff;
    var diffKnown = diffRaw !== null && diffRaw !== undefined && String(diffRaw).trim() !== '';
    var diffZero = diffKnown && isZeroAmount(diffRaw);

    // 差额配色必须以 matched 为准：金额刚好相等但仍有单据没参与核对时，
    // 不能因为差额是 0.00 就亮绿灯 —— 那是「部分一致」，不是通过。
    var diffEl = $('totalDiff');
    if (diffEl) {
      diffEl.classList.remove('is-zero', 'is-diff', 'is-warn');
      if (matched) {
        diffEl.classList.add('is-zero');
      } else if (diffKnown && diffZero) {
        diffEl.classList.add('is-warn');
      } else if (diffKnown) {
        diffEl.classList.add('is-diff');
      }
    }

    var conclusion = $('totalConclusion');
    if (conclusion) {
      conclusion.classList.remove('is-warn');
      if (matched) {
        // 一致时由下方「总额一致」大字面板承担，避免同一结论出现两次
        conclusion.hidden = true;
        conclusion.textContent = '';
      } else {
        conclusion.hidden = false;
        conclusion.textContent = total.conclusion ? String(total.conclusion) : '总额不一致';
        // 「已参与部分一致，但另有单据未参与」是待确认，不是错误
        if (diffKnown && diffZero) conclusion.classList.add('is-warn');
      }
    }

    card.hidden = false;
  }

  function renderResult(data) {
    var summary = data && data.summary ? data.summary : {};
    var total = summary.total && typeof summary.total === 'object' ? summary.total : null;
    var statement = summary.statement && typeof summary.statement === 'object' ? summary.statement : {};
    var matched = !!(total && total.matched === true);

    var head = '对账单：' + displayOrDash(statement.fileName);
    var extras = [];
    if (statement.rows !== undefined && statement.rows !== null) {
      extras.push(displayOrDash(statement.rows) + ' 行');
    }
    if (statement.amountMode) extras.push('取「' + String(statement.amountMode) + '」');
    if (extras.length) head += '（' + extras.join('，') + '）';
    setText($('resultHeadline'), head);

    renderTotal(summary);

    var identicalPanel = $('identicalPanel');
    if (identicalPanel) identicalPanel.hidden = !matched;

    var actions = $('resultActions');
    if (actions) actions.hidden = matched || !state.resultId;

    var download = $('btnDownload');
    if (download && state.resultId) {
      download.setAttribute('href', '/api/export/' + String(state.resultId));
    }

    renderStats(summary);
    renderMatchBar(summary);
    renderNotices(summary, data);
    renderGroups(data && Array.isArray(data.rows) ? data.rows : []);
  }

  /* ==================== 结果：下载 / 重置 ==================== */

  function onDownload(event) {
    if (!state.resultId) {
      event.preventDefault();
      toast('结果已失效，请重新核对。', 'error');
      return;
    }

    event.preventDefault();
    var url = '/api/export/' + String(state.resultId);

    fetch(url, { method: 'HEAD', credentials: 'same-origin', cache: 'no-store' })
      .then(function (res) {
        if (res.ok || res.status === 405 || res.status === 501) {
          window.location.href = url;
          return null;
        }
        return res.json().catch(function () { return null; }).then(function (data) {
          var message = data && data.error && data.error.message
            ? data.error.message
            : '下载失败，请稍后重试。';
          toast(message, 'error');
          return null;
        });
      })
      .catch(function () {
        window.location.href = url;
      });
  }

  function resetToUpload() {
    state.result = null;
    state.resultId = null;

    var actions = $('resultActions');
    if (actions) actions.hidden = true;
    var identicalPanel = $('identicalPanel');
    if (identicalPanel) identicalPanel.hidden = true;
    var matchBar = $('matchBar');
    if (matchBar) { matchBar.hidden = true; clear(matchBar); }

    var totalCard = $('totalCard');
    if (totalCard) totalCard.hidden = true;
    setText($('totalStatement'), '—');
    setText($('totalShipment'), '—');
    setText($('totalDiff'), '—');
    var diffReset = $('totalDiff');
    if (diffReset) diffReset.classList.remove('is-zero', 'is-diff', 'is-warn');
    var conclusionReset = $('totalConclusion');
    if (conclusionReset) conclusionReset.classList.remove('is-warn');
    setText($('totalFilesNote'), '');
    var totalConclusion = $('totalConclusion');
    if (totalConclusion) {
      totalConclusion.hidden = true;
      totalConclusion.textContent = '';
    }

    clear($('resultGroups'));
    clear($('resultNotices'));
    clear($('resultStats'));
    setText($('resultHeadline'), '');

    showView('upload');
  }

  /* ==================== 事件绑定 ==================== */

  function bindDropzone(dropEl, inputEl, handler) {
    ['dragenter', 'dragover'].forEach(function (type) {
      dropEl.addEventListener(type, function (event) {
        event.preventDefault();
        event.stopPropagation();
        if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy';
        dropEl.classList.add('is-drag');
      });
    });

    ['dragleave', 'dragend'].forEach(function (type) {
      dropEl.addEventListener(type, function (event) {
        event.preventDefault();
        dropEl.classList.remove('is-drag');
      });
    });

    dropEl.addEventListener('drop', function (event) {
      event.preventDefault();
      event.stopPropagation();
      dropEl.classList.remove('is-drag');
      if (state.busy) return;
      var files = event.dataTransfer ? event.dataTransfer.files : null;
      if (files && files.length) handler(files);
    });

    dropEl.addEventListener('click', function (event) {
      if (event.target && event.target.closest && event.target.closest('.btn')) return;
      if (state.busy) return;
      inputEl.click();
    });

    inputEl.addEventListener('change', function () {
      var files = inputEl.files;
      if (files && files.length) handler(files);
      inputEl.value = '';
    });
  }

  function bindUploadEvents() {
    var statementDrop = $('statementDrop');
    var statementInput = $('statementInput');
    var shipmentDrop = $('shipmentDrop');
    var shipmentInput = $('shipmentInput');

    bindDropzone(statementDrop, statementInput, handleStatementFiles);
    bindDropzone(shipmentDrop, shipmentInput, handleShipmentFiles);

    $('statementPick').addEventListener('click', function (event) {
      event.stopPropagation();
      if (state.busy) return;
      statementInput.click();
    });

    $('shipmentPick').addEventListener('click', function (event) {
      event.stopPropagation();
      if (state.busy) return;
      shipmentInput.click();
    });

    $('statementList').addEventListener('click', function (event) {
      var btn = event.target.closest ? event.target.closest('[data-action="remove-statement"]') : null;
      if (!btn || state.busy) return;
      state.statementFile = null;
      renderPicked();
    });

    $('shipmentList').addEventListener('click', function (event) {
      var btn = event.target.closest ? event.target.closest('[data-action="remove-shipment"]') : null;
      if (!btn || state.busy) return;
      var index = Number(btn.getAttribute('data-index'));
      if (!isFinite(index) || index < 0 || index >= state.shipmentFiles.length) return;
      state.shipmentFiles.splice(index, 1);
      renderPicked();
    });

    $('btnUpload').addEventListener('click', startUpload);
  }

  function bindMappingEvents() {
    $('btnBackUpload').addEventListener('click', function () {
      if (state.busy) return;
      showView('upload');
    });

    $('mappingErrorClose').addEventListener('click', hideMappingError);
    $('btnCheck').addEventListener('click', startCheck);
  }

  function bindResultEvents() {
    $('btnDownload').addEventListener('click', onDownload);
    $('btnReset').addEventListener('click', resetToUpload);
  }

  function bindNoticeEvents() {
    $('authNoticeClose').addEventListener('click', hideAuthNotice);
  }

  /* ==================== 启动 / 免登 ==================== */

  function loadConfig() {
    return fetch('/api/config', { credentials: 'same-origin', cache: 'no-store' })
      .then(function (res) {
        return res.json().catch(function () { return null; });
      })
      .then(function (data) {
        if (!data || data.ok !== true) throw new Error('配置不可用');

        state.config = data;
        state.mode = data.mode === 'dingtalk' ? 'dingtalk' : 'local';

        if (data.limits && typeof data.limits === 'object') {
          var maxMB = Number(data.limits.maxFileSizeMB);
          var maxShipments = Number(data.limits.maxShipments);
          if (isFinite(maxMB) && maxMB > 0) state.limits.maxFileSizeMB = maxMB;
          if (isFinite(maxShipments) && maxShipments > 0) state.limits.maxShipments = maxShipments;
        }

        if (Array.isArray(data.acceptExt) && data.acceptExt.length) {
          state.acceptExt = data.acceptExt.map(function (ext) { return String(ext); });
        }

        if (data.fieldLabels && typeof data.fieldLabels === 'object') {
          var merged = Object.assign({}, DEFAULT_FIELD_LABELS);
          FIELD_DEFS.forEach(function (def) {
            var label = data.fieldLabels[def.key];
            if (label) merged[def.key] = String(label);
          });
          state.fieldLabels = merged;
        }

        setUser(data.user, state.mode);
        renderPicked();
        updateUploadReady();
        return data;
      })
      .catch(function () {
        state.mode = 'local';
        setUser({ id: 'local', name: '本地模式' }, 'local');
        renderPicked();
        updateUploadReady();
        return null;
      });
  }

  function ensureDingTalkAuth() {
    var dd = window.dd;

    if (!dd || !dd.runtime || !dd.runtime.permission ||
        typeof dd.runtime.permission.requestAuthCode !== 'function') {
      showAuthNotice('当前不在钉钉容器内，钉钉免登不可用，已按本地模式继续使用，核对功能不受影响。');
      return;
    }

    var settled = false;
    var timer = setTimeout(function () {
      if (settled) return;
      settled = true;
      showAuthNotice('钉钉免登响应超时，已按本地模式继续使用，核对功能不受影响。');
    }, 6000);

    function finish(message) {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (message) showAuthNotice(message);
    }

    try {
      dd.runtime.permission.requestAuthCode({
        corpId: (state.config && state.config.corpId) || '',
        onSuccess: function (result) {
          if (settled) return;
          var code = result && result.code;
          if (!code) {
            finish('钉钉免登未返回授权码，已按本地模式继续使用，核对功能不受影响。');
            return;
          }
          request('/api/auth/dingtalk', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ code: code })
          })
            .then(function (data) {
              finish('');
              setUser(data.user, 'dingtalk');
            })
            .catch(function (err) {
              finish('钉钉免登失败：' + err.message + '。已按本地模式继续使用，核对功能不受影响。');
            });
        },
        onFail: function (err) {
          var reason = err && err.errorMessage ? '（' + err.errorMessage + '）' : '';
          finish('钉钉免登未完成' + reason + '，已按本地模式继续使用，核对功能不受影响。');
        }
      });
    } catch (e) {
      clearTimeout(timer);
      settled = true;
      showAuthNotice('钉钉免登调用异常，已按本地模式继续使用，核对功能不受影响。');
    }
  }

  function boot() {
    bindNoticeEvents();
    bindUploadEvents();
    bindMappingEvents();
    bindResultEvents();

    renderPicked();
    showView('upload', { focus: false });

    document.addEventListener('dragover', function (event) { event.preventDefault(); });
    document.addEventListener('drop', function (event) { event.preventDefault(); });

    loadConfig().then(function (data) {
      if (data && state.mode === 'dingtalk') ensureDingTalkAuth();
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();
