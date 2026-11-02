/**
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

define("values", function (module, exports, require, Buffer) {
'use strict';

/**
 * 数值归一化：金额一律换算成「分」的 BigInt，数量换算成固定 6 位小数的 BigInt。
 *
 * 为什么不用浮点：
 *   0.1 + 0.2 !== 0.3，多张发货单累加后会出现 1e-10 级的假差异，
 *   而需求要求「差 1 分钱也算不一致」，所以比较与累加全部走整数。
 *   真实样例里对账单的「应收增加」就是 1125.94199999999 这种浮点尾巴。
 */

const QTY_SCALE = 6; // 数量保留 6 位小数（吨/公斤换算足够）
const QTY_SCALE_N = 10n ** BigInt(QTY_SCALE);
const UNIT_RE = /^(件|个|箱|托|台|套|只|支|张|包|袋|桶|卷|米|m|km|kg|KG|Kg|公斤|吨|t|T|克|g|本|册|盒|批)$/;

/** 把表格单元格值统一转成文本（用于展示与文本解析） */
function toText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return String(v);
  if (typeof v === 'boolean') return String(v);
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((p) => p.text || '').join('');
    if ('result' in v) return toText(v.result);
    if ('text' in v) return toText(v.text);
    if ('error' in v) return String(v.error);
  }
  return String(v);
}

/** 是否「空」（只剩破折号占位也算空） */
function isBlankText(s) {
  const t = String(s ?? '').trim();
  return t === '' || t === '-' || t === '--' || t === '—' || t === '/' || t === '-- ';
}

function digitCount(s) {
  let n = 0;
  for (const ch of s) if (ch >= '0' && ch <= '9') n++;
  return n;
}

/**
 * 解析金额 → 分。
 * 支持：1200、1,200.5、¥1,200.00、1 200.00、（1,200.00）负数、-1200、1200元、1200-
 * 返回 { ok:true, cents:BigInt, blank, warnings, raw } | { ok:false, error }
 */
function parseMoney(raw) {
  const text = toText(raw).trim();
  if (isBlankText(text)) return { ok: true, cents: 0n, blank: true, warnings: [], raw: text };

  let s = text.normalize('NFKC');
  // 欧洲格式 "1.234,56" 明确报错，不静默解析错
  if (/,\d{1,2}$/.test(s) && s.includes('.')) {
    return { ok: false, error: `无法识别的金额「${text}」（疑似千分位/小数点格式混用）` };
  }
  s = s.replace(/[¥￥$€£]/g, '').replace(/元$/, '').replace(/[Rr][Mm][Bb]$/, '');
  s = s.replace(/\s+/g, '').replace(/,/g, '');

  let neg = false;
  const paren = /^\((.*)\)$/.exec(s);
  if (paren) {
    neg = true;
    s = paren[1];
  }
  if (/-$/.test(s) && digitCount(s) > 0) {
    neg = true;
    s = s.replace(/-$/, '');
  }
  s = s.replace(/^[—–−‒]/, '-').replace(/^\+/, '');
  if (/^-/.test(s)) {
    neg = !neg;
    s = s.slice(1);
  }

  const m = /^(\d*)(?:\.(\d*))?$/.exec(s);
  if (!m || (m[1] === '' && !m[2])) {
    return { ok: false, error: `无法识别的金额「${text}」` };
  }

  const intPart = m[1] || '0';
  let dec = m[2] || '';
  const warnings = [];
  if (dec.length > 2) {
    const extra = dec.slice(2);
    dec = dec.slice(0, 2);
    if (extra[0] >= '5') dec = String(BigInt(dec || '0') + 1n).padStart(2, '0');
    warnings.push(`金额「${text}」超过 2 位小数，已四舍五入为 ${neg ? '-' : ''}${intPart}.${dec}`);
    if (dec.length > 2) {
      // 99 + 1 = 100 的进位
      const cents = (BigInt(intPart) * 100n + BigInt(dec)) * (neg ? -1n : 1n);
      return { ok: true, cents, blank: false, warnings, raw: text };
    }
  } else {
    dec = dec.padEnd(2, '0');
  }

  const cents = (BigInt(intPart) * 100n + BigInt(dec)) * (neg ? -1n : 1n);
  return { ok: true, cents, blank: false, warnings, raw: text };
}

/**
 * 解析数量 → 6 位小数定标 BigInt。
 * 支持：10、10.5、1,000、10件、10 件、-3
 */
function parseQty(raw) {
  const text = toText(raw).trim();
  if (isBlankText(text)) return { ok: true, scaled: 0n, blank: true, warnings: [], raw: text };

  let s = text.normalize('NFKC').replace(/\s+/g, '').replace(/,/g, '');

  let neg = false;
  const paren = /^\((.*)\)$/.exec(s);
  if (paren) {
    neg = true;
    s = paren[1];
  }
  s = s.replace(/^[—–−‒]/, '-').replace(/^\+/, '');
  if (/^-/.test(s)) {
    neg = !neg;
    s = s.slice(1);
  }

  // 先按纯数字解析；失败再尝试去掉尾部单位（件/箱/吨…）。
  // 注意 "10." 是 Excel 数字格式的尾巴，属于纯数字，不能被当成单位。
  let m = /^(\d*)(?:\.(\d*))?$/.exec(s);
  if (!m) {
    const unitMatch = /^(\d+(?:\.\d*)?|\.\d+)([A-Za-z\u4e00-\u9fa5]+)$/.exec(s);
    if (!unitMatch || !UNIT_RE.test(unitMatch[2])) return { ok: false, error: `无法识别的数量「${text}」` };
    s = unitMatch[1];
    m = /^(\d*)(?:\.(\d*))?$/.exec(s);
  }
  if (!m || (m[1] === '' && !m[2])) {
    return { ok: false, error: `无法识别的数量「${text}」` };
  }

  const intPart = m[1] || '0';
  let dec = m[2] || '';
  const warnings = [];
  if (dec.length > QTY_SCALE) {
    const rest = dec.slice(QTY_SCALE);
    dec = dec.slice(0, QTY_SCALE);
    if (rest[0] >= '5') dec = String(BigInt(dec || '0') + 1n).padStart(QTY_SCALE, '0');
    warnings.push(`数量「${text}」超过 ${QTY_SCALE} 位小数，已四舍五入`);
  }
  const scaled = (BigInt(intPart) * QTY_SCALE_N + BigInt(dec.padEnd(QTY_SCALE, '0'))) * (neg ? -1n : 1n);
  return { ok: true, scaled, blank: false, warnings, raw: text };
}

/** 单价（分）× 数量（6 位定标）→ 金额（分），半进位，全整数运算 */
function mulPriceByQty(priceCents, qtyScaled) {
  const neg = (priceCents < 0n) !== (qtyScaled < 0n);
  const p = priceCents < 0n ? -priceCents : priceCents;
  const q = qtyScaled < 0n ? -qtyScaled : qtyScaled;
  let cents = (p * q + QTY_SCALE_N / 2n) / QTY_SCALE_N;
  return neg ? -cents : cents;
}

/** 分 → "1234.56"（固定 2 位小数） */
function formatCents(cents) {
  const neg = cents < 0n;
  const abs = neg ? -cents : cents;
  return `${neg ? '-' : ''}${abs / 100n}.${String(abs % 100n).padStart(2, '0')}`;
}

/** 6 位定标数量 → "10" / "10.5"（去尾零，整数不带小数点） */
function formatQty(scaled) {
  const neg = scaled < 0n;
  const abs = neg ? -scaled : scaled;
  const int = abs / QTY_SCALE_N;
  const dec = String(abs % QTY_SCALE_N).padStart(QTY_SCALE, '0').replace(/0+$/, '');
  return `${neg ? '-' : ''}${int}${dec ? '.' + dec : ''}`;
}

module.exports = {
  QTY_SCALE,
  QTY_SCALE_N,
  toText,
  isBlankText,
  parseMoney,
  parseQty,
  mulPriceByQty,
  formatCents,
  formatQty,
};

});

define("text", function (module, exports, require, Buffer) {
'use strict';

/**
 * 文本标准化与相似度。
 *
 * 真实样例里的写法差异：
 *   「Ａ产品」 vs 「A产品」        → NFKC 统一全半角
 *   「XX-100」 vs 「XX 100」       → 去空白、统一破折号
 *   「钢琴基础（第一册）」 vs「钢琴基础(第一册)」 → NFKC
 *   「湖南师范大学」 vs「湖南师范大学出版社」 → 出版社按"包含"判定相容
 */

const ZERO_WIDTH = /[\u200b-\u200f\u2028\u2029\ufeff]/g;
const DASHES = /[—–―‒−]/g;
const SPACES = /[\s\u00a0\u3000]+/g;

/** 名称/文本标准化：用于匹配，不用于展示 */
function normalizeText(input) {
  if (input === null || input === undefined) return '';
  let t = String(input).normalize('NFKC');
  t = t.replace(ZERO_WIDTH, '');
  t = t.replace(DASHES, '-');
  t = t.replace(SPACES, '');
  t = t.toLowerCase();
  // 去掉首尾的标点，保留中间的 "-" "/" "."（型号分隔有意义）
  t = t.replace(/^[：:；;、。．.,·・|]+/, '').replace(/[：:；;、。．.,·・|]+$/, '');
  return t;
}

/** 编码标准化：只留字母数字与中文，去空格与符号，统一大写 */
function normalizeCode(input) {
  if (input === null || input === undefined) return '';
  let t = String(input).normalize('NFKC').toUpperCase();
  t = t.replace(ZERO_WIDTH, '');
  t = t.replace(/[^0-9A-Z\u4e00-\u9fa5]/g, '');
  // Excel 常把长编码存成科学计数法或带 .0
  t = t.replace(/\.0+$/, '');
  return t;
}

/** 名称里去掉括号内容（用于"主体名"比较，如 "婴幼儿教育学（岳曼）" → "婴幼儿教育学"） */
function stripBracketContent(input) {
  return String(input ?? '')
    .normalize('NFKC')
    .replace(/\([^)]*\)/g, '')
    .replace(/【[^】]*】/g, '')
    .replace(/\[[^\]]*\]/g, '');
}

function bigrams(s) {
  const map = new Map();
  if (s.length === 1) {
    map.set(s, 1);
    return map;
  }
  for (let i = 0; i < s.length - 1; i++) {
    const g = s.slice(i, i + 2);
    map.set(g, (map.get(g) || 0) + 1);
  }
  return map;
}

/** Dice 相似度（0..1），对中文短串稳定 */
function dice(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  const A = bigrams(a);
  const B = bigrams(b);
  let inter = 0;
  let totalA = 0;
  let totalB = 0;
  for (const n of A.values()) totalA += n;
  for (const n of B.values()) totalB += n;
  for (const [g, n] of A) {
    const m = B.get(g);
    if (m) inter += Math.min(n, m);
  }
  return (2 * inter) / (totalA + totalB);
}

/**
 * 名称相似度 0..1：
 *   - 相同 → 1
 *   - 一方包含另一方（去括号后）→ ≥0.92
 *   - 否则 Dice
 */
function similarity(a, b) {
  const x = normalizeText(a);
  const y = normalizeText(b);
  if (!x || !y) return 0;
  if (x === y) return 1;
  const sx = normalizeText(stripBracketContent(x));
  const sy = normalizeText(stripBracketContent(y));
  if (sx && sy && sx === sy) return 0.95;
  if (sx && sy && (sx.includes(sy) || sy.includes(sx))) {
    const short = sx.length <= sy.length ? sx : sy;
    const long = sx.length <= sy.length ? sy : sx;
    // 短串至少 4 个字才认为可靠包含，避免「书」包含一切
    if (short.length >= 4) return 0.92;
  }
  return dice(x, y);
}

/** 出版社相容：全称/简称并存时视为同一家（"湖南师范大学" ⊂ "湖南师范大学出版社"） */
function publisherCompatible(a, b) {
  const x = normalizeText(a);
  const y = normalizeText(b);
  if (!x || !y) return true; // 缺一方就不参与消歧
  if (x === y) return true;
  const short = x.length <= y.length ? x : y;
  const long = x.length <= y.length ? y : x;
  return short.length >= 2 && long.includes(short);
}

module.exports = { normalizeText, normalizeCode, stripBracketContent, similarity, dice, publisherCompatible };

});

define("errors", function (module, exports, require, Buffer) {
'use strict';

/** 带错误码的业务异常，路由层统一转成 JSON 错误响应 */
class AppError extends Error {
  constructor(code, message, detail = null, status = 400) {
    super(message);
    this.name = 'AppError';
    this.code = code;
    this.detail = detail;
    this.status = status;
  }
}

module.exports = { AppError };

});

define("detect", function (module, exports, require, Buffer) {
'use strict';

/**
 * 表头探测 + 字段自动识别。
 *
 * 真实样例的三种表：
 *   对账单   表头第 1 行：书号 / 摘要 / 货号·科目 / 品牌名称 / 数量 / 标准售价1 / 标准金额 / 应收增加
 *   发货单A  表头第 10 行：ISBN号 / 图书名称 / 版别 / 定价 / 数量
 *   发货单B  表头第 3 行：数量 / 书号 / 书名 / 版别 / 价格 / 码洋 / 发行实洋
 * 前 9 行还被客户抬头、物流单号、联系人等占用，所以必须按"同义词命中数"扫描而不是写死行号。
 */

const { normalizeText } = require('text');

const FIELDS = ['name', 'code', 'code2', 'publisher', 'qty', 'amount', 'unitPrice'];

/** 金额列里"实洋"档的权重，只有 ≥ 这个档才允许自动采用 */
const AMOUNT_REAL_TIER = 3;

const FIELD_DEFS = [
  {
    key: 'qty',
    tiers: [{ weight: 1, words: ['数量', '发货数量', '出库数量', '销售数量', '订购数量', '采购数量', '订单数量', '本次发货数量', '数量合计', '件数', '册数', '本数'] }],
    exclude: /单价|金额|价格|代号|编码|编号/,
  },
  {
    // 金额列分三档：v3 口径是「应收增加」↔「发行实洋」，这两个必须压过「标准金额」「码洋」。
    // 真实订书单同时有 码洋 / 发行实洋 / 入账实洋 三列，按位置取第一个会选错。
    key: 'amount',
    tiers: [
      { weight: 3, words: ['应收增加', '发行实洋'] },
      { weight: 2, words: ['标准金额', '码洋', '应收金额', '结算金额', '含税金额', '货款金额', '金额合计', '合计金额', '总金额'] },
      { weight: 1, words: ['入账实洋', '金额', '小计', '合计'] },
    ],
    exclude: /折扣|单价|售价|价格|收款|欠款|余额/,
  },
  {
    key: 'unitPrice',
    tiers: [{ weight: 1, words: ['标准售价1', '标准售价', '销售单价', '结算单价', '零售价', '定价', '价格', '售价', '单价'] }],
    exclude: /折扣|金额|合计|数量/,
  },
  {
    key: 'code',
    tiers: [{ weight: 1, words: ['isbn号', 'isbn', '国际书号', '书号', '条形码', '条码', '产品编码', '产品编号', '商品编码', '商品编号', '物料编码', '物料号', 'sku', '自编码', '内码', '编码', '编号'] }],
    exclude: /序号|单号|日期|订单|客户|供应商|电话|物流/,
  },
  {
    key: 'code2',
    tiers: [{ weight: 1, words: ['货号/科目', '货号', '科目', '内部编码', '内部号', '商品货号', '自编号', '供应商货号'] }],
    exclude: /序号|单号|日期|订单|客户|供应商|电话|物流/,
  },
  {
    key: 'name',
    tiers: [{ weight: 1, words: ['产品名称', '商品名称', '货物名称', '物料名称', '图书名称', '货品名称', '品名规格', '产品名', '商品名', '书名', '品名', '摘要', '名称'] }],
    exclude: /客户|供应商|单位|公司|学校|收货|采购|业务员|部门|品牌|版别|出版社|联系人/,
  },
  {
    key: 'publisher',
    tiers: [{ weight: 1, words: ['出版社', '出版单位', '品牌名称', '供应商', '版别', '厂商', '品牌'] }],
    exclude: /客户|单位|公司|学校/,
  },
];

function matchColumnLabel(label) {
  const raw = String(label ?? '').trim();
  if (!raw) return null;
  const norm = normalizeText(raw);
  if (!norm) return null;

  let best = null;
  for (const def of FIELD_DEFS) {
    if (def.exclude && def.exclude.test(raw)) continue;
    for (const tier of def.tiers) {
      for (const syn of tier.words) {
        const s = normalizeText(syn);
        if (!s) continue;
        let quality = 0;
        if (norm === s) quality = 3;
        else if (s.length >= 2 && norm.includes(s)) quality = 2;
        else if (norm.length >= 2 && s.includes(norm)) quality = 1;
        if (quality === 0) continue;
        const cand = { field: def.key, quality, tier: tier.weight, len: s.length };
        if (
          !best ||
          cand.tier > best.tier ||
          (cand.tier === best.tier &&
            (cand.quality > best.quality || (cand.quality === best.quality && cand.len > best.len)))
        ) {
          best = cand;
        }
      }
    }
  }
  return best;
}

/**
 * 把一行表头分配成 7 个字段列位（1-based 列号）。
 * 同一字段命中多列时按「档位 > 匹配质量 > 靠左」选，不是简单取最左。
 */
function assignColumns(labels) {
  const columns = { name: null, code: null, code2: null, publisher: null, qty: null, amount: null, unitPrice: null };
  const candidates = {};
  const best = {};

  for (let i = 0; i < labels.length; i++) {
    const m = matchColumnLabel(labels[i]);
    if (!m) continue;
    if (!candidates[m.field]) candidates[m.field] = [];
    candidates[m.field].push({ col: i + 1, label: String(labels[i]).trim(), tier: m.tier });
    const cur = best[m.field];
    if (!cur || m.tier > cur.tier || (m.tier === cur.tier && m.quality > cur.quality)) {
      best[m.field] = { col: i + 1, tier: m.tier, quality: m.quality };
    }
  }

  for (const f of FIELDS) {
    if (!best[f]) continue;
    // 金额列只自动采用「实洋」档（应收增加 / 发行实洋）。
    // 只有 标准金额 / 码洋 / 金额 的表格，自动填进去会让人以为是实洋在核对，
    // 所以宁可不填，让用户手动确认 —— 这是业务上最不能猜的一处。
    if (f === 'amount' && best[f].tier < AMOUNT_REAL_TIER) continue;
    columns[f] = best[f].col;
  }
  return { columns, candidates };
}

function scoreColumns(c) {
  let score = 0;
  for (const f of FIELDS) if (c[f]) score++;
  return score;
}

function joinLabel(top, bottom) {
  const a = String(top ?? '').trim();
  const b = String(bottom ?? '').trim();
  if (!a) return b;
  if (!b) return a;
  if (a === b || a.includes(b)) return a;
  if (b.includes(a)) return b;
  return `${a} ${b}`;
}

function hasAny(rows, idx) {
  const row = rows[idx];
  return !!row && row.some((t) => String(t ?? '').trim() !== '');
}

function buildResult(sheetName, rows, rowIndex, labels, columns, twoRow, candidates) {
  const headerRow = rowIndex + 1;
  const headerSpan = twoRow ? 2 : 1;
  const columnLabels = {};
  labels.forEach((label, i) => {
    const t = String(label ?? '').trim();
    if (t) columnLabels[i + 1] = t;
  });

  const warnings = [];
  if (!columns.name && !columns.code && !columns.code2) {
    warnings.push('未识别到产品名称或编码列，请手动指定匹配字段');
  }
  if (!columns.amount) {
    const cands = candidates.amount || [];
    if (cands.length) {
      warnings.push(
        `识别到金额候选列（${cands.map((c) => c.label).join('、')}），但没有实洋列（应收增加 / 发行实洋）；如确认其中某列就是实洋，请手动指定，否则该文件不参与金额核对`,
      );
    } else {
      warnings.push('未识别到金额列（应收增加 / 发行实洋），该文件不参与金额核对');
    }
  }
  if (columns.amount && candidates.amount && candidates.amount.length > 1) {
    const others = candidates.amount.filter((c) => c.col !== columns.amount).map((c) => c.label);
    warnings.push(`识别到多个金额列（${others.join('、')}），已默认使用「${columnLabels[columns.amount]}」，可在下方修改`);
  }

  const previewStartRow = Math.max(1, headerRow - 2);
  const previewRows = [];
  for (let r = previewStartRow - 1; r < Math.min(rows.length, previewStartRow - 1 + 10); r++) {
    previewRows.push(rows[r].map((t) => String(t ?? '')));
  }

  return {
    sheet: sheetName,
    rowCount: rows.length,
    headerRow,
    headerSpan,
    columns,
    columnLabels,
    candidates,
    previewRows,
    previewStartRow,
    warnings,
  };
}

/**
 * 探测某个 sheet 的表头。
 * @param {string} sheetName
 * @param {string[][]} rows
 * @param {number|null} forcedHeaderRow 1-based，用户指定时按单行处理
 */
function detectHeader(sheetName, rows, forcedHeaderRow = null) {
  const scanEnd = Math.min(rows.length, 60);

  if (forcedHeaderRow) {
    const idx = forcedHeaderRow - 1;
    if (idx < 0 || idx >= rows.length) {
      return {
        sheet: sheetName,
        rowCount: rows.length,
        headerRow: forcedHeaderRow,
        headerSpan: 1,
        columns: { name: null, code: null, code2: null, publisher: null, qty: null, amount: null, unitPrice: null },
        columnLabels: {},
        candidates: {},
        previewRows: rows.slice(Math.max(0, idx - 2), Math.max(0, idx - 2) + 10).map((r) => r.map((t) => String(t ?? ''))),
        previewStartRow: Math.max(1, forcedHeaderRow - 2),
        warnings: [`指定的表头行（第 ${forcedHeaderRow} 行）没有可识别的字段，请同时手动指定列`],
      };
    }
    const labels = rows[idx].map((t) => String(t ?? '').trim());
    const { columns, candidates } = assignColumns(labels);
    const warnings = [];
    if (!columns.qty) warnings.push('未识别到数量列，请手动指定');
    if (!columns.amount && !columns.unitPrice) warnings.push('未识别到金额列，也没有单价列');
    if (!columns.name && !columns.code && !columns.code2) warnings.push('未识别到产品名称或编码列');
    return buildResult(sheetName, rows, idx, labels, columns, false, candidates);
  }

  let best = null;
  for (let r = 0; r < scanEnd; r++) {
    if (!hasAny(rows, r)) continue;
    const labels = rows[r].map((t) => String(t ?? '').trim());
    const { columns, candidates } = assignColumns(labels);
    const score = scoreColumns(columns);
    if (!best || score > best.score) best = { row: r, columns, candidates, score, labels, twoRow: false };
  }

  // 两行表头（分组行 + 字段行）兜底：样例里没出现，但客户模板常见
  for (let r = 0; r + 1 < scanEnd; r++) {
    if (!hasAny(rows, r) || !hasAny(rows, r + 1)) continue;
    const merged = rows[r].map((t, i) => joinLabel(t, rows[r + 1][i]));
    const { columns, candidates } = assignColumns(merged);
    const score = scoreColumns(columns) - 0.5;
    if (!best || score > best.score) best = { row: r, columns, candidates, score, labels: merged, twoRow: true };
  }

  if (!best || best.score < 2) {
    return {
      sheet: sheetName,
      rowCount: rows.length,
      headerRow: 1,
      headerSpan: 1,
      columns: { name: null, code: null, code2: null, publisher: null, qty: null, amount: null, unitPrice: null },
      columnLabels: {},
      candidates: {},
      previewRows: rows.slice(0, 10).map((r) => r.map((t) => String(t ?? ''))),
      previewStartRow: 1,
      warnings: ['未能自动识别表头（没有找到同时包含数量与金额/单价的行），请手动指定表头行与列映射'],
    };
  }

  return buildResult(sheetName, rows, best.row, best.labels, best.columns, best.twoRow, best.candidates);
}

module.exports = { detectHeader, assignColumns, matchColumnLabel, FIELDS };

});

define("extract", function (module, exports, require, Buffer) {
'use strict';

/**
 * 按列映射从网格里抽数据行，并做金额/数量归一。
 *
 * 真实数据里必须处理掉的行：
 *   对账单「期初 / 收款 / 合计」行、发货单尾部「合计」「采购签字：」行、
 *   xlsx 订书单第 1~2 行客户抬头 —— 否则汇总会多算。
 */

const { parseMoney, parseQty } = require('values');
const { normalizeText } = require('text');

const MAX_WARNINGS = 40;
const TOTAL_ROW_RE = /^(合计|总计|小计|累计|本页合计|本月合计|总计数|total|sum)$/;
// 对账单里的收付款/结转型行：不是产品，必须跳过，否则会凭空多出"仅对账单有"的产品
const NON_PRODUCT_RE =
  /^(收款|付款|收付款|期初|期初余额|结余|余额|结转|上年结转|本期结转|调账|记账|红冲|作废|合计|总计|小计)$/;

function cell(row, col) {
  if (!col) return '';
  const v = row[col - 1];
  return v === undefined || v === null ? '' : String(v);
}

function isTotalRow(name, code, code2) {
  // 「合计」可能写在名称列，也可能写在编码列（真实对账单里在客户代码列）
  for (const raw of [name, code, code2]) {
    const n = normalizeText(raw);
    if (!n) continue;
    if (TOTAL_ROW_RE.test(n)) return true;
    if (/^(合计|总计|小计)/.test(n)) return true;
  }
  return false;
}

/**
 * @param {object} p
 * @param {string[][]} p.rows 网格
 * @param {number} p.headerRow 1-based
 * @param {number} p.headerSpan 1 或 2
 * @param {object} p.columns 7 个列位（1-based 或 null）
 * @returns {{ items: object[], skipped: object[], warnings: string[], amountMode: string, derivedCount: number }}
 */
function extractRows({ rows, headerRow, headerSpan, columns }) {
  const dataStart = headerRow - 1 + (headerSpan || 1); // 0-based 起始下标
  const items = [];
  const skipped = [];
  const warnings = [];
  let derivedCount = 0;
  // v3：金额只认金额列（对账单「应收增加」/ 发货单「发行实洋」）。
  // 不映射金额列的文件不参与金额核对，不再用单价×数量派生（派生出来是码洋，口径不对）。
  const noAmountColumn = !columns.amount;
  let hasAmountValue = false;
  if (noAmountColumn) {
    warnings.push('该文件没有金额列，不参与金额核对');
  }

  const warn = (msg) => {
    if (warnings.length < MAX_WARNINGS) warnings.push(msg);
  };

  for (let r = dataStart; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const excelRow = r + 1;

    const name = cell(row, columns.name).trim();
    const code = cell(row, columns.code).trim();
    const code2 = cell(row, columns.code2).trim();
    const publisher = cell(row, columns.publisher).trim();
    const qtyRaw = cell(row, columns.qty).trim();
    const amountRaw = cell(row, columns.amount).trim();
    const priceRaw = cell(row, columns.unitPrice).trim();

    const allEmpty = !name && !code && !code2 && !qtyRaw && !amountRaw && !priceRaw;
    if (allEmpty) continue;

    // 合计 / 小计行
    if (isTotalRow(name, code, code2)) {
      skipped.push({ row: excelRow, reason: '合计行' });
      continue;
    }
    // 收付款 / 期初 / 结余行（对账单特有）
    if (!code && !code2 && NON_PRODUCT_RE.test(normalizeText(name))) {
      skipped.push({ row: excelRow, reason: `非产品行：${name}` });
      continue;
    }
    // 既没有编码，也没有数量/金额/单价：抬头、备注、签字行
    if (!code && !code2 && !qtyRaw && !amountRaw && !priceRaw) {
      skipped.push({ row: excelRow, reason: '无数量与金额' });
      continue;
    }
    // 没有任何匹配键：收款行、客户抬头
    if (!name && !code && !code2) {
      if (qtyRaw || amountRaw) {
        warn(`第 ${excelRow} 行缺少产品名称/编码，但有数量或金额，已跳过`);
      }
      skipped.push({ row: excelRow, reason: '缺少产品名称/编码' });
      continue;
    }
    // 数量金额都没有：抬头、备注行
    if (!qtyRaw && !amountRaw && !priceRaw) {
      skipped.push({ row: excelRow, reason: '无数量与金额' });
      continue;
    }

    const qtyParsed = parseQty(qtyRaw);
    if (!qtyParsed.ok) {
      warn(`第 ${excelRow} 行${qtyParsed.error}，已跳过该行`);
      skipped.push({ row: excelRow, reason: `数量无法识别：${qtyRaw}` });
      continue;
    }
    if (qtyParsed.warnings.length) warn(`第 ${excelRow} 行：${qtyParsed.warnings.join('；')}`);

    // 金额：只取金额列，空值记 null（未知），不计入总额也不当作 0
    let amountCents = null;
    if (columns.amount) {
      if (amountRaw) {
        const m = parseMoney(amountRaw);
        if (!m.ok) {
          warn(`第 ${excelRow} 行${m.error}，该行金额未计入`);
          skipped.push({ row: excelRow, reason: `金额无法识别：${amountRaw}` });
          continue;
        }
        if (m.warnings.length) warn(`第 ${excelRow} 行：${m.warnings.join('；')}`);
        amountCents = m.cents;
        hasAmountValue = true;
      } else {
        warn(`第 ${excelRow} 行金额为空，该行未计入总额`);
      }
    }

    if (qtyParsed.blank) warn(`第 ${excelRow} 行数量为空，已按 0 计`);

    items.push({
      row: excelRow,
      name,
      code,
      code2,
      publisher,
      qty: qtyParsed.scaled,
      amount: amountCents,
      qtyRaw,
      amountRaw: amountRaw || '',
    });
  }

  if (warnings.length >= MAX_WARNINGS) warnings.push(`……警告较多，仅显示前 ${MAX_WARNINGS} 条`);

  // 只有"映射了金额列且至少一行有金额"才算参与金额核对
  const amountChecked = !noAmountColumn && hasAmountValue;
  if (!noAmountColumn && !hasAmountValue) {
    warnings.push('金额列全部为空，该文件不参与金额核对');
  }

  return {
    items,
    skipped,
    warnings,
    amountChecked,
    amountMode: amountChecked ? '金额列' : '无金额列（未参与核对）',
  };
}

module.exports = { extractRows, isTotalRow };

});

define("engine", function (module, exports, require, Buffer) {
'use strict';

/**
 * 汇总比对引擎。
 *
 * 匹配策略是「级联多字段」而不是单键，原因见 docs/field-analysis.md：
 *   发货单 A 只有 ISBN，发货单 B 只有内部书号，对账单两列都有关键且编号体系不同，
 *   所以用并查集把所有行按 (编码1, 编码2, 标准化名称) 连成"产品实体"，再做同侧汇总。
 *
 * 连边顺序决定匹配依据：
 *   1. 编码相等（编码1 与 编码2 允许交叉，因为两个模板用不同列承载同一编号体系）
 *   2. 名称标准化后相等（出版社不相容时不连，避免同名不同书）
 *   3. 剩下单边实体之间做名称相似度配对（阈值 0.82），标记 fuzzy，不静默
 */

const { formatCents } = require('values');
const { normalizeText, normalizeCode, similarity, publisherCompatible } = require('text');
const { AppError } = require('errors');

const FUZZY_THRESHOLD = 0.82;
const FUZZY_SUSPECT_THRESHOLD = 0.7;
const FUZZY_MAX_PAIRS = 40000;

const CATEGORY_ORDER = { amount_diff: 0, statement_only: 1, shipment_only: 2, unknown_amount: 3, match: 4 };

class DSU {
  constructor(n) {
    this.parent = new Int32Array(n);
    for (let i = 0; i < n; i++) this.parent[i] = i;
  }
  find(x) {
    while (this.parent[x] !== x) {
      this.parent[x] = this.parent[this.parent[x]];
      x = this.parent[x];
    }
    return x;
  }
  union(a, b) {
    const ra = this.find(a);
    const rb = this.find(b);
    if (ra === rb) return false;
    this.parent[rb] = ra;
    return true;
  }
}

function makeNodes(side, file, nodes) {
  for (const item of file.items) {
    nodes.push({ side, fileId: file.fileId, fileName: file.fileName, item });
  }
}

/** 按名称 + 出版社相容性把一组节点聚成若干簇（避免同名不同出版社被并到一起） */
function clusterByPublisher(nodeIdxs, nodes) {
  const clusters = [];
  for (const idx of nodeIdxs) {
    const pub = nodes[idx].item.publisher;
    let placed = false;
    for (const cluster of clusters) {
      if (publisherCompatible(cluster.publisher, pub)) {
        cluster.members.push(idx);
        if (!cluster.publisher && pub) cluster.publisher = pub;
        placed = true;
        break;
      }
    }
    if (!placed) clusters.push({ publisher: pub || '', members: [idx] });
  }
  return clusters;
}

/**
 * @param {{ statement: object, shipments: object[] }} input
 *   statement/shipment = { fileId, fileName, items, warnings, skipped, amountMode }
 */
function reconcile({ statement, shipments }) {
  if (!statement.items || statement.items.length === 0) {
    throw new AppError(
      'PARSE_FAILED',
      `对账单「${statement.fileName}」没有解析出任何数据行，请检查表头行与列映射是否正确`,
      { fileName: statement.fileName },
    );
  }
  const shipmentItemCount = shipments.reduce((n, f) => n + f.items.length, 0);
  if (shipmentItemCount === 0) {
    throw new AppError('PARSE_FAILED', '发货单没有解析出任何数据行，请检查表头行与列映射是否正确');
  }

  const nodes = [];
  makeNodes('statement', statement, nodes);
  for (const f of shipments) makeNodes('shipment', f, nodes);

  const dsu = new DSU(nodes.length);

  // ---- 1. 编码相等（含 code ↔ code2 交叉） ----
  const codeMap = new Map();
  nodes.forEach((n, i) => {
    for (const raw of [n.item.code, n.item.code2]) {
      const k = normalizeCode(raw);
      if (!k || k.length < 2) continue;
      if (!codeMap.has(k)) codeMap.set(k, []);
      codeMap.get(k).push(i);
    }
  });
  for (const list of codeMap.values()) {
    for (let i = 1; i < list.length; i++) dsu.union(list[0], list[i]);
  }

  // ---- 2. 名称标准化后相等（出版社相容才连） ----
  const nameMap = new Map();
  nodes.forEach((n, i) => {
    const k = normalizeText(n.item.name);
    if (!k || k.length < 2) return;
    if (!nameMap.has(k)) nameMap.set(k, []);
    nameMap.get(k).push(i);
  });
  for (const list of nameMap.values()) {
    if (list.length < 2) continue;
    for (const cluster of clusterByPublisher(list, nodes)) {
      for (let i = 1; i < cluster.members.length; i++) dsu.union(cluster.members[0], cluster.members[i]);
    }
  }

  // ---- 组装实体 ----
  const groups = new Map();
  nodes.forEach((n, i) => {
    const root = dsu.find(i);
    if (!groups.has(root)) groups.set(root, []);
    groups.get(root).push(i);
  });

  let entities = [...groups.values()].map((idxs) => {
    const stmt = idxs.filter((i) => nodes[i].side === 'statement');
    const ship = idxs.filter((i) => nodes[i].side === 'shipment');
    return { idxs, stmt, ship, fuzzy: null };
  });

  // ---- 3. 单边实体做名称模糊配对 ----
  const suspects = [];
  let onlyStmt = entities.filter((e) => e.stmt.length > 0 && e.ship.length === 0);
  let onlyShip = entities.filter((e) => e.ship.length > 0 && e.stmt.length === 0);

  if (onlyStmt.length > 0 && onlyShip.length > 0) {
    const pairs = [];
    if (onlyStmt.length * onlyShip.length > FUZZY_MAX_PAIRS) {
      suspects.push({ left: '', right: '', reason: `单边产品过多（${onlyStmt.length} × ${onlyShip.length}），已跳过名称模糊匹配` });
    } else {
      for (const s of onlyStmt) {
        const sName = entityName(s, nodes);
        const sPub = entityPublisher(s, nodes);
        for (const p of onlyShip) {
          const pName = entityName(p, nodes);
          const pPub = entityPublisher(p, nodes);
          if (!publisherCompatible(sPub, pPub)) continue;
          const sim = similarity(sName, pName);
          if (sim < FUZZY_SUSPECT_THRESHOLD) continue;
          pairs.push({ s, p, sim });
        }
      }
      pairs.sort((a, b) => b.sim - a.sim);

      const usedS = new Set();
      const usedP = new Set();
      for (const pair of pairs) {
        if (pair.sim < FUZZY_THRESHOLD) {
          if (usedS.has(pair.s) || usedP.has(pair.p)) continue;
          suspects.push({
            left: entityName(pair.s, nodes),
            right: entityName(pair.p, nodes),
            reason: `仅一侧存在且名称高度相似（相似度 ${pair.sim.toFixed(2)}）`,
          });
          if (suspects.length >= 10) break;
          continue;
        }
        if (usedS.has(pair.s) || usedP.has(pair.p)) continue;
        usedS.add(pair.s);
        usedP.add(pair.p);
        // 合并两个单边实体
        const merged = {
          idxs: pair.s.idxs.concat(pair.p.idxs),
          stmt: pair.s.stmt,
          ship: pair.p.ship,
          fuzzy: { sim: pair.sim },
        };
        entities = entities.map((e) => (e === pair.s || e === pair.p ? null : e)).filter(Boolean);
        entities.push(merged);
      }
    }
  }

  // ---- 4. 生成结果行 ----
  const rows = entities.map((e) => buildRow(e, nodes));
  rows.sort((a, b) => {
    const ca = CATEGORY_ORDER[a.category];
    const cb = CATEGORY_ORDER[b.category];
    if (ca !== cb) return ca - cb;
    return String(a.key).localeCompare(String(b.key), 'zh');
  });

  const amountDiffCount = rows.filter((r) => r.category === 'amount_diff').length;
  const statementOnlyCount = rows.filter((r) => r.category === 'statement_only').length;
  const shipmentOnlyCount = rows.filter((r) => r.category === 'shipment_only').length;
  const unknownAmountCount = rows.filter((r) => r.category === 'unknown_amount').length;
  const matchCount = rows.filter((r) => r.category === 'match').length;
  const diffCount = amountDiffCount + statementOnlyCount + shipmentOnlyCount;

  const matchStats = { code: 0, code2: 0, name: 0, fuzzy: 0 };
  for (const r of rows) if (r.matchBy) matchStats[r.matchBy] = (matchStats[r.matchBy] || 0) + 1;

  const warnings = [
    ...(statement.warnings || []),
    ...shipments.flatMap((f) => (f.warnings || []).map((w) => `${f.fileName}：${w}`)),
  ];

  // ---- 总额：只看金额，且只统计参与金额核对的单据 ----
  const stmtSum = sumAmount(statement.items);
  const checkedFiles = shipments.filter((f) => f.amountChecked);
  const excludedFiles = shipments.filter((f) => !f.amountChecked);
  let shipTotal = 0n;
  const fileAmounts = new Map();
  for (const f of shipments) {
    const s0 = sumAmount(f.items);
    fileAmounts.set(f.fileId, s0.total);
    if (f.amountChecked) shipTotal += s0.total;
  }
  const amountDiff = shipTotal - stmtSum.total;
  const matched = checkedFiles.length > 0 && amountDiff === 0n && excludedFiles.length === 0;

  let totalConclusion;
  if (checkedFiles.length === 0) {
    totalConclusion = '没有可参与金额核对的发货单，无法核对总额';
  } else if (amountDiff === 0n && excludedFiles.length === 0) {
    totalConclusion = '总额一致';
  } else if (amountDiff === 0n) {
    totalConclusion = `已参与核对的部分金额一致，另有 ${excludedFiles.length} 张单据未参与核对`;
  } else {
    const word = amountDiff > 0n ? '多' : '少';
    const abs = formatCents(amountDiff < 0n ? -amountDiff : amountDiff);
    totalConclusion = `总额不一致，发货单${word} ${abs} 元`;
    if (excludedFiles.length > 0) totalConclusion += `（另有 ${excludedFiles.length} 张单据未参与核对）`;
  }

  const nameMerged = [];
  for (const r of rows) {
    if (r.aliases && r.aliases.length > 1) nameMerged.push({ key: r.key, aliases: r.aliases });
  }

  const summary = {
    statementFile: statement.fileName,
    statement: {
      fileId: statement.fileId,
      fileName: statement.fileName,
      rows: statement.items.length,
      amount: formatCents(stmtSum.total),
      amountMode: statement.amountMode || '金额列',
    },
    shipmentFiles: shipments.map((f) => ({
      fileId: f.fileId,
      fileName: f.fileName,
      rows: f.items.length,
      skippedRows: (f.skipped || []).length,
      amount: f.amountChecked ? formatCents(fileAmounts.get(f.fileId)) : null,
      amountChecked: !!f.amountChecked,
      amountMode: f.amountMode || '金额列',
    })),
    shipmentCount: shipments.length,
    total: {
      statementAmount: formatCents(stmtSum.total),
      shipmentAmount: formatCents(shipTotal),
      amountDiff: formatCents(amountDiff),
      matched,
      conclusion: totalConclusion,
      checkedFileCount: checkedFiles.length,
      excludedFileCount: excludedFiles.length,
    },
    productCount: rows.length,
    matchedCount: matchCount,
    amountDiffCount,
    statementOnlyCount,
    shipmentOnlyCount,
    unknownAmountCount,
    diffCount,
    identical: matched,
    matchStats,
    warnings,
    nameMerged: nameMerged.slice(0, 20),
  };

  return { summary, rows, suspects: suspects.slice(0, 10) };
}

function entityName(entity, nodes) {
  for (const i of entity.stmt) if (nodes[i].item.name) return nodes[i].item.name;
  for (const i of entity.ship) if (nodes[i].item.name) return nodes[i].item.name;
  return '';
}

function entityPublisher(entity, nodes) {
  for (const i of entity.idxs) if (nodes[i].item.publisher) return nodes[i].item.publisher;
  return '';
}

/** 判断实体两侧是通过哪种字段连上的 */
function resolveMatchBy(entity, nodes) {
  if (entity.fuzzy) return 'fuzzy';
  const s = entity.stmt.map((i) => nodes[i].item);
  const p = entity.ship.map((i) => nodes[i].item);

  const sCode = new Set(s.map((it) => normalizeCode(it.code)).filter(Boolean));
  const pCode = new Set(p.map((it) => normalizeCode(it.code)).filter(Boolean));
  if ([...sCode].some((c) => pCode.has(c))) return 'code';

  const sCode2 = new Set(s.map((it) => normalizeCode(it.code2)).filter(Boolean));
  const pCode2 = new Set(p.map((it) => normalizeCode(it.code2)).filter(Boolean));
  const crossA = [...sCode2].some((c) => pCode.has(c));
  const crossB = [...sCode].some((c) => pCode2.has(c));
  const same2 = [...sCode2].some((c) => pCode2.has(c));
  if (crossA || crossB || same2) return 'code2';

  const sName = new Set(s.map((it) => normalizeText(it.name)).filter(Boolean));
  const pName = new Set(p.map((it) => normalizeText(it.name)).filter(Boolean));
  if ([...sName].some((c) => pName.has(c))) return 'name';

  return null;
}

function sumAmount(items) {
  let total = 0n;
  let known = 0;
  let unknown = 0;
  for (const it of items) {
    if (it.amount === null || it.amount === undefined) unknown++;
    else {
      total += it.amount;
      known++;
    }
  }
  return { total, known, unknown };
}

function buildRow(entity, nodes) {
  const stmtItems = entity.stmt.map((i) => nodes[i].item);
  const shipItems = entity.ship.map((i) => nodes[i].item);

  const stmt = sumAmount(stmtItems);
  const ship = sumAmount(shipItems);

  const hasStmt = stmtItems.length > 0;
  const hasShip = shipItems.length > 0;
  const amountKnown = hasShip && ship.known > 0;

  const note = [];
  let category;
  let matchBy = null;

  if (hasStmt && hasShip) {
    matchBy = resolveMatchBy(entity, nodes);
    if (!amountKnown) {
      category = 'unknown_amount';
      note.push('发货单侧有记录，但来源单据没有金额列，金额未知');
    } else {
      const diff = ship.total - stmt.total;
      if (diff === 0n) {
        category = 'match';
        note.push('金额一致');
      } else {
        category = 'amount_diff';
        note.push(`金额差 ${formatCents(diff)} 元（${diff > 0n ? '发货单多' : '对账单多'} ${formatCents(diff < 0n ? -diff : diff)} 元）`);
      }
      if (ship.unknown > 0) note.push(`发货单有 ${ship.unknown} 行金额为空，未计入`);
    }
  } else if (hasStmt) {
    category = 'statement_only';
    note.push(`对账单有此产品：${formatCents(stmt.total)} 元，发货单未找到`);
  } else {
    category = 'shipment_only';
    note.push(
      amountKnown
        ? `发货单有此产品：${formatCents(ship.total)} 元，对账单未找到`
        : '发货单有此产品（金额未知），对账单未找到',
    );
  }

  if (hasStmt && stmt.unknown > 0) note.push(`对账单有 ${stmt.unknown} 行金额为空，未计入`);

  const code = firstNonEmpty(stmtItems, 'code') || firstNonEmpty(shipItems, 'code');
  const code2 = firstNonEmpty(stmtItems, 'code2') || firstNonEmpty(shipItems, 'code2');
  const displayName = firstNonEmpty(stmtItems, 'name') || firstNonEmpty(shipItems, 'name');
  const key = code || code2 || displayName || `未命名-${entity.idxs[0]}`;

  const aliases = [...new Set([...stmtItems, ...shipItems].map((it) => String(it.name || '').trim()).filter(Boolean))];

  const sources = [];
  const byFile = new Map();
  for (const i of entity.idxs) {
    const n = nodes[i];
    if (!byFile.has(n.fileId)) {
      const entry = { fileId: n.fileId, fileName: n.fileName, side: n.side, rows: [] };
      byFile.set(n.fileId, entry);
      sources.push(entry);
    }
    byFile.get(n.fileId).rows.push(n.item.row);
  }
  sources.sort((a, b) => (a.side === b.side ? 0 : a.side === 'statement' ? -1 : 1));
  for (const s of sources) s.rows.sort((x, y) => x - y);

  return {
    key,
    displayName,
    code,
    code2,
    statementAmount: formatCents(stmt.total),
    shipmentAmount: amountKnown ? formatCents(ship.total) : null,
    amountDiff: amountKnown ? formatCents(ship.total - stmt.total) : null,
    amountKnown,
    category,
    matchBy,
    note,
    aliases,
    sources,
  };
}

function firstNonEmpty(arr, field) {
  for (const it of arr) {
    const v = String(it[field] || '').trim();
    if (v) return v;
  }
  return '';
}

module.exports = { reconcile, FUZZY_THRESHOLD };

});

define("sheet", function (module, exports, require, Buffer) {
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

});

define("export", function (module, exports, require, Buffer) {
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

});

  window.ReconcileLib = {
  loadGrid: require("sheet").loadGrid,
  detectHeader: require("detect").detectHeader,
  extractRows: require("extract").extractRows,
  reconcile: require("engine").reconcile,
  buildDiffWorkbook: require("export").buildDiffWorkbook
  };
})();
