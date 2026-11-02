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

const { normalizeText } = require('./text');

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
