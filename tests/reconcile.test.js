'use strict';

const test = require('node:test');
const assert = require('node:assert');
const path = require('path');

const { parseMoney, parseQty, mulPriceByQty, formatCents, formatQty } = require('../server/lib/values');
const { normalizeText, normalizeCode, similarity } = require('../server/lib/text');
const { reconcile } = require('../server/lib/engine');
const { extractRows } = require('../server/lib/extract');
const { detectHeader } = require('../server/lib/detect');
const { loadGrid } = require('../server/lib/sheet');

// ---------- 金额：按分整数比较，不允许误差 ----------

test('金额解析：基本与千分位', () => {
  assert.strictEqual(parseMoney('1000').cents, 100000n);
  assert.strictEqual(parseMoney('1,234.56').cents, 123456n);
  assert.strictEqual(parseMoney('¥ 1,200.00').cents, 120000n);
  assert.strictEqual(parseMoney('100.01').cents, 10001n);
  assert.strictEqual(parseMoney('.5').cents, 50n);
});

test('金额解析：负数与括号', () => {
  assert.strictEqual(parseMoney('-1200').cents, -120000n);
  assert.strictEqual(parseMoney('(1,200.50)').cents, -120050n);
  assert.strictEqual(parseMoney('1200-').cents, -120000n);
  assert.strictEqual(parseMoney('1200元').cents, 120000n);
});

test('金额解析：超过 2 位小数四舍五入并给警告', () => {
  const r = parseMoney('1.005');
  assert.strictEqual(r.cents, 101n);
  assert.ok(r.warnings.length === 1);
});

test('金额解析：无法识别时报错而不是静默归零', () => {
  assert.strictEqual(parseMoney('abc').ok, false);
  assert.strictEqual(parseMoney('1.234,56').ok, false);
});

test('0.01 元差异必须能被区分（浮点会失真，整数不会）', () => {
  const a = parseMoney('100.00').cents;
  const b = parseMoney('100.01').cents;
  assert.notStrictEqual(a, b);
  assert.strictEqual(b - a, 1n);
  // 浮点对照：0.1+0.2 !== 0.3 这类误差不会进入比较
  assert.strictEqual(formatCents(b - a), '0.01');
});

test('金额格式：负数与补零', () => {
  assert.strictEqual(formatCents(0n), '0.00');
  assert.strictEqual(formatCents(-1n), '-0.01');
  assert.strictEqual(formatCents(123456n), '1234.56');
});

// ---------- 数量 ----------

test('数量解析：整数、小数、千分位、单位', () => {
  assert.strictEqual(parseQty('10').scaled, 10000000n);
  assert.strictEqual(parseQty('10.5').scaled, 10500000n);
  assert.strictEqual(parseQty('1,000').scaled, 1000000000n);
  assert.strictEqual(parseQty('10件').scaled, 10000000n);
  assert.strictEqual(parseQty('10abc').ok, false);
});

test('数量格式：去尾零', () => {
  assert.strictEqual(formatQty(10000000n), '10');
  assert.strictEqual(formatQty(10500000n), '10.5');
  assert.strictEqual(formatQty(-3000000n), '-3');
});

test('单价×数量：49.8 × 43 = 2141.40', () => {
  assert.strictEqual(mulPriceByQty(4980n, 43000000n), 214140n);
});

// ---------- 文本标准化 ----------

test('名称标准化：全半角、空格、破折号、大小写', () => {
  assert.strictEqual(normalizeText('Ａ－ 产品'), normalizeText('a-产品'));
  assert.strictEqual(normalizeText('XX—100'), normalizeText('xx-100'));
  assert.strictEqual(normalizeText(' 钢琴基础（第一册）'), '钢琴基础(第一册)');
});

test('编码标准化：去符号、统一大写', () => {
  assert.strictEqual(normalizeCode('978-7-5648-4290-1'), '9787564842901');
  assert.strictEqual(normalizeCode(' 026c060 '), '026C060');
});

test('相似度：括号作者后缀不影响主体判定', () => {
  assert.ok(similarity('婴幼儿教育学（岳曼）', '婴幼儿教育学') > 0.9);
  assert.ok(similarity('儿童保温杯', '不锈钢餐盘') < 0.3);
});

// ---------- 表头探测 ----------

test('表头探测：发货单表头在第 3 行（前面是客户抬头）', () => {
  const grid = loadGrid(path.join(__dirname, '..', 'samples', '验收-发货单-1.xlsx'));
  const sheet = grid.sheets.find((s) => !s.empty);
  const info = detectHeader(sheet.name, sheet.rows, null);
  assert.strictEqual(info.headerRow, 3);
  assert.strictEqual(info.columns.code, 2);
  assert.strictEqual(info.columns.name, 3);
  assert.strictEqual(info.columns.qty, 4);
  // 码洋在第 5 列、发行实洋在第 7 列，必须选后者
  assert.strictEqual(info.columns.amount, 7);
  assert.strictEqual(info.columnLabels[7], '发行实洋');
});

test('表头探测：对账单表头在第 3 行（有 2 行标题）', () => {
  const grid = loadGrid(path.join(__dirname, '..', 'samples', '验收-对账单.xlsx'));
  const sheet = grid.sheets.find((s) => !s.empty);
  const info = detectHeader(sheet.name, sheet.rows, null);
  assert.strictEqual(info.headerRow, 3);
  assert.strictEqual(info.columns.name, 2);
  assert.strictEqual(info.columnLabels[info.columns.amount], '应收增加');
});

test('表头探测：.xls 老格式也能读', () => {
  const grid = loadGrid(path.join(__dirname, '..', 'samples', '兼容-xls-发货单.xls'));
  const sheet = grid.sheets.find((s) => !s.empty);
  const info = detectHeader(sheet.name, sheet.rows, null);
  assert.strictEqual(info.headerRow, 2);
  assert.strictEqual(info.columns.qty, 4);
});

// ---------- 行抽取 ----------

test('行抽取：合计行、收款行、空编码行被跳过', () => {
  const rows = [
    ['产品编码', '产品名称', '数量', '金额'],
    ['P001', '甲产品', '10', '100.00'],
    ['', '收款', '', ''],
    ['合计', '', '10', '100.00'],
    ['', 'Ａ-100 型号', '2', '20.00'],
    ['P002', '', '1', '10.00'],
    ['', '', '', ''],
  ];
  const data = extractRows({
    rows,
    headerRow: 1,
    headerSpan: 1,
    columns: { name: 2, code: 1, code2: null, publisher: null, qty: 3, amount: 4, unitPrice: null },
  });
  assert.strictEqual(data.items.length, 3);
  const names = data.items.map((i) => i.name);
  assert.deepStrictEqual(names, ['甲产品', 'Ａ-100 型号', '']);
  assert.ok(data.skipped.some((s) => s.reason.includes('合计行')));
  assert.ok(data.skipped.some((s) => s.reason.includes('收款')));
});

test('行抽取：没有金额列时不派生金额，标记不参与核对', () => {
  const rows = [
    ['编码', '名称', '定价', '数量'],
    ['A1', '甲', '49.8', '43'],
  ];
  const data = extractRows({
    rows,
    headerRow: 1,
    headerSpan: 1,
    columns: { name: 2, code: 1, code2: null, publisher: null, qty: 4, amount: null, unitPrice: 3 },
  });
  assert.strictEqual(data.items.length, 1);
  assert.strictEqual(data.items[0].amount, null, '金额应为未知而不是用定价×数量凑出来');
  assert.strictEqual(data.amountChecked, false);
  assert.strictEqual(data.amountMode, '无金额列（未参与核对）');
});

test('行抽取：金额为空的行记未知，不按 0 计', () => {
  const rows = [
    ['编码', '名称', '数量', '金额'],
    ['A1', '甲', '5', ''],
    ['A2', '乙', '2', '10.00'],
  ];
  const data = extractRows({
    rows,
    headerRow: 1,
    headerSpan: 1,
    columns: { name: 2, code: 1, code2: null, publisher: null, qty: 3, amount: 4, unitPrice: null },
  });
  assert.strictEqual(data.items[0].amount, null);
  assert.strictEqual(data.items[1].amount, 1000n);
  assert.strictEqual(data.amountChecked, true);
});

// ---------- 比对引擎 ----------

function item(o) {
  return {
    row: 1,
    name: '',
    code: '',
    code2: '',
    publisher: '',
    qty: 0n,
    amount: 0n,
    ...o,
  };
}

function file(role, items, opts = {}) {
  const hasAmount = items.some((i) => i.amount !== null && i.amount !== undefined);
  const checked = opts.amountChecked !== undefined ? opts.amountChecked : hasAmount;
  return {
    fileId: role,
    fileName: `${role}.xlsx`,
    items,
    warnings: [],
    skipped: [],
    amountMode: checked ? '金额列' : '无金额列（未参与核对）',
    amountChecked: checked,
  };
}

test('比对：金额差 0.01 判为不一致', () => {
  const { summary, rows } = reconcile({
    statement: file('statement', [item({ name: '甲', code: 'P1', qty: 10000000n, amount: 10000n })]),
    shipments: [file('ship', [item({ name: '甲', code: 'P1', qty: 10000000n, amount: 10001n })])],
  });
  assert.strictEqual(summary.identical, false);
  assert.strictEqual(summary.amountDiffCount, 1);
  assert.strictEqual(summary.total.amountDiff, '0.01');
  assert.strictEqual(rows[0].category, 'amount_diff');
  assert.strictEqual(rows[0].amountDiff, '0.01');
});

test('比对：多张发货单同产品汇总后再比', () => {
  const { rows, summary } = reconcile({
    statement: file('statement', [item({ name: '甲', code: 'P1', qty: 10000000n, amount: 10000n })]),
    shipments: [
      file('ship1', [item({ name: '甲', code: 'P1', qty: 4000000n, amount: 4000n })]),
      file('ship2', [item({ name: '甲', code: 'P1', qty: 6000000n, amount: 6000n })]),
    ],
  });
  assert.strictEqual(rows[0].category, 'match');
  assert.strictEqual(rows[0].shipmentAmount, '100.00');
  assert.strictEqual(summary.total.amountDiff, '0.00');
  assert.strictEqual(summary.total.matched, true);
  assert.strictEqual(summary.identical, true);
});

test('比对：只在一方的产品单独归类', () => {
  const { rows } = reconcile({
    statement: file('statement', [
      item({ name: '甲', code: 'P1', qty: 10000000n, amount: 10000n }),
      item({ name: '乙', code: 'P2', qty: 10000000n, amount: 10000n }),
    ]),
    shipments: [file('ship', [item({ name: '丙', code: 'P3', qty: 10000000n, amount: 10000n })])],
  });
  const byCat = rows.reduce((m, r) => ((m[r.category] = m[r.category] || []).push(r), m), {});
  assert.strictEqual(byCat.statement_only.length, 2);
  assert.strictEqual(byCat.shipment_only.length, 1);
  assert.strictEqual(byCat.statement_only[0].amountKnown, false);
});

test('比对：编码不同但名称标准化后相同也能匹配', () => {
  const { rows } = reconcile({
    statement: file('statement', [item({ name: 'Ａ-100 型号', code: 'P9', qty: 10000000n, amount: 10000n })]),
    shipments: [file('ship', [item({ name: 'a-100型号', code: 'P9', qty: 10000000n, amount: 10000n })])],
  });
  assert.strictEqual(rows[0].category, 'match');
});

test('比对：编码1与编码2交叉匹配（两种发货单模板）', () => {
  const { rows } = reconcile({
    statement: file('statement', [item({ name: '甲', code: '9787564842901', code2: '026Q081', qty: 10000000n, amount: 10000n })]),
    shipments: [file('ship', [item({ name: '甲', code: '026Q081', qty: 10000000n, amount: 10000n })])],
  });
  assert.strictEqual(rows[0].category, 'match');
  assert.strictEqual(rows[0].matchBy, 'code2');
});

test('比对：名称相似度高且仅单边时按 fuzzy 匹配并标注', () => {
  const { rows, summary } = reconcile({
    statement: file('statement', [item({ name: '婴幼儿教育学', code: 'P1', qty: 10000000n, amount: 10000n })]),
    shipments: [file('ship', [item({ name: '婴幼儿教育学（岳曼）', code: 'P2', qty: 10000000n, amount: 10000n })])],
  });
  assert.strictEqual(rows[0].category, 'match');
  assert.strictEqual(rows[0].matchBy, 'fuzzy');
  assert.strictEqual(summary.matchStats.fuzzy, 1);
});

test('比对：疑似同一产品只提示、不并入计算', () => {
  const { rows, suspects } = reconcile({
    statement: file('statement', [item({ name: 'B产品-100', code: 'B1', qty: 10000000n, amount: 10000n })]),
    shipments: [file('ship', [item({ name: 'B产品 100', code: 'B2', qty: 10000000n, amount: 10000n })])],
  });
  assert.strictEqual(rows.length, 2);
  assert.ok(suspects.length >= 1);
  assert.strictEqual(rows.every((r) => r.category !== 'match'), true);
});

test('比对：对账单没有数据行时报错而不是给空结果', () => {
  assert.throws(
    () => reconcile({ statement: file('statement', []), shipments: [file('ship', [item({ name: '甲', code: 'P1' })])] }),
    /没有解析出任何数据行/,
  );
});

// ---------- v3：金额列识别优先级 ----------

test('表头识别：有多个金额候选时优先「应收增加」', () => {
  const rows = [
    ['书号', '摘要', '数量', '标准金额', '发行折扣', '应收增加', '总收款'],
    ['X1', '甲产品', '1', '100.00', '0.69', '69.00', ''],
  ];
  const info = detectHeader('对账单', rows, null);
  assert.strictEqual(info.columnLabels[info.columns.amount], '应收增加');
});

test('表头识别：订书单优先「发行实洋」而不是「码洋」「入账实洋」', () => {
  const rows = [
    ['数量', '书号', '书名', '价格', '码洋', '发行折扣', '发行实洋', '入账折扣', '入账实洋'],
    ['1', 'X1', '甲产品', '50', '50', '69%', '34.50', '0.60', '30.00'],
  ];
  const info = detectHeader('下单', rows, null);
  assert.strictEqual(info.columnLabels[info.columns.amount], '发行实洋');
});

// ---------- v3：总额判定 ----------

test('总额：多张单据金额累加后与对账单相等即一致', () => {
  const { summary } = reconcile({
    statement: file('statement', [
      item({ name: '甲', code: 'P1', amount: 6000n }),
      item({ name: '乙', code: 'P2', amount: 4000n }),
    ]),
    shipments: [
      file('ship1', [item({ name: '甲', code: 'P1', amount: 3000n })]),
      file('ship2', [item({ name: '甲', code: 'P1', amount: 3000n }), item({ name: '乙', code: 'P2', amount: 4000n })]),
    ],
  });
  assert.strictEqual(summary.total.statementAmount, '100.00');
  assert.strictEqual(summary.total.shipmentAmount, '100.00');
  assert.strictEqual(summary.total.amountDiff, '0.00');
  assert.strictEqual(summary.total.matched, true);
});

test('总额：差一分即不一致，并给出方向', () => {
  const { summary } = reconcile({
    statement: file('statement', [item({ name: '甲', code: 'P1', amount: 10000n })]),
    shipments: [file('ship', [item({ name: '甲', code: 'P1', amount: 9999n })])],
  });
  assert.strictEqual(summary.total.matched, false);
  assert.strictEqual(summary.total.amountDiff, '-0.01');
  assert.ok(summary.total.conclusion.includes('少'));
});

test('总额：有单据没参与核对时，即使金额相等也不能判为一致', () => {
  const { summary } = reconcile({
    statement: file('statement', [item({ name: '甲', code: 'P1', amount: 10000n })]),
    shipments: [
      file('ship1', [item({ name: '甲', code: 'P1', amount: 10000n })]),
      file('ship2', [item({ name: '甲', code: 'P1', amount: null })], { amountChecked: false }),
    ],
  });
  assert.strictEqual(summary.total.amountDiff, '0.00');
  assert.strictEqual(summary.total.matched, false, '有单据未参与核对，不能说总额一致');
  assert.strictEqual(summary.total.excludedFileCount, 1);
  assert.ok(summary.total.conclusion.includes('未参与核对'));
});

test('总额：没有一张单据提供金额时给出明确结论', () => {
  const { summary } = reconcile({
    statement: file('statement', [item({ name: '甲', code: 'P1', amount: 10000n })]),
    shipments: [file('ship', [item({ name: '甲', code: 'P1', amount: null })], { amountChecked: false })],
  });
  assert.strictEqual(summary.total.checkedFileCount, 0);
  assert.strictEqual(summary.total.matched, false);
  assert.ok(summary.total.conclusion.includes('无法核对'));
});

// ---------- v3：金额未知的产品 ----------

test('明细：只在无金额单据里出现的产品归为 unknown_amount，金额为 null', () => {
  const { rows, summary } = reconcile({
    statement: file('statement', [
      item({ name: '甲', code: 'P1', amount: 10000n }),
      item({ name: '乙', code: 'P2', amount: 20000n }),
    ]),
    shipments: [file('ship', [
      item({ name: '甲', code: 'P1', amount: 10000n }),
      item({ name: '乙', code: 'P2', amount: null }),
    ], { amountChecked: true })],
  });
  const byKey = Object.fromEntries(rows.map((r) => [r.key, r]));
  assert.strictEqual(byKey.P1.category, 'match');
  assert.strictEqual(byKey.P1.amountKnown, true);
  assert.strictEqual(byKey.P2.amountKnown, false);
  assert.strictEqual(byKey.P2.category, 'unknown_amount');
  assert.strictEqual(byKey.P2.shipmentAmount, null);
  assert.strictEqual(byKey.P2.amountDiff, null);
  assert.strictEqual(summary.unknownAmountCount, 1);
});
