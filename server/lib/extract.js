'use strict';

/**
 * 按列映射从网格里抽数据行，并做金额/数量归一。
 *
 * 真实数据里必须处理掉的行：
 *   对账单「期初 / 收款 / 合计」行、发货单尾部「合计」「采购签字：」行、
 *   xlsx 订书单第 1~2 行客户抬头 —— 否则汇总会多算。
 */

const { parseMoney, parseQty } = require('./values');
const { normalizeText } = require('./text');

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
