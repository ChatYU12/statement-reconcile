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
