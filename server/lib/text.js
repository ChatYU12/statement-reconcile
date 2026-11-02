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
