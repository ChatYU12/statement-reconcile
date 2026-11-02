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

const { formatCents } = require('./values');
const { normalizeText, normalizeCode, similarity, publisherCompatible } = require('./text');
const { AppError } = require('./errors');

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
