'use strict';

const express = require('express');
const config = require('../config');
const { AppError } = require('../lib/errors');
const { loadGrid } = require('../lib/sheet');
const { detectHeader } = require('../lib/detect');
const { extractRows } = require('../lib/extract');
const { reconcile } = require('../lib/engine');
const { buildDiffWorkbook } = require('../lib/export');
const { newId } = require('../lib/store');
const { uploads, results } = require('../stores');

const router = express.Router();

const COLUMN_KEYS = ['name', 'code', 'code2', 'publisher', 'qty', 'amount', 'unitPrice'];
const COLUMN_LABELS = {
  name: '产品名称',
  code: '编码1（书号/ISBN）',
  code2: '编码2（货号/内部号）',
  publisher: '出版社',
  qty: '数量（不参与核对）',
  amount: '金额（应收增加 / 发行实洋）',
  unitPrice: '单价（不参与核对）',
};
const MAX_COL = 200;

function normalizeMapping(fileName, columns, role) {
  if (!columns || typeof columns !== 'object') {
    throw new AppError('MAPPING_MISSING', `文件「${fileName}」缺少列映射`);
  }
  const normalized = {};
  for (const key of COLUMN_KEYS) {
    const raw = columns[key];
    if (raw === null || raw === undefined || raw === '') {
      normalized[key] = null;
      continue;
    }
    const num = Number(raw);
    if (!Number.isInteger(num) || num < 1 || num > MAX_COL) {
      throw new AppError('MAPPING_MISSING', `文件「${fileName}」的「${COLUMN_LABELS[key]}」列号不合法：${raw}`);
    }
    normalized[key] = num;
  }
  // v3：只核金额。对账单必须有金额列；发货单可以没有（则该文件不参与金额核对）。
  if (!normalized.amount && role === 'statement') {
    throw new AppError('MAPPING_MISSING', `对账单「${fileName}」没有指定金额列（应收增加），无法核对总额`);
  }
  if (!normalized.name && !normalized.code && !normalized.code2) {
    throw new AppError('MAPPING_MISSING', `文件「${fileName}」没有指定产品名称或任一编码列（否则无法匹配产品）`);
  }
  return normalized;
}

function asyncHandler(fn) {
  return (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);
}

router.post(
  '/check',
  asyncHandler(async (req, res) => {
    const { uploadId, files } = req.body || {};
    if (!uploadId || typeof uploadId !== 'string') {
      throw new AppError('BAD_REQUEST', '缺少 uploadId，请重新上传文件');
    }
    // 先判断会话是否过期，前端据此提示"重新上传"而不是"补映射"
    const batch = uploads.get(uploadId);
    if (!batch) throw new AppError('UPLOAD_EXPIRED', '上传会话已过期（超过 2 小时），请重新上传文件');

    if (!Array.isArray(files) || files.length === 0) {
      throw new AppError('BAD_REQUEST', '缺少文件列映射配置');
    }

    const byId = new Map();
    for (const f of files) {
      if (!f || typeof f.fileId !== 'string') throw new AppError('BAD_REQUEST', '列映射缺少 fileId');
      if (byId.has(f.fileId)) throw new AppError('BAD_REQUEST', `文件 ${f.fileId} 的列映射重复`);
      byId.set(f.fileId, f);
    }
    if (byId.size !== batch.files.length) {
      throw new AppError('BAD_REQUEST', `有 ${batch.files.length} 个文件，只收到 ${byId.size} 个文件的列映射`);
    }

    const extracted = [];

    for (const bf of batch.files) {
      const conf = byId.get(bf.fileId);
      if (bf.error) {
        throw new AppError('PARSE_FAILED', `文件「${bf.fileName}」无法解析：${bf.error}`);
      }
      const columns = normalizeMapping(bf.fileName, conf.columns, bf.role);
      const headerRow = Number(conf.headerRow);
      if (!Number.isInteger(headerRow) || headerRow < 1 || headerRow > 5000) {
        throw new AppError('BAD_REQUEST', `文件「${bf.fileName}」的表头行不合法：${conf.headerRow}`);
      }

      let grid;
      try {
        grid = loadGrid(bf.path);
      } catch (e) {
        throw new AppError('PARSE_FAILED', `文件「${bf.fileName}」读取失败：${e.message}`);
      }
      const sheet = grid.sheets.find((s) => s.name === conf.sheet);
      if (!sheet || sheet.empty) {
        throw new AppError('BAD_REQUEST', `文件「${bf.fileName}」的工作表「${conf.sheet}」不存在或为空`);
      }

      // 表头是两行时（分组 + 字段），数据从 headerRow+2 开始
      const auto = detectHeader(sheet.name, sheet.rows, null);
      const headerSpan = auto.headerRow === headerRow && auto.headerSpan === 2 ? 2 : 1;

      const data = extractRows({
        rows: sheet.rows,
        headerRow,
        headerSpan,
        columns,
      });

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

    const statement = extracted.find((e) => e.role === 'statement');
    const shipments = extracted.filter((e) => e.role === 'shipment');
    if (!statement) throw new AppError('BAD_REQUEST', '缺少对账单文件');
    if (shipments.length === 0) throw new AppError('BAD_REQUEST', '缺少发货单文件');

    const { summary, rows, suspects } = reconcile({ statement, shipments });

    const resultId = newId('r_');
    results.set(resultId, { summary, rows, suspects, createdAt: Date.now() });

    res.json({ ok: true, resultId, summary, rows, suspects });
  }),
);

router.get(
  '/export/:resultId',
  asyncHandler(async (req, res) => {
    const result = results.get(req.params.resultId);
    if (!result) throw new AppError('RESULT_EXPIRED', '核对结果已过期，请重新核对');

    if (result.summary.total && result.summary.total.matched) {
      throw new AppError('NO_DIFF', '本次核对总额一致，不生成差异表');
    }

    const buffer = await buildDiffWorkbook(result);
    const d = new Date();
    const stamp = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}-${String(d.getHours()).padStart(2, '0')}${String(d.getMinutes()).padStart(2, '0')}`;
    const name = `差异表-${stamp}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="diff-${stamp}.xlsx"; filename*=UTF-8''${encodeURIComponent(name)}`,
    );
    res.setHeader('Cache-Control', 'no-store');
    res.send(buffer);
  }),
);

module.exports = { router, COLUMN_LABELS };
