'use strict';

const fs = require('fs/promises');
const path = require('path');
const express = require('express');
const multer = require('multer');

const config = require('../config');
const { AppError } = require('../lib/errors');
const { loadGrid } = require('../lib/sheet');
const { detectHeader } = require('../lib/detect');
const { newId } = require('../lib/store');
const { uploads, ensureUploadDir } = require('../stores');

const router = express.Router();

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: config.maxFileSizeMB * 1024 * 1024, files: config.maxShipments + 1 },
});

/** 用文件头判断真实格式，防止把 .xls 改个名字成 .xlsx 导致解析失败 */
function sniff(buffer) {
  if (!buffer || buffer.length < 8) return null;
  if (buffer[0] === 0x50 && buffer[1] === 0x4b && buffer[2] === 0x03 && buffer[3] === 0x04) return 'xlsx';
  if (buffer[0] === 0xd0 && buffer[1] === 0xcf && buffer[2] === 0x11 && buffer[3] === 0xe0) return 'xls';
  return null;
}

function analyzeFile(filePath, fileName) {
  const grid = loadGrid(filePath);
  const sheetInfos = [];
  for (const sheet of grid.sheets) {
    if (sheet.empty) continue;
    const info = detectHeader(sheet.name, sheet.rows, null);
    sheetInfos.push(info);
  }
  if (sheetInfos.length === 0) {
    throw new AppError('PARSE_FAILED', `文件「${fileName}」里没有找到任何有数据的表`, { fileName });
  }
  const primary = sheetInfos[0];
  return {
    sheet: primary.sheet,
    sheets: sheetInfos.map((s) => s.sheet),
    sheetInfos,
    // 顶层也放一份默认 sheet 的探测结果，方便只关心单 sheet 的调用方直接用
    headerRow: primary.headerRow,
    columns: primary.columns,
    columnLabels: primary.columnLabels,
    previewRows: primary.previewRows,
    previewStartRow: primary.previewStartRow,
    warnings: [],
    error: null,
  };
}

router.post('/upload', (req, res, next) => {
  upload.fields([
    { name: 'statement', maxCount: 1 },
    { name: 'shipments', maxCount: config.maxShipments },
  ])(req, res, (err) => {
    if (err) return next(err);
    handleUpload(req, res).catch(next);
  });
});

async function handleUpload(req, res) {
  const statementFiles = (req.files && req.files.statement) || [];
  const shipmentFiles = (req.files && req.files.shipments) || [];

  if (statementFiles.length === 0) {
    throw new AppError('NO_FILE', '请先选择对账单文件');
  }
  if (shipmentFiles.length === 0) {
    throw new AppError('NO_FILE', '请至少选择一个发货单文件');
  }
  if (shipmentFiles.length > config.maxShipments) {
    throw new AppError('BAD_REQUEST', `发货单最多 ${config.maxShipments} 个，本次收到 ${shipmentFiles.length} 个`);
  }

  const all = [
    { role: 'statement', file: statementFiles[0] },
    ...shipmentFiles.map((f) => ({ role: 'shipment', file: f })),
  ];

  const uploadId = newId('u_');
  const dir = await ensureUploadDir(uploadId);
  const warnings = [];
  const files = [];

  for (let i = 0; i < all.length; i++) {
    const { role, file } = all[i];
    const fileId = `f_${i + 1}`;
    const originalName = Buffer.from(file.originalname, 'latin1').toString('utf8');
    const ext = path.extname(originalName).toLowerCase();
    const safeName = `${fileId}${ext}`;
    const filePath = path.join(dir, safeName);

    const entry = {
      fileId,
      role,
      fileName: originalName,
      size: file.size,
      sheet: null,
      sheets: [],
      sheetInfos: [],
      warnings: [],
      error: null,
    };

    const format = sniff(file.buffer);
    if (ext !== '.xlsx' && ext !== '.xls') {
      entry.error = `不支持的格式「${ext || '无扩展名'}」，请上传 .xlsx 或 .xls`;
      files.push(entry);
      continue;
    }
    if (!format) {
      entry.error = '文件内容不是有效的 Excel（可能已损坏或实际是 CSV），请用 Excel 另存为 .xlsx 后重试';
      files.push(entry);
      continue;
    }

    await fs.writeFile(filePath, file.buffer);

    try {
      const info = analyzeFile(filePath, originalName);
      Object.assign(entry, info);
      entry.warnings = info.sheetInfos.flatMap((s) => s.warnings || []).slice(0, 6);
    } catch (e) {
      entry.error = e instanceof AppError ? e.message : `解析「${originalName}」失败：${e.message}`;
    }

    if (!entry.error && entry.sheets.length > 1) {
      warnings.push(`文件「${originalName}」有 ${entry.sheets.length} 个工作表，当前使用「${entry.sheet}」`);
    }

    files.push(entry);
  }

  uploads.set(uploadId, {
    dir,
    createdAt: Date.now(),
    files: files.map((f, i) => ({
      fileId: f.fileId,
      role: f.role,
      fileName: f.fileName,
      path: path.join(dir, `${f.fileId}${path.extname(f.fileName).toLowerCase()}`),
      ext: path.extname(f.fileName).toLowerCase(),
      error: f.error,
    })),
    meta: files,
  });

  res.json({ ok: true, uploadId, files, warnings });
}

/** 切换工作表 / 改表头行后重新探测 */
router.post('/sheet', async (req, res, next) => {
  try {
    await handleSheet(req, res);
  } catch (err) {
    next(err);
  }
});

async function handleSheet(req, res) {
  const { uploadId, fileId } = req.body || {};
  const batch = uploads.get(uploadId);
  if (!batch) throw new AppError('UPLOAD_EXPIRED', '上传会话已过期，请重新上传文件');

  const target = batch.files.find((f) => f.fileId === fileId);
  if (!target) throw new AppError('BAD_REQUEST', '未找到对应的文件，请重新上传');

  const sheetName = String(req.body.sheet || '').trim();
  const headerRowRaw = req.body.headerRow;
  let headerRow = null;
  if (headerRowRaw !== undefined && headerRowRaw !== null && headerRowRaw !== '') {
    headerRow = Number(headerRowRaw);
    if (!Number.isInteger(headerRow) || headerRow < 1 || headerRow > 5000) {
      throw new AppError('BAD_REQUEST', '表头行必须是 1~5000 的整数');
    }
  }

  const grid = loadGrid(target.path);
  const sheet = grid.sheets.find((s) => s.name === sheetName);
  if (!sheet) {
    throw new AppError('BAD_REQUEST', `工作表「${sheetName}」不存在或为空`, { sheets: grid.sheets.filter((s) => !s.empty).map((s) => s.name) });
  }

  const info = detectHeader(sheet.name, sheet.rows, headerRow);
  res.json({ ok: true, ...info });
}

module.exports = { router, sniff };
