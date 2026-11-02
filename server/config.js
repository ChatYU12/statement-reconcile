'use strict';

/** 运行配置：全部通过环境变量注入，未配置钉钉凭据时自动进入本地模式 */
const path = require('path');

const appKey = process.env.DINGTALK_APP_KEY || '';
const appSecret = process.env.DINGTALK_APP_SECRET || '';
const corpId = process.env.DINGTALK_CORP_ID || '';

const hasDingtalk = Boolean(appKey && appSecret && corpId);

module.exports = {
  port: Number(process.env.PORT || 3000),
  host: process.env.HOST || '0.0.0.0',
  mode: hasDingtalk ? 'dingtalk' : 'local',
  appKey,
  appSecret,
  corpId,
  // 钉钉模式下是否允许匿名继续（默认不允许，避免绕过 50 人可见范围）
  allowAnon: process.env.ALLOW_ANON === '1',
  maxFileSizeMB: Number(process.env.MAX_FILE_SIZE_MB || 20),
  maxShipments: Number(process.env.MAX_SHIPMENTS || 30),
  uploadTtlMs: Number(process.env.UPLOAD_TTL_MIN || 120) * 60 * 1000,
  resultTtlMs: Number(process.env.RESULT_TTL_MIN || 120) * 60 * 1000,
  sessionTtlMs: 8 * 60 * 60 * 1000,
  tmpDir: path.join(__dirname, '..', '.tmp'),
  publicDir: path.join(__dirname, '..', 'public'),
};
