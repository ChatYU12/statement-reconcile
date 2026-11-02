'use strict';

const fs = require('fs/promises');
const path = require('path');
const config = require('./config');
const { TtlStore } = require('./lib/store');

/** 上传批次（含临时文件目录），过期时把目录删掉 */
const uploads = new TtlStore(config.uploadTtlMs, async (id, value) => {
  if (value && value.dir) {
    await fs.rm(value.dir, { recursive: true, force: true }).catch(() => {});
  }
});

/** 核对结果（差异表数据），导出时复用 */
const results = new TtlStore(config.resultTtlMs);

/** 登录会话 */
const sessions = new TtlStore(config.sessionTtlMs);

async function ensureUploadDir(uploadId) {
  const dir = path.join(config.tmpDir, 'uploads', uploadId);
  await fs.mkdir(dir, { recursive: true });
  return dir;
}

module.exports = { uploads, results, sessions, ensureUploadDir };
