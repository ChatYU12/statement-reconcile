'use strict';

/**
 * 对账单 × 发货单 自动核对工具
 *
 * 单进程 Node 服务：托管 public/ 静态页 + 提供上传/核对/导出 API。
 * 钉钉凭据通过环境变量注入；未配置时进入本地模式（无需登录，便于验收）。
 */

const path = require('path');
const express = require('express');
const multer = require('multer');

const config = require('./config');
const { AppError } = require('./lib/errors');
const { sessions } = require('./stores');
const uploadRoutes = require('./routes/upload');
const checkRoutes = require('./routes/check');
const authRoutes = require('./routes/auth');

const app = express();
app.disable('x-powered-by');
app.set('trust proxy', true);

app.use(express.json({ limit: '2mb' }));
app.use(express.urlencoded({ extended: false, limit: '2mb' }));

// 极简 cookie 解析（只有 sid 一个用途，不引入 cookie-parser）
app.use((req, _res, next) => {
  const raw = req.headers.cookie || '';
  const m = /(?:^|;\s*)sid=([^;]+)/.exec(raw);
  req.sid = m ? decodeURIComponent(m[1]) : null;
  next();
});

// 静态资源（H5 页面）
app.use(express.static(config.publicDir, { index: 'index.html', maxAge: '5m' }));

// 会话校验：本地模式放行；钉钉模式必须持有效会话
const OPEN_PATHS = new Set(['/config', '/health', '/auth/dingtalk']);
app.use('/api', (req, res, next) => {
  if (OPEN_PATHS.has(req.path)) return next();
  if (config.mode === 'local') {
    req.user = { id: 'local', name: '本地模式' };
    return next();
  }
  const session = req.sid ? sessions.get(req.sid) : null;
  if (session) {
    req.user = session.user;
    return next();
  }
  if (config.allowAnon) {
    req.user = { id: 'anon', name: '未登录' };
    return next();
  }
  return res.status(401).json({
    ok: false,
    error: { code: 'AUTH_FAILED', message: '登录已失效，请重新从钉钉工作台打开本应用' },
  });
});

app.get('/api/config', (req, res) => {
  let user = { id: '', name: '' };
  if (config.mode === 'local') {
    user = { id: 'local', name: '本地模式' };
  } else if (req.sid) {
    const session = sessions.get(req.sid);
    if (session) user = { id: session.user.userId, name: session.user.name };
  }
  res.json({
    ok: true,
    mode: config.mode,
    corpId: config.corpId,
    user,
    limits: { maxFileSizeMB: config.maxFileSizeMB, maxShipments: config.maxShipments },
    acceptExt: ['.xlsx', '.xls'],
    fieldLabels: {
      name: '产品名称',
      code: '编码1（书号/ISBN）',
      code2: '编码2（货号/内部号）',
      publisher: '出版社',
      qty: '数量（不参与核对）',
      amount: '金额（应收增加 / 发行实洋）',
      unitPrice: '单价（不参与核对）',
    },
  });
});

app.get('/api/health', (_req, res) => {
  res.json({ ok: true, mode: config.mode, uptime: Math.round(process.uptime()) });
});

app.use('/api', authRoutes);
app.use('/api', uploadRoutes.router);
app.use('/api', checkRoutes.router);

app.use('/api', (_req, res) => {
  res.status(404).json({ ok: false, error: { code: 'NOT_FOUND', message: '接口不存在' } });
});

// 统一错误出口
app.use((err, _req, res, next) => {
  if (res.headersSent) return next(err);

  if (err instanceof AppError) {
    return res.status(err.status).json({
      ok: false,
      error: { code: err.code, message: err.message, detail: err.detail || null },
    });
  }
  if (err instanceof multer.MulterError) {
    const tooLarge = err.code === 'LIMIT_FILE_SIZE';
    const tooMany = err.code === 'LIMIT_FILE_COUNT' || err.code === 'LIMIT_UNEXPECTED_FILE';
    return res.status(400).json({
      ok: false,
      error: {
        code: tooLarge ? 'FILE_TOO_LARGE' : 'BAD_REQUEST',
        message: tooLarge
          ? `单个文件不能超过 ${config.maxFileSizeMB}MB`
          : tooMany
            ? `文件数量超出限制（发货单最多 ${config.maxShipments} 个）`
            : `上传失败：${err.message}`,
      },
    });
  }
  if (err && err.type === 'entity.too.large') {
    return res.status(413).json({ ok: false, error: { code: 'BAD_REQUEST', message: '提交内容过大' } });
  }
  console.error('[unhandled]', err);
  return res.status(500).json({
    ok: false,
    error: { code: 'INTERNAL', message: `服务处理异常：${err && err.message ? err.message : '未知错误'}` },
  });
});

if (require.main === module) {
  const server = app.listen(config.port, config.host, () => {
    const modeText = config.mode === 'dingtalk' ? '钉钉模式（需要免登）' : '本地模式（未配置钉钉凭据，免登跳过）';
    console.log(`对账单核对服务已启动：http://127.0.0.1:${config.port}  ${modeText}`);
    console.log(`静态目录：${path.relative(process.cwd(), config.publicDir)}  临时目录：${path.relative(process.cwd(), config.tmpDir)}`);
  });
  server.on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
      console.error(`端口 ${config.port} 已被占用，请设置 PORT 环境变量换一个端口`);
      process.exit(1);
    }
    throw err;
  });
}

module.exports = app;
