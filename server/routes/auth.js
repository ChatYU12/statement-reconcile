'use strict';

const express = require('express');
const config = require('../config');
const { AppError } = require('../lib/errors');
const { getUserByCode } = require('../lib/dingtalk');
const { newId } = require('../lib/store');
const { sessions } = require('../stores');

const router = express.Router();

router.post('/auth/dingtalk', async (req, res, next) => {
  try {
    const code = String((req.body && req.body.code) || '').trim();
    if (!code) throw new AppError('BAD_REQUEST', '缺少钉钉免登 code');

    const user = await getUserByCode(code);
    const sid = newId('s_');
    sessions.set(sid, { user });

    res.setHeader(
      'Set-Cookie',
      `sid=${sid}; HttpOnly; Path=/; Max-Age=${Math.floor(config.sessionTtlMs / 1000)}; SameSite=Lax`,
    );
    res.json({ ok: true, user: { id: user.userId, name: user.name } });
  } catch (err) {
    next(err);
  }
});

module.exports = router;
