'use strict';

/**
 * 钉钉工作台免登。
 *
 * 前端在钉钉容器里用 dd.runtime.permission.requestAuthCode 拿 code，
 * 服务端用 appKey/appSecret 换 access_token，再用 code 换 userid。
 * 50 人可见范围由钉钉后台「应用可见范围」控制，服务端不做二次鉴权。
 */

const config = require('../config');
const { AppError } = require('./errors');

const API = 'https://oapi.dingtalk.com';
let tokenCache = { value: '', expiresAt: 0 };

async function fetchJson(url, options) {
  const res = await fetch(url, options);
  const text = await res.text();
  let json;
  try {
    json = JSON.parse(text);
  } catch {
    throw new AppError('AUTH_FAILED', `钉钉接口返回非 JSON：${text.slice(0, 120)}`, null, 502);
  }
  return json;
}

async function getAccessToken() {
  if (tokenCache.value && tokenCache.expiresAt > Date.now()) return tokenCache.value;
  const url = `${API}/gettoken?appkey=${encodeURIComponent(config.appKey)}&appsecret=${encodeURIComponent(config.appSecret)}`;
  const json = await fetchJson(url);
  if (!json.access_token) {
    throw new AppError('AUTH_FAILED', `获取钉钉 access_token 失败：${json.errmsg || '未知错误'}`, { errcode: json.errcode }, 502);
  }
  tokenCache = { value: json.access_token, expiresAt: Date.now() + Math.max(60, (json.expires_in || 7200) - 300) * 1000 };
  return tokenCache.value;
}

/** code → { userId, name } */
async function getUserByCode(code) {
  if (!config.appKey || !config.appSecret) {
    throw new AppError('AUTH_FAILED', '服务端未配置钉钉应用凭据，无法免登', null, 500);
  }
  const token = await getAccessToken();
  const info = await fetchJson(`${API}/topapi/v2/user/getuserinfo?access_token=${encodeURIComponent(token)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ code }),
  });
  if (info.errcode && info.errcode !== 0) {
    throw new AppError('AUTH_FAILED', `钉钉免登失败：${info.errmsg || info.errcode}`, { errcode: info.errcode }, 401);
  }
  const result = info.result || {};
  const userId = result.userid || result.userId;
  if (!userId) throw new AppError('AUTH_FAILED', '钉钉未返回用户身份', null, 401);

  let name = result.name || '';
  if (!name) {
    try {
      const detail = await fetchJson(`${API}/topapi/v2/user/get?access_token=${encodeURIComponent(token)}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ userid: userId }),
      });
      name = (detail.result && detail.result.name) || '';
    } catch {
      /* 名字拿不到不影响免登 */
    }
  }
  return { userId, name: name || userId };
}

/** 清空 token 缓存（仅测试用） */
function resetTokenCache() {
  tokenCache = { value: '', expiresAt: 0 };
}

module.exports = { getUserByCode, resetTokenCache };
