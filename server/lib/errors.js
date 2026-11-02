'use strict';

/** 带错误码的业务异常，路由层统一转成 JSON 错误响应 */
class AppError extends Error {
  constructor(code, message, detail = null, status = 400) {
    super(message);
    this.name = 'AppError';
    this.code = code;
    this.detail = detail;
    this.status = status;
  }
}

module.exports = { AppError };
