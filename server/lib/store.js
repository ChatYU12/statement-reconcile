'use strict';

const crypto = require('crypto');

/** 进程内 TTL 存储（上传批次 / 核对结果 / 会话）。单实例部署足够，过期自动清理。 */
class TtlStore {
  constructor(ttlMs, onExpire) {
    this.ttlMs = ttlMs;
    this.map = new Map();
    this.onExpire = onExpire;
    this.timer = setInterval(() => this.sweep(), Math.min(ttlMs, 10 * 60 * 1000));
    this.timer.unref();
  }

  set(id, value) {
    this.map.set(id, { value, expiresAt: Date.now() + this.ttlMs });
    return id;
  }

  get(id) {
    const entry = this.map.get(id);
    if (!entry) return null;
    if (entry.expiresAt < Date.now()) {
      this.drop(id);
      return null;
    }
    return entry.value;
  }

  drop(id) {
    const entry = this.map.get(id);
    if (!entry) return;
    this.map.delete(id);
    if (this.onExpire) {
      try {
        this.onExpire(id, entry.value);
      } catch {
        /* 清理失败不影响主流程 */
      }
    }
  }

  sweep() {
    const now = Date.now();
    for (const [id, entry] of this.map) {
      if (entry.expiresAt < now) this.drop(id);
    }
  }

  get size() {
    return this.map.size;
  }
}

function newId(prefix) {
  return `${prefix}${crypto.randomBytes(8).toString('hex')}`;
}

module.exports = { TtlStore, newId };
