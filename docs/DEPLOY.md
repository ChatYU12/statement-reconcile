# 钉钉工作台部署说明

目标形态：钉钉工作台里一个 H5 微应用，约 50 名销售可见，手机与电脑都能打开。

## 1. 服务端

### 运行环境

- Node.js ≥ 18（开发验证用 22）
- 无数据库、无外部服务；只有本地临时目录 `.tmp/`（上传文件与结果，默认 2 小时过期自动清理）

### 环境变量

| 变量 | 必填 | 说明 |
| --- | --- | --- |
| `PORT` | 否 | 监听端口，默认 `3000` |
| `HOST` | 否 | 默认 `0.0.0.0` |
| `DINGTALK_APP_KEY` | 钉钉模式必填 | 应用的 AppKey（原 SuiteKey） |
| `DINGTALK_APP_SECRET` | 钉钉模式必填 | 应用的 AppSecret |
| `DINGTALK_CORP_ID` | 钉钉模式必填 | 企业 CorpId（前端 `requestAuthCode` 需要） |
| `ALLOW_ANON` | 否 | `1` 时钉钉模式也允许未登录访问（**默认关闭**，别在生产开） |
| `MAX_FILE_SIZE_MB` | 否 | 单文件上限，默认 `20` |
| `MAX_SHIPMENTS` | 否 | 单次发货单上限，默认 `30` |
| `UPLOAD_TTL_MIN` | 否 | 上传会话有效期，默认 `120` 分钟 |
| `RESULT_TTL_MIN` | 否 | 核对结果（供下载）有效期，默认 `120` 分钟 |

**三个钉钉变量缺任意一个，服务自动进入本地模式**（免登跳过）。这是给开发和验收用的，不是降级方案。

### 启动

```bash
npm ci          # 或 npm install
DINGTALK_APP_KEY=xxx \
DINGTALK_APP_SECRET=yyy \
DINGTALK_CORP_ID=ding123 \
PORT=3000 \
node server/index.js
```

进程管理建议 systemd 或 pm2（`pm2 start server/index.js --name reconcile`）。服务是无状态单实例，重启不影响数据（只丢未过期的上传批次）。

### 反向代理（必须 HTTPS）

钉钉只在 HTTPS 页面下允许 JSAPI。用 Nginx：

```nginx
server {
    listen 443 ssl;
    server_name reconcile.example.com;

    ssl_certificate     /etc/ssl/reconcile.pem;
    ssl_certificate_key /etc/ssl/reconcile.key;

    client_max_body_size 40m;

    location / {
        proxy_pass http://127.0.0.1:3000;
        proxy_set_header Host $host;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_read_timeout 120s;
    }
}
```

## 2. 钉钉开放平台配置

1. **创建应用**：开放平台 → 应用开发 → 企业内部应用 → H5 微应用（或「网页应用」）。记下 **AppKey / AppSecret**。
2. **首页地址**：填 `https://reconcile.example.com/`。免登页面就是同一个地址，不需要单独的回调 URL。
3. **服务器出口 IP**：若企业开启了 IP 白名单，把服务器公网出口 IP 加进去（`oapi.dingtalk.com` 调用需要）。
4. **权限**：应用需要「通讯录个人信息读权限」相关的成员信息权限（用于 `topapi/v2/user/getuserinfo` 拿 userid，以及取姓名）。
5. **可见范围**：把 50 名销售加入。这是唯一的权限边界——服务端不做二次鉴权（需求里的约定：不按客户/区域隔离数据）。
6. **发布**：保存后发布版本，工作台才会出现入口。

## 3. 免登链路

```
浏览器（钉钉容器）
  1. GET /api/config            → { mode: "dingtalk", corpId }
  2. dd.runtime.permission.requestAuthCode({ corpId })
                                → code
  3. POST /api/auth/dingtalk { code }
       服务端：gettoken(appkey, appsecret) → access_token
               topapi/v2/user/getuserinfo(code) → userid
               topapi/v2/user/get(userid) → name
  4. 服务端下发 sid cookie（HttpOnly，8 小时）
  5. 之后所有 /api/* 请求带 cookie
```

未登录访问 `/api/*`（除 `config`/`health`/`auth`）返回 `401 AUTH_FAILED`，前端提示「登录已失效，请重新从钉钉工作台打开」。

**只能在钉钉客户端里验证免登**，浏览器直接打开会走不到 `window.dd`，前端会显示可关闭的提示并停在页面上（本地模式下没有这个限制）。

## 4. 上线前检查清单

- [ ] `https://<域名>/api/health` 返回 `{"ok":true,"mode":"dingtalk"}`
- [ ] 钉钉工作台点开应用 → 顶部徽标显示真实姓名（不是「本地模式」）
- [ ] 传 1 份真实对账单 + 至少 2 份真实发货单，确认自动识别出的表头行与 7 个字段列正确
- [ ] 核对结果里 99% 的产品落在「一致」或「数值不一致」，没有整批「仅对账单有」（若有，说明匹配字段没选对，回映射页调整）
- [ ] 下载差异表，用 Excel 打开确认列名与数字格式正常
- [ ] 构造一次「完全一致」的上传，确认页面显示「无差异」且没有下载按钮
- [ ] 用一名销售的真实账号在手机钉钉里跑一遍
- [ ] 确认 `.tmp/` 所在磁盘有空间（上传文件按 2 小时 TTL 清理）

## 5. 常见问题

**Q：提示「登录已失效」但刚打开？**
A：`sid` cookie 是 HttpOnly + SameSite=Lax，必须同域名 HTTPS；如果中间有代理改写了域名或剥离了 Set-Cookie，会丢会话。

**Q：某张发货单识别不出表头？**
A：映射页卡片会标红并说明原因（如「未识别到数量列」）。手动填表头行（1–200）后会自动重新探测；仍不行就逐列下拉指定。每张表的映射会按 sheet 记住。

**Q：结果里出现大量「仅对账单有」？**
A：三种可能：① 那次核对的发货单没传齐（对账单含多个学校）；② 匹配只用了名称而名称写法差异大——把编码列补上；③ 对账单的编码与发货单的编码是两套体系（如对账单用 ISBN、发货单用内部书号），此时需要把两列都对上。

**Q：金额差明明没有却报差异？**
A：先看结果页的总额卡片和参与核对清单。常见原因有三种：

1. **该发货单没有金额列**（`.xls` 老发货单只有定价），它不参与金额核对，会出现在「未参与金额核对」清单里。这种单据不能用来核金额，得传带「发行实洋」的订书单。
2. **金额列选错了**。对账单要选「应收增加」，发货单要选「发行实洋」。选成「标准金额」或「码洋」的话，比的是定价而不是实洋，差额会很大。
3. **漏传了发货单**。对账单是按客户汇总的，发货单是按学校/流向的，漏一个就差额。缺的产品会在明细里显示成「仅对账单有」。

明细行「备注」里会写明这一行属于哪种情况，以及差值方向和金额。

**Q：能改端口/上限吗？**
A：见上方环境变量表。
