# API 契约 v3（前后端唯一依据）

> **v3 变更（口径调整）**：核对目标从「逐产品比数量+金额」改为**只看金额总额**。
> - 对账单金额取**「应收增加」**，发货单金额取**「发行实洋」**（两者实测分毫不差）。
> - **不再核对数量**，数量列仅用于判断空行，结果与导出里不出现数量。
> - 发货单**没有金额列**时，该文件**不参与金额核对**（不再按 0 计），其产品仍出现在明细里并标 `unknown_amount`。
> - 保留产品级明细用于定位差异，但**结论只看总额**：`summary.total.matched`。
>
> v2 背景（仍然有效）：支持 `.xls`；取消全局主键，改为每文件映射 7 个列位 + 服务端级联多字段匹配；结果行带 `matchBy` / `displayName`。

服务：`server/index.js`，默认端口 `PORT || 3000`，静态资源目录 `public/`，页面入口 `/`。
JSON 响应统一：成功 `{ "ok": true, ... }`，失败 `{ "ok": false, "error": { "code", "message", "detail" } }`，HTTP 4xx/5xx。

错误码：`NO_FILE`、`FILE_TOO_LARGE`、`BAD_FORMAT`、`PARSE_FAILED`、`MAPPING_MISSING`、`UPLOAD_EXPIRED`、`RESULT_EXPIRED`、`BAD_REQUEST`、`AUTH_FAILED`、`NO_DIFF`。

---

## 0. GET /api/config

```json
{
  "ok": true,
  "mode": "local",
  "corpId": "",
  "user": { "id": "local", "name": "本地模式" },
  "limits": { "maxFileSizeMB": 20, "maxShipments": 30 },
  "acceptExt": [".xlsx", ".xls"],
  "fieldLabels": {
    "name": "产品名称",
    "code": "编码1（书号/ISBN）",
    "code2": "编码2（货号/内部号）",
    "publisher": "出版社",
    "qty": "数量（不参与核对）",
    "amount": "金额（应收增加 / 发行实洋）",
    "unitPrice": "单价（不参与核对）"
  }
}
```

## 1. POST /api/auth/dingtalk

`{ "code": "..." }` → `{ "ok": true, "user": { "id", "name" } }`，并下发 `sid` cookie（HttpOnly，8 小时）。

## 2. POST /api/upload  (multipart/form-data)

字段：`statement`（1 个文件，必填）、`shipments`（1..N，`maxShipments` 以内）。接受 `.xlsx` / `.xls`（扩展名 + 文件头校验）。

```json
{
  "ok": true,
  "uploadId": "u_8f2c...",
  "files": [
    {
      "fileId": "f_1",
      "role": "statement",
      "fileName": "客户应收往来流水帐_29834域航.xls",
      "size": 30720,
      "sheets": ["客户应收往来流水帐"],
      "sheet": "客户应收往来流水帐",
      "headerRow": 1,
      "columns": { "name": 7, "code": 6, "code2": 8, "publisher": 9, "qty": 10, "amount": 14, "unitPrice": 11 },
      "columnLabels": { "6": "书号", "7": "摘要", "14": "应收增加" },
      "previewRows": [["客户代码", "序号"]],
      "previewStartRow": 1,
      "sheetInfos": [
        {
          "name": "客户应收往来流水帐",
          "rowCount": 23,
          "headerRow": 1,
          "columns": { "name": 7, "code": 6, "code2": 8, "publisher": 9, "qty": 10, "amount": 14, "unitPrice": 11 },
          "columnLabels": { "1": "客户代码", "6": "书号", "7": "摘要", "8": "货号/科目", "9": "品牌名称", "10": "数量", "11": "标准售价1", "12": "标准金额", "14": "应收增加" },
          "previewRows": [["客户代码", "序号", "..."]],
          "previewStartRow": 1,
          "warnings": []
        }
      ],
      "warnings": [],
      "error": null
    }
  ],
  "warnings": ["文件 xxx 的 Sheet2 为空，已跳过"]
}
```

- `columns` 的每个值都是 1-based 列号，识别不到为 `null`。
- **金额列只自动采用实洋列**：仅当表头存在 `应收增加` / `发行实洋` 时才填 `columns.amount`（同表多个候选时取实洋那列，不取最靠左的）。只有 `标准金额` / `码洋` / `金额` / `小计` / `合计` 时 **填 `null`**，并在该 sheet 的 `warnings` 里列出候选列、提示用户手动指定。
  > 自动识别**绝不把码洋当成实洋**。`amount` 为 `null` 是前端「不参与金额核对」提示的唯一判据，前端不需要再比对列名。
- `previewRows` 最多 8 行原始文本（从 `previewStartRow` 开始），行号连续，与 `columnLabels` 的键对应。
- 服务端在此阶段**不做汇总、不插图**，只解析；原文件存临时目录（TTL 2 小时）。

## 3. POST /api/sheet  （重新探测某个 sheet / 某个表头行）

```json
{ "uploadId": "u_...", "fileId": "f_1", "sheet": "下单", "headerRow": 3 }
```
`headerRow` 省略时用探测值。→ 与 `sheetInfos[i]` 同结构，外加 `{ "ok": true }`。切换 sheet 或改表头行后前端应调用一次，刷新 `columnLabels`/`previewRows`/`columns`。传了 `headerRow` 时响应里的 `headerRow` 原样回显。

## 4. POST /api/check  (application/json)

```json
{
  "uploadId": "u_...",
  "files": [
    {
      "fileId": "f_1",
      "sheet": "客户应收往来流水帐",
      "headerRow": 1,
      "columns": { "name": 7, "code": 6, "code2": 8, "publisher": 9, "qty": 10, "amount": 14, "unitPrice": 11 }
    }
  ]
}
```

校验规则（违反即 `MAPPING_MISSING`，`message` 说明是哪个文件缺哪个字段）：
- **对账单必须有 `amount`**：否则无法核对总额。
- **发货单 `amount` 可缺**：缺则该文件不参与金额核对，前端需在结果页明确列出「未参与核对的文件」。
- `name`、`code`、`code2` 至少有一个（用于产品级明细定位）。
- `qty` 不再必填，也不参与任何计算。

响应：

```json
{
  "ok": true,
  "resultId": "r_3a91...",
  "summary": {
    "identical": false,
    "total": {
      "statementAmount": "40135.23",
      "shipmentAmount": "38645.52",
      "amountDiff": "-1489.71",
      "matched": false,
      "conclusion": "总额不一致，发货单少 1489.71 元",
      "checkedFileCount": 6,
      "excludedFileCount": 0
    },
    "statement": {
      "fileId": "f_1",
      "fileName": "客户应收往来流水帐_29834域航.xls",
      "rows": 14,
      "amount": "40135.23",
      "amountMode": "金额列"
    },
    "shipmentFiles": [
      { "fileId": "f_2", "fileName": "订书-兰州域航（陇西）7.14.xlsx", "rows": 9, "amount": "16023.74", "amountChecked": true,  "amountMode": "金额列" },
      { "fileId": "f_3", "fileName": "天水卫校康养（文智达）7.24.xls", "rows": 8, "amount": null, "amountChecked": false, "amountMode": "无金额列（未参与核对）" }
    ],
    "shipmentCount": 2,
    "productCount": 14,
    "matchedCount": 0,
    "amountDiffCount": 0,
    "statementOnlyCount": 0,
    "shipmentOnlyCount": 0,
    "unknownAmountCount": 14,
    "diffCount": 0,
    "matchStats": { "code": 14, "code2": 0, "name": 0, "fuzzy": 0 },
    "warnings": ["天水卫校康养（文智达）7.24.xls 没有金额列，未参与金额核对"],
    "nameMerged": [{ "key": "a产品", "aliases": ["A产品", "Ａ 产品"] }]
  },
  "rows": [
    {
      "key": "9787564842901",
      "displayName": "婴幼儿心理发展与教育-十四五国规",
      "code": "9787564842901",
      "code2": "026Q081",
      "statementAmount": "2365.00",
      "shipmentAmount": "2365.00",
      "amountDiff": "0.00",
      "amountKnown": true,
      "category": "match",
      "matchBy": "code",
      "note": [],
      "sources": [
        { "fileId": "f_1", "fileName": "客户应收往来流水帐_29834域航.xls", "rows": [12] },
        { "fileId": "f_2", "fileName": "订书-兰州域航（陇西）7.14.xlsx", "rows": [11, 12] }
      ]
    }
  ],
  "suspects": [{ "left": "B产品-100", "right": "B产品 100", "reason": "仅一侧存在且名称高度相似（相似度 0.93）" }]
}
```

### 字段语义

- 金额一律是"元"的 2 位小数**字符串**；差额带负号，`0` 为 `"0.00"`。**前端不做任何算术。**
- `amountDiff = shipmentAmount − statementAmount`（发货单 − 对账单）。
- `category`：
  - `match` —— 金额一致
  - `amount_diff` —— 金额不一致
  - `statement_only` —— 仅对账单有
  - `shipment_only` —— 仅发货单有
  - `unknown_amount` —— 有发货记录，但都来自未参与金额核对的文件，金额未知
- `amountKnown=false` 时 `shipmentAmount` / `amountDiff` 为 `null`，前端显示「—」，**不得显示 0.00**。
- **没有 `conclusion` 字段**：结论只看 `summary.total.matched` 与 `summary.total.conclusion`。明细行的 `note` 只是定位线索。
- `rows` 排序：`amount_diff` → `statement_only` → `shipment_only` → `unknown_amount` → `match`。

### 总额一致性

`summary.total.matched === true` 当且仅当：
1. 至少有一个发货单参与了金额核对（`checkedFileCount ≥ 1`）；
2. 对账单金额合计 === 参与核对的发货单金额合计（分毫不差）。

不满足第 1 条时 `matched=false`，`conclusion="没有可参与金额核对的发货单"`。

## 5. GET /api/export/:resultId

- 成功：`200`，`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`，`Content-Disposition: attachment; filename="差异表-YYYYMMDD-HHmm.xlsx"; filename*=UTF-8''...`
- 结果过期：`400 RESULT_EXPIRED`
- **总额一致（`total.matched === true`）时**：`400 NO_DIFF`，不生成文件。

导出内容（sheet 名 `差异表`）：
- 第 1~4 行说明区：对账单文件名 / 参与核对的发货单清单（含张数）/ 核对时间与金额口径 / 总额对比与结论。
- 第 5 行表头，**8 列**：`产品编码`、`产品名称`、`编码2`、`对账单金额`、`发货单金额`、`金额差`、`匹配依据`、`备注`。
- 第 6 行起是明细行；`amountKnown=false` 的金额列写 `—` 文本。
- 末尾附「未参与金额核对的文件」和「解析提示」两段。

## 6. 前端行为要点

- **上传页**：对账单 1 个文件；发货单多选。文案「仅支持 .xlsx / .xls」。
- **映射页**：
  - 7 个列位都可选；**阻断校验只有两条**：对账单必须有金额列；每个文件必须有名称/编码1/编码2 之一。
  - 发货单缺金额列时给**非阻断提示**：「该文件没有金额列，将不参与金额核对」。
  - 数量、单价两个列位标注「不参与核对」。
- **结果页**：
  - 顶部**总额卡片**（第一视觉层级）：对账单金额、发货单金额、差额、结论（一致 / 不一致）。
  - 结论只来自 `summary.total`，不得由明细行推断。
  - 明确列出**参与核对**与**未参与核对**的发货单（文件名 + 行数 + 各自金额）。
  - 下方产品级明细，8 列，按 `category` 分组；`—` 表示金额未知。
  - 总额一致时：不显示下载按钮，`/api/export` 返回 `NO_DIFF`。
  - 展示 `matchStats`、`suspects`、`warnings`、`nameMerged`。
