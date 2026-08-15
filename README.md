# 贸易单证系统（发票 + 订舱单生成）

面向**传统贸易**的发票（Commercial Invoice）与订舱单（BOOKING FORM）生成工具。纯前端单页应用，数据本地存储（IndexedDB），无需后端即可离线使用；预留远程数据源适配器，有服务器时可在线反查/同步。

## 核心能力

- **导入聚水潭订单**：支持 xlsx 手动导入（API 自动拉取接口已占位，配置凭证后可开通）。
- **导入装箱清单并解析**：自动检测表头（中/英）、剔除合计空行、箱号续行、0 数量行过滤（源忠实）。
- **装箱清单直生订单**：不导入聚水潭订单时，可针对上传的装箱清单直接生成一条系统订单数据（`source=from_packing`），用于后续开票/订舱。
- **多单合并开票**：勾选多个订单合并为一票，强校验「箱单单号↔所选订单」「SKU↔数量」勾稽一致。
- **收发货人解析**：订单有则抓取填充模板；没有则从「收发货人」主数据取。
- **申报信息**：独立主数据，取自飞书《申报信息》主文档+备用文档（主文档优先、缺字段自动补齐），支持从镜像同步到本地；申报价统一折算 USD，币种见「币种」列。
- **模板管理**：上传 / 解析 / 预览 / 停用 / 删除；占位符引擎 `{{field}}` + `{{items.xxx}}` 明细行自动复制展开。
- **全场景自测**：内置 `selftest.html` 浏览器自测页，与 Node 端 `tests/run_tests.js` 同源（42 用例）。

## 架构（七层解耦）

| 层 | 模块 | 职责 |
|----|------|------|
| L1 向导/UI | `js/ui.js` | 9 页签、订单/箱单/发票/订舱单向导、CRUD |
| L2 数据源适配器 | `js/adapters.js` | Local / RemoteHttp（超时+降级）/ JstApi（占位） |
| L3 解析层 | `js/parser.js` | 聚水潭订单、装箱清单解析、表头扫描、hash 防重 |
| L4 主数据 | `js/db.js` + `js/seed.js` | IndexedDB 封装 + 首次种子 |
| L5 生成引擎 | `js/engine.js` | 数据组装、重量分摊、金额大写、模板填充 |
| L6 校验反查 | `js/validator.js` | 单号匹配、SKU 勾稽、申报反查、状态机 |
| L7 导出 | `js/exporter.js` | xlsx 缓冲 / 浏览器下载 |

## 本地使用

直接用浏览器打开 `index.html` 即可（建议 Chrome/Edge）。首次打开自动创建本地数据库并写入示例主数据（收发货人、申报要素、内置模板）。

## 运行测试

```bash
# Node 端全场景测试（需 exceljs）：
NODE_PATH=<node_modules> node tests/run_tests.js

# 浏览器端冒烟测试（需 jsdom + fake-indexeddb）：
NODE_PATH=<node_modules> node tests/smoke_test.js

# 应用内自测：浏览器打开 selftest.html，点击「运行全部测试」
```

## 部署

- GitHub Pages：`https://heryma99.github.io/trade-docs-system/`
- 国内镜像（CloudStudio / EdgeOne）：见交付说明，github.io 在国内访问可能不稳定，建议用国内镜像或系统浏览器硬刷新。

## 已知边界（第一期骨架）

- 内置发票/订舱单模板为通用骨架版，待用户提供真实样张后替换/增模板。
- 聚水潭 API 为占位适配器，需配置 `app_key/app_secret` 后实现签名拉单。
- 远程数据源需自行提供 `{store}.json` 端点；未配置时全部走本地。
