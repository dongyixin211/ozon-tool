# Ozon Tool 企业级优化审计

## 已完成优化

- 页面架构：GUI 页面定义集中在 `App.PAGE_DEFINITIONS`，启动时只构建默认的“素材生成”页，其他页面首次点击时按需构建。
- 启动体验：macOS 启动脚本会优先选择 Tk 8.6+ 的 Python，规避系统 Python 3.9 / Tk 8.5 导致的白屏问题。
- 依赖治理：`requirements.txt` 只保留当前源码实际需要的 `Pillow` 和 `openpyxl`。
- 项目结构：移除了旧版打包源码副本、旧 Node Excel 脚本和静态 Excel 模板，主源码只保留在 `tool/`。
- 业务修复：富内容 JSON 图片与 `complex_attributes.image_url` 混合替换时共享图片顺序，避免 Ozon 上架图片错位。
- 配置服务化：配置文件路径、JSON 读写和损坏配置兜底已抽到 `tool/config_store.py`。
- API 路径治理：Ozon Seller API endpoint 集中在 `OZON_ENDPOINTS`，后续按官方文档核版本时有单一入口。
- API 边界防护：`/v3/product/import` 客户端层已限制非空且单次最多 100 项，避免未来批量调用越过平台约束。

## 当前架构现状

- `tool/app.py`：Tkinter GUI、页面构建、任务启动和日志调度仍在一个大类内。
- `tool/config_store.py`：本地配置文件路径、读取、写入和错误封装。
- `tool/batch_upload/`：Ozon Seller API、OSS 上传、上架、库存、条码、视频和已上架商品更新逻辑已经按业务模块拆分。
- `tool/builtin_mockups/`：本地场景图合成模板资产。
- `tool/exe_build/`：Windows PyInstaller 打包配置，当前 spec 直接引用 `tool/app.py`。

## Ozon API 对齐风险

代码中的 Ozon API 调用集中在 `OzonSellerClient`，endpoint 统一注册在 `OZON_ENDPOINTS`，每次请求统一发送 `Client-Id`、`Api-Key` 和 JSON body，符合 Seller API 的基本鉴权模式。

需要按官方文档持续核验的约束：

- Ozon Seller API 每次请求需要使用 API Key 和 Client ID。
- 通过 API 上传商品时，单次请求最多上传 100 项商品；当前 `import_products()` 已在客户端层防护。
- 库存更新相关帮助页提到仓库列表和库存更新方法。

需要后续逐项核对的接口版本：

- `/v3/product/import`：当前用于创建或更新商品；需持续对照 Ozon 官方 Seller API 文档确认版本和字段要求。
- `/v2/warehouse/list`：当前用于仓库列表；建议核对官方文档中最新仓库接口版本。
- `/v2/products/stocks`：当前用于库存更新；建议补充请求批量大小、限频和错误码测试。
- `/v1/barcode/generate`、`/v1/barcode/add`：当前用于条码生成和写入；建议补充真实错误响应样例解析。
- `/v4/product/info/attributes`、`/v4/product/info/stocks`：当前用于模板属性和库存信息；建议补充分页与字段缺失场景测试。

## 下一步企业级改造建议

- 拆分 GUI：把每个页面提取为独立类或模块，例如 `pages/generate.py`、`pages/upload.py`、`pages/inventory.py`。
- 配置服务化：继续把 `config.json` 的字段迁移、结构校验和敏感字段脱敏放入配置服务。
- API 客户端增强：为 Ozon API 增加统一错误对象、重试策略、限流处理和请求日志脱敏。
- 任务状态模型：将后台线程状态、取消、进度和错误汇总抽象为统一任务模型。
- UI 可用性：增加页面加载状态、任务进度条和禁用态，避免重复点击触发并发任务。
- 测试覆盖：补充 API 错误响应、配置迁移、页面懒加载和启动脚本的单元/集成测试。
