# Office AI 数据库说明

## 1. SQL 执行逻辑

### 执行时机
- **首次执行**：第一次调用 `OfficeAiDatabase.EnsureInitialized()` 时执行
- **后续执行**：同一进程内**不再执行**（通过 `_initialized` 标志控制）
- **新进程**：每次启动 Office 加载插件时，会重新执行一次（因进程重启后 `_initialized` 重置）

### 执行内容
- 迁移 SQL 使用 `CREATE TABLE IF NOT EXISTS`、`CREATE INDEX IF NOT EXISTS`，具有**幂等性**
- 多次执行不会报错，也不会重复建表
- 数据库文件路径：`%Documents%\OfficeAiAppData\office_ai.db`；从 `bin\Debug` 等本地调试目录运行时改用
  `%Documents%\OfficeAiAppData-Debug\office_ai.db`，与安装版数据分离（见 `OfficeAiDatabase.GetDatabasePath`）

## 2. 安装后何时执行
- 用户安装插件后，**首次使用涉及记忆/ Skills 的功能**时触发初始化
- 例如：打开 Skills 配置、发送聊天消息（若启用 ContextBuilder）、收藏回答等
- 之前未使用时，数据库文件不会被创建

## 3. SQL 版本管理

### 当前策略
- **基准 schema（版本 1）**：在 `GetMigrationSql()` 中用 `CREATE TABLE IF NOT EXISTS` 建表，并创建 `schema_version` 表、写入版本 1。
- **增量升级**：新增字段/索引等一律通过 **ALTER TABLE / CREATE INDEX** 在 `RunVersionedMigrations()` 中按版本号执行，便于升级与版本控制。
- 每次 `EnsureInitialized()` 会：先执行基准 SQL，再根据 `schema_version.version` 只执行**未应用过的**迁移（如 2、3…），执行后更新 `schema_version`。
- **当前最新版本：19**。版本 11 新增 `agent_run` / `agent_run_step` 轻量 RunTrace 表；版本 19（E2-04）把带 `app_type`
  的六张表收口到同一份归属口径（见第 4 节，纯 UPDATE、只填空、可重复执行，因此不需要改 `OfficeAiDbSchema.current.sql`）。
  新增迁移后必须同步 `OfficeAiDbSchema.current.sql`、`scripts/smoke-empty-db-initialization.ps1` 和 schema drift smoke。

### 如何新增迁移（例如新增字段）
1. 在 `OfficeAiDatabase.RunVersionedMigrations()` 的 `migrations` 字典中增加一档，例如：
   - 版本 3：`"ALTER TABLE xxx ADD COLUMN yyy TEXT DEFAULT ''; UPDATE schema_version SET version = 3;"`
2. 不要改基准 SQL 里已有表结构（旧库已存在），只通过迁移脚本从当前版本升级到新版本。
3. 如新增独立 SQL 文件，放在 `ShareRibbon/Storage/Migrations/NNN_name.sql`，并保持内联迁移、当前 schema snapshot 与 smoke 期望一致。
4. 迁移文件是 UTF-8 **无 BOM**、**LF**；由 `<Content Include="Storage\Migrations\*.sql">`（PreserveNewest）自动部署，不需要改 `.vbproj`。
5. 数据回填型迁移必须**只填空、不覆盖已解析的值**，并在 `scripts/smoke-empty-db-initialization.ps1` 里用 legacy 形状夹具
   跑一遍、再**重跑已部署的迁移文本**证明可重复执行；同时把 `schema_version` 期望值加一档。

## 4. 宿主归属列 `app_type`

带这一列的六张表：`conversation_event`、`session_summary`、`agent_run`、`agent_run_change`、`atomic_memory`、`memory_item`。
读侧一律是「等于当前宿主 **或** 空串（NULL/'' = 通用数据，三端都可见）」，所以：

- **合法值只有四种**：`Word` / `Excel` / `PowerPoint` / `''`。第三种值（例如早期构建落库的 `"Unknown"`）既不等于任何宿主、
  又拿不到空串豁免，那条数据会在三端同时查不到且不报错——现象是「升级后历史/记忆整段消失」。
- **归一只发生在写入边界**：统一走 `AppTypeContract.Normalize`（`Storage/AppTypeContract.vb`）。读侧过滤参数不要归一，
  那里 `''` 本来就是豁免，归一会把「按宿主精确查」变成「查全部」。
- **不要为了「老数据别丢」放宽读过滤**：那会把别的宿主的历史漏进当前端。正确做法是迁移 019 那种按证据强度回填。
- 新增第六处写入点或第七张带 `app_type` 的表时，`scripts/audit-p0-guardrails.ps1` 的 E2-04 段会要求：该文件引用
  `AppTypeContract.Normalize`，且迁移 019 的文件与内联兜底都逐表含归一语句。

### 发布新包时
- 无需用户手动执行 SQL。
- 新版本插件加载后，首次访问数据库时会自动执行到最新迁移版本。
- 建议在发布说明中注明「首次使用记忆功能时会初始化/升级本地数据库」。
