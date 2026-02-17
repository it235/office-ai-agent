# Office AI 数据库说明

## 1. SQL 执行逻辑

### 执行时机
- **首次执行**：第一次调用 `OfficeAiDatabase.EnsureInitialized()` 时执行
- **后续执行**：同一进程内**不再执行**（通过 `_initialized` 标志控制）
- **新进程**：每次启动 Office 加载插件时，会重新执行一次（因进程重启后 `_initialized` 重置）

### 执行内容
- 迁移 SQL 使用 `CREATE TABLE IF NOT EXISTS`、`CREATE INDEX IF NOT EXISTS`，具有**幂等性**
- 多次执行不会报错，也不会重复建表
- 数据库文件路径：`%Documents%\OfficeAi\office_ai.db`

## 2. 安装后何时执行
- 用户安装插件后，**首次使用涉及记忆/ Skills 的功能**时触发初始化
- 例如：打开 Skills 配置、发送聊天消息（若启用 ContextBuilder）、收藏回答等
- 之前未使用时，数据库文件不会被创建

## 3. SQL 版本管理

### 当前策略
- **基准 schema（版本 1）**：在 `GetMigrationSql()` 中用 `CREATE TABLE IF NOT EXISTS` 建表，并创建 `schema_version` 表、写入版本 1。
- **增量升级**：新增字段/索引等一律通过 **ALTER TABLE / CREATE INDEX** 在 `RunVersionedMigrations()` 中按版本号执行，便于升级与版本控制。
- 每次 `EnsureInitialized()` 会：先执行基准 SQL，再根据 `schema_version.version` 只执行**未应用过的**迁移（如 2、3…），执行后更新 `schema_version`。

### 如何新增迁移（例如新增字段）
1. 在 `OfficeAiDatabase.RunVersionedMigrations()` 的 `migrations` 字典中增加一档，例如：
   - 版本 3：`"ALTER TABLE xxx ADD COLUMN yyy TEXT DEFAULT ''; UPDATE schema_version SET version = 3;"`
2. 不要改基准 SQL 里已有表结构（旧库已存在），只通过迁移脚本从当前版本升级到新版本。

### 发布新包时
- 无需用户手动执行 SQL。
- 新版本插件加载后，首次访问数据库时会自动执行到最新迁移版本。
- 建议在发布说明中注明「首次使用记忆功能时会初始化/升级本地数据库」。
