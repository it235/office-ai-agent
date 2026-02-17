# Skills Config 能力规格

## ADDED Requirements

### Requirement: 系统提示词按场景管理

系统 SHALL 按场景（excel/word/ppt/common）维护系统级提示词。系统 SHALL 在构建上下文时，根据当前宿主类型加载对应场景的系统提示词并注入到 [1] 层。

#### Scenario: 加载系统提示词

- **WHEN** 构建上下文且当前宿主为 Excel
- **THEN** 系统 MUST 从 prompt_template 表读取 scenario=excel 且 is_skill=0 的记录，将 content 注入 [1] 层

#### Scenario: 编辑系统提示词

- **WHEN** 用户在配置窗口中编辑某场景的系统提示词并保存
- **THEN** 系统 MUST 更新 prompt_template 表中对应记录

---

### Requirement: Skills 增删改查与导入

系统 SHALL 支持 Skills 的创建、读取、更新、删除。系统 SHALL 支持从 JSON 或 Markdown 文件导入 Skill。Skill MUST 包含 skillName、description、promptTemplate、supportedApps、parameters（可选）。

#### Scenario: 添加 Skill

- **WHEN** 用户在配置窗口中添加新 Skill 并保存
- **THEN** 系统 MUST 将 Skill 写入 prompt_template 表（is_skill=1），parameters 与 supported_apps 存入 extra_json

#### Scenario: 从 JSON 导入 Skill

- **WHEN** 用户选择 JSON 文件并执行导入
- **THEN** 系统 MUST 解析 JSON 中的 skillName、description、promptTemplate、supportedApps、parameters，并写入 prompt_template 表

#### Scenario: 删除 Skill

- **WHEN** 用户删除某 Skill
- **THEN** 系统 MUST 从 prompt_template 表移除或标记该记录

---

### Requirement: Skills 按场景与宿主加载

系统 SHALL 在构建上下文时，根据当前宿主（Excel/Word/PPT）加载已启用的、supported_apps 包含当前宿主的 Skills。系统 SHALL 将 Skills 的 promptTemplate 与系统提示词一起注入 [1] 层。

#### Scenario: 按宿主过滤 Skills

- **WHEN** 当前宿主为 Excel，构建上下文
- **THEN** 系统 MUST 仅加载 supported_apps 包含 "Excel" 的 Skills

---

### Requirement: 变量替换

系统 SHALL 在运行时对提示词与 Skill 模板中的占位符进行变量替换。占位符格式为 `{{变量名}}`，如 `{{选中内容}}`、`{{operation}}`、`{{selectedCells}}`。系统 SHALL 由调用方提供实际值并替换。

#### Scenario: 替换选中内容

- **WHEN** 模板包含 `{{选中内容}}` 且用户已选中单元格
- **THEN** 系统 MUST 将占位符替换为实际选中内容文本

#### Scenario: 未提供变量时的处理

- **WHEN** 模板包含 `{{operation}}` 但调用方未提供该参数
- **THEN** 系统 SHALL 替换为空字符串或占位符原文（可配置）

---

### Requirement: 配置窗口入口

系统 SHALL 在 Ribbon「提示词配置」组提供「提示词配置」按钮。点击后 SHALL 弹出配置窗口，展示系统提示词与 Skills 列表，支持按场景筛选。

#### Scenario: 打开配置窗口

- **WHEN** 用户点击 Ribbon「提示词配置」按钮
- **THEN** 系统 MUST 弹出配置窗口，展示当前场景下的提示词与 Skills 列表
