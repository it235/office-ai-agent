# Memory 能力规格

## ADDED Requirements

### Requirement: 原子记忆存储与检索

系统 SHALL 提供原子记忆（atomic memory）的存储与检索能力。每条原子记忆 MUST 包含 id、timestamp、content（3～4 句完整语义）、tags（JSON 或逗号分隔）、session_id（可选）、create_time。系统 SHALL 支持按当前 query 进行被动 RAG 检索，返回 top-N 条相关记忆。

#### Scenario: 被动 RAG 检索原子记忆

- **WHEN** 用户发送消息，系统构建上下文
- **THEN** 系统 MUST 调用 MemoryService.GetRelevantMemories(currentQuery, topN)，将检索结果注入上下文 [3] 层

#### Scenario: 原子记忆异步写入

- **WHEN** 流式响应完成后，后台任务分析 (userPrompt, assistantReply)
- **THEN** 若包含值得记录的新信息，系统 MUST 异步写入 atomic_memory 表；可做去重/冲突检测

---

### Requirement: 用户画像

系统 SHALL 提供用户画像（user profile）的结构化存储与加载。用户画像 MUST 为可见、可编辑的文档，包含回复风格偏好、语言习惯、主要背景等。系统 SHALL 在每次请求时自动加载用户画像并注入到 system 或 [1] 层。

#### Scenario: 加载用户画像

- **WHEN** 构建上下文时
- **THEN** 若 memory.enable_user_profile 为真，系统 MUST 调用 MemoryService.GetUserProfile() 并将结果注入上下文

#### Scenario: 更新用户画像

- **WHEN** 异步记忆处理发现 user prompt 包含偏好或风格相关信息
- **THEN** 系统 MUST 更新 user_profile 表

---

### Requirement: 近期会话摘要

系统 SHALL 存储并按需检索近期会话摘要（session_summary）。每条摘要 MUST 包含 session_id、title、snippet、created_at。系统 SHALL 在每次请求时获取近期摘要并注入到上下文 [4] 层。

#### Scenario: 注入近期会话摘要

- **WHEN** 构建上下文时
- **THEN** 系统 MUST 调用 MemoryService.GetRecentSessionSummaries(limit)，将结果注入 [4] 层

#### Scenario: 新会话首条触发摘要写入

- **WHEN** 新会话首条消息得到回复后
- **THEN** 系统 MAY 触发会话摘要写入 session_summary 表

---

### Requirement: 会话持久化

系统 SHALL 将每条 user/assistant 消息持久化到 conversation 表。每条记录 MUST 包含 session_id、role、content、create_time、is_collected。新建会话时 SHALL 生成唯一 session_id（如 GUID）。

#### Scenario: 消息持久化

- **WHEN** 用户发送消息或收到 assistant 回复
- **THEN** 系统 MUST 将消息写入 conversation 表并关联当前 session_id

#### Scenario: 新建会话

- **WHEN** 用户点击「新建会话」
- **THEN** 系统 MUST 生成新 session_id，清空 ChatStateService 当前会话缓冲区

---

### Requirement: 滚动窗口管理

系统 SHALL 对当前会话消息实施滚动窗口管理。当消息数超过 context_limit 时，SHALL 从最早的消息开始裁剪，保证不超限。被裁掉的消息不自动进入长期记忆。

#### Scenario: 消息数超限裁剪

- **WHEN** ChatStateService 中消息数超过 context_limit + 2（含 system）
- **THEN** 系统 MUST 移除最早的一轮 user/assistant 对话

---

### Requirement: 记忆可配置与透明

系统 SHALL 支持以下可配置项：memory.rag_top_n、memory.enable_agentic_search、memory.enable_user_profile、memory.atomic_content_max_length、memory.session_summary_limit。所有记忆 SHALL 支持完整 CRUD，用户可查看、编辑、删除。

#### Scenario: 配置项生效

- **WHEN** 用户修改 memory.rag_top_n
- **THEN** 下次被动 RAG 检索时 MUST 使用新值作为 topN

---

### Requirement: 可选主动 RAG 工具

系统 MAY 将记忆搜索暴露为 MCP 工具。若启用，模型 SHALL 可通过工具调用按 keyword 和可选 timeRange 检索记忆。

#### Scenario: 主动记忆搜索（可选）

- **WHEN** memory.enable_agentic_search 为真且模型发起记忆搜索工具调用
- **THEN** 系统 MUST 执行 MemoryService.SearchMemories(keyword, timeRange) 并返回结果
