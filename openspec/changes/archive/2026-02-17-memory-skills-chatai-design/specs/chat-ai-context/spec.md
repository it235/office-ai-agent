# Chat AI Context 能力规格

## ADDED Requirements

### Requirement: 分层上下文组装

系统 SHALL 按 [0]～[6] 顺序组装发往 LLM 的上下文消息，而非全量历史堆砌。[0] 系统级配置、[1] 场景指令与 Skills、[2] Session Metadata（可选）、[3] 用户记忆 RAG、[4] 近期会话摘要、[5] 当前会话滚动窗口、[6] 用户最新消息。

#### Scenario: 完整上下文组装

- **WHEN** 用户发送消息，系统构建请求
- **THEN** 系统 MUST 按 [0]～[6] 顺序组装消息列表，并将结果传给 HttpStreamService

#### Scenario: 无记忆时的降级

- **WHEN** MemoryService 不可用或未启用
- **THEN** 系统 MUST 跳过 [3][4] 层，仅注入 [0][1][5][6]（[2] 可选）

---

### Requirement: ContextBuilder 职责

系统 SHALL 提供 ContextBuilder（或 ChatContextBuilder）类，负责协调 MemoryService、ChatStateService、ConfigManager，按分层顺序产出消息列表。HttpStreamService SHALL 接收「已组装的 List(Of Message)」作为输入，而非直接使用 ChatStateService.HistoryMessages。

#### Scenario: HttpStreamService 接收预组装消息

- **WHEN** BaseChatControl 发起流式请求
- **THEN** 系统 MUST 先调用 ContextBuilder 产出消息列表，再将该列表传入 HttpStreamService.SendStreamRequestAsync

---

### Requirement: 与 Memory 集成

系统 SHALL 在构建上下文时调用 MemoryService.GetRelevantMemories 获取 [3] 层内容，调用 MemoryService.GetRecentSessionSummaries 获取 [4] 层内容，调用 ChatStateService 获取 [5] 层（当前会话滚动窗口）。

#### Scenario: 注入用户记忆与会话摘要

- **WHEN** 构建上下文且 Memory 已启用
- **THEN** [3] 层 MUST 包含 MemoryService.GetRelevantMemories 的 top-N 结果；[4] 层 MUST 包含 GetRecentSessionSummaries 的结果

---

### Requirement: 与 Skills 集成

系统 SHALL 在构建 [1] 层时，从 prompt_template 表按当前场景加载系统提示词与已启用 Skills，执行变量替换后注入。

#### Scenario: 注入 Skills

- **WHEN** 构建 [1] 层且当前宿主为 Excel
- **THEN** 系统 MUST 加载 scenario=excel 的系统提示词与 supported_apps 含 Excel 的 Skills，替换变量后合并注入

---

### Requirement: 新建会话与持久化

系统 SHALL 在新建会话时生成 session_id 并清空当前会话缓冲区。系统 SHALL 在每条 user/assistant 消息产生时，将其持久化到 conversation 表。流式响应完成后，系统 SHALL 异步调用 MemoryService.SaveAtomicMemoryAsync。

#### Scenario: 新建会话

- **WHEN** 用户点击「新建会话」
- **THEN** 系统 MUST 生成新 session_id，清空 ChatStateService，保留 MemoryService 与用户画像

#### Scenario: 响应完成后异步记忆写入

- **WHEN** 流式响应完成
- **THEN** 系统 MUST 启动后台任务调用 MemoryService.SaveAtomicMemoryAsync，不阻塞 UI

---

### Requirement: 收藏回答

系统 SHALL 支持用户「收藏回答」。收藏时 SHALL 将对应 conversation 记录的 is_collected 置为 1。系统 MAY 在收藏时触发原子记忆写入。

#### Scenario: 收藏单条回答

- **WHEN** 用户对某条 assistant 回复点击「收藏」
- **THEN** 系统 MUST 将对应 conversation 记录的 is_collected 更新为 1
