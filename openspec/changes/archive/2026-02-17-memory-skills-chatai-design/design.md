# 技术设计：Memory / Skills / Chat AI 能力升级

## Context

**现状**：Office AI Agent 的 Chat 由 ShareRibbon 的 `BaseChatControl` 承载，依赖 `ChatStateService`（内存 `List(Of HistoryMessage)`）、`HttpStreamService`（流式请求）、`HistoryService`（仅管理 `saved_chat_*.html` 导出）。上下文组装直接使用 `HistoryMessages`，无分层、无 RAG、无 Skills。

**约束**：VB.Net、.NET Framework 4.7.2、VSTO、ShareRibbon 供 Excel/Word/PPT 三端复用；UI 更新须在主线程；配置通过 ConfigManager；已有 MCP 集成（McpService）。

**参考**：提案、`docs/01Memory.md`、`docs/02SkillsConfig.md`、`docs/03ChatAi.md`、`docs/roadmap.md` 2.5 节。

## Goals / Non-Goals

**Goals:**

- 实现分层上下文注入（[0]～[6]），替代全量历史堆砌，降低 lost-in-the-middle
- 实现 MemoryService：原子记忆、用户画像、近期会话摘要、被动 RAG
- 实现 Skills 与系统提示词按场景加载及变量替换
- 会话持久化到 SQLite `conversation` 表
- 异步写入原子记忆，不阻塞主对话

**Non-Goals:**

- 本阶段不实现文件夹记忆（可后续扩展）
- 主动 RAG（MCP 工具）为可选，非首版必选
- 不引入新的 UI 框架（沿用 WinForm/现有 Pane）
- 不改变现有 MCP 工具定义，仅可选新增「记忆搜索」工具

## Decisions

### 1. MemoryService 独立于 ChatStateService

**决策**：新建 `MemoryService`，不将长期记忆逻辑塞入 `ChatStateService`。

**理由**：ChatStateService 职责为当前会话状态（滚动窗口、选区映射、响应映射）；长期记忆涉及 SQLite、RAG、异步写入，职责边界清晰。二者通过「构建上下文时 MemoryService 提供 [3][4] 层、ChatStateService 提供 [5] 层」协作。

**备选**：扩展 ChatStateService → 会导致单类职责混杂，不利于测试与扩展。

---

### 2. RAG 初版：关键词 + 时间范围，向量为后续

**决策**：首版 RAG 采用「关键词匹配 + 可选时间范围过滤」，不依赖 Embedding API。

**理由**：减少外部依赖与实现成本；SQLite `LIKE` 或 FTS5 即可支持；多数 Office 场景下用户问题含明确关键词。后续若效果不足，再引入 Embedding（如 OpenAI Embedding、本地 sentence-transformers）做向量检索。

**备选**：直接上向量 RAG → 需 NuGet 依赖、API Key 或本地模型，增加首版复杂度。

---

### 3. 上下文组装：新增 ContextBuilder 类

**决策**：新增 `ContextBuilder`（或 `ChatContextBuilder`），负责按 [0]～[6] 顺序组装消息列表；HttpStreamService 接收「已组装的 `List(Of Message)`」作为请求体输入。

**理由**：上下文逻辑集中在一处，便于测试与演进；HttpStreamService 保持「发送 + 流处理」职责；BaseChatControl 或调用方负责协调 MemoryService、ChatStateService、ContextBuilder。

**备选**：在 HttpStreamService 内直接组装 → 会膨胀该类，且难以单测上下文逻辑。

---

### 4. Skills / 提示词存储：统一 prompt_template 表 + extra_json

**决策**：沿用 roadmap 2.6 的 `prompt_template` 表，用 `is_skill` 区分系统提示词与 Skill；Skill 的 `parameters`、`supported_apps` 存于 `extra_json`。不单独建 `skill` 表。

**理由**：与 roadmap 一致，减少表数量；02SkillsConfig 已有该约定。查询时按 `scenario`、`is_skill` 过滤即可。

**备选**：拆分为 `system_prompt` + `skill` 两表 → 更清晰但增加迁移与对接成本。

---

### 5. 异步记忆写入：Task.Run +  fire-and-forget

**决策**：流式响应完成后，使用 `Task.Run` 启动后台任务调用 `MemoryService.SaveAtomicMemoryAsync`，不 await，不阻塞 UI。

**理由**：VB.Net 中实现简单；记忆写入非关键路径，失败可仅记录日志。若需重试可后续加队列。

**备选**：引入消息队列或后台服务 → 对桌面端过重。

---

### 6. 滚动窗口：先按消息对数，再考虑 token

**决策**：首版滚动窗口仍按「消息对数」裁剪（如 `context_limit`），与现有 `ManageHistorySize` 兼容；token 估算为后续优化。

**理由**：token 估算需引入 tiktoken 或类似库，增加依赖；消息对数在多数场景下已能控制上下文长度。

**备选**：首版即做 token 级裁剪 → 增加实现与测试成本。

---

### 7. 会话 ID 生成与持久化时机

**决策**：新建会话时生成 `session_id`（GUID）；每条 user/assistant 写入 `conversation` 时关联 `session_id`。持久化在「消息发送前（user）」和「流式响应完成后（assistant）」执行。

**理由**：保证会话维度可查；与 HistoryService 的 `saved_chat_*` 可并行存在（conversation 为结构化存储，saved_chat 为导出快照）。

---

## Risks / Trade-offs

| 风险 | 缓解 |
|------|------|
| 关键词 RAG 召回率不足 | 后续引入向量 RAG；提供「记忆管理」界面让用户手动补充 |
| 异步写入失败导致记忆丢失 | 日志记录；可选增加重试或本地队列 |
| 上下文组装增加首包延迟 | 被动 RAG 与 session summary 查询保持轻量；可加缓存 |
| Skills 变量替换与 LLM 返回格式冲突 | 明确占位符规范（`{{xxx}}`），避免与 JSON 等混淆 |
| 多宿主（Excel/Word/PPT）配置冲突 | 按 `scenario` 严格隔离；ConfigManager 按当前宿主加载 |

---

## Migration Plan

1. **数据库**：执行 SQLite 迁移脚本，新增 `atomic_memory`、`user_profile`、`session_summary`、`conversation` 表（若不存在）；`prompt_template` 增加 `extra_json` 等字段（若需）。
2. **兼容**：保留现有 `ChatStateService` 接口，新增持久化与 MemoryService 调用为增量逻辑；旧会话无 `session_id` 时可按「无摘要」处理。
3. **回滚**：若出现问题，可通过配置开关禁用 Memory/Skills 加载，回退到「仅 ChatStateService 历史」的旧行为。
4. **部署**：随 ShareRibbon 发布；Excel/Word/PPT 三端自动获得能力。

---

## 统一智能体流程（阶段四）

**目标**：将「意图识别 + 记忆/上下文 + Ralph 式规划与执行」整合为一条流水线，使发送时自动：收集内容区引用与记忆 → 识别真实意图（必要时询问用户）→ 得到 Spec 执行方案 → 按步骤执行。

**流程**：
1. **上下文收集**：发送前统一收集「内容区引用（选中/附件摘要，可为引用而非全文）+ 当前会话 + RAG 相关记忆」作为意图与规划的输入。
2. **意图阶段**：用现有 `IntentRecognitionService.IdentifyIntentAsync` + 上述上下文做意图识别；不清晰时通过意图预览卡片询问用户。
3. **规划阶段**：意图明确后，调用 LLM 产出 **Spec 执行方案**（JSON 步骤列表，与 Ralph 的 `steps` 格式对齐），解析后得到可执行步骤。
4. **执行阶段**：按 Spec 步骤依次执行（复用 `RalphLoopController` 的解析与执行循环），在 Chat 中展示步骤与结果；支持「继续下一步」或一次性执行。

**与现有组件的整合**：
- **IntentRecognitionService**：继续作为意图识别与置信度/澄清入口；上下文由 `GetContextSnapshot()` 增强（见下）。
- **RalphLoopController / RalphLoopSession**：规划结果统一解析为 `RalphLoopStep` 列表，复用 `ExecuteNextStep`、`CompleteCurrentStep` 与前端「继续执行」逻辑，避免两套执行引擎。
- **GetContextSnapshot**：在基类或子类中注入「RAG 摘要（如 top 2 条）+ 选中/引用摘要」，供意图与规划 LLM 使用。

**配置与模式**：可通过「智能体模式」开关或 Chat/Agent 模式选择是否在意图确认后自动请求 Spec 并进入步骤执行；默认可先做成「Agent 模式下意图确认后请求 Spec 并展示步骤，用户点击执行」。

---

## Open Questions

1. **原子记忆写入的 LLM 调用**：异步分析 `(userPrompt, assistantReply)` 是否复用主对话的模型配置，还是使用单独的「轻量模型」以降低成本？
2. **会话摘要生成**：摘要由主模型生成还是独立摘要模型？首版可采用「取首条 user 消息前 N 字符」作为 snippet 的简化方案。
3. **ConfigManager 扩展**：记忆相关配置项是落在现有 ConfigManager 还是新建 `MemoryConfig` 子模块？
4. **Ribbon「记忆设置」入口**：提案提及「记忆设置」；需确认是新建配置页还是集成到现有设置弹窗。
