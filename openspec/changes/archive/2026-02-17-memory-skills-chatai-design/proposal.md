# 提案：Memory / Skills / Chat AI 能力升级

## Why

Office AI Agent 的 Chat 功能当前仅使用内存 List 存储历史，无跨会话记忆、无 RAG 召回、无 Skills 加载，无法为用户提供「记得你」的连贯体验。ChatGPT、Claude 等产品的实践证明：**长期记忆并非模型原生能力**，而是通过分层上下文注入、RAG 检索、原子记忆等工程手段实现。本提案旨在将这套成熟的设计理念复现并融入 Office AI，提升 Chat AI 的智能体能力。

## What Changes

- **记忆系统**：引入 MemoryService，实现原子记忆、用户画像、近期会话摘要；支持被动 RAG（构建上下文时自动检索）和可选主动 RAG（MCP 工具）
- ** Skills 与提示词**：按场景（Excel/Word/PPT）加载系统提示词与 Skills，支持变量替换（如 `{{选中内容}}`）
- **上下文工程**：请求组装改为 roadmap 2.5 的七层结构（System → 场景指令 → 用户记忆 RAG → 会话摘要 → 当前会话 → 本条消息），替代「全量历史堆砌」
- **ChatStateService**：由纯内存 List 改为滚动窗口 + 持久化到 `conversation` 表
- **HttpStreamService**：请求体组装前接收「已分层组装的上下文消息」，而非仅 `HistoryMessages`

## Capabilities

### New Capabilities

- `memory`：短期/长期记忆、原子记忆、用户画像、RAG 检索、会话摘要、可配置项与 SQLite 存储
- `skills-config`：系统提示词与 Skills 的配置、导入、按场景加载、变量替换
- `chat-ai-context`：Chat AI 分层上下文组装、与 Memory/Skills 集成、新建会话与持久化

### Modified Capabilities

<!-- 无现有 spec 需修改 -->

## Impact

- **ShareRibbon**：新增 MemoryService、扩展 ChatStateService、修改 HttpStreamService 请求组装逻辑、与 ConfigManager 对接
- **SQLite**：新增或扩展 `atomic_memory`、`user_profile`、`session_summary`、`conversation` 等表
- **依赖**：若采用向量 RAG，需 Embedding 模型或 API；初版可用关键词+时间简化

---

## 设计背景与依据（引言整理）

> 以下内容整理自社区对 ChatGPT / Claude 长期记忆系统的逆向工程与产品分析，作为本提案的设计依据。参考：Mathan 等对 ChatGPT/Claude 记忆系统的分析、Cherry Studio / Open-Web-UI 等开源实践。

### 1. 从 API 到客户端的上下文管理

**术语约定**：session/thread 指一个对话窗口内所有消息的集合；prompt/chat 指单条消息或模型回复。

**核心结论**：发送给 LLM 的并非原始消息数组，而是经 **context engineering** 处理的结果。以 ChatGPT 为例，上下文结构大致为：

| 层级 | 内容 |
|------|------|
| [0] System Instructions | 系统级配置、防 injection、工具定义 |
| [1] Developer Instructions | 场景指令 |
| [2] Session Metadata | 当前时间、地区等临时元信息 |
| [3] User Memory | 长期事实，**RAG 检索结果**，非全量 |
| [4] Recent Conversations Summary | 近期会话标题 + 摘要片段 |
| [5] Current Session Messages | 当前会话滚动窗口 |
| [6] 用户最新消息 | 本次 user prompt |

**启示**：长期记忆是 prompt engineering 与数据管理的结合；每次对话时将相关历史检索并动态注入，而非全量发送。

### 2. ChatGPT 的长期记忆架构

**三层设计**：

- **User Memory（显性）**：用户事实的结构化存储；设置中可查看、编辑；实际注入时使用 RAG 检索，非全量。
- **Recent Conversations Summary**：近期会话摘要；可能采用两阶段检索：先通过 bio 时间戳匹配时间区间，再在该区间内 RAG 检索 top-k 摘要；同时固定加入最近若干 session 摘要。
- **User Insight（隐性）**：用户偏好、风格、背景的高层总结；多为隐性、不可编辑；用于指导回复风格。

**滚动窗口**：当前会话消息有 token 上限；超限时从最早消息裁剪。被裁掉的内容若未写入记忆系统则永久丢失。

**信息密度**：上下文管理的目标是在有限 token 内放入**信息密度最高**的内容，避免冗余与 lost-in-the-middle。

### 3. Claude 的记忆系统

**Agentic RAG**：Claude 将会话/记忆搜索作为**工具**暴露给模型，由模型自主决定何时检索。灵活性高，但对模型能力要求高；弱模型可能出现漏检或过度检索。

**取舍**：ChatGPT 方案更工程化、稳定；Claude 方案更前沿、灵活。本提案以被动 RAG 为主，可选开放主动记忆搜索工具。

### 4. 开源实践

- **Cherry Studio**：直接 RAG，对话结束自动存储记忆；去重、去矛盾处理。
- **Open-Web-UI**：Agentic RAG，由模型决定搜索/更新/添加记忆。
- **共同点**：记忆透明、支持 CRUD；缺乏 ChatGPT 式的多层架构（bio + 会话摘要 + user insight）。

### 5. 融合方案要点（复现与改进）

| 组件 | 设计要点 |
|------|----------|
| **原子记忆** | ID、Timestamp、Content（3～4 句）、Tags；异步写入；去重/冲突检测 |
| **用户画像** | 结构化文档；可见可编辑（区别于 ChatGPT 的隐性 user insight） |
| **文件夹记忆** | 可选；按项目/主题隔离上下文 |
| **被动 RAG** | 每次请求前自动检索 top-N 原子记忆 + 加载用户画像 |
| **主动 RAG** | 可选；将记忆搜索暴露为 MCP 工具 |
| **可配置与透明** | 存储频率、top-N、主动检索开关、用户画像开关等；所有记忆支持 CRUD |

以上为提案的设计依据，与 `docs/01Memory.md` 中「Office AI 实现设计规格」及 `docs/roadmap.md` 2.5 节保持一致。
