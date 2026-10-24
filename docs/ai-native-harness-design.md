# Office AI Agent · AI Native Harness 详细设计方案

> **文档类型**：产品 + 架构设计（目标态基线）
> **状态**：目标设计基线；不等同当前代码已全部实现
> **落地设计**：[`ai-native-harness-implementation-design.md`](./ai-native-harness-implementation-design.md) 是后续编码的执行入口
> **目标读者**：产品、架构、Agent 平台、Word/Excel/PPT 执行层开发
> **相关文档**：
> - `openspec/changes/ai-native-productization-roadmap.md`（已落地阶段跟踪）
> - `openspec/changes/global-architecture-hardening-plan.md`（工程硬化）
> - `docs/project-issues-and-optimization-audit.md`（问题与优化审计）
> - `docs/build-and-installer.md` / `docs/signing-and-certificates.md`
> - 根目录 `AGENTS.md` / `CLAUDE.md`（AI Native 硬规则）

---

## 0. 一句话定位

把当前「Office 内嵌 Chat + 部分 Agent Loop」升级为：

> **在 Word / Excel / PowerPoint 本地文档上下文中，以 Harness 约束的 Agent 自主完成多步真实操作，并具备 Cursor 级 plan→act→observe→repair 闭环与 Copilot 365 / Claude for Office 级办公场景体验。**

不是「更聪明的聊天框」，而是 **Office 原生执行型 Agent 平台**。

---

## 1. 对标与差异化

### 1.1 对标矩阵（产品体验）

| 维度 | Microsoft 365 Copilot | Claude for Office（及同类） | Cursor 类 Coding Agent | **本产品目标（Office AI Agent）** |
|---|---|---|---|---|
| 宿主 | 云 + Office Web/桌面深度集成 | 侧栏/插件式，强对话 | IDE 内文件/终端/工具 | **VSTO 本地 COM：真实改文档** |
| 上下文 | Graph + 文档 + 组织数据 | 对话 + 上传/选区 | 代码库索引 + 打开文件 | **选区/结构/样式/表格/幻灯片 + Memory + Skills** |
| 执行 | 部分动作 + 草稿为主 | 偏生成与建议 | **强工具调用 + 多步自主** | **Harness 强制工具边界 + 多步自治** |
| 可撤销 | 版本/云协作 | 用户手动 | 本地 diff/undo | **Undo 点 + 预览 + 解释** |
| 隐私 | 企业租户 | 云 API | 本地/云可选 | **本地插件 + 用户自带 API Key** |
| 扩展 | Graph/插件生态 | MCP/工具有限 | Skills/Rules/MCP | **目录型 Skills + Tools JSON + MCP** |

### 1.2 我们不可让渡的差异化

1. **本地真实执行**：直接操作用户打开的 `.docx/.xlsx/.pptx`，不是只吐草稿。
2. **三端统一 Harness，宿主分离 Executor**：共享 Loop/Skill/Tool 契约，COM 永不进入 `ShareRibbon`。
3. **用户可控模型与密钥**：兼容 OpenAI/Anthropic 等，不锁死单一云。
4. **可解释、可撤销、可预览**：对标 Cursor 的 diff/explain，落到 Office Undo 与任务窗格。
5. **AI Native 硬约束**：入口不堆关键词 `if/else`；能读上下文就不问用户；明确目标默认执行/预览。

### 1.3 不对标（明确非目标）

| 非目标 | 原因 |
|---|---|
| 完整复刻 M365 企业 Graph 权限与合规中台 | 体量与商业模式不同，可后置「连接器」 |
| 替代 Word/Excel/PPT 全部 UI | 做 Agent，不是重做 Office |
| 无边界任意 VBA 自动运行 | 高风险，必须 Safety + 审批门 |
| 在 ShareRibbon 直接写 COM | 破坏宿主边界，长期不可维护 |

---

## 2. 当前已实现能力盘点（设计起点）

> 本章只描述当前代码事实。后文的 `IOfficeHarness`、`ContextPack`、`RunTrace`、`DocumentDiff`、`allowed-tools` 硬拒绝等是目标合同或部分实现项；具体落地顺序以 [`ai-native-harness-implementation-design.md`](./ai-native-harness-implementation-design.md) 为准。

### 2.1 已具备的平台骨架

```text
用户输入 / 引用 / 选区
        │
        ▼
ChatRoutingOrchestrator  ──► AiNativeRuntime.AnalyzeAsync
        │                         │ Intent + TaskSpec + Skills + Tools + Trace
        ▼                         ▼
   AgentKernel ──► LoopEngine (Think→Plan→Act→Observe→Repair/Reflect)
        │
        ├── ToolRegistry (72 JSON Tools + MCP + memory.* + skill_script.*)
        ├── SkillRegistry / SkillsDirectoryService (目录 SKILL.md + 旧 JSON)
        ├── AgentMemory + UnifiedMemory / memory_item 管线
        └── Host: ExecuteJsonCommand / WordActionHarness / 翻译·续写·排版
```

| 模块 | 代表实现 | 成熟度 |
|---|---|---|
| AI Native 分析中枢 | `AiNativeRuntime` | ★★★☆ 有统一入口，仍依赖 Intent 大服务 |
| 智能路由 | `ChatRoutingOrchestrator` + `ExecutionPathPolicy` | ★★★☆ 主路径收敛，Word 仍有 harness 快路径 |
| Agent Kernel | `AgentKernel` / `AgentSession` / `AgentMemory` | ★★★☆ 可跑通；VSTO shadow-copy 下 `Tools/Prompts/Skills` 定位已修复，上下文空兜底仍需收紧 |
| ReAct Loop | `LoopEngine` | ★★★☆ 有 repair/replan；0 工具迭代不再误报成功，观察结构化仍需扩展 |
| Self-check Loop | `SelfCheckLoopController` + DSL 校验 | ★★☆☆ 与主 Loop 并存，场景偏排版 |
| Tools | `ShareRibbon/Tools/**` ~72 | ★★★☆ Schema 全，执行质量与 observe 不均 |
| Skills | 4 目录 Skill + 5 旧 JSON | ★★★☆ 双格式；工具存在性/宿主可用性已硬校验，Skill `allowed-tools` 已按命中 Skill 收紧，门禁 Trace 待补 |
| Memory | atomic + memory_item + embedding + 晋升 | ★★★☆ 产品化有基础，双模型待收敛 |
| MCP | StreamJsonRpc client + 配置 UI | ★★★☆ 可用，健康度/超时需统一 |
| Word Capability | `WordActionHarness` / `WordCapabilityRegistry` / 排版校对编号 | ★★★★ Word 最强样板；四类 fast-path 已有结构化结果或启动结果回灌 |
| Excel 执行 | `ExcelDirectOperationService` + ExcelDna | ★★★☆ 命令强，Harness 弱 |
| PPT 执行 | ChatControl 命令分发 | ★★☆☆ 工具多，闭环弱 |
| 上下文 UI | 上下文控制台 / Agent 卡片 | ★★★☆ 可解释雏形 |
| 工程 | smoke、code/installer 二分、AppLogger、ToolResult 错误契约、P0 guardrails、code-only CI | ★★★☆ P0 治理后可继续产品化 |
| 网关 | `AiGateway` 非流式 + `HttpStreamService` 流式 | ★★☆☆ 双轨 |

### 2.2 与「完全 Harness」的差距（总览）

| 能力 | 现状 | 完全 Harness 目标 |
|---|---|---|
| 入口 | Chat/Ribbon 仍可能旁路 | **所有自然语言只进 Harness** |
| 规划 | LLM JSON 计划 + 部分规则 | **结构化 Plan + 工具约束 + 风险分级** |
| 执行 | JSON 命令 / VBA / 快路径混用 | **仅 Tool/Capability Executor** |
| 观察 | 部分 success 字符串 | **文档差分 + 结构化 ToolResult** |
| 修复 | 有限次 re-call | **基于观察的参数修复 + 换工具 + 降级** |
| 技能 | 召回弱、边界软 | **两阶段加载 + allowed-tools 硬门禁** |
| 上下文 | 有但不全 | **结构化文档图 + 预算管理** |
| 安全 | 零散 | **统一 SafetyGate（预检/审批/沙箱）** |
| 评测 | smoke 为主 | **场景黄金集 + 工具契约测试** |

---

## 3. 目标架构：Office Harness 全景

### 3.1 分层（强制）

```text
┌─────────────────────────────────────────────────────────────┐
│  Surface（表现层）                                            │
│  Chat WebView / Ribbon / 快捷键 / 选区右键 / 上下文控制台      │
│  只负责：采集输入、展示计划/进度/解释/审批、触发 Harness       │
└────────────────────────────┬────────────────────────────────┘
                             │ UserTurn + HostHints
                             ▼
┌─────────────────────────────────────────────────────────────┐
│  OfficeHarness（编排层 · 唯一产品入口）                        │
│  ┌──────────────┐  ┌─────────────┐  ┌────────────────────┐  │
│  │ ContextHub   │→ │ Planner     │→ │ RunController      │  │
│  │ 读文档/记忆   │  │ Spec+Plan   │  │ Loop 状态机        │  │
│  └──────────────┘  └─────────────┘  └─────────┬──────────┘  │
│         ▲                    │                 │             │
│         │                    ▼                 ▼             │
│  ┌──────────────┐  ┌─────────────┐  ┌────────────────────┐  │
│  │ SkillRouter  │  │ ToolBroker  │  │ SafetyGate         │  │
│  │ 两阶段 Skill │  │ 工具边界    │  │ 风险/审批/限额     │  │
│  └──────────────┘  └─────────────┘  └────────────────────┘  │
│         │                    │                 │             │
│         └──────────── Observe / Repair / Explain ───────────┘│
└────────────────────────────┬────────────────────────────────┘
                             │ ToolCall (schema-valid)
                             ▼
┌─────────────────────────────────────────────────────────────┐
│  Capability / Executor（宿主执行层）                           │
│  WordExecutor │ ExcelExecutor │ PptExecutor │ McpExecutor     │
│  只负责：COM 操作、返回 ToolResult、登记 Undo                 │
└────────────────────────────┬────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────┐
│  Infrastructure                                               │
│  AiGateway(stream/non-stream) · MemoryStore · SQLite          │
│  AppLogger · UiDispatcher · Config · Prompt Profiles          │
└─────────────────────────────────────────────────────────────┘
```

### 3.2 与现有类型的映射（迁移不推倒）

| 目标组件 | 现有落点 | 演进动作 |
|---|---|---|
| OfficeHarness | `ChatRoutingOrchestrator` + `AgentKernel` | 合并为唯一 façade：`IOfficeHarness.RunAsync(turn)` |
| ContextHub | `CaptureOfficeContext` + `ChatContextBuilder` + Memory | 统一 `IContextHub.SnapshotAsync`，输出预算化 ContextPack |
| Planner | `AiNativeRuntime` + `IntentRecognitionService` + Loop 规划 | Intent 只产 Spec；Plan 只在 Harness 内生成 |
| RunController | `LoopEngine` + `SelfCheckLoopController` | **合并主状态机**，SelfCheck 变 Pre/Post 插件 |
| SkillRouter | `SkillsDirectoryService` + `SkillRegistry` | 废弃 JSON skill 主路径；强制 front matter 两阶段 |
| ToolBroker | `ToolRegistry` | allowed-tools ∩ appType ∩ Safety 过滤后才可见 |
| SafetyGate | `SafetyChecker` + 预检 | 扩展为统一门：风险级、审批、配额、COM 线程 |
| Host Executor | `ExecuteJsonCommand` / `WordActionHarness` | 全部登记为 Capability，禁止 NL 旁路 |
| Observe | `ToolResult` / FormatObservation | 强制文档差分 + 错误契约 |
| Explain | Agent 卡片 / ExecutionExplanation | 统一用户可见时间线 |

### 3.3 唯一主路径（目标硬规则）

当前代码的主路径是 `ChatRoutingOrchestrator` → `AgentKernelService` → `OfficeHarness.RunAsync` → `AgentKernel` → `LoopEngine`。`IOfficeHarness` 第一版 adapter 已落地并进入主路径；后续继续补完整状态机、控制 API 和 ContextHub，不做推倒式重构。

```text
UserTurn
  → Surface 采集（选区、附件、模式）
  → OfficeHarness.RunAsync
       1. ContextHub.Snapshot
       2. SkillRouter.Recall（元数据）
       3. Planner.BuildSpec + Plan
       4. SafetyGate.PreCheck
       5. 展示 Plan（可自动执行低风险；高风险等 Approve）
       6. Loop:
            Think → ToolBroker.Select → SafetyGate.Act
            → HostExecutor.Execute → ToolResult
            → Observe(文档差分) → Repair|Continue|Stop
       7. Explain + MemoryWrite + Trace 落库
  → Surface 渲染结果 / 撤销入口
```

**禁止**（新代码零容忍；老路径按阶段迁移）：

- 在 `ChatControl` 用 `Select Case` 关键词决定业务主路径
- 自然语言直接 `ExecuteJsonCommand` 而不经 ToolBroker
- Skill 未声明工具却被调用
- 空 `OfficeContext` 进入生产 Loop（仅测试夹具允许）

---

## 4. 核心领域模型（需细化沉淀的契约）

> 下列模型建议落在 `ShareRibbon/Agent/Contracts/`（或 `Protocol/`），JSON 可序列化，供 UI / 日志 / 评测共用。

### 4.1 UserTurn

| 字段 | 类型 | 说明 |
|---|---|---|
| TurnId | string | 相关 id，贯通日志 |
| AppType | Excel\|Word\|PowerPoint | 宿主 |
| Text | string | 用户自然语言 |
| References | FileRef[] / SelectionRef[] | 附件与内容引用 |
| Mode | smart\|agent\|proofread\|template\|… | Surface 模式，不决定最终工具 |
| SessionId | string | 会话 |
| UIHints | object | 校对面板状态等非阻塞提示 |

### 4.2 ContextPack（预算化上下文）

| 段 | 内容 | 预算策略 |
|---|---|---|
| HostMeta | 应用、文档名、路径、只读、保护状态 | 必选短 |
| Selection | 选区地址、文本摘要、类型 | 优先 |
| Structure | 标题大纲 / 表头 / 幻灯片目录 | 分层截断 |
| Samples | 样例行/段落 | Top-N |
| Styles | 样式集、主题 | 按需 |
| Memory | 长期记忆 + 会话摘要 | RagTopN |
| Skills | 命中 Skill 元数据 | 仅 front matter |
| Tools | 可见工具摘要 | id+desc 短列表 |
| History | 近 N 轮 | 滑动窗口 |
| Risks | 保护表、宏、大表 | 安全提示 |

**细化项**：为 Word/Excel/PPT 分别定义 Structure Schema；统一 token 预算分配器 `ContextBudget`。

### 4.3 TaskSpec（意图升级产物）

| 字段 | 说明 |
|---|---|
| Goal | 一句话目标 |
| Objects | 作用对象（选区/全文/指定表/页） |
| Constraints | 不可改内容、保留编号等 |
| SuccessCriteria | 可观察成功标准 |
| RiskLevel | safe / medium / risky |
| RequiredCapabilities | 能力标签 |
| Complexity | simple / medium / complex / exploratory |
| Confidence | 0–1 |
| ClarifyingQuestions | **仅阻塞性问题**，默认可空 |

### 4.4 ExecutionPlan

| 字段 | 说明 |
|---|---|
| PlanId | |
| Steps[] | StepId, Description, ToolId?, CapabilityId?, DependsOn[], Risk, RequiresApproval |
| Summary | 用户可读 |
| AutoRunEligible | 是否可跳过人工点「开始」 |

### 4.5 ToolResult（已有，强制成为唯一执行回执）

已扩展字段需成为全链路强制：

- `Success`, `Message`, `Data`
- `ErrorCode`, `UserMessage`, `DebugDetail`, `Recoverable`
- `ElapsedMs`, `ToolId`
- `Observation`（文档差分摘要）、`UndoPointId`、`Artifacts[]`：字段已加入；Word 基础写工具已填 before/after/diff Observation，并通过 `ExecutionExplanation` 进入 Agent 卡片和 RunTrace。完整 Undo 绑定、Excel/PPT Diff 待后续。

### 4.6 RunTrace（轻量落库已实现，完整回放待补）

整轮：`UserTurn → ContextPack 摘要 → Spec → Plan → Steps[] → ToolResults[] → FinalExplain → MemoryWrites`
落库：当前已有 `agent_run` / `agent_run_step` 轻量表，由 `OfficeHarness` 通过 `SqliteRunTraceStore` 写入。`agent_run_event`、审批事件、回放 UI、崩溃 orphan 清理仍按落地设计后续补齐。

---

## 5. Harness 子系统详细设计

### 5.1 ContextHub ——「先读后想」

#### 目标

对标 Cursor 的 codebase awareness：Agent 永远先看到「当前 Office 世界」，而不是让用户复述。

#### 已有

- `OfficeContext` + 三端 `CaptureOfficeContext`
- `ChatContextBuilder` 记忆注入
- 上下文控制台 UI

#### 必须细化

| ID | 细化项 | 说明 | 优先级 |
|---|---|---|---|
| CTX-1 | **文档结构图** | Word：标题树/节/域；Excel：表区域/列类型/公式依赖摘要；PPT：幻灯片大纲/版式/母版 | P0 |
| CTX-2 | **选区语义** | 不仅是文本：单元格类型、段落样式名、形状类型 | P0 |
| CTX-3 | **大文档策略** | 采样 + 大纲 + 按需 `read_range` 工具二次拉取 | P0 |
| CTX-4 | **ContextBudget** | 按模型上下文窗口分配各段 token | P1 |
| CTX-5 | **脏读/快照** | 执行前后各采一次快照供 Observe 差分 | P0 |
| CTX-6 | **只读/保护/共享审阅** | 进入 Plan 前写入 Risk | P1 |
| CTX-7 | **跨打开工作簿/文档** | 明确 scope：仅活动文档 vs 允许切换 | P2 |

#### 接口草图

```text
IContextHub.SnapshotAsync(app, options) → ContextPack
IContextHub.DiffAsync(before, after) → DocumentDiff
IHostContextReader（Word/Excel/PPT 实现）
```

---

### 5.2 Planner ——「任务理解器，不是分类器」

#### 目标

对标 Cursor Agent 的任务拆解 + Copilot 的场景理解：输出 **可执行 Spec/Plan**，不是意图枚举标签。

#### 已有

- `IntentRecognitionService`（大）
- `AiNativeRuntime` 产 TaskSpec
- Loop 内 GeneratePlan

#### 必须细化

| ID | 细化项 | 说明 | 优先级 |
|---|---|---|---|
| PLN-1 | **双阶段规划** | Stage-A 廉价 Spec；Stage-B 带 Skill 正文与工具 schema 的 Plan | P0 |
| PLN-2 | **工具约束规划** | Plan 中每步只能引用 ToolBroker 可见工具 | P0 |
| PLN-3 | **探索式计划** | 低置信：先 List/Get/Read 再写 | P0 |
| PLN-4 | **拆分 Intent 巨类** | Classifier / SpecBuilder / PlanPrompt 分离 | P1 |
| PLN-5 | **Plan 校验器** | schema、环依赖、未知 tool、风险升级 | P0 |
| PLN-6 | **用户可编辑 Plan** | UI 改步骤后重跑（现 refine 弱） | P1 |
| PLN-7 | **多目标冲突** | 「校对并排版」拆并行/串行策略 | P1 |

#### 规划提示原则（写入 Prompt Profile）

1. 先观察后修改。
2. 优先原生 Tool，VBA 最后。
3. 写操作说明预期文档变化。
4. 不确定范围时用选区 → UsedRange → 询问（仅阻塞时）。
5. 成功标准必须可被 Observe 验证。

---

### 5.3 SkillRouter —— 能力发现与加载

#### 目标

对标 Cursor Rules/Skills、Claude Skills：Skill 是 **操作手册 + 工具边界**，不是关键词表。

#### 已有

- 目录型：`excel-table-agent` / `word-document-agent` / `powerpoint-deck-agent` / `office-skill-authoring`
- 旧 JSON skills
- Skills 索引与 usage registry

#### 必须细化

| ID | 细化项 | 说明 | 优先级 |
|---|---|---|---|
| SKL-1 | **废弃 JSON 主路径** | 迁移为目录 Skill 或只读兼容 | P0 |
| SKL-2 | **两阶段加载** | 元数据召回 → 命中后加载正文/references | P0 |
| SKL-3 | **allowed-tools 硬门禁** | ToolBroker 执行前校验 | P0 |
| SKL-4 | **application 隔离** | Excel Skill 不可见 Word tools | P0 |
| SKL-5 | **召回算法** | embedding + 意图标签 + 最近成功加权 | P1 |
| SKL-6 | **Skill 评测字段** | 成功率、失败码、样例 query | P1 |
| SKL-7 | **用户/企业 Skill 目录** | 文档目录外可配置路径 | P2 |
| SKL-8 | **Skill 作者体验** | 沿用 office-skill-authoring，补校验 CLI/smoke | P1 |

#### 目录规范（已有，固化为平台标准）

```text
ShareRibbon/Skills/<skill-name>/
  SKILL.md          # front matter + 操作手册
  references/       # 按需
  scripts/          # 可选
  assets/           # 可选
```

Front matter 必备：`name`, `description`, `application`, `allowed-tools`, `intent_types`, `risk_default`。

---

### 5.4 ToolBroker + Host Capability

#### 目标

对标 Cursor tools / MCP：模型只看见 **允许的、健康的、与宿主匹配的** 工具；执行永远经 Broker。

#### 已有

- 72 个 Tools JSON（excel/word/ppt/common）
- `ToolRegistry.ExecuteToolAsync`
- MCP 工具前缀 `mcp.*`
- memory.*
- WordCapabilityRegistry / ActionHarness

#### 必须细化

| ID | 细化项 | 说明 | 优先级 |
|---|---|---|---|
| TOL-1 | **统一 Capability 描述** | Tool JSON ∪ Host Capability 同一 Descriptor | P0 |
| TOL-2 | **执行器注册表** | `IToolExecutor` 按 toolId 路由到 Word/Excel/PPT | P0 |
| TOL-3 | **读工具回传结构化 Data** | 关闭 Word ListParagraphs 等 TODO | P0 |
| TOL-4 | **写工具强制 UndoPoint** | 每写操作前后挂钩 | P0 |
| TOL-5 | **幂等与重试语义** | 标注 idempotent / partial | P1 |
| TOL-6 | **工具健康** | unavailable/error 从 Plan 中剔除并解释 | P1 |
| TOL-7 | **VBA 降级策略** | 仅当无原生 tool 且用户允许 | P1 |
| TOL-8 | **Excel 大表批处理** | 分块 Write/Format，防 UI 卡死 | P0 |
| TOL-9 | **PPT 布局语义工具** | 不仅 InsertShape，要「对齐/分栏/母版」级 | P1 |
| TOL-10 | **未知命令** | 禁止「暂不支持」即停；进入 repair 换 tool | P0 |

#### Capability 生命周期

```text
Discover → Filter(app, skill, safety) → Model sees schema
  → Call → Validate args → UI-thread COM → ToolResult
  → Observe Diff → (optional) Verify hooks
```

#### Word 样板推广

`WordActionHarness` 当前是 **高置信快路径**。目标形态：

1. Harness 仍统一入口；
2. Planner 可直接选择 `word.proofread` / `word.numbering` Capability；
3. 快路径结果 **必须** 写成 ToolResult 并进入 Trace/Memory；
4. Excel/PPT 实现 `ExcelActionHarness` / `PptActionHarness` 同构。

---

### 5.5 RunController（Loop 状态机）—— Cursor 级自治核心

#### 目标

完整状态机，而不是「调用一次 LLM 后执行 JSON」。

#### 推荐状态

```text
Idle → Analyzing → Planning → AwaitingApproval?
  → RunningStep → Observing → Repairing | Replanning
  → Completed | Failed | Cancelled | NeedsInput
```

#### 已有

- `LoopEngine`：MaxIterations=15，修复 3 次，重规划 2 次
- 风险工具日志
- 结构化 ToolResult 摘要开始接入

#### 必须细化

| ID | 细化项 | 说明 | 优先级 |
|---|---|---|---|
| LOP-1 | **单一 Loop** | 吸收 SelfCheck 为 PreSend/PostFlush 钩子 | P0 |
| LOP-2 | **观察驱动修复** | repair prompt 必须含 DocumentDiff + ErrorCode | P0 |
| LOP-3 | **无进展检测** | 重复 tool+args 哈希熔断 | P0 |
| LOP-4 | **并行安全步** | 只读工具可并行；写串行 | P2 |
| LOP-5 | **中断/继续** | 用户 Cancel、Approve、补充信息后 resume | P1 |
| LOP-6 | **步骤级超时** | COM 卡死保护 | P1 |
| LOP-7 | **成功标准检查** | Plan.SuccessCriteria 自动 verifier | P1 |
| LOP-8 | **费用/轮次预算** | 按用户配置截断并解释 | P1 |

#### Observe 最小标准（P0）

每次写工具后至少返回：

- 变更范围（段落号 / 单元格地址 / 幻灯片索引）
- 前后指纹或抽样文本
- 错误码（若失败）
- 是否可 Undo

---

### 5.6 SafetyGate —— 办公场景安全

#### 目标

对标企业 Copilot 的「可控执行」+ Cursor 的「敏感命令确认」，但贴合 Office：删表、宏、保护工作表、批量替换。

#### 风险分级

| 级别 | 示例 | 默认策略 |
|---|---|---|
| safe | 读段落、列统计、生成预览 | 自动 |
| medium | 改格式、写公式到选区 | 自动或轻提示 |
| risky | 删行列、整表替换、VBA、取消保护、跨文档 | 需 Approve 或二次确认 |
| forbidden | 未授权网络、任意文件删除 | 拒绝 |

#### 必须细化

| ID | 细化项 | 优先级 |
|---|---|---|
| SAF-1 | 工具 schema 强制 risk 字段 | P0 |
| SAF-2 | 批量影响行数/页数阈值 | P0 |
| SAF-3 | 宏/VBA 独立开关 | P0 |
| SAF-4 | 保护工作表/文档加密交互 | P1 |
| SAF-5 | 审计日志（谁在何时改了什么摘要） | P1 |
| SAF-6 | 租户策略预留（日后） | P2 |

---

### 5.7 Memory —— 服务 Agent，而不是装饰 RAG

#### 已有

- 长短期、晋升、过期、冲突边
- Agent 主动 `memory.search`
- embedding

#### 必须细化

| ID | 细化项 | 优先级 |
|---|---|---|
| MEM-1 | 写入只走 memory_item 管线 | P0 |
| MEM-2 | atomic 只读淘汰期 | P1 |
| MEM-3 | 任务级记忆（本次 Plan 学到的文档约定） | P1 |
| MEM-4 | 用户偏好（语气、默认字号、公司术语）结构化 | P1 |
| MEM-5 | 记忆注入与 Skill 的优先级策略 | P1 |
| MEM-6 | 隐私：清除/导出/禁用开关产品化 | P0 |

---

### 5.8 Model Gateway —— 统一大脑接口

#### 已有

- `AiGateway` 非流式
- `HttpStreamService` 流式主聊天
- Provider 转换 smoke

#### 必须细化

| ID | 细化项 | 优先级 |
|---|---|---|
| GW-1 | Streaming 与 non-stream 统一适配层 | P0 |
| GW-2 | Tool-calling 原生协议（若模型支持） vs JSON 计划双模 | P1 |
| GW-3 | 超时/重试/限流/熔断 | P1 |
| GW-4 | 全链路脱敏日志 | P0 |
| GW-5 | 模型能力画像（是否支持 vision/tools/长上下文） | P2 |

---

### 5.9 Surface UX —— 对标 Copilot 侧栏体验

#### 已有

- WebView2 Chat、上下文控制台、Agent 卡片、校对/排版 UI

#### 必须细化

| ID | 细化项 | 对标 | 优先级 |
|---|---|---|---|
| UX-1 | **执行时间线** | Cursor 步骤流 | P0 |
| UX-2 | **Plan 卡片可编辑** | Agent 计划确认 | P1 |
| UX-3 | **文档内预览/高亮变更** | Copilot 建议批注感 | P0 |
| UX-4 | **一键撤销本轮 Agent** | Undo 栈 | P0 |
| UX-5 | **NeedsInput 最小提问卡** | 阻塞澄清 | P1 |
| UX-6 | **失败可理解文案** | 基于 UserMessage/ErrorCode | P0 |
| UX-7 | **多会话 / 任务历史** | 可回放 RunTrace | P2 |
| UX-8 | 收敛 Deepseek/Doubao 双入口或明确场景 | 降低分叉 | P2 |

**目标原则**：Surface **永不**实现业务路由；只调 `IOfficeHarness`。当前过渡期允许 `BaseChatControl` 调 `ChatRoutingOrchestrator`，但不得继续新增业务分支。

---

## 6. 三端场景能力地图（产品细化清单）

### 6.1 Word（最接近样板，继续做深）

| 场景 | 现状 | 要细化到 |
|---|---|---|
| 校对 | 有 IntentCompiler / UI | 问题列表 → 批量应用 → 验证 → 解释 |
| 直接格式 | FormattingIntentCompiler | 单位/样式冲突策略、范围推断 |
| 编号连续 | NumberingAgent | 多级列表、跨节 |
| 语义排版 | SmartFormatting / Reformat | 模板绑定、母版样式、失败 repair |
| 长文生成/续写 | 有服务 | 纳入 Harness，分段写入+观察 |
| 翻译 | 有 | 批处理 + 术语表 Memory |
| 结构编辑 | 工具有、回传弱 | 读工具 Data 完整、目录/题注 |

**Word 细化里程碑**：所有场景 `Capability → ToolResult → Diff → Explain` 全绿。

### 6.2 Excel（对标 Copilot 表格分析 + Cursor 工具精度）

| 场景 | 现状 | 要细化到 |
|---|---|---|
| 表理解 | 上下文服务有 | 自动探测表区域、类型、主键 |
| 公式生成/修复 | ApplyFormula | 相对引用、填充、错误检查观察 |
| 清洗 | CleanData 等 | 流水线 Skill + 可撤销 |
| 透视/图表 | 有 tool | 选型启发式写入 Skill |
| 多步分析报告 | 弱 | Skill 多步 + 新 sheet 输出 |
| 大表 | 风险 | 分块、后台、进度 |
| PowerQuery | 占位 | 明确做/不做；做则 Capability 化 |
| ExcelDna UDF | 同步 HTTP | 与 Agent 路径隔离策略 |

**Excel 细化里程碑**：`ExcelActionHarness` + 表级 Observe（值/公式/错误单元格）。

### 6.3 PowerPoint（对标「一句话生成演示」体验）

| 场景 | 现状 | 要细化到 |
|---|---|---|
| 生成大纲/页 | 有生成 handler | 版式选择、母版、占位符语义 |
| 美化 | 工具散 | 对齐/间距/主题一键 Capability |
| 图表/图 | Insert* | 数据绑定策略 |
| 演讲者备注 | 有 tool | 与讲稿 Memory 联动 |
| 未知命令 | warning | repair 换工具 |

**PPT 细化里程碑**：`PptActionHarness` + 幻灯片级 Diff（页数、标题、关键形状文本）。

---

## 7. 与 Copilot 365 / Claude for Office 的功能对标清单

### 7.1 必须对齐的用户故事（P0/P1）

| # | 用户故事 | 优先级 | 依赖子系统 |
|---|---|---|---|
| U1 | 选中内容不说话，AI 自动给出合理处理并执行/预览 | P0 | Context+Planner |
| U2 | 「把这份报告改成公司模板样式」多步完成 | P0 | Skill+Loop+Word |
| U3 | 「根据这张表做汇总和图表」 | P0 | Excel Skill+Tools |
| U4 | 「把纪要做成 8 页汇报 PPT」 | P0 | PPT Skill+Tools |
| U5 | 执行中可看每步在干什么 | P0 | UX 时间线 |
| U6 | 做错一键撤销本轮 | P0 | Undo |
| U7 | 记住我偏好的字体/术语 | P1 | Memory |
| U8 | 连接内部 MCP 工具查数据再写入表 | P1 | MCP+Broker |
| U9 | 低置信时先读后写，不瞎改 | P0 | 探索式 Plan |
| U10 | 高风险操作要我点头 | P0 | SafetyGate |

### 7.2 体验原则（写入产品宪法）

1. **默认行动**：可逆则先预览或直接做；不可逆则确认。
2. **最小提问**：只问阻塞执行的问题。
3. **可见推理**：展示依据（上下文/Skill/工具），不展示无用思维链刷屏。
4. **失败可恢复**：repair → 换工具 → 降级说明，而不是「暂不支持」。
5. **宿主诚实**：做不到的 COM 能力明确边界，给替代路径。

---

## 8. 与 Cursor / 主流 Agent 的能力对标（执行层）

| Cursor/Agent 能力 | Office 映射 | 我们缺口 |
|---|---|---|
| Repo 索引 | 文档结构索引 | CTX-1/3 |
| @file / @selection | 引用/选区 | 基本有，语义不足 |
| Tool calling | Tools JSON + MCP | 门禁与 observe 不足 |
| Apply patch + diff | 文档差分 + 高亮 | UX-3、CTX-5 |
| Terminal | 一般不需要；VBA≈危险终端 | SAF-3 |
| Rules / Skills | SKILL.md | SKL-1..4 |
| Multi-step agent | LoopEngine | LOP-1..3 |
| Stop / Continue | Abort/Approve | LOP-5 弱 |
| Test/lint loop | Observe/Verifier | LOP-7、场景 verifier |
| Composer 多文件 | 跨 sheet/slide/section | 部分有 |

**结论**：产品形态对标 Copilot 侧栏；**执行内核必须对标 Cursor Agent**，否则永远是「会聊天的插件」。

---

## 9. 工程与质量设计（支撑 Harness 可信）

### 9.1 测试金字塔

| 层 | 内容 | 现状 | 目标 |
|---|---|---|---|
| 单元 | Spec/Plan 校验、Skill front matter、ExceptionClassifier、Budget | 少 | `OfficeAi.Tests` |
| 契约 | Tool JSON schema、AiGateway provider | smoke 有 | 扩展每个 tool 最小 schema 测 |
| 场景黄金集 | U1–U10 固定文档夹具 | 无 | `fixtures/word|excel|ppt` + 脚本 |
| UI 手工 | 宿主冒烟 | 有清单 | 每版本必跑 |
| 安全 | 风险工具拒绝/审批 | 弱 | Safety 单测 |

### 9.2 可观测性

- `AppLogger` + correlation = TurnId
- RunTrace 落库
- 指标：步数、repair 次数、工具失败码 TopN、Skill 命中率、用户撤销率

### 9.3 性能预算（建议）

| 项 | 预算 |
|---|---|
| Context 快照（常文档） | < 300ms |
| 首 token / 首 Plan 展示 | < 3s（网络外） |
| 单步 COM 写 | < 2s 或出进度 |
| 整任务默认上限 | 15 步 / 可配置 |

### 9.4 线程模型（硬约束）

- UI / WebView2 / COM：UI 线程
- LLM / SQLite / embedding：线程池
- 禁止 UI 上裸 `.Result`（已用 SyncOverAsync 过渡）
- Executor 内部统一 `UiDispatcher.InvokeSync/Async`

---

## 10. 分阶段落地路线（设计 → 实现）

> 与历史 roadmap 衔接：不重复已勾选项，只列 **下一阶段必须细化并实现** 的包。

### Phase H0 —— Harness 契约冻结（1–2 周）

**目标**：所有新代码只依赖稳定契约。

**当前状态**：H0 adapter 已落地，`AgentKernelService.StartAgentAsync` 已经通过 `OfficeHarness` 调用现有 `AgentKernel`；Agent 原生工具主路径已强制 `ToolResult` 回执，不再回退 Boolean 假成功。P0 治理已落地工具定位、0 迭代失败防误报、code-only CI 与 P0 guardrails；轻量 RunTrace 和 Word 基础写工具最小 Diff 已完成，完整 ContextHub/审批控制/事件回放/Excel/PPT Diff 仍待后续。

- 冻结：`UserTurn` / `ContextPack` / `TaskSpec` / `ExecutionPlan` / `ToolResult` / `RunTrace`
- `IOfficeHarness` 接口 + 空壳实现转调现有 AgentKernel
- Surface 只调 Harness
- 文档：本设计 + 契约示例 JSON

**验收**：三端发送路径编译通过；旧路径标 obsolete。

### Phase H1 —— 主路径硬化（2–3 周）

- ContextHub 三端结构快照 + Diff
- ToolBroker allowed-tools 硬门禁
- Loop 合并 SelfCheck 钩子；Observe 强制差分字段
- Word 全部 Capability 回传 ToolResult
- 取消 NL 旁路（除 Safety 允许的快路径且写 Trace）

**当前进展**：工具存在性/宿主可用性已在 `ToolRegistry.TryNormalizeToolCall` 执行前校验；Word Numbering/DirectFormatting/Proofread/SemanticReformat fast-path 已有结构化结果或启动结果回灌；`ToolExecutionContext` 与执行期 Skill `allowed-tools` 硬拒绝已落地；Word `ListParagraphs/GetParagraphInfo` 已回 `Data`；Word `InsertText/FormatText/ReplaceText/DeleteText` 已有 before/after/diff Observation；最小 `SafetyGate` 已在 COM 前阻断 VBA/风险工具；轻量 RunTrace 已写 `agent_run`/`agent_run_step`。门禁事件、完整回放、Excel/PPT Observation 仍待实现。

**验收**：U1/U2/U9 黄金场景；失败可见 ErrorCode。

### Phase H2 —— Excel Harness 对齐（2–3 周）

- `ExcelActionHarness` + 表探测
- 公式/图表/清洗 Skill 实战
- 大表分块
- 场景 U3

### Phase H3 —— PPT Harness 对齐（2–3 周）

- `PptActionHarness` + 版式/母版语义
- 生成+美化闭环
- 场景 U4

### Phase H4 —— 体验与记忆产品化（2 周）

- 时间线、撤销本轮、变更高亮
- Memory 单写模型、隐私开关
- Plan 可编辑

### Phase H5 —— 平台化（持续）

- 原生 tool-calling
- 企业 Skill 目录
- 评测集 CI
- 连接器（可选 Graph/MCP 市场）

---

## 11. 专项设计清单

索引见 [`design/README.md`](./design/README.md)。

### 11.1 已完成（设计草案，评审用）

索引、**设计决策 D1–D15**：[`design/README.md`](./design/README.md) / [`design/design-review-record.md`](./design/design-review-record.md)。这些决策代表目标约束，不代表当前代码已全部实现。

| 专项 | 内容 |
|---|---|
| [`design/office-harness-api.md`](./design/office-harness-api.md) | IOfficeHarness、事件、WebView 协议（**H0 合同**） |
| [`design/context-pack-schema.md`](./design/context-pack-schema.md) | ContextPack |
| [`design/skill-runtime-and-gates.md`](./design/skill-runtime-and-gates.md) | Skill 两阶段与门禁 |
| [`design/safety-policy.md`](./design/safety-policy.md) | SafetyGate |
| [`design/tool-result-observation.md`](./design/tool-result-observation.md) | Observe / ToolResult |
| [`design/memory-runtime.md`](./design/memory-runtime.md) | memory_item 单写、task memory、隐私 |
| [`design/model-gateway-streaming.md`](./design/model-gateway-streaming.md) | IModelClient、流式 Pump、tool-calling 双模 |
| [`design/run-trace-storage.md`](./design/run-trace-storage.md) | agent_run 存储 |
| [`design/eval-golden-set.md`](./design/eval-golden-set.md) | 黄金评测集 |
| [`design/word-capability-map.md`](./design/word-capability-map.md) | Word Capability 地图 |
| [`design/excel-table-agent-runtime.md`](./design/excel-table-agent-runtime.md) | Excel 表运行时 |
| [`design/ppt-deck-agent-runtime.md`](./design/ppt-deck-agent-runtime.md) | PPT Deck 运行时 |

### 11.2 可选后续加深

| 专项 | 内容 |
|---|---|
| `design/prompt-profile-system.md` | 阶段提示装配 |

---

## 12. 决策记录（**目标设计决策**，详见 `design/design-review-record.md`）

> 下表是后续实现必须收敛到的目标约束。当前未完成项已经在 [`ai-native-harness-implementation-design.md`](./ai-native-harness-implementation-design.md) 拆成可执行任务。

| ID | 议题 | 冻结决策 |
|---|---|---|
| D1 | 产品入口 | 唯一 `IOfficeHarness.RunAsync` |
| D2 | 审批 | 模式 B：立即返回 awaiting_approval |
| D3 | 并发 | 同 session 仅 1 running |
| D4 | VBA | 默认关闭 |
| D5 | Skill 工具 | allowed-tools 硬拒绝 |
| D6–D7 | Memory | memory_item 单写；短期不进 RAG |
| D8–D9 | Observe | 写必有 observation；快路径必 Trace |
| D10–D11 | Gateway | Pump 降级；JSON Plan 默认，toolCalling=auto |
| D12 | PowerQuery | v1 不做完整执行 |
| D13 | CI | PR=L0+smoke+build-code；L1 Nightly |
| D14 | 无 Skill | 全 app 工具 + Safety |
| D15 | 阈值 | Safety T1/T2 权威 |
| — | UseNewAgentKernel | 废弃，始终 Harness |
| — | 云端索引 / 通用多 Agent 并行 | 暂缓；只读分诊子任务见执行计划 E17（写操作仍串行） |

---

## 13. 成功度量（产品 OKR 建议）

| 指标 | 3 个月目标（建议） |
|---|---|
| 任务自动完成率（无需用户改 prompt） | ≥ 60% 黄金集 |
| 平均 repair 次数 | ≤ 1.0 |
| 用户撤销率 | ≤ 15% |
| 工具失败可归类率（有 ErrorCode） | ≥ 95% |
| 空上下文执行率 | 0（生产） |
| Skill 命中后 allowed-tools 违规 | 0 |
| 首 Plan 展示时间 | P50 < 3s（排除网络） |

---

## 14. 风险与依赖

| 风险 | 缓解 |
|---|---|
| COM 不稳定 / 线程 | UiDispatcher + Executor 单线程队列 |
| 大文档爆上下文 | Budget + 二次 read 工具 |
| 模型乱调工具 | Plan 校验 + Broker 门禁 |
| 宿主能力不均 | Word 样板复制到 Excel/PPT |
| 工程债入口类过大 | 禁止在 ChatControl 加业务；只扩 Harness |
| 评测缺失导致回归 | 黄金集与 smoke 并行 |

---

## 15. 附录 A：建议的代码落位（实现时）

```text
ShareRibbon/
  Agent/
    Harness/
      IOfficeHarness.vb
      OfficeHarness.vb
      Models/          # UserTurn, ContextPack, ...
    Context/
      IContextHub.vb
      ContextBudget.vb
    Skills/            # 或沿用 Services
    Tools/
      IToolBroker.vb   # 包装 ToolRegistry
    Safety/
      SafetyGate.vb
  ...
WordAi/Services/Capabilities/
ExcelAi/Services/Capabilities/
PowerPointAi/Services/Capabilities/
```

Surface（`BaseChatControl`）最终应近似：

```text
Await OfficeHarness.RunAsync(turn)
' 订阅 OnPlan / OnStep / OnNeedApproval / OnExplain 事件刷新 UI
```

---

## 16. 附录 B：当前 → 目标 能力热力图

| 域 | 今 | 目标 |
|---|---|---|
| 统一入口 | ███░ | ████ |
| 上下文深度 | ██░░ | ████ |
| 规划质量 | ███░ | ████ |
| 工具边界 | ██░░ | ████ |
| 观察/修复 | ██░░ | ████ |
| Word 执行 | ████ | ████ |
| Excel 执行 | ███░ | ████ |
| PPT 执行 | ██░░ | ████ |
| Memory | ███░ | ████ |
| Safety | ██░░ | ████ |
| UX 可解释 | ███░ | ████ |
| 评测体系 | █░░░ | ████ |

---

## 17. 结语与下一步建议

当前项目 **已经站在 AI Native 门槛内**：有 Runtime、Kernel、Loop、Skills、Tools、Memory、Word 样板与工程治理基础。距离「对标 Copilot 体验 + Cursor 执行力」的关键跃迁，不是再堆入口功能，而是：

1. **把 Harness 做成唯一大脑与唯一手**；
2. **把观察（文档差分）做成一等公民**；
3. **把 Skill/Tool 边界做成硬门禁**；
4. **把 Word 样板复制到 Excel/PPT**；
5. **用黄金场景评测锁住回归**。

**建议的立即动作（设计落地，非本轮编码）**：

1. 以 [`ai-native-harness-implementation-design.md`](./ai-native-harness-implementation-design.md) 为唯一落地入口；
2. H0/H1/H2 的主路径基础已落地：`OfficeHarness` adapter、`ToolExecutionContext`、`allowed-tools` 执行期门禁、Word 读工具 Data 回传、Word 基础写工具 before/after/diff Observation、轻量 RunTrace；
3. 下一步补 Excel/PPT DocumentDiff、Golden 回归、`agent_run_event` 和回放 UI，避免一口气推倒重构。

---

*本文是规划基线，随 Phase H0–H5 推进应更新「热力图」与「已冻结契约」版本号，避免与过时 roadmap 脱节。*
