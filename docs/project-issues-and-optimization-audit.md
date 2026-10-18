# Office AI Agent 项目问题与优化审计报告

> **审计日期**：2026-07-11  
> **审计范围**：`ShareRibbon` / `WordAi` / `ExcelAi` / `PowerPointAi` / `OfficeAgent` / `scripts` / `openspec` / 前端资源  
> **方法**：全库源码规模与热点扫描 + 关键路径精读 + 既有 roadmap/评估文档对照  
> **相关历史文档**：`docs/project-assessment-optimization-plan.md`、`openspec/changes/global-architecture-hardening-plan.md`、`openspec/changes/ai-native-productization-roadmap.md`  
> **说明**：本报告基于当前仓库实态重新审计，不替代历史执行记录；已落地项会标注“部分缓解”，未完成项按优先级列入优化清单。2026-07-11 复核后修正了部分偏旧量化指标，并补充了“正确但需澄清”的事项。

## 0. 复核修订说明

本节记录对本报告的二次复核结论，避免后续执行任务基于偏旧或过度泛化的信息拆分。

### 0.1 复核后确认正确的主判断

- **入口类过重**、**多执行路径并存**、**UI/异步边界风险**、**Skills/Memory/LLM 双轨**、**发布验证偏手工** 这些主结论成立。
- `ToolResult` 类型已经存在，但结构化结果没有贯穿所有 Tool 和 observe/repair 闭环；问题不是“没有 ToolResult”，而是“没有成为强制契约”。
- `OfficeContext` 在 `BaseChatControl` 主路径中已有注入，但 `AgentKernel.ExecuteAsync` 仍允许空上下文兜底，生产路径仍应收紧。

### 0.2 复核后修正的事实

- 超大文件行数比初稿更高：`WordAi/ChatControl.vb` 约 4726 行、`ShareRibbon/Controls/BaseChatControl.vb` 约 3213 行、`PowerPointAi/ChatControl.vb` 约 2654 行。
- VB 源文件约 290 个，总行数约 85,441；初稿的 269 / 84,000 偏旧。
- 真正无条件裸 `Catch` 约 31 个，不是 280；`Catch ex As Exception` 约 1063 个仍准确。
- 当前工作区不是 git repo，无法确认 pfx/snk 是否被 Git 跟踪；只能确认签名密钥文件位于仓库目录内。
- LLM 请求路径问题的重点不是简单的 `New HttpClient` 数量；项目已有 `HttpClientPool`，真正风险是 provider 协议、请求体、鉴权、超时、错误码和脱敏日志没有统一。

### 0.3 复核后新增的文档漂移

- `ShareRibbon/AGENTS.md` 中 “Never access Office interop directly outside ShareRibbon” 与根目录规则相反。根规则要求 `ShareRibbon` 承载共享抽象，具体 Word/Excel/PowerPoint COM 实现应留在各宿主项目。
- `WordAi/AGENTS.md` 中 “Uses Word interop through ShareRibbon services” 也容易误导；当前推荐方向是 Word 具体执行器在 `WordAi/Services`，共享层只保留抽象和通用协议。

---

## 1. 执行摘要

### 1.1 一句话结论

项目已经具备 **AI Native Office Agent 的骨架**（`AiNativeRuntime`、`AgentKernel`/`LoopEngine`、Skills、Memory、MCP、DB 迁移、WebView2 Chat），但仍处于 **“聊天式插件 → 可执行 Agent” 的中后期迁移态**：新旧路径并存、超大入口类、线程边界、发布链路和验证体系是当前最大风险。

### 1.2 健康度评分

| 维度 | 分数 | 说明 |
|---|---:|---|
| 产品能力完整度 | 78/100 | 三端插件、Chat、排版/校对、翻译、MCP、Memory、Skills 均已有能力 |
| 架构清晰度 | 52/100 | 目标架构清楚，但运行时分叉多，主链路仍缠在入口类 |
| 可维护性 | 45/100 | 多个 1k–4k 行上帝类，局部修改影响面大 |
| 可靠性 / 线程安全 | 48/100 | 已引入 `UiDispatcher`，仍有 `.Result`/同步阻塞与宽泛 catch |
| 工程化（构建/安装/CI） | 55/100 | 有 smoke 与构建脚本，无完整 CI；`.vdproj` 仍脆弱 |
| 可测试性 | 40/100 | 仅有 PowerShell smoke，无单元/契约测试项目 |
| 安全与隐私 | 58/100 | API Key 有安全存储，日志/签名密钥/数据路径仍需产品化 |
| **综合** | **54/100** | 方向正确，工程债与迁移债集中，不适合继续“在入口层堆功能” |

### 1.3 问题统计（本轮）

| 级别 | 数量 | 含义 |
|---|---:|---|
| **P0 严重** | 5 + 1 待验证 | 直接影响稳定性、可交付性或架构主路径正确性；签名密钥项需确认真实 Git 跟踪状态后定级 |
| **P1 高** | 10 | 显著增加缺陷率、回归成本或产品体验分叉 |
| **P2 中** | 12 | 技术债、文档漂移、能力不均、长期治理项 |
| **P3 低** | 若干 | 命名、占位文案、重复工具类等 |

### 1.4 最关键的 5 个发现

1. **入口层仍然过重**：`WordAi/ChatControl.vb` ~4726 行、`BaseChatControl.vb` ~3213 行、`PowerPointAi/ChatControl.vb` ~2654 行，是变更热点与缺陷温床。  
2. **多套执行路径并存**：`AiNativeRuntime` → `AgentKernel`/`LoopEngine`、普通 Chat streaming、`ExecuteJsonCommand`、`WordActionHarness`、`SelfCheckLoopController` 同时存在，同一用户请求可能走不同闭环质量。  
3. **LLM / Memory / Skills 双轨**：`AiGateway` 与 `HttpStreamService`/`LLMUtil` 并存；`atomic_memory` 与 `memory_item` 并存；JSON Skill 与目录型 `SKILL.md` 并存。  
4. **线程与同步等待风险未清零**：`MemoryService`、`ExcelDnaFunctions`、`WebViewService`、`BaseChatControl` 等仍有 `.Result` / `GetAwaiter().GetResult()` / `Task.Run` 混用。  
5. **发布与验证偏手工**：代码项目可构建，但安装项目强依赖 `bin\Release`；无 CI 流水线；测试以 smoke 为主，Agent 行为回归成本高。

---

## 2. 系统规模与现状快照

### 2.1 代码规模（约）

| 模块 | VB 文件数 | 约行数 | 职责 |
|---|---:|---:|---|
| `ShareRibbon` | 206 | 59,327 | 共享 Chat、Agent、Loop、Config、MCP、Storage、前端资源 |
| `WordAi` | 34 | 12,007 | Word 宿主执行、校对/排版 harness、翻译 |
| `ExcelAi` | 28 | 7,986 | Excel 命令执行、ExcelDna、批量数据 |
| `PowerPointAi` | 22 | 6,121 | PPT 命令执行、续写/翻译 |
| **合计** | **~290** | **~85,441** | 另有 23 个 JS、34 个 SQL、72 个 Tool schema JSON |

### 2.2 超大文件 TOP（维护风险）

| 行数 | 文件 | 风险 |
|---:|---|---|
| 4726 | `WordAi/ChatControl.vb` | 宿主执行 + 路由 + 校对/排版 + 工具执行 |
| 3213 | `ShareRibbon/Controls/BaseChatControl.vb` | UI / WebView / 路由 / Agent / 发送主链路 |
| 2654 | `PowerPointAi/ChatControl.vb` | 与 Word 类似的宿主入口膨胀 |
| 2309 | `ShareRibbon/Config/ConfigApiForm.vb` | 配置 UI 与业务耦合 |
| 2278 | `ShareRibbon/Controls/BaseDataCapturePane.vb` | 采集面板职责过重 |
| 2118 | `ExcelAi/ExcelDirectOperationService.vb` | Excel 执行器上帝服务 |
| 2108 | `ShareRibbon/Controls/Services/IntentRecognitionService.vb` | 意图识别过大，含 LLM 调用与规则 |
| 1656 | `ShareRibbon/Mcp/MCPConfigForm.vb` | 配置窗体膨胀 |
| 1500 | `ShareRibbon/Services/Reformat/SmartFormattingOrchestrator.vb` | 排版编排复杂 |
| 1439 | `ShareRibbon/Controls/Services/HttpStreamService.vb` | 流式 LLM 与 provider 适配 |

### 2.3 已有正向资产（应保留）

- AI Native 中枢：`ShareRibbon/Agent/AiNativeRuntime.vb`
- 统一 Agent：`AgentKernel` + `LoopEngine`（ReAct：Think → Plan → Act → Observe）
- Word 样板：`WordActionHarness` + `WordCapabilityRegistry`
- 非流式 LLM 网关：`ShareRibbon/Services/Ai/AiGateway.vb`（Intent / Memory / Autocomplete 已部分迁移）
- UI 调度：`UiDispatcher`（部分 WebView 路径已接入）
- 请求拆分：`ChatRequestOrchestrator`、`ChatCommandRouter`、`WebViewBridge`、`MemoryTurnRecorder` 等
- 数据库：`Storage/Migrations/*.sql` + schema drift smoke
- 验证脚本：`scripts/smoke-*.ps1`、`scripts/build-code-projects.ps1`、`build/RunReleaseChecks.ps1`
- 目录型 Skills：`excel-table-agent` / `word-document-agent` / `powerpoint-deck-agent` / `office-skill-authoring`

### 2.4 目标架构（应收敛到）

```text
Chat / Ribbon / 快捷入口
        │  只收集上下文、展示结果
        ▼
   AiNativeRuntime（分析：意图 + 上下文 + Skills + Tools）
        │
        ▼
   AgentKernel / LoopEngine
   plan → act → observe → repair → explain
        │
        ├── Capability / Skill（allowed-tools 边界）
        ├── Tool schema（ShareRibbon/Tools/**）
        └── Host Executor（WordAi / ExcelAi / PowerPointAi）
                │
                ▼
        Observe / Verifier / Undo / 用户解释
```

**反模式（当前仍多处存在）**：在 `ChatControl` / `BaseChatControl` 中用 `Select Case`、关键词、直接 `ExecuteJsonCommand` 决定主业务路径。

---

## 3. P0 严重问题

### P0-1 超大入口类承担过多职责

**状态**：部分完成（2026-07-11）

#### 已完成

1. ~~`BaseChatControl` 智能路由拆到 `ChatRoutingOrchestrator`~~  
   - 文件：`ShareRibbon/Controls/Services/ChatRoutingOrchestrator.vb` + `IChatRoutingHost`  
   - `HandleSendMessageCore` 智能模式只委托 `RouteSmartModeAsync`  
   - MSBuild 四项目 Debug 通过  

#### 未完成（后续迭代，本轮仅标记）

| 编号 | 项 | 说明 | 建议优先级 |
|---|---|---|---|
| **P0-1-R1** | 三端 `ChatControl` 瘦身 | 只保留上下文采集、宿主 executor 注册、宿主特有 UI；当前 Word ~4k / PPT ~2k 行仍过大 | P1 迭代 |
| **P0-1-R2** | Word 工具执行下沉 | 将 `WordAi/ChatControl` 内大段工具/命令执行迁到 `WordAi/Services/*Executor` | P1 迭代 |
| **P0-1-R3** | >800 行类门禁 | `scripts/audit-p0-guardrails.ps1` 已固化现有超大文件基线，禁止新增 >800 行 VB 文件 | 已接入 code checks |

**证据（体量，复核值）**

- `WordAi/ChatControl.vb` ~4726 行  
- `BaseChatControl.vb` ~3213 行（路由拆出后会略降，仍远超健康阈值）  
- `PowerPointAi/ChatControl.vb` ~2654 行  
- `ConfigApiForm` / `BaseDataCapturePane` 均 >2000 行  

---

### P0-2 多套运行时 / 执行路径并存

**状态**：部分完成（2026-07-11）— 主路径策略与智能路由已收敛；Capability 回灌 Loop、删除 ralph-loop.js 文件仍待后续。

**证据**

| 路径 | 代表符号 | 角色（P0-2 后） |
|---|---|---|
| AI Native 分析 | `AiNativeRuntime` | 主路径第一步 |
| 新 Agent | `AgentKernel` / `LoopEngine` | **唯一智能模式产品主路径** |
| 策略 | `ExecutionPathPolicy` | 主路径/兼容边界单一真相源 |
| Word 专用 | `WordActionHarness` | 宿主 capability **快路径**（高置信确定动作），非并行 NL 路由器 |
| 旧 JSON 命令 | `ExecuteJsonCommand` | **Tool 后端 only**，不再作为 NL 入口 |
| 切换开关 | `UseNewAgentKernel` | **智能路由已忽略**；保留属性兼容，false 仅打警告日志 |
| Ralph | `startLoop` / `ralph-loop.js` | 兼容 shim → `startAgent`；`/loop` 前端已直发 `startAgent` |

**现象（收敛前）**

- 智能模式可按开关走 Agent 或旧 Chat  
- Word harness / JSON command / Ralph 形成多套闭环质量  

**本轮落地**

1. ~~明确唯一主路径~~：`ShareRibbon/Agent/ExecutionPathPolicy.vb` + `ChatRoutingOrchestrator` 始终 `Analyze → AgentKernel`（追问仍走 plain chat 以避免每轮重规划；Agent 启动失败可 fallback chat）。  
2. ~~`ExecuteJsonCommand` 降级为 Tool 后端~~：注释与 `CodeExecutionService` 标记 tool backend only。  
3. ~~`WordActionHarness` 角色澄清~~：标注 capability fast-path，受 `AllowHostCapabilityFastPath` 控制；observe/explain 回灌 Loop **仍待做**。  
4. ~~隔离 Ralph 主入口~~：`message-sender.js` 的 `/loop` 改为 `startAgent`；`chat-template-refactored.html` 不再加载 `ralph-loop.js`；后端 `startLoop` 仍 remap 到 `startAgent` 并打日志。  
5. ~~智能路由移除 `UseNewAgentKernel=false` 分支~~：属性保留但 smart-mode 不再读取。  
6. **VSTO shadow-copy 工具目录修复（追加）**：`AgentKernel` 不再只依赖 `Assembly.Location`，会优先使用 `Assembly.CodeBase` / AppDomain 候选目录解析 `Tools/Prompts/Skills`，避免 VSTO 缓存目录下找不到工具。  
7. **0 迭代误报成功修复（追加）**：`LoopEngine` 若没有执行任何工具调用，或仍存在未完成步骤，会返回失败而不是“任务完成，共执行 0 个迭代”。  

**仍待后续（P0-2 残余）**

| 编号 | 项 |
|---|---|
| **P0-2-R1** | Word/Excel/PPT capability 执行结果结构化回灌 Agent observe/explain（原生 Tool 后端已能回传宿主 Boolean；Word `DeleteText` 已补齐；Word Numbering/DirectFormatting 已有真实成功/失败回灌；Proofread 已回灌启动成功/失败；SemanticReformat 已回灌预览生成/回退/失败；Excel/PPT 待做） |
| **P0-2-R2** | `ralph-loop.js` 已从主 HTML 加载链路移除；文件和 CSS 暂保留为兼容/回滚资产，`audit-p0-guardrails` 防止主入口回归 |
| **P0-2-R3** | 配置 UI/持久化层移除对 `UseNewAgentKernel` 的可写暴露（若有） |  

---

### P0-3 UI 线程 / 同步等待风险未清零

**状态**：部分完成（2026-07-11）— 高风险 sync-over-async 已统一到安全桥；ChatContext 仍同步调 Memory（经桥接），全量 async 调用链与 Task.Run 清单仍待后续。

#### 本轮落地

1. **硬规则基础设施**  
   - `ShareRibbon/Common/SyncOverAsync.vb`：线程池 + 超时，禁止在 UI SynchronizationContext 上 `.Result`  
   - `UiDispatcher.InvokeSync`：后台 → UI 用 `Control.Invoke`，替代 `InvokeAsync(...).GetResult()`  
2. **Memory**  
   - `GetRelevantMemories` / `GetRelevantStructuredMemories` / `SearchMemories` 内部改 async + `SyncOverAsync`  
   - 新增 `*Async` 重载；`UnifiedMemoryService.RetrieveMemories` 向量同样走桥  
3. **LLM / ExcelDna**  
   - `LLMUtil.SendHttpRequestSync` 改为桥接 `SendHttpRequest`，去掉裸 `.Result`  
   - `ExcelDnaFunctions` 主路径改 `SendHttpRequestSync`  
4. **UI / WebView / 其它**  
   - `BaseChatControl.RunUiActionSync` → `UiDispatcher.InvokeSync`  
   - `WebViewService.InvokeIfRequired` → `InvokeSync`  
   - `LlmMemoryExtractor` / `ModelApiClient.GetModelsSync` / `ExcelAi/EnhancedPreviewAndConfirm` 走 `SyncOverAsync`  
   - `ReformatCoordinator` 去掉对 `tcs.Task.Result` 的阻塞（用 MessageBox 已得 decision）  
   - MSBuild 四项目 Debug 通过（无 warning）  

#### 仍待后续（P0-3 残余）

| 编号 | 项 |
|---|---|
| **P0-3-R1** | `ChatContextBuilder` / Agent 主路径全面 `Await *Async`，减少同步桥调用 |
| **P0-3-R2** | 全库 `Task.Run` 清单：`scripts/audit-p0-guardrails.ps1` 已建立 Task.Run + Office Interop 基线，禁止新增未审查文件；既有 5 个文件待分批消除 |
| **P0-3-R3** | ExcelDna UDF 可选改为队列 + 单元格异步刷新（彻底避免同步 HTTP） |

**影响（收敛前）**

- STA/UI 死锁或卡顿  
- WebView2 “must be accessed from UI thread”  
- Office COM 跨线程异常 / RPC_E_WRONG_THREAD  

---

### P0-4 宽泛异常吞噬，可观测性差

**状态**：部分完成（2026-07-11）— 日志/错误契约与 Agent 主路径已落地；全库 Catch 迁移与宿主 Executor 全覆盖仍待后续。

#### 本轮落地

1. **统一错误契约**  
   - `ToolResult` 扩展：`ErrorCode` / `UserMessage` / `DebugDetail` / `Recoverable` + `FromException` + `ToObserveSummary`  
   - `OperationResult` 对齐同名字段 + `FailFromException`  
   - `ExceptionClassifier`：COM / 网络 / 超时 / JSON / 参数 / IO 分类  
2. **结构化日志**  
   - 新增 `AppLogger`（级别、模块、correlation id、文件落盘 `%LocalAppData%\OfficeAiAppData\logs\`、密钥脱敏）  
   - `SimpleLogger` 改为 facade → `AppLogger`  
   - `ErrorHandler` 写日志改为 `AppLogger`  
3. **Agent 路径 observe/repair**  
   - `LoopEngine`：失败观察用结构化摘要；不可恢复错误跳过无效 repair；reflect/replan 带 error code  
   - `AgentKernel.ExecuteAsync`：BeginScope + 失败/异常记日志  
   - `AgentKernelService`：启动失败用户可见提示 + `AppLogger`  
   - `ToolRegistry`：MCP/原生执行异常走 `FromException`  
4. **宿主执行结果回灌（追加）**  
   - `CodeExecutionService.ExecuteCodeWithResult`：JSON/VBA/JS/公式执行统一返回 Boolean  
   - `AgentKernelService` / `AgentKernel` / `ToolRegistry`：原生 Office Tool 使用 `Func(Of String, String, Boolean, Boolean)` 获取宿主执行结果  
   - `ToolRegistry`：宿主返回 `False` 时生成失败 `ToolResult`，进入 observe/repair，而不是误报“执行成功”  
   - `WordAi/ChatControl`：补齐 `DeleteText` 执行器，使已有 Word schema 的删除能力可执行  

#### 仍待后续（P0-4 残余）

| 编号 | 项 |
|---|---|
| **P0-4-R1** | 三端 `ChatControl` / `ExcelDirectOperationService` 高密度 Catch 分批改为分类捕获 + AppLogger |
| **P0-4-R2** | 全库禁止空 `Catch`（复核后真正无条件裸 `Catch` 约 31；所有 Catch 变体约 1495，需按模块清理） |
| **P0-4-R3** | 宿主 Executor 统一返回 `ToolResult`/`OperationResult`，禁止只 `Debug.WriteLine`（Agent 原生 Tool 后端已完成 Boolean 回灌；Word capability fast-path 已开始结构化回灌；Excel/PPT 专用 capability 仍待覆盖） |
| **P0-4-R4** | 可选：UI 日志查看入口（打开当日 log 文件） |

**证据（收敛前）**

- `Catch ex As Exception` ≈ **1063**  
- 裸 `Catch` 扫描约 **280**（含大量“吞异常返回默认值”）  
- `Debug.WriteLine` ≈ **1510**  

---

### P0-5 安装包与发布链路脆弱

**状态**：部分完成（2026-07-11）— code/installer 二分脚本与文档已落地；code-only CI 门禁已落地；WiX/MSIX 迁移仍待后续。

#### 本轮落地

1. **文档二分**：`docs/build-and-installer.md`（code build vs installer prep vs MSI）  
2. **脚本**  
   - `scripts/build-installer-prep.ps1`：Release 代码构建 + `AuditInstallerInputs`  
   - `build-code.bat` / `build-installer-prep.bat`  
   - `build-all.bat` 明确为 code-only 别名  
3. **导航更新**：`AGENTS.md` Commands、`VS2026-InstallerProjects.md` 交叉引用  
4. **澄清**：`OfficeAgentSetupCustomActions/` **当前不存在**；需要时再新建，勿大改 vdproj  
5. **code-only CI 门禁（追加）**  
   - `scripts/run-code-checks.ps1`：聚合四代码项目 Debug build、DB/Memory/Skills/Word/AiGateway smoke、关键 JS `node --check`  
   - `scripts/audit-p0-guardrails.ps1`：P0-1 超大文件基线、P0-2 Ralph 主入口回归、P0-3 Task.Run+Office Interop 基线  
   - `.github/workflows/code-checks.yml`：PR/push 运行 code-only checks，不构建 `.vdproj` / MSI  
   - 本地已验证：`powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\run-code-checks.ps1 -Configuration Debug` 通过  

#### 仍待后续（P0-5 残余）

| 编号 | 项 |
|---|---|
| **P0-5-R1** | 中期评估 WiX / MSIX / 可 CI 的安装方案 |
| **P0-5-R2** | PR CI：仅 code build + smoke（不编 MSI）已落地；后续根据真实 CI 环境补 Office/VSTO 依赖说明 |
| **P0-5-R3** | 若业务需要，新建 `OfficeAgentSetupCustomActions` 承接安装逻辑 |

**证据（收敛前）**

- `OfficeAgent.vdproj` SourcePath 指向 `bin\Release`  
- 整解决方案 Debug Rebuild 易在安装项目失败  
- 无完整 CI  

---

### P0-6 签名密钥文件位于仓库目录内

**状态**：部分完成（2026-07-11）— 已核实 Git 跟踪并取消索引跟踪 + 策略文档/生成脚本；历史改写与正式证书轮换仍待维护者决策。

#### 核实结论

- 仓库 **是** git repo（`origin` → `github.com:it235/office-ai-agent.git`）。  
- `git ls-files "*.pfx" "*.snk"` **曾列出** 四个密钥文件（含私钥的自签名 TemporaryKey + 未引用 snk）→ **升级为 P0**。  
- 证书 Subject 为本机用户、约 2026-06 过期；**不是**商业代码签名证书，但仍不应入库。  
- `OfficeAiAgent.snk`：**无任何 vbproj 引用**，可安全移出版本库。  

#### 本轮落地

1. `git rm --cached` 四个密钥（工作区文件保留；`git ls-files` 现为空）  
2. `.gitignore` 增加 `*.pfx` / `*.snk` / secrets 约定；放行 `docs/signing-and-certificates.md` 等  
3. 文档：`docs/signing-and-certificates.md`（Debug TemporaryKey vs Release `SignArtifacts`）  
4. 脚本：`scripts/ensure-vsto-temp-keys.ps1`（生成本机开发证、更新 ManifestCertificateThumbprint）  
5. 本机已 `-Force` 重生有效期至约 2028 的 TemporaryKey（**不提交 pfx**）  
6. `AGENTS.md` / `build-and-installer.md` 交叉引用  

#### 仍待后续（P0-6 残余）

| 编号 | 项 |
|---|---|
| **P0-6-R1** | 维护者 commit「取消跟踪密钥」变更并 push；评估是否 `git filter-repo` 清历史 |
| **P0-6-R2** | 若曾用同一私钥做任何对外分发，轮换证书 |
| **P0-6-R3** | CI 仅注入 `OFFICE_AI_SIGN_*` secret，从不检出 pfx |

**发布签名路径（已有）**：`build/SignArtifacts.ps1` + `OFFICE_AI_SIGN_CERT_THUMBPRINT` / `OFFICE_AI_SIGN_PFX`。  

---

## 4. P1 高优先级问题

### P1-1 LLM 请求路径未统一

**现状**

| 组件 | 用途 | 状态 |
|---|---|---|
| `AiGateway` | 非流式 chat | Intent/Memory/Autocomplete 已用 |
| `HttpStreamService` | 主聊天 streaming | 仍是主路径，provider 逻辑独立 |
| `LLMUtil` | 同步/异步 HTTP 工具 | ExcelDna、Batch、WPS 检测等仍依赖 |
| 翻译服务 | 自建 `SendHttpRequestAsync` | 未走 AiGateway |

**优化项**

1. 增加 `AiStreamingGateway` 或让 `HttpStreamService` 仅做 SSE 泵，请求体/鉴权/Reasoning 统一。  
2. 翻译、补全 FIM、BatchData 分批迁移。  
3. 统一超时、重试、限流、错误码与脱敏日志。  
4. 保持 `HttpClientPool` 为默认 HTTP 池；重点收敛 provider 协议、鉴权、超时、重试、错误码和脱敏日志，而不是只统计 `New HttpClient`。  

---

### P1-2 Memory 双模型与同步检索

**现状**

- 表：`atomic_memory`（旧）+ `memory_item` / `memory_embedding`（新）  
- 服务：`MemoryService`（Controls）+ `UnifiedMemoryService` + `AgentMemoryRepository`  
- `skills_usage` 表与 `skills_usage.json` 迁移兼容仍在  
- 同步 RAG 使用 `Task.Run` + `Wait` 防死锁，但仍阻塞调用线程  

**优化项**

1. 新写入只走 `conversation_event → memory_job → memory_item`。  
2. `atomic_memory` 只读兼容，设淘汰期限。  
3. Chat 上下文构建只调用 async 检索。  
4. 用 `smoke-memory-pipeline.ps1` 扩展：晋升、过期、冲突版本。  

---

### P1-3 Skills 双格式与边界未强制

**现状**

- 旧：`ShareRibbon/Skills/*.json`（`triggerPatterns` 关键词风格）  
- 新：目录型 `*/SKILL.md`（front matter + allowed-tools）  
- 加载：`SkillsDirectoryService` + `SkillRegistry.LoadFromDirectory` + `SkillsService`  
- AgentKernel 注释写明 filesystem Skill 优先、JSON 兜底  

**优化项**

1. 将 JSON skills 迁移/归档为目录型 Skill 或自动生成器。  
2. Loop 执行前强制校验 tool ∈ `allowed-tools`。  
3. 第一阶段只加载 front matter，命中后再加载正文（避免 prompt 膨胀）。  
4. Excel/Word/PPT 用 `application` 字段隔离，禁止跨宿主 tool。  

---

### P1-4 三端能力不对称

| 能力 | Word | Excel | PowerPoint |
|---|---|---|---|
| ActionHarness / CapabilityRegistry | 有 | 弱/无对等 | 弱/无对等 |
| ChatControl 体量 | 最大 | 中 | 大 |
| 占位/未完成 | ToolResult 回传 TODO | PowerQuery“开发中” | 未知命令 warning |
| 上下文 Provider | 有 | 有 | 有 |

**优化项**

1. 以 Word 为样板，抽象 `IHostActionHarness`。  
2. Excel 优先补齐 table/formula/chart capability + observe。  
3. PPT 补齐 slide/shape 生成与“未知命令 → repair”而不是仅 warning。  
4. 统一 ToolResult 结构化回传（解决 Word TODO）。  

---

### P1-5 测试体系偏 smoke

**现状**

- 有：`smoke-ai-gateway-provider`、`smoke-db-schema-drift`、`smoke-memory-pipeline`、`smoke-skills-*`、`smoke-word-capability-registry` 等  
- 无：`*Tests.vbproj` / 单元测试项目  
- 无：CI 自动跑 smoke  

**优化项**

1. 新增 `OfficeAi.Tests`（.NET Framework 4.7.2）测纯逻辑：  
   - AiGateway provider 转换  
   - Skill front matter 解析  
   - Formatting/Proofread IntentCompiler  
   - DB migration 幂等  
2. smoke 接入 CI（至少 PR 必跑不依赖 Office 的脚本）。  
3. 关键 Agent 场景建立“黄金对话”回放（可先录请求/响应 fixture）。  

---

### P1-6 前端遗留与超大 JS

**现状**

- `ralph-loop.js`（11.7KB）仍被 `message-sender.js` / `chat-manager.js` 引用 `startLoop`  
- 后端已把 `startLoop` 映射到新 Agent 兼容处理  
- 大文件：`code-handler.js` 79KB、`reformat-template.js` 54KB、`message-sender.js` 32KB  

**优化项**

1. 前端统一到 `agent-card.js` / `agent-protocol.js`，删除 Ralph 主 UI。  
2. `code-handler.js` 拆分为 parse / preview / apply。  
3. 每次改 JS 必跑 `node --check`，并检查 Virtual Server 路径与 `.vbproj` 注册。  

---

### P1-7 数据库迁移双源与文档漂移

**现状**

- 外部：`ShareRibbon/Storage/Migrations/001_*.sql` … `010_*.sql`  
- 内联：`OfficeAiDatabase.GetMigrationSql()` 仍有大段 fallback SQL  
- `README_DATABASE.md` 写路径 `%Documents%\OfficeAi\office_ai.db`  
- 代码实际：`%Documents%\OfficeAiAppData\office_ai.db`（Debug 为 `OfficeAiAppData-Debug`）  

**优化项**

1. 内联 SQL 降为“最小应急 fallback”，权威源只保留 Migrations + `OfficeAiDbSchema.current.sql`。  
2. 修正 README 路径与初始化时机说明。  
3. 保持 `smoke-db-schema-drift.ps1` 为发布门禁。  

---

### P1-8 依赖与数据访问栈偏重

**现状**

- `ShareRibbon` 34 个 NuGet 包  
- 同时存在 `System.Data.SQLite` + `Microsoft.Data.Sqlite.Core` + `SQLitePCLRaw` + `EntityFramework 6`  
- JSON：`Newtonsoft.Json` 引用约 254 处，`System.Text.Json` 约 7 处  

**优化项**

1. 明确 SQLite 唯一入口（建议继续 `System.Data.SQLite` 主路径，评估是否可去掉 EF6 若几乎不用）。  
2. 新代码统一 Newtonsoft 或统一 STJ，避免双栈。  
3. 定期审计未使用包，减小安装体积与绑定重定向风险。  

---

### P1-9 重复基础设施

**证据**

- `ShareRibbon/Common/ComObjectHelper.vb` 与 `ShareRibbon/Utils/ComObjectHelper.vb`  
- Undo：`Core/UndoManager.vb`、`Undo/UndoStack.vb`、`Extensions/UndoManagerExtension.vb`  
- 三端重复：`DeepseekControl` / `DoubaoChat` / `Ribbon1` / `ChatControl` 平行结构  

**优化项**

1. 合并 ComObjectHelper 与 Undo 抽象，保留一个公共 API。  
2. Deepseek/Doubao 评估是否并入主 Chat（减少双 WebView 维护）。  
3. Ribbon 公共项尽量进 `ShareRibbon/Ribbon`。  

---

### P1-10 Agent 上下文与 Tool 结果闭环不完整

**证据**

- `AgentKernel.ExecuteAsync`：未传入 `officeContext` 时创建空上下文（TODO）  
- Word：`ExecuteListParagraphs` / `ExecuteGetParagraphInfo` 结果未结构化回传 AI（TODO）  
- `BaseChatControl`：新架构下“修改计划”暂不支持  
- `ToolResult` 类型已存在，且包含 `Data` 字段；缺口在于 read-tool/host executor 没有强制填充结构化数据，Loop 执行前也未强制校验 Skill `allowed-tools` 边界  

**影响**

- Agent 可能在“看不见文档”的情况下规划  
- observe 阶段信息不足，repair 质量差  

**优化项**

1. 强制生产路径 `StartAgentAsync` 必须注入 `CaptureOfficeContext`；空上下文只允许测试或明确降级路径。  
2. 所有 read-tool 必须返回 `ToolResult.Data` JSON。  
3. Loop 执行前校验 `toolId` 属于当前 app 可用工具与命中 Skill 的 `allowed-tools` / `RequiredTools` 交集。  
4. 支持基于 observe 的 replan（有限次数）。  

---

## 5. P2 中优先级问题

### P2-1 产品占位与体验落差

- Excel PowerQuery：`PowerQuery代码执行功能正在开发中`  
- PPT：`暂不支持的PPT命令: {command}`  
- 应改为：能力探测 → Agent 换 tool → 明确“可撤销的替代方案”，避免“开发中”文案  

### P2-2 配置模型拼写与开关膨胀

- `ConfigSettings.propmtName` / `propmtContent` 历史拼写错误  
- `UseNewAgentKernel`、`UseLoopFramework`、`UseAdvancedPlanning`、`MaxLoopIterations` 等多开关增加组合爆炸  

**建议**：配置分组、废弃开关列表、启动时打印有效配置快照（脱敏）。  

### P2-3 日志与隐私产品化不足

- 数据在 `Documents\OfficeAiAppData`，备份/清除/导出流程不清晰  
- Debug 与正式库分离是优点，但用户文档未同步  
- 需提供：清除记忆、关闭 RAG、导出会话、日志脱敏开关  

### P2-4 翻译链路重复实现

- `ShareRibbon/Translate/*`、Word/Excel/PPT DocumentTranslateService 多套 HTTP 与批处理  
- 应抽到共享 batch translator + AiGateway  

### P2-5 IntentRecognitionService 过大

- 2108 行，混合规则、LLM、prompt、计划预览  
- 建议拆：`IntentClassifier` / `IntentPromptFactory` / `ExecutionPlanPreviewBuilder`  

### P2-6 ExcelDirectOperationService 上帝服务

- 2118 行，命令执行集中  
- 建议按领域拆：Range / Sheet / Chart / Pivot / Protect  

### P2-7 前端资源与 Virtual Server 风险

- 历史问题：资源未进 vbproj 或路径大小写导致 `ERR_FILE_NOT_FOUND`  
- 建议增加 “HTML 引用 vs vbproj 注册” 的静态对账脚本  

### P2-8 openspec 与实现进度不同步风险

- roadmap 中大量 `[x]`，但代码仍有双路径与 TODO  
- 建议每个 `[x]` 必须链到验证命令与残留风险说明  

### P2-9 CodeGraph / 架构索引不完整

- 仓库存在 `.codegraph/`，但本轮架构图工具返回空社区/空 hub  
- 建议重建索引，便于后续影响面分析  

### P2-10 无自动化 CI

- 缺少 PR 级：restore → build 四项目 → smoke → node --check → git diff --check  

### P2-11 WPS 兼容分支分散

- `LLMUtil.IsWpsActive()` 在多处 ThisAddIn 使用  
- 建议集中 `HostEnvironment` 检测与任务窗格差异  

### P2-12 文档与导航轻微过时

- `AGENTS.md` 曾写 `OfficeAgentSetupCustomActions/`（如存在）；该目录未被 sln/vdproj 引用，已于 2026-10-18 删除并同步文档  
- `README_DATABASE.md` 路径错误  
- 局部 AGENTS 需与真实目录对齐  
- `ShareRibbon/AGENTS.md` 关于 Office interop 边界的表述与根规则冲突，应改为“共享层承载抽象，具体宿主 COM 实现留在对应 Office 项目”  
- `WordAi/AGENTS.md` 关于 “through ShareRibbon services” 的表述容易误导，应改为 Word 具体执行器优先在 `WordAi/Services`  

---

## 6. 架构专项观察

### 6.1 分层

| 层 | 现状 | 评价 |
|---|---|---|
| 表现层（WebView/Ribbon） | 较完整，但协议新旧并存 | 中 |
| 应用协调层 | `AiNativeRuntime` + `BaseChatControl` 路由 | 偏重，应下沉 |
| Agent/Loop | `AgentKernel`/`LoopEngine`/`SelfCheckLoop` | 方向正确 |
| 领域执行 | Word 较强，Excel/PPT 偏命令式 | 不均 |
| 基础设施 | SQLite、Http、MCP、Config | 可用但双轨 |

### 6.2 领域边界

- **优点**：`ShareRibbon` 未直接引用 Word/Excel/PowerPoint Interop 实现类型（检索未见具体 Interop using；`ExcelContextService` 仅处理数组数据，可接受）。  
- **风险**：共享层仍含 Excel 特化命名服务、巨大 Chat 基类知道过多宿主流程。  

### 6.3 隐式依赖

- WebView 字符串协议（`startAgent`、`startLoop`、JS 函数名）无强类型契约  
- `ConfigSettings` 全局可变静态状态  
- Skill/Tool 目录依赖运行时 `Assembly.Location`  

### 6.4 变更影响面最高的文件

优先保护（改前必看调用方 + 必跑三端构建）：

1. `ShareRibbon/Controls/BaseChatControl.vb`  
2. `ShareRibbon/Controls/Services/HttpStreamService.vb`  
3. `ShareRibbon/Agent/AgentKernel.vb` / `LoopEngine.vb`  
4. `WordAi/ChatControl.vb`  
5. `ShareRibbon/Storage/OfficeAiDatabase.vb`  
6. `ShareRibbon/Services/Ai/AiGateway.vb`  

---

## 7. 优化路线图（建议执行顺序）

### Phase 0 — 审计基线修正与门禁准备（1–2 天）

| # | 项 | 验收 |
|---|---|---|
| 0.1 | 固化审计统计命令 | 能复现 VB 文件数、LOC、Catch、Task.Run、Tool JSON、Skill 数量 |
| 0.2 | 修正导航文档漂移 | `ShareRibbon/AGENTS.md`、`WordAi/AGENTS.md` 与根规则一致 |
| 0.3 | 明确 P0/P1 分级口径 | P0 仅保留阻塞交付/稳定性主路径问题，安全泄露需先验证真实 Git 跟踪 |
| 0.4 | 建立高风险清单文件 | 同步等待、旧路径、TODO ToolResult、安装输入列成可勾选任务 |

### Phase A — 稳定交付面（1–2 周）

| # | 项 | 验收 |
|---|---|---|
| A1 | 文档化 code-build vs installer-build | `scripts/build-code-projects.ps1` 为默认开发入口 |
| A2 | Release 安装输入审计默认进打包手册 | `AuditInstallerInputs` 失败即停 |
| A3 | 清理/隔离 TemporaryKey 策略说明 | 真实 Git 跟踪状态已复核；发布不依赖仓库 pfx |
| A4 | 修正 `README_DATABASE.md` 路径 | 与代码一致 |
| A5 | 高危同步等待清单整改第一批 | Memory 异步化、去除 UI 路径 GetResult |

### Phase B — 主路径收敛（2–4 周）

| # | 项 | 验收 |
|---|---|---|
| B1 | Chat 路由编排器从 BaseChatControl 拆出 | BaseChatControl 行数明显下降 |
| B2 | 强制 Agent 注入 OfficeContext | 空上下文仅测试允许 |
| B3 | ToolResult 结构化回传 | Word 两个 TODO 关闭 |
| B4 | Tool/Skill 边界执行期校验 | toolId 不在 app 可用工具或 Skill 允许范围时返回结构化失败 |
| B5 | 前端废弃 Ralph 主路径 | `startLoop` 仅兼容或删除 |
| B6 | Skills JSON → 目录型迁移计划 | allowed-tools 执行期强制 |

### Phase C — LLM / Memory 统一（2–3 周）

| # | 项 | 验收 |
|---|---|---|
| C1 | Streaming 走统一网关适配层 | provider 行为一致 |
| C2 | 翻译/Batch 迁 AiGateway | LLMUtil 调用点减少 |
| C3 | atomic_memory 只读 + 迁移完成标记 | 新写入 100% memory_item |
| C4 | 日志脱敏与级别化 | 无 Key 前缀日志 |
| C5 | 明确 HTTP 池与请求策略边界 | `HttpClientPool` 继续池化，AiGateway/streaming 统一协议与错误模型 |

### Phase D — 三端 Capability 对齐（3–6 周）

| # | 项 | 验收 |
|---|---|---|
| D1 | ExcelActionHarness | 公式/图表/清洗闭环 |
| D2 | PptActionHarness | 生成/排版闭环 |
| D3 | 未知命令 → repair 而非 warning | 产品文案无“开发中”主路径 |
| D4 | 统一 Undo 语义 | 用户可理解可撤销范围 |
| D5 | 三端 read-tool 均返回 ToolResult.Data | observe/replan 能拿到宿主状态 JSON |

### Phase E — 工程化（并行）

| # | 项 | 验收 |
|---|---|---|
| E1 | `OfficeAi.Tests` 项目 | 纯逻辑测试进 CI |
| E2 | PR CI：build + smoke + node --check | 无 Office 依赖的门禁全绿 |
| E3 | HTML/vbproj 资源对账脚本 | 防止 WebView 404 |
| E4 | CodeGraph 索引重建 | 架构审计可重复 |
| E5 | 安装构建二分脚本化 | code build 与 installer build 输出、失败原因清晰分离 |

---

## 8. 建议的“立即不做”

1. **不要**继续在 `WordAi/ChatControl.vb` / `BaseChatControl.vb` 上堆新业务分支。  
2. **不要**大范围重写 `.vdproj` 或全文件格式化。  
3. **不要**同时开启多条记忆写入链路的新入口。  
4. **不要**在未强制执行 ToolResult/allowed-tools 契约前再加更多只写不读的 Office 命令。  
5. **不要**把宿主 COM 类型引入 `ShareRibbon`。  

---

## 9. 验证清单（审计后开发共用）

每完成一个优化项，至少执行与其范围匹配的检查：

```powershell
# 代码构建（开发默认）
powershell -File .\scripts\build-code-projects.ps1 -Configuration Debug

# 与改动相关的 smoke
powershell -File .\scripts\smoke-db-schema-drift.ps1
powershell -File .\scripts\smoke-memory-pipeline.ps1
powershell -File .\scripts\smoke-skills-registry.ps1
powershell -File .\scripts\smoke-ai-gateway-provider.ps1
powershell -File .\scripts\smoke-word-capability-registry.ps1

# 前端
node --check .\ShareRibbon\Resources\js\message-sender.js
node --check .\ShareRibbon\Resources\js\agent-card.js
node --check .\ShareRibbon\Resources\js\office-ai-bridge.js

# 发布前
powershell -File .\scripts\build-code-projects.ps1 -Configuration Release
powershell -File .\build\RunReleaseChecks.ps1
```

手工场景（改 Chat/Agent 必做）：

1. Word：发送普通问题 / 排版指令 / 校对 / 查看上下文面板  
2. Excel：选区分析 / 写公式 / 图表  
3. PPT：生成要点页 / 翻译  
4. 断网与错误 API Key：用户可见错误，而不是静默失败  

---

## 10. 附录：本轮量化指标

| 指标 | 数值 |
|---|---|
| VB 源文件（约） | 290 |
| 总 LOC（约） | 85,441 |
| ShareRibbon LOC（约） | 59,327 |
| `Catch ex As Exception` | ~1063 |
| 真正无条件裸 `Catch` | ~31 |
| 所有 `Catch` 变体 | ~1495 |
| `Debug.WriteLine` | ~1510 |
| `Task.Run` | ~68 |
| `Select Case` | ~223 |
| `UiDispatcher` 引用 | ~21 |
| `AiGateway` 引用文件 | 4 |
| `HttpStreamService` 引用文件 | 2 |
| `LLMUtil` 引用文件 | 7 |
| Tool JSON | 72 |
| Skills（JSON + 目录） | 5 JSON + 4 目录 Skill |
| 单元测试项目 | 0 |
| Smoke 脚本 | 9+ |
| 仓库目录内 pfx/snk | 4（真实 Git 跟踪状态需在 git repo 中复核） |

---

## 11. 附录：问题索引（便于建任务）

| ID | 级别 | 标题 | 主位置 |
|---|---|---|---|
| P0-1 | P0 | 超大入口类 | `BaseChatControl` / 三端 `ChatControl` |
| P0-2 | P0 | 多运行时并存 | `Agent/*` + Chat 路由 |
| P0-3 | P0 | 线程与同步等待 | Memory / ExcelDna / WebView |
| P0-4 | P0 | 异常与可观测性 | 全库 Catch / Debug |
| P0-5 | P0 | 安装发布链路 | `OfficeAgent.vdproj` |
| P0-6 | P1 / P0 待验证 | 签名密钥文件位于仓库目录内 | `*_TemporaryKey.pfx` / snk |
| P1-1 | P1 | LLM 双轨 | `AiGateway` / `HttpStreamService` / `LLMUtil` |
| P1-2 | P1 | Memory 双模型 | `Storage/*` / Memory 服务 |
| P1-3 | P1 | Skills 双格式 | `Skills/*` / Registry |
| P1-4 | P1 | 三端能力不均 | Word vs Excel/PPT |
| P1-5 | P1 | 测试不足 | `scripts` only |
| P1-6 | P1 | 前端遗留 | `ralph-loop.js` 等 |
| P1-7 | P1 | DB 双源与文档 | `OfficeAiDatabase` / README |
| P1-8 | P1 | 依赖栈偏重 | packages.config |
| P1-9 | P1 | 重复基础设施 | ComObjectHelper / Undo |
| P1-10 | P1 | Agent 闭环缺口 | ToolResult / Context |
| P2-* | P2 | 见第 5 章 | — |

---

## 12. 结论

当前项目 **不是从零开始的原型**，而是一个 **功能面已经较完整、架构正在收敛但未收敛完成** 的 VSTO AI 产品。最大的危险不是“缺功能”，而是：

1. **继续在旧入口上加功能**，使分叉永久化；  
2. **Agent 看起来启用了，但 observe/repair/工具边界未在所有路径生效**；  
3. **发布与线程问题以偶发形式消耗团队时间**。

建议下一阶段以 **“一条主路径 + 可验证闭环 + 入口持续瘦身”** 为唯一主题，而不是平行开发更多表面功能。  

若只选三件立刻做的事：

1. **拆分并冻结** `BaseChatControl` / `WordAi/ChatControl` 的新增业务；  
2. **清掉 UI 路径上的同步等待**，Memory/Agent 全异步；  
3. **强制 ToolResult + OfficeContext** 进入 Loop，让 Excel/PPT 按 Word 样板对齐 capability。  

---

*本报告为静态审计结果；未在本轮执行完整 Office 宿主 UI 手工回归。实施优化时请以第 9 章验证清单为准。*
