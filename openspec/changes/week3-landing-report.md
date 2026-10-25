# Week 3 落地报告与验收清单

> 日期：2026-10-24
> 范围：E1-06 replan 幂等保护、E1-08 信任层 UI 收口、E4-01 Word 选区工具条、E4-02 三端右键菜单、E3-02 每端一个一键示例。
> 结论：代码与自动化校验已通过；按用户指示本轮未做真机验收。

## 1. E1-06 replan 幂等保护

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 写入台账 | `ShareRibbon/Agent/Execution/AgentWriteLedger.vb`（新增） | `AppliedWrite` 记录“工具 + 规范化参数指纹 + 计划版本 + 步骤说明”；`AgentWriteLedger.Record` 只收宿主确认成功的写入 |
| 指纹归一 | `AgentWriteLedger.CanonicalJson` | 对象键按 `StringComparer.Ordinal` 排序后递归序列化，模型重规划时换一种键顺序不会被当成新写入 |
| 会话挂载 | `ShareRibbon/Agent/AgentSession.vb` | `session.Writes` 为本轮唯一台账，随会话生命周期存在，不入库、不跨轮 |
| 执行前拦截 | `ShareRibbon/Agent/LoopEngine.vb` | ACT 之前先 `FindAppliedInPriorPlan(toolCall)`；命中则直接产出“已在计划版本 N 完成”的成功结果，不进宿主执行、不建撤销点、不弹风险确认 |
| 计划版本推进 | `LoopEngine.vb` `Case "replan"` | 只有换成全新计划才 `BeginNextPlan()` 并把 `stepIndex` 归零；`retry`/`skip` 沿用原计划，台账不会误挡本步重试 |
| 台账回填 | `LoopEngine.vb` | `session.ChangeLog.AddFromStep` 返回的变更记录同时喂给 `session.Writes.Record`，成功写入与台账同源，不需要第二次观察 |
| 约束下发 | `LoopEngine.Planning.vb` `ReflectAndReplanAsync` | 已落盘写入清单作为系统级约束同时下发给反思与重新规划请求，两处共用同一份文本；文本带 `APPLIED_WRITES` 机器可校验标记 |
| 覆盖校验修正 | `ShareRibbon/Agent/LoopEngine.Coverage.vb`（新增，由 `LoopEngine.Planning.vb` 拆出） | `ValidateReplannedCoverage` 把上一版成功执行的工具并入覆盖判定，避免把已完成的工作误判成漏项而拒绝新计划 |
| 幂等结果形状 | `LoopEngine.Coverage.vb` `BuildAlreadyAppliedResult` | 观察 `kind=idempotence`、`changed=false`、`warnings=[duplicate_write_skipped]`，因此变更日志不会多出重复行，时间线又能看出该步已完成 |

关键约束：

- 判定“已经写入”的唯一依据是宿主返回的提交事实（`Success && Changed && !RolledBack`），不是计划步骤状态，也不是模型自述。
- 台账只在跨计划版本之间生效：同一版计划内重复出现同一步骤仍会照常执行，避免把幂等保护变成“本步不许重试”。
- 主链没有变：`OfficeHarness -> AgentKernel -> LoopEngine -> ToolRegistry -> CodeExecutionService -> 宿主 ExecuteJsonCommand`，幂等判断落在 LoopEngine 的 ACT 前闸门，未新增第二套计划器或执行器。
- 台账文本只写工具名、步骤说明与宿主摘要，不写文档原文、不写截图或 Base64。

顺带修掉一个长期存在的产品缺陷：`PromptManager.BuildReflectionPrompt` 里 `If(it.Observation, "成功", "失败")` 把字符串观察当成布尔条件，VB 运行时抛“从 String 到 Boolean 的转换无效”，反思提示每次都构造失败并被 `Catch` 吞掉——replan 分支此前从来没有真正跑到过。现在按长度收敛输出观察摘要，并把该 `Catch` 改为写脱敏 Warn 日志。

## 2. E1-08 信任层 UI 收口

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 计划信任模型 | `ShareRibbon/Agent/AgentSession.vb`（`PlanTrust`、`ExecutionPlan.Trust`） | 计划模型带上版本号、范围、数量、风险与需确认工具清单；这些事实由后端计算，模型只负责 steps/understanding/summary 文本 |
| 信任信息计算 | `LoopEngine.Coverage.vb` `BuildPlanTrust` / `BuildScopeText` | 风险取计划内工具在 `ToolRegistry` 里的 `RiskLevel` 上界并列出名称；范围取自 `lastContextPack`（选区项数或整篇文档 + 文档名），不采信模型自述的“我要改哪一段” |
| 重规划重发计划 | `LoopEngine.vb` `Case "replan"` | `BeginNextPlan()` 之后调用 `RaisePlanGenerated(session)`，第 2 版计划同样进入时间线；旧实现只在首版发一次，用户看到的始终是已被作废的计划 |
| 步骤状态令牌 | `LoopEngine.vb`（`OnStepStatusChanged` + 六个状态常量）、`AgentKernel.vb`、`OfficeHarness.vb`、`AgentKernelService.vb` | 事件由 `(index, Boolean, message)` 改为 `(index, status, message)`，`pending/running/completed/failed/skipped/aborted` 全链路透传；用户拒绝、审批放弃、执行失败在卡片上终于是三种样子 |
| 正在执行的步骤 | `LoopEngine.vb` 步骤边界 | 进入步骤时发 `running`。此前这个状态从未发出，卡片只有“待办→已完成”，迭代详情的 `.step-running` 兜底分支形同虚设 |
| 迭代归属 | `AgentKernelService.vb`（下发 `stepIndex`）、`agent-card.js` `updateAgentIteration` | 迭代详情按所属步骤落位而不是全局迭代号；重规划后两个编号必然错开，旧写法会把观察写进别的步骤甚至丢弃 |
| 计划卡渲染 | `agent-card.js` `buildPlanTrustHtml` / `showAgentPlanCard`、`styles.css` | 新增范围/数量/风险 chip；`planVersion > 1` 时复用同一容器原位重绘，并保留“终止执行”与运行中状态，不再退回“等待确认执行” |
| 终态变更摘要 | `AgentKernelService.vb` `BuildTerminalDescriptor`、`agent-card.js` `renderAgentChangeSummary` | 终态卡片列出实际留在文档里的改动（最多 8 条、单条摘要 120 字封顶、标注部分生效、另给回滚条数与“还有 N 处未列出”），数据源仍是已提交的变更日志 |
| 撤销入口与原因 | 同上 + `agent-card.js` / `styles.css` | 可撤销时按钮带撤销单元数；不出按钮时必须给出原因（本轮无改动 / 该宿主没有等价撤销单元 / 缺撤销依据或指纹 / 没有 runId），旧实现只是静默不显示 |
| 图标函数冲突 | `agent-card.js` `getAgentStepIcon` | `agent-card.js` 里的 `getStepIcon` 与 `code-handler.js` 同名且加载更晚，把执行计划卡使用的图标字典整个覆盖成步骤状态字典（`search/data/...` 全部退化成默认图标）；改名后两条链各自独立 |

关键约束：

- 卡片上的范围/数量/风险全部来自 `ToolRegistry` 与 `ContextPack`，模型文本只当标题渲染，不参与事实判定。
- 主链没有变，也没有新增第二套状态机：仍然是 `OfficeHarness -> AgentKernel -> LoopEngine -> ToolRegistry -> CodeExecutionService -> 宿主`，UI 只是把同一条链上的事实渲染出来。
- 终态摘要只取宿主观察里的短摘要并限长，写入前端一律经 `SerializeForScript`/`ToJsStringLiteral` 与 `textContent`/`escapeHtml`，不带文档原文、参数、截图或 Base64。

## 3. E4-01 Word 选区工具条

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 入口话术唯一清单 | `ShareRibbon/Agent/QuickActionCatalog.vb`（新增） | `QuickAction(Id/Label/RequestText/CapabilityId)` + `ActionsFor(appType)`；Word 三个动作（改写选区 / 生成摘要 / 校对选区），未接入的宿主返回空表而不是让调用方猜文案 |
| 能力对齐 | 同上 + `WordAi/Services/WordCapabilityRegistry.vb` | 「校对选区」声明 `CapabilityId = word.proofread`，由 smoke 校验该 id 真实存在于注册表，防止"入口说得到、能力做不到"；纯模型动作不声明能力 |
| 向导示例收口 | `ShareRibbon/Controls/Services/OnboardingService.vb` | `BuildExamples()` 删掉自带的那份 `Select Case appType` 示例表，改读 `QuickActionCatalog.SampleTasks(_getAppType())`（文案逐条未变）。选区工具条（E4-01）、向导示例（E3-01）与每端示例任务（E3-02）从此共用同一处，不再需要人工同步三张表 |
| 落点计算 | `WordAi/Selection/WordSelectionToolbarGeometry.vb`（新增，128 行） | 纯几何 + Win32 取点，不碰 Word COM、不碰窗体：鼠标锚点在窗口内则贴锚点下方居中，越出下边界翻到上方；锚点不可信（键盘选择、鼠标已移开）贴窗口顶部居中；结果一律 clamp 到锚点所在显示器工作区。工作区可显式传入的重载供隔离验证使用 |
| 浮层本体 | `WordAi/Selection/WordSelectionToolbar.vb`（新增，141 行） | 无边框 `Form`，`CreateParams` 叠 `WS_EX_NOACTIVATE(&H8000)` + `WS_EX_TOOLWINDOW(&H80)`，并覆写 `ShowWithoutActivation`：点按钮不抢焦点，因此不会把刚选中的文本清掉。只在动作签名（id 序列）变化时重建按钮 |
| 编排 | `WordAi/Selection/WordSelectionToolbarController.vb`（新增，197 行） | 订阅 `Application.WindowSelectionChange`，200ms 防抖（验收要求 300ms 内出现）；`MinSelectionChars = 4` 过滤光标停靠；同一处选区已显示则不再挪位。显示前置门：开关 → `ActiveWindow` 是文档窗口 → 面板可见 → 本轮 Agent 未在跑 |
| 宿主接线 | `WordAi/ThisAddIn.vb` | 控制器挂在 `ChatTaskPane_VisibleChanged` 的"变为可见"分支上（`EnsureSelectionToolbar`）：用户也能用任务窗格自带的折叠按钮打开面板，那条路不经过 `ShowChatTaskPaneAsync`；启动阶段面板不可见，所以不进 `WordAi_Startup` 关键路径。面板隐藏时收起浮层，`Shutdown` 里 `Dispose` 并置空。与 ghost text 各自订阅同一事件、互不调度，`Initialize` 以 `_bound` 幂等、`Dispose` 对称解绑 |
| 统一入口 | `ShareRibbon/Controls/BaseChatControl.vb` `StartAgentFromHostEntry`（新增） | 宿主原生入口只把一句话交给 `HandleStartAgent`：同一个模型配置闸门、同一次 `RefreshSessionDocumentContext`、同一条 `OfficeHarness -> AgentKernel -> LoopEngine` 主链，预览/风险确认/变更摘要/撤销全部沿用 E1-08 的计划卡。不再让宿主拼 `sendMessageToServer` 的 JS 字符串，不依赖 WebView 加载时序，也不多留一份话术副本 |
| 配置开关 | `ShareRibbon/Config/ChatSettings.vb` | `EnableSelectionToolbar` **默认 False**，`office_ai_chat_settings.json` 键 `enableSelectionToolbar`，`LoadSettings` 用 `ContainsKey` 兼容老配置；`HandleSaveSettings` 回填当前值，面板保存不会把开关悄悄重置 |

关键约束：

- 默认关闭是硬要求：浮层画在用户正在写的文档上，会碰到每一次选区操作。本轮只给配置开关，面板设置 UI 之后再补，因此用户能在不重新编译的前提下关掉它。
- 命名空间不能叫 `Agent.Entry`。VB 标识符大小写不敏感，该命名空间会劫持 `ShareRibbon/Agent/Execution/AgentChangeLog.vb` 里 `For Each entry In _entries` 的循环变量，整解决方案编不过（BC30112/BC30456）。清单落在 `Namespace Agent` 下。
- 窗口有效性用 `_app.ActiveWindow` 判定并要求 `wdWindowDocument`，不用文档计数：WordAi 是 Application 级加载项，"有文档但没有活动文档窗口"（开始页、邮件正文、WPS 其他窗口）是真实存在的状态。
- 隐私：`AppLogger` 只记动作 id 与数量，选区原文永不进日志；文本只在交给主链后由上下文采集读取。
- 没有新增第二套 Harness/Planner/Loop/Safety/执行器，也没有把入口动作登记进 `ToolRegistry`：入口只是"说话"，能力仍由主链上的 capability 承担。

## 4. E4-02 三端右键菜单

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 共享登记器 | `ShareRibbon/Agent/HostContextMenu.vb`（新增，179 行） | 只操作 OCOM 层的 `Office.CommandBar`/`CommandBarButton`：加项、按 `Tag` 前缀 `aiq:` 认领自己挂上去的项、倒序按 Tag 删除、`Temporary:=True` 不落进用户配置。共享层不引用任何 Word/Excel/PPT 互操作类型 |
| 宿主适配器 | `WordAi/HostEntry/WordContextMenu.vb`、`ExcelAi/HostEntry/ExcelContextMenu.vb`、`PowerPointAi/HostEntry/PowerPointContextMenu.vb`（各新增） | 每端只回答两个宿主问题："挂哪条菜栏"（按候选名试探，取不到就不出这个入口，不报错不重试）与"此刻的上下文够不够"。动作一律来自 `QuickActionCatalog.ActionsFor(...)`，不各留一张表 |
| 动作清单扩展 | `ShareRibbon/Agent/QuickActionCatalog.vb` | Excel 三项（AI 解读选区 / AI 生成公式 / AI 整理格式）、PowerPoint 三项（AI 改写选中文字 / AI 提炼本页要点 / AI 校对本页）。**Excel/PPT 一律不声明 `CapabilityId`**：这两个宿主还没有 capability 注册表（开放待办 #36），宁可少声称，也不要"入口说得到、能力做不到" |
| 降级方式 | 同共享登记器 + 三个 `CurrentAvailability()` | Word 要求 `Selection.Type <> wdSelectionIP`；Excel 要求 `Selection` 能 `TryCast` 成 `Range`（右键图表/形状时不算）；PowerPoint 要求 `ActiveWindow.Selection.Type` 落在 shapes/text/slides。不满足时点在菜单项上的回答是状态栏一句原因（`GlobalStatusStripAll.ShowProgress`），而不是硬跑一轮 Agent |
| 统一入口 | 三个 `ThisAddIn.vb` 的 `OnContextMenuAction` | 打开面板 → `chatControl.StartAgentFromHostEntry(action.RequestText)`，闸门、上下文、预览、风险确认、变更摘要、撤销全部沿用 E1-08 那条链。`EnsureContextMenu()` 只在 `ChatTaskPane_VisibleChanged` 的"变为可见"分支里挂，`ThisAddIn_Shutdown` 里 `Detach()` |
| Ribbon 收口 | `WordAi/Ribbon1.vb:471`、`ExcelAi/Ribbon1.vb:240`、`PowerPointAi/Ribbon1.vb:427` | `StartAgentFromRibbonAsync` 原来自己拼 `sendMessageToServer({type:'startAgent',...})` 注入 JS；三处改成 `chatCtrl.StartAgentFromHostEntry(request)`。`router.Register("startAgent", ...)` 本来就直通 `HandleStartAgent`，因此行为等价，但少了一份注入式载荷和一处 WebView 时序依赖 |
| 配置开关 | `ShareRibbon/Config/ChatSettings.vb` | `EnableContextMenu` **默认 False**，json 键 `enableContextMenu`，`ContainsKey` 兼容老配置，`HandleSaveSettings` 回填当前值 |

关键约束：

- 默认关闭的理由和 E4-01 同类且更强：它改的是用户宿主里**本来就存在**的右键菜单。没有面板设置开关之前只能是 opt-in。
- 命名空间一律 `HostEntry`，不用 `Selection`。VB 会把命名空间名和互操作类型名混在一起解析：`Namespace Selection` 让 `PowerPointAi` 里 `sel As Selection` 报 BC30182（应为类型），和之前 `Agent.Entry` 劫持 `For Each entry` 是同一类坑，且比逻辑错更难查。
- **不做置灰**：`packages/MicrosoftOfficeCore.15.0.0` 这份 PIA 的 `CommandBar` 没有任何事件面（反射确认只有 `_IMsoOleAccDispObj`/`IAccessible`，`CommandBarEvents` 在这份程序集里根本不存在），`BeforeRightClick` 编不过。可行的替代就是点击时探测；退一步说，即使能置灰，Attach 时算一次、之后不刷新的静止灰色也是骗人。
- 可撤销性：本加载项只按 `Tag` 前缀认领和删除自己加的项，删除按倒序走 COM 集合（正序枚举会错位、留下洗不掉的死项）；按钮 RCW 存在 `_buttons` 里，否则 GC 掉 RCW 后 Click 不再回调。
- 隐私：日志只记 `quick action id=...`；菜单项承载的请求原文、`Caption`、选区文本一律不进 `AppLogger`。
- 没有为右键菜单新建第二套 planner/loop/safety/executor，入口动作也没进 `ToolRegistry`。

已知不确定项（留给真机）：Word 的 `Text`、Excel 的 `Cell` 菜栏名是稳定约定；PowerPoint 各版本暴露的上下文菜栏更零散，`{"Text","Drawing","Object"}` 只是候选试探，真机需要按版本记录实际命中哪条、以及该菜栏是否真的会显示这些项。

## 5. E3-02 每端一个一键示例

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 示例数据形态 | `ShareRibbon/Agent/QuickActionCatalog.vb` | 每端示例从 `String()` 改为 `SampleTask(Text, RunRequestText?)`：`Text` 是面板文案，带 `RunRequestText` 的那条才是一键示例。**仍然是同一个清单**，没有新开一张表，也没有第二处 `Select Case appType` |
| 每端一键示例 | 同上 | Word=改写一段话（有选区改选区，没选区改第一段）、Excel=加一列同比增长率、PPT=提炼当前页要点。三端第一条各不同，smoke 断言"三条请求文本互不相同"，防止一端的话术被复制粘贴到另两端 |
| 空文档降级 | 同上（写在请求文本里） | 每条请求都自带"如果文档还是空的，先用一段内置示例（示例文字 / 示例数据表 / 内置主题一页）再来一次"。入口不猜用户有没有选区，也不为了这一判断去调 COM：主链本来就能读上下文，让它按观察结果决定 |
| 载荷 | `ShareRibbon/Controls/Services/OnboardingService.vb` | `examples` 由字符串数组改为 `[{text, request?}]`；`request` 只在 catalog 声明可跑时才出现，面板拿到的必须是 catalog 原文（smoke 逐条比对反射结果） |
| 面板行为 | `ShareRibbon/Resources/js/onboarding.js` | 带 `request` 的条目渲染成实心蓝按钮（`.onboarding-demo`），点击后只发 `startAgent` + `onboardingFinish` 并收起卡片；其余条目仍是"填进输入框"。`onboardingRunSample` 里没有配置判断、没有执行分支——模型没配好时由 E3-03 的发送闸门统一回落到配置 CTA，向导里不养第二份配置真值 |
| 样式 | `ShareRibbon/Resources/css/styles.css` | 新增 `.onboarding-demo`（及其 hover）。CSS→JS 类名守卫（原有）保证这条规则不会变成死样式 |

关键约束：

- 一键示例是**入口**，不是执行器：它只把 catalog 里那句话交给 `startAgent`，因此和面板输入共用同一个配置闸门、同一次 `RefreshSessionDocumentContext`、同一条 `OfficeHarness -> AgentKernel -> LoopEngine` 主链，预览/风险确认/计划卡/变更摘要/撤销全部沿用 E1-08。
- 没有为示例新增宿主端消息处理器。`BaseChatControl` 若出现 `Register("runSample")` 之类的第二个入口，P0 守卫直接失败。
- 未接入的宿主（如 Outlook）仍回落到 Word 清单：向导不能因为 appType 认不出就空着手。
- 前端资源改完必须重启宿主刷新 `%LOCALAPPDATA%\OfficeAI\www` 副本，否则面板跑的是旧 JS。

## 6. 自动化校验

- `scripts/smoke-replan-idempotence.ps1`（新增，已登记进 `scripts/run-code-checks.ps1`）：
  - Part A 直接反射 `AgentWriteLedger` 的合同：空台账不下发约束；拒绝 Nothing/失败/回滚记录；同版本不算“上一版已写”；跨版本命中键顺序不同的同一写入；不同工具或不同载荷不误命中。
  - Part B 反射驱动 `LoopEngine.RunAsync`，桩宿主按标记返回：`SMOKE_W1/W2` 成功写入、`SMOKE_FAIL_1..3` 抛 `COM_ERROR`、`SMOKE_W3` 是新计划才有的工作。断言：确实进入反思与 replan；`SMOKE_W1` 只到达宿主 1 次；W2/W3 各 1 次；失败注入 3 次；计划版本推进到 2；台账跳过标记恰好 1 条；变更日志 3 行且 3 行真实写入（无重复）；反思与重新规划两个请求的系统提示都带上了 `APPLIED_WRITES`。
- `scripts/audit-p0-guardrails.ps1` 新增守卫：ACT 前必须查台账、跳过分支不得建撤销点、replan 必须推进计划版本并做并入已完成项的覆盖校验、台账不得记录未提交写入、反思与重新规划不得各自复制一份约束文本、`PromptManager` 不得再对字符串观察用三参数 `If`。
- `scripts/audit-p0-guardrails.ps1` 新增 `.vbproj` 注册守卫：四个代码项目下每个 `.vb`（`bin/obj` 除外）都必须出现在本项目 `<Compile Include>` 中，直接拦住“新文件写了却不参与编译”。
- 因新守卫按 800 行判定 oversized，把 `LoopEngine.Planning.vb`（806 行）的计划/执行覆盖校验职责拆到 `LoopEngine.Coverage.vb`（217 行），Planning 收敛到 604 行；未走“加白名单”的路子。
- `scripts/check-agent-card-ui.js`（新增，已登记进 `scripts/run-code-checks.ps1`）：用最小 DOM 桩在 node 里加载 `agent-card.js`，直接跑 UI 契约——断言计划卡渲染出范围/数量/风险 chip、首卡有“开始执行”而重规划卡保持“终止执行”、重规划是原位重绘而不是追加第二张卡、旧计划的步骤节点被销毁、`skipped` 不残留失败样式、`aborted` 有独立图标、迭代详情按 `stepIndex` 落位、终态摘要列出已写入条目并标注部分生效与回滚、宿主摘要里的 HTML 片段只当纯文本、不可撤销时给出原因、可撤销时按钮带撤销单元数且点击后经 `agent:undoRun` 回写结果。
- `scripts/audit-p0-guardrails.ps1` 新增 E1-08 守卫段：步骤状态不得退回布尔、`RaisePlanGenerated` 必须至少两处调用（首版 + replan）、计划风险必须读 `ToolRegistry.GetTool`、范围必须读 `lastContextPack`、计划卡桥接载荷必须含 `scope/planVersion/riskLevel`、迭代载荷必须含 `stepIndex`、终态摘要必须由 `changeLog.Entries` 生成且受 `MaxTerminalChanges` 限制、`undoReason` 赋值不得缺失、卡片必须渲染 change summary 与不可撤销原因、`agent-card.js` 不得再声明 `getStepIcon`。

- `scripts/smoke-selection-toolbar.ps1`（新增，已登记进 `scripts/run-code-checks.ps1`）：
  - 结构合同：清单在 `Namespace Agent` 下且禁止回到 `Agent.Entry`；两个 `.vbproj` 的 `<Compile Include>` 覆盖四个新文件；Word 恰好 3 个动作、id 必须是 `word.selection.*`；每个 `CapabilityId` 都要能在 `WordCapabilityRegistry` 里找到同名 id；向导文案不得在 catalog 之外再出现一份；`DebounceMs` 必须落在 `(0, 280]`；`Initialize` 幂等 + `AddHandler/RemoveHandler` 对称；`ThisAddIn` 懒创建 + Shutdown 释放，且 `WordAi_Startup` 体内不得出现 `EnsureSelectionToolbar`；必须用 `ActiveWindow`/`wdWindowDocument`，出现 `Documents.Count` 直接失败；入口只能走 `StartAgentFromHostEntry -> HandleStartAgent`，出现 `sendMessageToServer/createChatSection/ExecuteJavaScript` 视为另开一条链；`AppLogger` 调用行里不得出现 `.Text` 或 `RequestText`；`WS_EX_NOACTIVATE` + `ShowWithoutActivation` + `CreateParams` 同一条语句叠两个扩展样式；开关默认 False、json 键存在、`LoadSettings` 有 `ContainsKey` 兜底、面板保存必须回填当前值。
  - 行为验证：把 `WordSelectionToolbarGeometry.vb` 现场编译（`Add-Type`，只依赖 System.Drawing/Windows.Forms，不需要 Word），用固定的工作区/窗口矩形断言 —— 条带尺寸非空、`ItemWidth` 容得下四个汉字、锚点贴右边缘时被拉回工作区、贴下边缘时翻到上方且不越界、锚点不可信时回落到工作区水平居中且且在正常落点之上。
- `scripts/audit-p0-guardrails.ps1` 新增 E4-01 守卫段：`Agent.Entry` 命名空间回归、宿主拼 JS 绕过面板、未经 `StartAgentFromHostEntry` 回主链、浮层丢掉不激活样式、日志里出现选区文本、开关改回默认开启、键不再持久化、Startup 关键路径重新创建控制器、smoke 未登记进 `run-code-checks.ps1`，以及向导不再读 `QuickActionCatalog.SampleTasks`（守卫文本全部 ASCII，中文串在 PS 5.1 下不能当锚点）。
- `scripts/smoke-host-context-menu.ps1`（新增，已登记进 `scripts/run-code-checks.ps1`）：
  - 行为验证：`Add-Type` 载入 `ShareRibbon\bin\Debug\ShareRibbon.dll` 后**直接调用** `QuickActionCatalog.ActionsFor('Word'/'Excel'/'PowerPoint')`，断言每端恰好 3 项、id/Label/RequestText 都非空、id 前端（`word.`/`excel.`/`ppt.`）、无重复 id，未接入的 `Outlook` 返回空表；同时断言 Excel/PPT 的动作**不得**声明 `CapabilityId`（那两个宿主没有注册表可证明它）。
  - 结构合同（读取源码前先剥掉 `'` 注释行，否则"禁止出现"的守卫会被解释性注释自己触发）：共享层必须在 `Namespace Agent` 下且不引用任何 `Microsoft.Office.Interop.*`；必须保留 `aiq:` Tag 前缀、`Temporary:=True`、倒序删除、按 Tag 删除、`_buttons` 持有 RCW、`RemoveHandler button.Click`；出现 `BeforeRightClick`/`CommandBarEvents` 直接失败（那份 PIA 没有这个面）；降级必须经 `GlobalStatusStripAll.ShowProgress`。三端适配器必须用 `Namespace HostEntry`（出现 `Namespace Selection` 失败）、必须 `New HostContextMenu()`、必须从 catalog 取动作并把 `AddressOf CurrentAvailability` 交给共享层、必须有带 Reason 的降级、不得自己 `AddHandler`、不得拼 JS。三端 `ThisAddIn` 必须过 `ChatSettings.EnableContextMenu` 门、必须经 `StartAgentFromHostEntry`、Shutdown 必须 `Detach()`、`EnsureContextMenu()` 必须出现在 `ChatTaskPane_VisibleChanged` 体内且**不在** `*_Startup` 体内；三端 `Ribbon1` 的 `StartAgentFromRibbonAsync` 不得再注入 JS。另加 500 行上限检查（项目规则是每类 ≤500 行）。
- `scripts/audit-p0-guardrails.ps1` 新增 E4-02 守卫段：共享层伸手进宿主对象模型、丢失 Tag 前缀、试图挂 CommandBar 事件、适配器回到 `Namespace Selection` 或不再委托共享层、入口层日志出现 `.Text`/`RequestText`/`.Caption`、任一宿主绕过开关或绕过 `StartAgentFromHostEntry`、Shutdown 不再 `Detach`、新文件没进 `.vbproj`、开关改回默认开启或不再持久化、smoke 未登记。
- 顺手修掉 E4-01 smoke 里两处因清单扩容而失真的断言：动作 id 计数改为只看 `word.*`（catalog 现在含 Excel/PPT 动作），`SaveSettings(...)` 回填断言允许 `EnableSelectionToolbar` 后再跟参数。
- `scripts/smoke-onboarding-wizard.ps1` 扩了 E3-02 段（同一份向导合同，不再新开脚本）：反射调用 `QuickActionCatalog.SampleTasks('word'/'excel'/'powerpoint')`，与面板实际收到的 `renderOnboarding` 载荷**逐条比对** `text`，可跑那条还必须逐字比对 `request`——这条断言就是"向导不再自己拥有一份话术"的证据，而不是靠注释里的一句承诺。另外断言：每端恰好 1 条可跑示例且必须排在第一条；三端请求文本互不相同；每条可跑请求都同时带"预览"“空文档分支”“内置示例”三个语义锚点且长度≥40 字；未接入的 `outlook` 仍返回一份成体系的清单。锚点字符用 `[char]0xXXXX` 现场拼装（整个脚本必须保持 ASCII，PS 5.1 会把无 BOM 脚本按代码页读）。
- `scripts/audit-p0-guardrails.ps1` 新增 E3-02 守卫段：`SampleTask` 丢掉 `RunRequestText`、`OnboardingService` 不再下发 `request`（一键示例会静默退回"只填输入框"）、`onboarding.js` 不再把示例发到 `startAgent`、`onboardingRunSample` 里出现 `needsSetup`/`sendChatMessage`/`startAgentExecution`/`executeCode`（入口层自己判断配置或自己开执行路径）、`BaseChatControl` 里冒出 `Register("runSample")` 一类第二个示例入口，任一命中即失败。已用变异探针验证这些守卫不是空跑：删掉 `item("request")` 那行 + 把 JS 的 `startAgent` 改成自定义消息类型，审计立刻报 3 条（含既有的"前端发的消息类型必须有宿主注册"守卫）。

门禁执行结果：`cmd //c build-code.bat` 四个代码项目全部 PASS；`powershell -File scripts/run-code-checks.ps1 -SkipBuild` 全部通过（含 P0 守卫、smoke 套件、Golden L0 12 例、JS 语法检查与 `check-agent-card-ui.js`）。

## 7. 真机验收清单（待执行）

| # | 验收动作 | 通过标准 |
| --- | --- | --- |
| 1 | Word 里让 Agent 一次写入多处，中途制造一次失败（例如临时禁用某步依赖的选区） | 时间线显示“已重新规划（共 N 步），已写入 M 处不会被重复执行” |
| 2 | replan 后检查文档 | 上一版已写成的内容没有第二份，新计划补上的工作正常出现 |
| 3 | 变更摘要卡 | 已写入处数与实际一致，不因重规划翻倍 |
| 4 | 本轮撤销 | 撤销一次即回到执行前状态，无残留重复段落 |
| 5 | `%LOCALAPPDATA%\OfficeAiAppData\logs\office-ai-YYYYMMDD.log` | 出现 `Duplicate write skipped after replan toolId=...`，且 Reflect/replan 不再出现转换异常 |
| 6 | 发送一个明确带选区的执行请求（例如"把选中的段落改成摘要"） | 计划卡"范围"chip 显示"当前选区（N 项）"，N 与实际选中项数一致，而不是模型自述 |
| 7 | 计划里混入一个高风险工具（例如删除/覆盖类） | 风险 chip 显示"需确认 1 项"并列出工具名；点开始执行后确实弹出风险确认，而不是直接执行 |
| 8 | 让计划执行到第 2 步时观察时间线 | 当前步显示进行中样式（`.step-running`），已完成步为绿、失败步为红，不会全是待执行 |
| 9 | 在审批弹窗点"拒绝"终止本轮 | 该步显示"已跳过"图标而非红色失败；被取消令牌终止的步显示终止图标 |
| 10 | replan 触发后观察计划卡 | 同一张卡更新为"第 2 版"，按钮区只剩"终止执行"，状态栏显示"按新计划继续执行…"，不出现第二张计划卡 |
| 11 | Word 本轮执行成功结束后查看变更摘要 | 摘要逐条列出实际写入处（超过 8 条显示"另有 N 处未列出"），撤销按钮文案带撤销单元数；点击后显示撤销结果且文档回到执行前 |
| 12 | 在 Excel/PPT 或本轮无改动时结束 | 不出现撤销按钮，改为显示不可撤销原因（如"当前宿主暂不支持一轮撤销"），而不是按钮消失让用户猜 |
| 13 | 关掉 Word，编辑 `我的文档\OfficeAiAppData\office_ai_chat_settings.json` 加 `"enableSelectionToolbar": true`，重启 Word 并打开 AI 面板 | 不开面板时选任何东西都不出浮层；面板打开后选中一句 4 字以上的文字，300ms 内在选区末端下方出现三个按钮（改写选区 / 生成摘要 / 校对选区） |
| 14 | 鼠标点一下"改写选区" | 面板没有失去焦点、选区仍在；时间线出现计划卡，"范围"chip 是当前选区，走的是与面板输入完全相同的一条链（预览、风险确认、变更摘要、撤销都在） |
| 15 | 连续拖动框选、小幅移动鼠标、键盘 Shift+方向键扩选 | 浮层不跟着每个中间态反复重画；同一处选区不跳动；换一处选区后浮层出现在新位置 |
| 16 | 在文档最右侧、最底部一行选字；再在副屏上重复 | 条带完整可见，不被任务栏或屏幕边缘切掉，也不会跑到另一个显示器 |
| 17 | 收起面板 / 让 Agent 跑起来后再选字 | 浮层立即收起；本轮 Agent 在跑时不再叠加浮层，不打断计划卡与风险审批 |
| 18 | 关掉开关（删除 json 键或置 false）重启 Word | 全程不再出现浮层，其余功能无变化；`%LOCALAPPDATA%\OfficeAiAppData\logs` 里只有 `quick action id=...`，没有选区原文 |
| 19 | Word 里编辑 `我的文档\OfficeAiAppData\office_ai_chat_settings.json`，把 `"enableSelectionToolbar": true` 换成再加一行 `"enableContextMenu": true`，重启 Word 并打开 AI 面板 | 选中文本后右键，菜单底部出现一条分组、三项 AI 动作（改写选区 / 生成摘要 / 校对选区）；不开面板时右键也应出现这些项，但点了要有承接方（面板会自动打开） |
| 20 | 在 Excel 选中一片单元格区域后右键 | 出现 AI 解读选区 / AI 生成公式 / AI 整理格式；把"AI 生成公式"点开，走的是与面板输入同一条链（计划卡、预览、风险确认、变更摘要都在） |
| 21 | 在 Excel 右键一个图表/图片（`Selection` 不是 `Range`）并点任一 AI 动作 | 不发 Agent 请求，状态栏出现"这些 AI 动作作用于单元格区域，请先选中单元格"；菜单项本身不置灰（该 PIA 没有右键事件面，静止的灰色会骗人） |
| 22 | 在 Word 光标停靠（不选任何字）时点任一 AI 动作 | 状态栏提示"先选中一段文字，再用这里的 AI 动作"，不会空跑一轮 Agent |
| 23 | PowerPoint 各版本记录实际命中哪条菜栏 | 打开面板后右键文本框/形状/幻灯片缩略图，确认 `{"Text","Drawing","Object"}` 候选里哪条真的承载了三项；若一条都取不到，日志里应有 `PowerPointContextMenu: context menu bar unavailable`，且右键菜单没有残留死项 |
| 24 | 关掉开关（`enableContextMenu` 置 false 或删除该键）重启宿主 | AI 动作不再出现在右键菜单，且**上一轮挂的项被摘干净**（`Detach` 按 `aiq:` Tag 倒序删除）；其余菜单项顺序、语言不受影响 |
| 25 | 三个 Ribbon 上原先进入 Agent 的按钮（模板排版等） | 行为与改动前一致：面板打开、Agent 起跑、日志里不再有注入式 JS 载荷；WPS 上同样可用 |
| 26 | 三端各打开一次向导第 2 步（未配置模型时也打开一次） | 每条清单的第一项是实心蓝的"一键示例"，文案与宿主场景对得上；模型没配好时点它只回落到配置 CTA，不会留下半截思考气泡 |
| 27 | 在**已有内容**的文档/工作簿/演示文稿里点一键示例 | 出现计划卡（范围 chip 是当前选区或第一段/当前区域），预览→确认→写入→变更摘要/撤销链路齐全，且实际只改了一处 |
| 28 | 新建一个空文档 / 空工作簿 / 空演示文稿后点一键示例 | Agent 先补一段内置示例（示例段落 / 月份+销售额两列 / 示例一页）再完成任务，不是回一句"没有可用内容" |
| 29 | 点第二、第三条示例 | 只把话术填进输入框，不自动发送；一键运行只属于第一条 |

## 8. 遗留与下一步

- W2 真机验收（E1-05、E2-02/03、E3-01/03）仍未执行；E1-06、E1-08、E4-01、E4-02、E3-02 真机验收同样待补（清单见上一节 1–29 行）。
- W3 排期（E1-06 → E1-08 → E4-01 → E4-02 → E3-02）代码与自动化校验已全部落地；下一步按排期进 W4：E1-07 Excel/PPT 撤销、E4-03 快捷键、E4-04/05 迁移与闸门、E2-04/05、E5-04/05。E4-03 复用同一份 `QuickActionCatalog` 与同一套 `HostEntry` 接线。
- E3-02 复用 `QuickActionCatalog.SampleTasks`，只是把清单从字符串升级为 `SampleTask`（文案 + 可选的可跑请求），没有新开第二张表。空文档兜底写在请求话术里、交给能读上下文的主链判断；如果真机发现某端在空文档下仍然卡住，应补的是那一端的 capability，而不是在入口层加 `if 选区为空`。
- 一键示例目前只在向导里出现。若真机验证有效，下一个候选落点是面板空态（无会话时）复用同一份 `SampleTasks`，仍不新开清单。
- E4-01/E4-02 的开关目前只能改 json；面板设置 UI 补上开关后，才可以讨论把默认值改成开启。
- E4-02 的两个已知待补项：① Excel/PPT 的入口动作没有 `CapabilityId` 可声明，等开放待办 #36（宿主 capability 注册表）落地后回填并让 smoke 继续校验"声明即存在"；② PowerPoint 上下文菜栏名按候选试探，真机需要按版本确认命中的是哪条，取不到时该宿主入口自然不出，不影响 Word/Excel。
- 三端 `Ribbon1.StartAgentFromRibbonAsync` 已从注入 JS 改为 `StartAgentFromHostEntry`，这是"宿主原生入口只说话"这条规则的收口，不是新功能；后续任何新入口都不得再拼 `sendMessageToServer`。
- E17 只读分诊的前置依赖 E1-06 已具备，可在 W3 之后启动。
