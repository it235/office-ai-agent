# SHARE RIBBON SHARED CORE

## OVERVIEW

`ShareRibbon/` 是 Excel、Word、PowerPoint 插件共享的 AI Native 核心库，负责 Harness/Agent/Loop 合同、工具注册、AI 请求、共享 UI、配置、Skill 与 MCP 基础设施。它不得承载任何宿主专属 COM 实现。

## STRUCTURE

```text
ShareRibbon/
├── Agent/                 # AgentKernel、LoopEngine、ToolRegistry、Harness、Safety、共享合同
├── Config/                # API、模型、运行时配置
├── Controls/              # WebView2 Chat UI 与共享服务
├── Mcp/                   # Model Context Protocol client
├── Services/Ai/           # AiGateway 与 Provider 请求适配
├── Skills/                # 目录型 Skill 与兼容 Skill 定义
├── Tools/                 # 工具 schema
└── Resources/             # HTML/JS/CSS 与嵌入资源
```

## WHERE TO LOOK

| Task | Location | Notes |
|---|---|---|
| Agent 主链 | `Agent/AgentKernel.vb`, `Agent/LoopEngine.vb` | 规划、执行、观察、修复和终态 |
| 宿主原生入口话术 | `Agent/QuickActionCatalog.vb` | 选区浮层/右键菜单/快捷键与命令面板、引导示例（含每端一键示例 `SampleTask`）的唯一清单；`CapabilityId` 必须能在宿主注册表查到真实 id |
| 右键菜单共享登记器 | `Agent/HostContextMenu.vb` | 只用 OCOM 的 `Office.CommandBar`；按 `aiq:` Tag 认领与倒序摘除自己挂的项 |
| 快捷键轮询与组合键合同 | `Controls/Services/HostHotkeyWatcher.vb` | `HostHotkeyChord.Parse`/`IsSafe` + WinForms Timer 轮询 `GetAsyncKeyState`；只在宿主处于前台时触发 |
| 命令面板状态服务 | `Controls/Services/CommandPaletteService.vb` | 由 `BaseChatControl` 惰性持有；把 `QuickActionCatalog` 推给面板并注册快捷键 |
| ToolResult / Tool 注册 | `Agent/ToolRegistry.vb` | 工具合同、临时视觉证据合同、执行适配 |
| Harness | `Agent/Harness/` | 唯一运行入口、上下文和 RunTrace |
| AI 请求与 Provider 适配 | `Services/Ai/AiGateway.vb` | OpenAI-compatible / Anthropic 请求转换 |
| 数据边界同意与阻断卡片 | `Controls/Services/PrivacyConsentGate.vb`, `Controls/Services/ModelConfigGate.vb` | 同意只此一处读写、三条边界文案只此一份下发；缺配置与缺同意共用同一张 CTA，靠 `action` 区分；「哪个出口由哪把门看着」的唯一清单是 `scripts/smoke-privacy-consent.ps1` 的 `$outboundInventory` |
| Chat UI 与请求编排 | `Controls/` | `BaseChatControl`、`AgentKernelService`、streaming 服务 |
| 配置 | `Config/` | API Key、模型、Prompt 与运行时设置 |
| 宿主归属口径（`app_type`） | `Storage/AppTypeContract.vb` | 六张表共用同一列：`Word`/`Excel`/`PowerPoint` 或空串（通用，三端可见）；写入前必须 `Normalize` |
| 历史列表契约（顺序与分组键） | `Controls/Services/HistorySessionService.vb`, `Resources/js/history-manager.js` | 排序只在后端 `GetRecentSessions`；payload 里的 `lastActiveAt`（排序键）与 `documentId`（路径哈希）是前端分组唯二可用的键 |
| Skill | `Skills/` | 召回元数据、allowed-tools、references/scripts/assets |
| MCP | `Mcp/` | MCP client、连接与协议模型 |
| WebView2 资源 | `Resources/` | HTML/JS/CSS；修改后检查项目注册和引用 |

## OWNERSHIP BOUNDARIES

- `ShareRibbon` 定义宿主无关合同，例如 `ToolResult`、视觉证据、Office object ref、operation batch、context、capability 和 loop 状态。
- PowerPoint、Word、Excel 的 COM resolver、executor、observer、renderer 必须放在各自宿主项目。
- 共享能力沿 `OfficeHarness -> AgentKernel -> LoopEngine -> ToolRegistry -> CodeExecutionService -> Host` 接入，不得建立平行 Harness/Loop/Safety。
- AI 请求统一经过共享请求层；宿主 Executor 不得自行创建模型客户端或发起 Repair 请求。
- 宿主原生入口的合同放这里（`Agent/QuickActionCatalog.vb`、`Agent/HostContextMenu.vb`），宿主专属的
  菜栏名探测与选区判断放 `<Host>/HostEntry/`。共享层只允许 `Microsoft.Office.Core` 的 `CommandBar` 层，
  出现 `Microsoft.Office.Interop.(Word|Excel|PowerPoint)` 即越界（smoke 与 P0 守卫都会拦）。
- 入口动作不是能力：不进 `ToolRegistry`、不新建 planner/loop/safety，只把动作本身交回
  `BaseChatControl.StartAgentFromHostEntry(action)`；形参是 `Agent.QuickAction` 而不是 `String`，
  因为 E6-05 的任务包断言靠入口 id 绑定，话术和 id 必须同源，漏掉 id 要编译不过。
- 命令面板（E4-03）也是同一个入口：快捷键监听与面板状态都住在 `ShareRibbon`，由 `BaseChatControl` 惰性创建
  `CommandPaletteService`，三个宿主各自只贡献一个 `EnsureHostTaskPaneVisible` 覆写。共享层仍不得出现
  `Microsoft.Office.Interop.(Word|Excel|PowerPoint)`，出现即越界。
- 快捷键安全是硬约束：组合键必须带 `Ctrl` 或 `Alt` 且只有一个主键（裸键会吃掉用户正常输入，Word 本身已占用
  `Ctrl+K`），修饰键按**精确集合**比对；只在宿主处于前台时触发（多宿主同时打开时不能互相抢键）；监听随面板
  销毁而停止；无效组合键在落盘前就拒绝，不留一个「按了没反应」的开关。
- 一键示例（E3-02）同样是入口：面板只发 `startAgent`，配置判断与执行分支一律留给主链；「当前文档有内容」
  与「文档还空着」两种情况写在请求话术里，不写一句「如果没选区就…」的 COM 探测。
- 选区自动引用（E4-04）默认开启，但「带不带选区」只允许判一次：`Controls/BaseChatControl.vb` 的
  `BuildMessageWithAutoSelection`。宿主 override 只回答「怎么写成提示」。默认开启的前提是面板里有真的能关掉的
  开关，加上带出去的内容要在对话里显示；引用摘要只从**实际附加上的增量**反推，不再采一次上下文，也不写进
  历史、日志、Observation 或 Artifacts。
- 事件与生命周期卫生（E4-05）：一个处理器只允许一份绑定，页面注入的粒度是「一份文档一次初始化」，读开关之前
  先 `ChatSettings.EnsureLoaded(...)`。细则与原因写在 `Controls/AGENTS.md` 的 EVENT AND LIFETIME HYGIENE，
  由 `scripts/smoke-event-binding.ps1` 钉住——重复执行在这里的症状是静默失效或值被反复洗回，不是异常。
- 宿主归属（E2-04）只有一个口径持有者：`Storage/AppTypeContract.vb`。读侧（历史与记忆）一律是
  「等于当前宿主 **或** 空串」，所以白名单之外的第三种值既不等于任何宿主、又拿不到空串豁免，那条数据会在
  三端同时消失且不报错（早期构建落库的 `"Unknown"` 就是这个形状，本机 `atomic_memory` 至今还有 123 行）。
  因此**写这六列之前必须 `AppTypeContract.Normalize`**（`conversation_event`、`session_summary`、`agent_run`、
  `agent_run_change`、`atomic_memory`、`memory_item`），读侧过滤参数不要归一——那里 `''` 本来就是豁免。
  历史数据的收口在迁移 019：按证据强度只填空、不覆盖，可重复执行；新增带 `app_type` 的表或写入点时，
  `scripts/audit-p0-guardrails.ps1` 的 E2-04 段会要求两处同步登记。
- 历史列表（E2-05）的顺序只有一个持有者：后端 `GetRecentSessions` 的 `ORDER BY`（当前文档整块最前，其余按
  最近活跃）。前端的分组只是呈现——按「首次出现」切桶、桶内保持后端顺序，所以在 JS 里重排会把「当前文档最前」
  这条后端不变式直接冲掉（P0 守卫见 history-manager.js 中出现 `sort(` 即失败）。分组需要的两个键必须由后端
  给：`lastActiveAt` = `MAX(created_at)` 就是排序键本身，用 `createdAt` 分时间桶会让分桶与顺序自相矛盾；
  `documentId` 是文档绝对路径的小写哈希，只用于归组、不含路径，文件名会撞车不能当键（未解析的老会话两键都为空，
  空值继续按「通用会话」展示，不能在被分组时丢掉）。
- 本项目的命名空间绝不能用 `Entry`：VB 标识符大小写不敏感，`Namespace Agent.Entry` 会劫持
  `Agent/Execution/AgentChangeLog.vb` 里的 `For Each entry`（BC30112/BC30456）。同理不得用与互操作类型
  同名的命名空间（`Selection`、`ContextMenu` 等），宿主项目里它会把 `sel As Selection` 压成 BC30182。

## VISUAL EVIDENCE RULES

- `AgentVisualEvidence` 只用于当前 Observe/Repair 轮次，必须保持宿主无关。
- 截图的 `DataUrl`/Base64 必须是内存临时载荷，并从普通 JSON 序列化中排除。
- 禁止把视觉载荷复制到 `Observation`、`Data`、`Artifacts`、History、Memory、RunTrace、日志或 UI。
- `LoopEngine` 负责限制证据数量和大小、构造多模态消息以及文本降级；Provider 格式转换属于 `AiGateway`。
- 图片通道失败后必须有限次降级，不能在同一 Run 内反复发送必然失败的多模态请求。

## CONVENTIONS

- 共享服务优先使用依赖注入或显式委托，避免从 UI 入口反向引用宿主实现。
- 新增 `.vb`、资源或 Skill 文件后同步检查 `ShareRibbon.vbproj`。
- 日志必须脱敏；不得记录 API Key、完整文档内容或视觉证据载荷。
- VB.NET/VSTO、构建和验证硬规则以根 `CLAUDE.md` 为准。

## ANTI-PATTERNS

- 在 `ShareRibbon` 中引用 PowerPoint/Word/Excel COM 类型。
- 在单个插件中复制 `AgentKernel`、`LoopEngine`、`AiGateway` 或共享 ToolResult 合同。
- 把 Base64 截图序列化进工具结果、Trace 或聊天历史。
- 在 Executor 内直接调用模型或实现第二套视觉修复循环。
- 绕过 `ToolRegistry`、`SafetyGate` 或 `CodeExecutionService` 执行宿主写操作。
- 硬编码 API Key、Provider URL 或模型能力假设。
- 用 `SetWindowsHookEx`、`RegisterHotKey`、`keybd_event`、`SendInput` 实现快捷键：前者是全局注入风险，后两者会
  伪造按键。快捷键只用 `GetAsyncKeyState` 轮询加前台判断，见 `Controls/Services/HostHotkeyWatcher.vb`。
- 把快捷键监听、面板状态或话术清单复制到某个宿主项目里；也禁止为命令面板另开一条执行/发送路径。
- 在宿主 `AppendCurrentSelectedContent` 里重新判 `selectedCellChecked`，或绕过 `BuildMessageWithAutoSelection`
  直接调用它：那等于把九个可漂移的开关重新放回宿主，开关就再也无法被「一次关掉」。
- 把调用方传来的 `app_type` 直接落库，或在读侧为了「老数据别丢」而放宽成「匹配不上也算」：前者造出三端都看不见
  的数据，后者把别的宿主的历史漏进当前端。归一只允许发生在写入边界（`AppTypeContract.Normalize`）。
- 在 `history-manager.js` 里对会话列表排序、或按 `documentName` / `createdAt` 分组：前者冲掉后端「当前文档最前」，
  后者让同名不同文件并成一组、时间桶与列表顺序自相矛盾。分组键只允许用后端回传的 `documentId` 与 `lastActiveAt`。
