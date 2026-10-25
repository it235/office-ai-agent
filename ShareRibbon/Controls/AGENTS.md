# SHARED UI AND REQUEST ORCHESTRATION

## OVERVIEW

`ShareRibbon/Controls/` 承载跨 Office 插件共享的 WebView2 Chat UI、请求编排、AgentKernel 接线和流式通信。这里是 UI/运行时适配层，不是宿主 COM 执行层。

## STRUCTURE

```text
Controls/
├── BaseChatControl.vb                 # 共享 Chat UI 与路由入口
├── Models/                            # History、Selection 等 UI 模型
├── Services/AgentKernelService.vb     # Harness/AgentKernel 与 UI 的接线
├── Services/ChatRequestOrchestrator.vb
├── Services/ChatRoutingOrchestrator.vb
├── Services/CodeExecutionService.vb   # 进入宿主 UI 线程执行入口
├── Services/HttpStreamService.vb      # 普通流式请求
└── Services/                          # 其他共享 UI/会话服务
```

## WHERE TO LOOK

| Task | Location | Notes |
|---|---|---|
| Chat UI / WebView2 | `BaseChatControl.vb` | 收集输入、上下文、展示状态，不堆业务路由 |
| Agent 启动与事件 | `Services/AgentKernelService.vb` | 绑定 AgentKernel、Harness、AI 请求和 UI 终态 |
| 失败步骤的单步纠偏 | `BaseChatControl.HandleAgentStepAction` + `Resources/js/agent-card.js` | 面板只收下并回执，落地在循环的步骤边界；被拒必须经 `setStepActionResult` 把原因送回卡片；终态统一 `lockAllStepActions`；不复用 `agent:approveStep`（超时/放弃语义） |
| 普通 Chat 请求 | `Services/ChatRequestOrchestrator.vb` | 构建文本请求和历史 |
| 路由 | `Services/ChatRoutingOrchestrator.vb` | 选择既有执行路径，不复制 Agent 逻辑 |
| 宿主执行边界 | `Services/CodeExecutionService.vb` | 统一调度到宿主执行回调/UI 线程 |
| Streaming / MCP | `Services/HttpStreamService.vb` | 流式响应和兼容工具调用 |
| 命令面板 / 快捷键 | `Services/CommandPaletteService.vb`、`Services/HostHotkeyWatcher.vb` | 向面板推状态、注册组合键；只发 `startAgent`，永不自己起 run |
| 选区自动引用 | `BaseChatControl.BuildMessageWithAutoSelection` | 「要不要带选区」的唯一决策点；三端 `AppendCurrentSelectedContent` 只回答「怎么写成提示」 |
| 页面装载与事件绑定 | `BaseChatControl` 的 `_pageBootstrapDone` + `OnWebViewNavigationCompleted` | 一份文档一次初始化；设置/MCP/快捷键注入的唯一入口 |
| 入口可达性 | `scripts/check-entry-reachability.js` | 面板 type、内联事件、DOM id、`vsto` 桥方法、三端 Ribbon 可见按钮、VB 注入载荷的字面量 id 与调用名、导航文档引用的路径十个方向的闭合合同（规则 A–J） |
| 门禁清单 | `scripts/run-code-checks.ps1` | 唯一持有者：`scripts\*.ps1` 默认全是门禁，退出必须在 `$notGates` 里具名写理由；`build\RunReleaseChecks.ps1` 委托给它，不再自备名单 |

## REQUEST BOUNDARIES

- `BaseChatControl` 只负责输入、上下文、路由和展示；专业 PowerPoint 构图、验证或修复逻辑不得写在这里。
- `AgentKernelService` 只负责把共享 AI 请求能力绑定到 `AgentKernel`，不得创建第二套 Agent/Loop。
- 普通聊天历史仍是文本模型；临时视觉证据不得写入 `HistoryMessage`、ChatState、Memory 或 WebView2 消息。
- 多模态 Repair 请求只能消费 `LoopEngine` 提供的临时 messages，并通过共享 `AiGateway` 发送；请求体不得记录或持久化。
- Provider 不支持图片时返回失败给 Loop 处理降级，不在 UI 服务层递归重试。
- 数据边界同意（E3-05）只有一个持有者：`Services/PrivacyConsentGate.vb`。判定（`IsAccepted`/`RefuseQuietly`）、
  写入（`Apply`）、拒绝理由（`RefusalReason`）和三条边界文案（`BoundaryItems`）都在那里，向导、设置行和阻断卡片
  都读同一份 payload，**前端不得再抄一份文案**。拦截点按出口而不是按入口：面板两条漏斗用
  `BaseChatControl.IsBlockedByPrivacy`（`Send` / `HandleStartAgent`），非流式发送器 `SendAndGetResponseAsync`
  与 `UtilsService`/`LLMUtil`/补全/embedding/翻译 各自在构造请求之前问一次；新增一个 HttpClient 出口必须
  登记进 `scripts/smoke-privacy-consent.ps1` 的 `$outboundInventory` 并写清它被哪把门看着，否则门禁红。
- 阻断卡片只有一张：缺配置与缺同意共用 `ModelConfigPrompt` + `showModelConfigRequired`，靠 `action`
  （`open_config` / `accept_privacy`）分支按钮动作，**不要**为新拦截再画第二张卡片或新增消息 type。
- 用量（E3-04）三层各只有一个持有者：**解析**在 `Services/Ai/TokenUsage.vb`（OpenAI 与 Anthropic 两套键名在那里
  收成一个形状，读不到 usage 返回 `Nothing` 而不是 0，provider 给了更大的 `total_tokens` 时以它为准）；
  **账本**在 `ChatStateService`，`SessionUsage` 是会话累计、`RunUsage` 是本轮切片，写入口只有 `RecordUsage`
  一个（`AddTokens` 已删）；**文案**在 `Controls/Services/UsageCostReporter.vb`，token 句、金额句、额度句、
  输入框哨兵渲染都在那里。终态卡片与流式脚注只是两个投递点，**JS 不得自己算钱或另拼一句费用**；
  没配单价时只显示 token 并说明原因，额度提醒永远不拦请求（能拦的只有同意）。
- 面板开关必须能真的关掉：`HandleSaveSettings` 一律从载荷读回每一个开关（缺键时保持当前值），**不得**再用静态字段
  原样写回，否则那个开关是只读的；`ChatSettings.SaveSettings` 读改写同一个 `JObject`，只覆盖自己认识的键，
  不认识的历史键保留。
- `CommandPaletteService` 只下发状态和唤起面板；起 run 的只有一处 JS `startAgent`。不要为面板按钮新增
  `router.Register`，也不要把快捷键监听、话术清单或执行分支复制进某个宿主。
- 选区自动引用（`selectedCellChecked`，E4-04 起默认开）只在 `BuildMessageWithAutoSelection` 里判定一次：
  宿主 override 里再判一次就会变成九个各自可漂移的开关。默认开启的前提是面板里有能真的关掉的开关，
  以及带出去的选区要在对话里显示出来——Agent 回合不渲染用户气泡，所以引用摘要必须走
  `appendTurnSelectionReference` 显示通道，摘要只从**实际附加上的增量**反推，不再采集一次上下文。
  关掉开关同时停掉 Word 的 `PendingSelectionInfo` 记录，这是「不引用也就不写回」的同一条语义。

## EVENT AND LIFETIME HYGIENE

- 一个处理器只允许一份绑定。`OnWebViewNavigationCompleted` 只用声明式 `Handles ChatBrowser.NavigationCompleted`；
  再 `AddHandler ChatBrowser.CoreWebView2.NavigationCompleted` 等于同一次导航跑两遍，而在处理器里
  `RemoveHandler` 那个 `CoreWebView2.*` 事件**摘不掉声明式那一份**——这种「解绑摘错对象」的写法已被门禁禁止。
- 页面侧初始化（`InitializeSettings`/`InitializeMcpSettings`）的粒度是**一份文档一次**：闸门
  `_pageBootstrapDone` 只在实际装载新文档（`NavigateToString`）之前复位，处理器必须
  「先判 → 立刻占位 → 才 `Await`」。占位跨过 `Await` 就失去意义，因为 `Task.Delay`/`InvokeAsync` 期间
  排队的第二次导航完成同样能通过检查。
- 新增注入点时先问「这段脚本重复执行会不会有副作用」。在这里，重复执行的**症状是静默失效或值被反复洗回**，
  不是异常，所以不能靠肉眼判断，必须由 `scripts/smoke-event-binding.ps1` 的顺序/次数守卫钉住。
- 读配置开关之前必须确保配置被读过一次：宿主在面板可见时才判断 `EnableContextMenu`/`EnableSelectionToolbar`，
  那一刻进程内可能还没有 `ChatSettings` 实例，拿到的是代码默认值。统一用
  `ChatSettings.EnsureLoaded(applicationInfo)`（首次才读盘，重复调用返回同一实例），不要在 `Startup` 里抢跑，
  也不要各宿主自己 `New ChatSettings(...)` 造第二份状态。

## ENTRY REACHABILITY (E5-04 / E5-05)

- 不可达入口在这里的症状是**点了没反应**，不是异常：`ChatCommandRouter.Dispatch` 对未注册的 type 只
  `Debug.WriteLine`，注释掉的按钮、指向不存在函数的 `onclick`、查不到的 `getElementById`、只有注释的
  Ribbon handler 全部静默通过。所以入口按合同核，不按肉眼核：`scripts/check-entry-reachability.js`
  同时钉住「面板发的 type 必须注册」「注册的 type 必须有发送方」「HTML 注释里不留可交互控件」
  「内联事件的函数必须有定义」「字面量 id 必须有节点」「`vsto` 桥方法必须有调用点」「可见 Ribbon 按钮必须有非空 handler」。
- 语料必须包含四个模块的 `.vb`。VB 会拼 HTML 标记（`SmartProofreadFocusMode.vb` 的 `id="proofread-summary"`），
  也会注入 JS（`UtilsService.GetVstoBridgeScript()` 里的发送方）。只扫 `Resources/` 会得出**相反结论**并误删正常功能。
- E5-05 起，「VB 注入的 JS 必须成立」这条合同只有规则 H/I 一个持有者（`audit-p0-guardrails.ps1` 的 P4-C 片段与
  `smoke-selection-reference.ps1` §6 的 id 循环都是同一合同的副本，已删）。副本的危害不是重复劳动，而是能力不同：
  那份 P4-C 把带 VB 插值的载荷整条跳过，而注入载荷绝大多数带插值，所以它长期零发现。
  H 查注入块里的字面量 id（全文扫，不受词法边界影响），I 查注入块调用的函数名（VB 词法读载荷，取下限）。
- 一条坏注入有两种合法收尾，判据是「谁拥有这块 DOM」：功能本来就该有就**接通**
  （`WordAi/ChatControl.vb` 的 `showProofreadLoading()` 面板里从没这个函数，改成既有的 `showLoadingIndicator()`，
  收起仍由 `changeSendButton()` 负责）；职责已有唯一持有者就**删除**（`AgentKernelService.vb` 那两处
  `planning-status-card` 只是声称复位规划卡，卡片实际归 `agent-card.js` 的 `'agent-plan-' + uuid`，
  终态复位实际由 `restoreAgentRequestUi()` 负责，两段一起删掉）。
- 同一条防漂移思路也用在文档上：规则 J 核导航文档（根 AGENTS.md/CLAUDE.md/README.md、各级 `AGENTS.md`、
  `Storage/README_DATABASE.md`）里引用的本地路径是否真的存在。本轮实测有 8 条坏引用（根地图指向
  docs/build-and-installer.md 与 docs/signing-and-certificates.md，两个文件从未存在），症状不是报错而是
  「下一个人重新考古一遍」。只核 markdown 链接与以顶层目录开头的反引号路径：把裸文件名也算进来会得到 13 条假阳性，
  而一条 noisy 的门禁最后只会被人关掉。
- 新增门禁脚本不需要在名单里登记：`scripts\*.ps1` 默认被 `run-code-checks.ps1` 发现并执行。
  反过来，让一条门禁**不**跑必须在 `$notGates` 里写下理由——「忘了登记」和「故意不跑」在磁盘上长得一样，
  只有写下来的理由能区分。缺 node 同样不再静默跳过，而是 `throw`。
- 新增入口只有两种合法收尾：接通到已注册 type（复用 `sendMessageToVB`/`startAgent` 那一条主链，
  不另开发送路径），或删掉。第三种「先留着以后接」在这里等于留一条验证不了的路径。
- `GetOfficeAppType()` 是 `MustOverride`：漏实现要在编译期暴露，而不是运行时变成一个既不属于任何宿主、
  又拿不到空串豁免的 `app_type`（见 `Storage/AppTypeContract.vb` 与迁移 019）。

## CONVENTIONS

- UI 组件继承 `UserControl`，WebView2 资源位于 `ShareRibbon/Resources/`。
- 长任务必须异步；Office COM 调用只能通过既有宿主 UI 线程入口。
- 所有 Agent 终态都必须恢复输入框、发送/停止按钮和 planning card 状态。
- 修改服务构造函数或委托签名时，同时检查三个宿主的 `ChatControl` 接线。

## ANTI-PATTERNS

- 在共享 UI 控件中直接访问 PowerPoint、Word 或 Excel COM。
- 把专业设计、Scene 或页面类型判断堆进 Chat 路由。
- 把截图 Data URL/Base64 放入 History、JavaScript 消息、日志或 RunTrace。
- 在多模态失败时由 UI 层无限重试或静默吞掉错误。
- 阻塞 UI 线程等待网络请求或文件 IO。
- 硬编码 API Key、Provider URL 或面向用户的敏感调试信息。
- 在注入的 JS 里引用面板中不存在的元素 id。注入块是顺序执行的，`getElementById(...)` 拿到 null 后取属性
  抛 TypeError，**其后每一行都不再执行**，症状是「设置项同步静默失效」而不是报错；`scripts/check-entry-reachability.js`
  的规则 H 会核对四模块全部 `.vb` 里注入块引用到的每个字面量 id 是否真的存在于面板标记。
- 在宿主 `AppendCurrentSelectedContent` 里重新判 `selectedCellChecked`，或把引用摘要顺手写进历史、日志、
  Observation、Artifacts。
- 在宿主 `MainForm_Load` 里自己再调一次 `InitializeSettings()`（Excel 曾这样写过），或用第二种绑定方式给已绑定的
  事件再挂一次同一个处理器。
- 为了「保险」在导航完成里反复重跑页面初始化，或反过来把闸门做成进程级一次性、导致重建文档后不再注入设置。
- 保留一条没有发送方的 `router.Register`，或给已经删掉的入口留 `Select Case`/注释掉的按钮——它不参与任何
  编译与运行检查，只会把「这条能力到底有没有」变成一次真机考古。删除时连注释实现一起删，注释里只留
  「为什么删」，由 `check-entry-reachability.js` 与 `audit-p0-guardrails.ps1` 负责让它不能悄悄回来。
