# POWERPOINT AI ADD-IN

## OVERVIEW

`PowerPointAi/` 是 PowerPoint 宿主项目，负责 Presentation/Slide/Shape COM 操作、专业 Deck 设计、Scene 编译、真实渲染验证和宿主级安全回滚。共享 Harness/Loop/ToolRegistry/AI 请求仍由 `ShareRibbon` 提供。

## STRUCTURE

```text
PowerPointAi/
├── ChatControl.vb                    # PowerPoint 命令入口与宿主执行回调
├── Context/PowerPointContextProvider.vb
├── HostEntry/PowerPointContextMenu.vb  # E4-02 右键菜单适配器（只挑菜栏 + 判选区类型）
├── Design/
│   ├── SlideDesignModels.vb          # Deck/Slide/Scene/Verifier 模型
│   ├── DesignSystemCatalog.vb        # 设计 Token 与主题
│   ├── SlideComponentLibrary.vb      # 可编辑 Scene 组件
│   ├── SlideLayoutEngine.vb          # Scene 语义到专业构图计划
│   ├── PowerPointSceneRenderer.vb    # PowerPoint COM 渲染
│   ├── PowerPointVisualVerifier.vb   # Preflight、渲染后验证与像素检查
│   ├── ScenePlanPreviewRenderer.vb   # 非写入预览
│   └── ProfessionalDeckExecutor.vb   # CreateSlides 事务、观察、回滚与证据采集
├── Runtime/                          # 动态 Office API Catalog/Resolver/Executor/Observer
├── Ribbon1.vb
└── ThisAddIn.vb
```

## WHERE TO LOOK

| Task | Location | Notes |
|---|---|---|
| PowerPoint 命令接入 | `ChatControl.vb` | 保持 `CodeExecutionService -> ExecuteJsonCommandWithToolResult` 主入口 |
| 专业 Deck 执行 | `Design/ProfessionalDeckExecutor.vb` | `CreateSlides`、preview、事务、观察和安全回滚 |
| Scene 语义与输入 | `Design/SlideDesignModels.vb` | slideType、variant、chart、table、metrics 等 |
| 构图 | `Design/SlideLayoutEngine.vb` | Deck/Slide 级专业布局与长文本适配 |
| 可编辑组件 | `Design/SlideComponentLibrary.vb` | Shape、图表、表格、连接线等 Scene primitives |
| COM 渲染 | `Design/PowerPointSceneRenderer.vb` | UI 线程执行与 COM 生命周期 |
| 视觉验证 | `Design/PowerPointVisualVerifier.vb` | 几何、层级、结构、像素与 COM 渲染差异 |
| 动态 PowerPoint API | `Runtime/` | `DiscoverOfficeCapability + OfficeObjectOperation` 的宿主实现 |
| PowerPoint 上下文 | `Context/PowerPointContextProvider.vb` | 活动文稿、选区和结构快照 |
| 右键菜单入口 | `HostEntry/PowerPointContextMenu.vb` | 只挑菜栏 + 判 `ActiveWindow.Selection.Type`；挂项/摘项在 `ShareRibbon/Agent/HostContextMenu.vb` |

## HOST ENTRY RULES

- 入口只说话：动作来自 `ShareRibbon/Agent/QuickActionCatalog.ActionsFor(PowerPointAppType)`，点击后一律
  `chatControl.StartAgentFromHostEntry(action)`（整个动作，含入口 id）`，不得再拼 `sendMessageToServer` JS。
- 入口层不自建清单、不声明 `CapabilityId`：PowerPoint 还没有宿主 capability 注册表（开放待办 #36）。
- 新入口文件放 `HostEntry/` + `Namespace HostEntry`。`Namespace Selection` 在本项目是**编译炸弹**：
  VB 里命名空间压过互操作类型名，`PowerPointContinuationService`/`PowerPointCompletionManager` 里的
  `sel As Selection` 会直接报 BC30182。
- 引用的 Office 15 PIA 的 `CommandBar` 没有事件面，所以右键菜单不做"弹出时置灰"，而是点击时探测
  （`ppSelectionShapes/Text/Slides` 才算可用）并把原因写进状态栏。
- 命令面板（E4-03）的快捷键轮询与面板状态全在 `ShareRibbon/Controls/Services/`，PowerPoint 侧只贡献一个
  `ChatControl.EnsureHostTaskPaneVisible` 覆写（`Globals.ThisAddIn.ShowChatTaskPane()`）；不在本项目里建
  watcher/timer，也不为面板另开第二条发送路径。快捷键只在 PowerPoint 处于前台时生效，多宿主共用一个组合键不互相抢键。
- 右键菜单受 `ChatSettings.EnableContextMenu`（默认关闭）与面板可见事件保护，不进 `PowerPointAi_Startup`；
  Shutdown 必须 `Detach()`，只按 `aiq:` Tag 摘自己挂的项。日志只记动作 id，幻灯片文字不进日志。
- 选区自动引用（E4-04）默认开启，「带不带选区」只在 `ShareRibbon` 的 `BuildMessageWithAutoSelection` 判一次；
  `ChatControl.AppendCurrentSelectedContent` 里那份重复的 `If Not selectedCellChecked` 已删除。引用卡片那条路
  （`GetSelectedContent` 的事件芯片）是另一个功能，仍自带闸门，不要混为一谈。
- PPT 各版本暴露的上下文菜栏零散，`{"Text","Drawing","Object"}` 只是候选试探：取不到就不出这个入口，
  不报错、不重试。真机需按版本记录实际命中哪条。
- 事件绑定卫生（E4-05）：`ChatTaskPane_VisibleChanged` 在 `EnsureContextMenu()` 之前先
  `ChatSettings.EnsureLoaded(New ApplicationInfo("PowerPoint", OfficeApplicationType.PowerPoint))`，否则进程内第一次
  读到的是代码默认值。`ChatControl.MainForm_Load` 不调 `InitializeSettings()`（页面注入由基类「一份文档一次初始化」闸门负责）；
  宽度定时器只在 `Is Nothing` 守卫里创建、`ThisAddIn_Shutdown` 统一 `Stop()`+`Dispose()`，不要在可见事件里再挂一份 Tick。

## PROFESSIONAL DESIGN PIPELINE

- 唯一专业生成入口保持 `CreateSlides + Scene`；不得为 cover、timeline、SmartArt、architecture 等页面类型新增专用 Tool。
- 主链必须保持 `OfficeHarness -> AgentKernel -> LoopEngine -> ToolRegistry -> CodeExecutionService -> PowerPoint host -> ProfessionalDeckExecutor`。
- `ProfessionalDeckExecutor` 负责整套 Deck 预编译、渲染、Observer/Verifier、ToolResult 和事务回滚，但不得直接调用模型。
- `SlideLayoutEngine` 负责把 Scene 语义编译为可验证计划；禁止在 `ChatControl` 或 Renderer 内堆用户话术关键词路由。
- `PowerPointSceneRenderer` 只渲染计划并管理 COM；不能自行重新规划版式或启动线程。
- `PowerPointVisualVerifier` 必须验证真实 PowerPoint 结果，而不只相信 Scene 计划或 COM 返回成功。
- preview 只能编译和预检，不得创建再删除幻灯片来模拟预览。

## VISUAL EVIDENCE AND ROLLBACK

- 渲染后验证失败时，可在回滚前导出实际幻灯片截图作为当前 Repair 轮次的视觉证据。
- 截图必须限制格式、像素尺寸和字节数，使用唯一临时文件，并在 `Finally` 中删除。
- 截图只能写入共享的临时内存证据合同；禁止放入 `Observation`、`Data`、`Artifacts`、History、Trace、日志或 Notes。
- Executor/Verifier 不得调用视觉模型；证据由现有 `LoopEngine` 消费并负责 Provider 降级。
- 回滚只能删除能通过本批 `office-ai-design:` 标记安全识别的页面或对象；不能按“初始页数之后全部删除”猜测回滚范围。
- 截图必须发生在回滚之前；截图失败不能阻止安全回滚，也不能把不完整修改伪装成成功。

## COM RULES

- PowerPoint COM 类型和实现属于 `PowerPointAi`，不得下沉到 `ShareRibbon`。
- 所有 COM 操作必须在共享宿主 UI 线程入口内执行；Executor/Renderer 不得创建工作线程访问 COM。
- 临时取得的 `Presentation`、`Slides`、`Slide`、`ShapeRange`、`FillFormat`、`TextFrame2` 等 COM 对象必须在 `Finally` 中显式释放。
- 不硬编码 Placeholder 索引；应按 `PpPlaceholderType` 查找并在写入后回读验证。
- `Slide.Export`、字体度量、透明度、阴影、文本边界等结果可能因 PowerPoint 版本和主题不同而变化，必须以渲染后 Observer 为准。

## CONVENTIONS

- PowerPoint 专属逻辑只放在本项目，不访问 Excel 或 Word 对象模型。
- 新增 `.vb` 或 fixture 后同步加入 `PowerPointAi.vbproj`。
- Scene 字段必须显式消费；不支持的 `slideType`、`variant`、`designSystem` 应返回可修复 schema 错误，不能静默退化。
- 长文本不能靠无限缩小字号解决；优先拆页、改变构图、压缩语义或明确拒绝。
- 视觉质量失败返回结构化 `VERIFY_FAILED` 和可解释指标，由现有 Loop 决定修复。

## ANTI-PATTERNS

- 在 `PowerPointAi` 内复制 Harness、Loop、ToolRegistry、SafetyGate 或 AI Gateway。
- 为每一种页面布局新增 Tool、意图枚举或 ChatControl 分支。
- 在 Executor/Verifier 内直接调用 LLM 或实现第二套 Repair Loop。
- 只依据 Shape 数量、COM 无异常或像素颜色数量宣称专业质量通过。
- 把 Base64 截图放入普通 ToolResult JSON、日志、Notes 或持久化存储。
- 失败时删除所有新增页、忽略部分应用状态，或返回假成功。
- 跨线程访问 COM，或遗漏释放临时 COM 对象。
