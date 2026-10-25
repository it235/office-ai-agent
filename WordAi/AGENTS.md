# WORD AI ADD-IN

## OVERVIEW
Word-specific add-in providing AI-driven document processing and text manipulation capabilities. Integrates with Word through VSTO and uses shared components from ShareRibbon.

## STRUCTURE
```
WordAi/
├── ChatControl.vb         # Word-specific chat interface
├── Ribbon1.vb             # Word ribbon implementation
├── HostEntry/             # 宿主原生入口（E4-02 右键菜单适配器），新增入口代码放这里
├── Selection/             # 选区浮层入口（E4-01）：几何 / 浮层 / 编排三个类
├── ThisAddIn.vb           # Word add-in entry point
└── WebDataCapturePane.vb  # Web content capture
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Word ribbon customization | Ribbon1.vb | Word tab and button definitions |
| Document processing | WebDataCapturePane.vb | Text extraction and manipulation |
| Chat interface | ChatControl.vb | Word-specific chat UI |
| Add-in initialization | ThisAddIn.vb | Startup and shutdown logic |
| Selection quick actions | Selection/WordSelectionToolbarController.vb | 只决定"何时出现、出现什么"；动作话术在 `ShareRibbon/Agent/QuickActionCatalog.vb`，执行一律交回 `StartAgentFromHostEntry` |
| 右键菜单入口 | HostEntry/WordContextMenu.vb | 只挑菜栏 + 探选区；挂项、按 Tag 摘项、事件解绑都在 `ShareRibbon/Agent/HostContextMenu.vb` |

## CONVENTIONS
- Word-specific functionality only
- Inherits from shared BaseChatControl
- Uses Word interop through ShareRibbon services
- Follows Word VSTO patterns

## HOST ENTRY RULES
- 宿主原生入口（选区浮层、右键菜单、之后的快捷键）只把入口动作交给 `BaseChatControl.StartAgentFromHostEntry(action)`，
  不拼 `sendMessageToServer` 的 JS 字符串、不新建执行链：闸门、上下文采集、预览/风险/撤销都已经在主链上。
- 入口层不得自建动作清单：话术只有 `ShareRibbon/Agent/QuickActionCatalog.vb` 一处，Word 侧取 `ActionsFor(WordAppType)`。
- 新增入口代码一律放 `HostEntry/` + `Namespace HostEntry`。`Selection/` 里的三个类是 E4-01 的历史落点，
  **不要再往那个命名空间加文件**：VB 解析名字时命名空间会压过互操作类型，同名会让本项目其它文件的
  `sel As Selection` 之类的声明报 BC30182（"应为类型"），PowerPointAi 就是这么被 `Namespace Selection` 打挂的。
- 浮层必须不抢焦点（`WS_EX_NOACTIVATE` + `ShowWithoutActivation`），否则点按钮会先清掉它要作用的选区。
- 右键菜单只能改自己认领的项：挂进宿主菜单的控件一律 `Temporary:=True` + `aiq:` Tag 前缀，摘除按 Tag 倒序删除，
  Shutdown 必须 `Detach()`。引用的 Office 15 PIA 的 `CommandBar` 没有事件面，所以不做"弹出时置灰"，
  而是点击时探测（`Selection.Type <> wdSelectionIP`）并把原因写进状态栏。
- WordAi 是 Application 级加载项，判断"能不能出浮层"看 `ActiveWindow.Type = wdWindowDocument`，不要数文档个数。
- 命令面板（E4-03）的快捷键轮询与面板状态全在 `ShareRibbon/Controls/Services/`，Word 侧只贡献一个
  `ChatControl.EnsureHostTaskPaneVisible` 覆写（`Globals.ThisAddIn.ShowChatTaskPane()`）；不要在本项目里
  建 watcher/timer，也不要为面板另开一条发送路径。Word 本身已占用 `Ctrl+K`（插入超链接），
  默认组合键 `Ctrl+Alt+K` 避开它。
- 选区自动引用（E4-04）默认开启，但「带不带选区」由 `ShareRibbon` 的 `BuildMessageWithAutoSelection` 判一次；
  `ChatControl.AppendCurrentSelectedContent` 只把选区写成提示，**不得**再在里面判 `selectedCellChecked`。
  这个方法同时负责记录 `PendingSelectionInfo`（供本轮选区对比/写回），所以关掉开关也就不再按选区写回——
  这条耦合是设计，已经写进设置项说明。
- 选区相关代码只订阅自己的事件并保持 `Initialize/Dispose` 对称（ghost text 也在 `WindowSelectionChange` 上），
  且入口创建挂在面板变为可见的事件上（折叠按钮那条路不经过 `ShowChatTaskPaneAsync`），不进 `WordAi_Startup` 关键路径。
- 事件绑定卫生（E4-05）：`ChatTaskPane_VisibleChanged` 在 `EnsureSelectionToolbar()`/`EnsureContextMenu()` **之前**
  先 `ChatSettings.EnsureLoaded(New ApplicationInfo("Word", OfficeApplicationType.Word))`——那一刻读的是配置文件里的开关，
  而进程内可能还没有任何 `ChatSettings` 实例读过盘。`CreateChatTaskPane` 有实例守卫，所以
  `AddHandler chatTaskPane.VisibleChanged` 每个面板只会挂一次；`HostContextMenu.Attach` 自带 `If _attached Then Detach()`，
  不要在外面再加一层「防止重复挂」的标记。
- 日志只记动作 id 与数量，选区原文永不进 `AppLogger`。三个原生入口的开关都在面板设置里可见、默认关闭：
  `enableSelectionToolbar` / `enableContextMenu` / `enableHotkey`（`我的文档\OfficeAiAppData\office_ai_chat_settings.json`，
  组合键为同一文件里的 `hotkeyChord`，默认 `Ctrl+Alt+K`）。同一文件里的 `selectedCellChecked` 默认开启，
  但一样能在面板里关掉。

## ANTI-PATTERNS
- Never access Excel or PowerPoint objects
- Never duplicate shared functionality from ShareRibbon
- Never bypass Word document protection