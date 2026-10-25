# EXCEL AI ADD-IN

## OVERVIEW
Excel-specific add-in providing AI-driven data analysis and automation capabilities. Integrates with Excel through VSTO and uses shared components from ShareRibbon.

## STRUCTURE
```
ExcelAi/
├── ChatControl.vb    # Excel-specific chat interface
├── HostEntry/        # 宿主原生入口（E4-02 右键菜单适配器）
├── Ribbon1.vb        # Excel ribbon implementation
├── ThisAddIn.vb      # Excel add-in entry point
└── ExcelFunctions.vb # Excel-specific functionality
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Excel ribbon customization | Ribbon1.vb | Excel tab and button definitions |
| Excel data analysis | ExcelFunctions.vb | Cell and range operations |
| Chat interface | ChatControl.vb | Excel-specific chat UI |
| Add-in initialization | ThisAddIn.vb | Startup and shutdown logic |
| 右键菜单入口 | HostEntry/ExcelContextMenu.vb | 只挑菜栏 + 判 `Selection` 是否 `Range`；挂项/摘项在 `ShareRibbon/Agent/HostContextMenu.vb` |

## HOST ENTRY RULES
- 入口只说话：动作来自 `ShareRibbon/Agent/QuickActionCatalog.ActionsFor(ExcelAppType)`，点击后一律
  `chatControl.StartAgentFromHostEntry(action)`（整个动作，含入口 id）`，不得再拼 `sendMessageToServer` JS。
- 入口层不自建清单、不声明 `CapabilityId`：Excel 还没有宿主 capability 注册表（开放待办 #36），
  「入口说得到、能力做不到」比少一个入口更糟。
- 新入口文件放 `HostEntry/` + `Namespace HostEntry`。**不要用 `Namespace Selection`**：VB 里命名空间会压过
  互操作类型名，`sel As Selection` 这类声明会直接报 BC30182。
- 右键菜单受 `ChatSettings.EnableContextMenu`（默认关闭）与面板可见事件保护，不进 `ExcelAi_Startup`；
  Shutdown 必须 `Detach()`，本加载项只按 `aiq:` Tag 摘自己挂的项。日志只记动作 id，区域地址与内容不进日志。
- 命令面板（E4-03）的快捷键轮询与面板状态全在 `ShareRibbon/Controls/Services/`，Excel 侧只贡献一个
  `ChatControl.EnsureHostTaskPaneVisible` 覆写（`Globals.ThisAddIn.ShowChatTaskPane()`）；不在本项目里建
  watcher/timer，也不为面板另开第二条发送路径。快捷键只在 Excel 处于前台时生效，与 Word 共用组合键不互相抢键。
- 选区自动引用（E4-04）默认开启，「带不带选区」只在 `ShareRibbon` 的 `BuildMessageWithAutoSelection` 判一次；
  `ChatControl.AppendCurrentSelectedContent` 只负责把区域写成提示文本（含 100 行 × 26 列上限），
  **不得**再在里面判 `selectedCellChecked`，也不要在 `DoubaoChat`/`DeepseekControl` 的覆写里各补一份闸门。
- 事件绑定卫生（E4-05）：`ChatControl.MainForm_Load` 只做 WebView2 初始化，**不再**自己调 `InitializeSettings()`
  （Excel 曾与基类的导航完成各调一次）；页面注入统一由 `BaseChatControl` 的「一份文档一次初始化」闸门负责。
  `ChatTaskPane_VisibleChanged` 在 `EnsureContextMenu()` 之前先
  `ChatSettings.EnsureLoaded(New ApplicationInfo("Excel", OfficeApplicationType.Excel))`，否则进程内第一次读到的是代码默认值。

## CONVENTIONS
- Excel-specific functionality only
- Inherits from shared BaseChatControl
- Uses Excel interop through ShareRibbon services
- Follows Excel VSTO patterns

## ANTI-PATTERNS
- Never access Word or PowerPoint objects
- Never duplicate shared functionality from ShareRibbon
- Never bypass Excel security model