# Week 4 落地报告与验收清单

> 日期：2026-10-25
> 范围：E1-07 Excel/PPT 接入本轮撤销、E4-03 全局快捷键 / 命令面板、E4-04 选区自动引用默认开启、E4-05 事件绑定卫生、E2-04 会话归属收口（迁移 019 + 写边界归一）、E2-05 历史 UI 分组与空状态引导、E5-04 不可达入口处理、E5-05 CI 防漂移，以及 Phase 2 的 E6-02/03/04 任务包合同、一致性门禁与面板归组和 E6-01 三端 Ribbon 话术收口。
> 结论：代码与自动化校验已通过；按用户指示本轮未做真机验收。

## 1. E1-07 Excel/PPT 接入本轮撤销

**验收标准**：两端取消/失败后可撤销本轮，行为与 Word 一致。

Word 的一致性不等于“同样调用宿主撤销栈”——Excel 与 PowerPoint 根本没有“一条命令 = 一个撤销记录”的边界（`UndoManager.CreateExcelUndoPoint` 明确把 `CanUndo` 置为 False，PowerPoint 只有 CreateSlides 链的按标记删页）。所以“行为一致”落在**用户看到的事实**上：本轮结束时会清楚说明能不能撤、依据是什么，能撤就一键回到执行前，不能撤就给出原因，而不是留下一个静默消失的按钮。

### 1.1 共享合同：两条互斥的可撤销路径

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 还原快照字段 | `ShareRibbon/Agent/ToolRegistry.vb`（`ToolResult.RestoreSnapshot As JObject`，`<JsonIgnore>`） | 与 `VisualEvidence` 同类，走**临时旁路**：不进 `Observation`、`Data`、`Artifacts`，因此不会被序列化进提示词、执行轨迹、聊天历史或日志 |
| 记录字段 | `ShareRibbon/Agent/Execution/AgentChangeLog.vb`（`AgentChangeRecord.RestoreJson`/`NativeUndoUnit`/`HasRestore`） | 变更记录同时知道“宿主给了撤销单元”和“这一步有没有可重放的快照”；`ToJson()` 只输出 `hasRestore` 布尔，不输出快照正文 |
| 体积闸门 | `AgentRunChangeLog.MaxRestoreChars = 60000`、`BuildRestoreNode(kind, state)` | 宿主采集后必须先过这里：超限直接返回 `Nothing`，这一步保持“不可撤销”，而不是把大块正文写进本地库或带着半份快照承诺一键回退 |
| 两路判定 | `NativeUndoable` / `RestoreCount` / `FullyUndoable` / `WrittenEntries` | `FullyUndoable = 全部残留改动都有宿主撤销单元且首尾文档指纹齐备` **或** `全部残留改动都有还原快照`；两条路都要求覆盖**全部**残留改动，半撤半留一律不算可撤销 |
| 范围语义 | `AgentRunUndoScope.NativeUndoable/RestoreCount/RestoreOnly` | 撤销范围能区分“走宿主撤销栈”和“只能按快照重放”，供宿主选择各自的回退实现 |
| 半应用不丢依据 | `ShareRibbon/Agent/LoopEngine.Planning.vb` `ValidateObservedOutcome` | 把 `Success=True` 的批次降级成 `CodePartialApply` 失败时显式带上 `RestoreSnapshot`——部分应用的这一步恰恰最需要“撤销本轮” |

主链未变：`OfficeHarness -> AgentKernel -> LoopEngine -> ToolRegistry -> CodeExecutionService -> 宿主 ExecuteJsonCommandWithToolResult`。快照由宿主在执行前后各采一次，LoopEngine 只把同一个 `toolResult` 交给 `session.ChangeLog.AddFromStep`，没有新增第二套 Loop、Broker 或执行入口。

### 1.2 存储：只有一个权威落点

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 迁移 018 | `ShareRibbon/Storage/Migrations/018_agent_run_change_restore.sql` + `OfficeAiDatabase` 内联字典与 `EnsureAgentRunTraceSchema` 的 `EnsureColumn` | `agent_run_change.restore_json TEXT`；老库升级走文件、新库走建表、缺列走自愈，三条路同一形状 |
| 读写 | `ShareRibbon/Storage/AgentChangeLogRepository.vb`（`SaveRun`/`GetByRun`/`GetRestoreSnapshotsByRun`/`GetUndoScopeByRun`） | 快照只落 `restore_json` 列；`GetRestoreSnapshotsByRun` 按写入顺序返回可重放节点，坏 JSON 跳过（少撤一步比撤错一步好）；`GetUndoScopeByRun` 的两路判定与内存合同一致，拒绝时带用户可读原因 |
| 快照不进观察 | 同上，`ToJson()` 只报 `hasRestore` | 单元测试直接断言快照正文不出现在变更记录 JSON 里 |

### 1.3 Excel 宿主

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 采集 | `ExcelAi/Undo/ExcelRunRestore.vb`（`kind = excel_range_state`） | 执行前抓目标区域 `Value2`/`Formula`/`NumberFormatLocal`（≤4000 单元格）并算 SHA256 签名；无法忠实映射的形状（非二维返回）直接放弃快照 |
| 结构闸门 | `ExcelAi/ChatControl.vb` `IsExcelStructureStable` | 工作簿/工作表/目标地址相同且 `usedRows`/`usedColumns`/`chartCount` 相同才算结构没动；插删行列、新建图表不产快照 |
| 重放 | `ExcelRunRestore.Replay` | 返回值空串才是“已还原并核对通过”；拒绝情形：缺工作表/地址、无活动工作簿、活动工作簿不是记录那个、找不到工作表、区域形状变了、当前签名既不等于执行前也不等于执行后（用户又改过）、快照内容不完整、写回后仍不一致 |
| 写回顺序 | 同上 | 数字格式 → `Value2` 整块 → 仅对以 `=` 开头的单元格逐格恢复公式；非公式单元格不回写 `Formula` 文本，否则 `"1-1"` 会被 Excel 重新解释成日期 |
| 一键撤销 | `ExcelAi/ChatControl.vb` `PerformUndoAgentRun` | 先查 `FullyUndoable`，再要求快照条数 == 残留改动条数；`ScreenUpdating` 关闭下按写入逆序重放，任一拒绝就如实报“已还原 N 处、剩余 M 处未还原”且**不**标记本轮已回滚，用户解决冲突后还能再点一次 |
| 异常路径 | `BuildExcelCrashRestore` + `BuildExcelRollbackNote(partialApply, hasRestore)` | 命令写一半抛异常时仍尝试补一份快照；观察里的 `rollback` 说明现在区分“可用「撤销本轮」还原”和“只能人工核对” |

### 1.4 PowerPoint 宿主

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 采集 | `PowerPointAi/Undo/PowerPointRunRestore.vb`（`kind = powerpoint_slide_text`） | 执行前按 1-based 序号记录目标页每个带文本框形状的形状名与文字，外加备注页文字，算 SHA256 签名；任一形状或备注读不到就整体作废（不留半份快照），文本上限 20000 字符 |
| 通道闸门 | `PowerPointAi/ChatControl.vb` `IsPowerPointTextOnlyChange` | 演示文稿名/页码/页数/形状数/动画数/格式签名全部相同、且文字或备注确实变了才产快照；格式、动画、形状增删类改动**不**声明可还原，因为写回 `Text` 还原不了它们 |
| 重放 | `PowerPointRunRestore.Replay` | 拒绝情形：演示文稿不是记录那个、页已被删掉、形状数量变了、按序号取到的形状名对不上、形状已无文本框、当前签名既不等于执行前也不等于执行后；写回时文字本就一致的形状跳过，尽量少动格式；备注只在执行前确有备注形状时才写，避免凭空造出备注占位符 |
| 一键撤销 | `PowerPointAi/ChatControl.vb` `PerformUndoAgentRun` | 与 Excel 同构：校验快照覆盖全部残留改动、逆序重放、如实汇报已还原/未还原、只在全部成功后 `MarkRunRolledBack` |
| 异常路径 | `BuildPowerPointCrashRestore` + `BuildPowerPointRollbackNote(partialApply, hasRestore)` | 半应用失败时同样先尝试补快照，再决定 `userMessage` 说“可以撤销本轮”还是“无法自动撤销” |

### 1.5 UI 与三端一致

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 去掉 Word 专属闸门 | `ShareRibbon/Controls/Services/AgentKernelService.vb` `BuildTerminalDescriptor` | 终态描述不再以宿主类型判断能否撤销，只看合同事实：写入数为 0 / `Not FullyUndoable` / 无 runId 三种情况各给一条 `undoReason`；成功时下发 `units` 与新增的 `restoreCount` |
| 按钮文案 | `ShareRibbon/Resources/js/agent-card.js` | `units > 0` 显示“撤销本轮（N 个撤销单元）”，`restoreCount > 0` 显示“撤销本轮（按快照还原 N 处）”，两者都不承诺“宿主撤销”，也不把两种机制混成一句话 |

## 2. E4-03 全局快捷键 / 命令面板

**验收标准**：任一宿主内可唤起；支持自然语言与命令。

关键决定是**不把它做成三端各一份的入口**。快捷键轮询、组合键合同、面板状态与话术下发全部落在 `ShareRibbon`，由 `BaseChatControl` 惰性持有一个 `CommandPaletteService`；三个宿主各贡献且只贡献一个覆写 `EnsureHostTaskPaneVisible` → `Globals.ThisAddIn.ShowChatTaskPane()`。于是「任一宿主内可唤起」是结构保证，而不是靠人记得在第三个插件里再接一次；设置改完立即生效，不需要重启宿主。

### 2.1 落点

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 组合键合同 | `ShareRibbon/Controls/Services/HostHotkeyWatcher.vb`（`HostHotkeyChord`） | `Parse` 支持 `+`/`-` 分隔、ctrl/alt/shift 别名、a-z、0-9、F1–F12、命名键与 OEM 标点，拒绝两个主键；`IsSafe = 有主键且带 Ctrl 或 Alt`——裸键会吃掉用户正常输入，Word 本身已占用 `Ctrl+K` |
| 唤起方式 | 同上（`HostHotkeyWatcher`） | WinForms `Timer`（80ms）轮询 `GetAsyncKeyState`，只认「未按下→按下」这条边沿；修饰键按**精确集合**比对（多按一个 Shift 不算命中）；沿用 `WordCompletionManager` 已有的轮询先例 |
| 前台判定 | 同上（`IsHostForeground`） | `GetForegroundWindow` + `GetWindowThreadProcessId` 与 `_hostProcessId` 比对，只有处于前台的 Office 进程响应。Word 和 Excel 同时开着、共用一个组合键时不会互相抢键——这是全局键盘状态必须付的代价 |
| 重入保护 | 同上（`OnPollTick`） | 回调期间停表（回调要开面板、走 COM、可能进模态），结束后仅在「没被 Dispose、没被关掉、还在监听」三个条件同时成立时才重启，所以关掉开关不会在同一个按键上被重新 Start |
| 面板状态与快捷键注册 | `ShareRibbon/Controls/Services/CommandPaletteService.vb` | `ApplyHotkey` 先停再注册（组合键刚改过时旧键不能继续吃这一次唤起）；注册失败用 `GlobalStatusStrip.ShowWarning` 说明原因，而不是留一个按了没反应的开关。`PushState` 记录页面就绪，`Open()` 早于就绪时**排队**而不是丢按键 |
| 监听生命周期 | `ShareRibbon/Controls/BaseChatControl.vb`（`CommandPaletteSvc` + `Disposed` 接线） | 服务随面板惰性创建，面板销毁时 `Shutdown()` 停表；关掉插件进程之后快捷键继续生效是不可接受的 |
| 面板 UI | `ShareRibbon/Resources/js/command-palette.js`、`chat-template-refactored.html`、`css/styles.css` | 顶栏新增 `#command-palette-button`（快捷键没开也能用），overlay + 输入框 + 结果列表；`↑↓` 选择、`Enter` 执行、`Esc` 关闭；列表 = `QuickActionCatalog` 过滤项 + 兜底「按原话执行：<用户输入>」，所以自然语言和命令走同一条出口 |
| 唯一出口 | `command-paletteRun` | 只发 `sendMessageToServer({ type: 'startAgent', request })`。服务代码里不存在 `startAgent`/`sendMessageToServer`，有反向守卫；`router.Register("openCommandPalette")` 一度加上后又删除——面板按钮本地就能开，多一条路等于多一套终态 |
| 设置 | `ShareRibbon/Config/ChatSettings.vb`（`EnableHotkey`、`DefaultHotkeyChord`、`HotkeyChord`） | 三个宿主共用同一个 `office_ai_chat_settings.json`（按应用区分的文件名早已被注释掉），因此组合键天然是全局的；`LoadSettings` 读回时再校验一次并归一化，非法写法直接回落默认值且关闭开关 |
| 面板设置 UI | `chat-template-refactored.html`、`settings-manager.js` | 补齐 E4-01/E4-02/E4-03 三个开关（选区工具条、右键菜单、快捷键）+ 组合键输入框；`settingsSave()` 改为先组一个 `payload` 再补 `type`，字段清单不再抄两遍 |

三道拒绝点：`ChatSettings.LoadSettings` 归一化、`HandleSaveSettings` 校验（无效即保持关闭并在状态栏说明）、`HostHotkeyWatcher.Start` 返回 False → `GlobalStatusStrip.ShowWarning`。刻意禁止 `SetWindowsHookEx`、`RegisterHotKey`、`keybd_event`、`SendInput`——前两个是全局钩子/系统级占用，后两个是伪造按键。

### 2.2 顺带修掉的两个持久化缺陷

1. `ChatSettings.SaveSettings` 此前用固定 `Dictionary` 重建整个 JSON，凡是表里没有的键都会在每次保存时被静默删除（已确认 `enableMcpList` 就是这样被吃掉的）。现在改成读改写同一个 `JObject`，文件不是合法 JSON 时才重建。
2. `HandleSaveSettings` 此前把 `EnableSelectionToolbar`/`EnableContextMenu` 从静态字段原样写回，等价于这两个开关是只读的——**这正是上一轮被驳回的方案里那条「默认开启却没有 UI 开关」**。现在一律从载荷读回（`ReadFlagSetting`，缺键时保持当前值，旧版前端不会把开关洗掉）。这两个缺陷都有反向守卫。

E4-04 把选区自动引用改成默认开启，前提就是这轮补齐了可见开关：能在面板里关掉的能力才谈得上默认开。该前提已在 §3 兑现。

### 2.3 门禁补强

| 脚本 | 新增断言 |
| --- | --- |
| `scripts/smoke-command-palette.ps1`（新增，已登记进 `run-code-checks.ps1`） | 组合键解析/归一化用例（`Ctrl+Alt+K`、`ctrl+k`→`Ctrl+K`、`Control+Shift+F12`、`Alt+/`、`Ctrl+.`）；拒绝用例（`K`、`Space`、`Ctrl+`、空串、空格串、`Ctrl+Shift+A+B`）并单独断言 `Shift+A`「解析得过但不安全」；监听器生命周期（坏组合键 `Start=False`、正常 `True`、重复 `Start=False`、`StopWatching`、二次 `Dispose` 不抛）；用反射真实实例驱动 `CommandPaletteService`：页面就绪前 `Open()` 一条都不发、`PushState` 恰好一次 `syncCommandPalette(` + 一次 `showCommandPalette(`、载荷含 `appType='Word'` 与 3 条动作、`ensurePaneVisible` 恰好被调 1 次、就绪后直接下发；默认值（关闭 + `Ctrl+Alt+K`）；持久化四个键与读改写、载荷回写、拒绝条件；话术单一来源与服务代码内**不得**出现 `startAgent`/`sendMessageToServer`；JS 行为模式与禁止 `fetch`/`XMLHttpRequest`/`invokeMethod`；资源四处登记（resx / `ResourceExtractor` / HTML / vbproj）；反向守卫（禁止 `SetWindowsHookEx`、`RegisterHotKey`、`keybd_event`、`SendInput`，必须有前台守卫、精确修饰键、回调停表后重启、`RemoveHandler _timer.Tick`）；500 行上限 |
| `scripts/run-code-checks.ps1` | 登记 `smoke-command-palette.ps1`；`node --check` 名单加入 `command-palette.js` 与 `settings-manager.js` |

```powershell
cmd /c build-code.bat
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run-code-checks.ps1 -SkipBuild
```

4 个代码项目构建与全部门禁（含 `Command palette smoke passed.`）已通过。两个新 `.vb` 与 `command-palette.js` 均已登记进 `ShareRibbon.vbproj` / `Resources.resx` / `ResourceExtractor` / HTML。

### 2.4 声明的不作为

1. **不做系统级全局快捷键**：快捷键只在对应 Office 进程处于前台时生效。跨应用抢键、Office 在后台也能唤起，属于钩子注入，风险与收益不成比例。
2. **不做命令语法解析**：面板只把「选中的那句话」交回主链，由 `AgentKernel`/`LoopEngine` 判断这是自然语言还是可执行指令；入口层不解析 `:` 前缀、不做斜杠命令表。
3. **不做搜索历史/ frecency 排序**：结果就是 `QuickActionCatalog` 的固定顺序 + 当前输入过滤。加排序需要在共享层持久化用户输入，收益低于隐私成本。
4. **Excel/PPT 动作仍不声明 `CapabilityId`**（开放待办 #36），命令面板继承同一口径。

### 2.5 真机验收清单（按用户指示暂缓，需要 Word 2021/Excel/PowerPoint + 重启宿主刷新 `%LOCALAPPDATA%\OfficeAI\www`）

1. 三端各按一次默认 `Ctrl+Alt+K`：任务窗格收起时也要先展开再出面板；面板输入框拿到焦点。
2. 面板里改组合键为 `Ctrl+Alt+J` 并保存，**不重启宿主**立即用新键唤起；旧键 `Ctrl+Alt+K` 不再生效。
3. 关掉「命令面板快捷键」开关并保存后，快捷键立即失效（`StopWatching` 生效），但顶栏按钮仍可唤起面板。
4. 输入 `K`、`Ctrl+`、`Ctrl+Shift+A+B` 三种写法保存：都应当被拒绝并在状态栏给出原因，开关保持关闭，文件里不会留下坏组合键。
5. 长按组合键、输入法候选期间重复按下：只唤起一次，不出现面板反复开关。
6. Word 与 Excel 同时打开、只有 Word 在前台时按键：只 Word 出面板。
7. 选一条目录项回车：进入与 Ribbon/选区工具条完全相同的链路——模型配置闸门、计划卡、风险确认、变更摘要都在；未配置 Key 时被拦住而不是静默失败。
8. 直接输入一句自然语言按 Enter：出现「按原话执行：…」并原话交给主链；空输入不发送。
9. `↑↓` 环绕、`Esc` 关闭、再次按键重新打开时光标回到第一项且输入被清空。
10. 关闭面板（进程内销毁控件）后快捷键不再响应。
11. 隐私：日志与 `agent_run_*` 表里不出现用户输入的原文之外的额外内容；面板输入不写聊天历史之外的存储。
12. 前端副本：`%LOCALAPPDATA%\OfficeAI\www\js\command-palette.js` 时间戳随重启刷新，面板顶栏出现命令按钮（旧副本会让整个注入块静默失效）。

### 2.6 已声明但推迟的缺陷（属于 E4-05 事件绑定卫生）

这两条已在 §4 兑现，此处保留原始描述作为缺陷出处。

1. 三宿主在 `ChatTaskPane_VisibleChanged` 里读 `ChatSettings.EnableContextMenu` 时，可能还没有任何 `ChatSettings` 实例加载过文件，因此**进程内第一次打开面板**右键菜单可能不挂。修法 belongs to E4-05（入口初始化前先确保配置已加载）。
2. 原先 `ChatBrowser.CoreWebView2.NavigateToString` 之后还 `AddHandler ChatBrowser.CoreWebView2.NavigationCompleted`，而处理器早已用 `Handles ChatBrowser.NavigationCompleted` 绑过一次；`OnWebViewNavigationCompleted` 顶部那句 `RemoveHandler` 摘的是**另一个事件**上的绑定，摘不掉声明式那份，于是 `InitializeSettings`——以及 E4-03 接上去的 `PushState`/`ApplyHotkey`——每次导航都会重跑。当前表现是重复同步而非错误，但它是 E4-05 的「不重复绑定」必查项。

---

## 3. E4-04 选区自动引用默认开启

**验收标准**：默认携带选区上下文；UI 明确显示引用了什么。

这件事的难点不在「把默认值改成 True」，而在改完之后仍然成立的三件事：开关必须真的能关掉（三端九个实现共用一个决策点）、带出去了什么必须让用户看得见（Agent 回合本来不渲染用户气泡）、摘要只能是显示而不能顺手变成第 N 份持久化数据。

### 3.1 落点

| 位置 | 内容 |
| --- | --- |
| 决策点 | [BaseChatControl.vb:2750](file:///f:/ai/code/AiHelper/ShareRibbon/Controls/BaseChatControl.vb#L2750) `BuildMessageWithAutoSelection`：`If Not selectedCellChecked Then Return message` 之后才调用 `AppendCurrentSelectedContent` |
| 契约 | `AppendCurrentSelectedContent`（MustOverride，三端 × ChatControl/DoubaoChat/DeepseekControl 共 9 个实现）从此只回答「怎么把选区写成提示」，不再回答「要不要写」；PowerPoint 里那份重复的 `If Not selectedCellChecked` 已删除 |
| 两条发送路径 | 聊天回合与 Agent 回合都改走同一个包装方法（`question = …` / `finalMessageToLLM = …`），共享层里除包装方法本身没有任何其它调用点 |
| 摘要 | `DescribeAutoSelection`：从**真正附加上的那段增量**（`appended.Substring(message.Length)`）生成，压成一行、预览上限 160 字符、只报「约 N 字」，不复制正文 |
| 显示通道 | `RenderAutoSelectionReference` 只在 Agent 回合把摘要发给前端；`chat-manager.js` 的 `window.appendTurnSelectionReference` 复用既有 `.chat-message-references` 样式，纯 `textContent`，同回合重复同步先删旧条 |
| 默认值 | `ChatSettings.selectedCellChecked = True` 与控件字段同值；`LoadSettings` 仍是 `ContainsKey` 才覆盖，所以显式存过 `false` 的用户保持关闭 |
| 面板 | 设置行改名「自动引用当前选区」，说明里写清代价：关掉后不再采集选区，也不再按选区定位写回 |

摘要坚持从增量反推，而不是改用 `CaptureOfficeContext` / `ContextPack.FromOfficeContext`（`Selection.Ref/Address/ItemCount/Preview`）再生成一份标签。后者看着更「结构化」，但会把选区第二次塞进提示，并且每回合多一次 COM 读取——这两条都比它省下的几行代码贵。

**明确接受的耦合**：Word 的 `AppendCurrentSelectedContent` 同时写 `PendingSelectionInfo`，共享层再把它绑到本轮 `requestUuid`（`ChatStateService._selectionPendingMap` → `MigrateSelectionToResponse` → `_responseSelectionMap`），选区对比与写回依赖这条链。开关关掉后这条链不再记录快照——这不是漏，而是「不引用也就不写回」的同一条语义，已经写进设置项说明和门禁。

顺带的行为收敛：Doubao/Deepseek 两个控件此前没有自己的闸门，等于不受该开关约束；收口到共享包装层之后，三端所有控件的自动引用都由同一个开关决定。

### 3.2 顺带修掉的注入缺陷

`InitializeSettings` 注入的脚本里有两行指向不存在的元素（`#scrollChecked`、`#selectedCell`；面板里实际是 `#settings-scroll-checked`、`#settings-selected-cell`）。WebView2 注入块是顺序执行的，`getElementById(...)` 返回 null 后取 `.checked` 抛 TypeError，**该行之后的每一行都不再执行**——也就是 E4-01/E4-02/E4-03 那三个原生入口开关和命令面板按钮的同步一直是静默失效的。修法不是补 id，而是删掉这两行：它们想设的控件在旧版面板上就存在过，现在的对应行本来就在别处。为了让这类缺陷不再复发，门禁新增「注入脚本引用到的每个 id 都必须在面板 HTML 中存在」的通用守卫，并把这两个 id 单独钉住。

### 3.3 门禁补强

| 脚本 | 新增断言 |
| --- | --- |
| `scripts/smoke-selection-reference.ps1`（新增，已登记进 `run-code-checks.ps1`） | 反射读取编译后的 `ChatSettings.selectedCellChecked` 必须为 True（在实例化之前读，取的是声明的默认值而非本机 profile）＋两处声明与 `LoadSettings` 的 `ContainsKey` 回读；包装层是唯一决策点（`AppendCurrentSelectedContent(` 在共享层只允许「MustOverride 声明 + 包装内一次调用」两处、`BuildMessageWithAutoSelection(` 恰好声明 + 两条发送路径三处）；九个宿主 override 的函数体一律**不得**再出现 `selectedCellChecked`；Word 侧仍写 `PendingSelectionInfo`（把耦合钉住而不是隐藏）；摘要必须由增量反推、`Char.IsWhiteSpace` 折行 + `>= 160` 截断；渲染只在 Agent 回合、两个参数都走 `ToJsStringLiteral`、摘要为空直接 Return；`_autoSelectionReference` 所在行不得出现 `AppLogger`/`AddMessage`/`History`/`Memory`/`Observation`；JS 通道缺元素返回 false、复用 `chat-message-references`、只 `textContent`、禁止 `innerHTML`/`fetch`/`XMLHttpRequest`/`invokeMethod`；注入块引用的 14 个 id 全部存在于 HTML，且 `scrollChecked`/`selectedCell` 两个死 id 不得回归；开关可即时切换（`HandleCheckedChange` 的 `"selectedCell"` 分支）与保存（`jsonDoc("selectedCell")`、`settings-manager.js` 载荷） |
| `scripts/run-code-checks.ps1` | 登记 `smoke-selection-reference.ps1`（`chat-manager.js` 早在 `node --check` 名单内） |

```powershell
cmd /c build-code.bat
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run-code-checks.ps1 -SkipBuild
```

4 个代码项目构建与全部门禁（含 `Selection reference smoke passed.`）已通过。本轮未新增 `.vb`/`.js` 文件，因此不涉及 `.vbproj` 与资源四处登记；新增的是一个门禁脚本，已登记。

### 3.4 声明的不作为

1. **不引入第二套引用展示**：聊天路径本来就把引用渲染成 `引用内容 (N)` 折叠块（`message-sender.js` + `window.selectedContentMap`），本轮只给不渲染用户气泡的 Agent 回合补一条通道，没有把两条合成一套。
2. **不做引用条的目的推断/置信度**：摘要只说明「带出去了哪一段、多大」，不声称模型会怎么用它；意图仍交给 `AgentKernel`/`LoopEngine`。
3. **不改采集逻辑本身**：三端 `AppendCurrentSelectedContent` 的选区读取、Excel 的 100×26 上限、Word 的原文直取都原样保留，本轮只收口「要不要附加」。
4. **不为摘要新增持久化字段**：它只在内存里活一回合，不写历史、不写日志、不进 Observation/Artifacts。
5. **`DescribeAutoSelection` 不做反射行为测试**：它挂在 `MustInherit` 的 WinForms 控件类型上，PowerShell 侧加载该类型会连带解析 WebView2/interop 面，反射不可用；因此截断与折行按源码不变量守卫，默认值则用可反射的 `ChatSettings` 静态属性做真实断言。

### 3.5 真机验收清单（按用户指示暂缓，需要 Word 2021/Excel/PowerPoint + 重启宿主刷新 `%LOCALAPPDATA%\OfficeAI\www`）

1. 全新配置（`office_ai_chat_settings.json` 无 `selectedCellChecked` 键）首次启动：设置面板里「自动引用当前选区」是勾选状态，且提问确实带上选区。
2. Word：选中一段文字→在面板提问 → Agent 回合出现「本轮已引用选区」折叠条，展开只有一行摘要（约 N 字 + 预览），点折叠/展开不报错。
3. Word：无选区（光标停留）提问 → 不出现引用条，也不出现空壳。
4. Excel：选中区域 → 摘要显示「约 N 字」，正文仍是表格化内容；选区很大时摘要不会变成第二份正文。
5. PowerPoint：选中文本框/形状 → 引用条出现，且关掉开关后不再采集。
6. 取消勾选并保存 → 立即生效，不需重启宿主；再提问时提示里不含选区，Word 的选区写回/对比也不再定位（这是声明的语义，不是故障）。
7. 显式存过 `false` 的旧配置：升级后保持关闭，默认值不得把用户已经关掉的采集再打开。
8. 聊天路径（非 Agent）回归：引用卡片仍按原样式显示，没有同时出现两条引用。
9. replan / 同回合重复同步：引用条只有一条，不堆叠。
10. 隐私：日志、`agent_run_*`、聊天历史里查不到这条摘要文本；摘要含文档原文时也不会被持久化。
11. 注入完整性（§3.2 的回归）：设置面板里三个原生入口开关与快捷键输入框都能反映真实配置；改一个保存后其余项没被洗掉。
12. 前端副本：`%LOCALAPPDATA%\OfficeAI\www\js\chat-manager.js` 时间戳随宿主重启刷新，否则引用条函数不存在（表现为「什么都不显示」而不是报错）。

---

## 4. E4-05 事件绑定卫生

**验收标准**：不重复绑定；重复打开/关闭不产生双触发。

§2.6 留下两条待修，§3.2 又暴露了一类更难发现的东西：重复执行在这个面板上的表现不是报错，而是「看起来没有副作用」。所以 E4-05 的价值不在于新增机制，而在于把「一份文档装载 → 一次初始化」「一次绑定 → 一次解绑」变成能被门禁钉住的形状。

### 4.1 落点

| 位置 | 内容 |
| --- | --- |
| 装载闸门 | [BaseChatControl.vb:36](file:///f:/ai/code/AiHelper/ShareRibbon/Controls/BaseChatControl.vb#L36) 新增 `_pageBootstrapDone`；处理器先 `If Not e.IsSuccess Then Return`、再 `If _pageBootstrapDone Then Return`，并**在第一个 `Await` 之前**占位（[718-719](file:///f:/ai/code/AiHelper/ShareRibbon/Controls/BaseChatControl.vb#L718)）；只有真正装载新文档的地方复位（[:566](file:///f:/ai/code/AiHelper/ShareRibbon/Controls/BaseChatControl.vb#L566)） |
| 绑定唯一化 | 删掉 `AddHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted` 和那句摘错事件的 `RemoveHandler`，只保留声明式 `Handles ChatBrowser.NavigationCompleted` 一份；[:569-572](file:///f:/ai/code/AiHelper/ShareRibbon/Controls/BaseChatControl.vb#L569) 留下「为什么不能再补回来」的注释 |
| 一次装载一次初始化 | `InitializeSettings()` 与 `InitializeMcpSettings()` 每份文档各一次；Excel 的 `MainForm_Load` 不再自己调 `InitializeSettings()`（[ExcelAi/ChatControl.vb:104](file:///f:/ai/code/AiHelper/ExcelAi/ChatControl.vb#L104)），Word/PPT 本来就没有，三端口径一致 |
| 配置先行加载 | [ChatSettings.vb:54](file:///f:/ai/code/AiHelper/ShareRibbon/Config/ChatSettings.vb#L54) 新增 `EnsureLoaded(applicationInfo)`：进程内第一次才真正读盘，重复调用返回同一实例；三宿主在 `ChatTaskPane_VisibleChanged` 里、`EnsureSelectionToolbar()`/`EnsureContextMenu()` **之前**各调一次（[WordAi/ThisAddIn.vb:214](file:///f:/ai/code/AiHelper/WordAi/ThisAddIn.vb#L214)、[ExcelAi/ThisAddIn.vb:363](file:///f:/ai/code/AiHelper/ExcelAi/ThisAddIn.vb#L363)、[PowerPointAi/ThisAddIn.vb:186](file:///f:/ai/code/AiHelper/PowerPointAi/ThisAddIn.vb#L186)） |

闸门放在「文档」而不是「控件」上：面板一生可能装载多份文档（CoreWebView2 重建后重新 `NavigateToString`），每一份都需要注入一次；面板内的后续导航不该再重跑——重跑会把开关值重复下发、把监听重复注册，而这类问题的表现恰好是「多点一次才看出来」。占位必须在第一个 `Await` 之前，因为 `Task.Delay` + `InvokeAsync` 之后可能又有导航完成排队进来，跨过 await 就变成两个都通过检查。

### 4.2 检查过、本来就卫生的部分（只钉住，不改）

- **右键菜单** `HostContextMenu`：`Attach` 里 `If _attached Then Detach()` 先清后挂、CommandBar `Temporary:=True`、按 `aiq:` Tag 反序删除、按钮 `AddHandler/RemoveHandler button.Click` 成对、`Detach` 里 `_attached = False` + `_buttons.Clear()`。
- **全局快捷键观察器** `HostHotkeyWatcher`：`AddHandler _timer.Tick` / `RemoveHandler _timer.Tick` 成对，面板销毁即停表。
- **命令面板服务**：懒创建 `If _commandPaletteService Is Nothing Then … AddHandler Me.Disposed`，Dispose 绑定只可能挂一次。
- **宽度定时器**：`EnsureWidthTimers` 用 `Is Nothing` 守卫，`ThisAddIn_Shutdown` 统一 `Stop()`+`Dispose()`；`CreateChatTaskPane`/`EnsureDataCapturePaneCreated` 有实例守卫，所以 `AddHandler chatTaskPane.VisibleChanged` 每个面板实例只挂一次。

这几处本轮一行代码都没动。把它们写进门禁而不是注释，是因为「已经对称」这个事实不属于任何一个人的记忆，下一次新增入口时最容易破坏的就是它。

### 4.3 顺带查明的一条不可达路径

`LoadLocalHtmlFile()` → `WebViewBridge.LoadHtmlFileAsync(ChatHtmlFilePath)` 是唯一没有配 `_pageBootstrapDone` 复位的导航调用点，看着像漏。实际不是：`ChatHtmlFilePath` 是**带时间戳的保存目标**（`saved_chat_yyyyMMdd_HHmmss*.html`，[:689-712](file:///f:/ai/code/AiHelper/ShareRibbon/Controls/BaseChatControl.vb#L689)），读取那一刻必然不存在，`LoadHtmlFileAsync` 在 `File.Exists` 处直接 Return、从不导航；宿主侧 `loadChatHtml` 又只在进程内第一次显示面板时调用。所以它是 E5-04 的候选清理项，而不是绑定卫生问题——本轮不为它写一条永不执行的复位分支。

### 4.4 门禁补强

| 脚本 | 新增断言 |
| --- | --- |
| `scripts/smoke-event-binding.ps1`（新增，已登记进 `run-code-checks.ps1`） | 反射调用 `ShareRibbon.ChatSettings.EnsureLoaded` 两次必须返回同一实例（`[object]::ReferenceEquals`）＋源码里存在 `If _loadedInstance IsNot Nothing Then Return`；禁止 `AddHandler`/`RemoveHandler ChatBrowser.CoreWebView2.NavigationCompleted` 回归、`OnWebViewNavigationCompleted` 只允许出现一次、必须有 `Handles ChatBrowser.NavigationCompleted`；`_pageBootstrapDone = False` 必须早于 `NavigateToString(htmlContent)`，处理器内部顺序必须是「判 → 占 → 第一个 `Await`」，标志位代码出现次数恒为 4；`AddHandler ChatBrowser.WebMessageReceived` 恰好一次；`AddHandler Me.Disposed` 必须在 `Is Nothing` 守卫内；三宿主 `MainForm_Load` 体不得出现 `InitializeSettings()`、共享层裸调用只允许 1 处；三宿主 `ChatTaskPane_VisibleChanged` 必须在 `EnsureContextMenu()` 之前调用 `ChatSettings.EnsureLoaded(New ApplicationInfo("<宿主>", OfficeApplicationType.<宿主>))`；宽度定时器 AddHandler 必须在 `Is Nothing` 守卫内、`ThisAddIn_Shutdown` 必须 Stop+Dispose 且 `contextMenu.Detach()`；Word 工具条保留 `EnableSelectionToolbar` 且 `EnsureLoaded` 早于 `EnsureSelectionToolbar`；`If _attached Then Detach()`、`AddHandler/RemoveHandler button.Click` 对称、watcher 的 Tick 增删计数相等 |

```powershell
cmd /c build-code.bat
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run-code-checks.ps1 -SkipBuild
```

4 个代码项目构建与全部门禁（含 `Event binding hygiene smoke passed.`）已通过。本轮未新增 `.vb`/`.js` 文件，只新增一个门禁脚本并登记。

### 4.5 声明的不作为

1. **不给 `Handles` 绑定补手工解绑**：控件与面板随 `ThisAddIn_Shutdown` 一起释放，为此写一套 Dispose 解绑等于凭空制造一份需要维护的状态。
2. **不动 Doubao/Deepseek/爬虫三套控件**：它们各有独立的 `OnWebViewNavigationCompleted`，加载的是远端 URL、不注入 AI 面板设置，本轮的「一次装载一次初始化」合同对它们不适用。
3. **不做配置热重载**：`EnsureLoaded` 只保证「读开关之前文件已被读过一次」；面板自身装载路径仍每次 `New ChatSettings(...)`，外部编辑文件的同步语义没变。
4. **不消除 §4.3 的不可达入口**：那是 E5-04 的范围，且它当前没有行为影响。
5. **不把三处 `EnsureLoaded` 收成一次 Startup 调用**：Startup 阶段读盘会把文件 IO 放进宿主启动关键路径，而这三处正是「第一次真正需要开关值」的时刻；成本由 `_loadedInstance` 缓存兜住。

### 4.6 真机验收清单（按用户指示暂缓，需要 Word 2021/Excel/PowerPoint + 重启宿主刷新 `%LOCALAPPDATA%\OfficeAI\www`）

1. 冷启动**第一次**点开面板：配置里勾过的「自动引用当前选区」、开过的右键菜单/选区工具条当场生效，不需要「第二次打开才对」。
2. 打开→收起→再打开，反复 5 次：一次提问只对应一次设置注入，DevTools 控制台不出现逐次增多的 `applySettings` 类日志。
3. 面板打开期间发生页面内导航（点文档链接、图片预览）后：设置面板各开关仍显示真实值，且不出现二次下发（表现为刚改完又被洗回）。
4. 关闭面板再打开：一次快捷键只出一个命令面板，没有双层浮层。
5. Excel：面板装载只有一次 `InitializeSettings`（此前 `MainForm_Load` 一次 + 导航完成一次）。
6. Word：反复开合面板后，右键菜单里 AI 子菜单只有一份、选区工具条只有一条。
7. 退出宿主（Word/Excel/PPT）：日志无定时器报错，退出不被拖慢。
8. 命令面板：面板销毁后快捷键不再响应（§2.5 第 10 项，本轮验的是绑定生命周期而非功能）。
9. 注入完整性回归：三个原生入口开关 + 快捷键输入框显示真实配置，保存一项不洗掉其它项——双注入消失后，§3.5 第 11 项应当稳定复现。
10. WPS 环境开合面板：只有一份宽度定时器在跑，面板宽度不被反复重置。

---

## 5. E2-04 旧数据迁移/回填（会话归属）

**验收标准**：老会话不丢失；无 `app_type` 的记录有明确归属策略。

E2-01 把历史列表改成按宿主过滤之后，`app_type` 实际上只剩两种合法值：三个规范宿主名，或者空串（通用数据，三端都可见）。第三种值既不等于任何宿主、又拿不到空串的豁免，于是那条数据在 Word/Excel/PowerPoint 里同时查不到——用户看到的现象就是「升级后历史整段消失」，而且没有任何报错。这个哨兵值确实进过库：`GetOfficeAppType()` 的基类默认实现至今仍是 `Return "Unknown"`（[BaseChatControl.vb:2801](file:///f:/ai/code/AiHelper/ShareRibbon/Controls/BaseChatControl.vb#L2801)），三端子类各自覆盖成规范名（[WordAi/ChatControl.vb:1998](file:///f:/ai/code/AiHelper/WordAi/ChatControl.vb#L1998)、[ExcelAi/ChatControl.vb:389](file:///f:/ai/code/AiHelper/ExcelAi/ChatControl.vb#L389)、[PowerPointAi/ChatControl.vb:402](file:///f:/ai/code/AiHelper/PowerPointAi/ChatControl.vb#L402)）。

本轮先量了本机两个库，结论直接决定了做法（都是 `Read Only=True` 探针，没有跑迁移、没有改库）：

| 库 | schema | 归属现状 |
| --- | --- | --- |
| 安装版 `Documents\OfficeAiAppData\office_ai.db` | 17 | 281 条消息 = 269 条空串 + 12 条 `Word`；27 条会话档案 `app_type` 全空，只有 2 条有文档名（`report.docx`、`文档1`）；3 条 `agent_run` 全是 `Word`，`agent_run_change` 0 条；81 个会话里 0 个能连到轨迹 |
| 同一库的 `atomic_memory` | — | **231 行里 123 行是 `Unknown`**；`memory_item` 24 行全在白名单内 |
| 调试版 `OfficeAiAppData-Debug` | 19（本轮门禁跑出来的） | 所有表 0 行，没有可回填的数据 |

那 123 行是本轮改变计划的东西：记忆链的读侧过滤（[MemoryRepository.vb:137](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/MemoryRepository.vb#L137) 等）用的是同一条「等于当前宿主 OR 空 OR NULL」规则，所以同一批哨兵值让老记忆也在三端同时失效。历史表看起来没事、记忆表病得很重——只修历史侧就是给同一个缺陷族留了一半。于是 E2-04 的口径落在**列**上而不是落在**会话列表**上：凡是带 `app_type` 的六张表（`conversation_event` / `agent_run` / `agent_run_change` / `session_summary` / `atomic_memory` / `memory_item`）一起收口。

### 5.1 归属策略（写成一个可校验的顺序）

迁移 019（[019_session_ownership_normalization.sql](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/Migrations/019_session_ownership_normalization.sql)，180 行）按证据强度从高到低填，**每一步只填空、不覆盖已经解析出来的归属**，因此整段可以重复执行：

| 步 | 证据 | 说明 |
| --- | --- | --- |
| 1 | 归一 | 六张表里白名单之外的任何值（含大小写与首尾空白变体）先压成规范名或空串（[:23](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/Migrations/019_session_ownership_normalization.sql#L23)）。必须先做：一条「最新记录是 `Unknown`、更早记录知道宿主」的数据，归一之后才能在第 2 步落到那个宿主 |
| 2 | `agent_run` 该会话最近一条带宿主的轨迹 | 与迁移 14 同一条规则（[:87](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/Migrations/019_session_ownership_normalization.sql#L87)） |
| 3 | `agent_run_change` 最近一条（按 `observed_at`） | 覆盖只有变更日志、没有轨迹头的老会话（[:100](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/Migrations/019_session_ownership_normalization.sql#L100)） |
| 4 | `session_summary.document_name` 扩展名分组 | Word `.doc/.docx/.docm/.dot/.dotx/.rtf/.wps`；Excel `.xls/.xlsx/.xlsm/.xlt/.et`；PowerPoint `.ppt/.pptx/.pptm/.pps/.ppsx/.dps`（[:112](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/Migrations/019_session_ownership_normalization.sql#L112)）。`.csv` 故意不推断（Excel 能打开但不独占），未保存文档名（本机就有 `文档1`）也不推断 |
| 5 | 档案行 → 同一会话仍空的事件行 | [:146](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/Migrations/019_session_ownership_normalization.sql#L146) |
| 6 | 事件行 → 仍空的档案行 | [:162](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/Migrations/019_session_ownership_normalization.sql#L162)。E2-05 分组读的是 `session_summary`，不反向补齐的话老会话在 UI 上会永远显示成「通用会话」 |
| 兜底 | 没有任何证据 | 保持空串，继续三端可见。老数据宁可多端可见，也不许在升级后凭空消失 |

### 5.2 写边界：口径只有一份

- 新增 [AppTypeContract.vb](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/AppTypeContract.vb)：三个规范名常量 + `Normalize`（[:22](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/AppTypeContract.vb#L22)）。它单独成文件而不是塞进 `OfficeAiDatabase.vb`，是因为后者已经 1000+ 行，而这份口径正是六处写入要共用的东西。
- 六处 `app_type` 绑定全部改为先归一：[ConversationRepository.vb:31](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/ConversationRepository.vb#L31)、[AgentMemoryRepository.vb:147](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/AgentMemoryRepository.vb#L147)（conversation_event）与 [:247](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/AgentMemoryRepository.vb#L247)（memory_item）、[MemoryRepository.vb:506](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/MemoryRepository.vb#L506)（atomic_memory）、[AgentChangeLogRepository.vb:34](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/AgentChangeLogRepository.vb#L34)、[RunTraceStore.vb:46](file:///f:/ai/code/AiHelper/ShareRibbon/Agent/Harness/RunTraceStore.vb#L46)。读侧的过滤参数没动——那里 `''` 本来就是豁免，归一读取值反而会把「按宿主精确查」变成「查全部」。
- 内联兜底同步扩到六张表（[OfficeAiDatabase.vb:252-258](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/OfficeAiDatabase.vb#L252)，字典登记在 [:613](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/OfficeAiDatabase.vb#L613)），不带 Migrations 目录的安装也会跑同一段。
- 选择在写边界归一，而不是放宽读过滤：读条件一旦变成「匹配不上也算」，别的宿主的历史就会漏进当前端。

### 5.3 本机数据上它会改变什么

按上面那张表算：1 个会话按 `.docx` 落到 Word（档案行与 2 条事件都跟着改）、3 个会话的档案行被事件反向补齐、123 条 `Unknown` 记忆重新可召回；81 个会话里 76 个完全没有任何宿主证据，加 1 个只有 `文档1` 这种未保存文档名，共 77 个继续空串三端可见。历史侧本机白名单外是 0 行——`Unknown` 只长在记忆表里，这恰好说明「按列收口」比「按会话列表收口」对。

### 5.4 门禁补强

| 脚本 | 新增断言 |
| --- | --- |
| `scripts/smoke-empty-db-initialization.ps1` | schema 断言 18→19；四类 legacy 会话夹具（扩展名定宿主、新哨兵盖住旧真实宿主、完全无证据、已正确归属不得重判）+ 记忆表夹具（`atomic_memory` 的 `Unknown` 与 `word ` 变体、`memory_item` 的 `Unknown` 与规范值）；断言六张表白名单外 0 行、记忆行必须能被真实的 `MemoryRepository.ListAtomicMemories` 在三端各召回一次、可见性走真实的 `ConversationRepository.GetRecentSessions`、最后**重跑已部署的 019 文件**再断言归属不变（同时证明文件与内联同义、可重复执行） |
| `scripts/audit-p0-guardrails.ps1` | `AppTypeContract.Normalize` 必须存在且六处写入文件都真的引用它；`AgentMemoryRepository` 两个写入点各自都要归一（计数 ≥2）；`InsertMessage` 必须写归一后的值；019 文件与内联兜底必须**逐表**都含 `UPDATE <表> SET app_type = ` 归一语句（六张表逐一比对，防止只修文件不修兜底）；必须登记 `{19, …}` |
| `scripts/smoke-change-log-and-session-scope.ps1` | 写边界：`InsertMessage(..., "Unknown ")` 之后必须能在三端历史里查到该会话，且 `conversation_event` 与 `session_summary` 落库值必须是空串（输出新增 `UnknownAppTypeNormalizedAtWrite`） |
| `scripts/smoke-db-schema-drift.ps1` | 019 是纯 UPDATE，快照无需变更；该脚本「执行全部 `Migrations\*.sql` vs 快照」仍为 `Drift: False`，顺带证明 019 文本本身可执行 |

```powershell
cmd /c build-code.bat
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run-code-checks.ps1 -SkipBuild
```

4 个代码项目构建通过，全部门禁通过（`P0 guardrails passed.`、`Drift : False`、空库初始化输出 `SchemaVersion : 19 / SessionOwnershipNormalized : True / MemoryOwnershipNormalized : True`）。本轮新增 1 个 `.vb` 文件，已登记 `ShareRibbon.vbproj`；迁移文件由 `Storage\Migrations\*.sql` 通配符自动部署，确认 `bin\Debug` 与源文件逐字节一致。

### 5.5 声明的不作为

1. **不把 `.csv` 归给 Excel**：Excel 能打开 CSV，但 Word/PPT 也常把它当文本处理，扩展名在这里不足以证明宿主。
2. **不给未保存文档名猜宿主**：`文档1` / `Book1` 这类名字没有扩展名，猜出来的归属比空串更糟——它会让学生会话只在某一个端出现。
3. **不放宽读过滤**：见 §5.2，那会把别的宿主历史漏进当前端。
4. **不改 `GetOfficeAppType()` 的 `"Unknown"` 默认值**：那个默认值的意义是「子类还没实现」。真要收口应该把它变成 `MustOverride`（编译期强制三端各自实现），那是 E5-04 不可达入口清理的范围，本轮不偷偷改成空串——那等于把信号抹掉。
5. **不引入第四种合法值**（如 `Generic`）：空串已经是「三端可见」的既有语义，再加一个同义词会让每条 SQL 都要写两种判断。
6. **不动 legacy `conversation` 表**：它没有 `app_type` 列，读取路径也只剩迁移 012 的回填与级联删除，257 行只是历史残留。
7. **不清 `atomic_memory` 的旧数据**：那 123 行是用户真实记忆，归一成通用数据即可召回，删除属于凭空丢用户内容。

### 5.6 真机验收清单（按用户指示暂缓，需要 Word 2021/Excel/PowerPoint 真机 + 重启宿主刷新 `%LOCALAPPDATA%\OfficeAI\www`）

1. 用安装版旧库（schema 17）启动 Word：升级后 `schema_version` 变 19，历史侧边栏仍能看到升级前的会话，特别是那 269 条空归属会话在 Word/Excel/PowerPoint 三端都可见。
2. `report.docx` 那个会话：升级后应只在 Word 端出现（扩展名证据生效），Excel/PPT 查不到它。
3. `文档1` 那个会话：必须继续三端可见，不得被猜成 Word。
4. 已正确归属的 12 条 Word 消息：升级前后归属不变，第 4/5/6 步不得重判它们。
5. 记忆管理界面：升级前查不到的老记忆（那 123 行）现在应能在任一宿主下查到，且不重复。
6. 重复启动宿主两次：第二次不报错、归属与第一次完全一致（019 只填空、可重跑）。
7. 不带 Migrations 目录的安装（只有 `ShareRibbon.dll`）：内联兜底跑到 version 19，结果与文件版一致。
8. 会话历史按文档分组（E2-05 的输入）：`session_summary` 的 `app_type` 已被补齐，分组标题不再全是「通用会话」。

---

## 6. E2-05 历史 UI 分组 + 空状态引导

**验收标准**（计划行）：按文档/时间分组；空状态有引导。

侧边栏本来就有「本文档 / 全部」切换、一个文档名标签和两条空态文案，缺的不是控件，是**结构**：`scope=all` 的列表是「当前文档整块在前，其余一律按最近活跃倒序」的一维流水，想找回上周那篇文档的会话只能逐行读文件名。这一行最容易走错成「前端再排一次序」——[history-manager.js:198](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/history-manager.js#L198) 那条「顺序由后端决定」的注释就是上一轮为此钉下的：排序键是 `MAX(e.created_at)`，前端只要按时间重排，当前文档在前的合同就没了。所以本轮把分组定成**呈现**，排序仍然是后端的合同，同时把分组真正需要的两个键补进后端 payload。

### 6.1 落点

| 位置 | 改动 |
| --- | --- |
| [ConversationRepository.vb:188](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/ConversationRepository.vb#L188) | `GetRecentSessions` 额外取 `MAX(e.created_at)`，并在 [:210](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/ConversationRepository.vb#L210) 把它映射到 `LastActiveAt`——这一列本来就是 [:195](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/ConversationRepository.vb#L195) 的排序键，以前只是没回传 |
| [MemoryRepository.vb:73](file:///f:/ai/code/AiHelper/ShareRibbon/Storage/MemoryRepository.vb#L73) | `SessionSummaryRecord.LastActiveAt`，默认空串，不影响 `ChatContextBuilder` 既有读取 |
| [HistorySessionService.vb:68](file:///f:/ai/code/AiHelper/ShareRibbon/Controls/Services/HistorySessionService.vb#L68) | 每条 payload 新增 `lastActiveAt`（老数据回落 `createdAt`）与 `documentId`（[:69](file:///f:/ai/code/AiHelper/ShareRibbon/Controls/Services/HistorySessionService.vb#L69)，路径哈希，不含路径本身） |
| [history-manager.js:206](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/history-manager.js#L206) | `buildGroups` 按**首次出现**切桶（[:222](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/history-manager.js#L222) `groupOf`、[:239](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/history-manager.js#L239) `timeBucket`、[:253](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/history-manager.js#L253) `renderGroup`），桶内顺序就是后端顺序 |
| [history-manager.js:270](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/history-manager.js#L270) | `renderEmptyState` 三种空态各带动作；[:155](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/history-manager.js#L155) `runEmptyAction` + [:167](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/history-manager.js#L167) `focusComposer` 把用户送回输入框，动作仍走 [:60](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/history-manager.js#L60) 那一个委托监听，没有新增绑定源 |
| [styles.css:935](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/css/styles.css#L935) | `.history-group*`（[:935-964](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/css/styles.css#L935)）与 `.empty-state-hint/-actions`（[:966-984](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/css/styles.css#L966)） |

### 6.2 分组规则，以及为什么必须是这两个键

- **`scope=all` 按文档分组**：当前文档（后端 `isCurrentDocument`）→ 其它文档（按各自最近一条的先后）→ 通用会话（`document_id` 与 `app_type` 都空、靠空串豁免可见的那批）。组键优先用 `documentId`：`document_name` 只是文件名，两篇 `预算.xlsx` 按文件名会并成一组，把人带到另一篇文档的会话里。
- **`scope=document` 按时间分组**：今天 / 昨天 / 最近 7 天 / 最近 30 天 / 更早。分桶用 `lastActiveAt` 而不是 `createdAt`——排序键就是它，否则「上周开、今天还在聊」的会话排在最前却挂着「更早」的标题，列表自己和自己矛盾。日期只截 `YYYY-MM-DD` 自己算天数差：`"2026-10-24 09:12:33"` 不是标准 ISO，交给 `new Date(字符串)` 各引擎解析结果不一致；解析不出来就落「更早」，不让一条脏时间戳毁掉整列。
- **条目 meta 不再重复文档名**（组头承载），时间精度收到分钟：分组已经说明是哪一天，300px 宽的侧边栏里再挂秒数只会挤行。
- 分组不请求后端、不调模型：一次 `getSessionList` 的 50 条在前端切桶。

### 6.3 三种空状态

| 场景 | 文案 + 动作 |
| --- | --- |
| `scope=document` 且当前文档已保存 | 《x.docx》还没有会话 / 「关闭面板后直接提问，这次会话会自动归到当前文档」→ **直接开始**（关面板 + 焦点回输入框）、**查看全部会话** |
| `scope=document` 且文档未保存 | 这篇文档还没保存，会话只能先记成通用会话 / 「保存文档后新会话就能按文档归组，已有会话可以在「全部」里找到」→ 同上两个动作 |
| `scope=all` 且真的没有任何会话 | 您还没有任何历史会话 + 一句可抄的具体任务示例（「把销售表按金额排序并高亮前 10 行」）→ **开始第一个任务**、**新手引导**（仅当页面已有 `window.showOnboarding` 才出现，复用 E3-01 的向导，不再造第二套引导） |

「直接开始」只关面板并把焦点交回 `smart-input`/`chat-input`，不发消息、不新建会话——发送路径仍然只有输入框那一条。

### 6.4 本机实测（`Read Only=True` 探针，没跑迁移、没改库）

| 观测 | 安装版 `Documents\OfficeAiAppData\office_ai.db`（schema 17） | 调试版 `OfficeAiAppData-Debug`（schema 19） |
| --- | --- | --- |
| 有 `chat_message` 的会话 | 81 | 0（门禁夹具跑完即清，所以行为验证只能落在桩与夹具上） |
| 会话档案 / 带 `document_id` | 27 / 2 | 0 / 0 |
| 按文档聚合 | 通用会话 79、`文档1` 1、`report.docx` 1 | — |
| 按 `lastActiveAt` 分桶 | 昨天 4、最近 7 天 4、更早 73（今天 0、最近 30 天 0），事件日期跨 2026-02-16 .. 2026-10-24 | — |
| 首条与末条不同一天的会话 | 1（恰好没跨桶边界，所以本机按 `createdAt` 分桶目前 0 条会落错组——形状上是错的，桩里 `t-1` 就是它的极端形态） | — |
| 同名不同 `document_id` | 0（哈希键是为这条形状准备的，不是为已发生的事故） | — |

分组标题不全是「通用会话」这件事依赖 E2-04 的第 5/6 步：`document_name` 没被反填进档案时，UI 就只能显示通用会话；本机那 79 条里绝大多数连 `document_id` 都没有，属于「没有任何证据」的兜底人口，本轮不动它们的归属，分组也只是把它们收进同一个「通用会话」组。

### 6.5 门禁补强

| 门禁 | 新增 |
| --- | --- |
| `scripts/check-history-grouping-ui.js`（新，323 行，已登记 [run-code-checks.ps1:108](file:///f:/ai/code/AiHelper/scripts/run-code-checks.ps1#L108)） | 最小 DOM 桩里加载 `history-manager.js`：4 个文档桶的标签/计数/同名拆分、组序 = 首次出现、桶内保持后端顺序、不丢不重、meta 只能是分钟精度时间戳、`lastActiveAt` 决定桶（2020 首条 + 今天末条必须落「今天」）、脏时间戳落「更早」、三种空态文案与动作、`data-empty=start` 关面板并把焦点还给 `smart-input` 且**不产生任何 post**、`data-empty=onboarding` 复用 `window.showOnboarding`（页面没有它就不出现）、组名与标题走 `escapeHtml`、删除确认取消时不落后端、连两次 `init()` 仍只有一个委托监听 |
| `scripts/audit-p0-guardrails.ps1`（[:1433](file:///f:/ai/code/AiHelper/scripts/audit-p0-guardrails.ps1#L1433)、[:1821](file:///f:/ai/code/AiHelper/scripts/audit-p0-guardrails.ps1#L1821)） | 字段合同双向钉住：`GetRecentSessions` 必须仍 SELECT `MAX(e.created_at)` 且映射到 `LastActiveAt`，服务必须仍发 `lastActiveAt`/`documentId`，JS 必须仍有 `buildGroups`/`timeBucket`/`runEmptyAction`/`s.documentId`/`data-empty="start"`，并且 JS 里出现 `sort(` 即失败；CSS↔JS 双向配对：`styles.css` 里的 `.history-group*`/`.empty-state-*` 必须仍被渲染，JS 渲染的 7 个新类必须仍有规则（防「无样式漂移」） |
| `scripts/smoke-change-log-and-session-scope.ps1`（[:388](file:///f:/ai/code/AiHelper/scripts/smoke-change-log-and-session-scope.ps1#L388)） | 真库夹具：把首条事件回拨到 2020-01-01 后，`CreatedAt` 必须仍是首条时间、`LastActiveAt` 必须不等于它、`DocumentId` 必须回传原哈希；输出新增 `GroupingKeysReturned` |

非空跑验证：用同一份 `history-manager.js` 拷贝做 6 个注入变异，全部被断言拦下（真实报错逐条为「time buckets wrong: 更早|今天|最近 7 天|最近 30 天」「scope=all must produce 4 document buckets, got 3」「rendered group wrappers mismatch」「document scope empty state must offer both…」「document name reached the group header as markup」「time buckets wrong: 更早|今天|最近 30 天|最近 7 天|昨天」）；CSS 侧删掉 `.history-group-header` 规则后审计如实失败（`history-manager.js renders '.history-group-header' but styles.css has no rule for it`）。

```powershell
cmd /c build-code.bat
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run-code-checks.ps1 -SkipBuild
```

实测输出：`PASS: ShareRibbon / WordAi / ExcelAi / PowerPointAi`（构建）、`P0 guardrails passed.`、`GroupingKeysReturned : True`、`history grouping UI contract checks passed`、`All code checks passed.`。

### 6.6 声明的不作为

1. **不改后端排序**：`ORDER BY` 仍是「当前文档在前 + `MAX(created_at)` 倒序」，分组只是把它切成可读的块。
2. **不做分组折叠/懒加载/搜索**：50 条上限下组头已经够用；折叠会引入「用户看不到被折走的会话」这类新反馈，需要真机数据再判。
3. **不给分组再排序或按宿主二级拆**：宿主过滤已经在 E2-01 由 `app_type` 决定，UI 不再重复表达。
4. **不动那 79 条无证据会话的归属**：它们继续靠空串豁免可见，属于 E2-04 的兜底人口。
5. **不新增发送路径**：空状态的「直接开始」只关面板+给焦点；`newSession` 仍由面板头部那个按钮负责。
6. **不做未保存文档名的归组猜测**：`文档1` 这类名字仍是一个独立组，不并到任何路径哈希上。
7. **不加「共 N 条 / 已截断」页脚**：payload 目前不带总数，凭 50 条上限写「还有更多」会是假信息；要做得先扩展后端合同。

### 6.7 真机验收清单（按用户指示暂缓，需要 Word/Excel/PowerPoint 真机 + 重启宿主刷新 `%LOCALAPPDATA%\OfficeAI\www`）

1. Word 打开面板 → 历史 → 「全部」：当前文档的会话整块排在第一组，组头就是文档名并高亮成蓝色。
2. 同一列表里另一篇文档的会话出现在第二个组，组头是它的文件名；未绑定文档的老会话集中在最后的「通用会话」组，且**仍然看得见**。
3. 切到「本文档」：组头变成时间（今天 / 昨天 / 最近 7 天 / …），顺序与「全部」里该文档那块一致。
4. 拿一条上周开、今天又聊过的会话：它必须出现在「今天」，行内时间显示今天的最近活跃时刻，而不是上周的首条时间。
5. 行内不再重复文档名，时间是 `YYYY-MM-DD HH:mm`；侧边栏宽度下不出现横向截断。
6. 空态一：一篇没聊过的新文档 + 「本文档」→ 文案带文档名，点「直接开始」面板关闭且光标落在输入框，没有发出任何请求。
7. 空态二：全新用户 + 「全部」→ 有一条可抄的任务示例，「新手引导」能打开 E3-01 向导。
8. 重命名/删除/点击条目在分组结构下与分组前一致：一次点击只触发一次（重点看删除是否弹两次确认）。
9. 文档名或会话标题里带 `<` `>` `&` 时，组头与行标题显示为原文而不是被解析的标签。
10. 三端各跑一次：Excel/PPT 的历史列表只看得到本端 + 通用会话，分组形状不变。

---

## 7. E5-04 不可达入口处理

计划行要求「模板库等隐藏入口要么接通，要么移除；不留只有开发者知道的能力」，验收物是「入口清单 + 真机检查」。这一轮先把清单做成**可执行的门禁**（`scripts/check-entry-reachability.js`，7 条规则），再用它把仓库里能静态证明的不可达入口逐个处理掉。

之所以必须做成门禁：`ChatCommandRouter.Dispatch` 对未注册的 type 只 `Debug.WriteLine`（[ChatCommandRouter.vb:18](file:///f:/ai/code/AiHelper/ShareRibbon/Controls/Services/ChatCommandRouter.vb#L18)），被注释掉的按钮、指向不存在函数的 `onclick`、查不到的 `getElementById` 同样全部静默通过——这类缺陷既不报错也不进日志，只有真机点击才暴露。

### 7.1 入口清单（本轮实测，不是回忆）

| 入口面 | 实测规模 | 判定 |
| --- | --- | --- |
| 面板 → VB 消息 type | 面板侧 56 个发送点，router 66 条注册（收口前 73） | 双向闭合：发得出必须有人接，接得住必须有人发 |
| 内联事件 `on*="fn("` | HTML + 20 个面板 JS 共 70 处引用、57 个函数名 | 70/70 解析到定义（修复前 69/70） |
| `getElementById('字面量')` | 103 处 | 103/103 有节点（HTML 声明 + JS 创建 + VB 拼出的标记） |
| 注入的 `vsto` 桥 | `executeCode`/`sendMessage`/`saveSettings` | 每个方法都有调用点（`checkedChange` 已删，见 §7.3） |
| 三端 Ribbon 按钮 | 18 个（共享基类定义），Word 可见 14、Excel 13、PowerPoint 11 | 可见项全部有非空 Click 处理；隐藏项见 §7.5 |
| 面板输入条按钮 | 命令面板 / 排版模板库（新）/ 设置 / 添加附件 / AI 续写 | 5 个都指向已注册 type 或已定义函数 |

### 7.2 接通：模板库（计划点名的那个隐藏入口）

事实核查结果比预想更糟：排版模板与规范库的**编辑器、列表、管理态、三端子类都实现了**，但三端没有任何一条可见路径能进去——

- `enterReformatTemplateMode()` 只由 VB 注入调用，而唯一调用它的是 `PowerPointAi/Ribbon1.vb:50` 的 `ReformatButton_Click`；
- 该按钮在 `PowerPointAi/Ribbon1.Designer.vb:80` 被设成 `Visible = False`；
- Excel 的 `ReformatButton` 同样被隐藏（`ExcelAi/Ribbon1.Designer.vb:74`），Word 的「排版」走的是 `TriggerSmartReformat()` 智能分析而不是模板库；
- 「模板排版」`TemplateFormatButton` 在共享基类 `BaseOfficeRibbon.Designer.vb:183` 就全局 `Visible = False`，可 Word/PPT 各自还留着 40 行的 `TemplateFormatButton_Click` 实现。

处理方式是在面板输入条上补一个真实入口，复用已经注册的消息而不是新造发送路径：

| 位置 | 改动 |
| --- | --- |
| [chat-template-refactored.html:195-203](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/chat-template-refactored.html#L195-L203) | 新增 `#template-library-button`，样式继承 `#chat-input-actions button`，不改 CSS |
| [reformat-template.js:70](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/reformat-template.js#L70) | `window.openTemplateLibrary()` = `enterReformatTemplateMode()` + 已注册的 `getReformatTemplates`，零新消息 type |
| [chat-template-refactored.html:64](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/chat-template-refactored.html#L64) | 恢复自 `68e5fab` 起被注释掉的「新建排版模板」按钮（`createNewTemplate` 一直有实现） |
| [reformat-template.js:773-776](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/reformat-template.js#L773-L776) | 管理态卡片补「编辑/复制」两个动作，预置模板在 `window.editTemplate`（[:1077](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/reformat-template.js#L1077)）里先拦住并提示「请先复制再编辑」 |
| [code-handler.js:977](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/code-handler.js#L977) | 补上 `exitReformatMode()`：排版横幅的「退出排版」按钮早在 :951 就存在，但函数从未定义，点了不会有任何反应；排版状态由服务端每次响应判定，所以这里只移除横幅 |

### 7.3 移除：能被静态证明走不到的入口

| 移除项 | 证据 |
| --- | --- |
| 7 条 router 注册：`startLoop`/`continueLoop`/`replanLoop`/`cancelLoop`（Ralph 别名，`/loop` 早已在 `message-sender.js` 改写成 `startAgent`）、`startAgentExecution`/`refineAgentPlan`（被 `agent:` 前缀那套取代）、`checkedChange`（旧版 BaseChat 面板有自己的 `Select Case`，不经过这个 router），连带 4 个只为它们存在的 handler | 面板与注入桥里没有任何发送点；注册数 73→66 |
| `BaseChatControl.LoadLocalHtmlFile()` + `WebViewBridge.LoadHtmlFileAsync()` | §4.3 已查明不可达；页面加载由 `InitializeWebView2` 的 `NavigateToString` 负责，`WebViewBridge` 每次调用前都 `_ensureInitialized()`，因此三端 `ShowChatTaskPane` 可退回同步 `Sub`，Word 的 `ShowChatTaskPaneAsync` 与那条从未使用的 `loadDataCaptureHtml`（外加 `loadChatHtml`）一并删除，`WordAi/Ribbon1.vb:72` 改调同步版 |
| `UtilsService.GetVstoBridgeScript()` 里的 `vsto.checkedChange` | 注入面只有 `BaseChatControl.vb:3113` 一处，面板里没有调用点；留着等于留一条往空 type 上 post 的死发送方 |
| `reformat-template.js` 的 `window.uploadStyleGuideDocument` | 无调用点，上传规范另有已接通的路径 |
| `run-code-checks.ps1` 里指向已删除 `ralph-loop.js` 的条目，以及「文件不存在就跳过」的静默兜底 | JS 语法清单改为**枚举真实目录**，纳入检查的文件 8→20 个；清单不再需要人工维护，也就不会再朝两个方向腐烂（漏掉新文件 / 留着死文件） |

### 7.4 门禁补强

| 门禁 | 新增 |
| --- | --- |
| `scripts/check-entry-reachability.js`（新，已登记 [run-code-checks.ps1:109](file:///f:/ai/code/AiHelper/scripts/run-code-checks.ps1#L109)） | A 面板发的 type 必须注册；B 注册的 type 必须有发送方；C HTML 注释里不得藏可交互控件；D 内联 `on*="fn("` 必须有定义；E `getElementById` 字面量 id 必须有节点；F `vsto` 桥方法必须有调用点；G 三端 Ribbon 可见按钮必须有非空 Click 处理。语料含 4 个模块全部 `.vb`：VB 拼出的标记（如 `SmartProofreadFocusMode.vb:361` 的 `id="proofread-summary"`）与注入脚本都算声明方/发送方，只看 `Resources/` 会得出相反结论 |
| `scripts/audit-p0-guardrails.ps1`（[:1884](file:///f:/ai/code/AiHelper/scripts/audit-p0-guardrails.ps1#L1884)） | 「删掉的不得复活」四组：`LoadLocalHtmlFile`/`LoadHtmlFileAsync` 不得再出现；7 条 legacy 注册不得再出现；`GetOfficeAppType` 必须仍是 `MustOverride`（[BaseChatControl.vb:2755](file:///f:/ai/code/AiHelper/ShareRibbon/Controls/BaseChatControl.vb#L2755)）且三端各自 `Overrides`；模板库三点合同（按钮 id + `openTemplateLibrary` + `editTemplate`）与「本门禁已登记进 `run-code-checks.ps1`」 |
| `scripts/smoke-selection-reference.ps1`（[:215](file:///f:/ai/code/AiHelper/scripts/smoke-selection-reference.ps1#L215)） | 原来钉的是已删除的 `HandleCheckedChange`。改为钉存活的那一条：`checkedChange` 注册不得复活，且 `selectedCellChecked = jsonDoc("selectedCell")` 必须仍在 `HandleSaveSettings` 体内——保存对话框是自动引用开关唯一的实时路径 |
| `scripts/run-code-checks.ps1`（[:79](file:///f:/ai/code/AiHelper/scripts/run-code-checks.ps1#L79)） | JS 语法检查改为目录枚举 + 空目录即失败；缺失文件不再被跳过 |

反空转验证：8 个变异探针逐个破坏一条真实链路，全部如实变红（每条都取到真实报错），探针脚本用 Buffer 原样读写以保住 BOM/CRLF，跑完逐文件比对字节一致：

| 探针 | 门禁输出 |
| --- | --- |
| 面板新增未注册 type | `面板发送了 type "probeUnregisteredType"，但 router 没有注册，消息会被静默丢弃` |
| 发送方常量改名（`agent-protocol.js`） | `router 注册了 "abortAgent"，但面板里没有任何发送方——这是一条走不到的入口` |
| HTML 注册改名 | `router 注册了 "abortAgentX"…没有任何发送方` |
| 把新按钮包进注释 | `chat-template-refactored.html:197 注释里藏着可交互控件 (button)` |
| `exitReformatMode` 改名 | `code-handler.js:951 内联事件调用 exitReformatMode()，但 JS 里没有这个函数` |
| HTML 里 `id="smart-input"` 改名 | `agent-card.js:19 查找 id "smart-input"，但 HTML/JS/VB 标记里都没有这个节点` |
| 桥里加一个没人调的方法 | `vsto 桥注入了 probeDeadBridge()，面板里没有任何调用点` |
| Word 隐藏按钮放开（处理体只剩注释） | `WordAi 的 Ribbon 上「聚光灯」(SpotlightButton) 可见，但 SpotlightButton_Click 只剩注释` |

```powershell
cmd /c build-code.bat
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run-code-checks.ps1 -SkipBuild
```

实测输出：`PASS: ShareRibbon / WordAi / ExcelAi / PowerPointAi`（构建）、`P0 guardrails passed.`、20 个 `PASS: node --check`、`Ribbon 入口：WordAi 可见 14/18 / ExcelAi 13/18 / PowerPointAi 11/18`、`入口可达性检查通过：消息 type、内联事件、DOM id、vsto 桥方法、Ribbon 可见按钮五个方向都闭合。`、`All code checks passed.`。

### 7.5 声明的不作为

1. **不改三端 Ribbon 的可见性**：Excel/PowerPoint 隐藏了「排版」「校对」，「模板排版」在共享基类就全局隐藏，可 Word/PPT 仍各自留着 40 行 `TemplateFormatButton_Click`。这些能力确实实现了但三端都摸不到，属于 E5-04 说的同一类问题——但放开可见性会直接改变用户 ribbon 上的按钮数量，是否接、接哪几个留给用户拍板（见 §12 的 D9）。
2. **不动遗留面板栈**：`BaseChat`/`BaseDeepseekChat`/`BaseDoubaoChat` 加三端 `DeepseekControl`/`DoubaoChat` 共 3290 行，自带内联桥与 `Select Case` 派发，挂在**可见**的「Deepseek」「Doubao」两个按钮上，是主链（`OfficeHarness -> AgentKernel -> LoopEngine -> ToolRegistry`）之外的第二套对话实现。它可达、用户看得见，因此不在「移除不可达入口」这一行里，但它与 AGENTS.md 的「禁止另建第二套」直接冲突，同样并入 D9。
3. **不把 `vsto.checkedChange` 在遗留面板里的同名实现删掉**：那三个类自己 `Select Case` 处理它，删桥不删派发只会造成不一致；随 D9 一起处理。
4. **不删 `SaveFullWebPageAsync`**：它每轮写 `saved_chat_*.html`，本轮删掉 `LoadLocalHtmlFile` 后已无人读回，但产物在用户磁盘上、属于可见副作用，先记为 D8 等用户决定。
5. **不扩展门禁到遗留面板的内联 HTML**：那套面板的标记写在 VB 字符串里，规则 A–G 只覆盖 `chat-template-refactored.html` + 面板 JS + 共享 router；先把主链的入口闭合钉死，避免一次性铺开成不可维护的正则。
6. **不新增任何发送路径**：模板库入口复用 `getReformatTemplates`，退出排版只是移除横幅；面板与宿主之间的 type 合同只做收口，不做扩张。

### 7.6 真机验收清单（按用户指示暂缓，需要 Word/Excel/PowerPoint 真机 + 重启宿主刷新 `%LOCALAPPDATA%\OfficeAI\www`）

1. 三端各点一次输入条上的「排版模板库」：进入模板/规范列表管理模式，能看到本端已实现的模板与规范条目。
2. 在该列表里点「新建排版模板」→ 编辑器打开且能保存；保存后列表出现新条目。
3. 对自建模板点「编辑」→ 进入编辑器；对预置模板点「编辑」→ 出现「预置模板不可直接修改，请先复制再编辑」且不落后端；点「复制」→ 生成副本。
4. Word 触发智能排版后顶部出现紫色排版横幅，点「退出排版」横幅立即消失（此前点了没有任何反应）。
5. Word/Excel/PPT 打开面板不再需要等本地 HTML 文件加载：`ShowChatTaskPane` 改为同步后，首次打开不出现空白面板或 `ERR_FILE_NOT_FOUND`。
6. Word 设置对话框里关掉「自动引用选区」并保存 → 当轮即生效（不再有 `checkedChange` 这条实时路径）；重新打开后仍是关的。
7. 三端 ribbon 与本轮之前一致：隐藏项没被误放开，可见项（Word 14、Excel 13、PPT 11）点击都有反应。
8. 面板控制台无新增 `is not defined` / `null` 报错（重点看模板库、退出排版、历史列表三处）。

---

## 8. E5-05 CI 防漂移

计划行的验收标准是「四份能力清单不一致时构建失败；新增能力必须更新清单 → CI 输出」。E5-01 已经把四条清单的一致性断言写进 `audit-p0-guardrails.ps1`（P7/P8），本轮量完才发现**那两条断言从来没在发布检查里跑过**。所以这条的产出不是再加一条断言，而是把两个更基础的事实收口：门禁清单只能有一个持有者，一条合同也只能有一个持有者。

### 8.1 改之前量到的漂移（以下全部出自 `git show HEAD:`，不是回忆）

| 位置 | HEAD 事实 | 后果 |
| --- | --- | --- |
| `scripts\run-code-checks.ps1` | 手写 16 条名单；磁盘上 22 个 `.ps1` 去掉 5 个非门禁（runner 自己、`build-code-projects`、`build-installer-prep`、`ensure-vsto-temp-keys`、只报告不失败的 `audit-installer-duplicates`）应是 17 条真门禁 | 唯一漏掉的 `run-professional-design-smoke.ps1` 是**会真失败**的门禁（[:38](file:///f:/ai/code/AiHelper/scripts/run-professional-design-smoke.ps1#L38) `throw "Preflight failed for slide ..."`），但从未被跑过 |
| `build\RunReleaseChecks.ps1` | 第二份私有名单，只有 5 条（ai-gateway-provider / word-capability-registry / memory-pipeline / skills-registry / db-schema-drift），不含 `audit-p0-guardrails.ps1` | **P7/P8 不是发布门禁**——计划行要的「不一致时构建失败」当时是空的 |
| 同文件的 `Invoke-Step` | `& $Action` 之后无条件 `Write-Host "PASS"` | 实测：子脚本以 `exit 1` 结束时父级照样打印 PASS，`LASTEXITCODE=1` 没人读，后面的步骤继续跑。任何以退出码失败的审计只是在报告里写了一句没人看的话 |
| JS 语法检查 | 手写 7 个路径，其中 `ralph-loop.js` 在 HEAD 已删除；循环体是 `if (Test-Path) { 检查 }`，没有 else | 名单里的死名字长期为绿；目录里实际 20 个 `.js`，被检查的只有 6 个 |
| node 缺失 | `Write-Warning` 后整段跳过 | 没装 node 的机器上三条面板合同门禁（agent 卡、历史分组、入口可达性）根本不跑，套件仍报 `All code checks passed.` |

### 8.2 门禁入口收口：让清单成为从磁盘推出的性质

- **发现规则反转**（[run-code-checks.ps1:42-64](file:///f:/ai/code/AiHelper/scripts/run-code-checks.ps1#L42)）：`scripts\*.ps1` 默认全是门禁，除非在 `$notGates` 里具名退出并写出理由。数量不写进代码，只留两条守卫：下限 `< 20` 即「发现规则自己坏了」，以及退出项必须仍在盘上（脚本被删后理由条目就成了死账）。前缀规则（`smoke-*.ps1`）看着通用，实测正是它漏掉 `run-professional-design-smoke.ps1` 的——名字不以 `smoke-` 开头的真门禁也是门禁。
- **发布路径不再自备名单**：`RunReleaseChecks.ps1` 的那 5 条私有清单删掉，改为一步委托 `scripts\run-code-checks.ps1`。不带 `-SkipBuild` 是有意的：反射类门禁读 `bin\Debug`，不是上一步刚产出的 Release 目录。
- **退出码进入判定**：两份 runner 的 `Invoke-Step` 统一为 `$global:LASTEXITCODE = 0` → 运行 → 非 0 即 `throw`，注释里留下实测症状。
- **缺失即失败**：JS 侧改目录枚举 + 空目录 `throw`；node 缺失由 warning 改为 `throw`，只有 `-SkipNodeCheck` 能显式声明「这台机器就是跳过」。

### 8.3 注入合同并轨：一份合同不能有第二个持有者

同一条「VB 注入的 JS 必须成立」此前有 **3 个持有者**：`audit-p0-guardrails.ps1` 的 P4-C 片段、`smoke-selection-reference.ps1` §6 的 id 存在性循环、E5-04 新建的 `check-entry-reachability.js` 规则 A–G。三份语料不同、能力不同、失效方式不同。合并为一条门禁的规则 H/I，两份副本删掉，并把「合并本身不许退回」钉住。

| 规则 | 覆盖 | 本轮实测 |
| --- | --- | --- |
| H 注入块里的字面量 id 必须有节点 | 语料是四模块全部 309 个 `.vb`（剥注释后全文扫），模式含 `getElementById('x')`、`querySelector('#x')`、`setNativeEntrySetting('x')`（[:296-311](file:///f:/ai/code/AiHelper/scripts/check-entry-reachability.js#L296)） | 16 处查询，16 个解析到节点 |
| I 注入块调用的 JS 函数必须有定义 | 载荷用 VB 词法读出（字符串内部不透明、剥掉插值空洞）再取调用名（[:61-121](file:///f:/ai/code/AiHelper/scripts/check-entry-reachability.js#L61)、[:338-353](file:///f:/ai/code/AiHelper/scripts/check-entry-reachability.js#L338)） | 124 段载荷、61 个调用名，61 个解析到定义或被 `typeof` 守卫 |

两条缺陷是这条扫描跑出来的，不是推理出来的：

1. **接通**：[WordAi/ChatControl.vb:1748](file:///f:/ai/code/AiHelper/WordAi/ChatControl.vb#L1748) 注入 `"showProofreadModeIndicator(); showProofreadLoading();"`，而 `showProofreadLoading` 面板里从来没有过——同一注入块顺序执行，前一句生效、这一句抛 ReferenceError，于是校对轮次右上角动画从不出现，除了面板控制台没人知道。改成 `showLoadingIndicator()`（[core.js:128](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/core.js#L128)，`showProofreadModeIndicator` 在 [code-handler.js:894](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/code-handler.js#L894)），收起仍走既有路径：`changeSendButton()`（[message-sender.js:204](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/message-sender.js#L204)）由宿主完成路径 [BaseChatControl.vb:1535/1791/2883](file:///f:/ai/code/AiHelper/ShareRibbon/Controls/BaseChatControl.vb#L1535) 调用。这条入口本身是产品要的，所以接线。
2. **删除**：`AgentKernelService.vb` 两处（HEAD `:328`、`:443`）执行 `document.getElementById('planning-status-card').remove()`，该 id 在面板里不存在；两处都带 `if(...)` 守卫所以不报错，属于「声称复位规划卡、实际一条 DOM 都没动」的死码。规划卡真正的 id 是 [agent-card.js:93](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/agent-card.js#L93) 的 `'agent-plan-' + uuid`，终态复位由 `restoreAgentRequestUi()` 负责（输入锁、发送/终止按钮、动画都在它里面）。兜底语句保留，去掉死片段（[:509](file:///f:/ai/code/AiHelper/ShareRibbon/Controls/Services/AgentKernelService.vb#L509)）。

同样是坏注入，一条接一条删，判据不是偏好而是「谁拥有这块 DOM」：第 1 条功能本来就该有，第 2 条职责已有唯一持有者。

已知边界（写在文件头注释里）：插值空洞内的字符串不需要 doubled quote，`""` 判定偶尔让载荷读偏一格，因此 I 按「解析得到的载荷」取下限，字面量 id 交给 H 全文扫，不受该边界影响。被删掉的 P4-C 副本恰恰是「带 VB 插值的载荷整条跳过」，而注入载荷绝大多数带插值——这就是它能长期零发现的原因。

### 8.4 同一形状推广到文档地图：规则 J

`AGENTS.md` 是 Agent 的地图，而地图坏掉时没有任何东西会响：它既不是运行时资源，也不参与编译。本轮把规则 H/I 的形状（「引用的目标必须存在于磁盘」）原样用在导航文档上，先量再钉。

量到的 HEAD 现状（语料 = git 跟踪的导航层 11 份：根 `AGENTS.md`/`CLAUDE.md`/`README.md` + 各级 `AGENTS.md` + `Storage/README_DATABASE.md`；解析 = markdown 链接按文档目录、以顶层目录开头的反引号路径按仓库根，两种基准都试）：

| 指标 | HEAD | 现在 |
| --- | --- | --- |
| 导航文档 | 11 | 11 |
| 可解析的本地引用 | 24 | 38 |
| 指向不存在文件的引用 | **5** | 0 |

那 5 条是：根 `AGENTS.md` 的 `docs/build-and-installer.md`、`docs/signing-and-certificates.md`，`CLAUDE.md` 的 `docs/VS2025-VSTO-Migration-Guide.md`，`README.md` 的同两条。三条被指向的文件从未进过仓库——症状不是报错，是下一个（人或 Agent）按地图去找文档、找不到、回去重新考古一遍。修法一律是**改指向真实存在的地方**（`README.md`「从源码构建」/「VSTO 开发证书与发布签名」两节、`docs/VS2026-InstallerProjects.md`、`build/SignArtifacts.ps1`），而不是补一个空文件去让门禁闭嘴。

范围边界是量出来的，不是拍脑袋的：

1. **只核两种能唯一解析的写法**。第一版把正文里随手写的裸文件名（`agent-card.js`）和模块内相对写法（`Storage/AppTypeContract.vb`）也算成引用，得到 13 条假阳性。一条 noisy 的门禁最后的归宿是被人关掉，所以这两种写法明确不判。
2. **不收 `docs/` 正文**。同一把尺子量 `docs/*.md`（21 份、120 条引用）在 HEAD 有 11 条断链，本轮修掉其中 3 条（`docs/VS2026-InstallerProjects.md` 的 `./build-and-installer.md`、`docs/ai-native-harness-design.md` 的两条），还剩 8 条：`.github/workflows/code-checks.yml`（2 处，写的是「计划要做」）、`build/reports/golden/latest.md`（评估集的产出物，跑一次就有）、`docs/project-assessment-optimization-plan.md`、`ShareRibbon/Agent/Harness/HarnessEvents.vb`、`ShareRibbon/Common/ComObjectHelper.vb`（3 处，审计文档描述的是当时的设想）。设计文档天生会描述尚未存在的文件，把它们算成坏引用会让门禁天天红——这 8 条留作已知账目，规则 J 只保证导航层不再说谎。
3. **扫描自己失效也要失败**：`navRefs.size < 30` 即报错（[:403-405](file:///f:/ai/code/AiHelper/scripts/check-entry-reachability.js#L403)）。38 条引用只比下限高 8，这是本轮故意留的紧下限——导航层再长出一批新引用它不会响，但语料目录一旦改名或 `collectNavDocs` 一旦失效，它立刻响。

### 8.5 门禁补强

| 门禁 | 变更 |
| --- | --- |
| `scripts/check-entry-reachability.js` | 新增规则 H、I、J；再加三条「扫描自己失效」下限：注入载荷段数 `< 100`（[:119](file:///f:/ai/code/AiHelper/scripts/check-entry-reachability.js#L119)）、面板 JS 函数名 `< 100`（[:323](file:///f:/ai/code/AiHelper/scripts/check-entry-reachability.js#L323)）、导航引用条数 `< 30`（[:403](file:///f:/ai/code/AiHelper/scripts/check-entry-reachability.js#L403)）都直接失败——否则一条空跑的正则会把「没查到问题」报成「没问题」 |
| `scripts/audit-p0-guardrails.ps1` | 删掉 P4-C 副本（[:781](file:///f:/ai/code/AiHelper/scripts/audit-p0-guardrails.ps1#L781) 记下为什么删）；把 3 条「我的脚本名在名单里」的旧钉法换成形状钉法（[:1508](file:///f:/ai/code/AiHelper/scripts/audit-p0-guardrails.ps1#L1508) `Get-ChildItem` + `$notGates`、发布路径必须委托 runner、`Invoke-Step` 必须看 `$LASTEXITCODE`）；新增 E5-05 钉：规则 H/I/J 的六个组成 token 必须齐（`injectApiRe`、`vbIdLookups`、`vbBridgeCalls`、`jsGlobals`、`collectNavDocs`、`navRefs`，防半并轨，[:1895](file:///f:/ai/code/AiHelper/scripts/audit-p0-guardrails.ps1#L1895)），`showProofreadLoading` 不得复活、`showProofreadModeIndicator(); showLoadingIndicator();` 不得消失、`planning-status-card` 不得回到 `AgentKernelService`（[:1904-1912](file:///f:/ai/code/AiHelper/scripts/audit-p0-guardrails.ps1#L1904)） |
| `scripts/smoke-selection-reference.ps1` | §6 的私有 id 循环与两条死 id 钉删除（[:182](file:///f:/ai/code/AiHelper/scripts/smoke-selection-reference.ps1#L182) 记下由 H 泛化覆盖），保留注入抽取与 `referencedIds.Count -lt 8` 下限 |
| `scripts/run-code-checks.ps1` | 门禁发现反转 + 数量下限 + 退出项存在性反向检查 + 打印 `discovered N gates`；node 缺失改 `throw`（[:83](file:///f:/ai/code/AiHelper/scripts/run-code-checks.ps1#L83)） |
| `build/RunReleaseChecks.ps1` | 私有名单删除，改委托；`Invoke-Step` 补退出码判定（[:18-25](file:///f:/ai/code/AiHelper/build/RunReleaseChecks.ps1#L18)） |

### 8.6 反空转验证

Node 侧 5 个探针：改真实源码 → 跑门禁 → 断言退出码非 0 且报错里出现被改的名字 → 用 latin1 Buffer 原样写回并逐文件比对字节。

| 探针 | 结果 |
| --- | --- |
| VB 注入查一个不存在的 id（`settings-selected-cell` → `...cellX`） | 变红，报错含 `settings-selected-cellX` |
| `setNativeEntrySetting('command-palette-button')` 首参改名 | 变红，报错含 `command-palette-buttonX` |
| 注入载荷调用一个没人定义的函数（`showContextHints(` → `showContextHintz(`） | 变红，报错含 `showContextHintz` |
| 把根 `AGENTS.md` 里的 `docs/VS2026-InstallerProjects.md` 改成不存在的 `...ProjectsX.md` | 变红，报错含 `VS2026-InstallerProjectsX` —— 规则 J 不是只在新写文档时有效 |
| 把 5 个注入 API 名整体改名（让扫描自己变瞎） | 变红：`只解析到 0 段注入载荷，是这条扫描自己失效了，不是桥变干净了` |

Runner 侧 3 个探针（新建临时失败门禁 / 改坏发现过滤 / 关掉退出码判定，跑完删文件并还原）：

| 探针 | 结果 |
| --- | --- |
| P1 新增 `scripts\zz-e505-probe-gate.ps1`（名字不匹配任何前缀规则，内容 `exit 1`） | 套件 `exit=1`，报错 `scripts\zz-e505-probe-gate.ps1 failed with exit code 1` —— 新门禁无需登记即被跑，且退出码真能拦 |
| P2 把 `-Filter *.ps1` 改成 `-Filter *.pszz` | `exit=1` 且 `discovery itself is broken` —— 发现规则坏掉时不会退化成「0 条门禁、全绿」 |
| P3 把 `if ($LASTEXITCODE -ne 0) {` 改成 `if (... -and $false) {` | 同一条失败门禁打出 `PASS: scripts\zz-e505-probe-gate.ps1` 且套件 `exit=0` —— 退出码判定是被需要的，不是装饰 |
| 还原 | `run-code-checks.ps1` 还原后字节一致 |

另有一条不靠变异、直接改环境的实测：把 `C:\Program Files\nodejs` 从 `PATH` 去掉后跑 `run-code-checks.ps1 -SkipBuild`，23 条 PowerShell 门禁照常跑完后在 [run-code-checks.ps1:83](file:///f:/ai/code/AiHelper/scripts/run-code-checks.ps1#L83) 抛 `node not found; the panel-contract gates ... cannot run`，进程退出码 1 —— HEAD 那种「没有 node 就跳过三条面板合同门禁、照样报 all passed」的形状已经不成立。

P3 第一次写成 `-ne 999` 时反而抛了 `failed with exit code 0`：条件被写成恒真、阈值成了实参。这条记下来是因为它正是本轮反复遇到的失效形状——**守卫看着在、判定条件却与意图相反**。

本轮最终实测：`discovered 23 gates under scripts\`，46 个 `PASS` 步骤（23 条 PowerShell 门禁 + 20 条 `node --check` + 3 条面板合同门禁）、`P0 guardrails passed.`、`All code checks passed.`。可达性门禁十条打印里本轮新增的两条是 `注入 id：VB 语料 16 个字面量查询，16 个能解析到节点` 与 `导航路径：11 份导航文档、38 条本地引用，38 条能解析到磁盘`；5 个 Node 探针与 3 个 Runner 探针在同一份代码上重跑，结果是 `all 5 probes turned the gate red and restored byte-identically`。

### 8.7 声明的不作为

1. **不把 E2-05 的 CSS↔JS 双向配对推广到全面板**：本轮实测 `styles.css` 有 330 个类名，整词边界匹配后 13 个查不到引用——其中 12 个是拼接前缀（`agent-trust-risk-${riskLevel}` [agent-card.js:232](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/agent-card.js#L232)、`step-msg-${status}` [:272](file:///f:/ai/code/AiHelper/ShareRibbon/Resources/js/agent-card.js#L272)），只有 `.reference-chip.file-chip`（styles.css:1772）是真正的死样式。要把它跑成门禁就得为拼接开一份白名单，而白名单就是 E5-04 刚删掉的那种腐烂清单，换个地方再长一遍。E2-05 那条只钉本轮新增的 7 个类，范围与生命周期都清楚，保持原样。
2. **不给 `audit-installer-duplicates.ps1` 补失败条件**：它只报告重复 `SourcePath`，安装门禁在 `build\`；本轮把它从「永远匹配不到的排除项」改成 `$notGates` 里带理由的具名退出——退出这件事从此可见，而不是一句没人读的 `-notcontains`。
3. **不改任何一条门禁的内部实现**，只改它们的调用方式与清单持有者。
4. **不假装已有 CI**：仓库里没有流水线定义（`.github/` 只有 `copilot-instructions.md`），所以「CI 输出」这一行落在 `build\RunReleaseChecks.ps1` 这个唯一带退出码的入口上；接进真实流水线前，它是能失败的那个地方。
5. **不在门禁里跑真 Office/COM**，继续用反射与桩 DOM。

### 8.8 真机验收清单（按用户指示暂缓，需要 Word/Excel/PowerPoint 真机 + 重启宿主刷新 `%LOCALAPPDATA%\OfficeAI\www`）

1. Word 触发一次智能校对：右上角加载动画出现（本轮唯一肉眼可见的功能修复），轮次结束后自动消失，且动画消失不是靠手动再点一次。
2. Agent 跑到终态与手动终止各一次：输入锁解除、发送按钮回来、规划卡不残留在页面上——确认删掉 `planning-status-card` 那两段没有让卡片留在聊天流里。
3. 三端各跑一次完整回合后面板控制台无新增 `is not defined` / `Cannot read properties of null`。

（原第 4 条「没有 node 的机器必须失败」不放在这里：它不需要 Office，本轮已直接改 `PATH` 实测，见 §8.6。）

---

## 9. 门禁补强（E1-07）

| 脚本 | 新增断言 |
| --- | --- |
| `scripts/smoke-change-log-and-session-scope.ps1` | E1-07 段：小快照被接受且声明 reversible；带快照的写入 `HasRestore=True`、`NativeUndoable=False`、`RestoreCount=1`、`FullyUndoable=True`；变更记录 JSON 里没有快照正文；“一步有快照 + 一步没有”必须拒绝整轮撤销并带 `Reason`；持久化后 `GetRestoreSnapshotsByRun` 按 kind 读回、`restore_json` 保留正文；超上限的快照被拒 |
| `scripts/audit-p0-guardrails.ps1` | `RestoreSnapshot` 必须仍是 `<JsonIgnore>` 旁路；变更日志必须从旁路读取且保留两路判定；仓储必须持久化 `restore_json` 并提供 `GetRestoreSnapshotsByRun`；迁移 018 文件存在；终态描述必须带 `restoreCount` 且**不得**再出现 Word 宿主类型判断；批次降级必须携带快照；两端 `PerformUndoAgentRun` 必须读快照、必须在写前建快照、宿主不得把 `restore` 塞进 `observation`；`agent-card.js` 必须区分撤销单元与快照重放。既有 PPT 观察守卫窗口 4000→6000 字符（函数体因接入快照变长，仍为局部性检查） |
| `scripts/smoke-db-schema-drift.ps1` 对照基线 | `ShareRibbon/Storage/OfficeAiDbSchema.current.sql` 增加 `restore_json TEXT`，迁移产物与快照形状继续逐列对齐 |
| `scripts/smoke-empty-db-initialization.ps1` | 空库初始化后的 schema 版本断言 17→18，并断言 `agent_run_change` 已有 `restore_json` 列 |

```powershell
cmd /c build-code.bat
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/run-code-checks.ps1 -SkipBuild
```

全部门禁已通过（含 4 个代码项目构建）。新增的 `PowerPointAi\Undo\PowerPointRunRestore.vb` 由通用「每个 `.vb` 必须被本项目 `.vbproj` 编译」守卫覆盖，已登记 `PowerPointAi.vbproj`。

## 10. 已知边界（E1-07；不是本轮遗漏，是明确声明的不作为）

1. **Excel 快照只覆盖所采区域的值/公式/数字格式**：字体、颜色、列宽、条件格式以及任何结构性改动都不在快照内；当前观察口径下“只改格式”的写入本来就不报 `changed`，因此也就不会被承诺可撤销。
2. **PowerPoint 快照只覆盖文字与备注**：格式、动画、形状增删、图片/图表一律不产快照；CreateSlides 专业链仍用它自己的标记删页回滚，E1-07 没有改动那条链，也没有把它接到“撤销本轮”上。
3. **声明式 `OfficeObjectOperation` 批次尚未产快照**：三宿主的操作执行器走到这条路时这一步没有还原依据，整轮 `FullyUndoable` 即为 False，UI 显示原因而不是给一个撤不动的按钮。
4. **文本重放的固有代价**：PowerPoint 用 `TextRange.Text =` 写回会重置该形状的字符级格式，因此重放只改“文字确实不一致”的形状，尽量缩小影响面；这一点必须在真机验收里被用户看见，而不是写在注释里。
5. **撤销不是无限期可用**：快照只在本地库，文档被用户大改、另存或页面/区域被删后，重放会拒绝执行并说明原因——这是设计目标（宁可不撤也不撤错），不是缺陷。

## 11. 真机验收清单（E1-07，按用户指示暂缓，需要 Word 2021/Excel/PowerPoint 真机 + 重启宿主刷新 `%LOCALAPPDATA%\OfficeAI\www`）

1. Excel：让 AI 改写某区域值 → 结束后按钮显示“撤销本轮（按快照还原 N 处）”，点击后区域回到执行前，日志无“还原失败”。
2. Excel：让 AI 同时改值与插行 → 不提供撤销按钮且给出原因（结构闸门生效）。
3. Excel：撤销前手动改一下目标区域 → 点击后如实拒绝“本轮之后又有其他改动”，且本轮仍标记为未回滚，可再点一次。
4. Excel：只改字体/颜色的指令 → 观察不报 `changed`，摘要与撤销口径一致，不出现“已写入 1 处但撤不掉”。
5. Excel：公式区域还原后公式仍是公式、`"1-1"` 类文本常量不被吃成日期。
6. PowerPoint：改写某页文字 → 撤销后该页文字与备注回到执行前，形状数量不变。
7. PowerPoint：加动画/改格式/加形状的指令 → 不提供撤销按钮并显示原因（通道闸门生效）。
8. PowerPoint：删掉目标页后再点撤销 → 拒绝并说明“第 N 页已不在演示文稿里”。
9. Word（回归）：原生撤销单元路径行为与文案不变，`units` 分支仍显示“N 个撤销单元”，首尾文档指纹校验未被削弱。
10. 三端对照：同一类“写 2 步、其中 1 步无法还原”的任务，三端都应当**不出现**撤销按钮而不是出现一个只撤一半的按钮。
11. 隐私与体积：`agent_run_change` 之外（`agent_run_step.observation_json`、聊天历史、日志、面板变更摘要）都查不到快照正文；长文档单条 `restore_json` 不超过 60000 字符。
12. 升级：旧库（version 17）打开后自动补 `restore_json` 列且不改写既有行；新库建表即含该列。

---

## 12. 下一步

W4 剩余行：E5-05 CI 防漂移，以及 M1 闸门收口。E2-04 已经把 `app_type` 变成一份闭合口径（六张表 + 六处写入 + 迁移 019 + 三条门禁），E2-05 才有可分组的输入：那一轮把历史列表切成可读的文档/时间两组、给出三种带动作的空态，并把分组真正需要的两个键（`lastActiveAt` = 排序键、`documentId` = 路径哈希）补进 `getSessionList` payload；**排序仍然只在后端**，前端按首次出现切桶这条不变式既写在注释里，也被 `audit-p0-guardrails.ps1` 的「JS 里出现 `sort(` 即失败」钉住。仍未解析归属的本机老会话（81 个会话里 79 个没有 `document_id`）必须继续靠空串豁免三端可见，分组只是把它们收进同一个「通用会话」组，不许在分组时丢掉。

E5-04（§7）已经落地两件事：§4.3 那条不可达的 `LoadLocalHtmlFile` 连同 `WebViewBridge.LoadHtmlFileAsync` 一起删除、`ShowChatTaskPane` 退回同步；§5.5 第 4 条把 `GetOfficeAppType()` 从「默认 `"Unknown"` 的 Overridable」改成 `MustOverride`，让「子类没实现」在编译期就暴露，并由审计要求三端各自 `Overrides`。更重要的是它把「入口」变成了可执行清单：`scripts/check-entry-reachability.js` 用 7 条规则闭合消息 type、内联事件、DOM id、`vsto` 桥方法与三端 Ribbon 可见按钮，8 个变异探针证明它不是空转。

E5-05（§8）已接手并落地那两件事：注入点守卫从「只有设置注入那一处」扩到四模块全部 309 个 `.vb`（规则 H 管字面量 id、规则 I 管调用名），并把 E2-05 的 CSS↔JS 配对、E5-04 的静态合同闭合与原本散在 `audit-p0-guardrails.ps1` 里的三份副本并成一条有唯一持有者的门禁；顺带量出并修掉了门禁自身更严重的两处漂移——发布检查只有 5 条私有名单、子脚本 `exit 1` 照样打印 PASS。**CSS↔JS 双向配对经实测没有推广**（330 个类名里 12 个是拼接前缀，白名单会变成 E5-04 刚删掉的那类腐烂清单），这一点在 §8.7 第 1 条如实记下。同一形状最后落在了文档地图上（规则 J，§8.4）：HEAD 的 11 份导航文档有 24 条本地引用、5 条指向从未存在的文件，现在 38 条引用 0 条断链，`docs/` 正文里剩下的 8 条是设计文档在描述还没建的东西，明确不归这条门禁管。

W4 只剩 **M1 闸门**。`plan` §6 把 M1 定义成五幕真机演示 + 四条通过线，不是「再补一条断言」，所以它天然要用户坐在机器前；本轮能收口的自动化侧已经收口（§8）。剩下唯一需要用户拍板的是承载：本仓库没有流水线定义（`.github/` 只有 `copilot-instructions.md`），`build\RunReleaseChecks.ps1` 是当前唯一有退出码的入口，「不一致就构建失败」要有第二台机器来跑，得先定这条链跑在哪、用什么触发。五幕的验收包见 §13。

M1 卡在真机与承载这两件用户侧的事上，按「先不用做真机验收，直接接着实现」的指示，主线没有停：已推进到 Phase 2 的 E6（任务包框架），并落地 E6-02/03/04（见 §14）与 E6-01（见 §15）。E6 剩下的行按依赖顺序是 **E6-05**（把包里的断言在 run 结束时真的核对一遍，复用 `ValidateObservedOutcome`，不新建验证链）→ **E6-06**（每包 ≥1 成功 + ≥1 失败黄金样例；现在三个包的 `goldenCases` 仍全为空数组，门禁输出的 `golden references=0` 就是这个事实）。E6-07（没有入口的任务包）依赖刚收口的 Ribbon 话术与 D9，E6-08（六包真机 ≥60%）依赖真机，仍排在其后。

等用户拍板的开放决策：#35（Skill 边界归属）、#36（Excel/PPT capability 清单——E6 之后它不再只是一个"要不要补注册表"的问题：`excel.clean-and-fix`、`excel.insight` 两个包的 `capabilityIds` 现在被门禁强制为空，一旦 #36 定下来就要同时改包和改那条空数组约束）、D7（E17 只读分诊范围）、D8（`SaveFullWebPageAsync` 每轮写 `saved_chat_*.html`，本轮删掉唯一读取者后已无人读回——留产物、留读取，还是整条去掉）、D9（§7.5 第 1/2 条：三端 Ribbon 隐藏但已实现的「排版/校对/模板排版」是否放开，以及挂在可见按钮上的 3290 行遗留面板栈 `BaseChat`/`BaseDeepseekChat`/`BaseDoubaoChat` 是否并入主链——它与 AGENTS.md「禁止另建第二套 Harness/Loop/派发」直接冲突）。W1/W2 提交未推送，W3/W4 全部改动仍在工作区，是否提交/推送等待用户指令。

---

## 13. M1 闸门验收包（等真机，一次性跑完）

`plan` §6 的 M1 是「五幕演示 + 四条通过线」。这一节只做一件事：把五幕翻译成**按编号点击就能跑**的清单，并把每一幕已经能静态证明的部分和必须肉眼确认的部分分开。各任务自己的真机清单（week1 §3、week2 §6、week3 §7、本报告 §2.5/3.5/4.6/5.6/6.7/7.6/8.8）不在此重复，遇到交叉只给指针。

### 13.1 前置（顺序不能换）

1. 关闭全部 Office 进程（本轮收尾时 WINWORD 仍在运行；宿主不关，第 2 步还原的配置会被运行中的实例再写回去）。
2. `powershell -File .qoder\realcheck.ps1 -Mode restore`：把之前为验证 E3-03 而置空的 key 还原。备份 `.qoder\realcheck-office_ai_config.json` 是 2026-10-24 20:54 的，实时配置 2026-10-25 11:48 被写过（比备份新），所以这一步等价于丢弃宿主这中间保存的配置——确认这不是你要留的东西再执行。
3. `.\build-code.bat` 编四项目（Debug 即可，反射类门禁读 `bin\Debug`），再跑 `powershell -File scripts\run-code-checks.ps1 -SkipBuild`：基线是 `discovered 23 gates`、46 个 `PASS`、`All code checks passed.`。这一步的作用是「先让机器证明没有假绿」，再去看人眼部分。
4. 启动 Word/Excel/PowerPoint，核对 `%LOCALAPPDATA%\OfficeAI\www` 里 `chat-template-refactored.html` 与 `js\*.js` 的时间戳晚于本轮改动（面板跑的是 www 副本，不是仓库文件；时间戳不新就等于在验证旧代码）。

### 13.2 五幕

| # | 动作 | 必须看到（可证伪） | 已经静态证明的部分 |
| --- | --- | --- | --- |
| 1 | 全新配置状态下走完向导（`privacyAccepted` 置假即回到第 1 步），配置一个真实 Key，跑 Word 一键示例 | 60 秒内出现文档写入结果；空 Key 时发送被拦住且只给一条 CTA，不静默失败 | week2 §3/§4、week3 §5；`scripts/smoke-onboarding-wizard.ps1` |
| 2 | 发起一个 5 步 Word 任务，中途点停止 | 摘要写明「已写入 N 处」且 N 与实际段落数相符；一键撤销后文档回到执行前 | week1 §1（E1-01/02/03/04）、本报告 §1（E1-07 快照路径）；变更日志与 `restore_json` 的落库形状由 `smoke-change-log-and-session-scope.ps1` 钉住 |
| 3 | 故意规划一次高风险写入（覆盖/删除类），在审批步拒绝 | 只拒当前步骤，任务继续给出「部分完成」摘要；不出现整轮静默终止，也不出现「成功」字样配 0 写入 | week2 §1（E1-05 审批粒度）、week3 §2（E1-08 信任 UI）；P7/P8 能力清单一致性在 `audit-p0-guardrails.ps1` |
| 4 | Excel 打开历史列表，再切到 Word 打开同一个列表 | 两端各只见自己宿主的会话 + 「通用会话」组；老会话（本机 79/81 条无归属）不许在任何一端消失 | 本报告 §5（E2-04 迁移 019、空串豁免三端可见）、§6（E2-05 分组只在后端排序）；排序不变式由「JS 里出现 `sort(` 即失败」钉住 |
| 5 | 抽查 10 个可见入口（Word 侧优先），逐个点击 | 每个都有真实行为，且没有「点了没反应」的按钮 | 门禁实测可见数：Word 14/18、Excel 13/18、PowerPoint 11/18（本报告 §7.1）；type/事件/id/`vsto` 桥/注入载荷/文档路径六个方向由 `check-entry-reachability.js` 十条规则闭合 |

### 13.3 四条通过线各自的判定点

- **假成功 0**：第 2、3 幕。判据不是「看起来对」，而是终态文案与实际 DOM/文档状态一致——这也是 E5-03 与 E1-07 反复处理的那类「声明与结果不符」。
- **未授权写入 0**：第 3 幕。拒绝之后文档里不得出现该步骤的产物。
- **撤销脚本 100% 通过**：第 2 幕，Word 走自定义撤销记录，Excel/PPT 走快照重放（§1）；两端是两条不同合同，必须分别跑。
- **跨端串会话 0**：第 4 幕，含第 1 次进入历史列表时空状态引导的三种形态（§6.3）。

### 13.4 这一轮之后仍然不存在的东西（不粉饰）

M1 的「闸门」目前只有两个承载：本机能失败的门禁套件，和 `build\RunReleaseChecks.ps1`。仓库里没有流水线定义，所以「不一致就在别人合并时失败」还没发生——它需要用户先决定这条链跑在哪台机器、由什么触发（GitHub Actions / 本地 pre-push / 发版机）。在这之前，把 M1 记成「代码侧就绪、真机与承载待用户」比记成「已过」准确。

---

## 14. E6-02/03/04 任务包合同、一致性门禁与面板归组

**验收标准**（`plan` §7 E6 三行）：E6-02「一个任务包一份 JSON，包能被加载器读出来并拒绝未知字段/未知断言词，而不是靠人记」；E6-03「包→动作→capability→Skill→工具→黄金样例逐级可解析，没有入口的包不许声明，任一级断裂即失败」；E6-04「命令面板里动作按任务包归组，用户看到的是工作目标而不是动词列表；零新消息 type、零新发送路径」。

M1 需要用户坐在机器前、且要先拍板流水线承载（§13.4），所以按「先不用做真机验收，直接接着实现」的指示，本轮转到 Phase 2 的 E6 前三行。

### 14.1 这张表不准有自己的字段

E6 唯一的坑是造出第五份平行清单（前四份是 `QuickActionCatalog`、Skill 的 `allowed-tools`、宿主 capability 注册表、`ToolRegistry`）。Phase 1 的教训（`plan` §9 第 20/22 条）是：任何需要人工同步的副本三周后就是腐烂清单。所以落笔前先逐个读回每个字段的当前持有者，只把**没有持有者**的字段写进合同：

| 字段 | 已有持有者（本轮实测读取处） | 任务包里的形态 |
| --- | --- | --- |
| 入口话术 | `ShareRibbon/Agent/QuickActionCatalog.vb` | 只写 `entryActions` 的 id，加载时逐个比对 `ActionsFor(application)` |
| capability | `WordAi/Services/WordCapabilityRegistry.vb`（Excel/PPT 没有注册表，属开放决策 #36） | 只写 `capabilityIds`；无注册表的宿主必须为空数组，否则门禁失败 |
| Skill 与工具边界 | `ShareRibbon/Skills/<name>/SKILL.md` 的 `name` / `allowed-tools` | 只写 `skill` 目录名；工具边界连字段都不设 |
| 变更事实 | `ShareRibbon/Agent/Execution/AgentChangeLog.vb` 的 `AgentRunChangeLog` | 断言只能引用从该类公开 Integer/Boolean 只读属性反射推导出的事实名 |
| 黄金样例 | `tests/golden/l0/catalog.json` | 只写 case id，门禁查目录 |
| 预览策略、验证器、成功标准、交付物、解释模板、失败恢复、已有实现 | 此前散在提示词与注释里，没有任何可校验载体 | 任务包唯一真正新增的七项 |

因此 schema 里**没有** `allowedTools`、`riskLevel`、`userPhrase` 字段——不是漏了，是它们已经有主人。这条刻在加载器的类注释里，门禁反过来断言注释仍然指向 `QuickActionCatalog`，防止日后有人偷偷把话术抄回包里。

### 14.2 三个包落地，另外三个显式不建壳

| 包 | 入口动作 | capability | Skill | 预览 | 断言 | onFail / recovery |
| --- | --- | --- | --- | --- | --- | --- |
| `word.review-polish` | `word.selection.proofread`、`word.selection.rewrite` | `word.proofread` | `word-document-agent` | `summary` + 写前确认 | `writtenCount gte 1`、`nativeUndoable isTrue` | repair / rollback |
| `excel.clean-and-fix` | `excel.selection.formula`、`excel.selection.format` | 无（#36） | `excel-table-agent` | `summary` + 写前确认 | `writtenCount gte 1`、`changedCount gte 1` | repair / rollback |
| `excel.insight` | `excel.selection.explain` | 无（#36） | `excel-table-agent` | `none`，不要求写前确认 | `writtenCount eq 0`、`hasChanges isFalse` | report / degrade |

报告 §3.2 前 6 包里的 Word 模板排版、PPT 成稿、PPT 品牌模板**没有对应 JSON**。它们在 E6-01/E6-07 之前没有可解析入口，而「没有入口的包不许声明」正是 E6-03 的验收标准；造占位壳就是 `plan` §1 明令禁止的假能力。这个决定写在 `TaskPackageCatalog` 的类注释里，不在门禁里维护"应该有几个包"的名单，只设 ≥3 的下限。

### 14.3 断言只能是代码真的会产生的事实

事实表是反射出来的 7 个属性，不是手写清单：`changedCount`、`writtenCount`、`hasChanges`、`undoableUnits`、`restoreCount`、`nativeUndoable`、`fullyUndoable`（数值 4 个、布尔 3 个，由类型区分）。布尔事实只许 `isTrue/isFalse`，数值事实不许用它们，反向也一样——这条双向规则让「用 `writtenCount isTrue` 蒙混过关」在解析期就报错。

三处按宿主的诚实口径特意分开：

- Word 断 `nativeUndoable`：Word 每一步都应有宿主撤销单元和首尾文档指纹，才配承诺「撤销本轮」。
- Excel 写包**不断** `fullyUndoable`：`ExcelAi/Undo/ExcelRunRestore.vb` 只快照 ≤4000 单元格的值/公式/数字格式，并明确拒绝声称插行插列、新建图表这类结构变更可还原。把它写成断言等于让合同替代码撒一次谎，所以"能不能撤"只留在 `successCriteria` 的文案里，不进断言。
- 只读包断 `writtenCount eq 0` + `hasChanges isFalse`，并配一条跨字段不变式：`previewPolicy.mode=none` 时 `requireBeforeWrite` 必须为 false。「只读却要写前确认」这种自相矛盾的合同，比缺字段更危险，因为它看起来是安全的。

解释模板的 `{...}` 占位符同样只能引用这 7 个事实。这条不是洁癖：摘要文案里出现一个不存在的事实名，运行时就是一句永远不会被替换的模板字面量。

### 14.4 加载器：候选目录只有一套算法，重复读取是预期的

`TaskPackageCatalog.LoadFromRuntimeDirectories` 沿用 `ToolRegistry.LoadFromRuntimeDirectories` 的向上走目录算法（`TaskPackages` 与 `ShareRibbon\TaskPackages` 两套候选 + 去重），避免出现第二套运行时视图。这带来一个必须当场判明的事实：同一次搜索会**同时命中源码目录和构建输出目录**，同一份合同被读两次是正常现象。所以「包 id 已存在」不能直接算冲突，只有同一 `SourceDirectory` 内出现两个同 id 才拒绝——而这种情况已被「文件名必须等于包 id」天然排除，两处规则不矛盾，第二条只是纵深防御。

面板和冒烟脚本只问「这个宿主观得上面的包」，不该各自记得先调一次加载，因此查询入口 `Find/PackagesFor/FindByAction` 前挂一个 `EnsureLoaded()`。这是刻意的取舍：不去改 `BaseChatControl` 的 `AiNativeRuntime` 惰性装配链，代价是第一次查询带一次磁盘扫描，换掉的风险是「合同在磁盘上、运行时视图是空的」这种看不见的空表。

加载期被拒绝的包不静默丢弃：`ValidationErrors` 非空时除 `AppLogger.Info` 外再写一条 Warn 带首条原因，理由是「被拒绝的包等于它在面板上没归组、在验证器上没断言」，不写清楚就是一次静默降级。

### 14.5 一致性门禁：词表从代码反查，不自带名单

`scripts/check-task-package-contract.ps1`（475 行，纯静态）不复制任何 schema，而是按名字从 VB 源码里读回来：`Allowed*` 十二张词表 + `AllowedSchemaVersion` 用正则匹配 `ToSet(...)` 的单行定义，匹配不到就 throw（词表被改写成人看不懂的形态，门禁必须瞎在明处）。逐级解析关系如下：

- 包 → 动作：按 `QuickActionCatalog.vb` 的 `_wordActions/_excelActions/_powerPointActions` 三块取 id 全集；
- 包 → capability：Word 读 `WordCapabilityRegistry.vb`，无注册表宿主强制空数组（把 #36 变成显式约束而不是隐性放任）；
- 包 → Skill：`ShareRibbon/Skills/<name>/SKILL.md` 必须存在，front matter 的 `name` 要等于目录名，`allowed-tools` 要非空；
- 包 → 黄金样例：id 必须在 `tests/golden/l0/catalog.json`；
- 包 → 已有实现：`existingImplementation.reuse` 每条路径必须在磁盘上真实存在——这一条让"先建壳"在经济上不划算，因为壳必须写它复用了谁，而写下的每个文件都得存在。

注册不变式：`TaskPackageCatalog.vb`、`TaskPackageCatalog.Parsing.vb`、`TaskPackageModels.vb` 三个编译项和 `TaskPackages\**\*.json` 的 CopyToOutputDirectory 必须仍在 `.vbproj` 里；`ActionsFor` 交叉校验仍在；Models 注释仍指向目录。合同再漂亮，没人加载就是装饰品。

### 14.6 反空转：三层验证，每层都能自己变红

1. **14 个内存探针**跑在两份干净基线（`excel.clean-and-fix` 写包、`excel.insight` 只读包）上：未知字段 `allowedTools`、幽灵入口动作、杜撰事实、布尔事实配 `gte`、失效 reuse 路径、只读包要求写前确认、Skill 不在磁盘、无入口的包、capability 未注册宿主却填了值、杜撰黄金样例 id、杜撰占位符、id 与文件名不符、`schemaVersion` 漂移、`mode` 越出词表。每个探针都要先证明基线是干净的、再证明变异变红——否则「14 个探针通过」只是 14 句空话。
2. **18 份坏样本喂给编译后的加载器**（反射调用 `LoadDirectory`）：17 份当场被拒、共产生 18 条拒绝消息（`probe-prefix` 那份同时命中「id 前缀≠application」和「入口动作不属于该宿主」两条）。唯一没被加载器拒的是 reuse 路径失效那份——引用能否落到磁盘按分工属于门禁，属于加载器的只有 schema 和入口动作；这条边界由探针里 `dead-reuse-path` 那一支负责变红，本轮实测确实红。同一次注入里放了一份真包做对照，`ACCEPTED=2`（对照包 + 无），证明这批拒绝不是"什么都不认"。
3. **1 次磁盘级变异**：把 `word.review-polish.json` 的 reuse 路径改成不存在的文件，门禁整体变红，随后按字节还原并核对 MD5 一致。前两层证明校验器有效，这一层证明它真的挂在仓库文件上跑。

拆分文件后（见 §14.8 第 1 条）三层全部重跑，结论不变。

### 14.7 面板按任务包归组：零新消息 type、零新发送路径

`CommandPaletteService.BuildState()`（`ShareRibbon/Controls/Services/CommandPaletteService.vb:110`）把状态从一维动作表改成 `packages`（头部：id/title/deliverable）+ `actions`（包内动作优先输出，携带 `packageId`；未被任何包收编的动作照旧下发、不带 `packageId`）。前端 `command-palette.js` 只在 `packageId` 变化且后端给得出标题时插一行 `div.command-palette-group`，方向键游标仍是**一条扁平列表**，不引入「组」这个第二状态机；后端没下发 `packages` 时按全未归组渲染。标题只从后端读，前端不另立包清单——否则就有了第二份归属表，而且是最难发现的那份。

发送侧一字未增：命令面板仍然只发 `startAgent`，快捷键路径、Ribbon 按钮、右键菜单共用同一闸门。`smoke-command-palette.ps1` 的 3b 段把这点钉住：Word 宿主下至少 1 个包头、恰好 2 个动作带 `packageId`、首个归组包是 `word.review-polish`、`word.selection.summarize` 必须仍未归组。

### 14.8 声明的不作为

1. **合同拆成两个文件不是风格问题**：`TaskPackageCatalog.vb` 原为 524 行，越过根规则「单个 class ≤500 行」。按仓库既有 `LoopEngine.Planning.vb` 的做法切成 `Partial Class`：目录/候选/查询留 278 行，严格校验留 263 行。为此门禁新增一条不变式（Parsing 那份必须仍声明 `Partial Class TaskPackageCatalog`），否则一次重构就能把 schema 变成两个持有者。
2. E6-05（验证器在 run 结束时真的被核对）本轮**没有实现**：断言目前只被静态校验其合法性，还没有执行者。计划复用 `LoopEngine.Planning.vb:495` 的 `ValidateObservedOutcome`（它已经把假成功降级成 `CodePartialApply`/`CodeVerifyFailed` 并带着 `RestoreSnapshot`），不新建第二套验证链。
3. E6-06 未做：三个包的 `goldenCases` 全是空数组，门禁输出里的 `golden references=0` 就是这个事实，不是省略。
4. E6-01 当时未做（三端 Ribbon 里仍有硬编码 Agent 请求），已作为下一行落地，见 §15。
5. 没有为 Excel/PPT 造 capability 清单（等 #36），没有把 `riskLevel` 引入合同（风险闸门归 `SafetyGate`）。
6. 没有动 `BaseChatControl` 的运行时装配链，理由见 §14.4。

### 14.9 真机验收清单（按用户指示暂缓，需 Word/Excel 真机 + 重启宿主刷新 `%LOCALAPPDATA%\OfficeAI\www`）

1. 重启 Word，`Ctrl+Alt+K` 唤起命令面板：`审阅润色` 这一行小标题出现在「校对选区」「改写选区」两条动作之上，而其余动作仍平铺、没有小标题。
2. 按 ↑↓ 穿过组头：游标应当在动作之间移动而不在组头停留，Enter 落在被指的动作上。
3. 输入一句没有任何动作命中的自然语言：列表只剩「按原话执行：…」一条，按 Enter 后面板进入 Agent 回合（证明归组没有改变发送路径）。
4. 把 chord 换成未占用组合并在设置里保存，重开宿主后仍能唤起（E4-03 的既有链路，本轮未改）。
5. 重启 Excel，命令面板小标题应为「清洗与公式修复」「表格问答与洞察」，而不是 Word 的两组（跨宿主串组是本次改动最直接的失败形态）。
6. 面板打开状态下查看 `%LOCALAPPDATA%\OfficeAI\www\js\command-palette.js` 与 `css\styles.css` 的时间戳：晚于本轮改动，否则是在验证旧副本。

### 14.10 本轮实测数字（全部为本次重新运行所得）

- `scripts\build-code-projects.ps1 -Configuration Debug`：`PASS: ShareRibbon / WordAi / ExcelAi / PowerPointAi`，`Code projects built successfully.`
- `scripts\check-task-package-contract.ps1`：`PASS task-package contract: 3 packages, 6 assertions, 15 schema keys, 7 facts, 14 probes turned red on 2 clean baselines, golden references=0`
- 反射加载 `bin\Debug\ShareRibbon.dll` 跑 `LoadFromRuntimeDirectories`：`LOADED=3 ERRORS=0 WORD=1 EXCEL=2`，事实表实到 7 项；`TaskPackageCatalog.vb` 278 行、`.Parsing.vb` 263 行、`TaskPackageModels.vb` 96 行，均在 500 行以内
- `scripts\smoke-command-palette.ps1`：`Command palette smoke passed.`（含新的 3b 归组断言）
- `scripts\run-code-checks.ps1 -SkipBuild`：`discovered 24 gates under scripts\`（新门禁自动被目录发现，未登记进 `$notGates`）、47 个 `PASS`、0 个 FAIL、`All code checks passed.`；其中 Golden L0 为 12 例（本轮未增删）
- `node scripts\check-entry-reachability.js`：消息 type 面板 56 / router 注册 66；内联事件 70 处引用、57 个函数名、70 处解析到定义；DOM id 103 个字面量查询、103 个解析到节点；注入 id 16/16；注入调用 124 段载荷、61 个调用名全部解析或被 `typeof` 守卫；导航文档 11 份、38 条本地引用、38 条解析到磁盘；Ribbon 可见入口 Word 14/18、Excel 13/18、PowerPoint 11/18（与 §7.1 一致，本轮未动 Ribbon 可见性）

## 15. E6-01 三端 Ribbon 话术迁入唯一清单

### 15.1 收的是哪十处

迁移前，Agent 请求文本写死在各宿主 `Ribbon1.vb` 的按钮 Click 里：Word 1 处（`WordAi/Ribbon1.vb:27` 数据分析）、Excel 6 处（`ExcelAi/Ribbon1.vb:17/60/156/160/219/226`：抓取网页、数据分析、校对、排版、续写、模板排版）、PowerPoint 3 处（`PowerPointAi/Ribbon1.vb:18/25/46`：抓取网页、数据分析、校对）。这十处与 E4 已经收口的选区浮层/右键菜单是同一类东西——「入口持有的话术」，只是没被 E4 覆盖，于是成了清单的第二个持有者：改一句请求要在四个文件里考古，而没有任何检查会响。

现在十句原文一字未改地进了 `ShareRibbon/Agent/QuickActionCatalog.vb:111-176`（10 条 `Public Shared ReadOnly <Host>Ribbon<Button> As QuickAction`，文件 250 行仍在 500 以内），`Label` 取自 Ribbon 设计器里的真实按钮文字而不是重新拟的文案。清点过程本身也留一笔：动手前按 9 处记，改完做全仓 grep 才发现漏了 `ExcelAi/Ribbon1.vb:17` 的抓取网页——**「我记得有几处」在这类收口里不算证据**。

### 15.2 用类型而不是查表，并把它证伪一次

三宿主的 `StartAgentFromRibbonAsync` 形参由 `request As String` 改为 `action As Agent.QuickAction`（`WordAi/Ribbon1.vb:460`、`ExcelAi/Ribbon1.vb:229`、`PowerPointAi/Ribbon1.vb:416`），函数体只在 `chatCtrl.StartAgentFromHostEntry(action.RequestText)` 用到文本（`:471`/`:240`/`:427`）。没有做成 `QuickActionCatalog.Find("word.ribbon.analyze")` 那种字符串查表：查表把「入口不再持有第二份话术」降为一条约定，签名把它变成类型错误。

这条不是声称，是做了一次注入验证：把 `WordAi/Ribbon1.vb:27` 换回字面量字符串后构建，MSBuild 报 `error BC30311: Value of type 'String' cannot be converted to 'QuickAction'`（位置 `(27,41)`），随后按 MD5 还原源码（前后同为 `79ac48daef722e565fc41980795a83e6`）并重跑构建，四项目 `PASS`。验证完的 DLL 也重建过，避免 §8 那次「DLL 比源码旧、白跑一轮」。

### 15.3 为什么这十条不进 `ActionsFor`

`ActionsFor` 喂的是选区浮层、右键菜单与命令面板（E4/E6-04 的运行时视图）。把这十条按钮文案一并返回，会当场改变面板行数——那是用户可见的产品改动，属于 E6-07 与开放决策 D9 要拍板的事，不该由一次「把文字搬走」的重构顺带做掉。所以 E6-01 只搬文本、不动行集，并把这条边界写成门禁规则而不是注释：ribbon id 一旦出现在 palette 宇宙就判失败（探针 `ribbon-id-in-palette`）。

顺带量到的事实：这十条对应的按钮有 9 条在设计器里是 `Visible = False`（8 条在宿主 `Ribbon1.Designer.vb`，`TemplateFormatButton` 在 `ShareRibbon/Ribbon/BaseOfficeRibbon.Designer.vb:183` 全局隐藏），唯一默认可见的是 Excel 数据分析。也就是说本轮收口的是一批「已实现但看不见」的入口，D9 的代价因此更可判定：文案已经只有一份，开不开按钮不再需要同步改代码。

### 15.4 门禁：围绕一条签名的闭包

`scripts/check-task-package-contract.ps1:410-565`。类型保证在编译器那边，静态门禁负责它周围会烂掉的部分：

- 每条宿主源里必须仍出现 `StartAgentFromRibbonAsync(action As Agent.QuickAction)`（签名退回 `String` 就是闸门消失）；
- 任何 `StartAgentFromRibbonAsync("` 形态的字面量直接判失败；
- 被引用的成员必须是声明过的 ribbon 话术，且前缀必须属于引用它的宿主（`Word`/`Excel`/`Ppt`）；
- 声明了却没人引用 = 装饰，声明数必须等于引用数（一个按钮一条话术）；
- ribbon id 必须是 `<host>.ribbon.<button>` 形状，且不得与 palette id 同集。

反空转靠 7 条内存探针（同一函数判定，干净状态先过再变异）：字面量回来、签名退回 String、引用未声明成员、一条话术被两个按钮复用、引用别宿主的话术、id 形状漂移、id 泄进 palette。

### 15.5 这套规则顺手逮到的旧假设

`scripts/smoke-selection-toolbar.ps1` 原先把 catalog 里所有 `word.` 前缀的 id 当作选区动作清单，第一次跑全量门禁就报「expects exactly 3 selection actions, found 4」，多出的正是 `word.ribbon.analyze`。这不是本轮引入的缺陷，是那条 smoke 的「id 前缀 = 那份清单」假设本来就站不住：只要同一个类里再放任何 `word.*` 就会误判。已改为按 `_wordActions` 块切片读取（`:52-64`，与 E6-03 门禁同一个「下一个访问修饰符即块尾」的切法），断言本身（恰好 3 条、必须 `word.selection.*`）没放松。

### 15.6 声明的不作为

1. 没有改变任何按钮的可见性、位置或点击行为：`check-entry-reachability.js` 的 Ribbon 可见入口仍是 Word 14/18、Excel 13/18、PowerPoint 11/18。
2. 没给 ribbon 动作填 `CapabilityId`（Excel/PPT 侧仍无 capability 清单，等 #36），也没把它们写进任务包 `entryActions`——包只能引用 palette 里的入口动作，这是 E6-03 定的。
3. 没做 E6-05/06：断言仍未在 run 结束时被核对，三包 `goldenCases` 仍空（门禁仍输出 `golden references=0`）。
4. 没把 `RequestText` 改成「按宿主上下文自动补齐」，仍是整句交给同一条主链，由 Agent 自己读选区/文档；`grep 'StartAgentFromRibbonAsync("'` 三端为 0，也没有新增第二套发送路径。

### 15.7 真机验收清单（按用户指示暂缓，需 Excel/Word 真机）

1. 重启 Excel，点「数据分析」：面板应发起一轮 Agent 回合，请求文本与 `excel.ribbon.analyze` 的 `RequestText` 逐字一致（本轮只搬文字，任何措辞差异都是迁移时的手误）。
2. 该按钮是唯一默认可见的，其余 9 条要在 D9 之前验证需临时改 `Visible`；建议改走 E4-02 右键菜单验证同一主链，不改设计器。
3. Word 侧确认「数据分析」按钮仍不可见、而选区浮层/右键菜单行数为 3（证明 ribbon 话术没有漏进 palette 宇宙）。
4. 重启宿主后核对 `bin\Debug\*.dll` 的时间戳晚于本轮改过的 3 个 `Ribbon1.vb` 与 `QuickActionCatalog.vb`，否则这一轮验证的是旧副本。

### 15.8 本轮实测数字（全部为本次重新运行所得）

- 迁移点：10 处 Ribbon 字面量 → 10 条 catalog 话术；3 处 helper 签名改为 `action As Agent.QuickAction`；引用数 Word 1 / Excel 6 / PowerPoint 3，与声明数 10 相等
- 注入验证：`WordAi\Ribbon1.vb(27,41): error BC30311`，源码 MD5 前后一致，随后 `PASS: ShareRibbon / WordAi / ExcelAi / PowerPointAi`、`Code projects built successfully.`
- `scripts\check-task-package-contract.ps1`：`PASS task-package contract: 3 packages, 6 assertions, 15 schema keys, 7 facts, 14 probes turned red on 2 clean baselines, golden references=0, ribbon utterances=10 at 10 button(s)`
- `scripts\smoke-command-palette.ps1`：`Command palette smoke passed.`（行集未变，故 3b 断言无需调整）
- `scripts\run-code-checks.ps1 -SkipBuild`：`discovered 24 gates under scripts\`、47 个 `PASS`、0 个 FAIL、`All code checks passed.`（第一次跑到 `smoke-selection-toolbar` 变红，即 §15.5 那条旧假设，修好后全绿）
- `node scripts\check-entry-reachability.js`：导航路径从 §14.10 的 38 条升到 42 条（本轮为 AGENTS.md 新增两行导航，42 条全部解析到磁盘）；Ribbon 可见入口计数与 §14.10 一致
- 规模：`QuickActionCatalog.vb` 250 行、`check-task-package-contract.ps1` 634 行

---

## 16. E6-05 包声明的成功标准在 run 收尾时真的核对

**验收标准**（`plan` §7 E6-05）：「包声明的成功标准在 run 结束时被真的核对（复用 `ValidateObservedOutcome` 与 `agent_run_change` 事实），失败进入既有 repair loop 而不是直接宣称成功」。真机部分按「先不用做真机验收」暂缓，清单见 §16.8。

### 16.1 §14 造出来的断言此前没有人执行

三个包共 6 条断言（`writtenCount gte 1`、`nativeUndoable isTrue`、`writtenCount eq 0` 等）在 E6-03 之后只被静态校验「填得合不合法」：词表对、事实名对、宿主口径对，然后就没有然后了——一轮 Agent 跑完，收尾时没有任何一处把台账和这些断言比一遍。这不是「差最后一步接线」，而是这类合同最常见的死法：字段齐全、门禁全绿、运行时无人读。所以本轮的验收不是「断言能跑」，而是「假成功会被拦下来」，并且拦下来之后仍然只有一条终路径。

前置条件只有一个，而且此前不存在：**这一轮是哪个入口动作发起的**。断言属于包，包靠 `entryActions` 绑定动作；一轮 run 如果不知道自己从哪个动作来，就没有任何东西能证明它该受哪个包的合同约束。

### 16.2 入口身份是参数，不是话术

原来 `BaseChatControl.StartAgentFromHostEntry(request As String)` 只带文本。文本不能反查包——从「帮我把这段润色一下」猜出 `word.selection.proofread` 是一次 NLP 分类任务，猜错的后果不是「没验证」，而是**把别人的成功标准扣到用户头上**（用户要洞察，系统按写包判定「你一个字都没改，未完成」）。所以身份必须由入口自己声明，收口方式沿用 E6-01 的「用类型而不是约定」：

- 形参改为 `action As Agent.QuickAction`（`ShareRibbon/Controls/BaseChatControl.vb:1769`），话术和 id 同源，取自同一个对象；
- payload 增 `entryActionId`（`:1773`），`HandleStartAgent` 读回（`:1806`），两条分支都必须显式带到 `HandleStartAgentCore`（`:1845/1847/1982`）和带文件解析路径（`:1859/1962`），启动时交 `StartAgentAsync(..., entryActionId)`（`:2010`）；
- 这两个下游方法的 `entryActionId As String` **不是 Optional**：漏传是编译错误，不是审查时才发现的约定；
- 自由文本/聊天输入路径显式写 `""`（`:1731-1732`），意思是「我知道我在问什么，这条没有合同」，而不是「我没想好」；
- 命令面板仍只发一条 `startAgent`（`ShareRibbon/Resources/js/command-palette.js:218`），零新消息 type、零新发送路径；自由文本行把 `__natural_language__` 映射成空串（`:217`），否则占位符会被当成一个真实入口 id 去查包；
- 链路往下：`AgentKernelService.StartAgentAsync(..., Optional entryActionId)`（`Services/AgentKernelService.vb:106`）→ `UserTurn.EntryActionId`（`Harness/HarnessModels.vb:26`）→ `OfficeHarness.vb:127` → `AgentKernel.ExecuteAsync`（`AgentKernel.vb:208`）→ `AgentSession.EntryActionId`（`:218`，字段在 `AgentSession.vb:22`）。

两处容易被忽略的细节。**其一**，UI 线程赋值与后台 `StartAgentAsync` 之间隔着线程切换，身份按每次调用显式携带而不是只放服务字段；`entryActionId Is Nothing`（调用方忘了传，写一条 Warn 到日志）与 `""`（自由文本，本就没有合同）是两种不同的事实，合并成一种就会把「漏接的入口」悄悄变成「合法的自由文本」。**其二**，`FinalizeAgentUi` 的 Finally 与 `AbortAgent` 都把服务字段复位，否则下一轮继承上一轮的包，用户问洞察却按写包判定未完成。三宿主 7 处调用点（Word 3、Excel 2、PowerPoint 2）全部改为交 `QuickAction` 对象，`grep 'StartAgentFromHostEntry(request'` 全仓 0 处。

### 16.3 变更事实仍然只有台账能报

验证器最容易变成第二份真相的地方是「自己数记录」。本轮把读取能力放回持有者：`AgentRunChangeLog.TryFactValue(fact, value)`（`ShareRibbon/Agent/Execution/AgentChangeLog.vb:162`）用反射读本实例的**公开、非索引、Integer/Long/Boolean 只读**属性，词表之外的事实读不出（`Nothing` + False），PascalCase 原名也读不出。命名规则同样只有一个持有者：`ToFactName`（`:179`）同时供 `TaskPackageCatalog.BuildVerifiableFacts`（`TaskPackageCatalog.vb:89`）使用，目录里那份私有 `ToCamel` 已删除——否则 E6-03 门禁反向推导出来的词表和运行时能读的名字会各自漂移，而漂移的表现正是「这条断言永远读不到事实」。

新增 `ShareRibbon/Agent/TaskPackageOutcomeVerifier.vb`（104 行，已加入 `ShareRibbon.vbproj` 编译项）。判定一律 fail-closed：事实名空、读不出、比较符不在词表、数值比较遇到布尔、`isTrue/isFalse` 遇到数值、本轮没有台账，每种都返回一条**点名是哪个事实**的原因；只有真的满足才返回空串。文案带上声明值与包里的 `why`，用户看到的是「`writtenCount gte 1`：实际 0（……为什么要求写过）」，而不是一句空泛的未完成。

### 16.4 判定落在既有收尾槽位的末尾，终态仍只有一个

`ValidateExecutionOutcome`（`LoopEngine.Coverage.vb:162`）本来就是「假成功」的最后一道闸——要求插图却没插、要求 N 页只做了 M 页都在这里拦。包断言是同一类判断，所以接在它末尾（`:189-191` 调 `ValidateTaskPackageOutcome`，函数体 `:199-213`），共用 `LoopEngine.vb:479-485` 那一条 `AgentResult.Failed(...)` 出口，没有新建第二套 verifier/loop/终态。没有入口身份就不核对（`:200`）；返回值仍走 `AppendWriteSummary`，所以「已经落地的改动」不会因为一句未完成而被吞掉——用户要同时知道没达成什么和已经改了什么。

### 16.5 反空转：15 条运行时用例 + 14 条源码变异探针

新门禁 `scripts/smoke-task-package-verifier.ps1`（172 行，由 `scripts\*.ps1` 自动发现，未进 `$notGates`）反射加载编译产物 `ShareRibbon.dll`，用 4 份真实 `AddFromStep` 载荷造出台账，把 3 个真包的断言喂进去。15 条用例同时覆盖两个方向：**满足→0 条原因**（`excel.clean-and-fix` 正常写入、`excel.insight` 保持只读、`word.review-polish` 写过且有原生撤销点）和**不满足→点名事实**（回滚后写次数归零、只读包发生写入、没有撤销栈、什么都没写、6 条 fail-closed 词表/类型/空台账用例），外加把全部 7 条事实逐个 `TryFactValue`、证明 PascalCase 读不出且空台账 `writtenCount=0`。只测「能变红」的验证器和只测「能变绿」的门禁一样不可信。

源码级闭包在 `scripts/check-task-package-contract.ps1:566-778`：12 个文件、约 20 条规则、14 条变异探针，每条都必须变红且必须被**它对应的那条规则**抓住（会话丢字段／harness 丢参数／内核从话术猜身份／宿主入口退回 String／聊天输入路径继承上一轮身份／面板不发 id／面板把占位当真 id／收尾跳过包／验证器自己数事实／数值分支永不满足／布尔分支永不满足／目录自带命名规则／`TryFactValue` 放开类型闸门／验证器不进 `.vbproj`），再加一条 `entryActionId` 出现次数下限（≥6，实测 10），防止「规则全对但链路已断」。

本轮第一次跑门禁就红了一次，报的正是 `identity probe 'verifier-never-satisfied' slipped through`：初版规则是 `verifier -notmatch 'Return ""'`，探针改掉数值分支的 satisfied 返回值后仍然绿，因为布尔分支里还留着另一条 `Return ""`。一条 grep 挡不住「两个 satisfied 分支坏掉一个」，已改为分别钉住两条分支的字面量并各配一条探针（探针 13→14）。这条记在这里，是因为它和 §14.6 那次「加载器什么都拒」的假绿是同一类错误。

### 16.6 `onFail: repair` 目前没有人接：如实记为 E6-05b

> **本节两条更正**：(1) 「三个包的 `onFail` 都写的是 `repair`」是错的，盘上实测 `excel.clean-and-fix`/`word.review-polish` 是 `repair`，`excel.insight` 是 `report`——这个错误是我给 E6-05b 写门禁前置检查时，被检查本身按真实包数据驳回才发现的（详见 §18.4）；(2) 「路由尚未接通」已随 §18 的 E6-05b 作废，下面的理由保留是为了记录当时的判断依据，不是当前状态。

按更正后的事实：两个写包声明 `repair`，只读包声明 `report`。run 收尾当时**不能**重进规划循环：`LoopEngine.vb` 的步进循环约 360 行，在 end-of-run 之后再包一层 re-plan 会牵动 replan 幂等（E1-06）和写入摘要（E1-04），而这些路径本轮没有真机在手。硬写会把「按合同去修」变成一次没有验证的重构，风险由用户文档承担。所以当时的处置是拦下假成功并写一条 Warn（`LoopEngine.Coverage.vb:207-211`，消息里明说 `run-end repair route is not wired`），接通这条路由记为 **E6-05b**。验收标准里「失败进入既有 repair loop 而不是直接宣称成功」的后半句已达成，前半句没有——把它写成达成就是这份报告最不该犯的那种错。

### 16.7 声明的不作为与代价

1. 没有新建第二套 Harness/Planner/Loop/Safety/verifier，没有新增消息 type 或发送路径；判定挂在既有 `ValidateExecutionOutcome`，终态仍只有一处。
2. 没有让验证器自己数 `agent_run_change` 记录，也没有把 `riskLevel`/工具边界塞回包（分别属于 `SafetyGate` 和 Skill 的 `allowed-tools`）。
3. 没有改 `ActionsFor` 行集与按钮可见性：Ribbon 可见入口仍是 Word 14/18、Excel 13/18、PowerPoint 11/18。
4. 没有做 E6-06：三包 `goldenCases` 仍空，门禁仍输出 `golden references=0`；本轮 §16.5 的 15 条用例是验证器自身的回归，不是黄金样例，两者不能互相顶替。（下一条 §17 补上，当时的如实状态。）
5. 没有为 `BaseChatControl.vb`（3370 行）做搬迁，本轮只加参数与注释；该类仍超「单 class ≤500 行」，是历史欠账，不是本轮引入。
6. **代价一**：自由文本发起的同一件事不受任何包合同约束。面板输入「帮我把这段润色一下」不核对，右键菜单「审阅润色」核对——这是「不猜包」换来的确定性，D9/E6-07 把入口补齐后应重新评估。
7. **代价二**：`writtenCount gte 1` 会让「用户要求润色，Agent 判断确实无需改动」这类合法的空写入被判未完成。合同没说谎（包确实要求写过），但这可能不是用户想要的答复；§16.8 第 2 条专门验证这条文案的观感，其结论决定 E6-06 的失败样例怎么写。

### 16.8 真机验收清单（按用户指示暂缓，需 Word/Excel 真机 + 重启宿主刷新 `%LOCALAPPDATA%\OfficeAI\www`）

1. Word：选区有可润色文本，走右键菜单「审阅润色」，成功收尾且日志里**不出现**「成功标准未达成」；`agent_run_change` 该 run 应有 ≥1 条写入、`nativeUndoable` 为真。
2. Word：同入口但选区无需改动，预期看到 `任务未完成（word.review-polish 的成功标准未达成）：writtenCount gte 1：实际 0（……）`。确认这是合同而非 bug，并判断这句答复对用户是清楚的还是误导的（§16.7 第 7 条）。
3. Word：写入后点「撤销本轮」，再发起同一入口的一轮，观察摘要仍列出本轮已落地的改动（`AppendWriteSummary` 未被未完成消息吞掉）。
4. Excel：`excel.insight`（只读）跑一次，确认台账 `writtenCount=0`、`hasChanges isFalse` 且**没有**误报未完成。
5. Excel：`excel.clean-and-fix` 制造「执行了但数据本来已经干净」的场景，应报未达成而不是成功。
6. 面板输入自由文本发起 4/5 同一件事，确认日志里不出现任何包 id（证明「不猜包」在真机上成立）。
7. 真机日志搜索 `StartAgentAsync called without entry identity`：若出现，说明有一条入口路径没接上身份，按 §16.2 链路补齐，而不是把 Warn 降噪。
8. 重启宿主后核对 `bin\Debug\*.dll` 时间戳晚于本轮改过的 `.vb`/`.js`，并核对 www 副本已刷新，否则验证的是旧副本（见 `Virtual server mode` 记忆）。

### 16.9 本轮实测数字（全部为本次重新运行所得）

- 身份链路：`entryActionId` 在 `BaseChatControl.vb` 出现 10 处；三宿主 7 处调用点交 `QuickAction`；`grep 'StartAgentFromHostEntry(request'` 全仓 0 处；链路上 12 个文件、约 20 条规则、14 条探针
- 规模：`TaskPackageOutcomeVerifier.vb` 104 行（新文件，已进 `.vbproj`）、`smoke-task-package-verifier.ps1` 172 行、`check-task-package-contract.ps1` 634→856 行；`AgentChangeLog.vb` 425 行、`LoopEngine.Coverage.vb` 325 行、`TaskPackageCatalog.vb` 274 行、`AgentKernel.vb` 371 行、`AgentSession.vb` 222 行、`AgentKernelService.vb` 567 行
- `scripts\smoke-task-package-verifier.ps1`：`PASS task package verifier: 3 real packages, 15 assertion cases, 7 facts readable at runtime`
- `scripts\check-task-package-contract.ps1`：`PASS task-package contract: 3 packages, 6 assertions, 15 schema keys, 7 facts, 14 probes turned red on 2 clean baselines, golden references=0, ribbon utterances=10 at 10 button(s), 14 entry-identity probes turned red at 10 panel site(s)`
- `scripts\run-code-checks.ps1 -SkipBuild`：`discovered 25 gates under scripts\`（24→25，新 smoke 被目录自动发现）、50 个 `PASS`、0 个 FAIL、`Golden L0 passed: 12 cases`、`All code checks passed.`
- 四项目 `PASS: WordAi / ExcelAi / PowerPointAi`（`ShareRibbon` 随其一并重建）、`Code projects built successfully.`
- 一次真实红灯：`identity probe 'verifier-never-satisfied' slipped through`（satisfied 分支只钉住一条 grep），修法与结论见 §16.5
- `node scripts\check-entry-reachability.js`：面板发送 56 个消息 type、DOM id 103/103、注入调用 61/61、导航文档 11 份、本地引用 42→45 条（本轮为根 `AGENTS.md` 补了两行导航：一致性门禁补入口身份闭包、新增「任务包断言的执行者」一行），45 条全部解析到磁盘；Ribbon 可见入口 Word 14/18、Excel 13/18、PowerPoint 11/18（与 §15.8 一致，本轮未动入口行集）

## 17. E6-06 黄金样例：把每个包的「满足长什么样」钉到磁盘上

### 17.1 这一行不是「多跑几个 case」

E6-05 之后断言真的会执行，但执行器的正确性只由 `smoke-task-package-verifier.ps1` 的 15 条合成载荷证明——那批载荷是我为了验证器造的，不是从包语义推出来的。如果包自己的口径写错（例如把「发生过的写」当成「留下来的写」），验证器越准就越忠实地核对一个错误的合同。黄金样例要补的正是这一层：**每个包在磁盘上留一份「满足的 run 长什么样」和一份「不满足的 run 长什么样」**，两份都用宿主营运时真会给出的 `ToolResult` 观察载荷形状来写，并且由 runner 用真实 `AddFromStep` 造台账、用真实 `TaskPackageOutcomeVerifier.UnmetReasons` 判定。数据与执行两边都得是真的，否则就又造了一个「字段齐全、门禁全绿、无人读取」的壳。

### 17.2 只有 3 个包，不是 6 个——这条如实写在开头

计划行叫「前 6 包黄金样例」，报告 §3.2 的前 6 包是 Word 审阅润色、Word 模板排版、Excel 清洗修复、Excel 洞察、PPT 成稿、PPT 品牌模板。磁盘上只有 3 份 JSON：另外 3 份在 E6-07 与开放决策 D9 之前没有可解析入口，而「没有入口的包不许声明」正是 E6-03 的验收标准（§14.2），先造包壳再造样例等于把那条标准撤掉。所以本轮把**存在的 3 个包**每个都补齐 ≥1 成功 + ≥1 失败，共 7 条 `task_package_outcome` 用例；第 4~6 包的黄金样例随 E6-07 一起落，这一行的完整验收没有在本轮达成。

### 17.3 数据侧：`tests/golden/l0/catalog.json:91-244`

12 条用例 → 19 条。新增的 7 条每条带 `packageId`、`entryActionId`、`expectedOutcome`、`expectedUnmetCount`、`expectedUnmetFacts`、`canUndo`、`note`（中文说明这条在防什么）与 `steps[]`；步骤里的 `observation` 就是宿主给 `ToolResult` 的原始形状（`kind`/`changed`/`summary`/`undoUnits`/`rollback.succeeded`/`before.documentTextHash`/`after.documentTextHash`），不额外发明字段——验证器读不到的字段在这里出现，就属于 §16.7 明令禁止的「包自己数事实」。

三条失败样例各钉住一种最容易糊过去的口径：

1. `L0-word-review-polish-fail-no-write`：`changed:false` 且首尾指纹相同 → `writtenCount` 与 `nativeUndoable` 两条未达成。这条是 §16.7 第 7 条代价的可执行版本：合同要求写过、run 一个字没改，就必须报未完成，而不是因为「没有变化所以没失败」而静默通过。
2. `L0-excel-clean-and-fix-fail-rolled-back`：写过但 `rollback.succeeded:true` → **只有** `writtenCount` 未达成，`changedCount` 仍为 1。这是本轮最重要的一条口径：`changedCount` 数「发生过的写」，`writtenCount` 数「留在文档里的写」。若断言读前者，用户点过「撤销本轮」的 run 会被判成功——合同替一个已经回退的改动背书。
3. `L0-word-review-polish-fail-no-undo`：每步都落地但没有 `undoUnits` → 只报 `nativeUndoable`，证明布尔事实与写次数各自独立变红，不会互相掩盖。

两条只读包样例：`L0-excel-insight-pass` 的唯一一步是 `kind:"read"` + `recorded:false`，验证器面对的是**空台账**而不是「有记录但都不算写」；`L0-excel-insight-fail-wrote` 用 `WriteData`（`changed:true`）报 `writtenCount`+`hasChanges` 两条，证明只读包的合同真的能反过来抓越界写入。

### 17.4 执行侧：`scripts/run-golden-l0.ps1:25-50` 与 `:53-96`

`New-TaskPackageChangeLog` 用编译产物里的真实 `AgentRunChangeLog.AddFromStep` 造台账，并**当场核对收录口径**：某步声明会被收录而 `AddFromStep` 返回 `Nothing` → 报「写观察合同漂移了」；某步声明不该收录却返回了记录 → 报「非写操作进了台账，只读的一轮会谎称改过文档」。这条双向检查不是多余的——§17.3 里那些 `recorded` 标记如果只由静态门禁凭规则推断，一旦 `AddFromStep` 的 `kind`/`changed` 口径改动，样例就会安静地描述一个不再存在的台账。分支体里 `FindByAction(entryActionId)` 解析包并比对 `package.Id`，再把 `UnmetReasons` 的条数、极性和点名事实与期望逐条对齐；`expectedOutcome` 只认 `satisfied`/`unmet`，第三种取值直接 throw。

### 17.5 覆盖检查顺手逮到一个真缺陷：旧构建副本会冒充当前合同

`:275-297` 的覆盖块跑完用例再要求「每个已加载的包 ≥1 pass + ≥1 fail，且包声明的 `goldenCases` 条数等于 catalog 里归属它的条数」。它第一次变红报的是：

```text
task package excel.clean-and-fix declares 0 golden case(s) but the catalog holds 2 for it
```

原因不在数据。`TaskPackageCatalog` 的 id 去重是「先加载的那份赢」，而 `LoadFromRuntimeDirectories` 的向上目录搜索会同时命中源码目录和 `bin\Debug\TaskPackages` 的构建副本——那一刻副本里还是旧内容。于是运行时视图里的合同是我没读过、也没写过的旧 JSON，任何基于合同的判定（门禁、验证器、黄金样例）都在绿灯下核对错的东西。这不是「脚本记得先 build」的纪律问题：一份没人读过的旧合同能冒充当前合同，正是 §14.2「避免第二套运行时视图」想要防到底的那类漂移。

修在持有者里：`ShareRibbon/Agent/TaskPackageCatalog.vb:71` 增 `_sourceTexts`，`:236` 把每个包规范化成一份比较用正文，`:241` 同目录重复仍按原规则拒绝，`:247-248` **跨目录同 id 而正文不同即 `AddError`**，`:254-255` 记录首次加载的目录与正文，`:181` 的 `Reset()` 一并清空。门禁反向钉住这两条不变式（`scripts/check-task-package-contract.ps1:560-564`，正则匹配不到直接 throw），防止下次重构把它当冗余代码删掉。

### 17.6 反空转：静态两侧规则 + 11 条内存探针 + 2 次盘上真实变异

**包侧**（`check-task-package-contract.ps1:326-347`）：每条 `goldenCases` 引用必须恰好对应 1 条 `task_package_outcome` 用例且 `packageId` 相同；每包必须同时存在 `satisfied` 与 `unmet` 用例，否则报「没有任何东西能证明这条合同可能被满足 / 可能变红」；声明条数与归属条数必须相等。

**用例侧**（`:425-509` 的 `Test-GoldenCases`）判 12 类：重复 id、`packageId` 指向盘上不存在的包、`entryActionId` 不属于该包、`expectedOutcome` 越词表、`satisfied` 却仍数事实或条数非 0、`unmet` 条数为 0 或不点名任何事实、事实名词表外、缺 `canUndo`、`steps` 为空、步骤缺 `toolId`/`observation.kind`、`recorded` 与 `kind` 互相矛盾（读操作声称入台账 / 写操作声称不入）、被收录的步骤缺 boolean `changed`。

**11 条内存变异探针**各对应上面一条规则：`golden-ghost-package`、`golden-wrong-entry`、`golden-outcome-outside-vocab`、`golden-satisfied-still-counts`、`golden-invented-fact`、`golden-unmet-without-facts`、`golden-hides-can-undo`、`golden-empty-steps`、`golden-read-counted-as-write`、`golden-write-hidden-from-ledger`、`golden-duplicate-id`。

**盘上真实变异**两次，都跑真脚本而不是内存副本：

1. 只改源码 `excel.insight.json` 的一条断言（`writtenCount eq 0` → `gte 1`）**不重构建** → `run-golden-l0.ps1` 当场红：`task package(s) rejected while loading: ... 同一任务包 id 在不同目录里内容不一致`。§17.5 的修复在真实漂移场景下有效，随后按 MD5（`a4bb5f7b…`）还原并复跑为绿。
2. 改 catalog 一条 `expectedUnmetCount` 2→1 → 红，且输出如实点名两条未达成事实，证明 runner 不是「跑完就打印 PASS」。

第二次还原时我用了「把第一个 1 换回 2」这种反向字符串替换，误伤了另一条本来就期望 1 的用例（`fail-no-undo`），当场又被同一条 runner 红灯抓出，最后改为按行号精确还原并复跑。**教训入册**：变异探针的还原必须按整文件字节/MD5 或按行定位，不能靠反向 replace——仓库在 §14.7 已经有「按字节还原并核对 MD5」的规矩，这次是我自己绕开它；被绕开的后果是往测试数据里注入了一个错误的期望值，而不是改坏代码。这里也算顺带证明了黄金 runner 的非空转性：它抓住的正是手工写错的期望数字。

### 17.7 不作为与代价

1. 第 4~6 包及其黄金样例未做（等 E6-07/D9），「前 6 包」这行的完整验收**未达成**，本轮只覆盖磁盘上存在的 3 包。
2. 没有新增门禁文件：所有规则加进已有的 `check-task-package-contract.ps1` 与 `run-golden-l0.ps1`，门禁数仍 25，避免「门禁清单本身变成第二份真相」（§13 那条）。
3. 没有把黄金样例落库、写进 `agent_run_change` 或任何 History/Memory：它们是测试数据，只走反射内存台账。
4. 没有为 `observation` 另造 schema 校验器：载荷形状仍由 `ToolResult`/`AgentChangeLog` 的既有收录口径决定，`recorded`↔`kind` 一致性检查只保证样例和台账不互相撒谎。
5. 代价：`canUndo` 是必填。写用例的人必须显式表态这轮宿主有没有撤销栈——好处是不会出现「默认 true 让 `nativeUndoable` 用例假绿」，代价是加一个非撤销宿主的样例时会先被门禁拦一次。
6. 代价：黄金 runner 现在会**拒绝加载**任何跨目录内容漂移的包，等于「改完 `TaskPackages/*.json` 不重建就直接跑 gates 会红」。这是响亮而非便利；若有人嫌烦去删那条加载器检查，旧合同冒充当前合同的漂移会退回静默——该改的是流程（改包即重建），不是那道闸门。
7. 真机按用户指示继续暂缓：§16.8 的 8 条清单一条未跑，E6-08（六包真机完成率 ≥60%）依赖它，而 §17.2 那 3 个包连包都还不存在。

### 17.8 本轮实测数字（全部为本次重新运行所得）

- `tests/golden/l0/catalog.json`：92→247 行、12→19 条用例（新增 7 条 `task_package_outcome`）；三包 `goldenCases` 由 `[]` 填为 3 / 2 / 2 条 id
- 规模：`run-golden-l0.ps1` 201→299 行、`check-task-package-contract.ps1` 856→1033 行、`TaskPackageCatalog.vb` 274→288 行（`TaskPackageOutcomeVerifier.vb` 104 行未变）
- `scripts\run-golden-l0.ps1`：`Golden L0 passed: 19 cases`
- `scripts\check-task-package-contract.ps1`：`PASS task-package contract: 3 packages, 6 assertions, 15 schema keys, 7 facts, 14 probes turned red on 2 clean baselines, golden references=7 (7 outcome cases, 11 golden probes turned red), ribbon utterances=10 at 10 button(s), 14 entry-identity probes turned red at 10 panel site(s)`
- `scripts\smoke-task-package-verifier.ps1`：`PASS task package verifier: 3 real packages, 15 assertion cases, 7 facts readable at runtime`
- `scripts\run-code-checks.ps1 -SkipBuild`：`discovered 25 gates under scripts\`（不变）、50 个 `PASS`、0 个 FAIL、`Golden L0 passed: 19 cases`（12→19）、`All code checks passed.`
- 四项目 `PASS: ShareRibbon / WordAi / ExcelAi / PowerPointAi`、`Code projects built successfully.`；构建后 `ShareRibbon\TaskPackages\*.json` 与 `ShareRibbon\bin\Debug\TaskPackages\*.json` MD5 三份两两相等（§17.5 那条闸门能成立的前提）
- 入口可达性（改完 `AGENTS.md` 后单独复跑 `node scripts\check-entry-reachability.js`）：面板发送 56 个消息 type、router 注册 66 个、DOM id 103/103、注入调用 61/61、导航文档 11 份、本地引用 45→46 条全部解析到磁盘（本轮给根 `AGENTS.md` 补了「包语义的黄金样例」一行，并在合同行写明改包须重建）；`Oversized baseline files: 22`（未新增超限文件）
- 两次真实红灯：旧构建副本冒充当前合同（§17.5）、手工还原误伤用例期望值（§17.6 第 2 条）

---

## 18. E6-05b 未达成的包真的回到既有 repair 路由

对应计划行 **E6-05** 的后半句「失败进入既有 repair loop 而不是直接宣称成功」；§16.6 记录的是它当时未达成，本节是接通它的一轮。真机按用户指示继续暂缓。

### 18.1 缺的不是「一条路由」，而是「不让它长成第二套」

E6-05 已经能在 run 末尾算出未达成并拦下假成功，缺的是拿着这份未达成去修。最省事的写法是在循环结束后再套一层「重规划 + 重跑」，那正是仓库明令禁止的形状：第二套 Loop、第二套终态、第二套 replan 幂等口径。所以这轮的验收不是「会不会修」，而是**修这条动作是否仍然由既有的那一个循环、既有的那一份预算、既有的那一条终态出口持有**。

### 18.2 闸门放在循环里面，而不是循环外面

`ShareRibbon/Agent/LoopEngine.vb:99-134`：`While` 的条件由 `stepIndex < session.Plan.Steps.Count` 扩成 `(stepIndex < session.Plan.Steps.Count OrElse Not runEndGated)`（`:103`）。计划跑完时 `runEndGated` 仍为 `False`，于是循环**再进一圈**，那一圈因为 `stepIndex >= Steps.Count` 命中 `:110` 分支，只走合同裁决：

- 四个前提（`:117-118`）：本轮真的执行过（`CurrentIteration > 0`）、**没有任何未落地也未跳过的步骤**（`IncompletePlanSteps(session).Count = 0`，失败步骤归 §14 的失败修复管，同一问题不得两个循环都来修）、包声明 `onFail=repair`（`PackageWantsRunEndRepair(session)`）、还有 replan 预算（`replanAttempts < MaxReplanAttempts`，常量 `:24` 仍是 2，**与失败修复共用同一份预算**，不是第二个计数器）。
- 接受一份新计划才做的四件事（`:126-133`）：`replanAttempts += 1`、`session.Plan = repairPlan`、`session.Writes.BeginNextPlan()`（E1-06 幂等：上一版已落盘的写入不会在新版里重跑）、`stepIndex = 0` + `runEndGated = False` 重新武装闸门，然后 `RaisePlanGenerated` 把第二版计划交给面板。
- 任何一条不成立就 `Exit While`（`:124`），落回循环后面原来的收尾顺序：`IncompletePlanSteps`（`:497`，闸门与终态判定**共用这一个口径**，内联 LINQ 已删除）→ `ValidateExecutionOutcome` → 唯一的失败出口，或唯一的成功出口 `AgentResult.SuccessResult`（`:519`，全文件计数 1）。

配套五个 helper 全在 `LoopEngine.Coverage.vb`：`MatchedTaskPackage`（`:214`，自由文本/无包直接 `Nothing`）、`UnmetTaskPackageReasons`（`:220`）、`PackageWantsRunEndRepair`（`:227`，策略**从包数据读**，不写死）、`RequestRunEndRepairPlanAsync`（`:238`）、`IncompletePlanSteps`（`:265`）。一个 VB 语言约束值得记：**Async 函数不能有 `ByRef` 参数**，所以「消耗一次预算」这件事不能在 helper 里改计数器——`RequestRunEndRepairPlanAsync` 只返回 `ExecutionPlan` 或 `Nothing`，`replanAttempts` 的加减留在调用方，预算的持有者因此仍然是循环本身。

### 18.3 观察结果长什么样：`RUN_END_CONTRACT_UNMET`

`RequestRunEndRepairPlanAsync` 把未达成的断言拼成一条观察文本，前缀是 ASCII 标记 `RUN_END_CONTRACT_UNMET {包id}`，正文带「本轮已落地写入 N 处，请只规划还缺的那部分，不要重复已经写入的内容」。前缀不是装饰：观察正文是中文，探针与日志检索无法靠中文判断走的是哪条路由，这条和 `PARTIAL_APPLY`、`CANCELLED` 是同一类机器可查合同。之后**整段交给既有的 `ReflectAndReplanAsync`**（`LoopEngine.Planning.vb`），因此反思与重规划自动继承 E1-06 那条「已落盘写入作为系统级约束下发」的机制；只接受 `strategy = replan`，`retry`/`skip` 在 run 末没有「当前失败步骤」可以落地，一律如实记 Warn 并按未达成报告；新计划还要过 `ValidateReplannedCoverage`，覆盖度丢了就不接受。`onFail=rollback` 依旧没有还原路由，`ValidateTaskPackageOutcome`（`:206-209`）保留那条明说 `no run-end restore route exists` 的 Warn。

### 18.4 我自己的一个错事实，被我自己写的前置检查驳回

§16.6 与计划第 25 条都写过「三个包的 `onFail` 都是 `repair`」。**盘上不是**：`excel.clean-and-fix` = `repair`、`word.review-polish` = `repair`、`excel.insight` = **`report`**。发现方式不是重读代码，而是我给新 smoke 写的前置检查按真实包数据断言每个包的声明策略，它直接抛出 `excel.insight no longer declares onFail=repair`——门禁把我记错的事实拦在了运行时用例之前。用例 5 于是改成反向证明「这条路由读的是包而不是硬编码」：给一个 `report` 包的违例，断言**反思次数为 0** 且终态如实报未达成。两处文档与项目记忆均已按实测更正。这是同一类错误的第三次（E6-01 的「9 处」实为 10 处、§14.6 的「加载器什么都拒」），结论没变：**「我记得」不是证据，量一遍才是**。

### 18.5 运行时门禁：新门禁 `scripts/smoke-run-end-repair.ps1`（259 行）

反射加载刚构建的 `ShareRibbon.dll`，用真实 `LoopEngine` + 脚本化模型 + 脚本化执行器跑 6 条合同用例，全程不落 Office/COM：

1. 未达成的 `excel.clean-and-fix` → 恰好 1 次反思、拿到新计划后成功收尾，`PlanGeneration=2`、`WrittenCount=1`，观察里同时出现 `writtenCount` 与 `changedCount`，提示词里出现 `RUN_END_CONTRACT_UNMET`；
2. 反复未达成 → 停在共享预算 `MaxReplanAttempts(2)`，终态仍点名包与事实，不装成功；
3. 自由文本（`EntryActionId=""`）→ **0 次反思**，闸门不绑包；
4. 合同已满足的包 → 0 次反思、`PlanGeneration=1`，闸门不空转；
5. `onFail=report` 的包违例 → **0 次反思**、失败消息点名 `excel.insight`（策略真的来自盘上数据）；
6. 修复轮拿不到可用计划 → 1 次反思、`PlanGeneration=1`、如实报未达成。

写这条 smoke 踩到的两个坑值得留档，因为它们都是「测试自己先错」：(a) 我的 react 分支按 `"toolId"` 派发，而 `BuildReactPrompt` 发的是 `{"action":{"tool":…,"params":…}}` 且不带工具清单，于是 react 请求被步骤修复分支吃掉、`Reflections=0`；E1-06 那份 smoke 用的是**无条件兜底分支**，改成兜底后才对得上。这是提示词形状的合同，不是测试技巧。(b) `ToolRegistry` 会校验必填参数，`DataAnalysis` 少 `sourceRange/type` 时返回 `[ARGUMENT_ERROR] recoverable=True`，用例里的工具参数必须由 `Get-ToolParams` 真实给出，未登记的工具直接 throw——否则执行器返回的观察是我编的，台账也就不再是真实收录口径。

### 18.6 静态侧：15 条新规则、15 条新探针、1 次盘上真实变异

`check-task-package-contract.ps1` 的入口身份语料从 12 个文件扩到 13 个（**加入 `LoopEngine.vb`**，闸门本体第一次进入静态视野），新增规则与探针一一对应：闸门可达（循环条件含 `OrElse Not runEndGated`）、失败步骤在场时不触发、策略读自包、预算共用 `replanAttempts < MaxReplanAttempts`、接受修复必须推进计划版本（`session.Writes.BeginNextPlan()`）、闸门必须重新武装、拿不到计划必须退回未达成报告、**终态出口计数必须等于 1**；Coverage 侧钉住「复用 `ReflectAndReplanAsync`」「本文件不得出现 `SendAIRequest(`」「只接受 `replan`」「必须过 `ValidateReplannedCoverage`」「自由文本不绑包」「rollback 未接这件事必须仍被如实记录」。入口身份探针 14→29 条，每条仍要求被**它自己那条规则**抓住。

反空转不止靠内存探针：把 `LoopEngine.vb:118` 的预算判断在**盘上**改成 `False Then`，四项目仍 `PASS`（说明这条路由不能靠编译器发现），跑新 smoke 当场红：`Smoke failed: unmet package must ask for repair exactly once, got 0`；随后按 MD5 还原（前后同为 `3d089d96f944ef15705b67a3a2bee8c4`）、重建四项目、smoke 与合同门禁复绿。这条红灯证明路由是**真在被执行**的，而不是只被 grep 读到。

### 18.7 顺手逮到的文档完整性缺陷

复核 §17.8 时发现三条路径被我自己早先用 python 写文档时打进了控制字符：`scripts`+CR+`un-golden-l0.ps1`、`scripts`+CR+`un-code-checks.ps1`（渲染成 `scriptsun-golden-l0.ps1`），以及 `ShareRibbon`+0x08+`in\Debug`。已按字节修复，并对 `openspec/**`、`docs/**`、三份根文档做了一次控制字符扫描，现为 0 处。**规则 J 没有抓到它**——含控制字符的引用不匹配任何路径模式，于是「断链」检查直接看不见这一条；这是防漂移门禁的一个已知盲区，记在这里而不是宣称它已覆盖。

### 18.8 真机验收清单（按用户指示暂缓，需 Excel/Word 真机 + 重启宿主刷新 www 副本）

1. Excel：走「数据分析」入口并制造「执行了但数据本来已经干净」的表，预期出现状态行「成功标准未达成，正在按任务包合同重新规划...」，日志含 `RUN_END_CONTRACT_UNMET excel.clean-and-fix`。
2. 同一次 run：面板应出现第二版计划卡，且第一版已落地的写入**不在**新版里重跑（对照 `agent_run_change` 该 run 的行数与文档实际内容）。
3. 连续不满足：最终答复点名包与未达成事实，绝不出现「任务完成」；确认预算耗尽后不再重规划（日志里 run-end repair 请求 ≤2 次）。
4. Word：「审阅润色」入口 + 一段本来无需改动的选区，确认要么一次修复后成功、要么如实未达成，且文档里没有重复插入。
5. 面板自由文本做同一件事：日志既不出现包 id 也不出现 `RUN_END_CONTRACT_UNMET`（证明不猜包在真机上成立）。
6. `excel.insight` 只读洞察：0 次反思、直接成功（`report` 策略不该触发修复）。
7. 观感判断：一轮任务给出两版计划是否让人困惑（`RaisePlanGenerated` 会重画计划卡）；结论决定要不要把第二版计划折叠成一条增量说明。

### 18.9 声明的不作为与代价

1. 没有新建第二套 Loop/Planner/verifier/Safety/终态：成功出口计数静态钉死为 1，预算与幂等都归既有循环持有。
2. `onFail=rollback` **仍未接通**——它需要 run 末把本轮改动交回宿主还原，那条能力不属于这一行；本轮只保证它继续被如实报告（并有探针防止「悄悄当成已处理」）。
3. 没有为自由文本绑包：面板输入同一件事仍不受合同约束，也不享受修复（用例 3 证明），等 E6-07/D9。
4. 没有改步骤失败修复的任何行为，两条路径只是共用同一份预算：一次 run 若在步内已耗尽 2 次 replan，run 末**不再**修复，直接按未达成报告。
5. 代价：修复轮会多花一轮模型调用与一次完整重跑，`MaxIterations` 也照旧消耗；用户会看到「一次任务两版计划」，观感待 §18.8 第 7 条判定。
6. 代价：修复计划的合法性由 `ValidateReplannedCoverage`（spec 的产出覆盖）判断，包断言本身要到下一圈 run 末才再被核对——这是刻意的单点顺序，不是遗漏，但意味着「修了仍不满足」的路径必须靠用例 2 而非静态规则来保证。
7. 真机全部未跑（§18.8 共 7 条），E6-08 六包真机完成率仍依赖它。

### 18.10 本轮实测数字（全部为本次重新运行所得）

- 规模：`LoopEngine.vb` 619→646 行、`LoopEngine.Coverage.vb` 325→382 行、新门禁 `scripts/smoke-run-end-repair.ps1` 259 行、`check-task-package-contract.ps1` 1033→1144 行（身份语料 12→13 个文件、探针 14→29 条）
- `scripts\smoke-run-end-repair.ps1`：`PASS run-end repair: 6 contract cases over the real LoopEngine`
- `scripts\check-task-package-contract.ps1`：`PASS task-package contract: 3 packages, 6 assertions, 15 schema keys, 7 facts, 14 probes turned red on 2 clean baselines, golden references=7 (7 outcome cases, 11 golden probes turned red), ribbon utterances=10 at 10 button(s), 29 entry-identity probes turned red at 10 panel site(s)`
- `scripts\run-code-checks.ps1 -SkipBuild`：`discovered 26 gates under scripts\`（25→26，新 smoke 被目录自动发现、未登记进 `$notGates`）、52 个 `PASS`（50→52：+1 条步骤行、+1 条 smoke 自己的 PASS）、0 个 FAIL、`Golden L0 passed: 19 cases`、`All code checks passed.`
- 四项目 `PASS: ShareRibbon / WordAi / ExcelAi / PowerPointAi`、`Code projects built successfully.`（变异还原后再复跑一次，仍 4 个 PASS）
- 入口可达性（改完 `AGENTS.md` 两行后复跑）：面板发送 56 个消息 type、router 注册 66 个、DOM id 103/103、注入调用 61/61、导航文档 11 份、本地引用 46→48 条全部解析到磁盘（本轮给合同行与执行者行各补了 run 末路由的路径）、Ribbon 可见入口 Word 14/18、Excel 13/18、PowerPoint 11/18（本轮未动入口行集）；`Oversized baseline files: 22`（两个 Loop 文件都在 650 行以下，未新增超限）
- 一次真实红灯：盘上把 run 末预算改成恒假 → smoke 红 `ask for repair exactly once, got 0`（§18.6）
- 一处文档缺陷：3 条路径含控制字符 → 字节级修复 + 全仓扫描归零（§18.7）

## 19. E3-05 隐私同意：把「有个字段」变成「真有人拦」

对应计划行 **E3-05 隐私同意 UI**（首次云调用前记录同意；可查看、可撤回；文案说明数据边界）。真机按用户指示继续暂缓，只写清单。

### 19.1 这一行缺的不是 UI，是「 enforcement 」

开工前先在 `HEAD` 上量了一遍：`PrivacyPolicyAccepted` 在仓库里只出现 6 次——3 次在 `SecureConfigManager.vb:231-236`（存储本体），3 次在 `OnboardingService.vb:40/:62/:73`（向导第 1 步的读写与步骤判定）。**没有任何一条发送链路读过它**。同目录的 `AllowSendDocumentContent` / `AllowSaveConversationHistory` / `AllowUsageStatistics` 三个兄弟字段更彻底：除了 getter/setter 全仓 0 引用。这就是仓库反复惩罚的「字段齐全、无人读取」形状——所以这一行的真正交付是「拦截 + 可撤回」，界面只是它的可见面。

第二条测量决定了设计：`git grep SendStreamRequestAsync` 只有 1 个调用方（`BaseChatControl.vb:2999`，在 `Send` 内），所以流式出口拦在 `Send` 就够；但 `AiGateway.SendChatAsync` 有 5 个调用方（`AgentKernelService.vb:503` 已在闸门内，`AutocompleteService.vb:211`、`IntentRecognitionService.vb:569/:1274`、`LlmMemoryExtractor.vb:58` 是绕开面板漏斗的长尾）。结论：**同意判定必须放在出口，而不是在每个入口各补一次**。

### 19.2 出口清单，而不是「我记得都拦了」

`HEAD` 上没有清单，只有印象。于是这一轮先把「能把字节发上网的地方」全量扫出来（`HttpRequestMessage(HttpMethod.Post` + `.PostAsync(`），实测 **16 个出口 / 12 个文件**，逐个决定它的口径并登记进 `scripts/smoke-privacy-consent.ps1` 的 `$outboundInventory`：

- **9 个文件自己问**（`coverage=guard`）：`AiGateway.vb:61`、`BaseChatControl.vb:2253`（`SendAndGetResponseAsync`，排版/审阅/AgentKernel 的非流式发送器）、`UtilsService.vb:265`、`LLMUtil.vb:96`（ExcelDna 自定义函数与批量数据生成）、`AutocompleteService.vb:145`、`OfficeCompletionService.vb:121/:171`（FIM 与 chat 两条补全形态）、`EmbeddingService.vb:94`、`DocumentTranslateService.vb:420`（Word/PPT 继承同一基类）、`ExcelDocumentTranslateService.vb:340`（Excel 那份是复制品，口径必须一致）。
- **两条面板漏斗拦在前面**（`coverage=funnel`）：`HttpStreamService.vb` 的 3 个 POST 只能从 `Send`（`:2920`）或 `HandleStartAgent`（`:1812`）到达，`IsBlockedByPrivacy` 在两者之后、构造请求体之前。豁免的前提是「它真的只有一个公开发送入口」，这条被写成断言：`Public Async Function Send\w*Async` 的计数必须等于 1，出现第二个入口就红（有探针）。
- **1 个固定探针**（`coverage=probe`）：`ModelConnectivityTester.vb` 的请求体是字面量 `{"content", "Hi"}`、`max_tokens=5`，不带任何用户内容，而且它是用户在**同意之前**用来验证接口的工具——拦它等于让人没法配置。**它是规则 9 上线第一次运行当场逮到的**：当时清单里没有它，门禁直接红 `an outbound request site is not registered in the consent inventory: ModelConnectivityTester.vb`。这条豁免同样带断言：那句话要是被换成文档正文，豁免立刻失效。
- **2 个 MCP 出口**（`coverage=mcp`）：`StreamJsonRpcMCPClient.vb` 打的是用户自己添加的工具服务器，不是模型接口，归 W2 的 MCP 授权面管。写进清单而不是留白，是为了让「以后有人把文档内容塞进 MCP 参数」必须回来改这一行、重新拍板。

清单的防漂移形状：**语料里扫到的出口数与清单对不上就红**（新出口没登记 → 红；登记了却不再发东西 → 红；`HttpStreamService` 从 3 个 POST 变 4 个 → 红），探针 3 条分别抓这三种漂移，外加「把 guard 字样抹掉」的一条。这条规则把「我们拦住了云调用」从一句承诺变成一条可重跑的计数。

### 19.3 拒绝形状沿用各出口自己的失败约定，不发明第三种终态

同一次「没有同意」在不同出口上的表现必须各不相同，但理由只有一条：

- 面板两条漏斗 → `PrivacyConsentGate.Evaluate` 产出卡片，正文是三条数据边界 + 「消息没有发送/任务没有开始执行，文档不会被修改」；
- `SendAndGetResponseAsync` → `Return ""`，正是它 `Catch` 里既有的失败返回值；
- `UtilsService` / `LLMUtil` → 空串与 `错误: ` + 同一条理由，各自既有约定；
- 补全（`Autocomplete`/`OfficeCompletion`）→ 空列表，即「这一轮没有建议」；embedding → `Nothing`，六个调用方本来就会降级；
- 翻译 → **抛异常**，因为它的调用方一律按异常处理并保留原文（`DocumentTranslateService.vb:257-260/:286-294`）。这里特意不返回空串：空串会变成 `JObject.Parse` 的语法错误，把一次隐私拦截伪装成「模型返回坏了」——这正是「不要用垃圾代码掩盖问题」在这一行的具体样子。

为此新增 `PrivacyConsentGate.RefusalReason`（唯一的机器理由串）与 `RefuseQuietly(caller)`（判定 + 统一日志），后台链路一行调用即可，`AppLogger.Debug` 落的是同一条 `privacy consent not accepted; request was not sent.`。

### 19.4 一张卡片，不是一个新家族

隐私卡片复用 E3-03 的 `ModelConfigPrompt` + `showModelConfigRequired` + `#model-config-cta`，只多了一颗判别字段 `action`（`open_config` / `accept_privacy`），`chat-manager.js:428` 在按钮回调里分支。**零新增面板消息 type**：写回用的是既有的 `onboardingSetPrivacy`（它的 `accepted` 字段本来就允许 `true|false`）。`ModelConfigGate.BuildShowScript` 因此从实例方法改为 `Shared`——阻断卡片只有一个持有者，画它的句子也该只有一个。代价当场显现：`scripts/smoke-model-config-gate.ps1:69/:89` 是用 `$gate.BuildShowScript(...)` 调的，改共享后 PowerShell 报 `MethodNotFound`，两处改成 `[ShareRibbon.ModelConfigGate]::BuildShowScript(...)` 才复绿。这是「顺手改签名要清点调用方」的又一次提醒，只不过这次调用方是门禁脚本。

### 19.5 可撤回，且撤回看得见

面板设置弹窗里加了一行「数据边界同意」（`chat-template-refactored.html:350-357`，复用 `.onboarding-privacy` 样式，0 行新 CSS）。三件事按 E3-01/E4-05 的既有口径收口：勾选框初始 `disabled`，只有宿主推来状态才可点（`syncPrivacyConsentRow`，`onboarding.js:169-186`）；打开设置时若还没有状态就补问一次 `getOnboardingState`（`settings-manager.js:8-16`），不显示未经证实的默认值；取消勾选即撤回，宿主写完立刻回推，`InitialStep` 回到 `privacy`，下一条消息/任务重新拦在发送之前。**不做**「同意后自动重放刚才那条请求」：卡片写的是「这一轮没有发出去」，偷偷重放会让承诺和实际行为不一致，也让 Agent 回合的幂等假设多出一种例外。

### 19.6 文案只有一个持有者，且规则是数据驱动的

三条边界说明写在 `PrivacyConsentGate.BoundaryItems`，随 `renderOnboarding` 的 `privacyBoundary` 字段下发；向导与设置行都只渲染 `state.privacyBoundary`，`chat-template-refactored.html` 里不再出现一句边界文案。门禁不是硬编码这三句话，而是**从 `PrivacyConsentGate.vb` 正则读出来再全仓计数**，要求每条恰好出现 1 次——改文案不需要改门禁，抄一份承诺才会红。这条规则当场抓到过一个真问题：`OnboardingService.BuildState` 里那句 `JArray.FromObject(PrivacyConsentGate.BoundaryItems)` 一度是我打算让前端自持的副本，探针 `wizard-drops-boundary-push` 把它钉成了「宿主必须下发」。

### 19.7 反空转：9 组静态规则、22 条探针红、3 次真实网络拒绝

`scripts/smoke-privacy-consent.ps1`（527 行，被 `scripts\*.ps1` 自动发现为第 27 个门禁）静态侧扫 4 个模块的 `.vb/.js/.html` 共 330 个文件外加 `.vbproj`（9 组规则、21 条内存探针 + 1 个合成第二写入者 = 22 条必须红，且每条要由**它自己那条规则**抓住）。运行时侧反射加载刚构建的 DLL 真跑：`Evaluate($true,…)` 必须给 `Nothing`；chat/agent 两版卡片文案必须不同且各自命中「消息/任务」标记；`BuildShowScript` 载荷往返后 `action` 仍是 `accept_privacy`；配置卡片仍必须是 `open_config`；向导同意与撤回**真落到 `secure.config`**（`PrivacyPolicyAccepted=true|false`，字节级备份还原）；`AiGateway.SendChatAsync` 打 `127.0.0.1:1` 必须在 31ms 内以同意理由失败（离线），`UtilsService` 7ms、`LLMUtil` 4ms 同理；`Apply($true)` 后同一个调用不得再被同意拦；`renderOnboarding` 推出去的 payload 必须带 `privacyAccepted` 与逐字相等的三条边界。

写这条门禁自己踩的四个坑，全是「测试先错」：`return ,$problems` 让前置检查把空数组读成 1 个缺陷；探针改 `Fail(...)` 文案没让规则红，因为同一句话在上一行的 `Debug.WriteLine` 里还在——规则 5 于是改成**位置合同**（守卫必须在 `_httpClient.SendAsync` 之前），探针改为把整段守卫搬到文件末尾；`BuildShowScript(Nothing)` 返回的是 `String.Empty` 而不是 `Nothing`，`$null -ne` 这条断言写反了；`$outboundInventory` 数组尾逗号在 PowerShell 5.1 里直接 `MissingExpressionAfterToken`。

### 19.8 明确不作为与代价

1. **`secure.config` 默认值是 `false`**，所以老用户升级后第一次发消息会被拦一次（点「同意并继续」即恢复）。这是同意的定义本身，不是缺陷，但确实是一次可见的行为变化。
2. 后台链路（补全、embedding、记忆沉淀、ExcelDna 函数、批量生成）拒绝时**没有卡片**——它们没有面板可画，表现为「这次没有结果」，日志里留同一条理由。用户可见的解释只有面板漏斗那两条有。
3. 三个 `Allow*` 装饰字段这一轮**没动**：它们既不读也不写业务，真要收口得先决定「会话历史开关」「用量统计」是不是产品要的能力，属于另立一行。
4. 没有为「同意后自动重放被拦下的那条请求」写代码；没有新建第二张卡片、第二个消息 type、第二套文案；没有把 `model-config-cta` / `showModelConfigRequired` 改名（纯 CSS+JS+VB 三方改名、零行为变化，风险大于收益）。
5. 翻译与 ExcelDna 那几条链路仍各自持有 HttpClient，本轮只在它们前面各补一把同一口径的锁；把它们收进 `AiGateway` 是 `ShareRibbon/AGENTS.md:46`「AI 请求统一经过共享请求层」的应有改造，属于独立一行（记为新任务，不在本轮偷偷做）。
6. 真机全部未跑（§19.9），面板设置行与卡片文案的观感未经人眼确认；`www` 副本要靠重启宿主刷新。

### 19.9 真机验收清单（按用户指示暂缓）

1. 全新配置（`secure.config` 无 `PrivacyPolicyAccepted`）下打开面板：向导第 1 步是隐私确认，三条边界文案由宿主下发。
2. 不点同意，直接在输入框发消息：出现「先确认数据边界」卡片，正文含「消息没有发送，文档不会被修改」；点「同意并继续」后卡片消失，`secure.config` 出现 `PrivacyPolicyAccepted=true`，**这条消息仍需再发一次**。
3. 不点同意直接点 Ribbon/命令面板的 Agent 任务：卡片文案变成「任务没有开始执行」，文档零改动。
4. 面板设置 →「数据边界同意」行：打开时即为已勾选、可交互；取消勾选后 `secure.config` 回到 `false`，此时再发消息应重新被拦。
5. Word/Excel/PPT 三端各跑一次「翻译」：未同意时进度/结果应显示「还没确认数据边界，翻译请求没有发送」，**原文不得被清空**（重点看这一条，它验证 §19.3 的异常口径没有变成数据丢失）。
6. 配置弹窗「测试连接」：未同意时仍应可用（它只发一个 `Hi`），验证豁免没有被误伤。
7. 打开 Word 输入 10 秒不确认同意：ghost text 不应出现；日志里能看到 `AutocompleteService`/`OfficeCompletionService` 的那条拒绝理由。
8. 同意后同样操作：补全恢复，聊天/Agent 正常，`agent_run` 与日志里没有把文档内容写进拒绝理由的痕迹（拒绝日志只有一行固定文本）。

### 19.10 本轮实测数字（全部为本次重新运行所得）

- 规模：新文件 `ShareRibbon/Controls/Services/PrivacyConsentGate.vb` 62 行、新门禁 `scripts/smoke-privacy-consent.ps1` 527 行；`ModelConfigGate.vb` 98 行（`+12` 行全是这一行改的）、`OnboardingService.vb` 110 行、`AiGateway.vb` 452 行（`+7`）；`BaseChatControl.vb` 新增 23 行（`:116-125` 助手 + `:1812`、`:2253`、`:2920` 三处调用）
- 出口清单：**16 个出口 / 12 个文件**，其中 9 个文件自带守卫、3 个文件按豁免登记（`funnel` 1 个文件 3 个 POST、`probe` 1 个、`mcp` 1 个文件 2 个）
- 静态规则 9 组、语料 330 个 `.vb/.js/.html` + `.vbproj`、探针 22 条全红；运行时 3 次真实离线拒绝：`RefusedMs=31`、`UtilsRefusedMs=7`、`LlmRefusedMs=4`，`ConsentMs=2034`（同意后同一请求走到死 socket）
- `scripts\smoke-privacy-consent.ps1`：`OK: consent gates every registered cloud exit, is revocable, and the boundary copy has one owner.`
- `scripts\smoke-model-config-gate.ps1`：`OK: model config gate blocks both entries and emits a parseable CTA.`（修掉两处实例调用后复绿）
- `scripts\run-code-checks.ps1`（含构建）：`discovered 27 gates under scripts\`（26→27）、55 个 `PASS`、0 个 FAIL、`Golden L0 passed: 19 cases`、`All code checks passed.`；四项目 `PASS: WordAi / ExcelAi / PowerPointAi`（`ShareRibbon` 随依赖链重编）+ `Code projects built successfully.`
- 入口可达性：面板发送 56 个消息 type（**未增加**，与上一轮同数）、router 注册 66 个、DOM id 103→105/105 全部解析到节点（新增设置行两个 id）、注入调用 61/61、导航引用 48/48、Ribbon 可见入口 Word 14/18、Excel 13/18、PowerPoint 11/18；`Oversized baseline files: 22`（未新增超限）
- 两次当场逮到的真问题：规则 9 第一次运行抓到未登记的 `ModelConnectivityTester.vb`；`BuildShowScript` 改 `Shared` 抓到旧门禁的实例调用
