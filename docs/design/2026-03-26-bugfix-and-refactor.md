# Office AI Agent — Bug 修复 + Send 重构 设计文档

> 版本：v1.0  日期：2026-03-26
> 范围：用户反馈的全部 Bug + 聊天 Send 提示词组装重构
> 优先级：P0 崩溃/功能完全不可用 → P1 影响主流程 → P2 体验问题 → P3 代码质量

---

## 目录

1. [Send 提示词组装重构（P1）](#1-send-提示词组装重构)
2. [PPT 附件无法识别解析（P0）](#2-ppt-附件无法识别解析)
3. [AI 聊天突然弹出"请求失败"（P0）](#3-ai-聊天突然弹出请求失败)
4. [PPT 翻译结果 JS 换行符丢失（P1）](#4-ppt-翻译结果-js-换行符丢失)
5. [Excel 批量数据处理无实现（P1）](#5-excel-批量数据处理无实现)
6. [API 配置项无法删除（P1）](#6-api-配置项无法删除)
7. [网页爬取跳转异常（P1）](#7-网页爬取跳转异常)
8. [Excel Office 2021 功能按钮灰色（P2）](#8-excel-office-2021-功能按钮灰色)
9. [AI 结果无法指定位置输出（P2）](#9-ai-结果无法指定位置输出)
10. [WPS 安装后 COM 加载项不显示（P2）](#10-wps-安装后-com-加载项不显示)
11. [MSI 被杀毒软件误报（P3）](#11-msi-被杀毒软件误报)
12. [待确认 Bug](#12-待确认-bug)

---

## 1. Send 提示词组装重构

### 问题描述

用户反馈：多功能组合时（记忆 + Skills + 意图 + 选区）提示词结构混乱，AI 难以区分各部分来源和权重。

### 根因分析

**文件：** `ShareRibbon/Controls/Services/ChatContextBuilder.vb`
**文件：** `ShareRibbon/Controls/BaseChatControl.vb`

**问题 1 — 结构化缺失**
`ChatContextBuilder.BuildMessages` 将所有层（baseSystem、场景提示、Skills、记忆）全部用 `vbCrLf & vbCrLf` 简单拼接到 `result(0).content`，无任何区分标记。AI 无法判断某段内容的来源（是系统指令？是记忆？还是 Skill 详情？）。

```vb
' 当前代码（有问题）
result(0).content = result(0).content & vbCrLf & vbCrLf & layer1   ' Skills拼进去
result(0).content = result(0).content & vbCrLf & vbCrLf & memoryBlock  ' 记忆再拼进去
```

**问题 2 — 记忆重复检索**
`Send()` 在调用 `CreateRequestBody()` 之后，又额外检索一次记忆只为获取 `ragCount`（行 2633-2635），浪费一次 SQLite 查询。

```vb
' 当前代码（有问题）—— BaseChatControl.vb ~2633
Dim mems = MemoryService.GetRelevantMemories(question, MemoryConfig.RagTopN, Nothing, Nothing, GetOfficeAppType())
ragCount = If(mems IsNot Nothing, mems.Count, 0)
```

**问题 3 — 硬编码 Fallback 提示词**
当 `PromptManager` 无配置时，`Send()` 注入了 5 条硬编码中文规则（行 2619-2625），这些规则与用户配置的系统提示词可能冲突，且难以维护。

### 设计方案

#### 1.1 ChatContextBuilder 结构化重构

将 system 消息拆为以 `---` 分隔的具名节，每节以 `### 标题` 开头：

```
### 角色与基础指令
{baseSystemPrompt}

---

### 场景能力
{场景数据库提示词，仅当存在时}

---

### 可用技能
{Skills目录}

#### 推荐技能（基于当前查询）
{topSkill详情，仅当 matchScore >= 10 时}
> 当前推荐: {name} | 标签: {tags} | 匹配关键词: {keywords}

---

### 用户上下文
#### 用户画像
{userProfile，仅当存在时}

#### 相关记忆
- {memory1}
- {memory2}

#### 近期会话
- {sessionTitle}: {snippet}
```

#### 1.2 ragCountOut ByRef 参数

`BuildMessages` 签名增加 `Optional ByRef ragCountOut As Integer = 0`，在记忆检索时同步赋值，省去 `Send()` 的重复查询。

```vb
' 新签名
Public Shared Function BuildMessages(..., Optional ByRef ragCountOut As Integer = 0) As List(Of HistoryMessage)

' CreateRequestBody 也透传
Private Function CreateRequestBody(uuid, question, systemPrompt, addHistory, Optional ByRef ragCountOut As Integer = 0) As String
```

#### 1.3 Send() Fallback 简化

```vb
' 原来（5 条硬编码规则）→ 改为
If String.IsNullOrWhiteSpace(systemPrompt) Then
    systemPrompt = If(Not String.IsNullOrWhiteSpace(ConfigSettings.propmtContent),
                      ConfigSettings.propmtContent,
                      "你是一个 Office AI 助手，请根据用户需求提供简洁、准确的回答。")
End If
```

### 改动文件

| 文件 | 改动点 |
|------|--------|
| `ShareRibbon/Controls/Services/ChatContextBuilder.vb` | `BuildMessages` 加结构化节标题 + `ragCountOut` 参数 |
| `ShareRibbon/Controls/BaseChatControl.vb` | `CreateRequestBody` 加 `ragCountOut`；`Send()` 移除重复记忆查询；简化 Fallback |

---

## 2. PPT 附件无法识别解析

**优先级：P0**

### 根因分析

**文件：** `PowerPointAi/ChatControl.vb` 第 434 行

```vb
' 当前代码（有问题）
Dim pptApp As New Microsoft.Office.Interop.PowerPoint.Application()
pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
```

在 VSTO 插件内部（已运行于 PPT 进程中）再 `New Application()` 会尝试创建**第二个 PPT COM 服务器进程**。在 Office 2021 / Microsoft 365 上，此行为常因 COM 单例限制或权限问题抛出异常，导致附件解析完全失败。此外，原代码在 `Finally` 中调用 `pptApp.Quit()` 会关闭这个新建实例，但如果创建本身就失败，则进入 Catch 返回错误信息。

### 设计方案

复用当前进程的 PPT Application 实例 (`Globals.ThisAddIn.Application`)，配合 `WithWindow:=msoFalse` 静默打开文件，文件关闭后不 Quit（不能关闭用户正在用的进程）。

```vb
' 修复后
Dim pptApp = Globals.ThisAddIn.Application   ' 复用现有进程

Dim presentation As Microsoft.Office.Interop.PowerPoint.Presentation = Nothing
Try
    presentation = pptApp.Presentations.Open(filePath,
        ReadOnly:=Microsoft.Office.Core.MsoTriState.msoTrue,
        Untitled:=Microsoft.Office.Core.MsoTriState.msoFalse,
        WithWindow:=Microsoft.Office.Core.MsoTriState.msoFalse)
    ' ... 解析内容 ...
Finally
    If presentation IsNot Nothing Then
        presentation.Close()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(presentation)
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End If
    ' 不调用 pptApp.Quit() — 不能关闭用户的 PPT 进程
End Try
```

### 改动文件

| 文件 | 改动点 |
|------|--------|
| `PowerPointAi/ChatControl.vb` | `ParseFile` 第 434 行：移除 `New Application()`，改用 `Globals.ThisAddIn.Application`；移除 `pptApp.Quit()` 和对应的 `ReleaseComObject(pptApp)` |

---

## 3. AI 聊天突然弹出"请求失败"

**优先级：P0**

### 根因分析

两处 `MessageBox.Show` 在 UI 线程上弹出模态对话框，任何 API 超时、网络抖动都会打断用户操作：

**位置 1：** `ShareRibbon/Controls/BaseChatControl.vb` 第 2665 行
```vb
Catch ex As Exception
    MessageBox.Show("请求失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
```

**位置 2：** `ShareRibbon/Translate/BaseTranslateService.vb` 第 118 行
```vb
Catch ex As Exception
    MessageBox.Show($"请求失败: {ex.Message}")
    Return String.Empty
```

**位置 3：** `ShareRibbon/Controls/BaseDataCapturePane.vb` 第 158 行（页面加载失败也用 MessageBox）
```vb
MessageBox.Show("页面加载失败，请检查网络连接或重试", "警告", ...)
```

### 设计方案

所有网络/请求错误改用 `GlobalStatusStrip.ShowWarning`，同时写入 `Debug.WriteLine` 保留调试信息。仅对需要用户**必须确认才能继续**的操作（如删除、覆盖）保留 MessageBox。

```vb
' 修复后 — BaseChatControl.vb
Catch ex As Exception
    Debug.WriteLine("Send 请求失败: " & ex.Message & vbCrLf & ex.StackTrace)
    GlobalStatusStrip.ShowWarning("请求失败: " & ex.Message)

' 修复后 — BaseTranslateService.vb
Catch ex As Exception
    Debug.WriteLine($"翻译请求失败: {ex.Message}")
    Return String.Empty   ' 已有 ShowWarning 在调用方处理

' 修复后 — BaseDataCapturePane.vb
If Not args.IsSuccess Then
    Debug.WriteLine($"页面加载失败，WebStatus: {ChatBrowser.CoreWebView2.Source}")
    GlobalStatusStrip.ShowWarning("页面加载失败，请检查网络或 URL")
End If
```

### 改动文件

| 文件 | 改动点 |
|------|--------|
| `ShareRibbon/Controls/BaseChatControl.vb` | 第 2665 行 `MessageBox` → `GlobalStatusStrip.ShowWarning` |
| `ShareRibbon/Translate/BaseTranslateService.vb` | 第 118 行 `MessageBox` → `Debug.WriteLine` |
| `ShareRibbon/Controls/BaseDataCapturePane.vb` | 第 158 行 `MessageBox` → `GlobalStatusStrip.ShowWarning` |

---

## 4. PPT 翻译结果 JS 换行符丢失

**优先级：P1**

### 根因分析

**文件：** `PowerPointAi/Ribbon1.vb` 第 177 行

```vb
' 当前代码（有问题）
Dim escapedText = displayText.Replace("\", "\\").Replace("'", "\'") _
                             .Replace(vbCr, "\n").Replace(vbLf, "")
```

执行顺序：先把 CR 替换成字面量 `\n`，再把 LF 删除。
- CRLF（`\r\n`）：CR→`\n`，LF→`` → 结果 `\n` ✓
- **LF-only**（AI 输出内容常见）：LF→`` → **行尾全部丢失** ✗
- CR-only：CR→`\n` ✓

对比项目中正确的参考实现（`BaseChat.vb` 第 532-539 行的 `EscapeJavaScriptString`）：

```vb
' 正确顺序：先删 CR，再把 LF 转为 \n
.Replace(vbCr, "").Replace(vbLf, "\n")
```

另外当前代码没有转义 `</script>`，存在潜在注入风险。

### 设计方案

对齐 `EscapeJavaScriptString` 的转义顺序：

```vb
' 修复后
Dim escapedText = displayText _
    .Replace("\", "\\") _
    .Replace("'", "\'") _
    .Replace("</script>", "<\/script>") _
    .Replace(vbCr, "") _
    .Replace(vbLf, "\n")
```

### 改动文件

| 文件 | 改动点 |
|------|--------|
| `PowerPointAi/Ribbon1.vb` | 第 177 行：调整 CR/LF 替换顺序，补充 `</script>` 转义 |

---

## 5. Excel 批量数据处理无实现

**优先级：P1**

### 根因分析

**文件：** `ExcelAi/Ribbon1.vb` 第 210-222 行

```vb
Protected Overrides Sub BatchDataGenButton_Click(...)
    Dim batchDataForm As New BatchDataGenerationForm()
    If batchDataForm.ShowDialog() = DialogResult.OK Then
        Dim excelApp As Excel.Application = Globals.ThisAddIn.Application
        Dim activeWorksheet As Excel.Worksheet = excelApp.ActiveSheet
        ' 这里实现数据生成逻辑
        ' ...          ← 空实现！用户点击OK后什么都不发生
    End If
End Sub
```

`BatchDataGenerationForm` 已有字段定义 UI，但点击确定后没有任何数据发送给 AI 或写入单元格。

**文件：** `ExcelAi/Ribbon1.Designer.vb` 第 72 行
```vb
Me.BatchDataGenButton.Visible = False   ' 按钮默认隐藏
```

### 设计方案

两步：
① 实现 `BatchDataGenButton_Click` 的实际逻辑：从表单收集字段定义 → 构造 AI 提示词 → 通过 `ChatControl.SendChatMessage` 发送 → AI 返回 JSON → 写入单元格
② 将 `BatchDataGenButton.Visible` 改为 `True`（或在配置中控制可见性）

提示词模板：
```
请根据以下字段定义，生成 {rowCount} 行测试数据，以 JSON 数组格式输出：
字段：{字段名} ({说明})，写入列 {列号}
...
返回格式：[{"A": "值1", "B": "值2"}, ...]
```

AI 返回后，在 `ExcelDirectOperationService` 中新增 `WriteBatchData(data, startRow)` 方法执行写入。

### 改动文件

| 文件 | 改动点 |
|------|--------|
| `ExcelAi/Ribbon1.vb` | `BatchDataGenButton_Click` 实现完整逻辑 |
| `ExcelAi/Ribbon1.Designer.vb` | `BatchDataGenButton.Visible = True` |
| `ExcelAi/ChatControl.vb` 或新文件 | `BatchDataWriter.WriteBatchData(data, worksheet, startRow)` |

---

## 6. API 配置项无法删除

**优先级：P1**

### 根因分析

**文件：** `ShareRibbon/Config/ConfigApiForm.vb` 第 1405-1417 行

```vb
Private Sub CloudDeleteButton_Click(...)
    If currentCloudConfig Is Nothing Then Return       ' ← 静默返回，无提示
    If currentCloudConfig.isPreset Then
        MessageBox.Show("预置配置不可删除")
        Return
    End If
    ...
End Sub
```

问题有两点：
1. 如果用户点击删除时列表没有选中项（`currentCloudConfig Is Nothing`），按钮静默无反应，用户不知道为什么删除没有效果
2. 通过 `Grep` 未找到 `cloudDeleteButton.Enabled` 的设置代码，说明删除按钮在窗体初始化时**可能默认是禁用状态**，且没有在选中列表项时动态启用

### 设计方案

在 `CloudProviderListBox_SelectedIndexChanged` 中根据选中项的 `isPreset` 属性动态设置删除按钮状态：

```vb
Private Sub CloudProviderListBox_SelectedIndexChanged(...)
    Dim selected = TryCast(cloudProviderListBox.SelectedItem, ConfigItem)
    currentCloudConfig = selected
    ' 非预设配置才允许删除
    cloudDeleteButton.Enabled = (selected IsNot Nothing AndAlso Not selected.isPreset)
    ' ... 其余字段填充逻辑
End Sub
```

同时在 `CloudDeleteButton_Click` 中把 `Is Nothing` 静默返回改为提示：

```vb
If currentCloudConfig Is Nothing Then
    GlobalStatusStrip.ShowWarning("请先选择要删除的配置")
    Return
End If
```

**本地模型** (`LocalProviderListBox`) 同样需要相同修复。

### 改动文件

| 文件 | 改动点 |
|------|--------|
| `ShareRibbon/Config/ConfigApiForm.vb` | `CloudProviderListBox_SelectedIndexChanged` 动态设置 `cloudDeleteButton.Enabled`；同样处理 `LocalProviderListBox_SelectedIndexChanged` |

---

## 7. 网页爬取跳转异常

**优先级：P1**

### 根因分析

**文件：** `ShareRibbon/Controls/BaseDataCapturePane.vb`

**问题 1：新窗口重定向逻辑不完整**（第 167-174 行）

```vb
AddHandler ChatBrowser.CoreWebView2.NewWindowRequested,
    Sub(s, args)
        args.Handled = True
        ChatBrowser.CoreWebView2.Navigate(args.Uri)   ' ← 直接导航，没有检查 URI 有效性
    End Sub
```

`args.Uri` 在某些情况下可能为空字符串（如 `window.open()` 无 URL），此时 `Navigate("")` 会导航到空页面，用户丢失当前页。

**问题 2：页面加载失败弹 MessageBox**（第 158 行） — 已在 Bug 3 中处理。

**问题 3：没有处理重定向循环**
部分站点（SSO 登录、防爬重定向）会产生连续跳转，当前没有跳转次数限制或异常 URL 检测。

### 设计方案

```vb
' 修复后 — NewWindowRequested
AddHandler ChatBrowser.CoreWebView2.NewWindowRequested,
    Sub(s, args)
        args.Handled = True
        Dim targetUri = args.Uri
        If Not String.IsNullOrWhiteSpace(targetUri) AndAlso
           (targetUri.StartsWith("http://") OrElse targetUri.StartsWith("https://")) Then
            ChatBrowser.CoreWebView2.Navigate(targetUri)
            Debug.WriteLine($"新窗口重定向: {targetUri}")
        Else
            Debug.WriteLine($"忽略无效新窗口 URI: {targetUri}")
        End If
    End Sub
```

同时，将 `NavigationCompleted` 中失败时的 `MessageBox` 改为状态栏提示（见 Bug 3）。

### 改动文件

| 文件 | 改动点 |
|------|--------|
| `ShareRibbon/Controls/BaseDataCapturePane.vb` | `NewWindowRequested` 增加 URI 有效性检查；`NavigationCompleted` 失败改为状态栏提示 |

---

## 8. Excel Office 2021 功能按钮灰色

**优先级：P2**

### 根因分析

**文件：** `ExcelAi/Ribbon1.Designer.vb` 第 72-76 行

```vb
Me.BatchDataGenButton.Visible = False
Me.WebCaptureButton.Visible = False
Me.ProofreadButton.Visible = False
Me.ReformatButton.Visible = False
Me.ContinuationButton.Visible = False
```

这些按钮在 Designer 中被**直接设置为不可见**。Office 2021 与其他版本行为一致，不是 Office 2021 特有的问题。用户看到的"灰色不可用"实际上是按钮被隐藏（Visible=False），Ribbon XML 渲染时显示为禁用灰色外观。

另外，`ExcelAi/Ribbon1.vb` 文件头注释写的是 `' WordAi\Ribbon1.vb`，这说明文件最初从 WordAi 复制而来，存在维护混乱的隐患。

### 设计方案

1. 修正 `Ribbon1.vb` 文件头注释（`' ExcelAi\Ribbon1.vb`）
2. 对已完成实现的按钮（`DataAnalysisButton`、`ChatButton`、`TranslateButton`）确认 `Visible = True`
3. 对尚未实现的功能（`ProofreadButton`、`ReformatButton`），保持隐藏或改为显示但在点击时给出"功能开发中"提示（当前 `ProofreadButton_Click` 已有 MessageBox，但按钮隐藏所以用户触达不到）
4. `BatchDataGenButton` 待 Bug 5 实现后改为 `Visible = True`

### 改动文件

| 文件 | 改动点 |
|------|--------|
| `ExcelAi/Ribbon1.vb` | 修正文件头注释 |
| `ExcelAi/Ribbon1.Designer.vb` | 根据实现情况调整各按钮 `Visible` |

---

## 9. AI 结果无法指定位置输出

**优先级：P2**

### 根因分析

通过代码搜索，未找到任何"指定输出位置"的相关实现。Excel ChatControl 的 AI 结果写入目前通过 JSON 命令（`SetCellValue`、`ApplyFormula` 等）完成，但这些命令由 AI 自行决定目标单元格，用户无法在发送前指定"将结果写到 B2"。

### 设计方案

在 Excel ChatControl 的输入区域增加一个"输出到"选区绑定控件（复用现有的 `PendingSelectionInfo` 机制）：

1. 用户可以在 Excel 中先选中目标单元格，然后点击"锁定输出位置"
2. 前端 JS 将该地址附加到 user message 的上下文中：`[输出位置: $B$2]`
3. `BaseChatControl.Send()` 中将锁定的输出位置注入到 `variableValues` 字典（key: `"输出位置"`），系统提示词中声明：`如果用户指定了输出位置，优先将结果写入该位置`

此功能较复杂，建议单独立项，本次文档仅记录设计方向。

---

## 10. WPS 安装后 COM 加载项不显示

**优先级：P2**

### 根因分析

**文件：** `ShareRibbon/Common/LLMUtil.vb` 第 13-19 行

```vb
Public Shared Function IsWpsActive() As Boolean
    Try
        Return Process.GetProcessesByName("WPS").Length > 0
    Catch
        Return False
    End Try
End Function
```

WPS Office 各组件的进程名不是 `"WPS"`，而是：
- `wps`（Writer）
- `et`（表格/Spreadsheets）
- `wpp`（演示/Presentation）
- `wpspdf`（PDF）

因此 `IsWpsActive()` 在 WPS 运行时始终返回 `False`，宽度修复定时器不触发，任务窗格显示宽度异常。

更深层的问题是 **COM 加载项注册路径**。VSTO 加载项通过 `HKCU\Software\Microsoft\Office\<AppName>\Addins` 注册，而 WPS 使用自己的加载项目录（`HKCU\Software\Kingsoft\Office\<AppName>\Addins` 或特定 WPS 插件格式），VSTO 注册表项对 WPS 无效。

### 设计方案

短期：修复进程名检测（覆盖更多 WPS 进程名）：

```vb
Public Shared Function IsWpsActive() As Boolean
    Try
        Dim wpsProcessNames = {"wps", "et", "wpp", "wpspdf", "WPS", "ET", "WPP"}
        Return wpsProcessNames.Any(Function(n) Process.GetProcessesByName(n).Length > 0)
    Catch
        Return False
    End Try
End Function
```

长期：需要调研 WPS 加载项 SDK（`.wll` 格式或 WPS JS 加载项），VSTO 不能直接作为 WPS COM 加载项加载，需要额外封装。建议在安装包中检测 WPS 并给出提示。

### 改动文件

| 文件 | 改动点 |
|------|--------|
| `ShareRibbon/Common/LLMUtil.vb` | `IsWpsActive()` 补充所有 WPS 进程名 |

---

## 11. MSI 被杀毒软件误报

**优先级：P3**

### 根因分析

未签名的 MSI 安装包是杀毒软件误报的最常见原因，与代码逻辑无关。

### 解决方案

1. 购买代码签名证书（EV 证书误报率更低），对 MSI、所有 DLL 和 EXE 进行签名
2. 向主流杀毒厂商（360、腾讯、火绒）提交白名单申请
3. 在官网提供 SHA256 校验码供用户验证

**不需要改动代码。**

---

## 12. 待确认 Bug

以下 Bug 从现有代码中**未能定位根因**，需要用户提供更详细的复现步骤或错误截图：

| Bug | 需要补充的信息 |
|-----|---------------|
| PPT 生成 PPT 报 JS 语法错误 | 哪个功能触发？AI 生成的 PPT JSON 内容示例？错误发生在哪个函数？ |
| 插件加载失败、功能无法启用 | Office 版本？安装方式（MSI/VSTO直装）？错误提示截图？ |
| Excel AI 结果插入格式混乱 | 具体是 Markdown 格式？表格格式？插入到哪里（单元格/注释/工作表）？ |

---

## 改动汇总

| 优先级 | Bug | 涉及文件 | 改动行数（估）|
|--------|-----|----------|-------------|
| P1 | Send 提示词重构 | ChatContextBuilder.vb, BaseChatControl.vb | ~150 行 |
| P0 | PPT 附件解析 | PowerPointAi/ChatControl.vb | ~10 行 |
| P0 | 请求失败弹窗 | BaseChatControl.vb, BaseTranslateService.vb, BaseDataCapturePane.vb | ~6 行 |
| P1 | PPT JS 换行符 | PowerPointAi/Ribbon1.vb | ~3 行 |
| P1 | Excel 批量数据（设计方向） | Ribbon1.vb + 新增 | ~100 行 |
| P1 | API 配置删除 | ConfigApiForm.vb | ~15 行 |
| P1 | 网页爬取跳转 | BaseDataCapturePane.vb | ~10 行 |
| P2 | Office 2021 按钮可见性 | Ribbon1.vb, Ribbon1.Designer.vb | ~8 行 |
| P2 | WPS 进程名检测 | LLMUtil.vb | ~5 行 |
| P3 | MSI 误报 | — | 代码外解决 |
