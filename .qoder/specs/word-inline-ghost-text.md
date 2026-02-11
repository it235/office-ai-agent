# Word 内联灰色补全文本 (Ghost Text) 实现方案

## 概述
将 Word 补全从弹窗方式改为内联灰色文本显示，类似代码编辑器的 ghost text 体验。同时优化 LLM 请求性能。

## 需求
1. 补全建议直接在光标位置显示为灰色文本
2. 按 `Ctrl+.` 接受补全（文本变为正常颜色）
3. 继续输入或移动光标时自动取消（删除灰色文本）
4. 优化 LLM 请求性能，减少等待时间

## 实现方案

### 1. 核心架构变更

**移除**:
- `CompletionPopupForm` 弹窗类（从 `WordCompletionManager.vb` 中移除）
- `PPTCompletionPopupForm`（从 `PowerPointCompletionManager.vb` 中移除）

**新增**:
- `WordGhostTextManager` - 管理 Word 中的灰色内联文本
- `PowerPointGhostTextManager` - 管理 PPT 中的灰色内联文本

### 2. 关键文件修改

| 文件 | 修改内容 |
|------|----------|
| `WordAi/WordCompletionManager.vb` | 移除弹窗逻辑，改用 GhostTextManager |
| `ShareRibbon/Controls/Services/OfficeCompletionService.vb` | HttpClient 单例化 + CancellationToken |
| `PowerPointAi/PowerPointCompletionManager.vb` | 同步修改，使用 GhostTextManager |

### 3. WordGhostTextManager 设计

```vb
Public Class WordGhostTextManager
    Private _wordApp As Word.Application
    Private _ghostRange As Word.Range  ' 跟踪灰色文本的 Range
    Private _originalCursorPos As Integer  ' 记录原始光标位置
    
    ' 显示灰色补全文本
    Public Sub ShowGhostText(suggestion As String)
        ClearGhostText()  ' 先清除旧的
        
        Dim sel = _wordApp.Selection
        _originalCursorPos = sel.Range.Start
        
        ' 在光标位置插入灰色文本
        _ghostRange = sel.Range.Duplicate
        _ghostRange.Collapse(WdCollapseDirection.wdCollapseEnd)
        _ghostRange.Text = suggestion
        _ghostRange.Font.Color = CType(RGB(150, 150, 150), WdColor)  ' 灰色
        
        ' 将光标移回原位（不选中灰色文本）
        sel.SetRange(_originalCursorPos, _originalCursorPos)
    End Sub
    
    ' 接受补全 - 将灰色文本变为正常颜色
    Public Sub AcceptGhostText()
        If _ghostRange IsNot Nothing Then
            _ghostRange.Font.ColorIndex = WdColorIndex.wdAuto
            _wordApp.Selection.SetRange(_ghostRange.End, _ghostRange.End)
            _ghostRange = Nothing
        End If
    End Sub
    
    ' 清除灰色文本
    Public Sub ClearGhostText()
        If _ghostRange IsNot Nothing Then
            Try
                _ghostRange.Delete()
            Catch
                ' Range 可能已失效
            End Try
            _ghostRange = Nothing
        End If
    End Sub
    
    ' 检查是否有活动的 ghost text
    Public ReadOnly Property HasGhostText As Boolean
        Get
            Return _ghostRange IsNot Nothing
        End Get
    End Property
End Class
```

### 4. 性能优化 - OfficeCompletionService

```vb
' 单例 HttpClient（避免频繁创建连接）
Private Shared ReadOnly _httpClient As New HttpClient() With {
    .Timeout = TimeSpan.FromSeconds(8)  ' 缩短超时
}

' 取消令牌支持
Private _cancellationTokenSource As CancellationTokenSource

Public Sub CancelPendingRequest()
    _cancellationTokenSource?.Cancel()
    _cancellationTokenSource = New CancellationTokenSource()
End Sub

Public Async Function GetCompletionsDirectAsync(inputText As String, appType As String, 
    Optional token As CancellationToken = Nothing) As Task(Of List(Of String))
    ' 使用 token 进行请求取消
End Function
```

### 5. 快捷键处理

由于 Word 编辑区无法直接捕获 `Ctrl+.`，需要在 `OnSelectionChange` 中检测：
- 当存在 ghost text 且光标在 ghost text 起始位置时
- 检测按键状态（通过 `GetAsyncKeyState` Win32 API）

**简化方案**：在 `WordCompletionManager` 中添加公开方法 `AcceptCurrentCompletion()`，通过 Ribbon 按钮或快捷键触发。

### 6. 交互流程

```
用户输入 → 800ms 防抖 → LLM 请求 → 显示灰色文本
                                          ↓
                              用户继续输入 → 清除灰色文本
                              用户按 Ctrl+. → 接受（变黑）
                              用户移动光标 → 清除灰色文本
```

### 7. 边界情况处理

1. **Range 失效**: 在 `ShowGhostText` 前检查文档状态
2. **Undo 栈**: Ghost text 的插入/删除会进入撤销列表（Word 的限制）
3. **多线程**: 所有 Word 操作通过 `SynchronizationContext.Post` 到主线程

## 验证方法

1. **功能测试**:
   - 在 Word 中输入文本，等待 800ms 后应显示灰色补全
   - 按 `Ctrl+.` 应接受补全（灰色变黑）
   - 继续输入应清除灰色文本
   
2. **性能测试**:
   - 连续快速输入不应导致界面卡顿
   - LLM 请求应能正确取消

3. **编译验证**:
   ```bash
   msbuild AiHelper.sln /t:Build /p:Configuration=Debug
   ```

## 实现顺序

1. 优化 `OfficeCompletionService`（HttpClient 单例 + CancellationToken）
2. 创建 `WordGhostTextManager` 类
3. 修改 `WordCompletionManager` 使用新的 ghost text 逻辑
4. 移除 `CompletionPopupForm` 相关代码
5. 同步修改 PowerPoint 部分
