Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ChatSettings
    Private ReadOnly _applicationInfo As ApplicationInfo
    Public Property EnabledMcpList As List(Of String)

    Public Sub New(applicationInfo As ApplicationInfo)
        _applicationInfo = applicationInfo

        ' 初始化MCP列表
        EnabledMcpList = New List(Of String)()

        LoadSettings()
    End Sub

    Public Shared Property topicRandomness As Double = 0.8  ' 默认值改为 Double
    Public Shared Property contextLimit As Integer = 5     ' 默认值改为 Integer
    ' E4-04 选区自动引用：默认开启——用户选中了一段内容再提问，最自然的期待就是「问的是这段」。
    ' 代价必须同时给出来：带出去的是什么要在对话里看得见（BaseChatControl.BuildMessageWithAutoSelection），
    ' 并且面板设置里有能真的关掉它的开关（缺键时才用这个默认值，显式存过 false 的用户保持关闭）。
    Public Shared Property selectedCellChecked As Boolean = True
    Public Shared Property executecodePreviewChecked As Boolean = True ' 执行代码前预览的默认选项
    Public Shared Property settingsScrollChecked As Boolean = True

    ' 自动补全设置（已禁用 - 观察期）
    Public Shared Property EnableAutocomplete As Boolean = False     ' 自动补全功能已禁用，默认关闭
    Public Shared Property AutocompleteDelayMs As Integer = 800      ' 防抖延迟（毫秒）
    Public Shared Property AutocompleteShortcut As String = "Ctrl+."  ' 接受补全的快捷键（默认Ctrl+.）

    ' E4-01 选区工具条：选中一段文字后在选区旁给出快捷动作。
    ' 默认关闭：浮层会自动出现在每一次选区旁边，不能未经用户同意就改变写作界面。
    ' 面板设置里有 `settings-enable-selection-toolbar` 开关，用户不改代码就能关掉它。
    Public Shared Property EnableSelectionToolbar As Boolean = False

    ' E4-02 右键菜单入口：把本宿主的 AI 动作挂进上下文菜单（三端共用同一份动作清单）。
    ' 同样默认关闭：它改的是用户宿主里既有的右键菜单；开关是 `settings-enable-context-menu`。
    Public Shared Property EnableContextMenu As Boolean = False

    ' E4-03 全局快捷键唤起命令面板：一个组合键在任一宿主里展开面板并打开命令面板。
    ' 默认关闭 + 必须带 Ctrl/Alt：快捷键是独占资源，抢用户既有按键（Word 的 Ctrl+K 是插入超链接）
    ' 比浮层更 intrusive，所以既不默认开，也不允许注册无修饰键的单键。
    Public Shared Property EnableHotkey As Boolean = False
    Public Const DefaultHotkeyChord As String = "Ctrl+Alt+K"
    Public Shared Property HotkeyChord As String = DefaultHotkeyChord

    ' E4-05：宿主在「面板变为可见」时就要读开关（挂右键菜单、出选区工具条），而那一刻可能还
    ' 没有任何 ChatSettings 实例读过文件，于是拿到的是声明的默认值——用户明明开过的开关，
    ' 进程内第一次就是不生效。EnsureLoaded 把「先加载再判断」固定成一次调用，重复调用不再读盘；
    ' 面板自身的装载路径照旧每次 New 一份，跨宿主的外部编辑仍然会在下一次装载时被刷进来。
    Private Shared _loadedInstance As ChatSettings

    Public Shared Function EnsureLoaded(applicationInfo As ApplicationInfo) As ChatSettings
        If _loadedInstance IsNot Nothing Then Return _loadedInstance
        ' 构造即 LoadSettings：静态属性就是这份文件投射出来的当前值。
        _loadedInstance = New ChatSettings(applicationInfo)
        Return _loadedInstance
    End Function

    ' 修改方法签名，参数类型改为 Double 和 Integer
    Public Sub SaveSettings(topicRandomness As Double, contextLimit As Integer,
                          selectedCell As Boolean, settingsScroll As Boolean, executecodePreview As Boolean,
                          Optional enableAutocomplete As Boolean = False, Optional autocompleteDelayMs As Integer = 800,
                          Optional autocompleteShortcut As String = "Ctrl+.",
                          Optional enableSelectionToolbar As Boolean = False,
                          Optional enableContextMenu As Boolean = False,
                          Optional enableHotkey As Boolean = False,
                          Optional hotkeyChord As String = "Ctrl+Alt+K")
        Try
            Dim settingsPath = _applicationInfo.GetChatSettingsFilePath()

            ' 确保目录存在
            Directory.CreateDirectory(Path.GetDirectoryName(settingsPath))

            ' 在既有文件上改这几个键：这里曾经是「新建一个字典整份覆盖」，
            ' 结果一保存设置就把 enableMcpList 之类没在这张表里的键全抹掉。
            Dim settings As JObject
            If File.Exists(settingsPath) Then
                Try
                    settings = JObject.Parse(File.ReadAllText(settingsPath))
                Catch ex As Exception
                    Debug.WriteLine($"设置文件不是合法 JSON，改为重建: {ex.Message}")
                    settings = New JObject()
                End Try
            Else
                settings = New JObject()
            End If

            settings("topicRandomness") = topicRandomness
            settings("contextLimit") = contextLimit
            settings("selectedCellChecked") = selectedCell
            settings("settingsScrollChecked") = settingsScroll
            settings("executecodePreviewChecked") = executecodePreview
            settings("enableAutocomplete") = enableAutocomplete
            settings("autocompleteDelayMs") = autocompleteDelayMs
            settings("autocompleteShortcut") = autocompleteShortcut
            settings("enableSelectionToolbar") = enableSelectionToolbar
            settings("enableContextMenu") = enableContextMenu
            settings("enableHotkey") = enableHotkey
            settings("hotkeyChord") = hotkeyChord

            File.WriteAllText(settingsPath, settings.ToString(Formatting.Indented))

            ' 更新静态属性
            ChatSettings.topicRandomness = topicRandomness
            ChatSettings.contextLimit = contextLimit
            ChatSettings.selectedCellChecked = selectedCell
            ChatSettings.settingsScrollChecked = settingsScroll
            ChatSettings.executecodePreviewChecked = executecodePreview
            ChatSettings.EnableAutocomplete = enableAutocomplete
            ChatSettings.AutocompleteDelayMs = autocompleteDelayMs
            ChatSettings.AutocompleteShortcut = autocompleteShortcut
            ChatSettings.EnableSelectionToolbar = enableSelectionToolbar
            ChatSettings.EnableContextMenu = enableContextMenu
            ChatSettings.EnableHotkey = enableHotkey
            ChatSettings.HotkeyChord = hotkeyChord

        Catch ex As Exception
            Debug.WriteLine($"保存设置失败: {ex.Message}")
        End Try
    End Sub

    ' 加载设置时进行类型转换
    Public Sub LoadSettings()
        Try
            Dim settingsPath = _applicationInfo.GetChatSettingsFilePath()

            If File.Exists(settingsPath) Then
                ' 读取JSON文件
                Dim json = File.ReadAllText(settingsPath)
                Dim settings = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(json)

                ' 更新静态属性，添加类型转换
                If settings.ContainsKey("topicRandomness") Then
                    topicRandomness = Convert.ToDouble(settings("topicRandomness"))
                End If
                If settings.ContainsKey("contextLimit") Then
                    contextLimit = Convert.ToInt32(settings("contextLimit"))
                End If
                If settings.ContainsKey("selectedCellChecked") Then
                    selectedCellChecked = CBool(settings("selectedCellChecked"))
                End If
                If settings.ContainsKey("settingsScrollChecked") Then
                    settingsScrollChecked = CBool(settings("settingsScrollChecked"))
                End If
                If settings.ContainsKey("executecodePreviewChecked") Then
                    executecodePreviewChecked = CBool(settings("executecodePreviewChecked"))
                End If
                If settings.ContainsKey("enableAutocomplete") Then
                    EnableAutocomplete = CBool(settings("enableAutocomplete"))
                End If
                If settings.ContainsKey("autocompleteDelayMs") Then
                    AutocompleteDelayMs = Convert.ToInt32(settings("autocompleteDelayMs"))
                End If
                If settings.ContainsKey("autocompleteShortcut") Then
                    AutocompleteShortcut = Convert.ToString(settings("autocompleteShortcut"))
                End If
                If settings.ContainsKey("enableSelectionToolbar") Then
                    EnableSelectionToolbar = CBool(settings("enableSelectionToolbar"))
                End If
                If settings.ContainsKey("enableContextMenu") Then
                    EnableContextMenu = CBool(settings("enableContextMenu"))
                End If
                If settings.ContainsKey("enableHotkey") Then
                    EnableHotkey = CBool(settings("enableHotkey"))
                End If
                If settings.ContainsKey("hotkeyChord") Then
                    ' 读回来就归一化一次：文件里写了个抢用户单键的组合，不能让它在任何一个宿主上注册成功。
                    Dim parsed = HostHotkeyChord.Parse(Convert.ToString(settings("hotkeyChord")))
                    If parsed IsNot Nothing AndAlso parsed.IsSafe Then
                        HotkeyChord = parsed.ToString()
                    Else
                        HotkeyChord = DefaultHotkeyChord
                        EnableHotkey = False
                    End If
                End If
                ' 加载MCP列表

                ' 加载MCP列表 - 修复了键不存在的问题
                If settings.ContainsKey("enableMcpList") AndAlso settings("enableMcpList") IsNot Nothing Then
                    Try
                        ' 处理不同可能的类型
                        If TypeOf settings("enableMcpList") Is JArray Then
                            EnabledMcpList = CType(settings("enableMcpList"), JArray).ToObject(Of List(Of String))()
                        ElseIf TypeOf settings("enableMcpList") Is String Then
                            EnabledMcpList = JsonConvert.DeserializeObject(Of List(Of String))(settings("enableMcpList").ToString())
                        Else
                            ' 尝试通过转换字符串再反序列化
                            EnabledMcpList = JsonConvert.DeserializeObject(Of List(Of String))(JsonConvert.SerializeObject(settings("enableMcpList")))
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"解析enableMcpList失败: {ex.Message}")
                        EnabledMcpList = New List(Of String)() ' 确保始终有一个有效的列表
                    End Try
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"加载ChatSettings失败: {ex.Message}")
        End Try
    End Sub

    Public Sub SaveEnabledMcpList(enabledList As List(Of String))
        EnabledMcpList = enabledList

        Dim settingsPath = _applicationInfo.GetChatSettingsFilePath()
        ' 读取现有设置
        Dim settingsObj As JObject
        If File.Exists(settingsPath) Then
            Dim jsonContent = File.ReadAllText(settingsPath)
            settingsObj = JObject.Parse(jsonContent)
        Else
            settingsObj = New JObject()
        End If

        ' 更新MCP列表
        settingsObj("enableMcpList") = JArray.FromObject(enabledList)

        ' 保存设置
        File.WriteAllText(settingsPath, settingsObj.ToString(Newtonsoft.Json.Formatting.Indented))
    End Sub
End Class