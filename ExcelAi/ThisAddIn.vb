Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports ExcelAi.ExcelAiFunctions
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Win32
Imports ShareRibbon

Public Class ThisAddIn

    ' 在类中添加以下变量
    Private _deepseekControl As DeepseekControl
    Private _deepseekTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private _doubaoControl As DoubaoChat
    Private _doubaoTaskPane As Microsoft.Office.Tools.CustomTaskPane

    Private chatTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Public Shared chatControl As ChatControl

    ' XLL 路径缓存：首次找到后不再重复搜索
    Private Shared _cachedXllPath As String = Nothing

    ' 延迟初始化：WebView2 和 SQLite 仅在首次使用时加载
    Private _lazyWebView2 As New Lazy(Of Boolean)(Function()
        WebView2Loader.EnsureWebView2Loader()
        Return True
    End Function)

    Private _lazySqlite As New Lazy(Of Boolean)(Function()
        SqliteNativeLoader.EnsureLoaded()
        Return True
    End Function)

    ' WPS 宽度修复定时器
    Private widthTimer As Timer
    Private widthTimer1 As Timer

    Private Sub ExcelAi_Startup() Handles Me.Startup
        ' SqliteAssemblyResolver 必须最先注册，确保后续加载 SQLite 程序集时能解析
        SqliteAssemblyResolver.EnsureRegistered()
        Try
            Debug.WriteLine("正在初始化GlobalStatusStrip...")
            GlobalStatusStripAll.InitializeApplication(Me.Application)
            Debug.WriteLine("GlobalStatusStrip初始化完成")
        Catch ex As Exception
            Debug.WriteLine("初始化GlobalStatusStrip时出错: " & ex.Message)
            MessageBox.Show("初始化状态栏时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ' WebView2 延迟加载：推迟到首次打开任务窗格时初始化
        ' SQLite 原生库延迟加载：首次访问数据库时由 OfficeAiDatabase.EnsureInitialized() 自动调用

        ' XLL 注册策略：
        '   - 已缓存路径 → 直接在主线程注册（无磁盘 I/O，极快）
        '   - 首次搜索  → 路径搜索（含磁盘扫描）放到后台线程，
        '                 RegisterXLL（COM 调用）通过 SynchronizationContext 回到 STA 主线程执行
        '                 这样 Office 启动不再被磁盘扫描阻塞
        If Not String.IsNullOrEmpty(_cachedXllPath) Then
            Try
                Application.RegisterXLL(_cachedXllPath)
            Catch ex As Exception
                Debug.WriteLine($"[XLL] 注册缓存路径失败: {ex.Message}")
            End Try
        Else
            Dim syncCtx = System.Threading.SynchronizationContext.Current
            Dim excelApp = Me.Application
            Task.Run(Sub()
                Try
                    Dim xllPath = FindXllPath()
                    If Not String.IsNullOrEmpty(xllPath) Then
                        ' RegisterXLL 是 COM 调用，必须回到 STA 主线程
                        syncCtx.Post(Sub(state)
                            Try
                                excelApp.RegisterXLL(CStr(state))
                                _cachedXllPath = CStr(state)
                                Debug.WriteLine($"[XLL] 已注册: {state}")
                            Catch regEx As Exception
                                Debug.WriteLine($"[XLL] RegisterXLL 失败: {regEx.Message}")
                            End Try
                        End Sub, xllPath)
                    Else
                        Debug.WriteLine("[XLL] 未找到 XLL 文件，跳过注册")
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[XLL] 后台搜索失败: {ex.Message}")
                End Try
            End Sub)
        End If
    End Sub

    ''' <summary>
    ''' 确保核心服务已加载（WebView2 + SQLite），首次调用时初始化
    ''' </summary>
    Private Sub EnsureCoreServicesLoaded()
        Try
            Dim webView2Init = _lazyWebView2.Value
        Catch ex As Exception
            MessageBox.Show($"WebView2 初始化失败: {ex.Message}")
        End Try
        Try
            Dim sqliteInit = _lazySqlite.Value
        Catch ex As Exception
            MessageBox.Show($"SQLite 原生库加载失败，Skills/记忆功能可能不可用: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 确保 WPS 宽度修复定时器已初始化（仅在需要时创建）
    ''' </summary>
    Private Sub EnsureWidthTimers()
        If widthTimer Is Nothing Then
            widthTimer = New Timer()
            AddHandler widthTimer.Tick, AddressOf WidthTimer_Tick
            widthTimer.Interval = 100
        End If
        If widthTimer1 Is Nothing Then
            widthTimer1 = New Timer()
            AddHandler widthTimer1.Tick, AddressOf WidthTimer1_Tick
            widthTimer1.Interval = 200
        End If
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ' 清理定时器资源
        If widthTimer IsNot Nothing Then
            widthTimer.Stop()
            widthTimer.Dispose()
            widthTimer = Nothing
        End If
        If widthTimer1 IsNot Nothing Then
            widthTimer1.Stop()
            widthTimer1.Dispose()
            widthTimer1 = Nothing
        End If
    End Sub

    ''' <summary>
    ''' 在后台线程中搜索 XLL 文件路径，不包含任何 COM 调用，可安全在非 STA 线程执行。
    ''' 搜索优先级：注册表安装路径 → 标准安装目录 → 当前目录 → 向上遍历 → 全盘扫描（最慢，最后执行）。
    ''' 返回找到的 XLL 路径，未找到则返回 Nothing。
    ''' </summary>
    Private Shared Function FindXllPath() As String
        Try
            Dim currentAssemblyPath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
            Dim currentDir As String = Path.GetDirectoryName(currentAssemblyPath)
            Dim searchPaths As New List(Of String)

            ' 1. 注册表路径（安装时写入，最可靠，几乎零开销）
            Try
                Using key As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\it235\OfficeAiAgent")
                    If key IsNot Nothing Then
                        Dim installPath As String = key.GetValue("InstallPath")?.ToString()
                        If Not String.IsNullOrEmpty(installPath) Then
                            Dim excelAiPath As String = Path.Combine(installPath, "ExcelAi")
                            If Directory.Exists(excelAiPath) Then searchPaths.Add(excelAiPath)
                        End If
                    End If
                End Using
            Catch ex As Exception
                Debug.WriteLine($"[XLL] 读取注册表失败: {ex.Message}")
            End Try

            ' 2. 标准安装路径（固定路径，Directory.Exists 开销极低）
            Dim standardInstallPaths As String() = {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "it235", "OfficeAiAgent", "ExcelAi"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "it235", "OfficeAiAgent", "ExcelAi"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "OfficeAiAgent", "ExcelAi"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "OfficeAiAgent", "ExcelAi")
            }
            For Each installPath In standardInstallPaths
                If Directory.Exists(installPath) Then searchPaths.Add(installPath)
            Next

            ' 3. 当前目录及父目录
            searchPaths.Add(currentDir)
            searchPaths.Add(Path.Combine(currentDir, ".."))

            ' 4. 开发环境：VSTO 影子复制目录时，尝试真实项目输出目录
            If currentDir.Contains("AppData\Local\assembly") Then
                Dim possibleDevPaths As String() = {
                    "F:\ai\code\AiHelper\ExcelAi\bin\Debug",
                    "F:\ai\code\AiHelper\ExcelAi\bin\Release",
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "source", "repos", "AiHelper", "ExcelAi", "bin", "Debug"),
                    Path.Combine("C:\", "ai", "code", "AiHelper", "ExcelAi", "bin", "Debug")
                }
                For Each devPath In possibleDevPaths
                    If Directory.Exists(devPath) Then searchPaths.Add(devPath)
                Next
            End If

            ' 5. 向上遍历目录树，查找 OfficeAiAgent 安装目录
            Dim currentDirInfo As New DirectoryInfo(currentDir)
            While currentDirInfo IsNot Nothing
                If currentDirInfo.Name.Equals("OfficeAiAgent", StringComparison.OrdinalIgnoreCase) OrElse
                   currentDirInfo.FullName.Contains("OfficeAiAgent") Then
                    Dim excelAiPath As String = Path.Combine(currentDirInfo.FullName, "ExcelAi")
                    If Directory.Exists(excelAiPath) Then searchPaths.Add(excelAiPath)
                    searchPaths.Add(currentDirInfo.FullName)
                    Exit While
                End If
                currentDirInfo = currentDirInfo.Parent
            End While

            ' 6. 全盘扫描（最慢，仅在前几步都找不到时才执行；后台线程中不阻塞 Office UI）
            Try
                For Each drive In DriveInfo.GetDrives()
                    If drive.IsReady AndAlso drive.DriveType = DriveType.Fixed Then
                        Dim customPaths As String() = {
                            Path.Combine(drive.Name, "OfficeAiAgent", "ExcelAi"),
                            Path.Combine(drive.Name, "Program Files", "OfficeAiAgent", "ExcelAi"),
                            Path.Combine(drive.Name, "Program Files (x86)", "OfficeAiAgent", "ExcelAi")
                        }
                        For Each customPath In customPaths
                            If Directory.Exists(customPath) Then searchPaths.Add(customPath)
                        Next
                    End If
                Next
            Catch ex As Exception
                Debug.WriteLine($"[XLL] 全盘扫描出错: {ex.Message}")
            End Try

            ' 按优先级逐路径检查，找到即返回
            Dim xllFileName As String = If(IntPtr.Size = 8, "ExcelAi-AddIn64-packed.xll", "ExcelAi-AddIn-packed.xll")
            Dim unpackedFileName As String = If(IntPtr.Size = 8, "ExcelAi-AddIn64.xll", "ExcelAi-AddIn.xll")
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each searchPath In searchPaths
                If String.IsNullOrEmpty(searchPath) OrElse Not seen.Add(searchPath) Then Continue For
                If Not Directory.Exists(searchPath) Then Continue For

                Dim packed = Path.Combine(searchPath, xllFileName)
                If File.Exists(packed) Then
                    Debug.WriteLine($"[XLL] 找到: {packed}")
                    Return packed
                End If

                Dim unpacked = Path.Combine(searchPath, unpackedFileName)
                If File.Exists(unpacked) Then
                    Debug.WriteLine($"[XLL] 找到（未打包）: {unpacked}")
                    Return unpacked
                End If
            Next

            Debug.WriteLine("[XLL] 所有路径均未找到 XLL 文件")
            Return Nothing

        Catch ex As Exception
            Debug.WriteLine($"[XLL] FindXllPath 出错: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ' 创建聊天任务窗格
    Private Sub CreateChatTaskPane()
        Try
            ' 为新工作簿创建任务窗格
            chatControl = New ChatControl()
            chatTaskPane = Me.CustomTaskPanes.Add(chatControl, "Excel AI智能助手")
            chatTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            chatTaskPane.Width = 420
            'AddHandler chatTaskPane.VisibleChanged, AddressOf ChatTaskPane_VisibleChanged
            'chatTaskPane.Visible = False
        Catch ex As Exception
            MessageBox.Show($"初始化任务窗格失败: {ex.Message}")
        End Try
    End Sub

    Private Sub CreateDeepseekTaskPane()
        Try
            If _deepseekControl Is Nothing Then
                ' 为新工作簿创建任务窗格
                _deepseekControl = New DeepseekControl()
                _deepseekTaskPane = Me.CustomTaskPanes.Add(_deepseekControl, "Deepseek AI智能助手")
                _deepseekTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
                _deepseekTaskPane.Width = 420
                'AddHandler _deepseekTaskPane.VisibleChanged, AddressOf DeepseekTaskPane_VisibleChanged
                '_deepseekTaskPane.Visible = False
            End If
        Catch ex As Exception
            MessageBox.Show($"初始化任务窗格失败: {ex.Message}")
        End Try
    End Sub

    Private Async Function CreateDoubaoTaskPane() As Task
        Try
            If _doubaoControl Is Nothing Then
                ' 为新工作簿创建任务窗格
                _doubaoControl = New DoubaoChat()
                _doubaoTaskPane = Me.CustomTaskPanes.Add(_doubaoControl, "Doubao AI智能助手")
                _doubaoTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
                _doubaoTaskPane.Width = 420
            End If
        Catch ex As Exception
            MessageBox.Show($"初始化Doubao任务窗格失败: {ex.Message}")
        End Try
    End Function

    ' 解决WPS中无法显示正常宽度的问题
    Private Sub ChatTaskPane_VisibleChanged(sender As Object, e As EventArgs)
        Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = CType(sender, Microsoft.Office.Tools.CustomTaskPane)
        If taskPane.Visible Then
            If LLMUtil.IsWpsActive() Then
                EnsureWidthTimers()
                widthTimer.Start()
            End If
        End If
    End Sub

    Private Sub DeepseekTaskPane_VisibleChanged(sender As Object, e As EventArgs)
        Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = CType(sender, Microsoft.Office.Tools.CustomTaskPane)
        If taskPane.Visible Then
            If LLMUtil.IsWpsActive() Then
                EnsureWidthTimers()
                widthTimer1.Start()
            End If
        End If
    End Sub

    Private Sub WidthTimer_Tick(sender As Object, e As EventArgs)
        widthTimer.Stop()
        If LLMUtil.IsWpsActive() AndAlso chatTaskPane IsNot Nothing Then
            chatTaskPane.Width = 420
        End If
    End Sub

    Private Sub WidthTimer1_Tick(sender As Object, e As EventArgs)
        widthTimer1.Stop()
        Debug.WriteLine($"Deepseek点击定时1")
        If LLMUtil.IsWpsActive() AndAlso _deepseekTaskPane IsNot Nothing Then
            Debug.WriteLine($"Deepseek点击定时2")
            _deepseekTaskPane.Width = 420
        End If
    End Sub

    Dim loadChatHtml As Boolean = True

    Public Async Sub ShowChatTaskPane()
        EnsureCoreServicesLoaded()
        CreateChatTaskPane()
        If chatTaskPane Is Nothing Then Return
        chatTaskPane.Visible = True
        If loadChatHtml Then
            loadChatHtml = False
            Await chatControl.LoadLocalHtmlFile()
        End If
    End Sub

    Public Async Sub ShowDeepseekTaskPane()
        Debug.WriteLine($"Deepseek点击事件")
        EnsureCoreServicesLoaded()
        CreateDeepseekTaskPane()
        If _deepseekTaskPane Is Nothing Then Return
        _deepseekTaskPane.Visible = True
    End Sub

    Public Async Sub ShowDoubaoTaskPane()
        EnsureCoreServicesLoaded()
        Await CreateDoubaoTaskPane()
        If _doubaoTaskPane Is Nothing Then Return
        _doubaoTaskPane.Visible = True
    End Sub
End Class
