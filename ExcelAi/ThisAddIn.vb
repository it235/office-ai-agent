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
        ' WebView2 延迟加载：推迟到首次打开任务窗格时初始化，减少启动耗时
        ' SQLite 原生库延迟加载：首次访问数据库时由 OfficeAiDatabase.EnsureInitialized() 自动调用
        ' 此处无需显式调用 SqliteNativeLoader.EnsureLoaded()

        ' 延迟加载 Excel-DNA XLL：仅在启动时执行一次路径搜索并缓存
        Try
            LoadExcelDnaAddIn()
        Catch ex As Exception
            Debug.WriteLine($"加载 Excel-DNA 失败: {ex.Message}")
            ' 继续执行，不要因为这个错误而中断启动
        End Try

    End Sub

    ''' <summary>
    ''' 确保核心服务已加载（WebView2 + SQLite），首次调用时初始化
    ''' </summary>
    Private Sub EnsureCoreServicesLoaded()
        Try
            Dim _ = _lazyWebView2.Value
        Catch ex As Exception
            MessageBox.Show($"WebView2 初始化失败: {ex.Message}")
        End Try
        Try
            Dim _ = _lazySqlite.Value
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

    Private Sub LoadExcelDnaAddIn()
        Try
            ' 使用缓存路径：若已找到 XLL，直接加载，跳过搜索
            If Not String.IsNullOrEmpty(_cachedXllPath) Then
                Debug.WriteLine($"[LoadExcelDnaAddIn] 使用缓存路径: {_cachedXllPath}")
                Application.RegisterXLL(_cachedXllPath)
                Return
            End If

            Debug.WriteLine("开始查找 XLL 文件...")

            ' 获取当前程序集路径
            Dim currentAssemblyPath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
            Dim currentDir As String = Path.GetDirectoryName(currentAssemblyPath)

            Debug.WriteLine($"当前程序集路径: {currentAssemblyPath}")
            Debug.WriteLine($"当前目录: {currentDir}")

            ' 创建搜索路径列表（按优先级排序）
            Dim searchPaths As New List(Of String)

            ' 1. 注册表路径（安装时写入，最可靠）
            Try
                Using key As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\it235\OfficeAiAgent")
                    If key IsNot Nothing Then
                        Dim installPath As String = key.GetValue("InstallPath")?.ToString()
                        If Not String.IsNullOrEmpty(installPath) Then
                            Dim excelAiPath As String = Path.Combine(installPath, "ExcelAi")
                            If Directory.Exists(excelAiPath) Then
                                searchPaths.Add(excelAiPath)
                                Debug.WriteLine($"从注册表添加路径: {excelAiPath}")
                            End If
                        End If
                    End If
                End Using
            Catch ex As Exception
                Debug.WriteLine($"从注册表读取安装路径时出错: {ex.Message}")
            End Try

            ' 2. 标准安装路径
            Dim standardInstallPaths As String() = {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "it235", "OfficeAiAgent", "ExcelAi"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "it235", "OfficeAiAgent", "ExcelAi"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "OfficeAiAgent", "ExcelAi"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "OfficeAiAgent", "ExcelAi")
            }
            For Each installPath In standardInstallPaths
                If Directory.Exists(installPath) Then
                    searchPaths.Add(installPath)
                    Debug.WriteLine($"添加标准安装路径: {installPath}")
                End If
            Next

            ' 3. 当前目录及其相关目录
            searchPaths.Add(currentDir)
            searchPaths.Add(Path.Combine(currentDir, ".."))

            ' 4. 开发环境：VSTO 临时目录时尝试真正的项目输出目录
            If currentDir.Contains("AppData\Local\assembly") Then
                Dim possibleDevPaths As String() = {
                    "F:\ai\code\AiHelper\ExcelAi\bin\Debug",
                    "F:\ai\code\AiHelper\ExcelAi\bin\Release",
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "source", "repos", "AiHelper", "ExcelAi", "bin", "Debug"),
                    Path.Combine("C:\", "ai", "code", "AiHelper", "ExcelAi", "bin", "Debug")
                }
                For Each devPath In possibleDevPaths
                    If Directory.Exists(devPath) Then
                        searchPaths.Add(devPath)
                        Debug.WriteLine($"添加开发路径: {devPath}")
                    End If
                Next
            End If

            ' 5. 从当前路径向上查找 OfficeAiAgent 目录
            Dim currentDirInfo As New DirectoryInfo(currentDir)
            While currentDirInfo IsNot Nothing
                If currentDirInfo.Name.Equals("OfficeAiAgent", StringComparison.OrdinalIgnoreCase) OrElse
                   currentDirInfo.FullName.Contains("OfficeAiAgent") Then
                    Dim excelAiPath As String = Path.Combine(currentDirInfo.FullName, "ExcelAi")
                    If Directory.Exists(excelAiPath) Then
                        searchPaths.Add(excelAiPath)
                        Debug.WriteLine($"添加安装路径: {excelAiPath}")
                    End If
                    searchPaths.Add(currentDirInfo.FullName)
                    Exit While
                End If
                currentDirInfo = currentDirInfo.Parent
            End While

            ' 6. 广泛搜索：遍历所有固定驱动器（最慢，最后执行）
            Try
                For Each drive In DriveInfo.GetDrives()
                    If drive.IsReady AndAlso drive.DriveType = DriveType.Fixed Then
                        Dim customPaths As String() = {
                            Path.Combine(drive.Name, "OfficeAiAgent", "ExcelAi"),
                            Path.Combine(drive.Name, "Program Files", "OfficeAiAgent", "ExcelAi"),
                            Path.Combine(drive.Name, "Program Files (x86)", "OfficeAiAgent", "ExcelAi")
                        }
                        For Each customPath In customPaths
                            If Directory.Exists(customPath) Then
                                searchPaths.Add(customPath)
                                Debug.WriteLine($"添加自定义路径: {customPath}")
                            End If
                        Next
                    End If
                Next
            Catch ex As Exception
                Debug.WriteLine($"搜索自定义路径时出错: {ex.Message}")
            End Try

            ' 去除重复路径
            Dim uniquePaths As New HashSet(Of String)(searchPaths, StringComparer.OrdinalIgnoreCase)

            ' 查找XLL文件
            Dim xllFileName As String = If(IntPtr.Size = 8, "ExcelAi-AddIn64-packed.xll", "ExcelAi-AddIn-packed.xll")
            Dim foundXllPath As String = String.Empty

            Debug.WriteLine($"正在查找文件: {xllFileName}")

            For Each searchPath In uniquePaths
                Debug.WriteLine($"  检查: {searchPath}")
                If Directory.Exists(searchPath) Then
                    Dim xllPath As String = Path.Combine(searchPath, xllFileName)
                    If File.Exists(xllPath) Then
                        foundXllPath = xllPath
                        Debug.WriteLine($"找到XLL文件: {xllPath}")
                        Exit For
                    End If
                    ' 也检查未打包的版本
                    Dim unpackedXllFileName As String = If(IntPtr.Size = 8, "ExcelAi-AddIn64.xll", "ExcelAi-AddIn.xll")
                    Dim unpackedXllPath As String = Path.Combine(searchPath, unpackedXllFileName)
                    If File.Exists(unpackedXllPath) Then
                        foundXllPath = unpackedXllPath
                        Debug.WriteLine($"找到未打包XLL文件: {unpackedXllPath}")
                        Exit For
                    End If
                End If
            Next

            ' 尝试加载XLL文件，成功后缓存路径供下次直接使用
            If Not String.IsNullOrEmpty(foundXllPath) Then
                Try
                    Debug.WriteLine($"正在加载 Excel-DNA XLL: {foundXllPath}")
                    Dim result As Boolean = Application.RegisterXLL(foundXllPath)
                    If result Then
                        ' 缓存成功找到的路径，避免下次重复搜索
                        _cachedXllPath = foundXllPath
                        Debug.WriteLine($"[LoadExcelDnaAddIn] 路径已缓存: {foundXllPath}")
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"加载 XLL 时出错: {ex.Message}")
                End Try
            Else
                Debug.WriteLine("未找到 XLL 文件，诊断信息:")
                For Each searchPath In uniquePaths
                    If Directory.Exists(searchPath) Then
                        Debug.WriteLine($"路径 {searchPath} 包含的文件:")
                        Try
                            For Each file In Directory.GetFiles(searchPath, "*.xll")
                                Debug.WriteLine($"  XLL文件: {Path.GetFileName(file)}")
                            Next
                            For Each file In Directory.GetFiles(searchPath, "ExcelAi*.*")
                                Debug.WriteLine($"  ExcelAi文件: {Path.GetFileName(file)}")
                            Next
                        Catch ex As Exception
                            Debug.WriteLine($"  无法读取目录内容: {ex.Message}")
                        End Try
                    Else
                        Debug.WriteLine($"路径不存在: {searchPath}")
                    End If
                Next
            End If

        Catch ex As Exception
            Debug.WriteLine($"LoadExcelDnaAddIn 出错: {ex.Message}")
            Debug.WriteLine($"堆栈跟踪: {ex.StackTrace}")
        End Try
    End Sub

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
