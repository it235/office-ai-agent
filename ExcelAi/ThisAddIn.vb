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
    Private translateService As ExcelTranslateService

    Private Sub ExcelAi_Startup() Handles Me.Startup
        Try
            Debug.WriteLine("正在初始化GlobalStatusStrip...")
            GlobalStatusStripAll.InitializeApplication(Me.Application)
            Debug.WriteLine("GlobalStatusStrip初始化完成")
        Catch ex As Exception
            Debug.WriteLine("初始化GlobalStatusStrip时出错: " & ex.Message)
            MessageBox.Show("初始化状态栏时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Try
            WebView2Loader.EnsureWebView2Loader()
        Catch ex As Exception
            MessageBox.Show($"WebView2 初始化失败: {ex.Message}")
        End Try

        ' 初始化 Timer，用于WPS中扩大聊天区域的宽度
        widthTimer = New Timer()
        AddHandler widthTimer.Tick, AddressOf WidthTimer_Tick
        widthTimer.Interval = 100 ' 设置延迟时间，单位为毫秒

        widthTimer1 = New Timer()
        AddHandler widthTimer1.Tick, AddressOf WidthTimer1_Tick
        widthTimer1.Interval = 200 ' 设置延迟时间，单位为毫秒

        ' 等待Excel完全启动
        WaitForExcelReady()

        ' 创建任务窗格
        'CreateChatTaskPane()
        'CreateDeepseekTaskPane()

        ' 确保有可用的工作簿
        'EnsureWorkbookAvailable()

        ' 加载 Excel-DNA XLL (添加这部分)
        Try
            LoadExcelDnaAddIn()
        Catch ex As Exception
            Debug.WriteLine($"加载 Excel-DNA 失败: {ex.Message}")
            ' 继续执行，不要因为这个错误而中断启动
        End Try
        translateService = New ExcelTranslateService()

    End Sub

    ' 添加这个方法来等待Excel启动
    Private Sub WaitForExcelReady()
        Try
            Dim startTime As DateTime = DateTime.Now
            While Not Application.Ready AndAlso DateTime.Now.Subtract(startTime).TotalSeconds < 10
                System.Threading.Thread.Sleep(100)
            End While
            Debug.WriteLine("Excel准备就绪")
        Catch ex As Exception
            Debug.WriteLine($"等待Excel准备就绪时出错: {ex.Message}")
        End Try
    End Sub

    ' 确保有可用的工作簿 - 增强版
    Private Sub EnsureWorkbookAvailable()
        Try
            ' 记录所有当前工作簿
            Debug.WriteLine($"当前工作簿数量: {Application.Workbooks.Count}")
            If Application.Workbooks.Count > 0 Then
                Debug.WriteLine("工作簿列表:")
                For i As Integer = 1 To Application.Workbooks.Count
                    Dim wb As Workbook = Application.Workbooks(i)
                    Debug.WriteLine($"  [{i}] 名称: {wb.Name}, 路径: {If(String.IsNullOrEmpty(wb.Path), "(未保存)", wb.Path)}")
                Next
            End If

            ' 关键修改: 只有在没有任何实际工作簿时才创建
            ' 检查是否有任何非临时工作簿
            Dim hasRealWorkbook As Boolean = False

            If Application.Workbooks.Count > 0 Then
                ' 检查是否所有工作簿都是新建的空白工作簿
                For i As Integer = 1 To Application.Workbooks.Count
                    Dim wb As Workbook = Application.Workbooks(i)
                    ' 如果工作簿不是默认名称(如Book1)或已保存过，则视为有效工作簿
                    If Not (wb.Name.StartsWith("Book") OrElse wb.Name.StartsWith("工作簿")) OrElse
                   Not String.IsNullOrEmpty(wb.Path) OrElse
                   wb.Saved = False Then
                        hasRealWorkbook = True
                        Debug.WriteLine($"找到有效工作簿: {wb.Name}")
                        Exit For
                    End If
                Next
            End If

            ' 仅当没有工作簿或只有临时工作簿且数量=1时才创建新的
            If Application.Workbooks.Count = 0 OrElse (Application.Workbooks.Count = 1 AndAlso Not hasRealWorkbook) Then
                Debug.WriteLine("需要创建新工作簿...")

                ' 如果已经有一个临时工作簿，先关闭它
                If Application.Workbooks.Count = 1 AndAlso Not hasRealWorkbook Then
                    Debug.WriteLine("关闭现有临时工作簿")
                    ' 不保存关闭
                    Application.DisplayAlerts = False
                    Application.Workbooks(1).Close(SaveChanges:=False)
                    Application.DisplayAlerts = True
                End If

                Application.Workbooks.Add()
                Debug.WriteLine("已创建新工作簿")
            Else
                Debug.WriteLine("已存在有效工作簿，无需创建")
            End If

            ' 确保有活动工作簿
            If Application.ActiveWorkbook Is Nothing AndAlso Application.Workbooks.Count > 0 Then
                Application.Workbooks(1).Activate()
                Debug.WriteLine("已激活第一个工作簿")
            End If

            ' 确保Excel是可见的
            If Not Application.Visible Then
                Application.Visible = True
                Debug.WriteLine("已设置Excel为可见")
            End If

            Debug.WriteLine("工作簿状态检查完成")

        Catch ex As Exception
            Debug.WriteLine($"确保工作簿可用时出错: {ex.Message}")
            MessageBox.Show($"初始化Excel工作簿时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub LoadExcelDnaAddIn()
        Try
            Debug.WriteLine("开始查找 XLL 文件...")

            ' 获取当前程序集路径
            Dim currentAssemblyPath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
            Dim currentDir As String = Path.GetDirectoryName(currentAssemblyPath)

            Debug.WriteLine($"当前程序集路径: {currentAssemblyPath}")
            Debug.WriteLine($"当前目录: {currentDir}")

            ' 创建搜索路径列表
            Dim searchPaths As New List(Of String)

            ' 1. 开发环境：尝试找到真正的项目输出目录
            If currentDir.Contains("AppData\Local\assembly") Then
                ' 这是VSTO临时目录，尝试找到真正的项目目录
                ' 根据您的项目结构：F:\ai\code\AiHelper\ExcelAi\bin\Debug
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

            ' 2. 当前目录及其相关目录
            searchPaths.Add(currentDir)
            searchPaths.Add(Path.Combine(currentDir, ".."))

            ' 3. 安装环境：查找所有可能的安装路径
            ' 首先尝试从当前路径向上查找OfficeAiAgent目录
            Dim currentDirInfo As New DirectoryInfo(currentDir)
            While currentDirInfo IsNot Nothing
                ' 检查是否包含OfficeAiAgent
                If currentDirInfo.Name.Equals("OfficeAiAgent", StringComparison.OrdinalIgnoreCase) OrElse
               currentDirInfo.FullName.Contains("OfficeAiAgent") Then

                    ' 找到OfficeAiAgent目录，检查ExcelAi子目录
                    Dim excelAiPath As String = Path.Combine(currentDirInfo.FullName, "ExcelAi")
                    If Directory.Exists(excelAiPath) Then
                        searchPaths.Add(excelAiPath)
                        Debug.WriteLine($"添加安装路径: {excelAiPath}")
                    End If

                    ' 也检查根目录
                    searchPaths.Add(currentDirInfo.FullName)
                    Exit While
                End If
                currentDirInfo = currentDirInfo.Parent
            End While

            ' 4. 标准安装路径
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

            ' 5. 用户自定义安装路径：搜索所有驱动器
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

            ' 6. 从注册表查找安装路径（如果MSI安装时写入了注册表）
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

            ' 去除重复路径
            Dim uniquePaths As New HashSet(Of String)(searchPaths, StringComparer.OrdinalIgnoreCase)

            ' 查找XLL文件
            Dim xllFileName As String = If(IntPtr.Size = 8, "ExcelAi-AddIn64-packed.xll", "ExcelAi-AddIn-packed.xll")
            Dim foundXllPath As String = String.Empty

            Debug.WriteLine($"正在查找文件: {xllFileName}")
            Debug.WriteLine("搜索路径列表:")

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

            ' 尝试加载XLL文件
            If Not String.IsNullOrEmpty(foundXllPath) Then
                Try
                    Debug.WriteLine($"正在加载 Excel-DNA XLL: {foundXllPath}")
                    Dim result As Boolean = Application.RegisterXLL(foundXllPath)
                Catch ex As Exception
                    Debug.WriteLine($"加载 XLL 时出错: {ex.Message}")
                End Try
            Else
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

    ' 显示VBA信任中心设置说明
    Private Sub ShowVbaTrustCenterInstructions()
        MessageBox.Show(
        "要使用 Excel AI 函数，请按以下步骤设置Excel安全选项:" & vbCrLf & vbCrLf &
        "1. 点击'文件' > '选项'" & vbCrLf &
        "2. 选择'信任中心' > '信任中心设置'" & vbCrLf &
        "3. 选择'宏设置'" & vbCrLf &
        "4. 勾选'信任访问 VBA 项目对象模型'" & vbCrLf &
        "5. 点击'确定'并重启Excel",
        "Excel AI 函数 - 安全设置",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information)
    End Sub


    Private Function IsWpsActive() As Boolean
        Try
            Return Process.GetProcessesByName("WPS").Length > 0
        Catch
            Return False
        End Try
    End Function


    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
    End Sub


    Private widthTimer As Timer
    Private widthTimer1 As Timer

    ' 解决WPS中无法显示正常宽度的问题
    Private Sub ChatTaskPane_VisibleChanged(sender As Object, e As EventArgs)
        Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = CType(sender, Microsoft.Office.Tools.CustomTaskPane)
        If taskPane.Visible Then
            If IsWpsActive() Then
                widthTimer.Start()
            End If
        End If
    End Sub

    Private Sub DeepseekTaskPane_VisibleChanged(sender As Object, e As EventArgs)
        Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = CType(sender, Microsoft.Office.Tools.CustomTaskPane)
        If taskPane.Visible Then
            If IsWpsActive() Then
                widthTimer1.Start()
            End If
        End If
    End Sub

    Private Sub WidthTimer_Tick(sender As Object, e As EventArgs)
        widthTimer.Stop()
        If IsWpsActive() AndAlso chatTaskPane IsNot Nothing Then
            chatTaskPane.Width = 420
        End If
    End Sub

    Private Sub WidthTimer1_Tick(sender As Object, e As EventArgs)
        widthTimer1.Stop()
        Debug.WriteLine($"Deepseek点击定时1")
        If IsWpsActive() AndAlso _deepseekTaskPane IsNot Nothing Then
            Debug.WriteLine($"Deepseek点击定时2")
            _deepseekTaskPane.Width = 420
        End If
    End Sub

    Private Sub AiHelper_Shutdown() Handles Me.Shutdown
    End Sub

    Dim loadChatHtml As Boolean = True

    Public Async Sub ShowChatTaskPane()
        CreateChatTaskPane()
        chatTaskPane.Visible = True
        If loadChatHtml Then
            loadChatHtml = False
            Await chatControl.LoadLocalHtmlFile()
        End If
    End Sub

    Public Async Sub ShowDeepseekTaskPane()
        Debug.WriteLine($"Deepseek点击事件")
        CreateDeepseekTaskPane()
        _deepseekTaskPane.Visible = True
    End Sub

    Public Async Sub ShowDoubaoTaskPane()
        CreateDoubaoTaskPane()
        _doubaoTaskPane.Visible = True
    End Sub
End Class
