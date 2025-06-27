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

'<ComVisible(True)>
'<ClassInterface(ClassInterfaceType.AutoDual)>
Public Class ThisAddIn

    Private chatTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Public Shared chatControl As ChatControl



    Private Sub ExcelAi_Startup() Handles Me.Startup
        ' ��ʼ�� GlobalStatusStrip
        Try
            Debug.WriteLine("���ڳ�ʼ��GlobalStatusStrip...")
            GlobalStatusStripAll.InitializeApplication(Me.Application)
            Debug.WriteLine("GlobalStatusStrip��ʼ�����")

            ' ����״̬���Ƿ���������
            'GlobalStatusStripAll.ShowWarning("Excel������������")
        Catch ex As Exception
            Debug.WriteLine("��ʼ��GlobalStatusStripʱ����: " & ex.Message)
            MessageBox.Show("��ʼ��״̬��ʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Try
            WebView2Loader.EnsureWebView2Loader()
        Catch ex As Exception
            MessageBox.Show($"WebView2 ��ʼ��ʧ��: {ex.Message}")
        End Try

        ' ��������͹������л��¼�
        'AddHandler Globals.ThisAddIn.Application.ActiveDocument, AddressOf Me.Application_WorkbookActivate
        'Application_WorkbookActivate()
        ' ��ʼ�� Timer������WPS��������������Ŀ��
        widthTimer = New Timer()
        AddHandler widthTimer.Tick, AddressOf WidthTimer_Tick
        widthTimer.Interval = 100 ' �����ӳ�ʱ�䣬��λΪ����


        ' �ȴ�Excel��ȫ����
        WaitForExcelReady()

        ' �������񴰸�
        CreateChatTaskPane()

        ' ȷ���п��õĹ�����
        EnsureWorkbookAvailable()

        ' ���� Excel-DNA XLL (����ⲿ��)
        Try
            LoadExcelDnaAddIn()
        Catch ex As Exception
            Debug.WriteLine($"���� Excel-DNA ʧ��: {ex.Message}")
            ' ����ִ�У���Ҫ��Ϊ���������ж�����
        End Try
    End Sub

    ' �������������ȴ�Excel����
    Private Sub WaitForExcelReady()
        Try
            Dim startTime As DateTime = DateTime.Now
            While Not Application.Ready AndAlso DateTime.Now.Subtract(startTime).TotalSeconds < 10
                System.Threading.Thread.Sleep(100)
            End While
            Debug.WriteLine("Excel׼������")
        Catch ex As Exception
            Debug.WriteLine($"�ȴ�Excel׼������ʱ����: {ex.Message}")
        End Try
    End Sub

    ' ȷ���п��õĹ�����
    Private Sub EnsureWorkbookAvailable()
        Try
            Debug.WriteLine($"��ǰ����������: {Application.Workbooks.Count}")

            ' ���û�й�����������һ���µ�
            If Application.Workbooks.Count = 0 Then
                Debug.WriteLine("û�й����������ڴ����¹�����...")
                Application.Workbooks.Add()
                Debug.WriteLine("�Ѵ����¹�����")
            End If

            ' ȷ���л������
            If Application.ActiveWorkbook Is Nothing AndAlso Application.Workbooks.Count > 0 Then
                Application.Workbooks(1).Activate()
                Debug.WriteLine("�Ѽ����һ��������")
            End If

            ' ȷ��Excel�ǿɼ���
            If Not Application.Visible Then
                Application.Visible = True
                Debug.WriteLine("������ExcelΪ�ɼ�")
            End If

            Debug.WriteLine("������״̬������")

        Catch ex As Exception
            Debug.WriteLine($"ȷ������������ʱ����: {ex.Message}")
            MessageBox.Show($"��ʼ��Excel������ʱ����: {ex.Message}", "����", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub Application_WorkbookActivate(Optional Wb As Workbook = Nothing)
        Try
            Debug.WriteLine("�����������¼�����")

            ' ���û�й����������Դ���һ��
            If Application.Workbooks.Count = 0 Then
                Debug.WriteLine("�����¼���û�й����������ڴ���...")
                Application.Workbooks.Add()
            End If

            ' ȷ�����񴰸���ȷ��ʼ��
            If chatTaskPane Is Nothing Then
                CreateChatTaskPane()
            End If

        Catch ex As Exception
            Debug.WriteLine($"����������ʱ����: {ex.Message}")
        End Try
    End Sub
    Private Sub LoadExcelDnaAddIn()
        Try
            Debug.WriteLine("��ʼ���� XLL �ļ�...")

            ' ��ȡ��ǰ����·��
            Dim currentAssemblyPath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
            Dim currentDir As String = Path.GetDirectoryName(currentAssemblyPath)

            Debug.WriteLine($"��ǰ����·��: {currentAssemblyPath}")
            Debug.WriteLine($"��ǰĿ¼: {currentDir}")

            ' ��������·���б�
            Dim searchPaths As New List(Of String)

            ' 1. ���������������ҵ���������Ŀ���Ŀ¼
            If currentDir.Contains("AppData\Local\assembly") Then
                ' ����VSTO��ʱĿ¼�������ҵ���������ĿĿ¼
                ' ����������Ŀ�ṹ��F:\ai\code\AiHelper\ExcelAi\bin\Debug
                Dim possibleDevPaths As String() = {
                "F:\ai\code\AiHelper\ExcelAi\bin\Debug",
                "F:\ai\code\AiHelper\ExcelAi\bin\Release",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "source", "repos", "AiHelper", "ExcelAi", "bin", "Debug"),
                Path.Combine("C:\", "ai", "code", "AiHelper", "ExcelAi", "bin", "Debug")
            }

                For Each devPath In possibleDevPaths
                    If Directory.Exists(devPath) Then
                        searchPaths.Add(devPath)
                        Debug.WriteLine($"��ӿ���·��: {devPath}")
                    End If
                Next
            End If

            ' 2. ��ǰĿ¼�������Ŀ¼
            searchPaths.Add(currentDir)
            searchPaths.Add(Path.Combine(currentDir, ".."))

            ' 3. ��װ�������������п��ܵİ�װ·��
            ' ���ȳ��Դӵ�ǰ·�����ϲ���OfficeAiAgentĿ¼
            Dim currentDirInfo As New DirectoryInfo(currentDir)
            While currentDirInfo IsNot Nothing
                ' ����Ƿ����OfficeAiAgent
                If currentDirInfo.Name.Equals("OfficeAiAgent", StringComparison.OrdinalIgnoreCase) OrElse
               currentDirInfo.FullName.Contains("OfficeAiAgent") Then

                    ' �ҵ�OfficeAiAgentĿ¼�����ExcelAi��Ŀ¼
                    Dim excelAiPath As String = Path.Combine(currentDirInfo.FullName, "ExcelAi")
                    If Directory.Exists(excelAiPath) Then
                        searchPaths.Add(excelAiPath)
                        Debug.WriteLine($"��Ӱ�װ·��: {excelAiPath}")
                    End If

                    ' Ҳ����Ŀ¼
                    searchPaths.Add(currentDirInfo.FullName)
                    Exit While
                End If
                currentDirInfo = currentDirInfo.Parent
            End While

            ' 4. ��׼��װ·��
            Dim standardInstallPaths As String() = {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "it235", "OfficeAiAgent", "ExcelAi"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "it235", "OfficeAiAgent", "ExcelAi"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "OfficeAiAgent", "ExcelAi"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "OfficeAiAgent", "ExcelAi")
        }

            For Each installPath In standardInstallPaths
                If Directory.Exists(installPath) Then
                    searchPaths.Add(installPath)
                    Debug.WriteLine($"��ӱ�׼��װ·��: {installPath}")
                End If
            Next

            ' 5. �û��Զ��尲װ·������������������
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
                                Debug.WriteLine($"����Զ���·��: {customPath}")
                            End If
                        Next
                    End If
                Next
            Catch ex As Exception
                Debug.WriteLine($"�����Զ���·��ʱ����: {ex.Message}")
            End Try

            ' 6. ��ע�����Ұ�װ·�������MSI��װʱд����ע���
            Try
                Using key As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\it235\OfficeAiAgent")
                    If key IsNot Nothing Then
                        Dim installPath As String = key.GetValue("InstallPath")?.ToString()
                        If Not String.IsNullOrEmpty(installPath) Then
                            Dim excelAiPath As String = Path.Combine(installPath, "ExcelAi")
                            If Directory.Exists(excelAiPath) Then
                                searchPaths.Add(excelAiPath)
                                Debug.WriteLine($"��ע������·��: {excelAiPath}")
                            End If
                        End If
                    End If
                End Using
            Catch ex As Exception
                Debug.WriteLine($"��ע����ȡ��װ·��ʱ����: {ex.Message}")
            End Try

            ' ȥ���ظ�·��
            Dim uniquePaths As New HashSet(Of String)(searchPaths, StringComparer.OrdinalIgnoreCase)

            ' ����XLL�ļ�
            Dim xllFileName As String = If(IntPtr.Size = 8, "ExcelAi-AddIn64-packed.xll", "ExcelAi-AddIn-packed.xll")
            Dim foundXllPath As String = String.Empty

            Debug.WriteLine($"���ڲ����ļ�: {xllFileName}")
            Debug.WriteLine("����·���б�:")

            For Each searchPath In uniquePaths
                Debug.WriteLine($"  ���: {searchPath}")

                If Directory.Exists(searchPath) Then
                    Dim xllPath As String = Path.Combine(searchPath, xllFileName)
                    If File.Exists(xllPath) Then
                        foundXllPath = xllPath
                        Debug.WriteLine($"�ҵ�XLL�ļ�: {xllPath}")
                        Exit For
                    End If

                    ' Ҳ���δ����İ汾
                    Dim unpackedXllFileName As String = If(IntPtr.Size = 8, "ExcelAi-AddIn64.xll", "ExcelAi-AddIn.xll")
                    Dim unpackedXllPath As String = Path.Combine(searchPath, unpackedXllFileName)
                    If File.Exists(unpackedXllPath) Then
                        foundXllPath = unpackedXllPath
                        Debug.WriteLine($"�ҵ�δ���XLL�ļ�: {unpackedXllPath}")
                        Exit For
                    End If
                End If
            Next

            ' ���Լ���XLL�ļ�
            If Not String.IsNullOrEmpty(foundXllPath) Then
                Try
                    Debug.WriteLine($"���ڼ��� Excel-DNA XLL: {foundXllPath}")
                    Dim result As Boolean = Application.RegisterXLL(foundXllPath)

                    If result Then
                        Debug.WriteLine($"�ɹ����� Excel-DNA XLL: {foundXllPath}")
                        'GlobalStatusStripAll.ShowWarning($"Excel DNA �����Ѽ��أ�����ʹ�� =DLLM() �� =ALLM() ����")
                    Else
                        Debug.WriteLine($"RegisterXLL����False: {foundXllPath}")
                        'GlobalStatusStripAll.ShowWarning($"Excel DNA ע��ʧ�ܣ�RegisterXLL����False")
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"���� XLL ʱ����: {ex.Message}")
                    'GlobalStatusStripAll.ShowWarning($"Excel DNA ���ش���: {ex.Message}")
                End Try
            Else
                Debug.WriteLine("δ�ҵ��κ�XLL�ļ�")

                ' ���������Ϣ����ʾÿ������·��������
                Debug.WriteLine("����·����ϸ��Ϣ:")
                For Each searchPath In uniquePaths
                    If Directory.Exists(searchPath) Then
                        Debug.WriteLine($"·�� {searchPath} �������ļ�:")
                        Try
                            For Each file In Directory.GetFiles(searchPath, "*.xll")
                                Debug.WriteLine($"  XLL�ļ�: {Path.GetFileName(file)}")
                            Next
                            For Each file In Directory.GetFiles(searchPath, "ExcelAi*.*")
                                Debug.WriteLine($"  ExcelAi�ļ�: {Path.GetFileName(file)}")
                            Next
                        Catch ex As Exception
                            Debug.WriteLine($"  �޷���ȡĿ¼����: {ex.Message}")
                        End Try
                    Else
                        Debug.WriteLine($"·��������: {searchPath}")
                    End If
                Next

                GlobalStatusStripAll.ShowWarning("δ�ҵ� Excel DNA �ļ���DLLM �� ALLM ����������")
            End If

        Catch ex As Exception
            Debug.WriteLine($"LoadExcelDnaAddIn ����: {ex.Message}")
            Debug.WriteLine($"��ջ����: {ex.StackTrace}")
            GlobalStatusStripAll.ShowWarning($"Excel DNA ���س���: {ex.Message}")
        End Try
    End Sub

    ' �����������񴰸�
    Private Sub CreateChatTaskPane()
        Try
            ' Ϊ�¹������������񴰸�
            chatControl = New ChatControl()
            chatTaskPane = Me.CustomTaskPanes.Add(chatControl, "Excel AI��������")
            chatTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            chatTaskPane.Width = 420
            AddHandler chatTaskPane.VisibleChanged, AddressOf ChatTaskPane_VisibleChanged
            chatTaskPane.Visible = False
        Catch ex As Exception
            MessageBox.Show($"��ʼ�����񴰸�ʧ��: {ex.Message}")
        End Try
    End Sub


    ' ��ʾVBA������������˵��
    Private Sub ShowVbaTrustCenterInstructions()
        MessageBox.Show(
        "Ҫʹ�� Excel AI �������밴���²�������Excel��ȫѡ��:" & vbCrLf & vbCrLf &
        "1. ���'�ļ�' > 'ѡ��'" & vbCrLf &
        "2. ѡ��'��������' > '������������'" & vbCrLf &
        "3. ѡ��'������'" & vbCrLf &
        "4. ��ѡ'���η��� VBA ��Ŀ����ģ��'" & vbCrLf &
        "5. ���'ȷ��'������Excel",
        "Excel AI ���� - ��ȫ����",
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
    ' ���WPS���޷���ʾ������ȵ�����
    Private Sub ChatTaskPane_VisibleChanged(sender As Object, e As EventArgs)
        Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = CType(sender, Microsoft.Office.Tools.CustomTaskPane)
        If taskPane.Visible Then
            If IsWpsActive() Then
                widthTimer.Start()
            End If
        End If
    End Sub

    Private Sub WidthTimer_Tick(sender As Object, e As EventArgs)
        widthTimer.Stop()
        If IsWpsActive() AndAlso chatTaskPane IsNot Nothing Then
            chatTaskPane.Width = 420
        End If
    End Sub
    Private Sub AiHelper_Shutdown() Handles Me.Shutdown
        ' ������Դ
        'RemoveHandler Globals.ThisAddIn.Application.WorkbookActivate, AddressOf Me.Application_WorkbookActivate
    End Sub

    Dim loadChatHtml As Boolean = True

    Public Async Sub ShowChatTaskPane()
        chatTaskPane.Visible = True
        If loadChatHtml Then
            loadChatHtml = False
            Await chatControl.LoadLocalHtmlFile()
        End If
    End Sub
End Class
