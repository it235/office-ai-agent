Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Reflection.Emit
Imports System.Text
Imports System.Text.JSON
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Windows.Forms
Imports System.Windows.Forms.ListBox
Imports Markdig
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon
Public Class ChatControl
    Inherits BaseChatControl


    Private sheetContentItems As New Dictionary(Of String, Tuple(Of System.Windows.Forms.Label, System.Windows.Forms.Button))


    Public Sub New()
        ' �˵��������ʦ������ġ�
        InitializeComponent()

        ' ȷ��WebView2�ؼ�������������
        ChatBrowser.BringToFront()

        '����ײ��澯��
        Me.Controls.Add(GlobalStatusStrip.StatusStrip)

        ' ����Word��SelectionChange �¼�
        ' ���Ҳ�ȫwordѡ��������¼�
        AddHandler Globals.ThisAddIn.Application.WindowSelectionChange, AddressOf GetSelectionContent
    End Sub

    '��ȡѡ�е�����
    Protected Overrides Sub GetSelectionContent(target As Object)
        Try
            If Not Me.Visible OrElse Not selectedCellChecked Then
                Return
            End If

            ' ת��Ϊ PowerPoint.Selection ����
            Dim selection = Globals.ThisAddIn.Application.ActiveWindow.Selection
            If selection Is Nothing Then
                Return
            End If

            ' ��ȡѡ�����ݵ���ϸ��Ϣ
            Dim content As String = String.Empty

            ' ����ѡ�����ʹ�������
            If selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes Then
                ' ������״ѡ��
                Dim shapeRange = selection.ShapeRange
                If shapeRange.Count > 0 Then
                    ' ����Ƿ��Ǳ��
                    If shapeRange(1).HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        ' ������
                        Dim table = shapeRange(1).Table
                        Dim sb As New StringBuilder()
                        For row As Integer = 1 To table.Rows.Count
                            For col As Integer = 1 To table.Columns.Count
                                sb.Append(table.Cell(row, col).Shape.TextFrame.TextRange.Text.Trim())
                                If col < table.Columns.Count Then sb.Append(vbTab)
                            Next
                            sb.AppendLine()
                        Next
                        content = sb.ToString()
                    Else
                        ' ������ͨ��״
                        content = "[��ѡ�� " & shapeRange.Count & " ����״]"
                        For i = 1 To shapeRange.Count
                            If shapeRange(i).HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                content &= vbCrLf & shapeRange(i).TextFrame.TextRange.Text
                            End If
                        Next
                    End If
                End If

            ElseIf selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText Then
                ' �����ı�ѡ��
                content = selection.TextRange.Text

            ElseIf selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides Then
                ' ����õ�Ƭѡ��
                content = "[��ѡ�� " & selection.SlideRange.Count & " �Żõ�Ƭ]"
            End If

            If Not String.IsNullOrEmpty(content) Then
                ' ��ӵ�ѡ�������б�
                AddSelectedContentItem(
                "PowerPoint�õ�Ƭ",  ' ʹ���ĵ�������Ϊ��ʶ
                content.Substring(0, Math.Min(content.Length, 50)) & If(content.Length > 50, "...", "")
            )
            End If

        Catch ex As Exception
            Debug.WriteLine($"��ȡPowerPointѡ������ʱ����: {ex.Message}")
        End Try
    End Sub

    Private Function GetSelectionDetails(selection As Object) As String
        Try
            Dim details As New StringBuilder()
            Dim ppSelection = TryCast(selection, Microsoft.Office.Interop.PowerPoint.Selection)

            If ppSelection Is Nothing Then
                Return "δѡ���κ�����"
            End If

            ' ��ӻ�����Ϣ
            details.AppendLine($"ѡ������: {ppSelection.Type}")

            If ppSelection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes Then
                Dim shapeRange = ppSelection.ShapeRange
                details.AppendLine($"��״����: {shapeRange.Count}")
                For i = 1 To shapeRange.Count
                    details.AppendLine($"��״ {i} ����: {shapeRange(i).Type}")
                    ' ����Ƿ��Ǳ��
                    If shapeRange(i).HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        Dim table = shapeRange(i).Table
                        details.AppendLine($"����С: {table.Rows.Count}�� x {table.Columns.Count}��")
                    ElseIf shapeRange(i).HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        details.AppendLine($"��״ {i} �ı�����: {shapeRange(i).TextFrame.TextRange.Length}")
                    End If
                Next

            ElseIf ppSelection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText Then
                Dim textRange = ppSelection.TextRange
                details.AppendLine($"�ı�����: {textRange.Length}")
                details.AppendLine($"�ַ���: {textRange.Length}")

            ElseIf ppSelection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides Then
                Dim slideRange = ppSelection.SlideRange
                details.AppendLine($"ѡ�лõ�Ƭ��: {slideRange.Count}")
                For i = 1 To slideRange.Count
                    details.AppendLine($"�õ�Ƭ {i} ����: {slideRange(i).Name}")
                Next
            End If

            Return details.ToString()
        Catch ex As Exception
            Return $"��ȡѡ������ʱ����: {ex.Message}"
        End Try
    End Function

    ' ��ʼ��ʱע����� HTML �ṹ
    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' ��ʼ�� WebView2
        Await InitializeWebView2()
        InitializeWebView2Script()
    End Sub


    Protected Overrides Function GetVBProject() As VBProject
        Try
            Dim project = Globals.ThisAddIn.Application.VBE.ActiveVBProject
            Return project
        Catch ex As Runtime.InteropServices.COMException
            VBAxceptionHandle(ex)
            Return Nothing
        End Try
    End Function

    'Protected Overrides Function RunCode(code As String) As Object
    '    Try
    '        Globals.ThisAddIn.Application.Run(code)
    '        Return True
    '    Catch ex As Runtime.InteropServices.COMException
    '        VBAxceptionHandle(ex)
    '        Return False
    '    Catch ex As Exception
    '        MessageBox.Show("ִ�д���ʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Return False
    '    End Try
    'End Function

    ' ִ��ǰ�˴����� VBA ����Ƭ��
    Protected Overrides Function RunCode(vbaCode As String) As Object
        ' ��ȡ VBA ��Ŀ
        Dim vbProj As VBProject = GetVBProject()

        ' ��ӿ�ֵ���
        If vbProj Is Nothing Then
            Return Nothing
        End If

        Dim vbComp As VBComponent = Nothing
        Dim tempModuleName As String = "TempMod" & DateTime.Now.Ticks.ToString().Substring(0, 8)

        Try
            ' ������ʱģ��
            vbComp = vbProj.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
            vbComp.Name = tempModuleName

            ' �������Ƿ��Ѱ��� Sub/Function ����
            If ContainsProcedureDeclaration(vbaCode) Then
                ' �����Ѱ�������������ֱ�����
                vbComp.CodeModule.AddFromString(vbaCode)

                ' ���ҵ�һ����������ִ��
                Dim procName As String = FindFirstProcedureName(vbComp)
                If Not String.IsNullOrEmpty(procName) Then
                    Globals.ThisAddIn.Application.Run(tempModuleName & "." & procName, vbaCode)
                Else
                    'MessageBox.Show("�޷��ڴ������ҵ���ִ�еĹ���")
                    GlobalStatusStrip.ShowWarning("�޷��ڴ������ҵ���ִ�еĹ���")
                End If
            Else
                ' ���벻�������������������װ�� Auto_Run ������
                Dim wrappedCode As String = "Sub Auto_Run()" & vbNewLine &
                                           vbaCode & vbNewLine &
                                           "End Sub"
                vbComp.CodeModule.AddFromString(wrappedCode)

                ' ִ�� Auto_Run ����
                Globals.ThisAddIn.Application.Run(tempModuleName & ".Auto_Run", vbaCode)
            End If

        Catch ex As Exception
            MessageBox.Show("ִ�� VBA ����ʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' ���۳ɹ�����ʧ�ܣ���ɾ����ʱģ��
            Try
                If vbProj IsNot Nothing AndAlso vbComp IsNot Nothing Then
                    vbProj.VBComponents.Remove(vbComp)
                End If
            Catch
                ' �����������
            End Try
        End Try
        Return True
    End Function

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("PowerPoint", OfficeApplicationType.PowerPoint)
    End Function

    Protected Overrides Sub SendChatMessage(message As String)
        ' �������ʵ��word�������߼�
        Send(message)
    End Sub

    Protected Overrides Function ParseFile(filePath As String) As FileContentResult

    End Function
    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Try
            ' ����Ƿ�������ѡ����
            If Not selectedCellChecked Then
                Return message
            End If

            ' ��ȡ��ǰ PowerPoint �е�ѡ��
            Dim selection = Globals.ThisAddIn.Application.ActiveWindow.Selection
            If selection Is Nothing Then
                Return message
            End If

            ' �������ݹ���������ʽ��ѡ������
            Dim contentBuilder As New StringBuilder()
            contentBuilder.AppendLine(vbCrLf & "--- �û�ѡ�е� PowerPoint ���� ---")

            ' �����ʾ�ĸ���Ϣ
            Dim activePresentation = Globals.ThisAddIn.Application.ActivePresentation
            If activePresentation IsNot Nothing Then
                contentBuilder.AppendLine($"��ʾ�ĸ�: {Path.GetFileName(activePresentation.FullName)}")
                contentBuilder.AppendLine($"��ǰ�õ�Ƭ: {Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex}")
            End If

            ' ����ѡ�����ʹ�������
            Select Case selection.Type
                Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes
                    ' ������״ѡ�񣨰������
                    Dim shapeRange = selection.ShapeRange
                    contentBuilder.AppendLine($"ѡ������: ��״ (�� {shapeRange.Count} ��)")

                    For i = 1 To shapeRange.Count
                        contentBuilder.AppendLine($"��״ {i}:")

                        ' ����Ƿ��Ǳ��
                        If shapeRange(i).HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            Dim table = shapeRange(i).Table
                            contentBuilder.AppendLine($"  ���: {table.Rows.Count} �� �� {table.Columns.Count} ��")

                            ' ��ӱ������
                            Dim maxRows As Integer = Math.Min(table.Rows.Count, 20)
                            Dim maxCols As Integer = Math.Min(table.Columns.Count, 10)

                            ' ������ͷ��
                            Dim headerBuilder As New StringBuilder("  ")
                            Dim separatorBuilder As New StringBuilder("  ")

                            For col = 1 To maxCols
                                Try
                                    Dim cellText = table.Cell(1, col).Shape.TextFrame.TextRange.Text.Trim()
                                    ' ���Ƶ�Ԫ���ı�����
                                    If cellText.Length > 20 Then
                                        cellText = cellText.Substring(0, 17) & "..."
                                    End If

                                    If col > 1 Then
                                        headerBuilder.Append(" | ")
                                        separatorBuilder.Append("-+-")
                                    End If
                                    headerBuilder.Append(cellText)
                                    separatorBuilder.Append(New String("-"c, Math.Max(cellText.Length, 3)))
                                Catch ex As Exception
                                    If col > 1 Then
                                        headerBuilder.Append(" | ")
                                        separatorBuilder.Append("-+-")
                                    End If
                                    headerBuilder.Append("N/A")
                                    separatorBuilder.Append("---")
                                End Try
                            Next

                            contentBuilder.AppendLine(headerBuilder.ToString())
                            contentBuilder.AppendLine(separatorBuilder.ToString())

                            ' ������������
                            For row = 2 To maxRows
                                Dim rowBuilder As New StringBuilder("  ")

                                For col = 1 To maxCols
                                    Try
                                        Dim cellText = table.Cell(row, col).Shape.TextFrame.TextRange.Text.Trim()
                                        ' ���Ƶ�Ԫ���ı�����
                                        If cellText.Length > 20 Then
                                            cellText = cellText.Substring(0, 17) & "..."
                                        End If

                                        If col > 1 Then
                                            rowBuilder.Append(" | ")
                                        End If
                                        rowBuilder.Append(cellText)
                                    Catch ex As Exception
                                        If col > 1 Then
                                            rowBuilder.Append(" | ")
                                        End If
                                        rowBuilder.Append("N/A")
                                    End Try
                                Next

                                contentBuilder.AppendLine(rowBuilder.ToString())
                            Next

                            ' ��ӱ��˵��
                            If table.Rows.Count > maxRows Then
                                contentBuilder.AppendLine($"  ... ���� {table.Rows.Count} �У�����ʾǰ {maxRows} ��")
                            End If

                            If table.Columns.Count > maxCols Then
                                contentBuilder.AppendLine($"  ... ���� {table.Columns.Count} �У�����ʾǰ {maxCols} ��")
                            End If
                        ElseIf shapeRange(i).HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            ' �����ı���
                            Dim textFrame = shapeRange(i).TextFrame
                            If textFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                Dim text = textFrame.TextRange.Text.Trim()
                                ' �����ı�����
                                If text.Length > 500 Then
                                    contentBuilder.AppendLine($"  �ı�: {text.Substring(0, 500)}...")
                                    contentBuilder.AppendLine($"  [�ı�̫��������ʾǰ500���ַ����ܼ�: {text.Length}���ַ�]")
                                Else
                                    contentBuilder.AppendLine($"  �ı�: {text}")
                                End If
                            Else
                                contentBuilder.AppendLine("  [���ı���]")
                            End If
                        ElseIf shapeRange(i).Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                            ' ����ͼƬ
                            contentBuilder.AppendLine("  [ͼƬ]")
                            If shapeRange(i).AlternativeText <> "" Then
                                contentBuilder.AppendLine($"  ����ı�: {shapeRange(i).AlternativeText}")
                            End If
                        Else
                            ' �������͵���״
                            contentBuilder.AppendLine($"  [��״����: {shapeRange(i).Type}]")
                        End If

                        ' ����״֮����ӷָ���
                        contentBuilder.AppendLine("  ---")
                    Next

                Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText
                    ' �����ı�ѡ��
                    contentBuilder.AppendLine("ѡ������: �ı�")

                    Dim textRange = selection.TextRange
                    If textRange IsNot Nothing Then
                        Dim text = textRange.Text.Trim()
                        ' �����ı�����
                        If text.Length > 1000 Then
                            contentBuilder.AppendLine(text.Substring(0, 1000) & "...")
                            contentBuilder.AppendLine($"[�ı�̫��������ʾǰ1000���ַ����ܼ�: {text.Length}���ַ�]")
                        Else
                            contentBuilder.AppendLine(text)
                        End If
                    Else
                        contentBuilder.AppendLine("[�޷���ȡ�ı�����]")
                    End If

                Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides
                    ' ����õ�Ƭѡ��
                    Dim slideRange = selection.SlideRange
                    contentBuilder.AppendLine($"ѡ������: �õ�Ƭ (�� {slideRange.Count} ��)")

                    ' ���ƴ���Ļõ�Ƭ����
                    Dim maxSlides = Math.Min(slideRange.Count, 5)

                    For i = 1 To maxSlides
                        Dim slide = slideRange(i)
                        contentBuilder.AppendLine($"�õ�Ƭ {slide.SlideIndex}:")

                        ' ��ȡ�õ�Ƭ����
                        Dim title As String = ""
                        For Each shape In slide.Shapes
                            If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                                If shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                                    If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                        title = shape.TextFrame.TextRange.Text.Trim()
                                        Exit For
                                    End If
                                End If
                            End If
                        Next

                        If title <> "" Then
                            contentBuilder.AppendLine($"  ����: {title}")
                        Else
                            contentBuilder.AppendLine("  [�ޱ���]")
                        End If

                        ' ��ȡ�õ�Ƭ�ϵ�����
                        Dim textShapesCount = 0

                        For Each shape In slide.Shapes
                            ' ����������״
                            If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder AndAlso
                           shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                                Continue For
                            End If

                            ' �����ı���״
                            If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
                           shape.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then

                                textShapesCount += 1
                                If textShapesCount > 3 Then Continue For ' ÿ�Żõ�Ƭ��ദ��3���ı���

                                Dim text = shape.TextFrame.TextRange.Text.Trim()
                                If text.Length > 0 Then
                                    ' �����ı�����
                                    If text.Length > 200 Then
                                        contentBuilder.AppendLine($"  �ı�: {text.Substring(0, 200)}...")
                                    Else
                                        contentBuilder.AppendLine($"  �ı�: {text}")
                                    End If
                                End If
                            ElseIf shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                contentBuilder.AppendLine("  [�������]")
                            ElseIf shape.Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                                contentBuilder.AppendLine("  [����ͼƬ]")
                            End If
                        Next

                        contentBuilder.AppendLine("  ---")
                    Next

                    ' ����и���õ�Ƭδ��ʾ�������ʾ
                    If slideRange.Count > maxSlides Then
                        contentBuilder.AppendLine($"[��ѡ�� {slideRange.Count} �Żõ�Ƭ������ʾǰ {maxSlides} ��]")
                    End If

                Case Else
                    contentBuilder.AppendLine($"ѡ������: δ֪ ({selection.Type})")
                    contentBuilder.AppendLine("[�޷�ʶ���ѡ������]")
            End Select

            contentBuilder.AppendLine("--- ѡ�����ݽ��� ---" & vbCrLf)

            ' ����ԭʼ��Ϣ����ѡ������
            Return message & contentBuilder.ToString()

        Catch ex As Exception
            Debug.WriteLine($"����PowerPointѡ������ʱ����: {ex.Message}")
            Return message ' ����ʱ����ԭʼ��Ϣ
        End Try
    End Function

    ' ������״ѡ�񣨰������
    Private Sub ProcessShapeSelection(builder As StringBuilder, selection As Microsoft.Office.Interop.PowerPoint.Selection)
        Try
            Dim shapeRange = selection.ShapeRange
            builder.AppendLine($"��״����: {shapeRange.Count}")

            ' ����ѡ�е���״
            For i = 1 To shapeRange.Count
                builder.AppendLine($"��״ {i}:")

                ' ����Ƿ��Ǳ��
                If shapeRange(i).HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    ProcessTable(builder, shapeRange(i).Table)
                ElseIf shapeRange(i).HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    ' ��������ı�����״
                    Dim textFrame = shapeRange(i).TextFrame
                    If textFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        Dim text = textFrame.TextRange.Text.Trim()
                        ' �����ı�����
                        If text.Length > 1000 Then
                            builder.AppendLine(text.Substring(0, 1000) & "...")
                            builder.AppendLine($"[�ı�̫��������ʾǰ1000���ַ����ܼ�: {text.Length}���ַ�]")
                        Else
                            builder.AppendLine(text)
                        End If
                    Else
                        builder.AppendLine("[���ı���]")
                    End If
                ElseIf shapeRange(i).Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                    ' ����ͼƬ
                    builder.AppendLine("[ͼƬ]")
                    ' ���Ի�ȡͼƬ������ı�������У�
                    If shapeRange(i).AlternativeText <> "" Then
                        builder.AppendLine($"����ı�: {shapeRange(i).AlternativeText}")
                    End If
                ElseIf shapeRange(i).Type = Microsoft.Office.Core.MsoShapeType.msoChart Then
                    ' ����ͼ��
                    builder.AppendLine("[ͼ��]")
                    If shapeRange(i).AlternativeText <> "" Then
                        builder.AppendLine($"ͼ��˵��: {shapeRange(i).AlternativeText}")
                    End If
                ElseIf shapeRange(i).Type = Microsoft.Office.Core.MsoShapeType.msoSmartArt Then
                    ' ����SmartArt
                    builder.AppendLine("[SmartArtͼ��]")
                Else
                    ' �������͵���״
                    builder.AppendLine($"[��״����: {shapeRange(i).Type}]")
                End If

                ' ��״֮����ӷָ���
                builder.AppendLine("---")
            Next

        Catch ex As Exception
            builder.AppendLine($"[������״ʱ����: {ex.Message}]")
        End Try
    End Sub

    ' ����������
    Private Sub ProcessTable(builder As StringBuilder, table As Microsoft.Office.Interop.PowerPoint.Table)
        Try
            builder.AppendLine($"���: {table.Rows.Count}�� �� {table.Columns.Count}��")

            ' ������ʾ��������
            Dim maxRows As Integer = Math.Min(table.Rows.Count, 20)
            Dim maxCols As Integer = Math.Min(table.Columns.Count, 10)

            ' ������ͷ��������һ�У�
            If table.Rows.Count > 0 Then
                ' ������ͷ�ͷָ���
                Dim headerBuilder As New StringBuilder()
                Dim separatorBuilder As New StringBuilder()

                For col As Integer = 1 To maxCols
                    Try
                        Dim cellText As String = table.Cell(1, col).Shape.TextFrame.TextRange.Text.Trim()

                        ' ���Ƶ�Ԫ���ı�����
                        If cellText.Length > 20 Then
                            cellText = cellText.Substring(0, 17) & "..."
                        End If

                        ' ����ͷ
                        If col > 1 Then
                            headerBuilder.Append(" | ")
                            separatorBuilder.Append("-+-")
                        End If
                        headerBuilder.Append(cellText)
                        separatorBuilder.Append(New String("-"c, Math.Max(cellText.Length, 3)))
                    Catch ex As Exception
                        ' ���Ե�Ԫ�������
                        If col > 1 Then
                            headerBuilder.Append(" | ")
                            separatorBuilder.Append("-+-")
                        End If
                        headerBuilder.Append("N/A")
                        separatorBuilder.Append("---")
                    End Try
                Next

                ' ��ӱ�ͷ�ͷָ���
                builder.AppendLine(headerBuilder.ToString())
                builder.AppendLine(separatorBuilder.ToString())
            End If

            ' ������������
            For row As Integer = 2 To maxRows ' �ӵ�2�п�ʼ��������ͷ��
                Dim rowBuilder As New StringBuilder()

                For col As Integer = 1 To maxCols
                    Try
                        Dim cellText As String = table.Cell(row, col).Shape.TextFrame.TextRange.Text.Trim()

                        ' ���Ƶ�Ԫ���ı�����
                        If cellText.Length > 20 Then
                            cellText = cellText.Substring(0, 17) & "..."
                        End If

                        ' ���������
                        If col > 1 Then
                            rowBuilder.Append(" | ")
                        End If
                        rowBuilder.Append(cellText)
                    Catch ex As Exception
                        ' ���Ե�Ԫ�������
                        If col > 1 Then
                            rowBuilder.Append(" | ")
                        End If
                        rowBuilder.Append("N/A")
                    End Try
                Next

                ' ���������
                builder.AppendLine(rowBuilder.ToString())
            Next

            ' ����и�����δ��ʾ�������ʾ
            If table.Rows.Count > maxRows Then
                builder.AppendLine($"... [����� {table.Rows.Count} �У�����ʾǰ {maxRows} ��]")
            End If

            ' ����и�����δ��ʾ�������ʾ
            If table.Columns.Count > maxCols Then
                builder.AppendLine($"... [����� {table.Columns.Count} �У�����ʾǰ {maxCols} ��]")
            End If

        Catch ex As Exception
            builder.AppendLine($"[����������ʱ����: {ex.Message}]")
        End Try
    End Sub

    ' �����ı�ѡ��
    Private Sub ProcessTextSelection(builder As StringBuilder, selection As Microsoft.Office.Interop.PowerPoint.Selection)
        Try
            Dim textRange = selection.TextRange

            If textRange IsNot Nothing Then
                builder.AppendLine($"�ı�����: {textRange.Length} ���ַ�")

                ' ��ȡ�ı����ݲ����Ƴ���
                Dim text = textRange.Text.Trim()
                Dim maxLength As Integer = 2000

                If text.Length > maxLength Then
                    builder.AppendLine(text.Substring(0, maxLength) & "...")
                    builder.AppendLine($"[�ı�̫��������ʾǰ{maxLength}���ַ����ܼ�: {text.Length}���ַ�]")
                Else
                    builder.AppendLine(text)
                End If
            Else
                builder.AppendLine("[�޷���ȡ�ı�����]")
            End If

        Catch ex As Exception
            builder.AppendLine($"[�����ı�ѡ��ʱ����: {ex.Message}]")
        End Try
    End Sub

    ' ����õ�Ƭѡ��
    Private Sub ProcessSlideSelection(builder As StringBuilder, selection As Microsoft.Office.Interop.PowerPoint.Selection)
        Try
            Dim slideRange = selection.SlideRange
            builder.AppendLine($"ѡ�лõ�Ƭ��: {slideRange.Count}")

            ' ���ƴ���Ļõ�Ƭ����
            Dim maxSlides As Integer = Math.Min(slideRange.Count, 10)

            For i = 1 To maxSlides
                Dim slide = slideRange(i)
                builder.AppendLine($"�õ�Ƭ {slide.SlideIndex}:")

                ' ��ȡ�õ�Ƭ����
                Dim title As String = GetSlideTitle(slide)
                If Not String.IsNullOrEmpty(title) Then
                    builder.AppendLine($"����: {title}")
                End If

                ' ��ȡ�õ�Ƭ�ϵ�����
                builder.AppendLine("����:")
                Dim slideContent = GetSlideContent(slide)
                builder.AppendLine(slideContent)

                ' ��ӷָ���
                builder.AppendLine("---")
            Next

            ' ����и���õ�Ƭδ��ʾ�������ʾ
            If slideRange.Count > maxSlides Then
                builder.AppendLine($"... [��ѡ�� {slideRange.Count} �Żõ�Ƭ������ʾǰ {maxSlides} ��]")
            End If

        Catch ex As Exception
            builder.AppendLine($"[����õ�Ƭѡ��ʱ����: {ex.Message}]")
        End Try
    End Sub

    ' ��ȡ�õ�Ƭ����
    Private Function GetSlideTitle(slide As Microsoft.Office.Interop.PowerPoint.Slide) As String
        Try
            ' ���õ�Ƭ�Ƿ��б���ռλ��
            For Each shape In slide.Shapes
                If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                    If shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                        If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            Return shape.TextFrame.TextRange.Text.Trim()
                        End If
                    End If
                End If
            Next

            ' ���û���ҵ�����ռλ�������Բ����κο��ܵı���
            For Each shape In slide.Shapes
                If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    Dim text = shape.TextFrame.TextRange.Text.Trim()
                    If Not String.IsNullOrEmpty(text) AndAlso text.Length < 100 Then
                        Return text ' �����һ������ı��Ǳ���
                    End If
                End If
            Next

            Return "[�ޱ���]"
        Catch ex As Exception
            Debug.WriteLine($"��ȡ�õ�Ƭ����ʱ����: {ex.Message}")
            Return "[��ȡ�������]"
        End Try
    End Function

    ' ��ȡ�õ�Ƭ����
    Private Function GetSlideContent(slide As Microsoft.Office.Interop.PowerPoint.Slide) As String
        Try
            Dim contentBuilder As New StringBuilder()
            Dim processedTextShapes As Integer = 0
            Dim maxTextShapes As Integer = 5 ' ����ÿ�Żõ�Ƭ������ı���״����

            ' ����õ�Ƭ�ϵ���״
            For Each shape In slide.Shapes
                ' ����������״����Ϊ�Ѿ������������
                If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder AndAlso
               shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                    Continue For
                End If

                ' �����ı���״
                If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
               shape.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then

                    If processedTextShapes >= maxTextShapes Then
                        contentBuilder.AppendLine("  [�����ı�����δ��ʾ...]")
                        Exit For
                    End If

                    Dim text = shape.TextFrame.TextRange.Text.Trim()
                    If Not String.IsNullOrEmpty(text) Then
                        ' �����ı�����
                        If text.Length > 200 Then
                            contentBuilder.AppendLine($"  �ı�: {text.Substring(0, 200)}...")
                        Else
                            contentBuilder.AppendLine($"  �ı�: {text}")
                        End If
                        processedTextShapes += 1
                    End If
                    ' ��������״
                ElseIf shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    contentBuilder.AppendLine("  [�������]")
                    ' ����ͼƬ��״
                ElseIf shape.Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                    contentBuilder.AppendLine("  [����ͼƬ]")
                    If shape.AlternativeText <> "" Then
                        contentBuilder.AppendLine($"  ͼƬ˵��: {shape.AlternativeText}")
                    End If
                    ' ����ͼ����״
                ElseIf shape.Type = Microsoft.Office.Core.MsoShapeType.msoChart Then
                    contentBuilder.AppendLine("  [����ͼ��]")
                    ' ����SmartArt��״
                ElseIf shape.Type = Microsoft.Office.Core.MsoShapeType.msoSmartArt Then
                    contentBuilder.AppendLine("  [����SmartArtͼ��]")
                End If
            Next

            ' ���û���ҵ��κ�����
            If contentBuilder.Length = 0 Then
                Return "  [�õ�Ƭ�޿���ȡ���ı�����]"
            End If

            Return contentBuilder.ToString()
        Catch ex As Exception
            Debug.WriteLine($"��ȡ�õ�Ƭ����ʱ����: {ex.Message}")
            Return $"  [��ȡ���ݳ���: {ex.Message}]"
        End Try
    End Function

    Protected Overrides Function GetCurrentWorkingDirectory() As String
        Try
            ' ��ȡ��ǰ���������·��
            If Globals.ThisAddIn.Application.ActiveWorkbook IsNot Nothing Then
                Return Globals.ThisAddIn.Application.ActiveWorkbook.Path
            End If
        Catch ex As Exception
            Debug.WriteLine($"��ȡ��ǰ����Ŀ¼ʱ����: {ex.Message}")
        End Try

        ' ����޷���ȡ������·�����򷵻�Ӧ�ó���Ŀ¼
        Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function
End Class

