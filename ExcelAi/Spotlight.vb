Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports ShareRibbon

Public Class Spotlight
    ' �����༶����������ڸ��پ۹�ƹ���״̬
    Private _spotlightActive As Boolean = False
    Private WithEvents _appEvents As Excel.Application
    Private _currentWorkbook As Excel.Workbook

    ' ���ٵ�ǰ�۹��λ��
    Private _currentRow As Integer = 0
    Private _currentColumn As Integer = 0

    ' ������ʽ����
    Private Const ROW_FORMAT_NAME As String = "SpotlightRow"
    Private Const COLUMN_FORMAT_NAME As String = "SpotlightColumn"
    Private Const CELL_FORMAT_NAME As String = "SpotlightCell"

    ' �۹������ - �޸�Ĭ����ɫΪǳ��ɫ
    Private _rowColor As Integer = RGB(230, 230, 230) ' ǳ��ɫ
    Private _columnColor As Integer = RGB(230, 230, 230) ' ǳ��ɫ
    Private _cellColor As Integer = RGB(200, 200, 200) ' ����Ļ�ɫ���û��Ԫ�������
    Private _rowDisplay As Boolean = True ' �Ƿ���ʾ�и���
    Private _columnDisplay As Boolean = True ' �Ƿ���ʾ�и���

    ' ����ģʽʵ��
    Private Shared _instance As Spotlight = Nothing

    ' ��ȡ����ʵ��
    Public Shared Function GetInstance() As Spotlight
        If _instance Is Nothing Then
            _instance = New Spotlight()
        End If
        Return _instance
    End Function

    ' ˽�й��캯������ֹ�ⲿֱ�Ӵ���ʵ��
    Private Sub New()
    End Sub

    ' ���۹���Ƿ񼤻�
    Public ReadOnly Property IsActive As Boolean
        Get
            Return _spotlightActive
        End Get
    End Property

    ' �л�����ʾ
    Public Sub ToggleRowDisplay()
        _rowDisplay = Not _rowDisplay
        If _spotlightActive Then
            UpdateHighlight()
        End If
    End Sub

    ' �л�����ʾ
    Public Sub ToggleColumnDisplay()
        _columnDisplay = Not _columnDisplay
        If _spotlightActive Then
            UpdateHighlight()
        End If
    End Sub

    ' ���þ۹����ɫ
    Public Sub SetColors(rowColor As Integer, columnColor As Integer, cellColor As Integer)
        _rowColor = rowColor
        _columnColor = columnColor
        _cellColor = cellColor

        ' ����۹���Ѽ������Ӧ������ɫ
        If _spotlightActive Then
            ApplyHighlight()
        End If
    End Sub

    ' ��ʾ��ɫѡ��Ի���
    Public Sub ShowColorDialog()
        Try
            ' ������ɫ�Ի���
            Using colorDialog As New ColorDialog()
                ' ���ó�ʼ��ɫ
                colorDialog.Color = System.Drawing.ColorTranslator.FromOle(_rowColor)
                colorDialog.FullOpen = True ' ��ʾ��������ɫ�Ի���
                colorDialog.CustomColors = New Integer() {
                    RGB(230, 230, 230), ' ǳ��ɫ
                    RGB(255, 255, 150), ' ǳ��ɫ
                    RGB(200, 255, 200), ' ǳ��ɫ
                    RGB(200, 200, 255), ' ǳ��ɫ
                    RGB(255, 200, 200)  ' ǳ��ɫ
                }

                ' ��ʾ�Ի���
                If colorDialog.ShowDialog() = DialogResult.OK Then
                    ' �û�ѡ������ɫ�����¾۹����ɫ
                    Dim selectedColor As Integer = System.Drawing.ColorTranslator.ToOle(colorDialog.Color)

                    ' �к���ʹ��ѡ�����ɫ
                    _rowColor = selectedColor
                    _columnColor = selectedColor

                    ' ���Ԫ��ʹ���������ɫ
                    Dim cellR As Integer = colorDialog.Color.R - 30
                    Dim cellG As Integer = colorDialog.Color.G - 30
                    Dim cellB As Integer = colorDialog.Color.B - 30

                    ' ȷ��RGBֵ��С��0
                    cellR = Math.Max(0, cellR)
                    cellG = Math.Max(0, cellG)
                    cellB = Math.Max(0, cellB)

                    _cellColor = RGB(cellR, cellG, cellB)

                    ' ����۹���Ѽ��Ӧ������ɫ
                    If _spotlightActive Then
                        ApplyHighlight()
                    End If

                    GlobalStatusStripAll.ShowWarning("�۹����ɫ�Ѹ���")
                End If
            End Using
        Catch ex As Exception
            Debug.WriteLine("��ʾ��ɫ�Ի���ʱ����: " & ex.Message)
            MessageBox.Show("��ʾ��ɫ�Ի���ʱ����: " & ex.Message, "�۹����ɫ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' �л��۹��״̬
    Public Function Toggle() As Boolean
        If _spotlightActive Then
            Deactivate()
            GlobalStatusStripAll.ShowWarning("�۹�ƹ����ѹر�,˫���۹�ư�ť���޸���ɫ")
        Else
            Activate()
            GlobalStatusStripAll.ShowWarning("�۹�ƹ����ѿ�����˫���۹�ư�ť���޸���ɫ")
        End If

        Return _spotlightActive
    End Function

    ' ����۹�ƹ���
    Public Sub Activate()
        Try
            If _appEvents Is Nothing Then
                _appEvents = Globals.ThisAddIn.Application
            End If

            ' ����Ե�ǰ�������������
            _currentWorkbook = _appEvents.ActiveWorkbook

            ' ����ԭʼ����
            _spotlightActive = True

            ' ����¼��������
            AddHandler _appEvents.SheetSelectionChange, AddressOf AppEvents_SheetSelectionChange
            AddHandler _appEvents.SheetActivate, AddressOf AppEvents_SheetActivate
            AddHandler _appEvents.SheetDeactivate, AddressOf AppEvents_SheetDeactivate
            AddHandler _appEvents.WorkbookActivate, AddressOf AppEvents_WorkbookActivate

            ' Ӧ�ø���
            ApplyHighlight()

            ' ������Ϣ
            Debug.WriteLine("�۹�ƹ����Ѽ���")
        Catch ex As Exception
            Debug.WriteLine("����۹�ƹ���ʱ����: " & ex.Message)
            MessageBox.Show("����۹�ƹ���ʱ����: " & ex.Message, "�۹�ƹ���", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ȡ������۹�ƹ���
    Public Sub Deactivate()
        Try
            If _appEvents IsNot Nothing Then
                ' �Ƴ��¼��������
                RemoveHandler _appEvents.SheetSelectionChange, AddressOf AppEvents_SheetSelectionChange
                RemoveHandler _appEvents.SheetActivate, AddressOf AppEvents_SheetActivate
                RemoveHandler _appEvents.SheetDeactivate, AddressOf AppEvents_SheetDeactivate
                RemoveHandler _appEvents.WorkbookActivate, AddressOf AppEvents_WorkbookActivate
            End If

            ' �Ƴ�����
            RemoveHighlight()

            ' ����״̬
            _spotlightActive = False
            _currentRow = 0
            _currentColumn = 0

            ' ������Ϣ
            Debug.WriteLine("�۹�ƹ�����ͣ��")
        Catch ex As Exception
            Debug.WriteLine("ȡ������۹�ƹ���ʱ����: " & ex.Message)
        End Try
    End Sub

    ' ������ѡ������¼��������
    Private Sub AppEvents_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)
        If _spotlightActive Then
            Debug.WriteLine("ѡ��λ���Ѹ���")
            UpdateHighlight()

            ' ȷ��ѡ�е�Ԫ����Ȼ��ѡ��״̬
            Target.Select()

            ' ��Excel���¼��㹫ʽ��ȷ����Ԫ��������Ч
            If Target.Count = 1 Then
                Try
                    Target.Calculate()
                Catch
                    ' ���Լ������
                End Try
            End If
        End If
    End Sub

    ' ���������¼��������
    Private Sub AppEvents_SheetActivate(ByVal Sh As Object)
        If _spotlightActive Then
            Debug.WriteLine("�������Ѽ���")
            ApplyHighlight()
        End If
    End Sub

    ' ������ͣ���¼��������
    Private Sub AppEvents_SheetDeactivate(ByVal Sh As Object)
        If _spotlightActive Then
            Debug.WriteLine("��������ͣ��")
            RemoveHighlight()
        End If
    End Sub

    ' �����������¼��������
    Private Sub AppEvents_WorkbookActivate(ByVal Wb As Excel.Workbook)
        If _spotlightActive Then
            Debug.WriteLine("�������Ѽ���")
            _currentWorkbook = Wb
            ApplyHighlight()
        End If
    End Sub

    ' Ӧ�ø���
    Private Sub ApplyHighlight()
        Try
            ' �����Ƴ����еĸ���
            RemoveHighlight()

            ' ��ȡ��ǰ���Ԫ��
            Dim activeCell As Excel.Range = _appEvents.ActiveCell
            Dim activeSheet As Excel.Worksheet = _appEvents.ActiveSheet

            ' ���浱ǰλ��
            _currentRow = activeCell.Row
            _currentColumn = activeCell.Column

            Debug.WriteLine("��ǰλ��: ��=" & _currentRow & ", ��=" & _currentColumn)

            ' Ӧ��������ʽ
            ApplyConditionalFormatting(activeCell, activeSheet)
        Catch ex As Exception
            Debug.WriteLine("Ӧ�ø���ʱ����: " & ex.Message)
        End Try
    End Sub

    ' ���¸���
    Private Sub UpdateHighlight()
        Try
            ' ��ȡ��ǰ���Ԫ��
            Dim activeCell As Excel.Range = _appEvents.ActiveCell

            ' ���λ�ñ仯�ˣ�����Ӧ�ø���
            If _currentRow <> activeCell.Row OrElse _currentColumn <> activeCell.Column Then
                ApplyHighlight()
            End If
        Catch ex As Exception
            Debug.WriteLine("���¸���ʱ����: " & ex.Message)
        End Try
    End Sub

    ' �Ƴ�����
    Private Sub RemoveHighlight()
        Try
            ' ��ȡ��ǰ�������
            Dim activeSheet As Excel.Worksheet = _appEvents.ActiveSheet

            ' ɾ��������ʽ
            activeSheet.Cells.FormatConditions.Delete()

            Debug.WriteLine("�������Ƴ�")
        Catch ex As Exception
            Debug.WriteLine("�Ƴ�����ʱ����: " & ex.Message)
        End Try
    End Sub

    ' Ӧ��������ʽ
    Private Sub ApplyConditionalFormatting(activeCell As Excel.Range, activeSheet As Excel.Worksheet)
        Try
            ' ������Ļ����ΪFalse���������
            _appEvents.ScreenUpdating = False

            ' Ӧ���и���
            If _rowDisplay Then
                Dim entireRow As Excel.Range = activeSheet.Rows(_currentRow)
                Dim rowFormat As Excel.FormatCondition = entireRow.FormatConditions.Add(
                    Type:=XlFormatConditionType.xlExpression,
                    Formula1:="=ROW()=" & _currentRow)

                With rowFormat
                    .Interior.Color = _rowColor
                    .StopIfTrue = False
                End With
            End If

            ' Ӧ���и���
            If _columnDisplay Then
                Dim entireColumn As Excel.Range = activeSheet.Columns(_currentColumn)
                Dim colFormat As Excel.FormatCondition = entireColumn.FormatConditions.Add(
                    Type:=XlFormatConditionType.xlExpression,
                    Formula1:="=COLUMN()=" & _currentColumn)

                With colFormat
                    .Interior.Color = _columnColor
                    .StopIfTrue = False
                End With
            End If

            ' Ӧ�õ�Ԫ��������Ḳ�����и�����
            Dim cellFormat As Excel.FormatCondition = activeCell.FormatConditions.Add(
                Type:=XlFormatConditionType.xlExpression,
                Formula1:="=TRUE")

            With cellFormat
                .Interior.Color = _cellColor
                .StopIfTrue = False
            End With

            ' �ָ���Ļ����
            _appEvents.ScreenUpdating = True

            Debug.WriteLine("������ʽ��Ӧ��")
        Catch ex As Exception
            _appEvents.ScreenUpdating = True
            Debug.WriteLine("Ӧ��������ʽʱ����: " & ex.Message)
        End Try
    End Sub

    ' ���һ��������и����Ĺ�������
    Public Sub ClearAllHighlights()
        RemoveHighlight()
    End Sub
End Class