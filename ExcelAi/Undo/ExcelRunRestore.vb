Imports System.Collections.Generic
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon

Namespace ExcelUndo

    ''' <summary>
    ''' Excel 没有“一条命令 = 一个撤销记录”的事务边界，Application.Undo 会连用户此前的操作一起撤掉，
    ''' 所以本轮撤销走还原快照：执行前把目标区域的值、公式、数字格式整块抓下来，撤销时整块写回去，
    ''' 再用同一套签名核对是否真回到执行前。
    ''' 快照只覆盖抓下来的那块区域：插行插列、新建图表这类结构性改动一律不声明可还原，
    ''' 采集不到、区域过大、形状对不上时都返回 Nothing/拒绝，让 UI 不给“撤销本轮”按钮。
    ''' </summary>
    Public Class ExcelRunRestore

        Public Const RestoreKind As String = "excel_range_state"

        ''' <summary>采集上限，只为挡住整表级的大区域；真正决定能否落库的是共享合同的体积闸门。</summary>
        Private Const MaxCells As Integer = 4000

        ''' <summary>抓取活动工作表上目标区域的执行前状态；区域拿不到或形状无法忠实映射时返回 Nothing。</summary>
        Public Shared Function Capture(app As Object, address As String) As JObject
            If String.IsNullOrWhiteSpace(address) Then Return Nothing
            Dim workbook As Object = Nothing
            Dim sheet As Object = Nothing
            Dim targetRange As Object = Nothing
            Try
                workbook = app.ActiveWorkbook
                sheet = app.ActiveSheet
                If workbook Is Nothing OrElse sheet Is Nothing Then Return Nothing
                targetRange = sheet.Range(address)
                Dim rows = CInt(targetRange.Rows.Count)
                Dim columns = CInt(targetRange.Columns.Count)
                If rows <= 0 OrElse columns <= 0 OrElse rows * columns > MaxCells Then Return Nothing

                Dim values = ToJagged(targetRange.Value2, rows, columns)
                Dim formulas = ToJagged(targetRange.Formula, rows, columns)
                Dim formats = ToJagged(targetRange.NumberFormatLocal, rows, columns)
                If values Is Nothing OrElse formulas Is Nothing OrElse formats Is Nothing Then Return Nothing

                Return New JObject From {
                    {"workbook", CStr(workbook.Name)},
                    {"sheet", CStr(sheet.Name)},
                    {"address", address},
                    {"rows", rows},
                    {"columns", columns},
                    {"values", values},
                    {"formulas", formulas},
                    {"numberFormats", formats},
                    {"signature", Signature(values, formulas, formats)}
                }
            Catch ex As Exception
                Debug.WriteLine($"[Excel] 还原快照采集失败: {ex.Message}")
                Return Nothing
            Finally
                ComObjectHelper.ReleaseComObject(targetRange)
                ComObjectHelper.ReleaseComObject(sheet)
                ComObjectHelper.ReleaseComObject(workbook)
            End Try
        End Function

        ''' <summary>
        ''' 组装 observation.restore：区域确实变了、且宿主判断工作表结构没动（行列数、图表数）才给快照。
        ''' 快照里补上执行后的签名，重放时用它区分“本轮留下的状态”和“用户后来又改的状态”。
        ''' </summary>
        Public Shared Function BuildNode(before As JObject, after As JObject, structureStable As Boolean) As JObject
            If Not structureStable Then Return Nothing
            If before Is Nothing OrElse after Is Nothing Then Return Nothing
            Dim beforeSignature = If(before("signature")?.ToString(), "")
            Dim afterSignature = If(after("signature")?.ToString(), "")
            If beforeSignature.Length = 0 OrElse afterSignature.Length = 0 Then Return Nothing
            If String.Equals(beforeSignature, afterSignature, StringComparison.Ordinal) Then Return Nothing
            before("afterSignature") = afterSignature
            Return ShareRibbon.Agent.AgentRunChangeLog.BuildRestoreNode(RestoreKind, before)
        End Function

        ''' <summary>
        ''' 按快照把目标区域写回执行前。返回空字符串表示已还原并核对通过；
        ''' 其余任何返回值都是“没有还原”的原因，调用方据此如实告诉用户。
        ''' </summary>
        Public Shared Function Replay(app As Object, state As JObject) As String
            If state Is Nothing Then Return "还原快照缺失，已拒绝自动回退。"
            Dim workbook As Object = Nothing
            Dim sheet As Object = Nothing
            Dim targetRange As Object = Nothing
            Try
                Dim address = If(state("address")?.ToString(), "")
                Dim sheetName = If(state("sheet")?.ToString(), "")
                If String.IsNullOrWhiteSpace(address) OrElse String.IsNullOrWhiteSpace(sheetName) Then
                    Return "还原快照缺少工作表或区域信息，已拒绝自动回退。"
                End If

                workbook = app.ActiveWorkbook
                If workbook Is Nothing Then Return "没有活动工作簿，无法还原本轮改动。"
                Dim expectedWorkbook = If(state("workbook")?.ToString(), "")
                If expectedWorkbook.Length > 0 AndAlso
                   Not String.Equals(CStr(workbook.Name), expectedWorkbook, StringComparison.Ordinal) Then
                    Return $"当前活动工作簿不是本轮改动所在的「{expectedWorkbook}」，已拒绝还原。"
                End If

                Dim sheets As Object = Nothing
                Try
                    sheets = workbook.Worksheets
                    sheet = sheets.Item(sheetName)
                Catch
                    Return $"找不到工作表「{sheetName}」，已拒绝还原。"
                Finally
                    ComObjectHelper.ReleaseComObject(sheets)
                End Try
                If sheet Is Nothing Then Return $"找不到工作表「{sheetName}」，已拒绝还原。"

                targetRange = sheet.Range(address)
                Dim rows = ReadInteger(state, "rows")
                Dim columns = ReadInteger(state, "columns")
                If rows <= 0 OrElse columns <= 0 Then Return "还原快照没有记录区域形状，已拒绝自动回退。"
                If CInt(targetRange.Rows.Count) <> rows OrElse CInt(targetRange.Columns.Count) <> columns Then
                    Return $"目标区域 {address} 的形状已与本轮记录不一致，直接写回会串位，已拒绝还原。"
                End If

                Dim beforeSignature = If(state("signature")?.ToString(), "")
                Dim expectedSignature = If(state("afterSignature")?.ToString(), "")
                Dim current = ReadSignature(targetRange, rows, columns)
                If current Is Nothing Then Return "无法读取目标区域当前内容，已拒绝自动回退。"
                If Not String.IsNullOrEmpty(beforeSignature) AndAlso
                   String.Equals(current, beforeSignature, StringComparison.Ordinal) Then
                    Return ""
                End If
                If Not String.IsNullOrEmpty(expectedSignature) AndAlso
                   Not String.Equals(current, expectedSignature, StringComparison.Ordinal) Then
                    Return $"区域 {address} 在本轮之后又有其他改动，已拒绝还原以免覆盖你的内容。"
                End If

                Dim values = FromJagged(TryCast(state("values"), JArray), rows, columns)
                Dim formulas = FromJagged(TryCast(state("formulas"), JArray), rows, columns)
                Dim formats = FromJagged(TryCast(state("numberFormats"), JArray), rows, columns)
                If values Is Nothing OrElse formulas Is Nothing OrElse formats Is Nothing Then
                    Return "还原快照内容不完整，已拒绝自动回退。"
                End If

                targetRange.NumberFormatLocal = formats
                targetRange.Value2 = values
                Dim formulaError = WriteFormulas(targetRange, formulas, rows, columns)
                If formulaError.Length > 0 Then Return formulaError

                Dim restored = ReadSignature(targetRange, rows, columns)
                If restored Is Nothing Then Return "写回后无法核对目标区域，请对照变更清单人工确认。"
                If Not String.IsNullOrEmpty(beforeSignature) AndAlso
                   Not String.Equals(restored, beforeSignature, StringComparison.Ordinal) Then
                    Return $"已写回 {sheetName}!{address}，但与执行前的记录仍不一致，请对照变更清单人工核对。"
                End If
                Return ""
            Catch ex As Exception
                Debug.WriteLine($"[Excel] 还原快照重放失败: {ex.Message}")
                Return $"还原失败: {AppLogger.Redact(ex.Message)}"
            Finally
                ComObjectHelper.ReleaseComObject(targetRange)
                ComObjectHelper.ReleaseComObject(sheet)
                ComObjectHelper.ReleaseComObject(workbook)
            End Try
        End Function

        ''' <summary>
        ''' 公式逐个写回：整块写 Value2 已经把常量按原类型放回去了，再把公式单元格恢复成公式。
        ''' 非公式单元格不能拿 Formula 文本回写，否则 "1-1" 这类文本会被 Excel 重新解释成日期。
        ''' </summary>
        Private Shared Function WriteFormulas(targetRange As Object, formulas As Object(,), rows As Integer, columns As Integer) As String
            Dim cells As Object = Nothing
            Try
                cells = targetRange.Cells
                For r = 1 To rows
                    For c = 1 To columns
                        Dim text = TryCast(formulas(r - 1, c - 1), String)
                        If text Is Nothing OrElse Not text.StartsWith("=", StringComparison.Ordinal) Then Continue For
                        Dim cell As Object = Nothing
                        Try
                            cell = cells.Item(r, c)
                            cell.Formula = text
                        Finally
                            ComObjectHelper.ReleaseComObject(cell)
                        End Try
                    Next
                Next
                Return ""
            Catch ex As Exception
                Return $"公式写回失败: {AppLogger.Redact(ex.Message)}"
            Finally
                ComObjectHelper.ReleaseComObject(cells)
            End Try
        End Function

        Private Shared Function ReadSignature(targetRange As Object, rows As Integer, columns As Integer) As String
            Try
                Dim values = ToJagged(targetRange.Value2, rows, columns)
                Dim formulas = ToJagged(targetRange.Formula, rows, columns)
                Dim formats = ToJagged(targetRange.NumberFormatLocal, rows, columns)
                If values Is Nothing OrElse formulas Is Nothing OrElse formats Is Nothing Then Return Nothing
                Return Signature(values, formulas, formats)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function Signature(values As JArray, formulas As JArray, formats As JArray) As String
            Return Sha256(values.ToString(Formatting.None) & "|" &
                           formulas.ToString(Formatting.None) & "|" &
                           formats.ToString(Formatting.None))
        End Function

        Private Shared Function Sha256(value As String) As String
            Dim bytes = System.Text.Encoding.UTF8.GetBytes(If(value, ""))
            Using sha = System.Security.Cryptography.SHA256.Create()
                Return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        ''' <summary>
        ''' 把 IDispatch 返回的区域值整理成 rows×columns 的二维数组。
        ''' 单格返回标量、多格返回二维数组是两种正常形状；其它形状（比如只给回一维）无法忠实映射，
        ''' 直接返回 Nothing——宁可没有快照，也不能按猜的布局写回一片区域。
        ''' </summary>
        Private Shared Function ToJagged(raw As Object, rows As Integer, columns As Integer) As JArray
            Dim array = TryCast(raw, Array)
            If array Is Nothing Then
                If rows <> 1 OrElse columns <> 1 Then Return Nothing
                Dim singleCell As New JArray()
                singleCell.Add(New JArray(New JToken(){ToToken(raw)}))
                Return singleCell
            End If
            If array.Rank < 2 Then Return Nothing

            Dim lowerRow = LBound(array, 1)
            Dim lowerColumn = LBound(array, 2)
            If UBound(array, 1) - lowerRow + 1 < rows OrElse UBound(array, 2) - lowerColumn + 1 < columns Then Return Nothing

            Dim matrix As New JArray()
            For r = 0 To rows - 1
                Dim rowValues As New JArray()
                For c = 0 To columns - 1
                    rowValues.Add(ToToken(array.GetValue(lowerRow + r, lowerColumn + c)))
                Next
                matrix.Add(rowValues)
            Next
            Return matrix
        End Function

        Private Shared Function ToToken(value As Object) As JToken
            If value Is Nothing OrElse TypeOf value Is DBNull Then Return JValue.CreateNull()
            If TypeOf value Is Date Then Return New JValue(CType(value, Date).ToOADate())
            Return JValue.FromObject(value)
        End Function

        Private Shared Function FromJagged(matrix As JArray, rows As Integer, columns As Integer) As Object(,)
            If matrix Is Nothing OrElse matrix.Count <> rows Then Return Nothing
            Dim values(rows - 1, columns - 1) As Object
            For r = 0 To rows - 1
                Dim rowValues = TryCast(matrix.Item(r), JArray)
                If rowValues Is Nothing OrElse rowValues.Count <> columns Then Return Nothing
                For c = 0 To columns - 1
                    values(r, c) = ToCellValue(rowValues.Item(c))
                Next
            Next
            Return values
        End Function

        Private Shared Function ReadInteger(state As JObject, name As String) As Integer
            Dim token = state?(name)
            If token Is Nothing Then Return 0
            Try
                Return token.Value(Of Integer)()
            Catch
                Return 0
            End Try
        End Function

        Private Shared Function ToCellValue(token As JToken) As Object
            If token Is Nothing OrElse token.Type = JTokenType.Null Then Return DBNull.Value
            If token.Type = JTokenType.String Then Return token.Value(Of String)()
            If token.Type = JTokenType.Boolean Then Return token.Value(Of Boolean)()
            Try
                Return token.Value(Of Double)()
            Catch
                Return token.ToString()
            End Try
        End Function

    End Class

End Namespace
