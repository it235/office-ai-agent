' ShareRibbon\Controls\Services\FileParserService.vb
' 文件解析服务：文本、CSV、编码检测等

Imports System.IO
Imports System.Text

''' <summary>
''' 文件解析服务，负责解析文本文件、CSV文件和检测文件编码
''' </summary>
Public Class FileParserService

#Region "文本文件解析"

        ''' <summary>
        ''' 解析文本文件
        ''' </summary>
        Public Function ParseTextFile(filePath As String) As FileContentResult
            Try
                Dim extension As String = Path.GetExtension(filePath).ToLower()

                ' 对 CSV 文件使用专门的处理逻辑
                If extension = ".csv" Then
                    Return ParseCsvFile(filePath)
                End If

                ' 对普通文本文件进行编码检测
                Dim encoding As Encoding = DetectFileEncoding(filePath)
                Dim content As String = File.ReadAllText(filePath, encoding)

                Return New FileContentResult With {
                    .FileName = Path.GetFileName(filePath),
                    .FileType = "Text",
                    .ParsedContent = content,
                    .RawData = content
                }
            Catch ex As Exception
                Return New FileContentResult With {
                    .FileName = Path.GetFileName(filePath),
                    .FileType = "Text",
                    .ParsedContent = $"[解析文本文件时出错: {ex.Message}]"
                }
            End Try
        End Function

#End Region

#Region "CSV 文件解析"

        ''' <summary>
        ''' 解析 CSV 文件
        ''' </summary>
        Public Function ParseCsvFile(filePath As String) As FileContentResult
            Try
                ' 检测文件编码
                Dim encoding As Encoding = DetectFileEncoding(filePath)
                Dim csvContent As String = File.ReadAllText(filePath, encoding)

                ' 创建一个格式化的 CSV 内容
                Dim formattedContent As New StringBuilder()
                formattedContent.AppendLine($"CSV 文件: {Path.GetFileName(filePath)} (编码: {encoding.EncodingName})")
                formattedContent.AppendLine()

                ' 分析 CSV 数据结构
                Dim rows As String() = csvContent.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

                If rows.Length > 0 Then
                    ' 检测分隔符
                    Dim delimiter As Char = DetectCsvDelimiter(rows(0))
                    Dim columns As String() = rows(0).Split(delimiter)

                    ' 添加表头
                    formattedContent.AppendLine("表头:")
                    formattedContent.AppendLine(FormatCsvRow(rows(0), delimiter))
                    formattedContent.AppendLine()

                    ' 添加数据行
                    Dim maxRows As Integer = Math.Min(rows.Length, 25)
                    formattedContent.AppendLine("数据:")

                    For i As Integer = 1 To maxRows - 1
                        formattedContent.AppendLine(FormatCsvRow(rows(i), delimiter))
                    Next

                    If rows.Length > maxRows Then
                        formattedContent.AppendLine("...")
                        formattedContent.AppendLine($"[文件包含 {rows.Length} 行，仅显示前 {maxRows - 1} 行数据]")
                    End If
                Else
                    formattedContent.AppendLine("[CSV 文件为空]")
                End If

                Return New FileContentResult With {
                    .FileName = Path.GetFileName(filePath),
                    .FileType = "CSV",
                    .ParsedContent = formattedContent.ToString(),
                    .RawData = csvContent
                }
            Catch ex As Exception
                Return New FileContentResult With {
                    .FileName = Path.GetFileName(filePath),
                    .FileType = "CSV",
                    .ParsedContent = $"[解析 CSV 文件时出错: {ex.Message}]"
                }
            End Try
        End Function

        ''' <summary>
        ''' 格式化 CSV 行
        ''' </summary>
        Private Function FormatCsvRow(row As String, delimiter As Char) As String
            Dim fields As String() = row.Split(delimiter)
            Dim formattedRow As New StringBuilder()

            For i As Integer = 0 To fields.Length - 1
                Dim field As String = fields(i).Trim(""""c)
                If i < fields.Length - 1 Then
                    formattedRow.Append($"{field} | ")
                Else
                    formattedRow.Append(field)
                End If
            Next

            Return formattedRow.ToString()
        End Function

        ''' <summary>
        ''' 检测 CSV 分隔符
        ''' </summary>
        Public Function DetectCsvDelimiter(sampleLine As String) As Char
            Dim possibleDelimiters As Char() = {","c, ";"c, vbTab, "|"c}
            Dim bestDelimiter As Char = ","c
            Dim maxCount As Integer = 0

            For Each delimiter In possibleDelimiters
                Dim count As Integer = sampleLine.Count(Function(c) c = delimiter)
                If count > maxCount Then
                    maxCount = count
                    bestDelimiter = delimiter
                End If
            Next

            Return bestDelimiter
        End Function

#End Region

#Region "编码检测"

        ''' <summary>
        ''' 检测文件编码
        ''' </summary>
        Public Function DetectFileEncoding(filePath As String) As Encoding
            Try
                Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                    ' 读取前几个字节来检测 BOM
                    Dim bom(3) As Byte
                    Dim bytesRead As Integer = fs.Read(bom, 0, bom.Length)

                    ' UTF-8 with BOM
                    If bytesRead >= 3 AndAlso bom(0) = &HEF AndAlso bom(1) = &HBB AndAlso bom(2) = &HBF Then
                        Return New UTF8Encoding(True)
                    End If

                    ' UTF-16 Big Endian
                    If bytesRead >= 2 AndAlso bom(0) = &HFE AndAlso bom(1) = &HFF Then
                        Return Encoding.BigEndianUnicode
                    End If

                    ' UTF-16 Little Endian
                    If bytesRead >= 2 AndAlso bom(0) = &HFF AndAlso bom(1) = &HFE Then
                        If bytesRead >= 4 AndAlso bom(2) = 0 AndAlso bom(3) = 0 Then
                            Return Encoding.UTF32
                        Else
                            Return Encoding.Unicode
                        End If
                    End If

                    ' UTF-32 Big Endian
                    If bytesRead >= 4 AndAlso bom(0) = 0 AndAlso bom(1) = 0 AndAlso bom(2) = &HFE AndAlso bom(3) = &HFF Then
                        Return New UTF32Encoding(True, True)
                    End If
                End Using

                ' 针对 CSV 文件优先尝试 GB18030/GBK 编码
                Dim fileExtension As String = Path.GetExtension(filePath).ToLower()
                If fileExtension = ".csv" Then
                    Dim gbkResult = TryGbkEncoding(filePath)
                    If gbkResult IsNot Nothing Then
                        Return gbkResult
                    End If
                End If

                ' 尝试其他编码
                Return DetectEncodingByContent(filePath)

            Catch ex As Exception
                Return Encoding.Default
            End Try
        End Function

        ''' <summary>
        ''' 尝试 GBK 编码
        ''' </summary>
        Private Function TryGbkEncoding(filePath As String) As Encoding
            Try
                Dim csvSampleBytes As Byte() = New Byte(4095) {}
                Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                    fs.Read(csvSampleBytes, 0, csvSampleBytes.Length)
                End Using

                Dim gbkEncoding As Encoding = Encoding.GetEncoding("GB18030")
                Dim gbkText As String = gbkEncoding.GetString(csvSampleBytes)

                If gbkText.Contains(",") AndAlso (gbkText.Contains(vbCr) OrElse gbkText.Contains(vbLf)) Then
                    Dim unicodeReplacementChar As Char = ChrW(&HFFFD)
                    Dim invalidCharCount As Integer = gbkText.Count(Function(c) c = "?"c Or c = unicodeReplacementChar)
                    Dim totalCharCount As Integer = gbkText.Length

                    If invalidCharCount <= totalCharCount * 0.05 Then
                        Return gbkEncoding
                    End If
                End If
            Catch
                ' 忽略错误
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' 通过内容检测编码
        ''' </summary>
        Private Function DetectEncodingByContent(filePath As String) As Encoding
            Dim encodingsToTry As Encoding() = {
                New UTF8Encoding(False),
                Encoding.GetEncoding("GB18030"),
                Encoding.Default
            }

            Dim sampleBytes As Byte() = New Byte(4095) {}
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read)
                fs.Read(sampleBytes, 0, sampleBytes.Length)
            End Using

            Dim bestEncoding As Encoding = encodingsToTry(0)
            Dim leastInvalidCharCount As Integer = Integer.MaxValue
            Dim unicodeReplacementChar As Char = ChrW(&HFFFD)

            For Each enc In encodingsToTry
                Try
                    Dim sample As String = enc.GetString(sampleBytes)
                    Dim invalidCharCount As Integer = sample.Count(Function(c) c = "?"c Or c = unicodeReplacementChar)

                    If invalidCharCount < leastInvalidCharCount Then
                        leastInvalidCharCount = invalidCharCount
                        bestEncoding = enc

                        If invalidCharCount = 0 Then
                            Exit For
                        End If
                    End If
                Catch
                    Continue For
                End Try
            Next

            Return bestEncoding
        End Function

#End Region

    End Class
