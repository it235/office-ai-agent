' ShareRibbon\Loop\Checkers\PreSendChecker.vb
' 发送前上下文校验器实现

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

''' <summary>
''' 发送前上下文校验器 - 在调用AI前校验用户请求和上下文完整性
''' </summary>
Public Class PreSendChecker
    Implements IContextChecker

    Public Async Function CheckAsync(context As ExecutionContext) As Task(Of ContextCheckResult) Implements IContextChecker.CheckAsync
        Dim result As New ContextCheckResult()

        ' 检查1: 用户输入是否包含明确的操作意图
        If Not HasClearIntent(context.UserMessage) Then
            result.Warnings.Add("用户输入可能不够明确，建议补充具体的排版/校对要求")
        End If

        ' 检查2: 文档内容是否可访问
        If context.OfficeContent Is Nothing OrElse context.OfficeContent.IsEmpty Then
            result.Errors.Add("无法获取当前文档内容，请确保文档已打开且有内容")
        End If

        ' 检查3: 选区状态检查（排版/校对通常需要选区）
        If context.RequiresSelection Then
            If context.SelectionInfo Is Nothing Then
                ' 如果是全文操作（如全文排版），允许无选区
                If Not IsFullDocumentOperation(context.RequestedOperation, context.UserMessage) Then
                    result.Warnings.Add("当前操作建议选中内容，未检测到有效选区将处理全文")
                End If
            End If
        End If

        ' 检查4: 文档是否处于可编辑状态
        If context.OfficeContent IsNot Nothing AndAlso context.OfficeContent.IsReadOnly Then
            result.Errors.Add("当前文档为只读状态，无法执行修改操作")
        End If

        ' 检查5: 历史上下文是否完整（追问模式下）
        If context.IsFollowUp AndAlso (context.ConversationHistory Is Nothing OrElse context.ConversationHistory.Count = 0) Then
            result.Warnings.Add("追问模式但无历史对话记录，可能导致上下文丢失")
        End If

        ' 检查6: 指令类型安全校验
        Select Case context.RequestedOperation
            Case RequestedOperation.Reformat
                If Not HasAvailableTemplateOrGuide() Then
                    result.Warnings.Add("未配置排版模板或规范，AI将使用默认规则")
                End If

            Case RequestedOperation.Proofread
                If context.OfficeContent IsNot Nothing AndAlso context.OfficeContent.ParagraphCount > 100 Then
                    result.Warnings.Add("文档段落较多（超过100段），校对可能需要较长时间")
                End If

            Case RequestedOperation.Translation
                ' 翻译通常需要明确的源语言和目标语言
                If Not ContainsLanguageHint(context.UserMessage) Then
                    result.Warnings.Add("未明确指定翻译目标语言，AI将自动判断")
                End If
        End Select

        ' 检查7: 用户消息长度
        If Not String.IsNullOrEmpty(context.UserMessage) AndAlso context.UserMessage.Length < 5 Then
            result.Warnings.Add("用户输入较短，AI可能无法准确理解意图")
        End If

        ' 设置校验结果
        result.IsValid = result.Errors.Count = 0
        If Not result.IsValid Then
            result.SuggestedClarification = GenerateClarification(result.Errors, result.Warnings)
        End If

        Return result
    End Function

    ''' <summary>
    ''' 判断用户输入是否包含明确意图
    ''' </summary>
    Private Function HasClearIntent(message As String) As Boolean
        If String.IsNullOrWhiteSpace(message) Then Return False

        ' 关键词判断
        Dim intentKeywords As String() = {
            "排版", "格式", "样式", "字体", "行距", "对齐", "缩进",
            "校对", "检查", "错别字", "标点", "语法",
            "翻译", "转成", "改为",
            "生成", "创建", "插入", "添加",
            "修改", "调整", "设置", "改成"
        }

        Dim lowerMsg = message.ToLower()
        Return intentKeywords.Any(Function(k) lowerMsg.Contains(k))
    End Function

    ''' <summary>
    ''' 判断是否为全文操作
    ''' </summary>
    Private Function IsFullDocumentOperation(operation As RequestedOperation, message As String) As Boolean
        If operation = RequestedOperation.Reformat OrElse operation = RequestedOperation.Proofread Then
            Dim fullDocKeywords As String() = {"全文", "整个", "全部", "所有", "文档"}
            Dim lowerMsg = message.ToLower()
            Return fullDocKeywords.Any(Function(k) lowerMsg.Contains(k))
        End If
        Return False
    End Function

    ''' <summary>
    ''' 检查是否有可用的模板或规范
    ''' </summary>
    Private Function HasAvailableTemplateOrGuide() As Boolean
        Try
            ' 检查是否有预置模板或用户自定义模板
            If ReformatTemplateManager.Instance.Templates.Count > 0 Then
                Return True
            End If
            ' 检查是否有语义映射
            If SemanticMappingManager.Instance.Mappings.Count > 0 Then
                Return True
            End If
        Catch
        End Try
        Return False
    End Function

    ''' <summary>
    ''' 检查是否包含语言提示
    ''' </summary>
    Private Function ContainsLanguageHint(message As String) As Boolean
        Dim langKeywords As String() = {
            "中文", "英文", "英语", "日语", "韩语", "法语", "德语", "俄语",
            "翻译成", "转为", "改为"
        }
        Dim lowerMsg = message.ToLower()
        Return langKeywords.Any(Function(k) lowerMsg.Contains(k))
    End Function

    ''' <summary>
    ''' 生成用户澄清提示
    ''' </summary>
    Private Function GenerateClarification(errors As List(Of String), warnings As List(Of String)) As String
        Dim sb As New StringBuilder()

        If errors.Count > 0 Then
            sb.AppendLine("当前无法执行操作，请解决以下问题：")
            For Each e In errors
                sb.AppendLine($"  - {e}")
            Next
        End If

        If warnings.Count > 0 Then
            If errors.Count > 0 Then
                sb.AppendLine()
            End If
            sb.AppendLine("温馨提示：")
            For Each e In warnings
                sb.AppendLine($"  - {e}")
            Next
        End If

        Return sb.ToString()
    End Function

End Class
