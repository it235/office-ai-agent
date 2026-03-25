' ShareRibbon\Controls\Services\ContextRelevanceService.vb
' 上下文相关性评估：判断话题是否转移，动态过滤历史消息

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Text.RegularExpressions

''' <summary>
''' 上下文相关性评估服务：纯本地关键词匹配，无需调用外部 API。
''' 用于判断话题是否转移，以便动态过滤历史消息，保留真正相关的上下文。
''' </summary>
Public Class ContextRelevanceService

    ''' <summary>话题转移阈值：低于此分数视为话题已切换</summary>
    Public Const TopicShiftThreshold As Double = 0.3

    ''' <summary>历史消息相关性保留阈值</summary>
    Public Const RelevanceRetainThreshold As Double = 0.2

    ' 中文/英文停用词（过滤掉无意义的高频词）
    Private Shared ReadOnly StopWords As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "的", "了", "在", "是", "我", "有", "和", "就", "不", "人", "都", "一", "一个",
        "上", "也", "很", "到", "说", "要", "去", "你", "会", "着", "没有", "看", "好",
        "吗", "这", "那", "他", "她", "它", "们", "什么", "怎么", "为什么", "如何",
        "请", "帮", "我", "能", "可以", "需要", "想", "让", "把", "给", "用",
        "the", "a", "an", "is", "are", "was", "were", "be", "been", "being",
        "have", "has", "had", "do", "does", "did", "will", "would", "could",
        "should", "may", "might", "shall", "can", "to", "of", "in", "for",
        "on", "with", "at", "by", "from", "as", "or", "and", "but", "not",
        "it", "its", "this", "that", "i", "you", "he", "she", "we", "they"
    }

    ''' <summary>
    ''' 评估新消息与历史消息列表的相关性（0=完全无关，1=完全相关）
    ''' </summary>
    Public Shared Function EvaluateRelevance(newMessage As String, historyMessages As List(Of HistoryMessage)) As Double
        If String.IsNullOrWhiteSpace(newMessage) OrElse historyMessages Is Nothing OrElse historyMessages.Count = 0 Then
            Return 1.0 ' 没有历史则默认相关
        End If

        Dim newKeywords = ExtractKeywords(newMessage)
        If newKeywords.Count = 0 Then Return 1.0

        ' 取最近 N 条非 system 消息的关键词合集
        Dim allNonSystem = historyMessages.
            Where(Function(m) m.role <> "system" AndAlso Not String.IsNullOrWhiteSpace(m.content)).ToList()
        Dim recentNonSystem = allNonSystem.Skip(Math.Max(0, allNonSystem.Count - 6)).ToList()

        If recentNonSystem.Count = 0 Then Return 1.0

        Dim historyKeywords As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each msg In recentNonSystem
            For Each kw In ExtractKeywords(msg.content)
                historyKeywords.Add(kw)
            Next
        Next

        Return JaccardSimilarity(newKeywords, historyKeywords)
    End Function

    ''' <summary>
    ''' 判断话题是否已转移（新消息与历史相关性低于阈值）
    ''' </summary>
    Public Shared Function IsTopicShift(newMessage As String, historyMessages As List(Of HistoryMessage)) As Boolean
        If historyMessages Is Nothing OrElse historyMessages.Where(Function(m) m.role <> "system").Count() < 2 Then
            Return False ' 历史太短，不判断话题转移
        End If
        Dim score = EvaluateRelevance(newMessage, historyMessages)
        Debug.WriteLine($"[ContextRelevance] 话题相关性分数: {score:F3}，阈值: {TopicShiftThreshold}")
        Return score < TopicShiftThreshold
    End Function

    ''' <summary>
    ''' 按相关性过滤历史消息：始终保留 system 消息 + 最近 2 条（保证对话连贯）+ 相关性高的消息
    ''' </summary>
    Public Shared Function FilterHistoryByRelevance(historyMessages As List(Of HistoryMessage), newMessage As String) As List(Of HistoryMessage)
        If historyMessages Is Nothing OrElse historyMessages.Count = 0 Then
            Return New List(Of HistoryMessage)()
        End If

        If String.IsNullOrWhiteSpace(newMessage) Then
            Return New List(Of HistoryMessage)(historyMessages)
        End If

        Dim newKeywords = ExtractKeywords(newMessage)
        Dim result As New List(Of HistoryMessage)()
        Dim nonSystemMessages = historyMessages.Where(Function(m) m.role <> "system").ToList()

        ' 始终保留 system 消息
        For Each msg In historyMessages.Where(Function(m) m.role = "system")
            result.Add(msg)
        Next

        If nonSystemMessages.Count = 0 Then Return result

        ' 最近 2 条无条件保留（保证对话连贯性）
        Dim alwaysKeep As New HashSet(Of HistoryMessage)(nonSystemMessages.Skip(Math.Max(0, nonSystemMessages.Count - 2)))

        ' 对其余消息按相关性评分过滤
        For Each msg In nonSystemMessages
            If alwaysKeep.Contains(msg) Then
                result.Add(msg)
                Continue For
            End If

            If String.IsNullOrWhiteSpace(msg.content) Then Continue For

            Dim msgKeywords = ExtractKeywords(msg.content)
            Dim score = JaccardSimilarity(newKeywords, msgKeywords)
            If score >= RelevanceRetainThreshold Then
                result.Add(msg)
                Debug.WriteLine($"[ContextRelevance] 保留历史消息（分数 {score:F3}）: {msg.content.Substring(0, Math.Min(40, msg.content.Length))}...")
            Else
                Debug.WriteLine($"[ContextRelevance] 过滤历史消息（分数 {score:F3}）: {msg.content.Substring(0, Math.Min(40, msg.content.Length))}...")
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' 计算两个关键词集合的 Jaccard 相似度
    ''' </summary>
    Public Shared Function JaccardSimilarity(setA As HashSet(Of String), setB As HashSet(Of String)) As Double
        If setA.Count = 0 OrElse setB.Count = 0 Then Return 0.0
        Dim intersection = setA.Intersect(setB).Count()
        Dim union = setA.Union(setB).Count()
        If union = 0 Then Return 0.0
        Return CDbl(intersection) / CDbl(union)
    End Function

    ''' <summary>
    ''' 从文本中提取有效关键词（去停用词、短词）
    ''' </summary>
    Public Shared Function ExtractKeywords(text As String) As HashSet(Of String)
        Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        If String.IsNullOrWhiteSpace(text) Then Return result

        ' 分词：按空白符和中文标点切分
        Dim tokens = Regex.Split(text.ToLowerInvariant(), "[\\s，。！？、；：""''【】（）《》\.,!?;:\[\]()""']+")

        For Each token In tokens
            Dim t = token.Trim()
            If t.Length < 2 Then Continue For
            If StopWords.Contains(t) Then Continue For
            ' 过滤纯数字
            If Regex.IsMatch(t, "^\d+$") Then Continue For
            result.Add(t)
        Next

        ' 对中文文本进行 2-gram 分词补充
        If ContainsChinese(text) Then
            Dim cleaned = Regex.Replace(text.ToLowerInvariant(), "[^\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff]", " ")
            Dim chars = cleaned.Replace(" ", "")
            For i = 0 To chars.Length - 2
                Dim bigram = chars.Substring(i, 2)
                If Not StopWords.Contains(bigram) Then
                    result.Add(bigram)
                End If
            Next
        End If

        Return result
    End Function

    Private Shared Function ContainsChinese(text As String) As Boolean
        Return Regex.IsMatch(text, "[\u4e00-\u9fff]")
    End Function

End Class
