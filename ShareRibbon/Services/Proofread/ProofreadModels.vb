' ShareRibbon/Services\Proofread\ProofreadModels.vb
' 校对相关数据模型

Imports System.Collections.Generic

''' <summary>
''' 校对问题数据模型
''' </summary>
Public Class ProofreadIssue
    Public Property Id As String = Guid.NewGuid().ToString()
    Public Property ParagraphIndex As Integer
    Public Property StartPosition As Integer
    Public Property EndPosition As Integer
    Public Property Original As String
    Public Property Suggestion As String
    Public Property IssueType As IssueType
    Public Property Severity As IssueSeverity
    Public Property Explanation As String
    Public Property IsIgnored As Boolean = False
    Public Property IsCorrected As Boolean = False
End Class

''' <summary>
''' 校对摘要
''' </summary>
Public Class ProofreadSummary
    Public Property TotalCount As Integer
    Public Property HighCount As Integer
    Public Property MediumCount As Integer
    Public Property LowCount As Integer
End Class

''' <summary>
''' 问题类型枚举
''' </summary>
Public Enum IssueType
    SpellingError       ' 拼写错误
    WordUsageError     ' 用词错误（的地得混用等）
    PunctuationError   ' 标点错误
    GrammaticalError   ' 语法错误
    ExpressionError   ' 表达问题
    FormatError       ' 格式问题
End Enum

''' <summary>
''' 严重程度枚举
''' </summary>
Public Enum IssueSeverity
    High    ' 必须修改
    Medium  ' 建议修改
    Low     ' 可选优化
End Enum

''' <summary>
''' 校对专注模式状态
''' </summary>
Public Class ProofreadFocusState
    Public Property IsActive As Boolean = False
    Public Property CurrentIssues As List(Of ProofreadIssue) = Nothing
    Public Property CurrentDocumentText As String = ""
    Public Property ProcessedParagraphs As List(Of String) = Nothing
    
    Public Sub New()
        CurrentIssues = New List(Of ProofreadIssue)()
        ProcessedParagraphs = New List(Of String)()
    End Sub
    
    Public Sub Reset()
        IsActive = False
        If CurrentIssues IsNot Nothing Then CurrentIssues.Clear()
        CurrentDocumentText = ""
        If ProcessedParagraphs IsNot Nothing Then ProcessedParagraphs.Clear()
    End Sub
End Class
