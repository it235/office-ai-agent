' ShareRibbon\Config\PromptConfiguration.vb
' 提示词配置数据结构

Imports System.Collections.Generic

''' <summary>
''' 提示词上下文 - 用于确定应该使用哪些提示词
''' </summary>
Public Class PromptContext
    ''' <summary>
    ''' Office应用类型 (Excel/Word/PowerPoint)
    ''' </summary>
    Public Property ApplicationType As String = "Excel"

    ''' <summary>
    ''' 意图识别结果
    ''' </summary>
    Public Property IntentResult As IntentResult

    ''' <summary>
    ''' 功能模式 (proofread/reformat/continuation/template_render等)
    ''' </summary>
    Public Property FunctionMode As String = ""

    ''' <summary>
    ''' 额外上下文信息（如选中内容摘要、工作表名等）
    ''' </summary>
    Public Property AdditionalContext As String = ""
End Class

''' <summary>
''' 提示词配置根结构
''' </summary>
Public Class PromptConfiguration
    ''' <summary>
    ''' 配置版本号
    ''' </summary>
    Public Property Version As Integer = 1

    ''' <summary>
    ''' 各Office应用的提示词配置
    ''' </summary>
    Public Property Applications As List(Of ApplicationPromptConfig) = New List(Of ApplicationPromptConfig)()
End Class

''' <summary>
''' 单个Office应用的提示词配置
''' </summary>
Public Class ApplicationPromptConfig
    ''' <summary>
    ''' 应用类型 (Excel/Word/PowerPoint)
    ''' </summary>
    Public Property Type As String = ""

    ''' <summary>
    ''' 意图专用提示词列表
    ''' </summary>
    Public Property IntentPrompts As List(Of IntentPromptTemplate) = New List(Of IntentPromptTemplate)()

    ''' <summary>
    ''' 功能模式提示词列表
    ''' </summary>
    Public Property FunctionModePrompts As List(Of FunctionModePromptTemplate) = New List(Of FunctionModePromptTemplate)()

    ''' <summary>
    ''' JSON格式约束提示词
    ''' </summary>
    Public Property JsonSchemaConstraint As String = ""
End Class

''' <summary>
''' 意图提示词模板
''' </summary>
Public Class IntentPromptTemplate
    ''' <summary>
    ''' 意图类型名称 (如 DATA_ANALYSIS, FORMULA_CALC 等)
    ''' </summary>
    Public Property IntentType As String = ""

    ''' <summary>
    ''' 提示词内容
    ''' </summary>
    Public Property Content As String = ""
End Class

''' <summary>
''' 功能模式提示词模板
''' </summary>
Public Class FunctionModePromptTemplate
    ''' <summary>
    ''' 功能模式名称 (如 proofread, reformat, continuation, template_render)
    ''' </summary>
    Public Property Mode As String = ""

    ''' <summary>
    ''' 提示词内容
    ''' </summary>
    Public Property Content As String = ""
End Class
