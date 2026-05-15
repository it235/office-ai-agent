' ShareRibbon\Loop\Interfaces\IInstructionGenerator.vb
' 指令生成器接口

Imports System.Threading.Tasks

''' <summary>
''' 指令生成器接口 - 生成AI请求的Prompt并获取响应
''' </summary>
Public Interface IInstructionGenerator

    ''' <summary>
    ''' 生成操作指令
    ''' </summary>
    ''' <param name="context">执行上下文</param>
    ''' <param name="plan">规划结果</param>
    ''' <returns>AI响应内容</returns>
    Function GenerateAsync(context As ExecutionContext, plan As PlanningResult) As Task(Of String)

    ''' <summary>
    ''' 生成修正请求（当校验失败时）
    ''' </summary>
    ''' <param name="originalResponse">原始AI响应</param>
    ''' <param name="validationErrors">校验错误列表</param>
    ''' <returns>修正后的AI响应</returns>
    Function GenerateCorrectionAsync(originalResponse As String, validationErrors As List(Of InstructionError)) As Task(Of String)

End Interface
