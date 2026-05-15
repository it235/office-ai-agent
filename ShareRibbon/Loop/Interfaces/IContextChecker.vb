' ShareRibbon\Loop\Interfaces\IContextChecker.vb
' 上下文校验器接口

Imports System.Threading.Tasks

''' <summary>
''' 上下文校验器接口 - 在发送AI请求前校验上下文完整性
''' </summary>
Public Interface IContextChecker

    ''' <summary>
    ''' 校验执行上下文是否满足发送条件
    ''' </summary>
    ''' <param name="context">执行上下文</param>
    ''' <returns>校验结果</returns>
    Function CheckAsync(context As ExecutionContext) As Task(Of ContextCheckResult)

End Interface
