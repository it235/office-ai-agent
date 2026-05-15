' ShareRibbon\Loop\Interfaces\IResultVerifier.vb
' 结果验证器接口

Imports System.Threading.Tasks

''' <summary>
''' 结果验证器接口 - 验证指令执行后的结果是否符合预期
''' </summary>
Public Interface IResultVerifier

    ''' <summary>
    ''' 验证执行结果
    ''' </summary>
    ''' <param name="executionResult">执行结果</param>
    ''' <param name="context">执行上下文</param>
    ''' <returns>验证结果</returns>
    Function VerifyAsync(executionResult As ExecutionResult, context As ExecutionContext) As Task(Of VerificationResult)

End Interface
