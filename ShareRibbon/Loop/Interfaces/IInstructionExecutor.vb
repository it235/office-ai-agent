' ShareRibbon\Loop\Interfaces\IInstructionExecutor.vb
' 指令执行器接口

Imports System.Threading.Tasks

''' <summary>
''' 指令执行器接口 - 将DSL指令转换为Office DOM操作并执行
''' </summary>
Public Interface IInstructionExecutor

    ''' <summary>
    ''' 执行指令列表
    ''' </summary>
    ''' <param name="instructions">要执行的指令列表</param>
    ''' <param name="context">执行上下文</param>
    ''' <returns>执行结果</returns>
    Function ExecuteAsync(instructions As List(Of Instruction), context As ExecutionContext) As Task(Of ExecutionResult)

End Interface
