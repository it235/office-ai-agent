' ShareRibbon\Loop\Interfaces\IInstructionPlanner.vb
' 指令规划器接口

Imports System.Threading.Tasks

''' <summary>
''' 指令规划器接口 - 将用户意图规划为操作步骤
''' </summary>
Public Interface IInstructionPlanner

    ''' <summary>
    ''' 根据上下文规划操作步骤
    ''' </summary>
    ''' <param name="context">执行上下文</param>
    ''' <returns>规划结果</returns>
    Function PlanAsync(context As ExecutionContext) As Task(Of PlanningResult)

End Interface
