' ShareRibbon\Loop\Interfaces\IInstructionValidator.vb
' 指令校验器接口

Imports System.Threading.Tasks

''' <summary>
''' 指令校验器接口 - 校验AI生成的指令格式和语义
''' </summary>
Public Interface IInstructionValidator

    ''' <summary>
    ''' 校验AI响应内容
    ''' </summary>
    ''' <param name="aiResponse">AI原始响应</param>
    ''' <param name="expectedFormat">期望的指令格式</param>
    ''' <param name="context">执行上下文</param>
    ''' <returns>校验结果</returns>
    Function ValidateAsync(aiResponse As String, expectedFormat As InstructionFormat, context As ExecutionContext) As Task(Of ValidationResult)

End Interface
