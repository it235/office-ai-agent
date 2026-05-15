' ShareRibbon\Loop\LoopPhase.vb
' 自检Loop阶段枚举

''' <summary>
''' 自检Loop阶段枚举
''' </summary>
Public Enum LoopPhase
    Idle
    PreSendCheck
    Planning
    Generating
    Validating
    Executing
    Verifying
    Correcting
    Completed
    Failed
End Enum
