' ShareRibbon\Agent\TaskPackageModels.vb
' E6-02 任务包合同的字段模型：与 ShareRibbon/TaskPackages/*.json 的 schema 一一对应。
' 字段词表（允许哪些键）由 TaskPackageCatalog 持有，门禁脚本从那里反查，本文件不另立名单。

Namespace Agent

    ''' <summary>一条验证断言：run 结束时用 AgentRunChangeLog 上的事实核对"是不是真的做成了"。</summary>
    Public Class TaskPackageAssertion

        Public Property Fact As String = ""

        Public Property Op As String = ""

        Public Property Value As Integer

        Public Property Why As String = ""
    End Class

    Public Class TaskPackageVerifier

        Public Property Assertions As List(Of TaskPackageAssertion) = New List(Of TaskPackageAssertion)()

        Public Property OnFail As String = ""
    End Class

    Public Class TaskPackagePreviewPolicy

        Public Property Mode As String = ""

        Public Property RequireBeforeWrite As Boolean

        Public Property Note As String = ""
    End Class

    Public Class TaskPackageExplanation

        Public Property WhatChanged As String = ""

        Public Property WhyHint As String = ""

        Public Property UndoHint As String = ""
    End Class

    Public Class TaskPackageRecovery

        Public Property Strategy As String = ""

        Public Property Notes As String = ""
    End Class

    Public Class TaskPackageExistingImplementation

        Public Property Reuse As List(Of String) = New List(Of String)()

        Public Property Gap As String = ""
    End Class

    ''' <summary>一个任务包的合同。字段与 JSON 一一对应，不接受 JSON 里没有的东西。</summary>
    Public Class TaskPackage

        Public Property Id As String = ""

        Public Property Title As String = ""

        Public Property Application As String = ""

        Public Property Skill As String = ""

        Public Property Deliverable As String = ""

        Public Property EntryActions As List(Of String) = New List(Of String)()

        Public Property CapabilityIds As List(Of String) = New List(Of String)()

        Public Property SuccessCriteria As List(Of String) = New List(Of String)()

        Public Property GoldenCases As List(Of String) = New List(Of String)()

        Public Property PreviewPolicy As TaskPackagePreviewPolicy = New TaskPackagePreviewPolicy()

        Public Property Verifier As TaskPackageVerifier = New TaskPackageVerifier()

        Public Property Explanation As TaskPackageExplanation = New TaskPackageExplanation()

        Public Property Recovery As TaskPackageRecovery = New TaskPackageRecovery()

        Public Property ExistingImplementation As TaskPackageExistingImplementation = New TaskPackageExistingImplementation()

        ''' <summary>合同文件的相对路径，只用于把校验错误指回磁盘上那一份。</summary>
        Public Property SourceFile As String = ""

        ''' <summary>合同所在目录；同一份合同从源码目录和构建输出目录各读到一次时按目录判重。</summary>
        Public Property SourceDirectory As String = ""
    End Class

End Namespace
