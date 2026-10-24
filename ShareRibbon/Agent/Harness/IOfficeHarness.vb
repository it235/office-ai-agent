Imports System.Threading
Imports System.Threading.Tasks

Namespace Agent.Harness

    Public Interface IOfficeHarness
        Event PhaseChanged As EventHandler(Of HarnessPhaseChangedEventArgs)
        Event StepChanged As EventHandler(Of HarnessStepChangedEventArgs)
        Event ContextReady As EventHandler(Of HarnessContextEventArgs)

        Function RunAsync(turn As UserTurn, cancellationToken As CancellationToken) As Task(Of HarnessRunResult)
        Function ApproveAsync(runId As String, approved As Boolean, cancellationToken As CancellationToken) As Task(Of HarnessRunResult)
        Function DecideAsync(runId As String, decision As ApprovalDecision, cancellationToken As CancellationToken) As Task(Of HarnessRunResult)
        Function ResumeAsync(runId As String, cancellationToken As CancellationToken) As Task(Of HarnessRunResult)
        Function CancelAsync(runId As String, cancellationToken As CancellationToken) As Task(Of HarnessRunResult)
    End Interface

End Namespace
