' ShareRibbon\Common\ExceptionClassifier.vb
' Classifies exceptions for structured ToolResult / OperationResult (P0-4).

Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports Newtonsoft.Json

''' <summary>
''' Maps exceptions to stable error codes and user-facing messages.
''' Prefer catching specific types, then fall back to Classify(ex).
''' </summary>
Public NotInheritable Class ExceptionClassifier
    Private Sub New()
    End Sub

    Public Const CodeUnknown As String = "UNKNOWN"
    Public Const CodeCom As String = "COM_ERROR"
    Public Const CodeNetwork As String = "NETWORK_ERROR"
    Public Const CodeTimeout As String = "TIMEOUT"
    Public Const CodeJson As String = "JSON_ERROR"
    Public Const CodeArgument As String = "ARGUMENT_ERROR"
    Public Const CodeNotFound As String = "NOT_FOUND"
    Public Const CodeToolNotAllowed As String = "TOOL_NOT_ALLOWED"
    Public Const CodeHostUnsupported As String = "HOST_UNSUPPORTED"
    Public Const CodeSafetyBlocked As String = "SAFETY_BLOCKED"
    Public Const CodeSafetyNeedsApproval As String = "SAFETY_NEEDS_APPROVAL"
    Public Const CodeApprovalTimeout As String = "APPROVAL_TIMEOUT"
    Public Const CodeVbaDisabled As String = "VBA_DISABLED"
    Public Const CodeCancelled As String = "CANCELLED"
    Public Const CodeIo As String = "IO_ERROR"
    Public Const CodeVerifyFailed As String = "VERIFY_FAILED"
    Public Const CodeObservationFailed As String = "OBSERVATION_FAILED"
    Public Const CodePartialApply As String = "PARTIAL_APPLY"
    Public Const CodeCapabilityNotFound As String = "CAPABILITY_NOT_FOUND"
    Public Const CodeMemberNotExecutable As String = "MEMBER_NOT_EXECUTABLE"
    Public Const CodeOperationSchemaInvalid As String = "OPERATION_SCHEMA_INVALID"
    Public Const CodeObjectRefInvalid As String = "OBJECT_REF_INVALID"
    Public Const CodeObjectNotFound As String = "OBJECT_NOT_FOUND"
    Public Const CodeObjectTypeMismatch As String = "OBJECT_TYPE_MISMATCH"
    Public Const CodeDocMissing As String = "DOC_MISSING"

    Public Class ClassifiedError
        Public Property ErrorCode As String = CodeUnknown
        Public Property UserMessage As String = ""
        Public Property DebugDetail As String = ""
        Public Property Recoverable As Boolean = True
    End Class

    Public Shared Function Classify(ex As Exception) As ClassifiedError
        Dim result As New ClassifiedError()
        If ex Is Nothing Then
            result.UserMessage = "发生未知错误"
            result.DebugDetail = "Exception was Nothing"
            result.Recoverable = False
            Return result
        End If

        Dim baseEx = If(TypeOf ex Is AggregateException, ex.GetBaseException(), ex)
        result.DebugDetail = $"{baseEx.GetType().FullName}: {AppLogger.Redact(baseEx.Message)}"

        If TypeOf baseEx Is TaskCanceledException OrElse TypeOf baseEx Is OperationCanceledException Then
            result.ErrorCode = CodeTimeout
            result.UserMessage = "操作超时或已取消，请重试"
            result.Recoverable = True
            Return result
        End If

        If TypeOf baseEx Is HttpRequestException OrElse TypeOf baseEx Is WebException Then
            result.ErrorCode = CodeNetwork
            result.UserMessage = "网络请求失败，请检查网络与 API 配置后重试"
            result.Recoverable = True
            Return result
        End If

        If TypeOf baseEx Is JsonException OrElse
           TypeOf baseEx Is JsonReaderException OrElse
           TypeOf baseEx Is JsonSerializationException OrElse
           (baseEx.GetType().Name.IndexOf("Json", StringComparison.OrdinalIgnoreCase) >= 0) Then
            result.ErrorCode = CodeJson
            result.UserMessage = "数据解析失败，请重试或调整指令"
            result.Recoverable = True
            Return result
        End If

        Dim isComInterfaceCastFailure = TypeOf baseEx Is InvalidCastException AndAlso
            IsComInterfaceUnavailableMessage(baseEx.Message)
        If TypeOf baseEx Is COMException OrElse TypeOf baseEx Is InvalidComObjectException OrElse isComInterfaceCastFailure Then
            result.ErrorCode = CodeCom
            result.UserMessage = If(isComInterfaceCastFailure,
                                    "当前 Office/WPS 宿主不支持所需的 COM 接口，请更新宿主或使用兼容路径",
                                    "Office 文档操作失败，请确认文档未锁定且选区有效")
            ' QueryInterface/E_NOINTERFACE is deterministic for the current host.
            ' Retrying the same tool with AI-generated parameters cannot add the missing interface.
            result.Recoverable = Not isComInterfaceCastFailure
            Return result
        End If

        ' RPC_E_WRONG_THREAD / COM often surfaces as COMException; also match message.
        Dim msg = If(baseEx.Message, "")
        If msg.IndexOf("RPC_E_WRONG_THREAD", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           msg.IndexOf("被调用的对象已与其客户端断开连接", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           msg.IndexOf("COM", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso msg.IndexOf("线程", StringComparison.OrdinalIgnoreCase) >= 0 Then
            result.ErrorCode = CodeCom
            result.UserMessage = "Office 对象跨线程访问失败，请重试该操作"
            result.Recoverable = True
            Return result
        End If

        If TypeOf baseEx Is ArgumentException OrElse TypeOf baseEx Is ArgumentNullException OrElse TypeOf baseEx Is ArgumentOutOfRangeException Then
            result.ErrorCode = CodeArgument
            result.UserMessage = "参数无效，请检查指令参数后重试"
            result.Recoverable = True
            Return result
        End If

        If TypeOf baseEx Is FileNotFoundException OrElse TypeOf baseEx Is DirectoryNotFoundException Then
            result.ErrorCode = CodeNotFound
            result.UserMessage = "找不到所需文件或目录"
            result.Recoverable = False
            Return result
        End If

        If TypeOf baseEx Is IOException Then
            result.ErrorCode = CodeIo
            result.UserMessage = "文件读写失败，请检查路径与权限"
            result.Recoverable = True
            Return result
        End If

        result.ErrorCode = CodeUnknown
        result.UserMessage = "操作失败，请重试；若反复出现请查看日志"
        result.Recoverable = True
        Return result
    End Function

    Public Shared Function IsComInterfaceUnavailableMessage(message As String) As Boolean
        Dim value = If(message, "")
        Return value.IndexOf("QueryInterface", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               value.IndexOf("E_NOINTERFACE", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               value.IndexOf("不支持此接口", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    Public Shared Function ToUserMessage(ex As Exception, Optional fallback As String = Nothing) As String
        Dim c = Classify(ex)
        If Not String.IsNullOrWhiteSpace(fallback) Then Return fallback
        Return c.UserMessage
    End Function
End Class
