Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports Newtonsoft.Json.Linq
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon
Imports ShareRibbon.Agent
Imports ShareRibbon.Agent.OfficeOperations

Namespace OfficeRuntime

    Friend NotInheritable Class PowerPointOperationExecutor
        Private Sub New()
        End Sub

        Public Shared Function Execute(params As JObject) As ToolResult
            Const toolId As String = "OfficeObjectOperation"
            Dim batch As OfficeOperationBatch = Nothing
            Try
                Dim batchToken = params?("batch")
                If batchToken Is Nothing OrElse batchToken.Type <> JTokenType.Object Then
                    Return ToolResult.Failed(toolId,
                                             "OfficeObjectOperation requires a batch object",
                                             errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                             userMessage:="声明式 PowerPoint 操作格式无效",
                                             recoverable:=True)
                End If
                batch = batchToken.ToObject(Of OfficeOperationBatch)()
                Dim validation = OfficeOperationValidation.ValidateBatch(batch)
                If Not validation.IsValid Then
                    Return ToolResult.Failed(toolId,
                                             validation.ToErrorMessage(),
                                             errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                             userMessage:="声明式 PowerPoint 操作未通过合同校验",
                                             recoverable:=True)
                End If
                If Not String.Equals(OfficeObjectRef.NormalizeAppType(batch.AppType), "PowerPoint", StringComparison.OrdinalIgnoreCase) Then
                    Return ToolResult.Failed(toolId,
                                             $"Unsupported batch appType {batch.AppType}",
                                             errorCode:=ExceptionClassifier.CodeHostUnsupported,
                                             userMessage:="当前 PowerPoint 宿主不能执行其他 Office 应用的操作",
                                             recoverable:=False)
                End If

                Dim beforeState = PowerPointOperationObserver.CaptureState(batch)
                Dim operationResults As New JArray()
                Dim targetRefs As New List(Of String)()
                Dim warnings As New List(Of String)()
                Dim succeededCount As Integer = 0

                For Each operation In batch.Operations
                    Try
                        Dim execution = ExecuteOperation(operation)
                        succeededCount += 1
                        targetRefs.Add(operation.TargetRef)
                        If Not String.IsNullOrWhiteSpace(execution.ResultRef) Then targetRefs.Add(execution.ResultRef)
                        operationResults.Add(New JObject From {
                            {"id", operation.Id},
                            {"status", "succeeded"},
                            {"memberId", operation.MemberId},
                            {"targetRef", operation.TargetRef},
                            {"resultRef", If(execution.ResultRef, "")},
                            {"data", If(execution.Data, JValue.CreateNull())}
                        })
                    Catch ex As PowerPointOperationException
                        operationResults.Add(New JObject From {
                            {"id", operation.Id},
                            {"status", "failed"},
                            {"memberId", operation.MemberId},
                            {"targetRef", operation.TargetRef},
                            {"errorCode", ex.ErrorCode},
                            {"message", ex.Message}
                        })
                        If batch.Atomic AndAlso succeededCount > 0 Then
                            warnings.Add("Atomic batch partially applied; automatic rollback was not guaranteed by PowerPoint COM")
                        End If
                        Dim afterFailure = PowerPointOperationObserver.CaptureState(batch, targetRefs)
                        Dim failureObservation = PowerPointOperationObserver.BuildObservation(batch,
                                                                                             operationResults,
                                                                                             beforeState,
                                                                                             afterFailure,
                                                                                             targetRefs,
                                                                                             warnings)
                        Dim code = If(succeededCount > 0, ExceptionClassifier.CodePartialApply, ex.ErrorCode)
                        Return ToolResult.Failed(toolId,
                                                 ex.Message,
                                                 data:=New JObject From {{"targetRefs", JArray.FromObject(targetRefs)}},
                                                 errorCode:=code,
                                                 userMessage:="PowerPoint 声明式操作未全部完成",
                                                 recoverable:=ex.Recoverable,
                                                 observation:=failureObservation)
                    Catch ex As Exception
                        Dim baseException = If(TypeOf ex Is TargetInvocationException AndAlso ex.InnerException IsNot Nothing,
                                               ex.InnerException,
                                               ex)
                        Dim classified = ExceptionClassifier.Classify(baseException)
                        operationResults.Add(New JObject From {
                            {"id", operation.Id},
                            {"status", "failed"},
                            {"memberId", operation.MemberId},
                            {"targetRef", operation.TargetRef},
                            {"errorCode", classified.ErrorCode},
                            {"message", classified.UserMessage}
                        })
                        If batch.Atomic AndAlso succeededCount > 0 Then warnings.Add("Atomic batch partially applied")
                        Dim afterFailure = PowerPointOperationObserver.CaptureState(batch, targetRefs)
                        Dim failureObservation = PowerPointOperationObserver.BuildObservation(batch,
                                                                                             operationResults,
                                                                                             beforeState,
                                                                                             afterFailure,
                                                                                             targetRefs,
                                                                                             warnings)
                        Return ToolResult.Failed(toolId,
                                                 classified.DebugDetail,
                                                 data:=New JObject From {{"targetRefs", JArray.FromObject(targetRefs)}},
                                                 errorCode:=If(succeededCount > 0, ExceptionClassifier.CodePartialApply, classified.ErrorCode),
                                                 userMessage:=classified.UserMessage,
                                                 recoverable:=classified.Recoverable,
                                                 observation:=failureObservation)
                    End Try
                Next

                Dim afterState = PowerPointOperationObserver.CaptureState(batch, targetRefs)
                Dim observation = PowerPointOperationObserver.BuildObservation(batch,
                                                                               operationResults,
                                                                               beforeState,
                                                                               afterState,
                                                                               targetRefs,
                                                                               warnings)
                Dim verification = PowerPointOperationObserver.VerifyExpectedEffects(batch, operationResults, afterState)
                observation("verification") = verification
                Dim data As New JObject From {
                    {"schemaVersion", batch.SchemaVersion},
                    {"targetRefs", JArray.FromObject(targetRefs.Distinct(StringComparer.OrdinalIgnoreCase))},
                    {"operations", operationResults.DeepClone()}
                }
                If PowerPointOperationObserver.HasRequiredVerificationFailure(verification) Then
                    Return ToolResult.Failed(toolId,
                                             "PowerPoint operation verification failed",
                                             data:=data,
                                             errorCode:=ExceptionClassifier.CodeVerifyFailed,
                                             userMessage:="PowerPoint 已执行操作，但实际结果未满足成功标准",
                                             recoverable:=True,
                                             observation:=observation)
                End If
                Return ToolResult.Succeed(toolId, observation("summary").ToString(), data:=data, observation:=observation)
            Catch ex As PowerPointOperationException
                Return ToolResult.Failed(toolId,
                                         ex.Message,
                                         errorCode:=ex.ErrorCode,
                                         userMessage:="PowerPoint 对象操作失败",
                                         recoverable:=ex.Recoverable)
            Catch ex As Exception
                Return ToolResult.FromException(toolId, ex)
            End Try
        End Function

        Private Shared Function ExecuteOperation(operation As OfficeOperation) As OperationExecutionResult
            Dim capability As OfficeCapabilityMember = Nothing
            Dim reflectedMember As MemberInfo = Nothing
            If Not PowerPointApiCatalogProvider.TryGetMemberBinding(operation.MemberId, capability, reflectedMember) Then
                Throw New PowerPointOperationException(ExceptionClassifier.CodeCapabilityNotFound,
                                                       $"Catalog member was not found: {operation.MemberId}")
            End If
            If Not capability.Executable Then
                Throw New PowerPointOperationException(ExceptionClassifier.CodeMemberNotExecutable,
                                                       If(capability.UnsupportedReason, $"Member is not executable: {operation.MemberId}"),
                                                       recoverable:=False)
            End If

            Using resolved = PowerPointObjectResolver.Resolve(operation.TargetRef)
                EnsureTargetCompatibility(resolved, reflectedMember)
                Dim boundComObjects As New List(Of Object)()
                Dim returnValue As Object = Nothing
                Dim isReadOnlyMember = String.Equals(capability.RiskLevel, "safe", StringComparison.OrdinalIgnoreCase)
                Try
                    If TypeOf reflectedMember Is PropertyInfo Then
                        returnValue = ExecuteProperty(operation,
                                                      resolved.Value,
                                                      DirectCast(reflectedMember, PropertyInfo),
                                                      boundComObjects)
                    ElseIf TypeOf reflectedMember Is MethodInfo Then
                        returnValue = ExecuteMethod(operation,
                                                    resolved.Value,
                                                    DirectCast(reflectedMember, MethodInfo),
                                                    boundComObjects,
                                                    isReadOnlyMember)
                    Else
                        Throw New PowerPointOperationException(ExceptionClassifier.CodeMemberNotExecutable,
                                                               "Only catalog properties and methods are executable",
                                                               recoverable:=False)
                    End If

                    Dim resultRef = BuildResultRef(operation, resolved.Value, reflectedMember, returnValue)
                    Dim scalarData = ConvertReturnValue(returnValue)
                    Return New OperationExecutionResult With {
                        .ResultRef = resultRef,
                        .Data = scalarData
                    }
                Finally
                    If returnValue IsNot Nothing AndAlso
                       Marshal.IsComObject(returnValue) AndAlso
                       Not Object.ReferenceEquals(returnValue, resolved.Value) Then
                        ReleaseCom(returnValue)
                    End If
                    For index = boundComObjects.Count - 1 To 0 Step -1
                        ReleaseCom(boundComObjects(index))
                    Next
                End Try
            End Using
        End Function

        Private Shared Function ExecuteProperty(operation As OfficeOperation,
                                                target As Object,
                                                prop As PropertyInfo,
                                                boundComObjects As List(Of Object)) As Object
            Dim action = operation.Action.Trim().ToLowerInvariant()
            Select Case action
                Case "get", "collection_item"
                    If Not prop.CanRead Then ThrowMemberAction(operation, "Property is not readable")
                    Dim indexArgs = BindParameters(prop.GetIndexParameters(), operation.Arguments, boundComObjects)
                    Return prop.GetValue(target, indexArgs)

                Case "set"
                    If Not prop.CanWrite Then ThrowMemberAction(operation, "Property is not writable")
                    Dim valueToken = GetArgument(operation.Arguments, "value")
                    If valueToken Is Nothing Then
                        Throw New PowerPointOperationException(ExceptionClassifier.CodeOperationSchemaInvalid,
                                                               $"Operation {operation.Id} requires arguments.value")
                    End If
                    Dim value = ConvertArgument(valueToken, prop.PropertyType, boundComObjects)
                    Dim indexArgs = BindParameters(prop.GetIndexParameters(), operation.Arguments, boundComObjects)
                    prop.SetValue(target, value, indexArgs)
                    Return Nothing

                Case Else
                    ThrowMemberAction(operation, $"Action {operation.Action} cannot execute a property")
                    Return Nothing
            End Select
        End Function

        Private Shared Function ExecuteMethod(operation As OfficeOperation,
                                              target As Object,
                                              method As MethodInfo,
                                              boundComObjects As List(Of Object),
                                              isReadOnlyMember As Boolean) As Object
            Dim action = operation.Action.Trim().ToLowerInvariant()
            Select Case action
                Case "get", "collection_item"
                    ' 读语义只能用于只读访问器：绝不允许通过 get 触发有副作用的方法（如 Save/Export/Delete）。
                    If Not isReadOnlyMember Then
                        ThrowMemberAction(operation,
                                          "Action " & operation.Action & " cannot invoke a side-effecting method; use invoke")
                    End If
                Case "invoke", "create", "delete"
                    ' 写语义：可执行性与 SafetyGate 已在主链上游裁决。
                Case Else
                    ThrowMemberAction(operation, $"Action {operation.Action} cannot execute a method")
            End Select
            Dim args = BindParameters(method.GetParameters(), operation.Arguments, boundComObjects)
            Return method.Invoke(target, args)
        End Function

        Private Shared Function BindParameters(parameters As ParameterInfo(),
                                               arguments As JObject,
                                               boundComObjects As List(Of Object)) As Object()
            If parameters Is Nothing OrElse parameters.Length = 0 Then Return Array.Empty(Of Object)()
            Dim result(parameters.Length - 1) As Object
            For index = 0 To parameters.Length - 1
                Dim parameter = parameters(index)
                Dim token = GetArgument(arguments, parameter.Name)
                If token Is Nothing Then
                    If parameter.IsOptional Then
                        result(index) = Type.Missing
                        Continue For
                    End If
                    Throw New PowerPointOperationException(ExceptionClassifier.CodeOperationSchemaInvalid,
                                                           $"Missing required argument '{parameter.Name}'")
                End If
                result(index) = ConvertArgument(token, parameter.ParameterType, boundComObjects)
            Next
            Return result
        End Function

        Private Shared Function ConvertArgument(token As JToken,
                                                expectedType As Type,
                                                boundComObjects As List(Of Object)) As Object
            If expectedType.IsByRef Then expectedType = expectedType.GetElementType()
            If token Is Nothing OrElse token.Type = JTokenType.Null Then
                If Not expectedType.IsValueType OrElse Nullable.GetUnderlyingType(expectedType) IsNot Nothing Then Return Nothing
                Throw New PowerPointOperationException(ExceptionClassifier.CodeOperationSchemaInvalid,
                                                       $"Null is not valid for {expectedType.FullName}")
            End If

            If String.Equals(expectedType.FullName, "Microsoft.Office.Core.SmartArtLayout", StringComparison.Ordinal) Then
                Dim layout = ResolveSmartArtLayout(token)
                boundComObjects.Add(layout)
                Return layout
            End If
            If expectedType Is GetType(String) Then Return token.ToString()
            If expectedType Is GetType(Boolean) Then Return token.Value(Of Boolean)()
            If expectedType Is GetType(Integer) Then Return token.Value(Of Integer)()
            If expectedType Is GetType(Long) Then Return token.Value(Of Long)()
            If expectedType Is GetType(Short) Then Return token.Value(Of Short)()
            If expectedType Is GetType(Single) Then Return Single.Parse(token.ToString(), CultureInfo.InvariantCulture)
            If expectedType Is GetType(Double) Then Return Double.Parse(token.ToString(), CultureInfo.InvariantCulture)
            If expectedType Is GetType(Decimal) Then Return Decimal.Parse(token.ToString(), CultureInfo.InvariantCulture)
            If expectedType.IsEnum Then
                If token.Type = JTokenType.Integer Then Return [Enum].ToObject(expectedType, token.Value(Of Integer)())
                Return [Enum].Parse(expectedType, token.ToString(), ignoreCase:=True)
            End If
            If expectedType Is GetType(Object) Then Return token.ToObject(Of Object)()

            Throw New PowerPointOperationException(ExceptionClassifier.CodeOperationSchemaInvalid,
                                                   $"Unsupported argument type {expectedType.FullName}")
        End Function

        Private Shared Function ResolveSmartArtLayout(token As JToken) As SmartArtLayout
            Dim key As Object = Nothing
            If token.Type = JTokenType.Integer Then
                key = token.Value(Of Integer)()
            ElseIf token.Type = JTokenType.String Then
                key = token.ToString()
            ElseIf token.Type = JTokenType.Object Then
                Dim obj = DirectCast(token, JObject)
                Dim keyToken = obj.GetValue("index", StringComparison.OrdinalIgnoreCase)
                If keyToken Is Nothing Then keyToken = obj.GetValue("layoutId", StringComparison.OrdinalIgnoreCase)
                If keyToken Is Nothing Then keyToken = obj.GetValue("id", StringComparison.OrdinalIgnoreCase)
                If keyToken Is Nothing Then keyToken = obj.GetValue("name", StringComparison.OrdinalIgnoreCase)
                If keyToken IsNot Nothing Then key = If(keyToken.Type = JTokenType.Integer, CType(keyToken.Value(Of Integer)(), Object), keyToken.ToString())
            End If
            If key Is Nothing Then
                Throw New PowerPointOperationException(ExceptionClassifier.CodeOperationSchemaInvalid,
                                                       "SmartArtLayout requires an index, layoutId, id, or name")
            End If

            Dim layouts As SmartArtLayouts = Nothing
            Try
                layouts = Globals.ThisAddIn.Application.SmartArtLayouts
                If layouts Is Nothing Then Throw New PowerPointOperationException(ExceptionClassifier.CodeObjectNotFound, "SmartArt layouts are unavailable")
                Dim semanticKey = NormalizeLayoutKey(key.ToString())
                If semanticKey = "basicprocess" OrElse semanticKey = "process" OrElse
                   semanticKey = "threestage" OrElse semanticKey = "三阶段" OrElse semanticKey = "基本流程" Then
                    key = "urn:microsoft.com/office/officeart/2005/8/layout/process1"
                End If
                Try
                    Dim direct = layouts.Item(key)
                    If direct IsNot Nothing Then Return direct
                Catch
                End Try

                Dim search = key.ToString()
                Dim normalizedSearch = NormalizeLayoutKey(search)
                If String.IsNullOrWhiteSpace(normalizedSearch) Then
                    Throw New PowerPointOperationException(ExceptionClassifier.CodeOperationSchemaInvalid,
                                                           "SmartArt layout key cannot be empty")
                End If
                For index = 1 To layouts.Count
                    Dim candidate As SmartArtLayout = Nothing
                    Try
                        candidate = layouts.Item(index)
                        If String.Equals(candidate.Id, search, StringComparison.OrdinalIgnoreCase) OrElse
                           String.Equals(candidate.Name, search, StringComparison.OrdinalIgnoreCase) OrElse
                           NormalizeLayoutKey(candidate.Name).IndexOf(normalizedSearch, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Return candidate
                        End If
                    Catch
                    End Try
                    ReleaseCom(candidate)
                Next
            Finally
                ReleaseCom(layouts)
            End Try
            Throw New PowerPointOperationException(ExceptionClassifier.CodeObjectNotFound,
                                                   $"SmartArt layout '{key}' was not found")
        End Function

        Private Shared Function NormalizeLayoutKey(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return ""
            Dim builder As New Text.StringBuilder()
            For Each ch In value.Trim().ToLowerInvariant()
                If Char.IsLetterOrDigit(ch) Then builder.Append(ch)
            Next
            Return builder.ToString()
        End Function

        Private Shared Sub EnsureTargetCompatibility(resolved As ResolvedOfficeObject, member As MemberInfo)
            If resolved Is Nothing OrElse resolved.Value Is Nothing OrElse member?.DeclaringType Is Nothing Then
                Throw New PowerPointOperationException(ExceptionClassifier.CodeObjectTypeMismatch, "Operation target is unavailable")
            End If
            If member.DeclaringType.IsInstanceOfType(resolved.Value) Then Return
            Dim expected = member.DeclaringType.Name.TrimStart("_"c)
            If String.Equals(expected, resolved.ObjectKind, StringComparison.OrdinalIgnoreCase) Then Return
            Throw New PowerPointOperationException(ExceptionClassifier.CodeObjectTypeMismatch,
                                                   $"Member {member.Name} requires {expected}, target is {resolved.ObjectKind}")
        End Sub

        Private Shared Function BuildResultRef(operation As OfficeOperation,
                                               target As Object,
                                               member As MemberInfo,
                                               returnValue As Object) As String
            If returnValue Is Nothing OrElse Not Marshal.IsComObject(returnValue) Then Return ""

            If TypeOf returnValue Is PowerPoint.Shape AndAlso TypeOf target Is PowerPoint.Shapes Then
                Dim shapes = DirectCast(target, PowerPoint.Shapes)
                Dim resultShape = DirectCast(returnValue, PowerPoint.Shape)
                For index = 1 To shapes.Count
                    Dim candidate As PowerPoint.Shape = Nothing
                    Try
                        candidate = shapes(index)
                        If candidate.Id = resultShape.Id Then Return operation.TargetRef.TrimEnd("/"c) & "/" & index
                    Finally
                        If Not Object.ReferenceEquals(candidate, returnValue) Then ReleaseCom(candidate)
                    End Try
                Next
            End If
            If TypeOf returnValue Is SmartArtNode AndAlso TypeOf target Is SmartArtNodes Then
                Dim nodes = DirectCast(target, SmartArtNodes)
                For index = 1 To nodes.Count
                    Dim candidate As SmartArtNode = Nothing
                    Try
                        candidate = nodes.Item(index)
                        If ComIdentityEquals(candidate, returnValue) Then Return operation.TargetRef.TrimEnd("/"c) & "/" & index
                    Finally
                        If Not Object.ReferenceEquals(candidate, returnValue) Then ReleaseCom(candidate)
                    End Try
                Next
            End If

            Select Case member.Name.ToLowerInvariant()
                Case "smartart"
                    Return operation.TargetRef.TrimEnd("/"c) & "/smartart"
                Case "nodes", "allnodes"
                    Return operation.TargetRef.TrimEnd("/"c) & "/nodes"
                Case "textframe"
                    Return operation.TargetRef.TrimEnd("/"c) & "/textframe"
                Case "textframe2"
                    Return operation.TargetRef.TrimEnd("/"c) & "/textframe2"
                Case "textrange"
                    Return operation.TargetRef.TrimEnd("/"c) & "/textrange"
            End Select
            Return operation.TargetRef
        End Function

        Private Shared Function ComIdentityEquals(left As Object, right As Object) As Boolean
            If left Is Nothing OrElse right Is Nothing Then Return False
            If Object.ReferenceEquals(left, right) Then Return True
            If Not Marshal.IsComObject(left) OrElse Not Marshal.IsComObject(right) Then Return False
            Dim leftIdentity As IntPtr = IntPtr.Zero
            Dim rightIdentity As IntPtr = IntPtr.Zero
            Try
                leftIdentity = Marshal.GetIUnknownForObject(left)
                rightIdentity = Marshal.GetIUnknownForObject(right)
                Return leftIdentity = rightIdentity
            Finally
                If leftIdentity <> IntPtr.Zero Then Marshal.Release(leftIdentity)
                If rightIdentity <> IntPtr.Zero Then Marshal.Release(rightIdentity)
            End Try
        End Function

        Private Shared Function ConvertReturnValue(value As Object) As JToken
            If value Is Nothing OrElse Marshal.IsComObject(value) Then Return JValue.CreateNull()
            Try
                Return JToken.FromObject(value)
            Catch
                Return New JValue(value.ToString())
            End Try
        End Function

        Private Shared Function GetArgument(arguments As JObject, name As String) As JToken
            If arguments Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then Return Nothing
            Return arguments.GetValue(name, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Sub ThrowMemberAction(operation As OfficeOperation, message As String)
            Throw New PowerPointOperationException(ExceptionClassifier.CodeOperationSchemaInvalid,
                                                   $"Operation {operation.Id}: {message}")
        End Sub

        Private Shared Sub ReleaseCom(value As Object)
            If value Is Nothing Then Return
            Try
                If Marshal.IsComObject(value) Then Marshal.ReleaseComObject(value)
            Catch
            End Try
        End Sub

        Private Class OperationExecutionResult
            Public Property ResultRef As String
            Public Property Data As JToken
        End Class
    End Class

End Namespace
