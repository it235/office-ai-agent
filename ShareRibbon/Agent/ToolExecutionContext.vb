Imports System.Collections.Generic
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' Carries per-run tool visibility and execution policy.
    ''' H1: appType + selected Skill allowed-tools are enforced at execution time.
    ''' </summary>
    Public Class ToolExecutionContext
        Public Property AppType As String
        Public Property RunId As String
        Public Property CorrelationId As String
        Public Property PrimarySkillName As String
        Public Property AllowedTools As HashSet(Of String)
        Public Property ApprovedTools As HashSet(Of String)
        Public Property ApprovedToolCalls As HashSet(Of String)
        ''' <summary>本会话内放行的工具；与一次性审批不同，命中后不消费。</summary>
        Public Property SessionApprovedTools As HashSet(Of String)
        Public Property EnforceAllowedTools As Boolean = True

        Public Shared Function FromSession(session As AgentSession, skill As AgentSkill) As ToolExecutionContext
            Dim context As New ToolExecutionContext With {
                .AppType = If(session?.AppType, ""),
                .RunId = If(session?.Id, ""),
                .CorrelationId = If(session?.Id, ""),
                .PrimarySkillName = If(skill?.Name, ""),
                .AllowedTools = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase),
                .ApprovedTools = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase),
                .ApprovedToolCalls = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase),
                .SessionApprovedTools = If(session?.SessionApprovedTools IsNot Nothing,
                                            session.SessionApprovedTools,
                                            New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)),
                .EnforceAllowedTools = False
            }

            If skill IsNot Nothing AndAlso skill.RequiredTools IsNot Nothing Then
                For Each toolId In skill.RequiredTools
                    If Not String.IsNullOrWhiteSpace(toolId) Then
                        context.AllowedTools.Add(toolId.Trim())
                    End If
                Next

                context.EnforceAllowedTools = context.AllowedTools.Count > 0
            End If

            Return context
        End Function

        Public Function HasPrimarySkillGate() As Boolean
            Return EnforceAllowedTools AndAlso
                   AllowedTools IsNot Nothing AndAlso
                   AllowedTools.Count > 0
        End Function

        Public Function IsToolAllowed(toolId As String) As Boolean
            If Not HasPrimarySkillGate() Then Return True
            If String.IsNullOrWhiteSpace(toolId) Then Return False
            Return AllowedTools.Contains(toolId.Trim())
        End Function

        Public Function AllowedToolsText() As String
            If AllowedTools Is Nothing OrElse AllowedTools.Count = 0 Then Return ""
            Return String.Join(", ", AllowedTools.OrderBy(Function(id) id))
        End Function

        Public Sub ApproveTool(toolId As String, Optional params As JObject = Nothing)
            If String.IsNullOrWhiteSpace(toolId) Then Return
            If params IsNot Nothing Then
                If ApprovedToolCalls Is Nothing Then ApprovedToolCalls = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                ApprovedToolCalls.Add(BuildApprovalKey(toolId, params))
                Return
            End If
            If ApprovedTools Is Nothing Then ApprovedTools = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            ApprovedTools.Add(toolId.Trim())
        End Sub

        ''' <summary>允许本会话内该工具的所有调用；不做参数匹配，也不在放行后移除。</summary>
        Public Sub ApproveToolForSession(toolId As String)
            If String.IsNullOrWhiteSpace(toolId) Then Return
            If SessionApprovedTools Is Nothing Then
                SessionApprovedTools = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            End If
            SessionApprovedTools.Add(toolId.Trim())
        End Sub

        Public Function IsToolApprovedForSession(toolId As String) As Boolean
            Return SessionApprovedTools IsNot Nothing AndAlso
                   Not String.IsNullOrWhiteSpace(toolId) AndAlso
                   SessionApprovedTools.Contains(toolId.Trim())
        End Function

        Public Function IsToolApproved(toolId As String, Optional params As JObject = Nothing) As Boolean
            If params IsNot Nothing AndAlso ApprovedToolCalls IsNot Nothing Then
                Return ApprovedToolCalls.Contains(BuildApprovalKey(toolId, params))
            End If
            Return ApprovedTools IsNot Nothing AndAlso
                   Not String.IsNullOrWhiteSpace(toolId) AndAlso
                   ApprovedTools.Contains(toolId.Trim())
        End Function

        Public Function ConsumeToolApproval(toolId As String, Optional params As JObject = Nothing) As Boolean
            If params IsNot Nothing AndAlso ApprovedToolCalls IsNot Nothing Then
                Dim approvalKey = BuildApprovalKey(toolId, params)
                Return ApprovedToolCalls.Remove(approvalKey)
            End If
            If Not IsToolApproved(toolId, Nothing) Then Return False
            Return ApprovedTools.Remove(toolId.Trim())
        End Function

        Public Shared Function BuildApprovalKey(toolId As String, params As JObject) As String
            Dim normalizedToolId = If(toolId, "").Trim().ToLowerInvariant()
            Dim canonicalParams = CanonicalizeToken(If(params, New JObject())).ToString(Formatting.None)
            Dim bytes = Encoding.UTF8.GetBytes(canonicalParams)
            Using sha = SHA256.Create()
                Dim hash = BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant()
                Return normalizedToolId & ":" & hash
            End Using
        End Function

        Private Shared Function CanonicalizeToken(token As JToken) As JToken
            If token Is Nothing Then Return JValue.CreateNull()

            If token.Type = JTokenType.Object Then
                Dim source = DirectCast(token, JObject)
                Dim result As New JObject()
                For Each prop In source.Properties().OrderBy(Function(item) item.Name, StringComparer.Ordinal)
                    result.Add(prop.Name, CanonicalizeToken(prop.Value))
                Next
                Return result
            End If

            If token.Type = JTokenType.Array Then
                Dim result As New JArray()
                For Each item In DirectCast(token, JArray)
                    result.Add(CanonicalizeToken(item))
                Next
                Return result
            End If

            Return token.DeepClone()
        End Function
    End Class

End Namespace
