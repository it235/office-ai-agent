' ShareRibbon\Storage\ConversationRepository.vb
' 会话消息表 CRUD

Imports System.Data.SQLite

''' <summary>
''' 单条会话消息 DTO（用于加载历史会话）
''' </summary>
Public Class ConversationMessageDto
    Public Property Role As String
    Public Property Content As String
    Public Property CreateTime As String
End Class

''' <summary>
''' 会话消息 CRUD
''' </summary>
Public Class ConversationRepository

    ''' <summary>
    ''' 插入一条会话消息
    ''' </summary>
    Public Shared Sub InsertMessage(sessionId As String, role As String, content As String, Optional isCollected As Boolean = False)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "INSERT INTO conversation (session_id, role, content, is_collected) VALUES (@sid, @role, @content, @collected)", conn)
                cmd.Parameters.AddWithValue("@sid", sessionId)
                cmd.Parameters.AddWithValue("@role", role)
                cmd.Parameters.AddWithValue("@content", If(content, ""))
                cmd.Parameters.AddWithValue("@collected", If(isCollected, 1, 0))
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 更新消息收藏状态
    ''' </summary>
    Public Shared Sub SetCollected(conversationId As Long, isCollected As Boolean)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("UPDATE conversation SET is_collected=@c WHERE id=@id", conn)
                cmd.Parameters.AddWithValue("@c", If(isCollected, 1, 0))
                cmd.Parameters.AddWithValue("@id", conversationId)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 按 responseUuid 更新收藏（需通过 session_id + 最新 assistant 消息定位，简化实现：按 session 最后一条 assistant 更新）
    ''' 若调用方有 conversation_id 可直接用 SetCollected
    ''' </summary>
    Public Shared Sub SetLastAssistantCollected(sessionId As String, isCollected As Boolean)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "UPDATE conversation SET is_collected=@c WHERE id=(SELECT id FROM conversation WHERE session_id=@sid AND role='assistant' ORDER BY create_time DESC LIMIT 1)", conn)
                cmd.Parameters.AddWithValue("@c", If(isCollected, 1, 0))
                cmd.Parameters.AddWithValue("@sid", sessionId)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 按会话 ID 获取该会话下所有消息（按 create_time 升序），用于加载历史会话到界面
    ''' </summary>
    Public Shared Function GetMessagesBySession(sessionId As String) As List(Of ConversationMessageDto)
        OfficeAiDatabase.EnsureInitialized()
        Dim list As New List(Of ConversationMessageDto)()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT role, content, create_time FROM conversation WHERE session_id=@sid ORDER BY create_time ASC", conn)
                cmd.Parameters.AddWithValue("@sid", sessionId)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        list.Add(New ConversationMessageDto With {
                            .Role = rdr.GetString(0),
                            .Content = If(rdr.IsDBNull(1), "", rdr.GetString(1)),
                            .CreateTime = If(rdr.IsDBNull(2), "", rdr.GetString(2))
                        })
                    End While
                End Using
            End Using
        End Using
        Return list
    End Function
End Class
