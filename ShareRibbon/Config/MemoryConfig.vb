' ShareRibbon\Config\MemoryConfig.vb
' 记忆相关配置项

Imports System.IO
Imports Newtonsoft.Json

''' <summary>
''' 记忆配置：memory.* 配置项读写
''' </summary>
Public Class MemoryConfig

    Private Shared _configPath As String
    Private Shared _config As MemoryConfigData

    Private Shared ReadOnly _lockObj As New Object()

    Private Shared ReadOnly Property ConfigPath As String
        Get
            If String.IsNullOrEmpty(_configPath) Then
                _configPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    ConfigSettings.OfficeAiAppDataFolder,
                    "memory_config.json")
            End If
            Return _configPath
        End Get
    End Property

    Private Shared Function LoadConfig() As MemoryConfigData
        SyncLock _lockObj
            If _config IsNot Nothing Then Return _config

            _config = New MemoryConfigData()
            Try
                Dim dir = Path.GetDirectoryName(ConfigPath)
                If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If
                If File.Exists(ConfigPath) Then
                    Dim json = File.ReadAllText(ConfigPath)
                    Dim loaded = JsonConvert.DeserializeObject(Of MemoryConfigData)(json)
                    If loaded IsNot Nothing Then _config = loaded
                End If
            Catch ex As Exception
                Debug.WriteLine("MemoryConfig load failed: " & ex.Message)
            End Try
            Return _config
        End SyncLock
    End Function

    Private Shared Sub SaveConfig()
        SyncLock _lockObj
            If _config Is Nothing Then Return
            Try
                Dim json = JsonConvert.SerializeObject(_config, Formatting.Indented)
                File.WriteAllText(ConfigPath, json)
            Catch ex As Exception
                Debug.WriteLine("MemoryConfig save failed: " & ex.Message)
            End Try
        End SyncLock
    End Sub

    Public Shared Property RagTopN As Integer
        Get
            Return LoadConfig().RagTopN
        End Get
        Set(value As Integer)
            LoadConfig()
            _config.RagTopN = Math.Max(1, Math.Min(20, value))
            SaveConfig()
        End Set
    End Property

    Public Shared Property EnableAgenticSearch As Boolean
        Get
            Return LoadConfig().EnableAgenticSearch
        End Get
        Set(value As Boolean)
            LoadConfig()
            _config.EnableAgenticSearch = value
            SaveConfig()
        End Set
    End Property

    Public Shared Property EnableUserProfile As Boolean
        Get
            Return LoadConfig().EnableUserProfile
        End Get
        Set(value As Boolean)
            LoadConfig()
            _config.EnableUserProfile = value
            SaveConfig()
        End Set
    End Property

    Public Shared Property AtomicContentMaxLength As Integer
        Get
            Return LoadConfig().AtomicContentMaxLength
        End Get
        Set(value As Integer)
            LoadConfig()
            _config.AtomicContentMaxLength = Math.Max(50, Math.Min(500, value))
            SaveConfig()
        End Set
    End Property

    Public Shared Property SessionSummaryLimit As Integer
        Get
            Return LoadConfig().SessionSummaryLimit
        End Get
        Set(value As Integer)
            LoadConfig()
            _config.SessionSummaryLimit = Math.Max(1, Math.Min(15, value))
            SaveConfig()
        End Set
    End Property

    ''' <summary>
    ''' 是否使用 ContextBuilder 分层组装上下文（含 Memory/Skills）
    ''' </summary>
    Public Shared Property UseContextBuilder As Boolean
        Get
            Return LoadConfig().UseContextBuilder
        End Get
        Set(value As Boolean)
            LoadConfig()
            _config.UseContextBuilder = value
            SaveConfig()
        End Set
    End Property

    Private Class MemoryConfigData
        Public Property RagTopN As Integer = 5
        Public Property EnableAgenticSearch As Boolean = False
        Public Property EnableUserProfile As Boolean = True
        Public Property AtomicContentMaxLength As Integer = 200
        Public Property SessionSummaryLimit As Integer = 5
        Public Property UseContextBuilder As Boolean = True
    End Class
End Class
