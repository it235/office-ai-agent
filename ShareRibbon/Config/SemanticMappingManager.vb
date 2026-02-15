' ShareRibbon\Config\SemanticMappingManager.vb
' 语义映射管理器（单例模式）

Imports System.IO
Imports Newtonsoft.Json

''' <summary>
''' 语义映射管理器（单例模式）
''' 管理 SemanticStyleMapping 的持久化、CRUD操作
''' </summary>
Public Class SemanticMappingManager
    Private Shared _instance As SemanticMappingManager
    Private _mappings As List(Of SemanticStyleMapping)
    Private ReadOnly _configPath As String

    ''' <summary>获取单例实例</summary>
    Public Shared ReadOnly Property Instance As SemanticMappingManager
        Get
            If _instance Is Nothing Then
                _instance = New SemanticMappingManager()
            End If
            Return _instance
        End Get
    End Property

    ''' <summary>获取所有映射</summary>
    Public ReadOnly Property Mappings As List(Of SemanticStyleMapping)
        Get
            Return _mappings
        End Get
    End Property

    Private Sub New()
        _configPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder,
            "semantic_mappings.json")
        LoadMappings()
    End Sub

    ''' <summary>加载映射配置</summary>
    Private Sub LoadMappings()
        _mappings = New List(Of SemanticStyleMapping)()

        If File.Exists(_configPath) Then
            Try
                Dim json = File.ReadAllText(_configPath, Text.Encoding.UTF8)
                Dim loaded = JsonConvert.DeserializeObject(Of List(Of SemanticStyleMapping))(json)
                If loaded IsNot Nothing Then
                    _mappings = loaded
                End If
            Catch ex As Exception
                Debug.WriteLine($"加载语义映射配置失败: {ex.Message}")
            End Try
        End If
    End Sub

    ''' <summary>保存映射配置</summary>
    Public Sub SaveMappings()
        Try
            Dim dir = Path.GetDirectoryName(_configPath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Dim json = JsonConvert.SerializeObject(_mappings, Formatting.Indented)
            File.WriteAllText(_configPath, json, Text.Encoding.UTF8)
        Catch ex As Exception
            Debug.WriteLine($"保存语义映射配置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>添加映射</summary>
    Public Sub AddMapping(mapping As SemanticStyleMapping)
        If mapping Is Nothing Then Return
        _mappings.Add(mapping)
        SaveMappings()
    End Sub

    ''' <summary>根据ID获取映射</summary>
    Public Function GetMappingById(id As String) As SemanticStyleMapping
        Return _mappings.FirstOrDefault(Function(m) m.Id = id)
    End Function

    ''' <summary>根据来源ID获取映射（查找已转换缓存）</summary>
    Public Function GetMappingBySourceId(sourceId As String) As SemanticStyleMapping
        Return _mappings.FirstOrDefault(Function(m) m.SourceId = sourceId)
    End Function

    ''' <summary>更新映射</summary>
    Public Sub UpdateMapping(mapping As SemanticStyleMapping)
        If mapping Is Nothing Then Return
        Dim index = _mappings.FindIndex(Function(m) m.Id = mapping.Id)
        If index >= 0 Then
            mapping.LastModified = DateTime.Now
            _mappings(index) = mapping
            SaveMappings()
        End If
    End Sub

    ''' <summary>删除映射</summary>
    Public Sub DeleteMapping(id As String)
        _mappings.RemoveAll(Function(m) m.Id = id)
        SaveMappings()
    End Sub

    ''' <summary>刷新（重新从磁盘加载）</summary>
    Public Sub Refresh()
        LoadMappings()
    End Sub
End Class
