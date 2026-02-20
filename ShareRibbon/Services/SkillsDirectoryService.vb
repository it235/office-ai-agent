' ShareRibbon\Services\SkillsDirectoryService.vb
' Skills目录管理服务：从文件系统读取Claude规范的Skills

Imports System.IO
Imports System.Collections.Generic
Imports System.Linq
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Skills文件定义（Claude规范格式，支持Front Matter）
''' </summary>
Public Class SkillFileDefinition
    ' 必需字段
    Public Property Name As String
    Public Property Description As String
    Public Property Content As String

    ' 可选字段（Front Matter元数据）
    Public Property License As String
    Public Property Compatibility As String
    Public Property AllowedTools As List(Of String)
    Public Property Metadata As Dictionary(Of String, Object)
    Public Property ArgumentHint As String
    Public Property DisableModelInvocation As Boolean = False
    Public Property UserInvocable As Boolean = True
    Public Property Model As String
    Public Property Context As String
    Public Property Agent As String
    Public Property Hooks As Dictionary(Of String, String)

    ' 扩展字段
    Public Property FilePath As String

    Public ReadOnly Property AllowedToolsText As String
        Get
            If AllowedTools Is Nothing OrElse AllowedTools.Count = 0 Then Return ""
            Return String.Join(", ", AllowedTools)
        End Get
    End Property

    Public ReadOnly Property Author As String
        Get
            If Metadata IsNot Nothing AndAlso Metadata.ContainsKey("author") Then
                Return Metadata("author")?.ToString()
            End If
            Return ""
        End Get
    End Property

    Public ReadOnly Property Version As String
        Get
            If Metadata IsNot Nothing AndAlso Metadata.ContainsKey("version") Then
                Return Metadata("version")?.ToString()
            End If
            Return "1.0"
        End Get
    End Property
End Class

''' <summary>
''' Skills目录服务
''' 管理Skills文件的读取、解析和缓存
''' </summary>
Public Class SkillsDirectoryService
    Private Shared _skillsDirectory As String = ""
    Private Shared _cachedSkills As New List(Of SkillFileDefinition)()
    Private Shared _lastRefreshTime As DateTime = DateTime.MinValue
    Private Shared ReadOnly _cacheDuration As TimeSpan = TimeSpan.FromMinutes(5)

    ''' <summary>
    ''' 获取Skills目录路径
    ''' </summary>
    Public Shared Function GetSkillsDirectory() As String
        If Not String.IsNullOrEmpty(_skillsDirectory) Then
            Return _skillsDirectory
        End If

        ' 默认路径：Documents\OfficeAiAppData\Skills
        Dim appDataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder,
            "Skills")
        _skillsDirectory = appDataPath
        Return _skillsDirectory
    End Function

    ''' <summary>
    ''' 设置Skills目录路径
    ''' </summary>
    Public Shared Sub SetSkillsDirectory(path As String)
        _skillsDirectory = path
        _cachedSkills.Clear()
        _lastRefreshTime = DateTime.MinValue
    End Sub

    ''' <summary>
    ''' 确保Skills目录存在
    ''' </summary>
    Public Shared Sub EnsureDirectoryExists()
        Dim dir = GetSkillsDirectory()
        If Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If
    End Sub

    ''' <summary>
    ''' 获取所有Skills（带缓存）
    ''' </summary>
    Public Shared Function GetAllSkills(Optional forceRefresh As Boolean = False) As List(Of SkillFileDefinition)
        If Not forceRefresh AndAlso _cachedSkills.Count > 0 AndAlso (DateTime.Now - _lastRefreshTime) < _cacheDuration Then
            Return _cachedSkills.ToList()
        End If

        RefreshSkills()
        Return _cachedSkills.ToList()
    End Function

    ''' <summary>
    ''' 刷新Skills缓存
    ''' </summary>
    Public Shared Sub RefreshSkills()
        _cachedSkills.Clear()

        Dim dir = GetSkillsDirectory()
        If Not Directory.Exists(dir) Then
            _lastRefreshTime = DateTime.Now
            Return
        End If

        ' 读取所有目录（每个目录是一个Skill）
        Dim skillDirs = Directory.GetDirectories(dir)
        For Each skillDir In skillDirs
            Try
                Dim skill = ParseSkillDirectory(skillDir)
                If skill IsNot Nothing Then
                    _cachedSkills.Add(skill)
                End If
            Catch ex As Exception
                Debug.WriteLine($"[SkillsDirectoryService] 解析目录失败: {skillDir}, 错误: {ex.Message}")
            End Try
        Next

        ' 也读取根目录下的JSON文件（兼容旧格式）
        Dim jsonFiles = Directory.GetFiles(dir, "*.json", SearchOption.TopDirectoryOnly)
        For Each file In jsonFiles
            Try
                Dim skill = ParseSkillJsonFile(file)
                If skill IsNot Nothing Then
                    _cachedSkills.Add(skill)
                End If
            Catch ex As Exception
                Debug.WriteLine($"[SkillsDirectoryService] 解析JSON文件失败: {file}, 错误: {ex.Message}")
            End Try
        Next

        _lastRefreshTime = DateTime.Now
    End Sub

    ''' <summary>
    ''' 解析Skill目录（Claude规范）
    ''' 目录结构：
    ''' my-skill/
    '''   ├── SKILL.md (required)
    '''   ├── reference.md (optional)
    '''   ├── examples.md (optional)
    '''   ├── scripts/ (optional)
    '''   └── templates/ (optional)
    ''' </summary>
    Private Shared Function ParseSkillDirectory(dirPath As String) As SkillFileDefinition
        Dim skillMdPath = Path.Combine(dirPath, "SKILL.md")
        If Not File.Exists(skillMdPath) Then
            Return Nothing
        End If

        Dim skillName = Path.GetFileName(dirPath)
        Dim fileContent = File.ReadAllText(skillMdPath)

        Dim skill As New SkillFileDefinition()
        skill.Name = skillName
        skill.FilePath = dirPath

        ' 解析Front Matter和内容
        ParseFrontMatterAndContent(fileContent, skill)

        ' 尝试读取reference.md
        Dim refPath = Path.Combine(dirPath, "reference.md")
        If File.Exists(refPath) Then
            skill.Content &= vbCrLf & vbCrLf & "---" & vbCrLf & vbCrLf & File.ReadAllText(refPath)
        End If

        ' 尝试读取examples.md
        Dim examplesPath = Path.Combine(dirPath, "examples.md")
        If File.Exists(examplesPath) Then
            skill.Content &= vbCrLf & vbCrLf & "---" & vbCrLf & vbCrLf & File.ReadAllText(examplesPath)
        End If

        Return skill
    End Function

    ''' <summary>
    ''' 解析Front Matter和内容
    ''' </summary>
    Private Shared Sub ParseFrontMatterAndContent(fileContent As String, skill As SkillFileDefinition)
        ' 查找Front Matter分隔符
        Dim lines = fileContent.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
        Dim frontMatterLines As New List(Of String)()
        Dim contentLines As New List(Of String)()
        Dim inFrontMatter = False
        Dim frontMatterEnded = False
        Dim frontMatterStartIndex = -1
        Dim frontMatterEndIndex = -1

        For i = 0 To lines.Length - 1
            Dim line = lines(i).Trim()

            If line = "---" Then
                If Not frontMatterEnded Then
                    If frontMatterStartIndex = -1 Then
                        frontMatterStartIndex = i
                        inFrontMatter = True
                    Else
                        frontMatterEndIndex = i
                        inFrontMatter = False
                        frontMatterEnded = True
                    End If
                End If
            ElseIf inFrontMatter Then
                frontMatterLines.Add(lines(i))
            ElseIf frontMatterEnded OrElse frontMatterStartIndex = -1 Then
                contentLines.Add(lines(i))
            End If
        Next

        ' 解析Front Matter
        If frontMatterLines.Count > 0 Then
            ParseFrontMatterLines(frontMatterLines, skill)
        End If

        ' 剩余部分作为内容
        skill.Content = String.Join(vbCrLf, contentLines).Trim()

        ' 如果没有从Front Matter获取到name，从第一个标题获取
        If String.IsNullOrWhiteSpace(skill.Name) OrElse skill.Name = Path.GetFileName(skill.FilePath) Then
            For Each line In contentLines
                If line.StartsWith("# ") Then
                    skill.Name = line.Substring(2).Trim()
                    Exit For
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 解析Front Matter行
    ''' </summary>
    Private Shared Sub ParseFrontMatterLines(lines As List(Of String), skill As SkillFileDefinition)
        skill.Metadata = New Dictionary(Of String, Object)()
        Dim inMetadata = False

        For Each line In lines
            Dim trimmedLine = line.Trim()

            If String.IsNullOrWhiteSpace(trimmedLine) Then
                Continue For
            End If

            If trimmedLine.StartsWith("metadata:") Then
                inMetadata = True
                Continue For
            End If

            If inMetadata AndAlso trimmedLine.StartsWith("  ") Then
                ' metadata下的子项
                Dim colonIndex = trimmedLine.IndexOf(":")
                If colonIndex > 0 Then
                    Dim key = trimmedLine.Substring(0, colonIndex).Trim()
                    Dim value = trimmedLine.Substring(colonIndex + 1).Trim()
                    ' 移除引号
                    If value.StartsWith("""") AndAlso value.EndsWith("""") Then
                        value = value.Substring(1, value.Length - 2)
                    ElseIf value.StartsWith("'") AndAlso value.EndsWith("'") Then
                        value = value.Substring(1, value.Length - 2)
                    End If
                    skill.Metadata(key) = value
                End If
                Continue For
            ElseIf inMetadata Then
                inMetadata = False
            End If

            ' 标准字段
            Dim colonIndex2 = trimmedLine.IndexOf(":")
            If colonIndex2 > 0 Then
                Dim key = trimmedLine.Substring(0, colonIndex2).Trim()
                Dim value = trimmedLine.Substring(colonIndex2 + 1).Trim()
                ' 移除引号
                If value.StartsWith("""") AndAlso value.EndsWith("""") Then
                    value = value.Substring(1, value.Length - 2)
                ElseIf value.StartsWith("'") AndAlso value.EndsWith("'") Then
                    value = value.Substring(1, value.Length - 2)
                End If
                ' 移除#注释
                Dim hashIndex = value.IndexOf("#")
                If hashIndex > 0 Then
                    value = value.Substring(0, hashIndex).Trim()
                End If

                Select Case key.ToLowerInvariant()
                    Case "name"
                        skill.Name = value
                    Case "description"
                        skill.Description = value
                    Case "license"
                        skill.License = value
                    Case "compatibility"
                        skill.Compatibility = value
                    Case "allowed-tools"
                        If Not String.IsNullOrWhiteSpace(value) Then
                            skill.AllowedTools = value.Split({","c}, StringSplitOptions.RemoveEmptyEntries).Select(Function(s) s.Trim()).ToList()
                        End If
                    Case "argument-hint"
                        skill.ArgumentHint = value
                    Case "disable-model-invocation"
                        Dim boolVal As Boolean
                        If Boolean.TryParse(value, boolVal) Then
                            skill.DisableModelInvocation = boolVal
                        End If
                    Case "user-invocable"
                        Dim boolVal As Boolean
                        If Boolean.TryParse(value, boolVal) Then
                            skill.UserInvocable = boolVal
                        End If
                    Case "model"
                        skill.Model = value
                    Case "context"
                        skill.Context = value
                    Case "agent"
                        skill.Agent = value
                End Select
            End If
        Next
    End Sub

    ''' <summary>
    ''' 解析单个Skill JSON文件（兼容旧格式）
    ''' </summary>
    Private Shared Function ParseSkillJsonFile(filePath As String) As SkillFileDefinition
        Dim content = File.ReadAllText(filePath)
        Dim jo = JObject.Parse(content)

        Dim skill As New SkillFileDefinition()
        skill.FilePath = filePath

        ' 读取基本信息
        skill.Name = If(jo("name")?.ToString(), Path.GetFileNameWithoutExtension(filePath))
        skill.Description = If(jo("description")?.ToString(), "")

        ' 读取keywords兼容
        Dim keywordsToken = jo("keywords")
        If keywordsToken IsNot Nothing AndAlso TypeOf keywordsToken Is JArray Then
            ' 不做任何处理，保持兼容性
        ElseIf keywordsToken IsNot Nothing AndAlso TypeOf keywordsToken Is JValue Then
            ' 不做任何处理，保持兼容性
        End If

        ' 读取Skill内容
        Dim contentToken = jo("content")
        If contentToken IsNot Nothing Then
            skill.Content = contentToken.ToString()
        Else
            Dim promptToken = jo("prompt")
            If promptToken IsNot Nothing Then
                skill.Content = promptToken.ToString()
            Else
                Dim promptTemplateToken = jo("promptTemplate")
                If promptTemplateToken IsNot Nothing Then
                    skill.Content = promptTemplateToken.ToString()
                End If
            End If
        End If

        Return skill
    End Function

    ''' <summary>
    ''' 打开Skills目录
    ''' </summary>
    Public Shared Sub OpenSkillsDirectory()
        EnsureDirectoryExists()
        Dim dir = GetSkillsDirectory()
        Process.Start("explorer.exe", dir)
    End Sub

    ''' <summary>
    ''' 打开指定Skill的目录
    ''' </summary>
    Public Shared Sub OpenSkillDirectory(skill As SkillFileDefinition)
        If skill Is Nothing OrElse String.IsNullOrWhiteSpace(skill.FilePath) Then
            OpenSkillsDirectory()
            Return
        End If

        If Directory.Exists(skill.FilePath) Then
            Process.Start("explorer.exe", skill.FilePath)
        ElseIf File.Exists(skill.FilePath) Then
            Process.Start("explorer.exe", Path.GetDirectoryName(skill.FilePath))
        Else
            OpenSkillsDirectory()
        End If
    End Sub

End Class
