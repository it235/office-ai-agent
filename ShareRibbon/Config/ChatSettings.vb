﻿Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ChatSettings
    Private ReadOnly _applicationInfo As ApplicationInfo
    Public Property EnabledMcpList As List(Of String)

    Public Sub New(applicationInfo As ApplicationInfo)
        _applicationInfo = applicationInfo

        ' 初始化MCP列表
        EnabledMcpList = New List(Of String)()

        LoadSettings()
    End Sub

    Public Shared Property topicRandomness As Double = 0.8  ' 默认值改为 Double
    Public Shared Property contextLimit As Integer = 5     ' 默认值改为 Integer
    Public Shared Property selectedCellChecked As Boolean = False
    Public Shared Property executecodePreviewChecked As Boolean = True ' 执行代码前预览的默认选项
    Public Shared Property settingsScrollChecked As Boolean = True
    Public Shared Property chatMode As String = "chat"

    ' 修改方法签名，参数类型改为 Double 和 Integer
    Public Sub SaveSettings(topicRandomness As Double, contextLimit As Integer,
                          selectedCell As Boolean, settingsScroll As Boolean, executecodePreview As Boolean, chatMode As String)
        Try
            ' 创建设置对象
            Dim settings As New Dictionary(Of String, Object) From {
                {"topicRandomness", topicRandomness},
                {"contextLimit", contextLimit},
                {"selectedCellChecked", selectedCell},
                {"settingsScrollChecked", settingsScroll},
                {"executecodePreviewChecked", executecodePreview},
                {"chatMode", chatMode}
            }

            ' 将设置保存到JSON文件
            Dim settingsPath = _applicationInfo.GetChatSettingsFilePath()

            ' 确保目录存在
            Directory.CreateDirectory(Path.GetDirectoryName(settingsPath))

            ' 将设置序列化为JSON并保存
            File.WriteAllText(settingsPath, JsonConvert.SerializeObject(settings, Formatting.Indented))

            ' 更新静态属性
            ChatSettings.topicRandomness = topicRandomness
            ChatSettings.contextLimit = contextLimit
            ChatSettings.selectedCellChecked = selectedCell
            ChatSettings.settingsScrollChecked = settingsScroll
            ChatSettings.executecodePreviewChecked = executecodePreview
            ChatSettings.chatMode = chatMode

        Catch ex As Exception
            Debug.WriteLine($"保存设置失败: {ex.Message}")
        End Try
    End Sub

    ' 加载设置时进行类型转换
    Public Sub LoadSettings()
        Try
            Dim settingsPath = _applicationInfo.GetChatSettingsFilePath()

            If File.Exists(settingsPath) Then
                ' 读取JSON文件
                Dim json = File.ReadAllText(settingsPath)
                Dim settings = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(json)

                ' 更新静态属性，添加类型转换
                If settings.ContainsKey("topicRandomness") Then
                    topicRandomness = Convert.ToDouble(settings("topicRandomness"))
                End If
                If settings.ContainsKey("contextLimit") Then
                    contextLimit = Convert.ToInt32(settings("contextLimit"))
                End If
                If settings.ContainsKey("selectedCellChecked") Then
                    selectedCellChecked = CBool(settings("selectedCellChecked"))
                End If
                If settings.ContainsKey("settingsScrollChecked") Then
                    settingsScrollChecked = CBool(settings("settingsScrollChecked"))
                End If
                If settings.ContainsKey("executecodePreviewChecked") Then
                    executecodePreviewChecked = CBool(settings("executecodePreviewChecked"))
                End If
                If settings.ContainsKey("chatMode") Then
                    chatMode = Convert.ToString(settings("chatMode"))
                End If
                ' 加载MCP列表

                ' 加载MCP列表 - 修复了键不存在的问题
                If settings.ContainsKey("enableMcpList") AndAlso settings("enableMcpList") IsNot Nothing Then
                    Try
                        ' 处理不同可能的类型
                        If TypeOf settings("enableMcpList") Is JArray Then
                            EnabledMcpList = CType(settings("enableMcpList"), JArray).ToObject(Of List(Of String))()
                        ElseIf TypeOf settings("enableMcpList") Is String Then
                            EnabledMcpList = JsonConvert.DeserializeObject(Of List(Of String))(settings("enableMcpList").ToString())
                        Else
                            ' 尝试通过转换字符串再反序列化
                            EnabledMcpList = JsonConvert.DeserializeObject(Of List(Of String))(JsonConvert.SerializeObject(settings("enableMcpList")))
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"解析enableMcpList失败: {ex.Message}")
                        EnabledMcpList = New List(Of String)() ' 确保始终有一个有效的列表
                    End Try
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"加载ChatSettings失败: {ex.Message}")
        End Try
    End Sub

    Public Sub SaveEnabledMcpList(enabledList As List(Of String))
        EnabledMcpList = enabledList

        Dim settingsPath = _applicationInfo.GetChatSettingsFilePath()
        ' 读取现有设置
        Dim settingsObj As JObject
        If File.Exists(settingsPath) Then
            Dim jsonContent = File.ReadAllText(settingsPath)
            settingsObj = JObject.Parse(jsonContent)
        Else
            settingsObj = New JObject()
        End If

        ' 更新MCP列表
        settingsObj("enableMcpList") = JArray.FromObject(enabledList)

        ' 保存设置
        File.WriteAllText(settingsPath, settingsObj.ToString(Formatting.Indented))
    End Sub
End Class