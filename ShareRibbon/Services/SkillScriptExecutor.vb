' ShareRibbon\Services\SkillScriptExecutor.vb
' Skill脚本执行器：执行Python/PowerShell/Shell/Batch脚本

Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks

''' <summary>
''' 脚本执行结果
''' </summary>
Public Class ScriptExecutionResult
    ''' <summary>是否成功</summary>
    Public Property Success As Boolean

    ''' <summary>标准输出</summary>
    Public Property StdOut As String

    ''' <summary>标准错误</summary>
    Public Property StdErr As String

    ''' <summary>退出代码</summary>
    Public Property ExitCode As Integer

    ''' <summary>执行耗时（毫秒）</summary>
    Public Property ElapsedMs As Long

    ''' <summary>错误消息</summary>
    Public Property ErrorMessage As String

    ''' <summary>执行的脚本路径</summary>
    Public Property ScriptPath As String
End Class

''' <summary>
''' Skill脚本执行器
''' 支持 Python、PowerShell、Shell、Bat/CMD 脚本执行
''' </summary>
Public Class SkillScriptExecutor

    ''' <summary>
    ''' 执行 Skill 脚本
    ''' </summary>
    ''' <param name="script">脚本定义</param>
    ''' <param name="args">脚本参数（将作为环境变量传递）</param>
    ''' <param name="skillDir">Skill根目录（用于设置工作目录）</param>
    ''' <returns>执行结果</returns>
    Public Shared Async Function ExecuteScriptAsync(
        script As SkillScript,
        Optional args As Dictionary(Of String, String) = Nothing,
        Optional skillDir As String = "") As Task(Of ScriptExecutionResult)

        Dim result As New ScriptExecutionResult()
        result.ScriptPath = script.FilePath
        Dim sw = Stopwatch.StartNew()

        Try
            ' 确定工作目录
            Dim workingDir As String
            If Not String.IsNullOrEmpty(script.WorkingDirectory) Then
                workingDir = If(Path.IsPathRooted(script.WorkingDirectory),
                    script.WorkingDirectory,
                    Path.Combine(skillDir, script.WorkingDirectory))
            ElseIf Not String.IsNullOrEmpty(skillDir) Then
                workingDir = skillDir
            Else
                workingDir = Path.GetDirectoryName(script.FilePath)
            End If

            ' 构建环境变量
            Dim envVars As New Dictionary(Of String, String)
            For Each dict In Environment.GetEnvironmentVariables()
                envVars(dict.Key) = dict.Value
            Next
            If args IsNot Nothing Then
                For Each kvp In args
                    envVars($"SKILL_ARG_{kvp.Key}") = kvp.Value
                Next
            End If

            ' 根据脚本类型执行
            Select Case script.ScriptType.ToLower()
                Case "python"
                    Dim pyResult = Await ExecutePythonAsync(script.FilePath, workingDir, args, envVars)
                    result.Success = pyResult.Success
                    result.StdOut = pyResult.StdOut
                    result.StdErr = pyResult.StdErr
                    result.ExitCode = pyResult.ExitCode
                    result.ErrorMessage = pyResult.ErrorMessage

                Case "powershell"
                    Dim psResult = Await ExecutePowerShellAsync(script.FilePath, workingDir, args, envVars)
                    result.Success = psResult.Success
                    result.StdOut = psResult.StdOut
                    result.StdErr = psResult.StdErr
                    result.ExitCode = psResult.ExitCode
                    result.ErrorMessage = psResult.ErrorMessage

                Case "shell"
                    Dim shResult = Await ExecuteShellAsync(script.FilePath, workingDir, args, envVars)
                    result.Success = shResult.Success
                    result.StdOut = shResult.StdOut
                    result.StdErr = shResult.StdErr
                    result.ExitCode = shResult.ExitCode
                    result.ErrorMessage = shResult.ErrorMessage

                Case "batch"
                    Dim batResult = Await ExecuteBatchAsync(script.FilePath, workingDir, args, envVars)
                    result.Success = batResult.Success
                    result.StdOut = batResult.StdOut
                    result.StdErr = batResult.StdErr
                    result.ExitCode = batResult.ExitCode
                    result.ErrorMessage = batResult.ErrorMessage

                Case Else
                    result.Success = False
                    result.ErrorMessage = $"不支持的脚本类型: {script.ScriptType}"
            End Select

        Catch ex As Exception
            result.Success = False
            result.ErrorMessage = $"脚本执行异常: {ex.Message}"
            result.StdErr = ex.ToString()
        End Try

        sw.Stop()
        result.ElapsedMs = sw.ElapsedMilliseconds
        Return result
    End Function

    ''' <summary>
    ''' 执行 Python 脚本
    ''' </summary>
    Private Shared Async Function ExecutePythonAsync(
        scriptPath As String,
        workingDir As String,
        args As Dictionary(Of String, String),
        envVars As Dictionary(Of String, String)) As Task(Of ScriptExecutionResult)

        Dim result As New ScriptExecutionResult()

        ' 查找 python.exe
        Dim pythonPath = FindPython()
        If String.IsNullOrEmpty(pythonPath) Then
            result.Success = False
            result.ErrorMessage = "未找到 Python 解释器，请确保已安装 Python 并添加到 PATH"
            Return result
        End If

        ' 构建命令行参数
        Dim arguments = $" ""{scriptPath}"""
        If args IsNot Nothing Then
            For Each kvp In args
                arguments &= $" --{kvp.Key}=""{kvp.Value}"""
            Next
        End If

        Using proc = New Process()
            proc.StartInfo.FileName = pythonPath
            proc.StartInfo.Arguments = arguments
            proc.StartInfo.WorkingDirectory = workingDir
            proc.StartInfo.UseShellExecute = False
            proc.StartInfo.RedirectStandardOutput = True
            proc.StartInfo.RedirectStandardError = True
            proc.StartInfo.CreateNoWindow = True
            proc.StartInfo.StandardOutputEncoding = Encoding.UTF8
            proc.StartInfo.StandardErrorEncoding = Encoding.UTF8

            ' 设置环境变量
            For Each kvp In envVars
                proc.StartInfo.EnvironmentVariables(kvp.Key) = kvp.Value
            Next

            AddHandler proc.OutputDataReceived, Sub(s, e)
                If e.Data IsNot Nothing Then result.StdOut &= e.Data & Environment.NewLine
            End Sub
            AddHandler proc.ErrorDataReceived, Sub(s, e)
                If e.Data IsNot Nothing Then result.StdErr &= e.Data & Environment.NewLine
            End Sub

            Try
                proc.Start()
                proc.BeginOutputReadLine()
                proc.BeginErrorReadLine()
                Await Task.Run(Sub() proc.WaitForExit())
                result.ExitCode = proc.ExitCode
                result.Success = proc.ExitCode = 0
            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = $"启动 Python 失败: {ex.Message}"
            End Try
        End Using

        Return result
    End Function

    ''' <summary>
    ''' 执行 PowerShell 脚本
    ''' </summary>
    Private Shared Async Function ExecutePowerShellAsync(
        scriptPath As String,
        workingDir As String,
        args As Dictionary(Of String, String),
        envVars As Dictionary(Of String, String)) As Task(Of ScriptExecutionResult)

        Dim result As New ScriptExecutionResult()

        ' 查找 powershell.exe
        Dim psPath = FindPowerShell()
        If String.IsNullOrEmpty(psPath) Then
            result.Success = False
            result.ErrorMessage = "未找到 PowerShell 解释器"
            Return result
        End If

        ' 构建命令行参数
        Dim argStr = $"-ExecutionPolicy Bypass -NoProfile -File ""{scriptPath}"""
        If args IsNot Nothing Then
            For Each kvp In args
                argStr &= $" -{kvp.Key} ""{kvp.Value}"""
            Next
        End If

        Using proc = New Process()
            proc.StartInfo.FileName = psPath
            proc.StartInfo.Arguments = argStr
            proc.StartInfo.WorkingDirectory = workingDir
            proc.StartInfo.UseShellExecute = False
            proc.StartInfo.RedirectStandardOutput = True
            proc.StartInfo.RedirectStandardError = True
            proc.StartInfo.CreateNoWindow = True
            proc.StartInfo.StandardOutputEncoding = Encoding.UTF8
            proc.StartInfo.StandardErrorEncoding = Encoding.UTF8

            ' 设置环境变量
            For Each kvp In envVars
                proc.StartInfo.EnvironmentVariables(kvp.Key) = kvp.Value
            Next

            AddHandler proc.OutputDataReceived, Sub(s, e)
                If e.Data IsNot Nothing Then result.StdOut &= e.Data & Environment.NewLine
            End Sub
            AddHandler proc.ErrorDataReceived, Sub(s, e)
                If e.Data IsNot Nothing Then result.StdErr &= e.Data & Environment.NewLine
            End Sub

            Try
                proc.Start()
                proc.BeginOutputReadLine()
                proc.BeginErrorReadLine()
                Await Task.Run(Sub() proc.WaitForExit())
                result.ExitCode = proc.ExitCode
                result.Success = proc.ExitCode = 0
            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = $"启动 PowerShell 失败: {ex.Message}"
            End Try
        End Using

        Return result
    End Function

    ''' <summary>
    ''' 执行 Shell 脚本 (bash/sh)
    ''' </summary>
    Private Shared Async Function ExecuteShellAsync(
        scriptPath As String,
        workingDir As String,
        args As Dictionary(Of String, String),
        envVars As Dictionary(Of String, String)) As Task(Of ScriptExecutionResult)

        Dim result As New ScriptExecutionResult()

        ' 查找 bash.exe
        Dim bashPath = FindBash()
        If String.IsNullOrEmpty(bashPath) Then
            result.Success = False
            result.ErrorMessage = "未找到 Bash 解释器，请确保已安装 Git Bash 或 WSL"
            Return result
        End If

        ' 构建命令行参数
        Dim arguments = $"""{scriptPath}"""
        Dim argsStr = ""
        If args IsNot Nothing Then
            argsStr = " " & String.Join(" ", args.Values.Select(Function(v) $"""{v}"""))
        End If

        Using proc = New Process()
            proc.StartInfo.FileName = bashPath
            proc.StartInfo.Arguments = $"-c ""cd '{workingDir}' && '{scriptPath}'{argsStr}"""
            proc.StartInfo.WorkingDirectory = workingDir
            proc.StartInfo.UseShellExecute = False
            proc.StartInfo.RedirectStandardOutput = True
            proc.StartInfo.RedirectStandardError = True
            proc.StartInfo.CreateNoWindow = True
            proc.StartInfo.StandardOutputEncoding = Encoding.UTF8
            proc.StartInfo.StandardErrorEncoding = Encoding.UTF8

            ' 设置环境变量
            For Each kvp In envVars
                proc.StartInfo.EnvironmentVariables(kvp.Key) = kvp.Value
            Next

            AddHandler proc.OutputDataReceived, Sub(s, e)
                If e.Data IsNot Nothing Then result.StdOut &= e.Data & Environment.NewLine
            End Sub
            AddHandler proc.ErrorDataReceived, Sub(s, e)
                If e.Data IsNot Nothing Then result.StdErr &= e.Data & Environment.NewLine
            End Sub

            Try
                proc.Start()
                proc.BeginOutputReadLine()
                proc.BeginErrorReadLine()
                Await Task.Run(Sub() proc.WaitForExit())
                result.ExitCode = proc.ExitCode
                result.Success = proc.ExitCode = 0
            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = $"启动 Bash 失败: {ex.Message}"
            End Try
        End Using

        Return result
    End Function

    ''' <summary>
    ''' 执行 Batch 脚本
    ''' </summary>
    Private Shared Async Function ExecuteBatchAsync(
        scriptPath As String,
        workingDir As String,
        args As Dictionary(Of String, String),
        envVars As Dictionary(Of String, String)) As Task(Of ScriptExecutionResult)

        Dim result As New ScriptExecutionResult()

        Dim batchArgs = ""
        If args IsNot Nothing Then
            batchArgs = " " & String.Join(" ", args.Values.Select(Function(v) $"""{v}"""))
        End If

        Using proc = New Process()
            proc.StartInfo.FileName = "cmd.exe"
            proc.StartInfo.Arguments = $"/c ""{scriptPath}{batchArgs}"""
            proc.StartInfo.WorkingDirectory = workingDir
            proc.StartInfo.UseShellExecute = False
            proc.StartInfo.RedirectStandardOutput = True
            proc.StartInfo.RedirectStandardError = True
            proc.StartInfo.CreateNoWindow = True
            proc.StartInfo.StandardOutputEncoding = Encoding.UTF8
            proc.StartInfo.StandardErrorEncoding = Encoding.UTF8

            ' 设置环境变量
            For Each kvp In envVars
                proc.StartInfo.EnvironmentVariables(kvp.Key) = kvp.Value
            Next

            AddHandler proc.OutputDataReceived, Sub(s, e)
                If e.Data IsNot Nothing Then result.StdOut &= e.Data & Environment.NewLine
            End Sub
            AddHandler proc.ErrorDataReceived, Sub(s, e)
                If e.Data IsNot Nothing Then result.StdErr &= e.Data & Environment.NewLine
            End Sub

            Try
                proc.Start()
                proc.BeginOutputReadLine()
                proc.BeginErrorReadLine()
                Await Task.Run(Sub() proc.WaitForExit())
                result.ExitCode = proc.ExitCode
                result.Success = proc.ExitCode = 0
            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = $"启动 Batch 失败: {ex.Message}"
            End Try
        End Using

        Return result
    End Function

    ''' <summary>
    ''' 查找 Python 解释器
    ''' </summary>
    Private Shared Function FindPython() As String
        ' 先尝试 python3，再尝试 python
        Dim candidates = New String() {"python3", "python", "py"}
        For Each candidate In candidates
            Try
                Using proc = New Process()
                    proc.StartInfo.FileName = candidate
                    proc.StartInfo.Arguments = "--version"
                    proc.StartInfo.UseShellExecute = False
                    proc.StartInfo.RedirectStandardOutput = True
                    proc.StartInfo.CreateNoWindow = True
                    proc.Start()
                    proc.WaitForExit()
                    If proc.ExitCode = 0 Then
                        Return candidate
                    End If
                End Using
            Catch
            End Try
        Next

        ' Windows 下尝试查找注册表
        Try
            Using key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Python\PythonCore\3.x\InstallPath")
                If key IsNot Nothing Then
                    Dim installPath = key.GetValue(Nothing)?.ToString()
                    If Not String.IsNullOrEmpty(installPath) Then
                        Dim pythonExe = Path.Combine(installPath, "python.exe")
                        If File.Exists(pythonExe) Then Return pythonExe
                    End If
                End If
            End Using
        Catch
        End Try

        Return ""
    End Function

    ''' <summary>
    ''' 查找 PowerShell 解释器
    ''' </summary>
    Private Shared Function FindPowerShell() As String
        Dim paths = New String() {
            "powershell.exe",
            "pwsh.exe",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell\v1.0\powershell.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "PowerShell\7\pwsh.exe")
        }

        For Each p In paths
            If File.Exists(p) Then Return p
        Next

        Return "powershell.exe"  ' 尝试直接调用，系统 PATH 应该能找到
    End Function

    ''' <summary>
    ''' 查找 Bash 解释器
    ''' </summary>
    Private Shared Function FindBash() As String
        Dim paths = New String() {
            "bash.exe",
            "C:\Program Files\Git\bin\bash.exe",
            "C:\Program Files (x86)\Git\bin\bash.exe"
        }

        For Each p In paths
            If File.Exists(p) Then Return p
        Next

        Return "bash.exe"  ' 尝试直接调用，系统 PATH 应该能找到
    End Function

    ''' <summary>
    ''' 执行 Skill 的默认脚本（便捷方法）
    ''' </summary>
    Public Shared Async Function ExecuteDefaultScriptAsync(
        skill As SkillFileDefinition,
        Optional args As Dictionary(Of String, String) = Nothing) As Task(Of ScriptExecutionResult)

        Dim defaultScript = skill.GetDefaultScript()
        If defaultScript Is Nothing Then
            Return New ScriptExecutionResult() With {
                .Success = False,
                .ErrorMessage = $"Skill '{skill.Name}' 没有可执行的默认脚本"
            }
        End If

        Return Await ExecuteScriptAsync(defaultScript, args, skill.FilePath)
    End Function

    ''' <summary>
    ''' 根据 Skill 名称和脚本名称执行脚本
    ''' </summary>
    Public Shared Async Function ExecuteSkillScriptByNameAsync(
        skillName As String,
        scriptFileName As String,
        Optional args As Dictionary(Of String, String) = Nothing) As Task(Of ScriptExecutionResult)

        Dim allSkills = SkillsDirectoryService.GetAllSkills()
        Dim skill = allSkills.FirstOrDefault(Function(s) s.Name.Equals(skillName, StringComparison.OrdinalIgnoreCase))
        If skill Is Nothing Then
            Return New ScriptExecutionResult() With {
                .Success = False,
                .ErrorMessage = $"未找到 Skill: {skillName}"
            }
        End If

        Dim script = skill.Scripts.FirstOrDefault(Function(s) s.FileName.Equals(scriptFileName, StringComparison.OrdinalIgnoreCase))
        If script Is Nothing Then
            Return New ScriptExecutionResult() With {
                .Success = False,
                .ErrorMessage = $"未找到脚本: {scriptFileName} (在 Skill {skillName} 中)"
            }
        End If

        Return Await ExecuteScriptAsync(script, args, skill.FilePath)
    End Function

    ''' <summary>
    ''' 获取可用脚本解释器的状态信息
    ''' </summary>
    Public Shared Function GetInterpreterStatus() As Dictionary(Of String, Boolean)
        Return New Dictionary(Of String, Boolean) From {
            {"python", Not String.IsNullOrEmpty(FindPython())},
            {"powershell", Not String.IsNullOrEmpty(FindPowerShell())},
            {"bash", Not String.IsNullOrEmpty(FindBash())},
            {"batch", True}  ' cmd.exe 始终可用
        }
    End Function
End Class