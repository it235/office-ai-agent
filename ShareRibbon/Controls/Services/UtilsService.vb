' ShareRibbon\Controls\Services\UtilsService.vb
' 工具服务：通用工具方法封装

Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq

''' <summary>
''' 工具服务，封装通用工具方法
''' </summary>
Public Class UtilsService

    ''' <summary>
    ''' 获取文本前20个字符并清理为有效文件名
    ''' </summary>
    Public Shared Function GetFirst10Characters(text As String) As String
        If String.IsNullOrEmpty(text) Then Return String.Empty
        Dim result As String = If(text.Length > 20, text.Substring(0, 20), text)
        Dim invalidChars As Char() = System.IO.Path.GetInvalidFileNameChars()
        For Each invalidChar In invalidChars
            result = result.Replace(invalidChar, "_"c)
        Next
        result = result.Replace(" ", "_").Replace(".", "_").Replace(",", "_").
                        Replace(":", "_").Replace("?", "_").Replace("!", "_")
        Return result
    End Function

    ''' <summary>
    ''' 转义字符串以用于 JavaScript
    ''' </summary>
    Public Shared Function EscapeJavaScriptString(input As String) As String
        If String.IsNullOrEmpty(input) Then Return String.Empty
        Return input.Replace("\", "\\").
                     Replace("'", "\'").
                     Replace("""", "\""").
                     Replace(vbCr, "\r").
                     Replace(vbLf, "\n").
                     Replace(vbTab, "\t")
    End Function

    ''' <summary>
    ''' 转义问题字符串（用于 JSON 构建）
    ''' </summary>
    Public Shared Function StripQuestion(question As String) As String
        Return question.Replace("\", "\\").Replace("""", "\""").
                        Replace(vbCr, "\r").Replace(vbLf, "\n").
                        Replace(vbTab, "\t").Replace(vbBack, "\b").
                        Replace(Chr(12), "\f")
    End Function

    ''' <summary>
    ''' Base64 解码
    ''' </summary>
    Public Shared Function DecodeBase64(base64 As String) As String
        Dim bytes = Convert.FromBase64String(base64)
        Return Encoding.UTF8.GetString(bytes)
    End Function

    ''' <summary>
    ''' 反转义 HTML 内容
    ''' </summary>
    Public Shared Function UnescapeHtmlContent(htmlContent As String) As String
        If String.IsNullOrEmpty(htmlContent) Then Return String.Empty
        Return htmlContent.
            Replace("\\n", vbLf).
            Replace("\\r", vbCr).
            Replace("\\t", vbTab).
            Replace("\""", """").
            Replace("\\", "\")
    End Function

    ''' <summary>
    ''' 从文本中提取 JSON 数组
    ''' </summary>
    Public Shared Function TryExtractJsonArrayFromText(text As String) As JArray
        Try
            If String.IsNullOrWhiteSpace(text) Then Return Nothing
            Dim m As Match = Regex.Match(text, "\[.*\]", RegexOptions.Singleline)
            If m.Success Then
                Dim jsonCandidate As String = m.Value.Trim()
                Try
                    Return JArray.Parse(jsonCandidate)
                Catch
                    Return Nothing
                End Try
            End If
        Catch
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' 创建错误响应 JSON
    ''' </summary>
    Public Shared Function CreateErrorResponse(errorMessage As String) As JObject
        Return New JObject From {
            {"isError", True},
            {"content", New JArray From {
                New JObject From {
                    {"type", "text"},
                    {"text", errorMessage}
                }
            }}
        }
    End Function

    ''' <summary>
    ''' 获取保存HTML时需要注入的必要JavaScript代码
    ''' </summary>
    Public Shared Function GetEssentialJavaScript() As String
        Return "
<script>
// 代码复制功能
function copyCode(button) {
    const codeBlock = button.closest('.code-block');
    const codeElement = codeBlock.querySelector('code');
    const code = codeElement.textContent;
    const textarea = document.createElement('textarea');
    textarea.value = code;
    textarea.style.position = 'fixed';
    textarea.style.opacity = '0';
    document.body.appendChild(textarea);
    try {
        textarea.select();
        textarea.setSelectionRange(0, 99999);
        document.execCommand('copy');
        const originalText = button.innerHTML;
        button.innerHTML = '已复制';
        setTimeout(() => { button.innerHTML = originalText; }, 2000);
    } catch (err) {
        console.error('复制失败:', err);
        alert('复制失败');
    } finally {
        document.body.removeChild(textarea);
    }
}

// 聊天消息引用展开/折叠功能
function toggleChatMessageReference(headerElement) {
    const container = headerElement.closest('.chat-message-references');
    if (container) {
        container.classList.toggle('collapsed');
        const arrow = headerElement.querySelector('.chat-message-reference-arrow');
        if (arrow) {
            arrow.innerHTML = container.classList.contains('collapsed') ? '&#9658;' : '&#9660;';
        }
    }
}

// 页面初始化
document.addEventListener('DOMContentLoaded', function() {
    document.querySelectorAll('.code-toggle-label').forEach(label => {
        label.onclick = function(e) {
            e.stopPropagation();
            const preElement = this.nextElementSibling;
            if (preElement && preElement.tagName.toLowerCase() === 'pre') {
                preElement.classList.toggle('collapsed');
                this.textContent = preElement.classList.contains('collapsed') ? '点击展开代码' : '点击折叠代码';
            }
        };
    });
    document.querySelectorAll('pre.collapsible').forEach(preElement => {
        preElement.onclick = function(e) {
            if (e.target.closest('.code-button') || e.target.closest('.code-buttons')) return;
            e.stopPropagation();
            this.classList.toggle('collapsed');
            const toggleLabel = this.previousElementSibling;
            if (toggleLabel && toggleLabel.classList.contains('code-toggle-label')) {
                toggleLabel.textContent = this.classList.contains('collapsed') ? '点击展开代码' : '点击折叠代码';
            }
        };
    });
    document.querySelectorAll('.chat-message-reference-header').forEach(header => {
        header.onclick = function(e) {
            e.preventDefault();
            e.stopPropagation();
            toggleChatMessageReference(this);
        };
    });
    document.querySelectorAll('.reasoning-header').forEach(header => {
        header.onclick = function() {
            const container = this.closest('.reasoning-container');
            if (container) container.classList.toggle('collapsed');
        };
    });
});

if (document.readyState !== 'loading') {
    const event = new Event('DOMContentLoaded');
    document.dispatchEvent(event);
}
</script>"
    End Function

    ''' <summary>
    ''' 获取 VSTO 桥接脚本（用于 WebView2 与 VB.NET 通信）
    ''' </summary>
    Public Shared Function GetVstoBridgeScript() As String
        Return "
        window.vsto = {
            executeCode: function(code, language,preview) {
                window.chrome.webview.postMessage({
                    type: 'executeCode',
                    code: code,
                    language: language,
                    executecodePreview: preview
                });
                return true;
            },
            checkedChange: function(thisProperty,checked) {
                return window.chrome.webview.postMessage({
                    type: 'checkedChange',
                    isChecked: checked,
                    property: thisProperty
                });
            },
            sendMessage: function(payload) {
                let messageToSend;
                if (typeof payload === 'string') {
                    messageToSend = { type: 'sendMessage', value: payload };
                } else {
                    messageToSend = payload;
                }
                window.chrome.webview.postMessage(messageToSend);
                return true;
            },
            saveSettings: function(settingsObject){
                return window.chrome.webview.postMessage({
                    type: 'saveSettings',
                    topicRandomness: settingsObject.topicRandomness,
                    contextLimit: settingsObject.contextLimit,
                    selectedCell: settingsObject.selectedCell,
                    executeCodePreview: settingsObject.executeCodePreview,
                });
            }
        };
    "
    End Function

    ''' <summary>
    ''' 发送 HTTP POST 请求
    ''' </summary>
    Public Shared Async Function SendHttpRequestAsync(apiUrl As String, apiKey As String, requestBody As String) As Threading.Tasks.Task(Of String)
        Try
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12
            Using client As New System.Net.Http.HttpClient()
                client.Timeout = TimeSpan.FromSeconds(120)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim content As New System.Net.Http.StringContent(requestBody, Encoding.UTF8, "application/json")
                Dim response As System.Net.Http.HttpResponseMessage = Await client.PostAsync(apiUrl, content)
                response.EnsureSuccessStatusCode()
                Return Await response.Content.ReadAsStringAsync()
            End Using
        Catch ex As Exception
            Debug.WriteLine($"HTTP请求失败: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' 处理 VBA 相关的 COM 异常
    ''' </summary>
    Public Shared Sub HandleVbaException(ex As Runtime.InteropServices.COMException)
        If ex.Message.Contains("程序访问不被信任") OrElse
           ex.Message.Contains("Programmatic access to Visual Basic Project is not trusted") Then
            ShowVbaTrustDialog()
        Else
            System.Windows.Forms.MessageBox.Show("执行 VBA 代码时出错: " & ex.Message, "错误",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        End If
    End Sub

    ''' <summary>
    ''' 显示 VBA 信任设置对话框
    ''' </summary>
    Public Shared Sub ShowVbaTrustDialog()
        System.Windows.Forms.MessageBox.Show(
            "无法执行 VBA 代码，请按以下步骤设置：" & vbCrLf & vbCrLf &
            "1. 点击 '文件' -> '选项' -> '信任中心'" & vbCrLf &
            "2. 点击 '信任中心设置'" & vbCrLf &
            "3. 选择 '宏设置'" & vbCrLf &
            "4. 勾选 '信任对 VBA 项目对象模型的访问'",
            "需要设置信任中心权限",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Warning)
    End Sub

    ''' <summary>
    ''' 将本地 CSS 资源内联到 HTML 中（用于离线历史文件）
    ''' </summary>
    Public Shared Function InlineCssResources(html As String) As String
        ' 内联 styles.css
        Dim stylesPattern As String = "<link[^>]*href=[""']https://officeai\.local/css/styles\.css[""'][^>]*/?>"
        html = Regex.Replace(html, stylesPattern, "<style>" & My.Resources.styles & "</style>", RegexOptions.IgnoreCase)

        ' 内联 github.min.css (highlight.js theme)
        Dim githubPattern As String = "<link[^>]*href=[""']https://officeai\.local/css/github\.min\.css[""'][^>]*/?>"
        html = Regex.Replace(html, githubPattern, "<style>" & My.Resources.github_min & "</style>", RegexOptions.IgnoreCase)

        Return html
    End Function

End Class
