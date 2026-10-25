Imports System.Diagnostics
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Microsoft.Web.WebView2.WinForms

Public Class WebViewBridge
    Private ReadOnly _browser As WebView2
    Private ReadOnly _ensureInitialized As Func(Of Task)

    Public Sub New(browser As WebView2, ensureInitialized As Func(Of Task))
        If browser Is Nothing Then Throw New ArgumentNullException(NameOf(browser))
        If ensureInitialized Is Nothing Then Throw New ArgumentNullException(NameOf(ensureInitialized))
        _browser = browser
        _ensureInitialized = ensureInitialized
    End Sub

    Public Async Function ExecuteScriptAsync(js As String) As Task
        Await _ensureInitialized()

        Await UiDispatcher.InvokeAsync(_browser,
            Async Function()
                If _browser.CoreWebView2 Is Nothing Then Return
                Await _browser.ExecuteScriptAsync(js)
            End Function)
    End Function

    Public Async Function WaitForRendererMapAsync(uuid As String) As Task
        If String.IsNullOrWhiteSpace(uuid) Then Return

        For i = 1 To 30
            Try
                Dim checkJs = $"Boolean(window.rendererMap && window.rendererMap['{EscapeJavaScriptString(uuid)}']);"
                Dim result = Await UiDispatcher.InvokeAsync(Of String)(_browser,
                    Async Function()
                        If _browser.CoreWebView2 Is Nothing Then Return "false"
                        Return Await _browser.ExecuteScriptAsync(checkJs)
                    End Function)

                If result IsNot Nothing AndAlso result.Trim().ToLowerInvariant().Contains("true") Then
                    Debug.WriteLine($"[WebViewBridge.WaitForRendererMapAsync] rendererMap[{uuid}] ready, retry={i}")
                    Return
                End If
            Catch ex As Exception
                Debug.WriteLine($"[WebViewBridge.WaitForRendererMapAsync] Check failed: {ex.Message}")
            End Try

            Await Task.Delay(100)
        Next

        Debug.WriteLine($"[WebViewBridge.WaitForRendererMapAsync] rendererMap[{uuid}] wait timed out")
    End Function

    Public Async Function GetFullHtmlContentAsync() As Task(Of String)
        Return Await UiDispatcher.InvokeAsync(Of String)(_browser,
            Async Function()
                Await EnsureCoreWebView2InitializedAsync()

                Dim js As String = "
                (function(){
                    const serializer = new XMLSerializer();
                    return serializer.serializeToString(document.documentElement);
                })();"

                Dim rawResult As String = Await _browser.CoreWebView2.ExecuteScriptAsync(js)
                Dim decodedHtml As String = UtilsService.UnescapeHtmlContent(rawResult)
                decodedHtml = decodedHtml.TrimStart(""""c).TrimEnd(""""c)

                Dim bottomBarPattern As String = "<div[^>]*id=[""']chat-bottom-bar[""'][^>]*>.*?</div>\s*</div>\s*</div>"
                decodedHtml = Regex.Replace(decodedHtml, bottomBarPattern, "", RegexOptions.Singleline)

                Dim scriptPattern As String = "<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>"
                decodedHtml = Regex.Replace(decodedHtml, scriptPattern, String.Empty, RegexOptions.IgnoreCase)

                decodedHtml = UtilsService.InlineCssResources(decodedHtml)
                Dim essentialScript As String = UtilsService.GetEssentialJavaScript()

                If decodedHtml.Contains("</body>") Then
                    decodedHtml = decodedHtml.Replace("</body>", essentialScript & Environment.NewLine & "</body>")
                Else
                    decodedHtml &= essentialScript
                End If

                Return decodedHtml
            End Function)
    End Function

    Private Async Function EnsureCoreWebView2InitializedAsync() As Task
        If _browser.CoreWebView2 Is Nothing Then
            Await _browser.EnsureCoreWebView2Async()
        End If
    End Function

    Private Shared Function EscapeJavaScriptString(input As String) As String
        Return UtilsService.EscapeJavaScriptString(input)
    End Function
End Class
