' ShareRibbon\Common\JsUtil.vb
' JavaScript 字符串转义工具

Public Class JsUtil
    ''' <summary>转义文本以安全嵌入 JavaScript 模板字符串</summary>
    Public Shared Function EscapeForJs(text As String) As String
        Return text.Replace("\", "\\").Replace("`", "\`").Replace("$", "\$").Replace(vbCr, "").Replace(vbLf, "\n")
    End Function
End Class
