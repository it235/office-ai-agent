Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Reflection.Emit
Imports System.Text
Imports System.Text.JSON
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Windows.Forms
Imports System.Windows.Forms.ListBox
Imports Markdig
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports HtmlAgilityPack
Imports HtmlDocument = HtmlAgilityPack.HtmlDocument
Imports Timer = System.Windows.Forms.Timer
Public MustInherit Class BaseDataCapturePane
    Inherits UserControl
    ' 添加成员变量
    Private isNavigating As Boolean = False
    Private navigationTimer As Timer
    Private Const NAVIGATION_TIMEOUT As Integer = 10000 ' 10秒超时

    Private domSelectionMode As Boolean = False
    Private selectedDomPath As String = ""

    Private isInitialized As Boolean = False
    Private isWebViewInitialized As Boolean = False
    Private pendingUrl As String = Nothing

    Private isCapturing As Boolean = False

    ' 在构造函数或初始化方法中初始化定时器
    Private Sub InitializeNavigationTimer()
        navigationTimer = New Timer With {
            .Interval = NAVIGATION_TIMEOUT,
            .Enabled = False
        }
        AddHandler navigationTimer.Tick, AddressOf OnNavigationTimeout
    End Sub

    Protected Async Function InitializeWebView2() As Task
        If isInitialized Then
            Debug.WriteLine("WebView2 already initialized, skipping...")
            Return
        End If

        isInitialized = True
        Try
            Debug.WriteLine("Starting WebView2 initialization...")

            ' 初始化导航定时器
            InitializeNavigationTimer()

            ' 自定义用户数据目录
            Dim userDataFolder As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MyAppWebView2Cache")

            If Not Directory.Exists(userDataFolder) Then
                Directory.CreateDirectory(userDataFolder)
            End If

            ' 创建 WebView2 环境
            Dim env = Await CoreWebView2Environment.CreateAsync(
            Nothing, userDataFolder, New CoreWebView2EnvironmentOptions())

            ' 初始化 WebView2
            Await ChatBrowser.EnsureCoreWebView2Async(env)

            ' 确保 CoreWebView2 已初始化
            If ChatBrowser.CoreWebView2 IsNot Nothing Then
                Debug.WriteLine("CoreWebView2 initialized successfully")

                ' 设置 WebView2 的安全选项 - 允许弹窗
                With ChatBrowser.CoreWebView2.Settings
                    .IsScriptEnabled = True
                    .AreDefaultScriptDialogsEnabled = True
                    .IsWebMessageEnabled = True
                    .AreDevToolsEnabled = True
                    .AreHostObjectsAllowed = True
                    .IsGeneralAutofillEnabled = True
                End With

                ' 允许弹窗权限
                ' 允许弹窗权限 - 修复后的代码
                AddHandler ChatBrowser.CoreWebView2.PermissionRequested, AddressOf OnPermissionRequested

                ' 移除现有的事件处理程序（如果有）
                RemoveEventHandlers()

                ' 添加新的事件处理程序
                AddEventHandlers()

                isWebViewInitialized = True
                Debug.WriteLine("WebView2 initialization completed successfully")

            Else
                Throw New Exception("CoreWebView2 initialization failed")
            End If

        Catch ex As Exception
            isInitialized = False
            isWebViewInitialized = False
            Debug.WriteLine($"WebView2 initialization failed: {ex.Message}")
            Throw
        End Try
    End Function

    ' 添加权限请求处理方法
    Private Sub OnPermissionRequested(sender As Object, e As CoreWebView2PermissionRequestedEventArgs)
        ' 允许所有权限请求（包括弹窗）
        Select Case e.PermissionKind
            Case CoreWebView2PermissionKind.Camera,
             CoreWebView2PermissionKind.Microphone,
             CoreWebView2PermissionKind.Geolocation,
             CoreWebView2PermissionKind.Notifications
                e.State = CoreWebView2PermissionState.Allow
            Case Else
                e.State = CoreWebView2PermissionState.Allow
        End Select
    End Sub


    ' 新增：事件处理程序管理方法
    Private Sub AddEventHandlers()
        Debug.WriteLine("Adding event handlers...")

        ' 导航事件
        AddHandler ChatBrowser.CoreWebView2.NavigationStarting,
            Sub(s, args)
                Debug.WriteLine($"Navigation starting to: {args.Uri}")
                UrlTextBox.Text = args.Uri
                ' 禁用导航按钮并启动超时计时器
                SetNavigationState(True)
            End Sub

        AddHandler ChatBrowser.CoreWebView2.NavigationCompleted,
            Sub(s, args)
                Debug.WriteLine($"Navigation completed: {args.IsSuccess}")
                ' 停止计时器并恢复按钮状态
                SetNavigationState(False)
                If Not args.IsSuccess Then
                    ' 获取更详细的错误信息
                    Dim errorStatus = ChatBrowser.CoreWebView2.GetDevToolsProtocolEventReceiver("Network.loadingFailed")
                    Debug.WriteLine($"Navigation failed with status: {errorStatus}")
                    MessageBox.Show("页面加载失败，请检查网络连接或重试", "警告",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Else
                    Debug.WriteLine("页面加载成功")
                    ' 可以在这里添加成功加载的处理逻辑
                End If
            End Sub

        ' 处理新窗口打开请求，重定向到当前窗口
        AddHandler ChatBrowser.CoreWebView2.NewWindowRequested,
            Sub(s, args)
                ' 取消新窗口打开
                args.Handled = True
                ' 在当前窗口导航到目标URL
                ChatBrowser.CoreWebView2.Navigate(args.Uri)
                Debug.WriteLine($"拦截到新窗口请求，已重定向到当前窗口: {args.Uri}")
            End Sub

        ' WebMessage事件
        AddHandler ChatBrowser.CoreWebView2.WebMessageReceived,
            AddressOf WebView2_MessageReceived

        ' 按钮事件
        AddHandler NavigateButton.Click, AddressOf NavigateButton_Click
        AddHandler CaptureButton.Click, AddressOf CaptureButton_Click
        AddHandler UrlTextBox.KeyPress, AddressOf UrlTextBox_KeyPress
        AddHandler SelectDomButton.Click, AddressOf SelectDomButton_Click

        ' 添加前进后退按钮事件
        AddHandler BackButton.Click, AddressOf BackButton_Click
        AddHandler ForwardButton.Click, AddressOf ForwardButton_Click

        ' 监听历史记录状态变化
        AddHandler ChatBrowser.CoreWebView2.HistoryChanged,
        Sub(s, args)
            UpdateNavigationButtons()
        End Sub

        Debug.WriteLine("Event handlers added successfully")
    End Sub

    ' 移除事件处理程序时也需要移除新增的事件
    Private Sub RemoveEventHandlers()
        Try
            If ChatBrowser?.CoreWebView2 IsNot Nothing Then
                RemoveHandler ChatBrowser.CoreWebView2.PermissionRequested, AddressOf OnPermissionRequested
                RemoveHandler ChatBrowser.CoreWebView2.WebMessageReceived,
                AddressOf WebView2_MessageReceived
                RemoveHandler ChatBrowser.CoreWebView2.NewWindowRequested,
                Sub(s, args)
                    args.Handled = True
                    ChatBrowser.CoreWebView2.Navigate(args.Uri)
                End Sub
                RemoveHandler ChatBrowser.CoreWebView2.HistoryChanged,
                Sub(s, args)
                    UpdateNavigationButtons()
                End Sub
            End If

            RemoveHandler NavigateButton.Click, AddressOf NavigateButton_Click
            RemoveHandler CaptureButton.Click, AddressOf CaptureButton_Click
            RemoveHandler UrlTextBox.KeyPress, AddressOf UrlTextBox_KeyPress
            RemoveHandler SelectDomButton.Click, AddressOf SelectDomButton_Click
            RemoveHandler BackButton.Click, AddressOf BackButton_Click
            RemoveHandler ForwardButton.Click, AddressOf ForwardButton_Click
        Catch ex As Exception
            Debug.WriteLine($"Error removing event handlers: {ex.Message}")
        End Try
    End Sub

    ' 添加前进后退按钮点击事件处理
    Private Sub BackButton_Click(sender As Object, e As EventArgs)
        If ChatBrowser?.CoreWebView2 IsNot Nothing AndAlso ChatBrowser.CoreWebView2.CanGoBack Then
            ChatBrowser.CoreWebView2.GoBack()
        End If
    End Sub

    Private Sub ForwardButton_Click(sender As Object, e As EventArgs)
        If ChatBrowser?.CoreWebView2 IsNot Nothing AndAlso ChatBrowser.CoreWebView2.CanGoForward Then
            ChatBrowser.CoreWebView2.GoForward()
        End If
    End Sub

    ' 更新前进后退按钮状态
    Private Sub UpdateNavigationButtons()
        If ChatBrowser?.CoreWebView2 IsNot Nothing Then
            BackButton.Enabled = ChatBrowser.CoreWebView2.CanGoBack
            ForwardButton.Enabled = ChatBrowser.CoreWebView2.CanGoForward
        Else
            BackButton.Enabled = False
            ForwardButton.Enabled = False
        End If
    End Sub

    ' 添加导航状态控制方法
    Private Sub SetNavigationState(isNavigating As Boolean)
        Me.isNavigating = isNavigating
        NavigateButton.Enabled = Not isNavigating
        UrlTextBox.Enabled = Not isNavigating

        If isNavigating Then
            navigationTimer.Start()
        Else
            navigationTimer.Stop()
        End If
    End Sub

    ' 添加超时处理方法
    Private Sub OnNavigationTimeout(sender As Object, e As EventArgs)
        ' 在UI线程中执行
        If Me.InvokeRequired Then
            Me.Invoke(Sub() OnNavigationTimeout(sender, e))
            Return
        End If

        navigationTimer.Stop()
        If isNavigating Then
            ' 如果仍在导航状态，则强制恢复按钮
            SetNavigationState(False)
            Debug.WriteLine("Navigation timeout - restoring button state")
            'MessageBox.Show("页面加载超时，已恢复导航按钮", "提示",
            '              MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub NavigateButton_Click(sender As Object, e As EventArgs)
        NavigateToUrl(UrlTextBox.Text)
    End Sub

    Private Sub UrlTextBox_KeyPress(sender As Object, e As KeyPressEventArgs)
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True
            NavigateToUrl(UrlTextBox.Text)
        End If
    End Sub

    ' 修改：导航方法添加更多调试信息
    Private Sub NavigateToUrl(url As String)
        If String.IsNullOrWhiteSpace(url) Then
            Debug.WriteLine("Navigation cancelled: Empty URL")
            Return
        End If

        ' 如果正在导航中，忽略新的导航请求
        If isNavigating Then
            Debug.WriteLine("Navigation in progress, ignoring new request")
            Return
        End If

        Try
            ' 标准化URL
            If Not url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) AndAlso
               Not url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                url = "https://" & url
            End If

            If Not isWebViewInitialized Then
                pendingUrl = url
                Return
            End If

            If ChatBrowser.CoreWebView2 IsNot Nothing Then
                ChatBrowser.CoreWebView2.Navigate(url)
            Else
                Debug.WriteLine("Navigation failed: CoreWebView2 is null")
                MessageBox.Show("WebView2 组件未就绪", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            Debug.WriteLine($"Navigation error: {ex.Message}")
            MessageBox.Show($"导航失败: {ex.Message}", "错误",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' 确保在发生错误时恢复按钮状态
            SetNavigationState(False)
        End Try
    End Sub
    Private Async Sub CaptureButton_Click(sender As Object, e As EventArgs)
        If isCapturing Then
            MessageBox.Show("正在抓取中，请稍候...", "提示")
            Return
        End If

        Try
            ' 显示抓取类型选择对话框
            Dim captureTypeResult = MessageBox.Show(
            "请选择抓取内容类型：" & vbCrLf & vbCrLf &
            "【是】- 抓取文本内容（智能解析，保持格式）" & vbCrLf &
            "【否】- 抓取HTML代码（包含完整标签结构）" & vbCrLf &
            "【取消】- 取消操作",
            "选择抓取类型",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button1)

            Select Case captureTypeResult
                Case DialogResult.Yes
                    ' 抓取文本内容
                    Await CaptureTextContent()
                Case DialogResult.No
                    ' 抓取HTML代码
                    Await CaptureHtmlContent()
                Case DialogResult.Cancel
                    ' 取消操作
                    Return
            End Select

        Catch ex As Exception
            Debug.WriteLine($"抓取内容时出错: {ex.Message}")
            MessageBox.Show($"抓取内容时出错: {ex.Message}", "错误")
        End Try
    End Sub

    ' 抓取文本内容方法
    Private Async Function CaptureTextContent() As Task
        Try
            isCapturing = True
            CaptureButton.Text = "抓取文本中..."
            CaptureButton.Enabled = False

            ' 等待页面完全渲染（对Vue/React等SPA很重要）
            Await EnsurePageFullyLoaded()

            ' 获取处理后的页面内容
            Dim extractedContent As String

            If Not String.IsNullOrEmpty(selectedDomPath) Then
                ' 使用选定的DOM路径抓取特定元素
                extractedContent = Await ExtractSelectedElement()
            Else
                ' 抓取整个页面内容
                extractedContent = Await ExtractFullPageContent()
            End If

            If Not String.IsNullOrEmpty(extractedContent) Then
                ' 清理和格式化内容
                Dim cleanContent = CleanAndFormatContent(extractedContent)

                ' 添加页面信息头部
                Dim pageInfo = Await GetPageMetaInfo()
                Dim finalContent = pageInfo & vbCrLf & vbCrLf & cleanContent

                HandleExtractedContent(finalContent)

                Debug.WriteLine($"成功抓取文本内容，长度: {finalContent.Length} 字符")
                'MessageBox.Show("文本内容抓取完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("未能获取到有效的文本内容", "提示")
            End If

        Catch ex As Exception
            Debug.WriteLine($"抓取文本内容时出错: {ex.Message}")
            MessageBox.Show($"抓取文本内容时出错: {ex.Message}", "错误")
        Finally
            isCapturing = False
            CaptureButton.Text = "抓取内容"
            CaptureButton.Enabled = True
        End Try
    End Function

    ' 确保页面完全加载（包括动态内容）
    Private Async Function EnsurePageFullyLoaded() As Task
        Try
            ' 等待页面加载完成的脚本
            Dim waitScript = "
        (async function() {
            // 等待基本DOM加载
            if (document.readyState !== 'complete') {
                await new Promise(resolve => {
                    if (document.readyState === 'complete') {
                        resolve();
                    } else {
                        window.addEventListener('load', resolve, { once: true });
                    }
                });
            }

            // 等待Vue/React等框架渲染完成
            await new Promise(resolve => setTimeout(resolve, 1000));

            // 检查是否有动态加载的内容
            let retryCount = 0;
            const maxRetries = 5;
            
            while (retryCount < maxRetries) {
                // 触发滚动以加载懒加载内容
                window.scrollTo(0, document.body.scrollHeight);
                await new Promise(resolve => setTimeout(resolve, 500));
                
                // 检查是否有加载指示器
                const loadingIndicators = document.querySelectorAll(
                    '[class*=""loading""], [class*=""spinner""], [class*=""skeleton""], .loading, .spinner'
                );
                
                if (loadingIndicators.length === 0) {
                    break;
                }
                
                retryCount++;
                await new Promise(resolve => setTimeout(resolve, 1000));
            }

            // 滚动回顶部
            window.scrollTo(0, 0);
            
            // 最后等待一点时间确保渲染完成
            await new Promise(resolve => setTimeout(resolve, 500));
            
            return 'ready';
        })();
        "

            Dim result = Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(waitScript)
            Debug.WriteLine($"页面加载等待完成: {result}")

        Catch ex As Exception
            Debug.WriteLine($"等待页面加载时出错: {ex.Message}")
        End Try
    End Function

    ' 抓取选定元素内容
    Private Async Function ExtractSelectedElement() As Task(Of String)
        Try
            Dim script = $"
        (function() {{
            const element = document.querySelector('{selectedDomPath}');
            if (!element) return null;
            
            // 获取元素的可见文本内容，保持基本结构
            function getCleanText(el) {{
                if (!el) return '';
                
                // 克隆元素避免修改原DOM
                const clone = el.cloneNode(true);
                
                // 移除脚本和样式标签
                const scripts = clone.querySelectorAll('script, style, noscript');
                scripts.forEach(s => s.remove());
                
                // 处理特殊元素
                const links = clone.querySelectorAll('a[href]');
                links.forEach(link => {{
                    const href = link.getAttribute('href');
                    if (href && !href.startsWith('#')) {{
                        link.textContent = `${{link.textContent}} [${{href}}]`;
                    }}
                }});
                
                const images = clone.querySelectorAll('img[src], img[alt]');
                images.forEach(img => {{
                    const alt = img.getAttribute('alt') || '';
                    const src = img.getAttribute('src') || '';
                    img.outerHTML = `[图片: ${{alt || '无描述'}}]${{src ? ' - ' + src : ''}}`;
                }});
                
                return clone.innerText || clone.textContent || '';
            }}
            
            return getCleanText(element);
        }})();
        "

            Dim result = Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            If Not String.IsNullOrEmpty(result) AndAlso result <> "null" Then
                Return JsonConvert.DeserializeObject(Of String)(result)
            End If

            Return String.Empty

        Catch ex As Exception
            Debug.WriteLine($"抓取选定元素时出错: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ' 抓取完整页面内容
    Private Async Function ExtractFullPageContent() As Task(Of String)
        Try
            Dim script = "
        (function() {
            // 智能内容提取函数
            function extractMainContent() {
                // 移除不需要的元素
                const elementsToRemove = [
                    'script', 'style', 'noscript', 'iframe', 'embed', 'object',
                    '[style*=""display: none""]', '[style*=""visibility: hidden""]',
                    '.advertisement', '.ads', '.ad', '[class*=""popup""]', 
                    '[class*=""modal""]', '.cookie-banner', '.cookie-notice'
                ];
                
                const clone = document.cloneNode(true);
                const body = clone.body || clone.documentElement;
                
                elementsToRemove.forEach(selector => {
                    const elements = body.querySelectorAll(selector);
                    elements.forEach(el => el.remove());
                });
                
                // 查找主要内容区域
                const mainSelectors = [
                    'main', 'article', '[role=""main""]', '.main-content', 
                    '.content', '.post', '.entry', '#content', '#main'
                ];
                
                let mainContent = null;
                for (const selector of mainSelectors) {
                    const element = body.querySelector(selector);
                    if (element && element.textContent.trim().length > 100) {
                        mainContent = element;
                        break;
                    }
                }
                
                const sourceElement = mainContent || body;
                
                // 递归处理元素，保持结构
                function processElement(element, level = 0) {
                    if (!element || level > 10) return '';
                    
                    const tagName = element.tagName ? element.tagName.toLowerCase() : '';
                    let result = '';
                    
                    // 处理不同类型的元素
                    switch (tagName) {
                        case 'h1':
                        case 'h2':
                        case 'h3':
                        case 'h4':
                        case 'h5':
                        case 'h6':
                            const headerLevel = '='.repeat(parseInt(tagName.charAt(1)));
                            result += `\n\n${headerLevel} ${element.textContent.trim()} ${headerLevel}\n\n`;
                            break;
                            
                        case 'p':
                            const text = element.textContent.trim();
                            if (text) result += `${text}\n\n`;
                            break;
                            
                        case 'br':
                            result += '\n';
                            break;
                            
                        case 'hr':
                            result += '\n---\n\n';
                            break;
                            
                        case 'blockquote':
                            const quote = element.textContent.trim();
                            if (quote) result += `> ${quote}\n\n`;
                            break;
                            
                        case 'ul':
                        case 'ol':
                            const listItems = element.querySelectorAll('li');
                            listItems.forEach((li, index) => {
                                const bullet = tagName === 'ul' ? '•' : `${index + 1}.`;
                                result += `${bullet} ${li.textContent.trim()}\n`;
                            });
                            result += '\n';
                            break;
                            
                        case 'table':
                            result += processTable(element);
                            break;
                            
                        case 'a':
                            const href = element.getAttribute('href');
                            const linkText = element.textContent.trim();
                            if (href && !href.startsWith('#') && linkText) {
                                result += `${linkText} [${href}]`;
                            } else {
                                result += linkText;
                            }
                            break;
                            
                        case 'img':
                            const alt = element.getAttribute('alt') || '';
                            const src = element.getAttribute('src') || '';
                            result += `[图片: ${alt || '无描述'}]${src ? ` - ${src}` : ''}\n`;
                            break;
                            
                        case 'div':
                        case 'section':
                        case 'article':
                        case 'aside':
                        case 'header':
                        case 'footer':
                        case 'nav':
                            // 递归处理子元素
                            for (const child of element.children) {
                                result += processElement(child, level + 1);
                            }
                            // 如果没有子元素，处理文本内容
                            if (element.children.length === 0) {
                                const text = element.textContent.trim();
                                if (text && text.length > 0) {
                                    result += `${text}\n\n`;
                                }
                            }
                            break;
                            
                        default:
                            // 处理其他元素的文本内容
                            const childText = element.textContent.trim();
                            if (childText && element.children.length === 0) {
                                result += `${childText} `;
                            } else {
                                // 递归处理子元素
                                for (const child of element.children) {
                                    result += processElement(child, level + 1);
                                }
                            }
                            break;
                    }
                    
                    return result;
                }
                
                // 处理表格
                function processTable(table) {
                    let tableResult = '\n';
                    const rows = table.querySelectorAll('tr');
                    
                    rows.forEach((row, rowIndex) => {
                        const cells = row.querySelectorAll('td, th');
                        const cellTexts = Array.from(cells).map(cell => 
                            cell.textContent.trim().replace(/\s+/g, ' ')
                        );
                        
                        if (cellTexts.length > 0) {
                            tableResult += `| ${cellTexts.join(' | ')} |\n`;
                            
                            // 添加表头分隔线
                            if (rowIndex === 0 && row.querySelector('th')) {
                                const separator = cellTexts.map(() => '---').join(' | ');
                                tableResult += `| ${separator} |\n`;
                            }
                        }
                    });
                    
                    return tableResult + '\n';
                }
                
                return processElement(sourceElement);
            }
            
            return extractMainContent();
        })();
        "

            Dim result = Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            If Not String.IsNullOrEmpty(result) AndAlso result <> "null" Then
                Return JsonConvert.DeserializeObject(Of String)(result)
            End If

            Return String.Empty

        Catch ex As Exception
            Debug.WriteLine($"抓取完整页面时出错: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ' 获取页面元信息
    Private Async Function GetPageMetaInfo() As Task(Of String)
        Try
            Dim script = "
        (function() {
            const title = document.title || '无标题';
            const url = window.location.href;
            const description = document.querySelector('meta[name=""description""]')?.content || '';
            const author = document.querySelector('meta[name=""author""]')?.content || '';
            const publishDate = document.querySelector('meta[property=""article:published_time""]')?.content || 
                               document.querySelector('meta[name=""date""]')?.content || '';
            
            let info = `页面标题: ${title}\n`;
            info += `页面链接: ${url}\n`;
            if (description) info += `页面描述: ${description}\n`;
            if (author) info += `作者: ${author}\n`;
            if (publishDate) info += `发布时间: ${publishDate}\n`;
            info += `抓取时间: ${new Date().toLocaleString('zh-CN')}\n`;
            info += '----------------------------------------';
            
            return info;
        })();
        "

            Dim result = Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            If Not String.IsNullOrEmpty(result) AndAlso result <> "null" Then
                Return JsonConvert.DeserializeObject(Of String)(result)
            End If

            Return "页面信息获取失败"

        Catch ex As Exception
            Debug.WriteLine($"获取页面信息时出错: {ex.Message}")
            Return "页面信息获取失败"
        End Try
    End Function

    ' 清理和格式化内容
    Private Function CleanAndFormatContent(content As String) As String
        If String.IsNullOrWhiteSpace(content) Then
            Return String.Empty
        End If

        Try
            ' 基本清理
            Dim cleaned = content.Trim()

            ' 移除多余的空行（超过2个连续换行符替换为2个）
            cleaned = Regex.Replace(cleaned, "\n{3,}", vbCrLf & vbCrLf)

            ' 移除行首行尾空格
            Dim lines = cleaned.Split({vbCrLf, vbLf}, StringSplitOptions.None)
            For i = 0 To lines.Length - 1
                lines(i) = lines(i).Trim()
            Next
            cleaned = String.Join(vbCrLf, lines)

            ' 移除开头和结尾的多余换行
            cleaned = cleaned.Trim()

            ' 限制最大长度（防止内容过长）
            If cleaned.Length > 50000 Then
                cleaned = cleaned.Substring(0, 50000) & vbCrLf & vbCrLf & "... [内容过长，已截断]"
            End If

            Return cleaned

        Catch ex As Exception
            Debug.WriteLine($"清理内容时出错: {ex.Message}")
            Return content
        End Try
    End Function

    ' 修改抓取HTML代码方法 - 支持大文件处理
    Private Async Function CaptureHtmlContent() As Task
        Try
            isCapturing = True
            CaptureButton.Text = "抓取HTML中..."
            CaptureButton.Enabled = False

            ' 等待页面完全渲染
            Await EnsurePageFullyLoaded()

            ' 获取渲染后的HTML内容
            Dim htmlContent As String

            If Not String.IsNullOrEmpty(selectedDomPath) Then
                ' 抓取选定元素的HTML
                htmlContent = Await ExtractSelectedElementHtml()
            Else
                ' 抓取整个页面的HTML
                htmlContent = Await ExtractFullPageHtml()
            End If

            If Not String.IsNullOrEmpty(htmlContent) Then
                Debug.WriteLine($"获取到HTML内容，长度: {htmlContent.Length} 字符")

                ' 检查内容大小，决定处理方式
                If htmlContent.Length > 200000 Then ' 200KB
                    Dim choice = MessageBox.Show(
                    $"HTML内容很大 ({htmlContent.Length:N0} 字符)，选择处理方式：" & vbCrLf & vbCrLf &
                    "【是】- 完整显示（可能较慢）" & vbCrLf &
                    "【否】- 保存到文件" & vbCrLf &
                    "【取消】- 截断显示",
                    "大文件处理",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question)

                    Select Case choice
                        Case DialogResult.Yes
                            ' 完整处理
                            Await ProcessLargeHtmlContent(htmlContent)
                        Case DialogResult.No
                            ' 保存到文件
                            Await SaveHtmlToFile(htmlContent)
                        Case DialogResult.Cancel
                            ' 截断处理
                            Await ProcessTruncatedHtmlContent(htmlContent)
                    End Select
                Else
                    ' 正常处理
                    Await ProcessNormalHtmlContent(htmlContent)
                End If

            Else
                MessageBox.Show("未能获取到有效的HTML内容", "提示")
            End If

        Catch ex As Exception
            Debug.WriteLine($"抓取HTML内容时出错: {ex.Message}")
            MessageBox.Show($"抓取HTML内容时出错: {ex.Message}", "错误")
        Finally
            isCapturing = False
            CaptureButton.Text = "抓取内容"
            CaptureButton.Enabled = True
        End Try
    End Function

    ' 处理大HTML内容
    Private Async Function ProcessLargeHtmlContent(htmlContent As String) As Task
        Try
            ' 在后台线程处理格式化以避免UI卡死
            Dim formattedHtml As String = Nothing

            Await Task.Run(Sub()
                               formattedHtml = FormatHtmlContent(htmlContent)
                           End Sub)

            ' 添加页面信息头部
            Dim pageInfo = Await GetPageMetaInfo()
            Dim finalContent = pageInfo & vbCrLf & vbCrLf &
                          "========== 大型HTML代码内容 ==========" & vbCrLf &
                          $"原始大小: {htmlContent.Length:N0} 字符" & vbCrLf &
                          $"格式化后: {formattedHtml.Length:N0} 字符" & vbCrLf & vbCrLf &
                          formattedHtml

            HandleExtractedContent(finalContent)
            MessageBox.Show("大型HTML内容处理完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"处理大型HTML内容时出错: {ex.Message}", "错误")
        End Try
    End Function

    ' 保存HTML到文件
    Private Async Function SaveHtmlToFile(htmlContent As String) As Task
        Try
            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "HTML文件 (*.html)|*.html|所有文件 (*.*)|*.*"
                saveDialog.DefaultExt = "html"
                saveDialog.FileName = $"captured_page_{DateTime.Now:yyyyMMdd_HHmmss}.html"

                If saveDialog.ShowDialog() = DialogResult.OK Then
                    ' 在后台线程保存文件
                    Await Task.Run(Sub()
                                       File.WriteAllText(saveDialog.FileName, htmlContent, Encoding.UTF8)
                                   End Sub)

                    ' 同时在文档中插入文件信息
                    Dim pageInfo = Await GetPageMetaInfo()
                    Dim fileInfo = $"{pageInfo}{vbCrLf}{vbCrLf}" &
                              "========== HTML文件已保存 ==========" & vbCrLf &
                              $"文件路径: {saveDialog.FileName}" & vbCrLf &
                              $"文件大小: {htmlContent.Length:N0} 字符" & vbCrLf &
                              $"保存时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}" & vbCrLf

                    HandleExtractedContent(fileInfo)
                    MessageBox.Show($"HTML内容已保存到：{vbCrLf}{saveDialog.FileName}", "保存成功")
                End If
            End Using

        Catch ex As Exception
            MessageBox.Show($"保存HTML文件时出错: {ex.Message}", "错误")
        End Try
    End Function

    ' 处理截断的HTML内容
    Private Async Function ProcessTruncatedHtmlContent(htmlContent As String) As Task
        Try
            ' 截断处理，但保留结构完整性
            Dim truncatedHtml = TruncateHtmlSafely(htmlContent, 100000)
            Dim formattedHtml = FormatHtmlContent(truncatedHtml)

            Dim pageInfo = Await GetPageMetaInfo()
            Dim finalContent = pageInfo & vbCrLf & vbCrLf &
                          "========== HTML代码内容（已截断） ==========" & vbCrLf &
                          $"原始大小: {htmlContent.Length:N0} 字符" & vbCrLf &
                          $"显示大小: {formattedHtml.Length:N0} 字符" & vbCrLf & vbCrLf &
                          formattedHtml & vbCrLf & vbCrLf &
                          "... [内容过长，已截断，完整内容请选择保存到文件]"

            HandleExtractedContent(finalContent)
            MessageBox.Show("HTML内容已截断显示！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"处理截断HTML内容时出错: {ex.Message}", "错误")
        End Try
    End Function

    ' 处理正常大小的HTML内容
    Private Async Function ProcessNormalHtmlContent(htmlContent As String) As Task
        Try
            Dim formattedHtml = FormatHtmlContent(htmlContent)
            Dim pageInfo = Await GetPageMetaInfo()
            Dim finalContent = pageInfo & vbCrLf & vbCrLf &
                          "========== HTML代码内容 ==========" & vbCrLf & vbCrLf &
                          formattedHtml

            HandleExtractedContent(finalContent)
            'MessageBox.Show("HTML代码抓取完成！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"处理HTML内容时出错: {ex.Message}", "错误")
        End Try
    End Function

    ' 安全截断HTML内容，保持标签完整性
    Private Function TruncateHtmlSafely(htmlContent As String, maxLength As Integer) As String
        If htmlContent.Length <= maxLength Then
            Return htmlContent
        End If

        Try
            ' 找到合适的截断点（在标签之间）
            Dim truncatePos = maxLength
            While truncatePos > 0 AndAlso htmlContent(truncatePos) <> ">"c
                truncatePos -= 1
            End While

            If truncatePos > 0 Then
                Return htmlContent.Substring(0, truncatePos + 1)
            Else
                Return htmlContent.Substring(0, maxLength)
            End If

        Catch ex As Exception
            Return htmlContent.Substring(0, Math.Min(maxLength, htmlContent.Length))
        End Try
    End Function

    ' 抓取选定元素的HTML代码
    Private Async Function ExtractSelectedElementHtml() As Task(Of String)
        Try
            Dim script = $"
        (function() {{
            const element = document.querySelector('{selectedDomPath}');
            if (!element) return null;
            
            // 获取元素的完整HTML，包括所有属性和子元素
            function getElementHtml(el) {{
                if (!el) return '';
                
                // 克隆元素以避免修改原DOM
                const clone = el.cloneNode(true);
                
                // 清理一些可能影响显示的内联样式
                function cleanElement(element) {{
                    // 移除一些调试相关的属性
                    element.removeAttribute('data-reactid');
                    element.removeAttribute('data-react-checksum');
                    
                    // 递归处理子元素
                    Array.from(element.children).forEach(child => {{
                        cleanElement(child);
                    }});
                }}
                
                cleanElement(clone);
                
                // 返回格式化的HTML
                return clone.outerHTML;
            }}
            
            return getElementHtml(element);
        }})();
        "

            Dim result = Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            If Not String.IsNullOrEmpty(result) AndAlso result <> "null" Then
                Return JsonConvert.DeserializeObject(Of String)(result)
            End If

            Return String.Empty

        Catch ex As Exception
            Debug.WriteLine($"抓取选定元素HTML时出错: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ' 最简单有效的解决方案 - 去掉JSON序列化的限制
    Private Async Function ExtractFullPageHtml() As Task(Of String)
        Try
            ' 首先检查内容大小
            Dim sizeCheckScript = "document.documentElement.outerHTML.length;"
            Dim sizeResult = Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(sizeCheckScript)
            Dim contentSize = CInt(sizeResult)

            Debug.WriteLine($"HTML内容大小: {contentSize:N0} 字符")

            If contentSize > 2000000 Then ' 2MB
                ' 内容太大，提示用户
                Dim choice = MessageBox.Show(
                $"HTML内容非常大 ({contentSize:N0} 字符)，可能会导致传输问题。" & vbCrLf & vbCrLf &
                "建议选择：" & vbCrLf &
                "【是】- 尝试完整传输（可能失败）" & vbCrLf &
                "【否】- 使用简化HTML" & vbCrLf &
                "【取消】- 取消操作",
                "内容过大警告",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning)

                Select Case choice
                    Case DialogResult.No
                        Return Await ExtractSimplifiedHtml()
                    Case DialogResult.Cancel
                        Return String.Empty
                End Select
            End If

            ' 尝试直接获取（无JSON包装）
            Dim script = "
        (function() {
            const docClone = document.cloneNode(true);
            
            // 清理不需要的元素
            const elementsToRemove = [
                'script[src*=""webview""]',
                'script[src*=""devtools""]', 
                '.inspector-overlay',
                '[data-inspector]'
            ];
            
            elementsToRemove.forEach(selector => {
                const elements = docClone.querySelectorAll(selector);
                elements.forEach(el => el.remove());
            });
            
            const head = docClone.head ? docClone.head.outerHTML : '';
            const body = docClone.body ? docClone.body.outerHTML : '';
            
            let fullHtml = '<!DOCTYPE html>\\n<html';
            
            if (document.documentElement.attributes) {
                Array.from(document.documentElement.attributes).forEach(attr => {
                    fullHtml += ` ${attr.name}=\""${attr.value}\""`;
                });
            }
            
            fullHtml += '>\\n' + head + '\\n' + body + '\\n</html>';
            
            return fullHtml;
        })();
        "

            Dim result = Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            If Not String.IsNullOrEmpty(result) AndAlso result <> "null" Then
                ' 直接反序列化，不需要额外包装
                Dim htmlContent = JsonConvert.DeserializeObject(Of String)(result)
                Debug.WriteLine($"HTML传输成功，最终长度: {htmlContent.Length:N0} 字符")
                Return htmlContent
            End If

            Return String.Empty

        Catch ex As Exception
            Debug.WriteLine($"抓取HTML时出错: {ex.Message}")
            ' 降级到简化版本
            Return CreateSimplifiedHtml()
        End Try
    End Function

    ' 添加缺失的简化HTML创建方法
    Private Function CreateSimplifiedHtml() As String
        Try
            Dim simplified = $"<!DOCTYPE html>
<html>
<head>
    <title>页面内容（简化版）</title>
    <meta charset='utf-8'>
</head>
<body>
    <h1>页面内容（简化版）</h1>
    <p>原始页面过大，这是简化版本</p>
    <p>页面URL: {ChatBrowser.CoreWebView2?.Source}</p>
    <p>页面标题: {ChatBrowser.CoreWebView2?.DocumentTitle}</p>
    <hr>
    <div>由于原始HTML内容过大，无法完整传输。请考虑使用文本抓取模式或保存到文件。</div>
</body>
</html>"

            Return simplified

        Catch ex As Exception
            Debug.WriteLine($"创建简化HTML时出错: {ex.Message}")
            Return "<!DOCTYPE html><html><head><title>错误</title></head><body><h1>HTML抓取失败</h1><p>无法获取页面内容</p></body></html>"
        End Try
    End Function
    ' 简化版HTML提取（移除大部分内容）
    Private Async Function ExtractSimplifiedHtml() As Task(Of String)
        Try
            Dim script = "
        (function() {
            const simplified = document.createElement('html');
            simplified.innerHTML = `
                <head>
                    <title>${document.title}</title>
                    <meta charset='utf-8'>
                </head>
                <body>
                    <h1>页面内容（简化版）</h1>
                    <p>原始页面过大，这是简化版本</p>
                    <p>页面URL: ${window.location.href}</p>
                    <p>页面标题: ${document.title}</p>
                    <hr>
                    <div>${document.body.innerText.substring(0, 5000)}</div>
                </body>
            `;
            return simplified.outerHTML;
        })();
        "

            Dim result = Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            If Not String.IsNullOrEmpty(result) AndAlso result <> "null" Then
                Return JsonConvert.DeserializeObject(Of String)(result)
            End If

            Return String.Empty

        Catch ex As Exception
            Debug.WriteLine($"获取简化HTML时出错: {ex.Message}")
            Return String.Empty
        End Try
    End Function
    ' 格式化HTML内容
    Private Function FormatHtmlContent(htmlContent As String) As String
        If String.IsNullOrWhiteSpace(htmlContent) Then
            Return String.Empty
        End If

        Try
            ' 使用HtmlAgilityPack格式化HTML
            Dim doc As New HtmlDocument()
            doc.LoadHtml(htmlContent)

            ' 格式化设置
            doc.OptionOutputAsXml = False
            doc.OptionAutoCloseOnEnd = True
            doc.OptionFixNestedTags = True

            ' 获取格式化后的HTML
            Dim formattedHtml As String

            Using stringWriter As New StringWriter()
                doc.Save(stringWriter)
                formattedHtml = stringWriter.ToString()
            End Using

            ' 基本的格式化处理
            formattedHtml = formattedHtml.Replace("><", ">" & vbCrLf & "<")

            ' 添加适当的缩进
            Dim lines = formattedHtml.Split({vbCrLf, vbLf}, StringSplitOptions.None)
            Dim indentLevel = 0
            Dim formattedLines As New List(Of String)

            For Each line In lines
                Dim trimmedLine = line.Trim()
                If String.IsNullOrEmpty(trimmedLine) Then Continue For

                ' 减少缩进（闭合标签）
                If trimmedLine.StartsWith("</") Then
                    indentLevel = Math.Max(0, indentLevel - 1)
                End If

                ' 添加缩进
                Dim indent = New String(" "c, indentLevel * 2)
                formattedLines.Add(indent & trimmedLine)

                ' 增加缩进（开放标签，但不是自闭合标签）
                If trimmedLine.StartsWith("<") AndAlso
               Not trimmedLine.StartsWith("</") AndAlso
               Not trimmedLine.EndsWith("/>") AndAlso
               Not trimmedLine.Contains("<img ") AndAlso
               Not trimmedLine.Contains("<br") AndAlso
               Not trimmedLine.Contains("<hr") AndAlso
               Not trimmedLine.Contains("<input ") AndAlso
               Not trimmedLine.Contains("<meta ") AndAlso
               Not trimmedLine.Contains("<link ") Then
                    indentLevel += 1
                End If
            Next

            formattedHtml = String.Join(vbCrLf, formattedLines)

            ' 限制长度
            'If formattedHtml.Length > 100000 Then
            '    formattedHtml = formattedHtml.Substring(0, 100000) &
            '               vbCrLf & vbCrLf & "... [HTML内容过长，已截断]"
            'End If

            Return formattedHtml

        Catch ex As Exception
            Debug.WriteLine($"格式化HTML时出错: {ex.Message}")
            ' 如果格式化失败，返回原始内容
            Return htmlContent
        End Try
    End Function

    ' 添加表格数据模型
    Protected Class TableData
        Public Property Rows As Integer
        Public Property Columns As Integer
        Public Property Data As List(Of List(Of String))
        Public Property Headers As List(Of String)

        Public Sub New()

            Data = New List(Of List(Of String))
            Headers = New List(Of String)

        End Sub
    End Class


    ' 添加抽象方法
    Protected MustOverride Function CreateTable(tableData As TableData) As String

    ' 抽象方法：处理提取的内容（由具体实现类实现）
    Protected MustOverride Sub HandleExtractedContent(content As String)

    ' 选择DOM元素按钮点击事件处理程序
    Private Async Sub SelectDomButton_Click(sender As Object, e As EventArgs)
        Try
            Dim selectScript As String = "
        (function() {
            // 移除旧的选择器
            if(window._domSelector) {
                try {
                    document.removeEventListener('mouseover', window._domSelector.onMouseOver);
                    document.removeEventListener('mouseout', window._domSelector.onMouseOut);
                    document.removeEventListener('click', window._domSelector.onClick);
                    document.removeEventListener('keydown', window._domSelector.onKeyDown);
                    document.removeEventListener('keyup', window._domSelector.onKeyUp);
                    document.removeEventListener('wheel', window._domSelector.onWheel);
                    if(window._domSelector.tip) window._domSelector.tip.remove();
                } catch(e) {
                    console.log('Error removing old selector:', e);
                }
            }

            // 创建新的选择器
            window._domSelector = {
                lastHighlight: null,
                lastParentHighlight: null,
                isShiftKey: false,
                isAltKey: false,  // 改为Alt键
                parentLevel: 0,
                maxParentLevel: 5,
                _lastChildElement: null,
                _currentTarget: null,
                _highlightTimer: null,
                _parentCandidates: [],
                
                // 创建改进的提示框
                createTip: function() {
                    const tip = document.createElement('div');
                    tip.style.cssText = `
                        position: fixed;
                        top: 10px;
                        left: 50%;
                        transform: translateX(-50%);
                        background: linear-gradient(135deg, rgba(0, 0, 0, 0.9), rgba(33, 150, 243, 0.8));
                        color: white;
                        padding: 12px 20px;
                        border-radius: 8px;
                        font-family: Arial, sans-serif;
                        font-size: 14px;
                        z-index: 2147483647;
                        pointer-events: none;
                        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
                        border: 1px solid rgba(255, 255, 255, 0.2);
                    `;
                    tip.innerHTML = `
                        <div style='margin-bottom: 8px; font-weight: bold; color: #FFD700;'>🎯 智能元素选择器</div>
                        <div style='margin-bottom: 4px;'>• 点击选择元素</div>
                        <div style='margin-bottom: 4px;'>• Shift + 点击：选择父元素</div>
                        <div style='margin-bottom: 4px;'>• Alt + 滚轮：调整父元素层级</div>
                        <div style='color: #90EE90;'>• 支持抓取图片和视频</div>
                    `;
                    document.body.appendChild(tip);
                    this.tip = tip;
                },

                // 安全检查函数
                isValidSelector: function() {
                    return window._domSelector && 
                           typeof window._domSelector === 'object' && 
                           !window._domSelector._destroyed;
                },

                // 改进的高亮显示
                highlight: function(element, options = {}) {
                    if (!this.isValidSelector() || !element) return;
                    
                    const { isParent = false, level = 0 } = options;
                    
                    // 清除之前的高亮
                    this.removeHighlight();

                    let target = element;
                    if (isParent) {
                        target = this.getParentAtLevel(element, level);
                    }
                    
                    if (!target) return;

                    this._currentTarget = target;
                    
                    // 根据元素类型设置不同的高亮样式
                    const elementType = this.getElementType(target);
                    const highlightStyle = this.getHighlightStyle(elementType, isParent, level);
                    
                    // 设置高亮样式
                    target.style.transition = 'all 0.3s ease-in-out';
                    target.style.outline = highlightStyle.outline;
                    target.style.outlineOffset = highlightStyle.outlineOffset;
                    target.style.backgroundColor = highlightStyle.backgroundColor;
                    target.style.transform = highlightStyle.transform;
                    
                    // 保存当前高亮的元素
                    this.lastHighlight = target;
                    
                    // 显示详细信息框
                    this.showDetailedInfo(target, target.getBoundingClientRect(), { isParent, level, elementType });

                    // 如果是父元素选择，记住子元素
                    if (isParent) {
                        this._lastChildElement = element;
                    }
                },

                // 获取元素类型
                getElementType: function(element) {
                    if (!element || !element.tagName) return 'text';
                    const tag = element.tagName.toLowerCase();
                    if (tag === 'img') return 'image';
                    if (tag === 'video') return 'video';
                    if (tag === 'audio') return 'audio';
                    if (tag === 'canvas') return 'canvas';
                    if (tag === 'svg') return 'svg';
                    if (tag === 'table') return 'table';
                    if (tag === 'form') return 'form';
                    if (element.querySelector && element.querySelector('img, video, audio')) return 'media-container';
                    return 'text';
                },

                // 获取高亮样式
                getHighlightStyle: function(elementType, isParent, level) {
                    const baseStyles = {
                        image: {
                            outline: '4px solid #FF5722',
                            outlineOffset: '3px',
                            backgroundColor: 'rgba(255, 87, 34, 0.1)',
                            transform: 'scale(1.02)'
                        },
                        video: {
                            outline: '4px solid #9C27B0',
                            outlineOffset: '3px',
                            backgroundColor: 'rgba(156, 39, 176, 0.1)',
                            transform: 'scale(1.02)'
                        },
                        audio: {
                            outline: '4px solid #FF9800',
                            outlineOffset: '3px',
                            backgroundColor: 'rgba(255, 152, 0, 0.1)',
                            transform: 'scale(1.02)'
                        },
                        'media-container': {
                            outline: '4px dashed #E91E63',
                            outlineOffset: '3px',
                            backgroundColor: 'rgba(233, 30, 99, 0.1)',
                            transform: 'scale(1.01)'
                        },
                        table: {
                            outline: '4px solid #4CAF50',
                            outlineOffset: '2px',
                            backgroundColor: 'rgba(76, 175, 80, 0.1)',
                            transform: 'none'
                        },
                        default: {
                            outline: '3px solid #2196F3',
                            outlineOffset: '2px',
                            backgroundColor: 'rgba(33, 150, 243, 0.1)',
                            transform: 'none'
                        }
                    };

                    let style = baseStyles[elementType] || baseStyles.default;
                    
                    if (isParent) {
                        const parentColors = ['#FF9800', '#9C27B0', '#4CAF50', '#F44336', '#00BCD4'];
                        const color = parentColors[level % parentColors.length];
                        style = {
                            outline: `4px dashed ${color}`,
                            outlineOffset: `${3 + level}px`,
                            backgroundColor: `${color}20`,
                            transform: 'none'
                        };
                    }

                    return style;
                },

                // 获取指定层级的父元素
                getParentAtLevel: function(element, level) {
                    let parent = element;
                    for (let i = 0; i < level + 1 && parent && parent !== document.body; i++) {
                        parent = parent.parentElement;
                    }
                    return parent;
                },

                // 修改移除高亮函数 - 增强稳定性
                removeHighlight: function() {
                    try {
                        // 防止重复执行
                        if (this._removing) return;
                        this._removing = true;
        
                        if (this.lastHighlight) {
                            try {
                                this.lastHighlight.style.outline = '';
                                this.lastHighlight.style.outlineOffset = '';
                                this.lastHighlight.style.backgroundColor = '';
                                this.lastHighlight.style.transform = '';
                            } catch(e) {
                                // 忽略样式设置错误
                            }
                            this.lastHighlight = null;
                        }
        
                        if (this.infoBox) {
                            try {
                                if (this.infoBox.parentNode) {
                                    this.infoBox.parentNode.removeChild(this.infoBox);
                                }
                            } catch(e) {
                                // 忽略移除错误
                            }
                            this.infoBox = null;
                        }
        
                    } catch(e) {
                        console.log('Error removing highlight:', e);
                    } finally {
                        // 释放锁
                        setTimeout(() => {
                            this._removing = false;
                        }, 10);
                    }
                },

                // 显示详细信息
                showDetailedInfo: function(element, rect, options = {}) {
                    if (!this.isValidSelector() || !element) return;
                    
                    try {
                        // 防抖处理 - 如果正在显示信息框则跳过
                        if (this._showingInfo) return;
                        this._showingInfo = true;
        
                        // 设置超时自动释放锁
                        setTimeout(() => {
                            this._showingInfo = false;
                        }, 100);
        
                        if (this.infoBox) {
                            try {
                                this.infoBox.remove();
                            } catch(e) {
                                // 忽略移除错误
                            }
                            this.infoBox = null;
                        }
        
                        const { isParent = false, level = 0, elementType = 'text' } = options;
                        
                        const info = document.createElement('div');
                        info.style.cssText = `
                            position: absolute;
                            background: linear-gradient(135deg, rgba(0, 0, 0, 0.95), rgba(33, 150, 243, 0.9));
                            color: white;
                            padding: 12px 20px;
                            border-radius: 8px;
                            font-size: 14px;
                            pointer-events: none;
                            z-index: 2147483647;
                            font-family: Arial, sans-serif;
                            max-width: 450px;
                            min-width: 300px;
                            text-align: left;
                            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
                            border: 1px solid rgba(255, 255, 255, 0.2);
                        `;
        
                        // 安全地获取元素信息
                        const tag = (element.tagName || '').toLowerCase() || 'unknown';
                        const id = element.id ? '#' + element.id : '';
                        const classes = element.classList ? Array.from(element.classList).slice(0, 3).map(c => '.' + c).join('') : '';
                        const textContent = (element.textContent || element.innerText || '').trim();
                        const contentPreview = textContent.length > 100 ? textContent.substring(0, 100) + '...' : textContent;

                        // 获取媒体信息
                        const mediaInfo = this.getMediaInfo(element);
                        
                        info.innerHTML = `
                            <div style='font-size: 16px; margin-bottom: 8px; color: #FFD700;'>
                                ${this.getElementIcon(elementType)} &lt;${tag}${id}${classes}&gt;
                                ${isParent ? ` (父元素层级: ${level + 1})` : ''}
                            </div>
                            ${mediaInfo ? `<div style='margin-bottom: 8px; color: #90EE90;'>${mediaInfo}</div>` : ''}
                            <div style='font-size: 13px; opacity: 0.9; margin-bottom: 8px;'>
                                ${contentPreview}
                            </div>
                            <div style='font-size: 12px; color: #FFA726;'>
                                ${this.getActionHint(elementType, isParent)}
                            </div>
                        `;
        
                        // 智能定位
                        this.positionInfoBox(info, rect);
                        
                        // 安全地添加到DOM
                        try {
                            document.body.appendChild(info);
                            this.infoBox = info;
                        } catch(e) {
                            console.log('Error adding info box to DOM:', e);
                        }
                    } catch(e) {
                        console.log('Error showing detailed info:', e);
                    } finally {
                        // 确保释放锁
                        setTimeout(() => {
                            this._showingInfo = false;
                        }, 50);
                    }
                },

                // 新增安全的媒体信息获取函数
                getMediaInfoSafe: function(element) {
                    if (!element || !element.tagName) return '';
    
                    try {
                        const tag = element.tagName.toLowerCase();
                        let info = '';
        
                        if (tag === 'img') {
                            const width = element.naturalWidth || element.width || 0;
                            const height = element.naturalHeight || element.height || 0;
                            const alt = (element.alt || '无描述').substring(0, 20);
                            info = `📷 图片: ${alt} (${width}×${height})`;
                        } else if (tag === 'video') {
                            const width = element.videoWidth || element.width || 0;
                            const height = element.videoHeight || element.height || 0;
                            const duration = element.duration ? Math.round(element.duration) + 's' : '未知';
                            info = `🎬 视频: ${duration} (${width}×${height})`;
                        } else if (tag === 'audio') {
                            const duration = element.duration ? Math.round(element.duration) + 's' : '未知';
                            info = `🎵 音频: ${duration}`;
                        } else if (element.querySelectorAll) {
                            // 限制查询范围以防止性能问题
                            const mediaElements = element.querySelectorAll('img, video, audio');
                            if (mediaElements.length > 0 && mediaElements.length < 50) {
                                info = `📦 包含 ${mediaElements.length} 个媒体元素`;
                            }
                        }
        
                        return info;
                    } catch(e) {
                        return '';
                    }
                },

                // 获取媒体信息
                getMediaInfo: function(element) {
                    if (!element || !element.tagName) return '';
                    
                    try {
                        const tag = element.tagName.toLowerCase();
                        let info = '';
                        
                        if (tag === 'img') {
                            const src = element.src || '';
                            const alt = element.alt || '无描述';
                            const width = element.naturalWidth || element.width || 0;
                            const height = element.naturalHeight || element.height || 0;
                            info = `📷 图片: ${alt} (${width}×${height})`;
                        } else if (tag === 'video') {
                            const src = element.src || (element.querySelector && element.querySelector('source') ? element.querySelector('source').src : '');
                            const duration = element.duration ? Math.round(element.duration) + 's' : '未知';
                            const width = element.videoWidth || element.width || 0;
                            const height = element.videoHeight || element.height || 0;
                            info = `🎬 视频: ${duration} (${width}×${height})`;
                        } else if (tag === 'audio') {
                            const src = element.src || (element.querySelector && element.querySelector('source') ? element.querySelector('source').src : '');
                            const duration = element.duration ? Math.round(element.duration) + 's' : '未知';
                            info = `🎵 音频: ${duration}`;
                        }
                        
                        // 检查是否包含媒体元素
                        if (!info && element.querySelectorAll) {
                            const mediaElements = element.querySelectorAll('img, video, audio');
                            if (mediaElements.length > 0) {
                                info = `📦 包含 ${mediaElements.length} 个媒体元素`;
                            }
                        }
                        
                        return info;
                    } catch(e) {
                        console.log('Error getting media info:', e);
                        return '';
                    }
                },

                // 获取元素图标
                getElementIcon: function(elementType) {
                    const icons = {
                        image: '📷',
                        video: '🎬',
                        audio: '🎵',
                        canvas: '🎨',
                        svg: '🖼️',
                        table: '📊',
                        form: '📝',
                        'media-container': '📦',
                        text: '📄'
                    };
                    return icons[elementType] || '📄';
                },

                // 获取操作提示
                getActionHint: function(elementType, isParent) {
                    if (isParent) {
                        return '🔄 Alt+滚轮调整层级，点击确认选择';
                    }
                    
                    const hints = {
                        image: '🖼️ 将保存图片到文档',
                        video: '🎬 将保存视频信息和链接',
                        audio: '🎵 将保存音频信息和链接',
                        table: '📊 将转换为Word表格',
                        'media-container': '📦 将保存所有媒体内容',
                        text: '📄 将保存文本内容'
                    };
                    return hints[elementType] || '点击选择此元素';
                },

                // 智能定位信息框
                positionInfoBox: function(info, rect) {
                    try {
                        const infoWidth = 450;
                        const infoHeight = 120;
                        const margin = 10;
                        
                        let left = rect.left + (rect.width / 2) - (infoWidth / 2);
                        let top = rect.top + window.scrollY - infoHeight - margin;
                        
                        // 确保不超出视口
                        left = Math.max(margin, Math.min(left, window.innerWidth - infoWidth - margin));
                        
                        if (top < margin) {
                            top = rect.bottom + window.scrollY + margin;
                        }
                        
                        info.style.left = left + 'px';
                        info.style.top = top + 'px';
                    } catch(e) {
                        console.log('Error positioning info box:', e);
                    }
                },

                // 事件处理程序
                onMouseOver: function(e) {
                    if (!window._domSelector || !window._domSelector.isValidSelector()) return;
                    
                    try {
                        e.stopPropagation();
                        const target = e.target;
                        
                        if (target === window._domSelector._currentTarget) return;
                        
                        const options = {
                            isParent: window._domSelector.isShiftKey || window._domSelector.isAltKey,
                            level: window._domSelector.parentLevel
                        };
                        
                        window._domSelector.highlight(target, options);
                    } catch(e) {
                        console.log('Error in onMouseOver:', e);
                    }
                },

                onMouseOut: function(e) {
                    if (!window._domSelector || !window._domSelector.isValidSelector()) return;
                    
                    try {
                        if (window._domSelector._highlightTimer) {
                            clearTimeout(window._domSelector._highlightTimer);
                        }

                        const relatedTarget = e.relatedTarget;
                        if (!window._domSelector.lastHighlight || 
                            !window._domSelector.lastHighlight.contains(relatedTarget)) {
                            window._domSelector.removeHighlight();
                            window._domSelector._currentTarget = null;
                        }

                        e.stopPropagation();
                    } catch(e) {
                        console.log('Error in onMouseOut:', e);
                    }
                },

                // 改进的点击处理
                onClick: function(e) {
                    if (!window._domSelector || !window._domSelector.isValidSelector()) return;
                    
                    try {
                        e.preventDefault();
                        e.stopPropagation();
                        
                        let element = e.target;
                        
                        // 根据按键状态选择元素
                        if (window._domSelector.isShiftKey || window._domSelector.isAltKey) {
                            element = window._domSelector.getParentAtLevel(element, window._domSelector.parentLevel);
                        }
                        
                        if (!element) return;
                        
                        // 收集元素信息
                        const elementInfo = window._domSelector.collectElementInfo(element);
                        
                        if (window.chrome && window.chrome.webview && window.chrome.webview.postMessage) {
                            window.chrome.webview.postMessage({
                                type: 'elementSelected',
                                ...elementInfo
                            });
                        }
                    } catch(e) {
                        console.log('Error in onClick:', e);
                    }
                },

                // 收集元素信息
                collectElementInfo: function(element) {
                    if (!element) return {};
                    
                    try {
                        const tag = element.tagName ? element.tagName.toLowerCase() : 'unknown';
                        const elementType = this.getElementType(element);
                        const rect = element.getBoundingClientRect ? element.getBoundingClientRect() : {};
                        
                        let info = {
                            html: element.outerHTML || '',
                            path: this.getPath(element),
                            tag: tag,
                            text: element.innerText || element.textContent || '',
                            rect: rect,
                            elementType: elementType,
                            isTable: tag === 'table'
                        };
                        
                        // 收集媒体信息
                        if (elementType === 'image') {
                            info.mediaInfo = {
                                src: element.src || '',
                                alt: element.alt || '',
                                width: element.naturalWidth || element.width || 0,
                                height: element.naturalHeight || element.height || 0
                            };
                        } else if (elementType === 'video') {
                            const sourceElement = element.querySelector ? element.querySelector('source') : null;
                            info.mediaInfo = {
                                src: element.src || (sourceElement ? sourceElement.src : ''),
                                poster: element.poster || '',
                                duration: element.duration || 0,
                                width: element.videoWidth || element.width || 0,
                                height: element.videoHeight || element.height || 0
                            };
                        } else if (elementType === 'audio') {
                            const sourceElement = element.querySelector ? element.querySelector('source') : null;
                            info.mediaInfo = {
                                src: element.src || (sourceElement ? sourceElement.src : ''),
                                duration: element.duration || 0
                            };
                        }
                        
                        // 收集包含的媒体元素
                        if (element.querySelectorAll) {
                            const mediaElements = element.querySelectorAll('img, video, audio');
                            if (mediaElements.length > 0) {
                                info.containedMedia = Array.from(mediaElements).map(media => ({
                                    tag: media.tagName ? media.tagName.toLowerCase() : '',
                                    src: media.src || '',
                                    alt: media.alt || '',
                                    width: media.naturalWidth || media.videoWidth || media.width || 0,
                                    height: media.naturalHeight || media.videoHeight || media.height || 0
                                }));
                            }
                        }
                        
                        return info;
                    } catch(e) {
                        console.log('Error collecting element info:', e);
                        return {};
                    }
                },

                // 修改滚轮事件处理 - 更严格的事件阻止
                onWheel: function(e) {
                    if (!window._domSelector || !window._domSelector.isValidSelector()) return;
    
                    try {
                        // 检查是否同时按下 Alt 键
                        if (e.altKey && window._domSelector.isAltKey) {
                            // 立即阻止默认行为和冒泡
                            e.preventDefault();
                            e.stopPropagation();
                            e.stopImmediatePropagation();
            
                            const delta = e.deltaY > 0 ? 1 : -1;
                            window._domSelector.parentLevel = Math.max(0, 
                                Math.min(window._domSelector.maxParentLevel, 
                                    window._domSelector.parentLevel + delta));
            
                            // 重新高亮当前元素
                            if (window._domSelector._currentTarget) {
                                const originalTarget = window._domSelector._lastChildElement || window._domSelector._currentTarget;
                                window._domSelector.highlight(originalTarget, {
                                    isParent: true,
                                    level: window._domSelector.parentLevel
                                });
                            }
            
                            return false; // 额外保险
                        }
                    } catch(e) {
                        console.log('Error in onWheel:', e);
                    }
                },

                // 键盘事件处理
                onKeyDown: function(e) {
                    if (!window._domSelector || !window._domSelector.isValidSelector()) return;
                    
                    try {
                        if (e.key === 'Shift' && !window._domSelector.isShiftKey) {
                            window._domSelector.isShiftKey = true;
                            window._domSelector.parentLevel = 0;
                            
                            if (window._domSelector._currentTarget) {
                                window._domSelector.highlight(window._domSelector._currentTarget, {
                                    isParent: true,
                                    level: window._domSelector.parentLevel
                                });
                            }
                        } else if (e.key === 'Alt' && !window._domSelector.isAltKey) {
                            window._domSelector.isAltKey = true;
                            window._domSelector.parentLevel = 0;
                            
                            if (window._domSelector._currentTarget) {
                                window._domSelector.highlight(window._domSelector._currentTarget, {
                                    isParent: true,
                                    level: window._domSelector.parentLevel
                                });
                            }
                        }
                    } catch(e) {
                        console.log('Error in onKeyDown:', e);
                    }
                },

                onKeyUp: function(e) {
                    if (!window._domSelector || !window._domSelector.isValidSelector()) return;
                    
                    try {
                        if (e.key === 'Shift') {
                            window._domSelector.isShiftKey = false;
                            window._domSelector.parentLevel = 0;
                            
                            if (window._domSelector._lastChildElement) {
                                window._domSelector.highlight(window._domSelector._lastChildElement, { isParent: false });
                            }
                        } else if (e.key === 'Alt') {
                            window._domSelector.isAltKey = false;
                            window._domSelector.parentLevel = 0;
                            
                            if (window._domSelector._lastChildElement) {
                                window._domSelector.highlight(window._domSelector._lastChildElement, { isParent: false });
                            }
                        }
                    } catch(e) {
                        console.log('Error in onKeyUp:', e);
                    }
                },

                // 获取元素路径
                getPath: function(element) {
                    try {
                        const path = [];
                        while(element && element.nodeType === Node.ELEMENT_NODE) {
                            let selector = element.tagName ? element.tagName.toLowerCase() : 'unknown';
                            if(element.id) {
                                selector += '#' + element.id;
                                path.unshift(selector);
                                break;
                            } else {
                                let sibling = element, nth = 1;
                                while(sibling = sibling.previousElementSibling) {
                                    if(sibling.tagName === element.tagName) nth++;
                                }
                                if(nth > 1) selector += ':nth-of-type(' + nth + ')';
                            }
                            path.unshift(selector);
                            element = element.parentNode;
                        }
                        return path.join(' > ');
                    } catch(e) {
                        console.log('Error getting path:', e);
                        return '';
                    }
                },

                // 初始化
                init: function() {
                    try {
                        if (window._domSelector) {
                            window._domSelector.cleanup();
                        }
                        
                        // 添加防卡死标志
                        this._showingInfo = false;
                        this._removing = false;

                        this._destroyed = false;
                        this.createTip();
                        
                        // 绑定事件处理程序
                        this.onMouseOver = this.onMouseOver.bind(this);
                        this.onMouseOut = this.onMouseOut.bind(this);
                        this.onClick = this.onClick.bind(this);
                        this.onKeyDown = this.onKeyDown.bind(this);
                        this.onKeyUp = this.onKeyUp.bind(this);
                        this.onWheel = this.onWheel.bind(this);
                        
                        // 添加事件监听器
                        document.addEventListener('mouseover', this.onMouseOver, true);
                        document.addEventListener('mouseout', this.onMouseOut, true);
                        document.addEventListener('click', this.onClick, true);
                        document.addEventListener('keydown', this.onKeyDown, true);
                        document.addEventListener('keyup', this.onKeyUp, true);
                        document.addEventListener('wheel', this.onWheel, true);
                        
                        document.body.style.cursor = 'crosshair';
                        
                        console.log('DOM selector initialized successfully');
                    } catch(e) {
                        console.log('Error initializing DOM selector:', e);
                    }
                },

                // 清理
                cleanup: function() {
                    try {
                        this._destroyed = true;
                        this.removeHighlight();
                        if (this.tip) this.tip.remove();
                        
                        // 移除所有事件监听器
                        document.removeEventListener('mouseover', this.onMouseOver, true);
                        document.removeEventListener('mouseout', this.onMouseOut, true);
                        document.removeEventListener('click', this.onClick, true);
                        document.removeEventListener('keydown', this.onKeyDown, true);
                        document.removeEventListener('keyup', this.onKeyUp, true);
                        document.removeEventListener('wheel', this.onWheel, true);
                        
                        // 清理所有状态
                        this._currentTarget = null;
                        this._lastChildElement = null;
                        this.lastHighlight = null;
                        this.isShiftKey = false;
                        this.isAltKey = false;
                        this.parentLevel = 0;
                        
                        document.body.style.cursor = '';
                        
                        console.log('DOM selector cleaned up successfully');
                    } catch(e) {
                        console.log('Error cleaning up DOM selector:', e);
                    }
                }
            };

            // 初始化选择器
            window._domSelector.init();
        })();
        "
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(selectScript)

        Catch ex As Exception
            Debug.WriteLine($"DOM选择器错误: {ex.Message}")
            MessageBox.Show($"初始化选择器失败: {ex.Message}", "错误")
        End Try
    End Sub

    ' 修改消息处理程序
    Private Async Sub WebView2_MessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Debug.WriteLine($"收到消息: {e.WebMessageAsJson}")

            Dim message = JsonConvert.DeserializeObject(Of JObject)(e.WebMessageAsJson)
            If message("type")?.ToString() = "elementSelected" Then
            ' 获取完整信息
            selectedDomPath = message("path").ToString()
            Dim html = message("html").ToString()
            Dim text = message("text").ToString()
            Dim tag = message("tag").ToString()
                Dim elementType = If(message("elementType")?.ToString(), "text")

                ' 处理媒体信息
                Dim mediaInfo As JObject = Nothing
            If message("mediaInfo") IsNot Nothing Then
                mediaInfo = DirectCast(message("mediaInfo"), JObject)
            End If
            
            ' 处理包含的媒体元素
            Dim containedMedia As JArray = Nothing
            If message("containedMedia") IsNot Nothing Then
                containedMedia = DirectCast(message("containedMedia"), JArray)
            End If

            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync("
                        if(window._domSelector) {
                            window._domSelector.cleanup();
                            window._domSelector = null;
                        }
                    ")
            
            ' 根据元素类型显示不同的确认对话框
            Select Case elementType
                Case "image"
                    HandleImageSelection(mediaInfo, html, text, selectedDomPath)
                Case "video"
                    HandleVideoSelection(mediaInfo, html, text, selectedDomPath)
                Case "audio"
                    HandleAudioSelection(mediaInfo, html, text, selectedDomPath)
                Case "media-container"
                    HandleMediaContainerSelection(containedMedia, html, text, selectedDomPath)
                Case Else
                    ' 使用原有的确认对话框
                    ShowStandardConfirmDialog(text, tag, selectedDomPath)
            End Select
        End If
        Catch ex As Exception
            Debug.WriteLine($"处理消息错误: {ex.Message}")
            MessageBox.Show($"处理选择消息失败: {ex.Message}", "错误")
        End Try
    End Sub

    ' 处理图片选择
    Private Sub HandleImageSelection(mediaInfo As JObject, html As String, text As String, path As String)
        Dim src = If(mediaInfo("src")?.ToString(), "")
        Dim alt = If(mediaInfo("alt")?.ToString(), "")
        Dim width = If(mediaInfo("width")?.ToString(), "0")
        Dim height = If(mediaInfo("height")?.ToString(), "0")


        Dim message = $"🖼️ 发现图片元素{vbCrLf}描述: {alt}{vbCrLf}尺寸: {width}×{height}{vbCrLf}链接: {src}"

        Dim result = MessageBox.Show(message & vbCrLf & vbCrLf & "是否要抓取此图片？",
                                "图片选择确认",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Question)

        Select Case result
            Case DialogResult.Yes
                DownloadAndInsertImage(src, alt)
            Case DialogResult.No
                OnAiChatRequested($"图片信息: {message}")
            Case DialogResult.Cancel
                selectedDomPath = ""
        End Select
    End Sub

    ' 处理视频选择
    Private Sub HandleVideoSelection(mediaInfo As JObject, html As String, text As String, path As String)
        Dim src = If(mediaInfo("src")?.ToString(), "")
        Dim poster = If(mediaInfo("poster")?.ToString(), "")
        Dim duration = If(mediaInfo("duration")?.ToString(), "0")
        Dim width = If(mediaInfo("width")?.ToString(), "0")
        Dim height = If(mediaInfo("height")?.ToString(), "0")

        Dim message = $"🎬 发现视频元素{vbCrLf}时长: {duration}秒{vbCrLf}尺寸: {width}×{height}{vbCrLf}链接: {src}"

        Dim result = MessageBox.Show(message & vbCrLf & vbCrLf & "是否要抓取此视频信息？",
                                "视频选择确认",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Question)

        Select Case result
            Case DialogResult.Yes
                HandleVideoContent(src, poster, duration, width, height)
            Case DialogResult.No
                OnAiChatRequested($"视频信息: {message}")
            Case DialogResult.Cancel
                selectedDomPath = ""
        End Select
    End Sub

    ' 处理音频选择
    Private Sub HandleAudioSelection(mediaInfo As JObject, html As String, text As String, path As String)
        Dim src = If(mediaInfo("src")?.ToString(), "")
        Dim duration = If(mediaInfo("duration")?.ToString(), "0")

        Dim message = $"🎵 发现音频元素{vbCrLf}时长: {duration}秒{vbCrLf}链接: {src}"

        Dim result = MessageBox.Show(message & vbCrLf & vbCrLf & "是否要抓取此音频信息？",
                                "音频选择确认",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Question)

        Select Case result
            Case DialogResult.Yes
                HandleAudioContent(src, duration)
            Case DialogResult.No
                OnAiChatRequested($"音频信息: {message}")
            Case DialogResult.Cancel
                selectedDomPath = ""
        End Select
    End Sub

    ' 处理包含媒体的容器
    Private Sub HandleMediaContainerSelection(containedMedia As JArray, html As String, text As String, path As String)
        Dim mediaCount = If(containedMedia?.Count, 0)
        Dim message = $"📦 发现包含 {mediaCount} 个媒体元素的容器{vbCrLf}内容预览: {text.Substring(0, Math.Min(text.Length, 100))}"

        Dim result = MessageBox.Show(message & vbCrLf & vbCrLf & "是否要抓取此容器及其媒体内容？",
                                "媒体容器选择确认",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Question)

        Select Case result
            Case DialogResult.Yes
                HandleMediaContainerContent(containedMedia, text)
            Case DialogResult.No
                OnAiChatRequested($"媒体容器信息: {message}")
            Case DialogResult.Cancel
                selectedDomPath = ""
        End Select
    End Sub
    ' 替换标准确认对话框 - 异步版本，避免卡死
    Private Sub ShowStandardConfirmDialog(text As String, tag As String, path As String)
        ' 使用 Task.Run 在后台线程显示对话框
        Task.Run(Sub()
                     Try
                         Dim result As DialogResult

                         ' 在UI线程中显示对话框
                         Me.Invoke(Sub()
                                       Try
                                           Using dialog As New WebSiteContentConfirmDialog(text, tag, path)
                                               result = dialog.ShowDialog(Me)
                                           End Using
                                       Catch ex As Exception
                                           Debug.WriteLine($"对话框显示错误: {ex.Message}")
                                           result = DialogResult.Cancel
                                       End Try
                                   End Sub)

                         ' 处理结果也在UI线程中执行
                         Me.Invoke(Sub()
                                       Select Case result
                                           Case DialogResult.Cancel
                                               selectedDomPath = ""
                                           Case DialogResult.Yes
                                               HandleExtractedContent(text)
                                           Case DialogResult.No
                                               OnAiChatRequested(text)
                                       End Select
                                   End Sub)

                     Catch ex As Exception
                         Debug.WriteLine($"异步对话框处理错误: {ex.Message}")
                         Me.Invoke(Sub()
                                       selectedDomPath = ""
                                   End Sub)
                     End Try
                 End Sub)
    End Sub

    ' 下载并插入图片（需要在子类中实现）
    Protected MustOverride Sub DownloadAndInsertImage(src As String, alt As String)

    ' 处理视频内容（需要在子类中实现）
    Protected MustOverride Sub HandleVideoContent(src As String, poster As String, duration As String, width As String, height As String)

    ' 处理音频内容（需要在子类中实现）
    Protected MustOverride Sub HandleAudioContent(src As String, duration As String)

    ' 处理媒体容器内容（需要在子类中实现）
    Protected MustOverride Sub HandleMediaContainerContent(containedMedia As JArray, text As String)

    ' 添加事件以供子类处理AI聊天请求
    Protected Event AiChatRequested As EventHandler(Of String)
    Protected Sub OnAiChatRequested(content As String)
        RaiseEvent AiChatRequested(Me, content)
    End Sub
End Class