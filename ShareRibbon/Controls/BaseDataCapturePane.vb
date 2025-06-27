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

                ' 设置 WebView2 的安全选项
                With ChatBrowser.CoreWebView2.Settings
                    .IsScriptEnabled = True
                    .AreDefaultScriptDialogsEnabled = True
                    .IsWebMessageEnabled = True
                    .AreDevToolsEnabled = True
                End With

                ' 移除现有的事件处理程序（如果有）
                RemoveEventHandlers()

                ' 添加新的事件处理程序
                AddEventHandlers()

                isWebViewInitialized = True
                Debug.WriteLine("WebView2 initialization completed successfully")

                ' 加载初始页面
                'NavigateToUrl("https://www.officeso.cn")
                'NavigateToUrl("https://piaofang.maoyan.com/dashboard")

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

    '' 新增：移除事件处理程序方法
    'Private Sub RemoveEventHandlers()
    '    Try
    '        If ChatBrowser?.CoreWebView2 IsNot Nothing Then
    '            RemoveHandler ChatBrowser.CoreWebView2.WebMessageReceived,
    '                AddressOf WebView2_MessageReceived
    '        End If

    '        RemoveHandler NavigateButton.Click, AddressOf NavigateButton_Click
    '        RemoveHandler CaptureButton.Click, AddressOf CaptureButton_Click
    '        RemoveHandler UrlTextBox.KeyPress, AddressOf UrlTextBox_KeyPress
    '        RemoveHandler SelectDomButton.Click, AddressOf SelectDomButton_Click
    '    Catch ex As Exception
    '        Debug.WriteLine($"Error removing event handlers: {ex.Message}")
    '    End Try
    'End Sub


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
            isCapturing = True

            ' 获取HTML内容
            Dim script As String
            If Not String.IsNullOrEmpty(selectedDomPath) Then
                ' 使用选定的DOM路径
                script = $"
                (function() {{
                    const element = document.querySelector('{selectedDomPath}');
                    return element ? element.outerHTML : null;
                }})();
            "
            Else
                ' 获取整个页面内容
                script = "document.documentElement.outerHTML;"
            End If

            Dim html = Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            If Not String.IsNullOrEmpty(html) Then
                html = JsonConvert.DeserializeObject(Of String)(html)
                HandleExtractedContent(html)
            Else
                MessageBox.Show("未能获取到内容", "提示")
            End If

        Catch ex As Exception
            MessageBox.Show($"抓取内容时出错: {ex.Message}", "错误")
        Finally
            isCapturing = False
        End Try
    End Sub



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
                document.removeEventListener('mouseover', window._domSelector.onMouseOver);
                document.removeEventListener('mouseout', window._domSelector.onMouseOut);
                document.removeEventListener('click', window._domSelector.onClick);
                if(window._domSelector.tip) window._domSelector.tip.remove();
            }

            // 创建新的选择器
            window._domSelector = {
                lastHighlight: null,
                lastParentHighlight: null,
                isShiftKey: false,
                _lastChildElement: null,
                _currentTarget: null,  // 新增：记录当前目标元素
                _highlightTimer: null, // 新增：用于防抖动的定时器
                
                // 创建提示框
                createTip: function() {
                    const tip = document.createElement('div');
                    tip.style.cssText = `
                        position: fixed;
                        top: 10px;
                        left: 50%;
                        transform: translateX(-50%);
                        background: rgba(0, 0, 0, 0.8);
                        color: white;
                        padding: 8px 16px;
                        border-radius: 4px;
                        font-family: Arial;
                        font-size: 14px;
                        z-index: 2147483647;
                        pointer-events: none;
                    `;
                    tip.innerHTML = '按住 Shift 键可选择父元素<br>点击选择要抓取的内容';
                    document.body.appendChild(tip);
                    this.tip = tip;
                },

                // 高亮显示
                highlight: function(element, isParent) {
        if (!element) return;
        
        // 清除之前的高亮
        this.removeHighlight();

        const target = isParent ? this.findParentElement(element) : element;
        if (!target) return;

        this._currentTarget = target;
        
        // 设置高亮样式
        target.style.transition = 'outline 0.2s ease-in-out';
        target.style.outline = isParent ? '3px dashed #FF9800' : '3px solid #2196F3';
        target.style.outlineOffset = '2px';
        
        // 保存当前高亮的元素
        this.lastHighlight = target;
        
        // 显示信息框
        this.showInfo(target, target.getBoundingClientRect(), isParent);

        // 如果是按住Shift键，记住子元素
        if (isParent) {
            this._lastChildElement = element;
        }
    },

                // 查找合适的父元素
                findParentElement: function(element) {
                    let parent = element;
                    while (parent && parent !== document.body) {
                        // 如果是表格相关元素，优先选择整个表格
                        if (parent.tagName === 'TD' || parent.tagName === 'TH') {
                            parent = this.findClosest(parent, 'table');
                            if (parent) break;
                        }
                        // 对于其他元素，查找有意义的父容器
                        if (this.isSignificantElement(parent)) {
                            break;
                        }
                        parent = parent.parentElement;
                    }
                    return parent;
                },

                // 判断是否是有意义的元素
                isSignificantElement: function(element) {
                    const tag = element.tagName.toLowerCase();
                    const significantTags = ['table', 'article', 'section', 'div', 'form', 'main'];
                    
                    if (significantTags.includes(tag)) {
                        // 检查是否包含足够的内容
                        if (element.textContent.trim().length > 50) return true;
                        // 检查是否有特定的类名或ID
                        if (element.id || element.className) return true;
                        // 检查是否包含多个子元素
                        if (element.children.length > 2) return true;
                    }
                    return false;
                },

                // 查找最近的指定标签祖先元素
                findClosest: function(element, tagName) {
                    while (element && element !== document.body) {
                        if (element.tagName.toLowerCase() === tagName.toLowerCase()) {
                            return element;
                        }
                        element = element.parentElement;
                    }
                    return null;
                },

                // 移除高亮
                removeHighlight: function() {
                    if (this.lastHighlight) {
                        this.lastHighlight.style.outline = '';
                        this.lastHighlight.style.outlineOffset = '';
                        this.lastHighlight = null;
                    }
                    if (this.lastParentHighlight) {
                        this.lastParentHighlight.style.outline = '';
                        this.lastParentHighlight.style.outlineOffset = '';
                        this.lastParentHighlight = null;
                    }
                    if (this.infoBox) {
                        this.infoBox.remove();
                        this.infoBox = null;
                    }
                },

                // 显示元素信息
                showInfo: function(element, rect, isParent) {
                    if (this.infoBox) this.infoBox.remove();
    
                    const info = document.createElement('div');
                    info.style.cssText = `
                        position: absolute;
                        background: ${isParent ? 'rgba(255, 152, 0, 0.9)' : 'rgba(33, 150, 243, 0.9)'};
                        color: white;
                        padding: 8px 16px;
                        border-radius: 4px;
                        font-size: 14px;
                        pointer-events: none;
                        z-index: 2147483647;
                        font-family: Arial;
                        max-width: 400px;
                        text-align: center;
                        transform: translateX(-50%);
                        box-shadow: 0 2px 8px rgba(0,0,0,0.2);
                    `;
    
                    const tag = element.tagName.toLowerCase();
                    const id = element.id ? '#' + element.id : '';
                    const classes = Array.from(element.classList).map(c => '.' + c).join('');
                    const contentPreview = element.textContent.trim().substring(0, 50) + 
                        (element.textContent.trim().length > 50 ? '...' : '');
    
                    info.innerHTML = `
                        <div style=""font-size: 16px; margin-bottom: 4px;"">
                            <${tag}${id}${classes}> ${isParent ? '(获取父元素)' : ''}
                        </div>
                        <div style=""font-size: 13px; opacity: 0.9;"">
                            ${contentPreview}
                        </div>
                    `;
    
                    // 计算位置，显示在元素正上方中央
                    const infoWidth = 400; // 固定宽度
                    const verticalOffset = 10; // 与元素的垂直距离
    
                    // 确保信息框在可视区域内
                    let left = rect.left + (rect.width / 2);
                    left = Math.min(Math.max(infoWidth / 2, left), document.documentElement.clientWidth - infoWidth / 2);
    
                    let top = rect.top + window.scrollY - verticalOffset;
                    top = Math.max(10, top); // 确保不会超出顶部
    
                    info.style.left = left + 'px';
                    info.style.top = top - info.offsetHeight + 'px';
    
                    document.body.appendChild(info);
                    this.infoBox = info;
                },

                // 事件处理程序
                onMouseOver: function(e) {
        if (window._domSelector) {
            e.stopPropagation();
            const target = e.target;
            
            // 如果目标元素相同则不重复处理
            if (target === window._domSelector._currentTarget) return;
            
            window._domSelector.highlight(target, window._domSelector.isShiftKey);
        }
    },

                // 修改鼠标移出事件处理
                onMouseOut: function(e) {
                    // 清除定时器
                    if (window._domSelector._highlightTimer) {
                        clearTimeout(window._domSelector._highlightTimer);
                    }

                    // 检查是否真的需要移除高亮
                    const relatedTarget = e.relatedTarget;
                    if (!window._domSelector.lastHighlight || 
                        !window._domSelector.lastHighlight.contains(relatedTarget)) {
                        window._domSelector.removeHighlight();
                        window._domSelector._currentTarget = null;
                    }

                    e.stopPropagation();
                },


                onClick: function(e) {
                    e.preventDefault();
                    e.stopPropagation();
                    
                    const element = window._domSelector.isShiftKey ? 
                        window._domSelector.findParentElement(e.target) : e.target;
                    
                    if (!element) return;
                    
                    const html = element.outerHTML;
                    const path = window._domSelector.getPath(element);
                    
                    window.chrome.webview.postMessage({
                        type: 'elementSelected',
                        html: html,
                        path: path,
                        tag: element.tagName.toLowerCase(),
                        text: element.innerText,
                        rect: element.getBoundingClientRect(),
                        isTable: element.tagName.toLowerCase() === 'table'
                    });
                },

                
                onKeyDown: function(e) {
        if (e.key === 'Shift' && !window._domSelector.isShiftKey) {
            window._domSelector.isShiftKey = true;
            
            // 如果当前有高亮的元素，切换到其父元素
            if (window._domSelector._currentTarget) {
                const parentElement = window._domSelector.findParentElement(window._domSelector._currentTarget);
                if (parentElement) {
                    window._domSelector.highlight(window._domSelector._currentTarget, true);
                }
            }
        }
    },

                // 修改键盘释放事件
    onKeyUp: function(e) {
        if (e.key === 'Shift') {
            window._domSelector.isShiftKey = false;
            
            // 恢复到子元素
            if (window._domSelector._lastChildElement) {
                window._domSelector.highlight(window._domSelector._lastChildElement, false);
            }
        }
    },

                // 获取元素路径
                getPath: function(element) {
                    const path = [];
                    while(element && element.nodeType === Node.ELEMENT_NODE) {
                        let selector = element.tagName.toLowerCase();
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
                },

                // 初始化
                // 修改初始化方法
    init: function() {
        // 确保清理之前的实例
        if (window._domSelector) {
            window._domSelector.cleanup();
        }
        
        this.createTip();
        
        // 使用 bind 确保事件处理程序中的 this 指向正确
        this.onMouseOver = this.onMouseOver.bind(this);
        this.onMouseOut = this.onMouseOut.bind(this);
        this.onClick = this.onClick.bind(this);
        this.onKeyDown = this.onKeyDown.bind(this);
        this.onKeyUp = this.onKeyUp.bind(this);
        
        // 添加事件监听器
        document.addEventListener('mouseover', this.onMouseOver, true);
        document.addEventListener('mouseout', this.onMouseOut, true);
        document.addEventListener('click', this.onClick, true);
        document.addEventListener('keydown', this.onKeyDown, true);
        document.addEventListener('keyup', this.onKeyUp, true);
        
        document.body.style.cursor = 'pointer';
    },

                // 修改清理方法
    cleanup: function() {
        this.removeHighlight();
        if (this.tip) this.tip.remove();
        
        // 移除所有事件监听器
        document.removeEventListener('mouseover', this.onMouseOver, true);
        document.removeEventListener('mouseout', this.onMouseOut, true);
        document.removeEventListener('click', this.onClick, true);
        document.removeEventListener('keydown', this.onKeyDown, true);
        document.removeEventListener('keyup', this.onKeyUp, true);
        
        // 清理所有状态
        this._currentTarget = null;
        this._lastChildElement = null;
        this.lastHighlight = null;
        this.isShiftKey = false;
        
        // 恢复鼠标样式
        document.body.style.cursor = '';
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

                Await ChatBrowser.CoreWebView2.ExecuteScriptAsync("
                            if(window._domSelector) {
                                window._domSelector.cleanup();
                                window._domSelector = null;
                            }
                        ")
                ' 显示自定义确认对话框
                Using dialog As New WebSiteContentConfirmDialog(text, tag, selectedDomPath)
                    Dim result = dialog.ShowDialog()
                    Select Case result
                        Case DialogResult.Cancel
                            ' 取消操作，清除路径
                            selectedDomPath = ""

                        Case DialogResult.Yes
                            ' 直接使用内容
                            HandleExtractedContent(text)

                        Case DialogResult.No
                            ' 调用AI聊天
                            ' 在子类 WebDataCapturePane 中实现这个方法
                            OnAiChatRequested(text)
                    End Select
                End Using
            End If
        Catch ex As Exception
            Debug.WriteLine($"处理消息错误: {ex.Message}")
            MessageBox.Show($"处理选择消息失败: {ex.Message}", "错误")
        End Try
    End Sub

    ' 添加事件以供子类处理AI聊天请求
    Protected Event AiChatRequested As EventHandler(Of String)
    Protected Sub OnAiChatRequested(content As String)
        RaiseEvent AiChatRequested(Me, content)
    End Sub
End Class