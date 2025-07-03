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