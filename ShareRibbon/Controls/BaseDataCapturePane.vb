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
    ' ��ӳ�Ա����
    Private isNavigating As Boolean = False
    Private navigationTimer As Timer
    Private Const NAVIGATION_TIMEOUT As Integer = 10000 ' 10�볬ʱ

    Private domSelectionMode As Boolean = False
    Private selectedDomPath As String = ""

    Private isInitialized As Boolean = False
    Private isWebViewInitialized As Boolean = False
    Private pendingUrl As String = Nothing

    Private isCapturing As Boolean = False

    ' �ڹ��캯�����ʼ�������г�ʼ����ʱ��
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

            ' ��ʼ��������ʱ��
            InitializeNavigationTimer()

            ' �Զ����û�����Ŀ¼
            Dim userDataFolder As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MyAppWebView2Cache")

            If Not Directory.Exists(userDataFolder) Then
                Directory.CreateDirectory(userDataFolder)
            End If

            ' ���� WebView2 ����
            Dim env = Await CoreWebView2Environment.CreateAsync(
                Nothing, userDataFolder, New CoreWebView2EnvironmentOptions())

            ' ��ʼ�� WebView2
            Await ChatBrowser.EnsureCoreWebView2Async(env)

            ' ȷ�� CoreWebView2 �ѳ�ʼ��
            If ChatBrowser.CoreWebView2 IsNot Nothing Then
                Debug.WriteLine("CoreWebView2 initialized successfully")

                ' ���� WebView2 �İ�ȫѡ��
                With ChatBrowser.CoreWebView2.Settings
                    .IsScriptEnabled = True
                    .AreDefaultScriptDialogsEnabled = True
                    .IsWebMessageEnabled = True
                    .AreDevToolsEnabled = True
                End With

                ' �Ƴ����е��¼������������У�
                RemoveEventHandlers()

                ' ����µ��¼��������
                AddEventHandlers()

                isWebViewInitialized = True
                Debug.WriteLine("WebView2 initialization completed successfully")

                ' ���س�ʼҳ��
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


    ' �������¼�������������
    Private Sub AddEventHandlers()
        Debug.WriteLine("Adding event handlers...")

        ' �����¼�
        AddHandler ChatBrowser.CoreWebView2.NavigationStarting,
            Sub(s, args)
                Debug.WriteLine($"Navigation starting to: {args.Uri}")
                UrlTextBox.Text = args.Uri
                ' ���õ�����ť��������ʱ��ʱ��
                SetNavigationState(True)
            End Sub

        AddHandler ChatBrowser.CoreWebView2.NavigationCompleted,
            Sub(s, args)
                Debug.WriteLine($"Navigation completed: {args.IsSuccess}")
                ' ֹͣ��ʱ�����ָ���ť״̬
                SetNavigationState(False)
                If Not args.IsSuccess Then
                    ' ��ȡ����ϸ�Ĵ�����Ϣ
                    Dim errorStatus = ChatBrowser.CoreWebView2.GetDevToolsProtocolEventReceiver("Network.loadingFailed")
                    Debug.WriteLine($"Navigation failed with status: {errorStatus}")
                    MessageBox.Show("ҳ�����ʧ�ܣ������������ӻ�����", "����",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Else
                    Debug.WriteLine("ҳ����سɹ�")
                    ' ������������ӳɹ����صĴ����߼�
                End If
            End Sub

        ' �����´��ڴ������ض��򵽵�ǰ����
        AddHandler ChatBrowser.CoreWebView2.NewWindowRequested,
            Sub(s, args)
                ' ȡ���´��ڴ�
                args.Handled = True
                ' �ڵ�ǰ���ڵ�����Ŀ��URL
                ChatBrowser.CoreWebView2.Navigate(args.Uri)
                Debug.WriteLine($"���ص��´����������ض��򵽵�ǰ����: {args.Uri}")
            End Sub

        ' WebMessage�¼�
        AddHandler ChatBrowser.CoreWebView2.WebMessageReceived,
            AddressOf WebView2_MessageReceived

        ' ��ť�¼�
        AddHandler NavigateButton.Click, AddressOf NavigateButton_Click
        AddHandler CaptureButton.Click, AddressOf CaptureButton_Click
        AddHandler UrlTextBox.KeyPress, AddressOf UrlTextBox_KeyPress
        AddHandler SelectDomButton.Click, AddressOf SelectDomButton_Click

        ' ���ǰ�����˰�ť�¼�
        AddHandler BackButton.Click, AddressOf BackButton_Click
        AddHandler ForwardButton.Click, AddressOf ForwardButton_Click

        ' ������ʷ��¼״̬�仯
        AddHandler ChatBrowser.CoreWebView2.HistoryChanged,
        Sub(s, args)
            UpdateNavigationButtons()
        End Sub

        Debug.WriteLine("Event handlers added successfully")
    End Sub

    ' �Ƴ��¼��������ʱҲ��Ҫ�Ƴ��������¼�
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

    ' ���ǰ�����˰�ť����¼�����
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

    ' ����ǰ�����˰�ť״̬
    Private Sub UpdateNavigationButtons()
        If ChatBrowser?.CoreWebView2 IsNot Nothing Then
            BackButton.Enabled = ChatBrowser.CoreWebView2.CanGoBack
            ForwardButton.Enabled = ChatBrowser.CoreWebView2.CanGoForward
        Else
            BackButton.Enabled = False
            ForwardButton.Enabled = False
        End If
    End Sub

    ' ��ӵ���״̬���Ʒ���
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

    ' ��ӳ�ʱ������
    Private Sub OnNavigationTimeout(sender As Object, e As EventArgs)
        ' ��UI�߳���ִ��
        If Me.InvokeRequired Then
            Me.Invoke(Sub() OnNavigationTimeout(sender, e))
            Return
        End If

        navigationTimer.Stop()
        If isNavigating Then
            ' ������ڵ���״̬����ǿ�ƻָ���ť
            SetNavigationState(False)
            Debug.WriteLine("Navigation timeout - restoring button state")
            'MessageBox.Show("ҳ����س�ʱ���ѻָ�������ť", "��ʾ",
            '              MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    '' �������Ƴ��¼�������򷽷�
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

    ' �޸ģ�����������Ӹ��������Ϣ
    Private Sub NavigateToUrl(url As String)
        If String.IsNullOrWhiteSpace(url) Then
            Debug.WriteLine("Navigation cancelled: Empty URL")
            Return
        End If

        ' ������ڵ����У������µĵ�������
        If isNavigating Then
            Debug.WriteLine("Navigation in progress, ignoring new request")
            Return
        End If

        Try
            ' ��׼��URL
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
                MessageBox.Show("WebView2 ���δ����", "����",
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            Debug.WriteLine($"Navigation error: {ex.Message}")
            MessageBox.Show($"����ʧ��: {ex.Message}", "����",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' ȷ���ڷ�������ʱ�ָ���ť״̬
            SetNavigationState(False)
        End Try
    End Sub

    Private Async Sub CaptureButton_Click(sender As Object, e As EventArgs)
        If isCapturing Then
            MessageBox.Show("����ץȡ�У����Ժ�...", "��ʾ")
            Return
        End If

        Try
            isCapturing = True

            ' ��ȡHTML����
            Dim script As String
            If Not String.IsNullOrEmpty(selectedDomPath) Then
                ' ʹ��ѡ����DOM·��
                script = $"
                (function() {{
                    const element = document.querySelector('{selectedDomPath}');
                    return element ? element.outerHTML : null;
                }})();
            "
            Else
                ' ��ȡ����ҳ������
                script = "document.documentElement.outerHTML;"
            End If

            Dim html = Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            If Not String.IsNullOrEmpty(html) Then
                html = JsonConvert.DeserializeObject(Of String)(html)
                HandleExtractedContent(html)
            Else
                MessageBox.Show("δ�ܻ�ȡ������", "��ʾ")
            End If

        Catch ex As Exception
            MessageBox.Show($"ץȡ����ʱ����: {ex.Message}", "����")
        Finally
            isCapturing = False
        End Try
    End Sub



    ' ��ӱ������ģ��
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


    ' ��ӳ��󷽷�
    Protected MustOverride Function CreateTable(tableData As TableData) As String

    ' ���󷽷���������ȡ�����ݣ��ɾ���ʵ����ʵ�֣�
    Protected MustOverride Sub HandleExtractedContent(content As String)


    ' ѡ��DOMԪ�ذ�ť����¼��������
    Private Async Sub SelectDomButton_Click(sender As Object, e As EventArgs)
        Try
            Dim selectScript As String = "
        (function() {
            // �Ƴ��ɵ�ѡ����
            if(window._domSelector) {
                document.removeEventListener('mouseover', window._domSelector.onMouseOver);
                document.removeEventListener('mouseout', window._domSelector.onMouseOut);
                document.removeEventListener('click', window._domSelector.onClick);
                if(window._domSelector.tip) window._domSelector.tip.remove();
            }

            // �����µ�ѡ����
            window._domSelector = {
                lastHighlight: null,
                lastParentHighlight: null,
                isShiftKey: false,
                _lastChildElement: null,
                _currentTarget: null,  // ��������¼��ǰĿ��Ԫ��
                _highlightTimer: null, // ���������ڷ������Ķ�ʱ��
                
                // ������ʾ��
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
                    tip.innerHTML = '��ס Shift ����ѡ��Ԫ��<br>���ѡ��Ҫץȡ������';
                    document.body.appendChild(tip);
                    this.tip = tip;
                },

                // ������ʾ
                highlight: function(element, isParent) {
        if (!element) return;
        
        // ���֮ǰ�ĸ���
        this.removeHighlight();

        const target = isParent ? this.findParentElement(element) : element;
        if (!target) return;

        this._currentTarget = target;
        
        // ���ø�����ʽ
        target.style.transition = 'outline 0.2s ease-in-out';
        target.style.outline = isParent ? '3px dashed #FF9800' : '3px solid #2196F3';
        target.style.outlineOffset = '2px';
        
        // ���浱ǰ������Ԫ��
        this.lastHighlight = target;
        
        // ��ʾ��Ϣ��
        this.showInfo(target, target.getBoundingClientRect(), isParent);

        // ����ǰ�סShift������ס��Ԫ��
        if (isParent) {
            this._lastChildElement = element;
        }
    },

                // ���Һ��ʵĸ�Ԫ��
                findParentElement: function(element) {
                    let parent = element;
                    while (parent && parent !== document.body) {
                        // ����Ǳ�����Ԫ�أ�����ѡ���������
                        if (parent.tagName === 'TD' || parent.tagName === 'TH') {
                            parent = this.findClosest(parent, 'table');
                            if (parent) break;
                        }
                        // ��������Ԫ�أ�����������ĸ�����
                        if (this.isSignificantElement(parent)) {
                            break;
                        }
                        parent = parent.parentElement;
                    }
                    return parent;
                },

                // �ж��Ƿ����������Ԫ��
                isSignificantElement: function(element) {
                    const tag = element.tagName.toLowerCase();
                    const significantTags = ['table', 'article', 'section', 'div', 'form', 'main'];
                    
                    if (significantTags.includes(tag)) {
                        // ����Ƿ�����㹻������
                        if (element.textContent.trim().length > 50) return true;
                        // ����Ƿ����ض���������ID
                        if (element.id || element.className) return true;
                        // ����Ƿ���������Ԫ��
                        if (element.children.length > 2) return true;
                    }
                    return false;
                },

                // ���������ָ����ǩ����Ԫ��
                findClosest: function(element, tagName) {
                    while (element && element !== document.body) {
                        if (element.tagName.toLowerCase() === tagName.toLowerCase()) {
                            return element;
                        }
                        element = element.parentElement;
                    }
                    return null;
                },

                // �Ƴ�����
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

                // ��ʾԪ����Ϣ
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
                            <${tag}${id}${classes}> ${isParent ? '(��ȡ��Ԫ��)' : ''}
                        </div>
                        <div style=""font-size: 13px; opacity: 0.9;"">
                            ${contentPreview}
                        </div>
                    `;
    
                    // ����λ�ã���ʾ��Ԫ�����Ϸ�����
                    const infoWidth = 400; // �̶����
                    const verticalOffset = 10; // ��Ԫ�صĴ�ֱ����
    
                    // ȷ����Ϣ���ڿ���������
                    let left = rect.left + (rect.width / 2);
                    left = Math.min(Math.max(infoWidth / 2, left), document.documentElement.clientWidth - infoWidth / 2);
    
                    let top = rect.top + window.scrollY - verticalOffset;
                    top = Math.max(10, top); // ȷ�����ᳬ������
    
                    info.style.left = left + 'px';
                    info.style.top = top - info.offsetHeight + 'px';
    
                    document.body.appendChild(info);
                    this.infoBox = info;
                },

                // �¼��������
                onMouseOver: function(e) {
        if (window._domSelector) {
            e.stopPropagation();
            const target = e.target;
            
            // ���Ŀ��Ԫ����ͬ���ظ�����
            if (target === window._domSelector._currentTarget) return;
            
            window._domSelector.highlight(target, window._domSelector.isShiftKey);
        }
    },

                // �޸�����Ƴ��¼�����
                onMouseOut: function(e) {
                    // �����ʱ��
                    if (window._domSelector._highlightTimer) {
                        clearTimeout(window._domSelector._highlightTimer);
                    }

                    // ����Ƿ������Ҫ�Ƴ�����
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
            
            // �����ǰ�и�����Ԫ�أ��л����丸Ԫ��
            if (window._domSelector._currentTarget) {
                const parentElement = window._domSelector.findParentElement(window._domSelector._currentTarget);
                if (parentElement) {
                    window._domSelector.highlight(window._domSelector._currentTarget, true);
                }
            }
        }
    },

                // �޸ļ����ͷ��¼�
    onKeyUp: function(e) {
        if (e.key === 'Shift') {
            window._domSelector.isShiftKey = false;
            
            // �ָ�����Ԫ��
            if (window._domSelector._lastChildElement) {
                window._domSelector.highlight(window._domSelector._lastChildElement, false);
            }
        }
    },

                // ��ȡԪ��·��
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

                // ��ʼ��
                // �޸ĳ�ʼ������
    init: function() {
        // ȷ������֮ǰ��ʵ��
        if (window._domSelector) {
            window._domSelector.cleanup();
        }
        
        this.createTip();
        
        // ʹ�� bind ȷ���¼���������е� this ָ����ȷ
        this.onMouseOver = this.onMouseOver.bind(this);
        this.onMouseOut = this.onMouseOut.bind(this);
        this.onClick = this.onClick.bind(this);
        this.onKeyDown = this.onKeyDown.bind(this);
        this.onKeyUp = this.onKeyUp.bind(this);
        
        // ����¼�������
        document.addEventListener('mouseover', this.onMouseOver, true);
        document.addEventListener('mouseout', this.onMouseOut, true);
        document.addEventListener('click', this.onClick, true);
        document.addEventListener('keydown', this.onKeyDown, true);
        document.addEventListener('keyup', this.onKeyUp, true);
        
        document.body.style.cursor = 'pointer';
    },

                // �޸�������
    cleanup: function() {
        this.removeHighlight();
        if (this.tip) this.tip.remove();
        
        // �Ƴ������¼�������
        document.removeEventListener('mouseover', this.onMouseOver, true);
        document.removeEventListener('mouseout', this.onMouseOut, true);
        document.removeEventListener('click', this.onClick, true);
        document.removeEventListener('keydown', this.onKeyDown, true);
        document.removeEventListener('keyup', this.onKeyUp, true);
        
        // ��������״̬
        this._currentTarget = null;
        this._lastChildElement = null;
        this.lastHighlight = null;
        this.isShiftKey = false;
        
        // �ָ������ʽ
        document.body.style.cursor = '';
    }
            };

            // ��ʼ��ѡ����
            window._domSelector.init();
        })();
        "
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(selectScript)

        Catch ex As Exception
            Debug.WriteLine($"DOMѡ��������: {ex.Message}")
            MessageBox.Show($"��ʼ��ѡ����ʧ��: {ex.Message}", "����")
        End Try
    End Sub

    ' �޸���Ϣ�������
    ' �޸���Ϣ�������
    Private Async Sub WebView2_MessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Debug.WriteLine($"�յ���Ϣ: {e.WebMessageAsJson}")

            Dim message = JsonConvert.DeserializeObject(Of JObject)(e.WebMessageAsJson)
            If message("type")?.ToString() = "elementSelected" Then
                ' ��ȡ������Ϣ
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
                ' ��ʾ�Զ���ȷ�϶Ի���
                Using dialog As New WebSiteContentConfirmDialog(text, tag, selectedDomPath)
                    Dim result = dialog.ShowDialog()
                    Select Case result
                        Case DialogResult.Cancel
                            ' ȡ�����������·��
                            selectedDomPath = ""

                        Case DialogResult.Yes
                            ' ֱ��ʹ������
                            HandleExtractedContent(text)

                        Case DialogResult.No
                            ' ����AI����
                            ' ������ WebDataCapturePane ��ʵ���������
                            OnAiChatRequested(text)
                    End Select
                End Using
            End If
        Catch ex As Exception
            Debug.WriteLine($"������Ϣ����: {ex.Message}")
            MessageBox.Show($"����ѡ����Ϣʧ��: {ex.Message}", "����")
        End Try
    End Sub

    ' ����¼��Թ����ദ��AI��������
    Protected Event AiChatRequested As EventHandler(Of String)
    Protected Sub OnAiChatRequested(content As String)
        RaiseEvent AiChatRequested(Me, content)
    End Sub
End Class