Imports System.Diagnostics
Imports System.Threading
Imports System.Threading.Tasks

''' <summary>
''' 阶段化启动管理器, 统一调度 VSTO 插件的四阶段启动流程
''' Phase 0: Critical  - 同步, 主线程, &lt;50ms, 仅注册事件处理器
''' Phase 1: Required - 后台并行, 程序集预加载 + 配置加载
''' Phase 2: Background - 空闲预加载, WebView2Loader/SqliteNativeLoader/ResourceExtractor
''' Phase 3: OnDemand - 首次使用时, 保持现有 Lazy 模式
''' </summary>
Public Class PhaseStartupManager

    ''' <summary>
    ''' 全局单例 — 供 ThisAddIn 和 BaseOfficeRibbon 共享
    ''' </summary>
    Public Shared ReadOnly Instance As New PhaseStartupManager()

    Private _timings As New List(Of (String, Long))
    Private _backgroundPhaseComplete As Boolean = False
    Private _requiredPhaseStarted As Boolean = False

    ''' <summary>
    ''' Phase 2 是否已完成（供 EnsureCoreServicesLoaded 判断是否已预热）
    ''' </summary>
    Public ReadOnly Property IsBackgroundReady As Boolean
        Get
            Return _backgroundPhaseComplete
        End Get
    End Property

    ''' <summary>
    ''' Phase 0: 关键路径初始化（同步，主线程，必须在 VSTO Startup 事件中调用）
    ''' 仅注册事件处理器，不做任何 I/O 密集操作
    ''' </summary>
    Public Sub RunCriticalPhase(application As Object)
        Dim sw = Stopwatch.StartNew()

        ' 仅注册 AssemblyResolve 事件，不做预加载
        Measure("AssemblyResolve.Register", Sub() SqliteAssemblyResolver.EnsureRegistered())

        ' 初始化全局状态栏
        Try
            Measure("GlobalStatusStrip.Init", Sub() GlobalStatusStripAll.InitializeApplication(application))
        Catch ex As Exception
            Debug.WriteLine($"[Startup] Phase0 GlobalStatusStrip failed: {ex.Message}")
        End Try

        sw.Stop()
        Debug.WriteLine($"[Startup] Phase0-Critical 完成: {sw.ElapsedMilliseconds}ms")
        LogTimings("Phase0")
    End Sub

    ''' <summary>
    ''' Phase 1: 必需组件初始化（后台并行，在 Ribbon Load 后调用）
    ''' 程序集预加载与配置加载并行执行
    ''' </summary>
    Public Sub StartRequiredPhase()
        If _requiredPhaseStarted Then Return
        _requiredPhaseStarted = True

        Task.Run(Sub()
                     Dim sw = Stopwatch.StartNew()

                     ' 并行：预加载程序集
                     Dim preloadTask = Task.Run(Sub()
                                                    Try
                                                        Measure("PreloadAssemblies", Sub() SqliteAssemblyResolver.PreloadAssemblies())
                                                    Catch ex As Exception
                                                        Debug.WriteLine($"[Startup] Phase1 PreloadAssemblies failed: {ex.Message}")
                                                    End Try
                                                End Sub)

                     Task.WaitAll(preloadTask)
                     sw.Stop()
                     Debug.WriteLine($"[Startup] Phase1-Required 完成: {sw.ElapsedMilliseconds}ms")
                     LogTimings("Phase1")
                 End Sub)
    End Sub

    ''' <summary>
    ''' Phase 2: 后台预加载（低优先级，fire-and-forget，在 Ribbon1_Load 末尾调用）
    ''' 预热 WebView2Loader、SqliteNativeLoader、ResourceExtractor
    ''' </summary>
    Public Sub StartBackgroundPhase()
        Task.Run(Sub()
                     Dim sw = Stopwatch.StartNew()
                     Debug.WriteLine("[Startup] Phase2-Background 开始...")

                     ' WebView2Loader 和 SqliteNativeLoader 无依赖关系，可并行
                     Dim wv2Task = Task.Run(Sub()
                                                Try
                                                    Measure("WebView2Loader", Sub() WebView2Loader.EnsureWebView2Loader())
                                                Catch ex As Exception
                                                    Debug.WriteLine($"[Startup] Phase2 WebView2Loader failed: {ex.Message}")
                                                End Try
                                            End Sub)

                     Dim sqliteTask = Task.Run(Sub()
                                                   Try
                                                       Measure("SqliteNativeLoader", Sub() SqliteNativeLoader.EnsureLoaded())
                                                   Catch ex As Exception
                                                       Debug.WriteLine($"[Startup] Phase2 SqliteNativeLoader failed: {ex.Message}")
                                                   End Try
                                               End Sub)

                     ' ResourceExtractor 独立，也可并行
                     Dim resTask = Task.Run(Sub()
                                                Try
                                                    Measure("ResourceExtractor", Sub() ResourceExtractor.ExtractResources())
                                                Catch ex As Exception
                                                    Debug.WriteLine($"[Startup] Phase2 ResourceExtractor failed: {ex.Message}")
                                                End Try
                                            End Sub)

                     Task.WaitAll(wv2Task, sqliteTask, resTask)
                     sw.Stop()
                     _backgroundPhaseComplete = True
                     Debug.WriteLine($"[Startup] Phase2-Background 完成: {sw.ElapsedMilliseconds}ms")
                     LogTimings("Phase2")
                 End Sub)
    End Sub

    Private Sub Measure(name As String, action As Action)
        Dim sw = Stopwatch.StartNew()
        action()
        sw.Stop()
        _timings.Add((name, sw.ElapsedMilliseconds))
    End Sub

    Private Sub LogTimings(phase As String)
        For Each t In _timings
            Debug.WriteLine($"[Startup] {phase}: {t.Item1} = {t.Item2}ms")
        Next
    End Sub
End Class
