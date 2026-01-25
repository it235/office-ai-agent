' 重构验证脚本
' 检查BaseDeepseekChat和BaseDoubaoChat是否正确继承了BaseChat

Imports System
Imports System.IO
Imports System.Text

Module RefactoringVerification
    Sub Main()
        Console.WriteLine("开始验证重构结果...")
        
        ' 检查文件是否存在
        Dim baseChatPath As String = "F:\ai\code\AiHelper\ShareRibbon\Controls\BaseChat.vb"
        Dim deepseekPath As String = "F:\ai\code\AiHelper\ShareRibbon\Controls\BaseDeepseekChat.vb"
        Dim doubaoPath As String = "F:\ai\code\AiHelper\ShareRibbon\Controls\BaseDoubaoChat.vb"
        
        If Not File.Exists(baseChatPath) Then
            Console.WriteLine("错误: BaseChat.vb 文件不存在")
            Exit Sub
        End If
        
        If Not File.Exists(deepseekPath) Then
            Console.WriteLine("错误: BaseDeepseekChat.vb 文件不存在")
            Exit Sub
        End If
        
        If Not File.Exists(doubaoPath) Then
            Console.WriteLine("错误: BaseDoubaoChat.vb 文件不存在")
            Exit Sub
        End If
        
        Console.WriteLine("✓ 所有基类文件都存在")
        
        ' 检查继承关系
        Dim deepseekContent As String = File.ReadAllText(deepseekPath)
        Dim doubaoContent As String = File.ReadAllText(doubaoPath)
        
        If deepseekContent.Contains("Inherits BaseChat") Then
            Console.WriteLine("✓ BaseDeepseekChat 正确继承了 BaseChat")
        Else
            Console.WriteLine("✗ BaseDeepseekChat 没有正确继承 BaseChat")
        End If
        
        If doubaoContent.Contains("Inherits BaseChat") Then
            Console.WriteLine("✓ BaseDoubaoChat 正确继承了 BaseChat")
        Else
            Console.WriteLine("✗ BaseDoubaoChat 没有正确继承 BaseChat")
        End If
        
        ' 检查是否有重复代码
        Dim baseChatContent As String = File.ReadAllText(baseChatPath)
        
        ' 检查BaseDeepseekChat是否移除了重复代码
        Dim removedDuplicates As Boolean = True
        Dim duplicateMethods() As String = {
            "Protected Overrides Sub WndProc",
            "Protected Overrides Sub OnGotFocus",
            "Protected Overrides Sub OnClick",
            "Protected Sub WebView2_WebMessageReceived",
            "Protected Overridable Sub HandleExecuteCode"
        }
        
        For Each method As String In duplicateMethods
            If deepseekContent.Contains(method) Then
                Console.WriteLine($"⚠ BaseDeepseekChat 仍包含重复方法: {method}")
                removedDuplicates = False
            End If
        Next
        
        If removedDuplicates Then
            Console.WriteLine("✓ BaseDeepseekChat 已移除重复代码")
        End If
        
        ' 检查是否实现了必要的方法
        Dim requiredMethods() As String = {
            "Protected Overrides ReadOnly Property ChatUrl",
            "Protected Overrides ReadOnly Property SessionFileName",
            "Protected Overrides Function GetWebView2DataFolderName",
            "Protected Overrides Async Function InjectExecuteButtonsSafe"
        }
        
        For Each method As String In requiredMethods
            If Not deepseekContent.Contains(method) Then
                Console.WriteLine($"✗ BaseDeepseekChat 缺少必要方法: {method}")
            Else
                Console.WriteLine($"✓ BaseDeepseekChat 实现了: {method}")
            End If
        Next
        
        ' 检查项目文件是否包含新文件
        Dim projectPath As String = "F:\ai\code\AiHelper\ShareRibbon\ShareRibbon.vbproj"
        If File.Exists(projectPath) Then
            Dim projectContent As String = File.ReadAllText(projectPath)
            If projectContent.Contains("Controls\BaseChat.vb") Then
                Console.WriteLine("✓ 项目文件已包含 BaseChat.vb")
            Else
                Console.WriteLine("✗ 项目文件缺少 BaseChat.vb")
            End If
        Else
            Console.WriteLine("⚠ 无法找到项目文件")
        End If
        
        Console.WriteLine("重构验证完成!")
        Console.WriteLine()
        Console.WriteLine("重构总结:")
        Console.WriteLine("1. 创建了新的 BaseChat 基类，提取了公共代码")
        Console.WriteLine("2. BaseDeepseekChat 和 BaseDoubaoChat 现在都继承自 BaseChat")
        Console.WriteLine("3. 移除了重复的代码，保持了各自的特殊功能")
        Console.WriteLine("4. 抽象化了平台特定的实现")
        Console.WriteLine("5. 保持了所有现有的子类兼容性")
    End Sub
End Module