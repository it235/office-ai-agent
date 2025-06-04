Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security

' 有关程序集的一般信息由以下
' 控制。更改这些特性值可修改
' 与程序集关联的信息。

'查看程序集特性的值

<Assembly: AssemblyTitle("ExcelAi")> 
<Assembly: AssemblyDescription("")> 
<Assembly: AssemblyCompany("")> 
<Assembly: AssemblyProduct("ExcelAi")> 
<Assembly: AssemblyCopyright("Copyright ©  2025")> 
<Assembly: AssemblyTrademark("")> 

'将 ComVisible 设置为 false 将使此程序集中的类型
'对 COM 组件不可见。  如果需要从 COM 访问此程序集中的类型，
'请将此类型的 ComVisible 特性设置为 true。
<Assembly: ComVisible(False)>

'如果此项目向 COM 公开，则下列 GUID 用于类型库的 ID
<Assembly: Guid("9d7e14b7-cfa3-45df-90c3-43fc080611cf")> 

' 程序集的版本信息由下列四个值组成:
'
'      主版本
'      次版本
'      生成号
'      修订号
'
'可以指定所有这些值，也可以使用“生成号”和“修订号”的默认值，
' 方法是按如下所示使用“*”: :
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.0")> 
<Assembly: AssemblyFileVersion("1.0.0.0")> 

Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module
