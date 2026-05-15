' ShareRibbon\Protocol\InstructionRegistry.vb
' 指令注册表 - 维护所有支持的指令及其参数Schema

Imports System.Collections.Generic
Imports Newtonsoft.Json.Linq

''' <summary>
''' 指令注册表 - 维护所有支持的指令及其参数Schema
''' </summary>
Public Class InstructionRegistry

    Private Shared ReadOnly _registry As New Dictionary(Of String, InstructionDefinition)()
    Private Shared _isInitialized As Boolean = False

    ''' <summary>
    ''' 初始化注册表（首次访问时自动调用）
    ''' </summary>
    Public Shared Sub Initialize()
        If _isInitialized Then Return

        ' ========== 排版指令 ==========

        Register("setParagraphStyle", New InstructionDefinition With {
            .Operation = "setParagraphStyle",
            .Category = "reformat",
            .DisplayName = "设置段落样式",
            .Description = "设置段落的样式名称、字体、对齐方式、行距等",
            .RequiredParams = New List(Of String) From {"target"},
            .OptionalParams = New List(Of String) From {"params.styleName", "params.font", "params.alignment", "params.spacing", "params.indent", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"semantic", "index", "range", "selection"})},
                {"target.selector", ParamType.StringType()},
                {"target.index", ParamType.NumberType()},
                {"params.styleName", ParamType.StringType()},
                {"params.font.name", ParamType.StringType()},
                {"params.font.size", ParamType.NumberType()},
                {"params.font.bold", ParamType.BooleanType()},
                {"params.font.italic", ParamType.BooleanType()},
                {"params.font.color", ParamType.StringType()},
                {"params.alignment", ParamType.EnumType(New String() {"left", "center", "right", "justify"})},
                {"params.spacing.before", ParamType.NumberType()},
                {"params.spacing.after", ParamType.NumberType()},
                {"params.spacing.line", ParamType.NumberType()},
                {"params.indent.firstLine", ParamType.NumberType()},
                {"params.indent.left", ParamType.NumberType()},
                {"params.indent.right", ParamType.NumberType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"paragraph", "selection"}
        })

        Register("setCharacterFormat", New InstructionDefinition With {
            .Operation = "setCharacterFormat",
            .Category = "reformat",
            .DisplayName = "设置字符格式",
            .Description = "设置选定文本的字符格式（字体、颜色、粗体等）",
            .RequiredParams = New List(Of String) From {"target"},
            .OptionalParams = New List(Of String) From {"params.font", "params.color", "params.bold", "params.italic", "params.underline", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"semantic", "range", "selection"})},
                {"target.selector", ParamType.StringType()},
                {"params.font.name", ParamType.StringType()},
                {"params.font.size", ParamType.NumberType()},
                {"params.color", ParamType.StringType()},
                {"params.bold", ParamType.BooleanType()},
                {"params.italic", ParamType.BooleanType()},
                {"params.underline", ParamType.EnumType(New String() {"none", "single", "double", "wavy"})}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"range", "selection"}
        })

        Register("insertTable", New InstructionDefinition With {
            .Operation = "insertTable",
            .Category = "reformat",
            .DisplayName = "插入表格",
            .Description = "在指定位置插入表格",
            .RequiredParams = New List(Of String) From {"target", "params.rows", "params.cols"},
            .OptionalParams = New List(Of String) From {"params.style", "params.data", "params.headerRow", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"position", "after", "before"})},
                {"target.position", ParamType.EnumType(New String() {"cursor", "end", "start"})},
                {"params.rows", ParamType.NumberType()},
                {"params.cols", ParamType.NumberType()},
                {"params.style", ParamType.StringType()},
                {"params.headerRow", ParamType.BooleanType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"position"}
        })

        Register("formatTable", New InstructionDefinition With {
            .Operation = "formatTable",
            .Category = "reformat",
            .DisplayName = "格式化表格",
            .Description = "格式化已有表格的样式、边框等",
            .RequiredParams = New List(Of String) From {"target"},
            .OptionalParams = New List(Of String) From {"params.style", "params.borders", "params.headerRow", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"index", "semantic"})},
                {"target.index", ParamType.NumberType()},
                {"target.selector", ParamType.StringType()},
                {"params.style", ParamType.StringType()},
                {"params.borders", ParamType.BooleanType()},
                {"params.headerRow", ParamType.BooleanType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"table"}
        })

        Register("setPageSetup", New InstructionDefinition With {
            .Operation = "setPageSetup",
            .Category = "reformat",
            .DisplayName = "页面设置",
            .Description = "设置页面边距、方向、纸张大小",
            .RequiredParams = New List(Of String) From {},
            .OptionalParams = New List(Of String) From {"params.margins", "params.orientation", "params.paperSize", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"params.margins.top", ParamType.NumberType()},
                {"params.margins.bottom", ParamType.NumberType()},
                {"params.margins.left", ParamType.NumberType()},
                {"params.margins.right", ParamType.NumberType()},
                {"params.orientation", ParamType.EnumType(New String() {"portrait", "landscape"})},
                {"params.paperSize", ParamType.StringType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"document"}
        })

        Register("insertBreak", New InstructionDefinition With {
            .Operation = "insertBreak",
            .Category = "reformat",
            .DisplayName = "插入分隔符",
            .Description = "插入分页符、分节符或换行符",
            .RequiredParams = New List(Of String) From {"params.type"},
            .OptionalParams = New List(Of String) From {"target", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"position", "after"})},
                {"target.position", ParamType.EnumType(New String() {"cursor", "end"})},
                {"params.type", ParamType.EnumType(New String() {"page", "section", "line"})}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"position", "paragraph"}
        })

        Register("applyListFormat", New InstructionDefinition With {
            .Operation = "applyListFormat",
            .Category = "reformat",
            .DisplayName = "应用列表格式",
            .Description = "将段落设置为列表（项目符号或编号）",
            .RequiredParams = New List(Of String) From {"target"},
            .OptionalParams = New List(Of String) From {"params.listType", "params.numberFormat", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"semantic", "index", "range", "selection"})},
                {"params.listType", ParamType.EnumType(New String() {"bullet", "number", "outline"})},
                {"params.numberFormat", ParamType.StringType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"paragraph", "selection"}
        })

        Register("setColumnFormat", New InstructionDefinition With {
            .Operation = "setColumnFormat",
            .Category = "reformat",
            .DisplayName = "分栏设置",
            .Description = "设置文档分栏",
            .RequiredParams = New List(Of String) From {"params.columnCount"},
            .OptionalParams = New List(Of String) From {"params.columnWidth", "params.separator", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"params.columnCount", ParamType.NumberType()},
                {"params.columnWidth", ParamType.NumberType()},
                {"params.separator", ParamType.BooleanType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"section"}
        })

        Register("insertHeaderFooter", New InstructionDefinition With {
            .Operation = "insertHeaderFooter",
            .Category = "reformat",
            .DisplayName = "插入页眉页脚",
            .Description = "插入或修改页眉页脚内容",
            .RequiredParams = New List(Of String) From {"params.type", "params.content"},
            .OptionalParams = New List(Of String) From {"params.alignment", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"params.type", ParamType.EnumType(New String() {"header", "footer"})},
                {"params.content", ParamType.StringType()},
                {"params.alignment", ParamType.EnumType(New String() {"left", "center", "right"})}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"section"}
        })

        Register("generateToc", New InstructionDefinition With {
            .Operation = "generateToc",
            .Category = "reformat",
            .DisplayName = "生成目录",
            .Description = "自动生成文档目录",
            .RequiredParams = New List(Of String) From {},
            .OptionalParams = New List(Of String) From {"target.position", "params.levels", "params.includePageNumbers", "expected", "rollback"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.position", ParamType.EnumType(New String() {"start", "cursor"})},
                {"params.levels", ParamType.NumberType()},
                {"params.includePageNumbers", ParamType.BooleanType()}
            },
            .IsDestructive = True,
            .SupportedTargetTypes = New List(Of String) From {"document"}
        })

        ' ========== 校对指令 ==========

        Register("suggestCorrection", New InstructionDefinition With {
            .Operation = "suggestCorrection",
            .Category = "proofread",
            .DisplayName = "建议修正",
            .Description = "建议将原文修正为新的文本",
            .RequiredParams = New List(Of String) From {"target", "params.original", "params.suggestion"},
            .OptionalParams = New List(Of String) From {"params.issueType", "params.severity", "params.explanation", "expected"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"textMatch", "semantic", "range"})},
                {"target.match", ParamType.StringType()},
                {"params.original", ParamType.StringType()},
                {"params.suggestion", ParamType.StringType()},
                {"params.issueType", ParamType.EnumType(New String() {"spellingError", "wordUsageError", "punctuationError", "grammaticalError", "expressionError", "formatError"})},
                {"params.severity", ParamType.EnumType(New String() {"high", "medium", "low"})},
                {"params.explanation", ParamType.StringType()}
            },
            .IsDestructive = False,
            .RequiresConfirmation = True,
            .SupportedTargetTypes = New List(Of String) From {"text", "range"}
        })

        Register("suggestFormatFix", New InstructionDefinition With {
            .Operation = "suggestFormatFix",
            .Category = "proofread",
            .DisplayName = "建议格式修正",
            .Description = "建议修正格式问题",
            .RequiredParams = New List(Of String) From {"target"},
            .OptionalParams = New List(Of String) From {"params.currentFormat", "params.expectedFormat", "params.explanation", "expected"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"semantic", "range", "paragraph"})},
                {"params.currentFormat", ParamType.StringType()},
                {"params.expectedFormat", ParamType.StringType()},
                {"params.explanation", ParamType.StringType()}
            },
            .IsDestructive = False,
            .RequiresConfirmation = True,
            .SupportedTargetTypes = New List(Of String) From {"range", "paragraph"}
        })

        Register("suggestStyleUnify", New InstructionDefinition With {
            .Operation = "suggestStyleUnify",
            .Category = "proofread",
            .DisplayName = "建议样式统一",
            .Description = "建议将不一致的样式统一",
            .RequiredParams = New List(Of String) From {"params.targetStyle", "params.inconsistentRanges"},
            .OptionalParams = New List(Of String) From {"params.expectedStyle", "expected"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"params.targetStyle", ParamType.StringType()},
                {"params.expectedStyle", ParamType.StringType()},
                {"params.inconsistentRanges", ParamType.ArrayType()}
            },
            .IsDestructive = False,
            .RequiresConfirmation = True,
            .SupportedTargetTypes = New List(Of String) From {"document"}
        })

        Register("markForReview", New InstructionDefinition With {
            .Operation = "markForReview",
            .Category = "proofread",
            .DisplayName = "标记待审核",
            .Description = "标记某处内容供用户审核",
            .RequiredParams = New List(Of String) From {"target"},
            .OptionalParams = New List(Of String) From {"params.note", "params.category", "expected"},
            .ParamSchema = New Dictionary(Of String, ParamType) From {
                {"target.type", ParamType.EnumType(New String() {"semantic", "range", "paragraph"})},
                {"params.note", ParamType.StringType()},
                {"params.category", ParamType.StringType()}
            },
            .IsDestructive = False,
            .SupportedTargetTypes = New List(Of String) From {"range", "paragraph"}
        })

        _isInitialized = True
    End Sub

    ''' <summary>
    ''' 注册指令定义
    ''' </summary>
    Public Shared Sub Register(operation As String, definition As InstructionDefinition)
        _registry(operation.ToLower()) = definition
    End Sub

    ''' <summary>
    ''' 检查操作是否有效
    ''' </summary>
    Public Shared Function IsValidOperation(operation As String) As Boolean
        EnsureInitialized()
        Return _registry.ContainsKey(operation.ToLower())
    End Function

    ''' <summary>
    ''' 获取指令定义
    ''' </summary>
    Public Shared Function GetDefinition(operation As String) As InstructionDefinition
        EnsureInitialized()
        Dim key = operation.ToLower()
        If _registry.ContainsKey(key) Then
            Return _registry(key)
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' 校验指令参数
    ''' </summary>
    Public Shared Function ValidateParameters(operation As String, params As JToken) As ParamValidationResult
        EnsureInitialized()

        Dim key = operation.ToLower()
        If Not _registry.ContainsKey(key) Then
            Return ParamValidationResult.Failure($"未知操作类型: {operation}")
        End If

        Dim def = _registry(key)

        If params Is Nothing OrElse params.Type <> JTokenType.Object Then
            Return ParamValidationResult.Failure("params必须是对象类型")
        End If

        Dim paramsObj = CType(params, JObject)

        ' 检查必需参数
        For Each required In def.RequiredParams
            Dim token = paramsObj.SelectToken(required)
            If token Is Nothing OrElse token.Type = JTokenType.Null Then
                Return ParamValidationResult.Failure($"缺少必需参数: {required}")
            End If
        Next

        ' 校验参数类型
        For Each kvp In def.ParamSchema
            Dim token = paramsObj.SelectToken(kvp.Key)
            If token IsNot Nothing AndAlso token.Type <> JTokenType.Null Then
                If Not IsTokenTypeMatch(token, kvp.Value) Then
                    Return ParamValidationResult.Failure($"参数 {kvp.Key} 类型不匹配，期望 {kvp.Value.BaseType}")
                End If
            End If
        Next

        Return ParamValidationResult.Success()
    End Function

    ''' <summary>
    ''' 获取某类别的所有指令
    ''' </summary>
    Public Shared Function GetOperationsByCategory(category As String) As List(Of InstructionDefinition)
        EnsureInitialized()
        Dim result As New List(Of InstructionDefinition)()
        For Each def In _registry.Values
            If def.Category.ToLower() = category.ToLower() Then
                result.Add(def)
            End If
        Next
        Return result
    End Function

    ''' <summary>
    ''' 获取所有已注册的操作名称
    ''' </summary>
    Public Shared Function GetAllOperations() As List(Of String)
        EnsureInitialized()
        Return New List(Of String)(_registry.Keys)
    End Function

    ''' <summary>
    ''' 生成AI Prompt用的指令说明文本
    ''' </summary>
    Public Shared Function BuildPromptDocumentation(category As String) As String
        EnsureInitialized()
        Dim defs = GetOperationsByCategory(category)
        Dim sb As New System.Text.StringBuilder()

        sb.AppendLine($"【{If(category = "reformat", "排版", "校对")}可用指令】")
        sb.AppendLine()

        For Each def In defs
            sb.AppendLine($"- {def.Operation}: {def.DisplayName}")
            sb.AppendLine($"  描述: {def.Description}")
            If def.RequiredParams.Count > 0 Then
                sb.AppendLine($"  必需参数: {String.Join(", ", def.RequiredParams)}")
            End If
            sb.AppendLine()
        Next

        Return sb.ToString()
    End Function

    Private Shared Sub EnsureInitialized()
        If Not _isInitialized Then
            Initialize()
        End If
    End Sub

    Private Shared Function IsTokenTypeMatch(token As JToken, paramType As ParamType) As Boolean
        Select Case paramType.BaseType.ToLower()
            Case "string"
                Return token.Type = JTokenType.String OrElse token.Type = JTokenType.Integer OrElse token.Type = JTokenType.Float
            Case "number"
                Return token.Type = JTokenType.Integer OrElse token.Type = JTokenType.Float
            Case "boolean"
                Return token.Type = JTokenType.Boolean
            Case "array"
                Return token.Type = JTokenType.Array
            Case "object"
                Return token.Type = JTokenType.Object
            Case "enum"
                If token.Type <> JTokenType.String Then Return False
                If paramType.EnumValues IsNot Nothing AndAlso paramType.EnumValues.Count > 0 Then
                    Return paramType.EnumValues.Contains(token.ToString())
                End If
                Return True
            Case Else
                Return True
        End Select
    End Function

End Class
