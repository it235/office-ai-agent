# OFFICE AI AGENT KNOWLEDGE BASE

**Generated:** Sat Jan 24 2026
**Commit:** N/A
**Branch:** main

> 请用中文与我交互

## OVERVIEW
Office AI Agent is a Visual Studio solution containing multiple VSTO add-ins for Microsoft Office applications (Excel, Word, PowerPoint) with shared components. Built with Visual Basic.NET and VSTO, it provides AI-driven assistance for office automation tasks.


## STRUCTURE
```
./
├── ExcelAi/          # Excel add-in with data analysis capabilities
├── WordAi/           # Word add-in with document processing features
├── PowerPointAi/     # PowerPoint add-in with presentation tools
├── ShareRibbon/      # Shared components and core services
└── OfficeAgent/      # Installation package
```

## Important
ShareRibbon中的html非常重要，是与用户交互的入口，html采用了office virtual server，css/js都使用virtual server来访问

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Core shared functionality | ShareRibbon/ | Referenced by all Office add-ins |
| Excel-specific features | ExcelAi/ | Data analysis, cell operations |
| Word-specific features | WordAi/ | Document processing, text manipulation |
| PowerPoint-specific features | PowerPointAi/ | Presentation creation, slide operations |
| MCP protocol integration | ShareRibbon/Mcp/ | Model Context Protocol client implementation |
| UI components | ShareRibbon/Controls/ | Shared user interface elements |
| Configuration management | ShareRibbon/Config/ | API keys, settings, prompts |
| AI communication | ShareRibbon/Controls/Services/ | HTTP streaming, message handling |

## CODE MAP
| Symbol | Type | Location | Refs | Role |
|--------|------|----------|------|------|
| ShareRibbon | Library | ShareRibbon/ | 3 | Shared components for all add-ins |
| BaseChatControl | Class | ShareRibbon/Controls/ | 3 | Core chat UI component |
| HttpStreamService | Class | ShareRibbon/Controls/Services/ | 1 | AI API communication |
| MCPConnectionConfig | Class | ShareRibbon/Mcp/ | 1 | MCP protocol configuration |
| ConfigManager | Class | ShareRibbon/Config/ | 1 | Settings management |
| ExcelAi | Add-in | ExcelAi/ | 0 | Excel-specific functionality |
| WordAi | Add-in | WordAi/ | 0 | Word-specific functionality |
| PowerPointAi | Add-in | PowerPointAi/ | 0 | PowerPoint-specific functionality |

## CONVENTIONS
- All Office add-ins reference ShareRibbon for shared functionality
- UI components are in ShareRibbon/Controls/
- Configuration management is in ShareRibbon/Config/
- MCP protocol implementation is in ShareRibbon/Mcp/
- Services are in ShareRibbon/Controls/Services/
- Each add-in has its own ribbon implementation
- WebView2 is used for modern UI rendering

## ANTI-PATTERNS (THIS PROJECT)
- Direct Office interop calls should go through shared services
- Configuration should be managed through ConfigManager
- MCP connections should use StreamJsonRpcMCPClient
- UI updates should happen on the main thread

## UNIQUE STYLES
- Heavy use of VSTO for Office integration
- Shared ribbon components across Office applications
- WebView2 for modern HTML-based UI
- MCP protocol for AI model integration
- DeepSeek API integration

## COMMANDS
```bash
# Build solution
msbuild AiHelper.sln

# Build individual projects
msbuild ExcelAi/ExcelAi.vbproj
msbuild WordAi/WordAi.vbproj
msbuild PowerPointAi/PowerPointAi.vbproj
msbuild ShareRibbon/ShareRibbon.vbproj
```

## NOTES
- Requires Visual Studio 2022 with VSTO tools
- .NET Framework 4.7.2 dependency
- Office 2016+ required for deployment
- NuGet packages must be restored before building
- MCP protocol supports multiple AI backends

---

## LESSONS LEARNED / 经验与避坑

（基于 memory-skills-chatai-design 等近期实现总结，供后续开发与重构参考。）

### 安装与工程
- **vdproj 慎改**：安装包项目（OfficeAgent/*.vdproj）被自动修改后容易加载失败。如需改安装逻辑，尽量小范围编辑或单独分支；出问题可先回退 vdproj 恢复加载。

### 配置与 UI 形态
- **配置模态框**：布局和风格要与 Chat 主界面协调（间距、字体、按钮组）；原子记忆/用户画像等列表建议按当前应用过滤（如 Excel 只展示 `app_type=Excel` 的记录）。

### 前端资源与 Virtual Server
- **新增 JS/CSS 必须加入投放路径**：Chat 使用 Office virtual server 加载 html/css/js。新增脚本若未加入资源或路径错误，会出现「点击无反应、控制台 ERR_FILE_NOT_FOUND」。检查：该 js 是否被主 html 引用、是否通过 virtual server 可访问；嵌入资源需在对应 csproj 中正确配置。


### 数据库与升级
- **新字段要走迁移/ALTER TABLE**：新增表或字段时，应用 SQLite 迁移脚本（如 `ALTER TABLE ... ADD COLUMN`）并配合版本号或迁移记录做升级控制，避免「只在新环境建表」导致老用户升级后缺字段或报错。

### 意图识别与智能体流程
- **意图识别要结合上下文**：发送前应收集「内容区引用（选中/附件摘要）+ RAG 相关记忆 + 当前会话」，再调用 `IntentRecognitionService.IdentifyIntentAsync`；意图 LLM 的 context 中应包含 `referenceSummary`、`ragSnippets`（见 `EnrichContextForIntent`）。
- **意图 → 规划 → 执行一条龙**：意图确认后（尤其 Agent 模式）可进入「请求 Spec 步骤 → 解析 JSON steps → 按 RalphLoopController 逐步执行」；与 `/loop` 共用同一套规划格式与执行循环，避免两套引擎。
- **RAG/意图要在 UI 有反馈**：发请求前若用了 RAG，在 Chat 中给简短提示（如「已检索 N 条相关记忆」）；若有意图识别结果，展示「识别意图：xxx」，便于用户感知并信任行为。

### 多 Agent 协作编码避坑（2026-04-29 智能排版 V2 实战总结）

#### VB.NET 语法 — Agent 极易犯错

AI Agent 在写 VB.NET 时会混入 C# 语法。以下模式必须在代码审查时**主动扫描**：

| C# 写法（错误） | VB.NET 写法（正确） | 检测关键词 |
|----------------|-------------------|-----------|
| `var x = ...` | `Dim x = ...` | `var ` |
| `List<T>` | `List(Of T)` | `<T>` |
| `Dictionary<K,V>` | `Dictionary(Of K, V)` | `<K,V>` |
| `x ?? y` | `If(x, y)` | `??` |
| `x?.Prop` | `If(x?.Prop, ...)` 或先判 Nothing | `?.` 在复杂表达式中 |
| `new()` | `New T()` | `new()` |
| `string`/`int`/`bool`/`void` | `String`/`Integer`/`Boolean`/`Sub` | 小写类型名 |
| `=>` lambda | `Function(x) ...` 或 `Sub(x) ...` | `=>` |
| `for (int i=0; i<n; i++)` | `For i = 0 To n-1` | `for (` |

#### VB.NET 关键字冲突

以下 C# 无问题的标识符在 VB.NET 中**是关键字**，用作属性/枚举值必须加方括号：

| 关键字 | 正确写法 | 场景 |
|--------|---------|------|
| `Error` | `[Error]` | 枚举值/属性名 |
| `Resume` | `[Resume]` | 枚举值/属性名（错误处理关键字） |
| `Structure` | `[Structure]` 或改名 `DocStructure` | 属性名（建议改名避免） |
| `String`/`Integer` 等 | 可在适当上下文直接使用 | 类型名 vs 标识符 |

#### Async 语法限制

- `Async Sub ... As Task` **不合法**，必须用 `Async Function ... As Task`
- `Async Function` 中**不能有 `ByRef` 参数**
- `Async Function` 中的 `Return`（无值）可以用于提前退出，编译器自动处理

#### 多 Agent 独立编码的 API 不一致

当多个 Agent 同时创建有依赖关系的文件时，**类型 API 会不一致**。预防措施：

1. **先定义接口/契约**（数据模型类），再分派实现
2. **或**：先让一个 Agent 写出核心类型文件，验证编译后，再让其他 Agent 基于该文件编写消费者
3. **或**：所有 Agent 完成后，必须做一次 **API 一致性检查**：
   - 对比各文件对同一类型的属性/方法引用是否匹配
   - 用 `grep "Public (Function|Sub|Property|Class)"` 导出公开 API，交叉验证

#### 新文件必须注册到 .vbproj

- `.vb` 文件 → `<Compile Include="...">`
- `.js` 文件 → `<None Include="...">`
- 遗漏会导致「类型未定义」编译错误
- **检查方法**：`grep 新文件名.vb 项目.vbproj` 确认存在

#### Agent 编写规范

给 Agent 的 prompt 必须明确：
1. **语言**：必须写 VB.NET，不是 C#
2. **关键类型**：明确列出要复用的现有类型及其属性名
3. **文件路径**：给出完整绝对路径
4. **只新建/只修改**：明确边界，避免 Agent 自行扩展范围

#### 语义排版 AI 标注避坑（2026-04-29 实测发现）

**问题 1：AI 标注不认得原有标题 → 全部变成统一格式**
- 根因：`SemanticPromptBuilder` 只把纯文本发给 AI，AI 不知道原有样式名、字号大小
- 修法：Prompt 中为每个段落附加「原文样式: 标题 1」和「AI 自动检测到的标题结构」
- 要诀：**AI 需要看到原有格式线索才能做准确判断**，不能只喂纯文本

**问题 2：图片/表格段落以空文本发给 AI → 标注错乱**
- 根因：非文本段落的文本为空串（""），AI 看到 `[3] ` 不知道这是什么
- 修法：构建 Prompt 前过滤掉非文本段落，同时记录原始索引映射
- 要诀：**发往 AI 的段落必须是纯文本**，图片/表格段落应在渲染引擎层自动跳过

**设计原则：**
- 排版不是"推倒重来"，而是"识别 + 规范化"。保留原有结构 > 套模板
- AI prompt 是排版质量的核心瓶颈，不能吝啬上下文