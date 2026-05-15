# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Office AI 智能体是基于 **Visual Studio Community 2026 + Visual Basic.NET + VSTO** 开发的 Office AI 插件，为 Excel/Word/PowerPoint 提供 AI 驱动的辅助功能。

- **开发环境**: Windows 10 + Visual Studio Community 2026 + VB.NET + VSTO + .NET Framework 4.7.2
- **支持 Office**: Microsoft Office 2016+ / WPS Office
- **官网**: https://www.officeso.cn
- **License**: Apache 2.0

## Repository Structure

```
AiHelper/
├── ExcelAi/                    # Excel VSTO 插件
│   ├── ChatControl.vb         # Excel 聊天面板
│   ├── Ribbon1.vb             # Excel 功能区
│   ├── ExcelJsonCommandSchema.vb   # Excel JSON 命令模式
│   ├── ExcelDirectOperationService.vb  # Excel 直接操作服务
│   ├── ExcelDnaFunctions.vb   # Excel DNA 函数
│   └── BatchDataGenerationForm.vb  # 批量数据生成
├── WordAi/                    # Word VSTO 插件
│   ├── ChatControl.vb        # Word 聊天面板
│   ├── Ribbon1.vb             # Word 功能区
│   ├── WordJsonCommandSchema.vb    # Word JSON 命令模式
│   ├── WordDocumentTranslateService.vb  # Word 翻译服务
│   ├── WordCompletionManager.vb     # Word 续写管理
│   └── OpenXmlWordTranslator.vb     # Word 翻译器
├── PowerPointAi/              # PowerPoint VSTO 插件
│   ├── ChatControl.vb        # PowerPoint 聊天面板
│   ├── Ribbon1.vb             # PowerPoint 功能区
│   ├── PowerPointJsonCommandSchema.vb  # PowerPoint JSON 命令模式
│   ├── PowerPointCompletionManager.vb   # PPT 续写管理
│   └── PowerPointDocumentTranslateService.vb  # PPT 翻译服务
├── ShareRibbon/               # 共享组件（所有插件引用）
├── OfficeAgent/               # 安装包项目 (.vdproj)
└── AiHelper.sln              # 主解决方案文件
```

## Tech Stack

| 类别 | 技术 |
|------|------|
| **Framework** | .NET Framework 4.7.2 |
| **Language** | Visual Basic.NET |
| **Office Integration** | VSTO (Visual Studio Tools for Office) |
| **UI** | WebView2 + HTML/CSS/JS (Office Virtual Server) |
| **Database** | SQLite (System.Data.SQLite, EntityFramework 6) |
| **AI Protocol** | MCP (Model Context Protocol) via StreamJsonRpc |
| **JSON** | Newtonsoft.Json, System.Text.Json |
| **Markdown** | Markdig |

## Key Features

### Core Capabilities
- **AI 翻译**: 多语言、多模型 AI 翻译（段落/整页）
- **智能排版**: 文档/PPT 内容排版，支持模板定制
- **AI 续写**: 智能续写文档或 PPT 内容
- **内容校对**: 错句词审阅修订
- **Ralph Loop**: 首个将智能规划 Ralph Loop 集成到 Office 的插件
- **MCP 客户端**: 支持 MCP-Server 配置
- **DeepSeek/Doubao 增强**: 针对这两个模型优化

### Excel 特有功能
- 数据分析、图表生成、公式辅助
- ALLM/CLLM 函数（ExcelDNA）
- 批量数据生成
- 选中单元格/Sheet 引用分析

### Word 特有功能
- 文档处理、内容生成/补全
- OpenXml 翻译引擎
- 多语言文档翻译

### PowerPoint 特有功能
- 演示文稿创建、幻灯片设计
- PPT 内容续写与翻译

## Build & Development

### Prerequisites
- Visual Studio 2026 (with VSTO 工作负载)
- .NET Framework 4.7.2
- Office 2016+ / WPS
- NuGet packages (restore before build)

### Build Commands
```bash
# 还原 NuGet 包
msbuild AiHelper.sln -t:Restore

# 构建整个解决方案
msbuild AiHelper.sln

# 构建单个项目
msbuild ShareRibbon/ShareRibbon.vbproj
msbuild ExcelAi/ExcelAi.vbproj
msbuild WordAi/WordAi.vbproj
msbuild PowerPointAi/PowerPointAi.vbproj
```

### Debug
- 使用 `docs/VisualStudio调试问题诊断.md` 排查 VSTO 调试问题
- 查看 `docs/VS2025-VSTO-Migration-Guide.md` 了解 VSTO 迁移指南

## Key Architecture

### ShareRibbon Core Namespaces

| Namespace | Purpose |
|-----------|---------|
| `ShareRibbon.Config` | 配置管理 (ConfigManager, PromptManager, ModelApiClient, ChatSettings) |
| `ShareRibbon.Controls` | UI组件 (BaseChatControl, BaseDeepseekChat, BaseDoubaoChat) |
| `ShareRibbon.Controls.Services` | 核心服务 (HttpStreamService, MessageService, IntentRecognitionService, MemoryService, McpService, AgentKernelService) |
| `ShareRibbon.Controls.Models` | 数据模型 (SelectionInfo, HistoryMessage, ExecutionStep) |
| `ShareRibbon.Mcp` | MCP协议 (StreamJsonRpcMCPClient, MCPConnectionConfig) |
| `ShareRibbon.Storage` | 数据存储 (OfficeAiDatabase, ConversationRepository, FormatTemplateRepository) |
| `ShareRibbon.Loop` | Ralph Loop 智能体 (RalphLoopController, RalphAgentController) |
| `ShareRibbon.Ribbon` | 共享 Ribbon 基类 (BaseOfficeRibbon) |
| `ShareRibbon.Agent` | 智能体核心 (AgentKernel, AgentMemory, LoopEngine, ToolRegistry, SkillRegistry) |
| `ShareRibbon.Services` | 业务服务 (EmbeddingService, SkillsService, UnifiedMemoryService) |
| `ShareRibbon.Services.Reformat` | 排版服务 (DocumentParserService, FormatPreviewService) |
| `ShareRibbon.Translate` | 翻译服务 (TranslateActionForm, TranslateSettings) |
| `ShareRibbon.Formatting` | 格式化服务 |
| `ShareRibbon.Common` | 公共工具 (LLMUtil) |
| `ShareRibbon.Log` | 日志 (SimpleLogger) |
| `ShareRibbon.Prompts` | 提示词模板 |
| `ShareRibbon.Skills` | Skills 定义 |
| `ShareRibbon.Tools` | 工具类 |
| `ShareRibbon.Continuation` | 续写服务 (ContinuationService) |
| `ShareRibbon.Resources` | 共享资源 (ShareResources) |

### Office Application-Specific Structure

Each Office app plugin (ExcelAi/WordAi/PowerPointAi) references ShareRibbon and provides:

1. **Ribbon 实现** (`Ribbon1.vb`): 继承 `BaseOfficeRibbon`
2. **聊天面板** (`ChatControl.vb`): 继承 `BaseChatControl` / `BaseDeepseekChat`
3. **JSON 命令模式** (`*JsonCommandSchema.vb`): 应用特定的 JSON 命令解析
4. **文档服务** (`*DocumentTranslateService.vb`): 应用特定的翻译服务
5. **续写服务** (`*ContinuationService.vb`): 应用特定的续写服务
6. **完成管理** (`*CompletionManager.vb`): AI 补全管理

## Important Conventions & Pitfalls

### 前端资源 (HTML/JS/CSS)
- Chat 使用 Office Virtual Server 加载 `html/css/js`
- 新增脚本必须加入 `.vbproj` 的 `<None Include="...">` 或 `<EmbeddedResource>`
- 检查：JS 是否被主 html 引用、是否通过 virtual server 可访问
- 错误表现：「点击无反应、控制台 ERR_FILE_NOT_FOUND」

### vbproj 文件注册
- `.vb` 文件 → `<Compile Include="...">`
- `.js` 文件 → `<None Include="...">`
- 遗漏会导致「类型未定义」编译错误

### 安装项目 (vdproj)
- `OfficeAgent/*.vdproj` 被自动修改后容易加载失败
- 如需修改，尽量小范围编辑或单独分支；出问题先回退

### 数据库迁移
- 新增字段必须走 `ALTER TABLE` 迁移脚本
- 配合版本号或迁移记录做升级控制
- 避免「只在新环境建表」导致老用户升级后缺字段

### VB.NET 语法规范（AI Agent 极易犯错）
| C# 写法（错误） | VB.NET 写法（正确） | 检测关键词 |
|----------------|-------------------|-----------|
| `var x = ...` | `Dim x = ...` | `var ` |
| `List<T>` | `List(Of T)` | `<T>` |
| `Dictionary<K,V>` | `Dictionary(Of K, V)` | `<K,V>` |
| `x ?? y` | `If(x, y)` | `??` |
| `new()` | `New T()` | `new()` |
| `string`/`int`/`bool`/`void` | `String`/`Integer`/`Boolean`/`Sub` | 小写类型名 |
| `=>` lambda | `Function(x) ...` | `=>` |
| `for (int i=0; i<n; i++)` | `For i = 0 To n-1` | `for (` |

### VB.NET 关键字冲突
| 关键字 | 正确写法 | 场景 |
|--------|---------|------|
| `Error` | `[Error]` | 枚举值/属性名 |
| `Resume` | `[Resume]` | 枚举值/属性名 |
| `Structure` | `[Structure]` 或改名 | 属性名（建议改名） |

### Async 语法限制
- `Async Sub ... As Task` 不合法，必须用 `Async Function ... As Task`
- `Async Function` 中不能有 `ByRef` 参数

### 意图识别流程
- 需要结合 `referenceSummary` + `ragSnippets` + 当前会话上下文
- 发送前应收集「内容区引用 + RAG 相关记忆 + 当前会话」
- 意图确认后可进入「请求 Spec 步骤 → 解析 JSON steps → 按 RalphLoopController 逐步执行」

### 中文交互
- 请用中文与项目维护者交互
- 代码注释也使用中文

## Reference Documents

| 文档 | 说明 |
|------|------|
| `AGENTS.md` | 更详细的代码库知识库 |
| `docs/roadmap.md` | 产品路线图 |
| `docs/01Memory.md` | 记忆能力设计文档 |
| `docs/02SkillsConfig.md` | Skills 配置文档 |
| `docs/03ChatAi.md` | Chat AI 设计文档 |
| `docs/agent-architecture-redesign.md` | 智能体架构重设计文档 |
| `docs/smart-reformat-design.md` | 智能排版设计文档 |
| `docs/smart-reformat-v2.md` | 智能排版 V2 设计文档 |
| `docs/CODE_ANALYSIS.md` | 代码分析与优化报告 |
| `docs/VS2025-VSTO-Migration-Guide.md` | VSTO 迁移指南 |
| `docs/WordAi-Startup-Optimization.md` | Word 启动优化文档 |
| `docs/需求整理与设计方案.md` | 需求整理与设计方案 |
| `docs/VisualStudio调试问题诊断.md` | VS 调试问题诊断 |
| `.github/copilot-instructions.md` | Copilot 指令 |

## 代码审查重点 (AI Agent 协作时)

1. **VB.NET 语法检查**: 确认没有 C# 语法污染
2. **新文件注册**: 确认新增 .vb/.js 文件已加入 .vbproj
3. **API 一致性**: 多 Agent 协作时验证类型 API 一致性
4. **上下文完整**: 给 Agent 的 prompt 必须包含语言、关键类型、文件路径、边界
