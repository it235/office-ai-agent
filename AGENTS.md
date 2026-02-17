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