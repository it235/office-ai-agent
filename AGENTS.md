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