# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

我在windows10上面开发，使用的开发工具是Visual Studio Community 2022 + Visual Basic.NET + VSTO，使用的语言是Visual Basic.NET，使用的框架是.NET Framework 4.7.2，使用的Office插件是VSTO，使用的Office版本是Office 2016+ / WPS。

- **官网**: https://www.officeso.cn
- **License**: Apache 2.0

## Repository Structure

```
office-ai-agent/
├── ExcelAi/           # Excel VSTO 插件
├── WordAi/            # Word VSTO 插件
├── PowerPointAi/      # PowerPoint VSTO 插件
├── ShareRibbon/       # 共享组件（核心逻辑、UI、服务、MCP等）
├── OfficeAgent/       # 安装包项目 (.vdproj)
└── AiHelper.sln       # 主解决方案文件
```

### Key Projects

| Project | Purpose |
|---------|---------|
| **ShareRibbon** | 所有插件共享的核心库 - UI组件(BaseChatControl)、配置管理(ConfigManager)、MCP客户端、数据库、AI通信服务等 |
| **ExcelAi** | Excel 特定功能 - 数据分析、单元格操作、ExcelDna函数 |
| **WordAi** | Word 特定功能 - 文档处理、文本操作 |
| **PowerPointAi** | PowerPoint 特定功能 - 演示文稿操作 |

## Tech Stack

- **Framework**: .NET Framework 4.7.2
- **Language**: Visual Basic.NET
- **Office Integration**: VSTO (Visual Studio Tools for Office)
- **UI**: WebView2 + HTML/CSS/JS (Office Virtual Server)
- **Database**: SQLite (System.Data.SQLite, EntityFramework 6)
- **AI Protocol**: MCP (Model Context Protocol) via StreamJsonRpc
- **Markdown**: Markdig
- **JSON**: Newtonsoft.Json, System.Text.Json

## Build & Development

### Prerequisites

- Visual Studio 2022 (with VSTO 工作负载)
- .NET Framework 4.7.2
- Office 2016+ / WPS

### Build Commands

```bash
# 还原 NuGet 包
# (在 Visual Studio 中右键解决方案 -> "还原 NuGet 包")

# 构建整个解决方案
msbuild AiHelper.sln

# 构建单个项目
msbuild ShareRibbon/ShareRibbon.vbproj
msbuild ExcelAi/ExcelAi.vbproj
msbuild WordAi/WordAi.vbproj
msbuild PowerPointAi/PowerPointAi.vbproj
```

### Important Configuration

- **ShareRibbon HTML/JS/CSS**: 必须配置为嵌入资源或通过 Office Virtual Server 访问，新增资源需在 `.vbproj` 中正确配置
- **vdproj 安装项目**: 谨慎修改，容易加载失败，出问题先回退
- **SQLite 迁移**: 新增字段需通过 `ALTER TABLE` 迁移脚本处理

## Key Architecture

### ShareRibbon Core Services

| Namespace | Purpose |
|-----------|---------|
| `ShareRibbon.Config` | 配置管理 (ConfigManager, PromptManager, API设置) |
| `ShareRibbon.Controls` | UI组件 (BaseChatControl, BaseDeepseekChat, BaseDoubaoChat) |
| `ShareRibbon.Controls.Services` | 服务 (HttpStreamService, MessageService, IntentRecognitionService, MemoryService, McpService) |
| `ShareRibbon.Mcp` | MCP协议实现 (StreamJsonRpcMCPClient, MCPConnectionConfig) |
| `ShareRibbon.Storage` | 数据存储 (OfficeAiDatabase, MemoryRepository, ConversationRepository) |
| `ShareRibbon.Loop` | Ralph Loop 智能体 (RalphLoopController, RalphAgentController) |
| `ShareRibbon.Ribbon` | 共享 Ribbon 基类 (BaseOfficeRibbon) |

### Office Application-Specific

Each Office app plugin (ExcelAi/WordAi/PowerPointAi) references ShareRibbon and provides:
- 继承自 `BaseOfficeRibbon` 的功能区实现
- 继承自 `BaseChatControl` (或 `BaseDeepseekChat`/`BaseDoubaoChat`) 的聊天面板
- 应用特定的 JSON 命令模式与直接操作服务

## Important Conventions & Pitfalls

1. **前端资源**: 新增 JS/CSS/HTML 必须在 `ShareRibbon.vbproj` 中配置为 `None`/`EmbeddedResource`，并确保通过 Office Virtual Server 可访问
2. **安装项目**: `OfficeAgent/OfficeAgent.vdproj` 慎改，自动修改后易加载失败
3. **数据库迁移**: 新字段要走 `ALTER TABLE` 迁移，不要只在新环境建表
4. **意图识别**: 需要结合 `referenceSummary` + `ragSnippets` + 当前会话上下文
5. **中文交互**: 请用中文与项目维护者和代码注释交互

## Reference

- **AGENTS.md**: 更详细的代码库知识库
- **.github/copilot-instructions.md**: Copilot 指令
