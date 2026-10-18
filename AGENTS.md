# AGENTS.md

> 请用中文与用户交互。代码符号、路径、命令、API 名称保持原文。

本文件是仓库级 Agent 导航图，帮助 AI Agent 快速判断“要改某类功能应该看哪里”。工程级硬规则、构建验证、VB.NET/VSTO 避坑见根目录 `CLAUDE.md`。

## Purpose

- `CLAUDE.md`：项目级工作协议，规定必须遵守的规则、避坑和验证清单。
- `AGENTS.md`：仓库级导航，说明目录职责、任务入口和局部规则位置。
- 子目录 `AGENTS.md`：进入某个模块后的局部补充规则。距离目标文件最近的 `AGENTS.md` 优先作为局部上下文，但不得覆盖根 `CLAUDE.md` 的工程级安全/质量规则。

## Read Order for Agents

处理任务前按顺序读取：

1. 根目录 `CLAUDE.md`
2. 根目录 `AGENTS.md`
3. 目标目录最近的 `AGENTS.md`
4. 相关 `.vbproj` / `packages.config`
5. 目标功能源码

如果任务跨多个项目，例如同时修改 `ShareRibbon/` 和 `WordAi/`，需要同时阅读这些目录下的局部 `AGENTS.md`。

## Top-level Structure

```text
AiHelper/
├── AiHelper.sln                         # Visual Studio 解决方案
├── CLAUDE.md                            # AI Agent 项目级规则
├── AGENTS.md                            # 仓库级导航
├── ExcelAi/                             # Excel VSTO 插件
├── WordAi/                              # Word VSTO 插件
├── PowerPointAi/                        # PowerPoint VSTO 插件
├── ShareRibbon/                         # 共享组件库
├── OfficeAgent/                         # 安装包项目 (.vdproj)
├── docs/                                # 调试、迁移、说明文档
├── openspec/                            # 需求/变更规格与归档
├── packages/                            # NuGet packages 目录
└── .codegraph/                          # CodeGraph 语义索引（如存在）
```

## Where to Look

| 任务 | 位置 | 说明 |
|---|---|---|
| 共享 Chat UI / WebView2 / 服务 | `ShareRibbon/Controls/` | `BaseChatControl`、HTTP/streaming 服务、共享 UI 组件 |
| 前端 HTML/JS/CSS 资源 | `ShareRibbon/Resources/` 及相关项目资源项 | Office Virtual Server 加载入口，改动后检查 `.vbproj` 注册和 HTML 引用 |
| 配置、API Key、Prompt | `ShareRibbon/Config/` | 配置管理和持久化逻辑 |
| MCP 协议集成 | `ShareRibbon/Mcp/` | StreamJsonRpc / MCP client 相关实现 |
| Excel 功能 | `ExcelAi/` | 单元格、Range、Sheet、图表、ExcelDna、批量数据生成 |
| Word 功能 | `WordAi/` | 文档、段落、续写、翻译、OpenXml 处理 |
| PowerPoint 功能 | `PowerPointAi/` | 演示文稿、幻灯片、形状、续写、翻译 |
| PowerPoint Professional Design Agent | `PowerPointAi/Design/` | Scene 语义、设计系统、专业构图、COM 渲染、视觉验证与安全回滚 |
| 安装包定义 | `OfficeAgent/` | `.vdproj` 安装项目，慎改，保持最小 diff |
| 调试/迁移文档 | `docs/` | Visual Studio/VSTO 调试、迁移指南等 |
| 需求/变更规格 | `openspec/` | 变更说明、规格草案、归档记录 |
| AI Native Harness 总设计 | `docs/ai-native-harness-design.md` | 对标 Copilot/Claude/Cursor 的平台设计与细化清单 |
| AI Native Harness 落地实施设计 | `docs/ai-native-harness-implementation-design.md` | 基于现有代码的渐进改造任务、文件范围与验收清单 |
| Harness 专项设计 | `docs/design/` | ContextPack / Observe / Safety 等目标合同级专项 |
| 声明式 Office 对象操作 | `docs/design/office-object-operation-integration.md` | 基于现有 Harness/ToolRegistry/Loop 接入动态 API 发现、通用对象操作和宿主 Executor |
| PowerPoint Deck Agent 运行时 | `docs/design/ppt-deck-agent-runtime.md` | `CreateSlides + Scene + Observer/Verifier/Repair` 的专业设计主链 |
| ToolResult 与视觉证据 | `docs/design/tool-result-observation.md` | Observe/Repair 合同、Diff、临时多模态证据与持久化边界 |
| NuGet 依赖 | 各项目 `packages.config` 与根 `packages/` | 判断技术栈、依赖版本、目标框架 |

## Local AGENTS.md Files

当前仓库存在多层局部 Agent 文档：

- `ExcelAi/AGENTS.md`
- `WordAi/AGENTS.md`
- `PowerPointAi/AGENTS.md`
- `ShareRibbon/AGENTS.md`
- `ShareRibbon/Controls/AGENTS.md`
- `ShareRibbon/Mcp/AGENTS.md`
- `ShareRibbon/Config/AGENTS.md`

进入对应目录前必须阅读局部文件。局部文件用于说明该模块的入口、约束和反模式；根目录 `CLAUDE.md` 仍是通用质量规则来源。

## Project Ownership Boundaries

### ShareRibbon

`ShareRibbon/` 是跨 Office 插件共享的核心库。以下能力优先放这里：

- 共享 Chat 控件与 WebView2 UI 基础设施
- AI 请求、streaming、HTTP 通信
- MCP client 与工具调用协议
- 配置、Prompt、API Key 管理
- 可被 Excel/Word/PowerPoint 共同复用的服务、模型、工具类

不要把共享服务复制到单个 Office 插件项目中。

### ExcelAi

`ExcelAi/` 只处理 Excel 相关能力：

- Workbook / Worksheet / Range / Cell 操作
- 数据分析、公式辅助、图表生成
- ExcelDna 自定义函数
- Excel 专属 Ribbon、任务窗格、命令 schema

不要访问 Word 或 PowerPoint 对象模型。

### WordAi

`WordAi/` 只处理 Word 相关能力：

- Document / Paragraph / Range / Selection 操作
- 文档生成、续写、翻译
- OpenXml 文档处理
- Word 专属 Ribbon、任务窗格、命令 schema

不要访问 Excel 或 PowerPoint 对象模型。

### PowerPointAi

`PowerPointAi/` 只处理 PowerPoint 相关能力：

- Presentation / Slide / Shape 操作
- 幻灯片生成、排版、续写、翻译
- PowerPoint 专属 Ribbon、任务窗格、命令 schema

不要访问 Excel 或 Word 对象模型。

### OfficeAgent

- `OfficeAgent/` 的 `.vdproj` 是安装包定义，自动化修改风险高，除非必要不要大范围编辑。
- 仓库曾有 `OfficeAgentSetupCustomActions/`，它未被解决方案和安装包引用、也不参与任何构建校验，已于 2026-10-18 删除。
- 若将来需要可代码化的安装动作，新建自定义动作项目并同时接入 `AiHelper.sln` 与 `.vdproj` 的 `CustomAction` 节点，否则等于写了一套永不验证的代码。
- 打 MSI 前必须先 `build-installer-prep`（Release 代码 + SourcePath 审计），见 `docs/build-and-installer.md`。

## AI Native Architecture Rules

本项目的核心方向是 AI Native Office 插件。新增或重构智能能力时，默认采用 `Harness -> Agent -> Loop -> Capability/Skill -> Executor -> Observe/Repair/Explain` 的结构，而不是在入口处继续堆叠关键词 `if/else`。

### 设计原则

- Chat、Ribbon、快捷入口只负责收集上下文、调用 harness、展示结果；不要在入口层写大量业务判断。
- 意图识别应优先交给 planner/harness/agent 结合当前 Office 上下文、选区、文档结构和历史会话来判断。
- 确定性规则只允许作为轻量安全门、成本优化门或兜底门；不要把它扩展成主逻辑。
- 新能力应登记为可发现的 capability/skill，包含名称、适用场景、输入 schema、执行器、验证器和可解释输出。
- Agent Loop 至少包含 `plan -> act -> observe -> repair/continue -> explain`。执行失败时应基于观察结果修复计划，而不是直接把错误抛给用户。
- 能读当前 Office 文档上下文的地方，必须先读上下文再决策；不要让用户重复描述插件已经能读取的信息。
- 用户给出明确操作目标时，默认进入执行/预览流程；不要退化成普通聊天长篇解释工具边界。
- 需要澄清时，只问会阻塞执行的最小问题；可推断、可预览、可撤销的操作应先形成计划。

### 模块边界

- `ShareRibbon/` 可以承载共享抽象，例如上下文模型、capability 契约、loop 状态、MCP/Skill 协议、通用 UI 组件。
- `ShareRibbon/` 不应引入 Word/Excel/PowerPoint 的具体 COM 实现。
- Word 具体执行器、文档读取、编号、排版、校对等能力放在 `WordAi/`；已有方向可参考 `WordAi/Services/WordActionHarness.vb`。
- Excel、PowerPoint 的具体能力分别放在 `ExcelAi/`、`PowerPointAi/`，通过相同抽象接入共享 harness/loop。
- `openspec/` 中的 AI Native 设计文档是产品/架构意图来源；实现时要优先保持与 roadmap 一致。

### 视觉证据与多模态修复

- 宿主 Observer/Verifier 负责从真实 Office 渲染结果采集视觉证据；`ShareRibbon/` 只定义宿主无关的证据合同和模型请求适配，不得引用 PowerPoint/Word/Excel COM 类型。
- 截图、Base64、Data URL 等大载荷只能作为当前执行轮次的临时内存证据；禁止写入 `Observation`、`Data`、`Artifacts`、History、Memory、RunTrace、日志或 UI 消息。
- 多模态判断和修复必须由现有 `LoopEngine` 发起，并继续经过共享 `AiGateway`；Executor 只执行、观察和回滚，不得直接调用模型或建立第二套视觉 Repair Loop。
- 视觉证据必须限制格式、数量、尺寸和字节数，并在模型或 Provider 不支持图片时有限次降级到文本 Observation；禁止因视觉通道不可用形成无限重试。
- PowerPoint 渲染失败需要回滚时，必须先采集必要证据，再按安全标记回滚本批生成对象；不得为了保留截图而遗留不确定的部分修改。

### 动态 Office API 与声明式操作

- 长尾 Office 能力必须沿现有唯一主链接入：`OfficeHarness -> AgentKernel -> LoopEngine -> ToolRegistry -> CodeExecutionService -> 宿主 ExecuteJsonCommandWithToolResult`。
- 禁止为动态能力另建第二套 Harness、Planner、Loop、ToolBroker、SafetyGate 或宿主旁路执行入口。
- 动态发现统一使用稳定工具 `DiscoverOfficeCapability`；声明式执行统一使用 `OfficeObjectOperation`。不得为 SmartArt、艺术字、时间线等长尾需求继续新增意图枚举或一需求一 Tool。
- API Catalog 可以枚举宿主 Type Library，但一次只向 Agent 返回少量相关成员；禁止把 Word/Excel/PowerPoint 全量 API 注入 Prompt。
- `ShareRibbon/` 只承载 `OfficeObjectRef`、`OfficeOperationBatch`、Capability/Observation 等共享合同；PowerPoint/Word/Excel 的 Catalog Provider、ObjectResolver、OperationExecutor、Observer 必须放在各自宿主项目。
- Catalog 中“可发现”不等于“可执行”。未评级、宏、外部进程、退出/覆盖类成员默认拒绝，所有写操作继续经过现有 `SafetyGate`。
- 通用 COM 操作必须通过共享宿主 UI 线程入口执行；禁止在 Executor 内各自创建线程或绕过 `CodeExecutionService`。
- 现有高频 Tool 保持兼容，稳定后逐个转为 `OperationBatch` adapter；禁止大爆炸式删除或一次跨三宿主迁移。
- 权威实施方案见 `docs/design/office-object-operation-integration.md`。

### 新能力落地清单

新增一个 AI Native 能力时，至少补齐：

1. capability/skill 描述：自然语言触发范围、输入输出、风险级别。
2. context reader：能读取选区、全文、结构、样式或宿主状态。
3. planner/agent：基于上下文产生结构化计划，而不是字符串拼接命令。
4. executor：执行最小可验证动作，支持预览或撤销。
5. observer/verifier：验证实际 Office 文档是否达到目标。
6. repair loop：失败时让 AI 根据观察结果修复计划。
7. explanation：向用户解释“准备改什么、已经改什么、哪里需要确认”。

## Skills Official Authoring Rules

Skill 是 Office Agent 能力发现和扩展的核心入口，不是关键词路由表。新增 Excel/Word/PowerPoint 智能能力时，应优先落成目录型 Skill，并由 harness/agent 根据当前 Office 上下文选择、加载和执行。

### 目录结构

目录型 Skill 必须使用如下结构：

```text
ShareRibbon/Skills/<skill-name>/
├── SKILL.md
├── references/        # 可选：长说明、领域规则、示例，命中后再加载
├── scripts/           # 可选：可执行辅助脚本
└── assets/            # 可选：模板、示例或静态资源
```

`SKILL.md` 必须以简洁 front matter 开头，至少包含：

```yaml
---
name: excel-table-agent
description: Use when Excel needs table understanding, calculation, charting, cleanup, or multi-step spreadsheet automation.
---
```

本项目推荐补充：

```yaml
application: Excel
tags: excel, formula, chart
allowed-tools: ApplyFormula, CreateChart, TransformData
intent_types: data_analysis, formula, chart
```

### 运行时要求

- 第一阶段只加载 Skill 目录和 front matter 元数据，用于低成本召回；不要把完整长文档全部塞进系统提示词。
- 第二阶段只对命中的 Skill 加载 `SKILL.md` 正文、`references/`、`scripts/` 和 `assets/`。
- `allowed-tools` 是工具边界，agent 只能在声明的工具集合中规划动作；需要新执行能力时先补 Tool schema 和 executor。
- `application` 用于隔离宿主范围。Excel Skill 不得假设 Word/PPT COM 能力，Word/PPT 同理。
- Skill 文档描述“何时使用、如何判断、如何计划、如何观察、如何修复”，不要写成用户话术关键词列表。
- 可确定的兜底规则可以存在，但只能服务于安全、成本和失败恢复，不能替代 Skill 选择与 Agent Loop。

## Common Agent Mistakes

修改本仓库时重点避免：

1. 把 C# 语法写进 VB.NET，例如 `var`、`List<T>`、`??`、`=>`、`new()`。
2. 新增 `.vb` 文件后忘记加入 `.vbproj` 的 `<Compile Include="...">`。
3. 新增 `.js` / `.css` / `.html` 后忘记加入 `.vbproj`，或没有被 HTML 引用。
4. 修改前端资源但没有考虑 Office Virtual Server 访问路径，导致 WebView2 控制台 `ERR_FILE_NOT_FOUND`。
5. 在 Excel/Word/PowerPoint 项目里复制本应位于 `ShareRibbon/` 的共享逻辑。
6. 跨 Office 应用直接访问对象模型，例如在 `WordAi/` 中访问 Excel 对象。
7. 新增 SQLite 字段只改建表逻辑，没有写升级迁移。
8. 大范围格式化 `.vdproj`、`.vbproj`、Designer 文件或生成文件。
9. 多 Agent 并行创建类型时，没有先统一共享模型/接口，导致 API 不一致。
10. 用关键词 `if/else` 堆叠用户意图，绕过 harness/agent/loop，导致场景永远补不完。
11. 明明能读取 Office 上下文，却让用户反复回答文档范围、编号格式、选区内容等插件可自行观察的信息。
12. 对明确执行请求只输出聊天解释或 JSON 示例，不进入预览、执行、观察和修复流程。
13. 为每个新 Office 对象或业务话术新增意图枚举和专用 Tool，而不先评估 `DiscoverOfficeCapability + OfficeObjectOperation`。
14. 在通用 Executor 中开放任意字符串反射、`CallByName`、宏或未经过 Catalog/Safety 的 COM 成员。
15. 新建平行 Harness/Loop/Safety 来实现动态 API，造成两套状态机、两套终态和不同错误合同。
16. 把 PowerPoint 截图或 Base64 塞进 Observation、Artifacts、History、Trace 或日志，造成隐私、体积和持久化风险。
17. 在 PowerPoint Executor/Verifier 内直接调用视觉模型，绕过现有 Loop 的重试、降级和终态管理。

## Commands

常用命令索引如下；更完整规则见 `CLAUDE.md`。
**构建二分（P0-5）**：日常只编代码，不要用整解决方案 Rebuild 判断代码是否可编。详见 `docs/build-and-installer.md`。

```bash
# 还原 NuGet 包
msbuild AiHelper.sln -t:Restore

# 推荐：只构建四个代码项目（不含 OfficeAgent.vdproj）
.\build-code.bat
# 或
powershell -File .\scripts\build-code-projects.ps1 -Configuration Debug

# 打 MSI 前准备：Release 代码 + 安装 SourcePath 审计（不生成 MSI）
.\build-installer-prep.bat
# 或
powershell -File .\scripts\build-installer-prep.ps1

# 完整 Release 门禁（构建 + 审计 + smoke 等）
powershell -File .\build\RunReleaseChecks.ps1

# 构建共享库 / 单插件（MSBuild 直调）
msbuild ShareRibbon/ShareRibbon.vbproj
msbuild ExcelAi/ExcelAi.vbproj
msbuild WordAi/WordAi.vbproj
msbuild PowerPointAi/PowerPointAi.vbproj

# 安装包 MSI：在 Visual Studio 中打开 OfficeAgent/OfficeAgent.vdproj 并 Build（需 Installer Projects 扩展）

# 本地 VSTO 开发证书（*.pfx 不入库；克隆后若清单签名失败请运行）
powershell -File .\scripts\ensure-vsto-temp-keys.ps1

# Release 签名（正式证书经环境变量，禁止用 TemporaryKey）
# OFFICE_AI_SIGN_CERT_THUMBPRINT=... 或 OFFICE_AI_SIGN_PFX=...
powershell -File .\build\SignArtifacts.ps1
```

签名策略见 `docs/signing-and-certificates.md`。禁止提交 `*.pfx` / `*.snk` / 正式签名私钥。

## Documentation Maintenance Rules

- 不要在本文件写容易过时的生成日期、当前分支、临时 commit 信息。
- 新增顶层目录、重要子系统或局部 `AGENTS.md` 后，应更新本文件的导航表。
- 技术栈、构建命令、验证清单的权威位置是 `CLAUDE.md`；本文件只保留导航级摘要。
- 如果局部 `AGENTS.md` 与实际代码明显不一致，优先相信代码和项目文件，并修正文档。
