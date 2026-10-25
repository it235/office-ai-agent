# CLAUDE.md

本文件为 Claude Code / AI Agent 在本仓库工作的项目级协议。请优先遵守本文件；进入具体目录后，再阅读距离目标文件最近的 `AGENTS.md` 作为局部补充。

## 交互与协作语言

- 默认使用中文与用户交流。
- 代码符号、文件路径、命令、API 名称保持原文。
- 修改代码前先确认目标目录职责、相关 `.vbproj` 注册方式、以及是否存在局部 `AGENTS.md`。

## Project Overview

Office AI 智能体是基于 **Visual Basic.NET + VSTO + .NET Framework 4.7.2** 的 Office AI 插件解决方案，为 Excel、Word、PowerPoint 提供 AI 驱动的辅助能力。

- **开发环境**：Windows + Visual Studio（VSTO 工作负载）+ .NET Framework 4.7.2
- **语言**：Visual Basic.NET
- **Office 集成**：VSTO，支持 Microsoft Office 2016+ / WPS Office
- **官网**：https://www.officeso.cn
- **License**：Apache 2.0

## Repository Structure

```text
AiHelper/
├── AiHelper.sln                         # 主解决方案
├── CLAUDE.md                            # AI Agent 项目级规则
├── AGENTS.md                            # 仓库级目录导航
├── ExcelAi/                             # Excel VSTO 插件
├── WordAi/                              # Word VSTO 插件
├── PowerPointAi/                        # PowerPoint VSTO 插件
├── ShareRibbon/                         # 共享组件库，所有 Office 插件引用
├── OfficeAgent/                         # Visual Studio Installer 安装包项目 (.vdproj)
├── docs/                                # 调试、迁移、说明文档
├── openspec/                            # 需求/变更规格与归档
├── packages/                            # NuGet packages 还原目录
└── .codegraph/                          # CodeGraph 语义索引（如存在）
```

各子目录可能包含自己的 `AGENTS.md`。进入目标目录工作时，应按以下顺序读取：

1. 根目录 `CLAUDE.md`
2. 根目录 `AGENTS.md`
3. 目标目录最近的 `AGENTS.md`
4. 相关 `.vbproj`、`packages.config`
5. 目标功能源码

## Tech Stack

| 类别 | 技术 |
|---|---|
| Framework | .NET Framework 4.7.2 |
| Language | Visual Basic.NET |
| Office Integration | VSTO (Visual Studio Tools for Office) |
| UI | WebView2 + HTML/CSS/JS + Office Virtual Server |
| Database | SQLite (`System.Data.SQLite`, `Microsoft.Data.Sqlite.Core`) + EntityFramework 6 |
| AI / Protocol | MCP (Model Context Protocol) via StreamJsonRpc |
| JSON | Newtonsoft.Json, System.Text.Json |
| Markdown | Markdig |
| Excel 特有 | ExcelDna.AddIn / ExcelDna.Integration |
| Word/文档处理 | DocumentFormat.OpenXml, AngleSharp, HtmlAgilityPack |
| Async/Runtime | Microsoft.Bcl.AsyncInterfaces, System.Threading.Tasks.Dataflow |

## Build & Development

### Prerequisites

- Windows + Visual Studio（带 VSTO 工作负载）
- .NET Framework 4.7.2 targeting pack
- Office 2016+ 或 WPS Office
- NuGet packages 已还原
- 构建安装包项目 `OfficeAgent/*.vdproj` 需要 Visual Studio Installer Projects 支持

### Build Commands

```bash
# 还原 NuGet 包
msbuild AiHelper.sln -t:Restore

# 构建整个解决方案
msbuild AiHelper.sln

# 构建共享库
msbuild ShareRibbon/ShareRibbon.vbproj

# 构建单个 Office 插件
msbuild ExcelAi/ExcelAi.vbproj
msbuild WordAi/WordAi.vbproj
msbuild PowerPointAi/PowerPointAi.vbproj
```

### Debug

- VSTO 迁移/兼容性问题（含 `OfficeAgent.vdproj` 在 VS 2026 打不开）查看 `docs/VS2026-InstallerProjects.md`。
- 插件启动、Ribbon、任务窗格、WebView2 相关问题通常需要在 Visual Studio + Office 宿主进程中调试。

## Architecture Rules

### 顶层职责

- `ShareRibbon/` 是共享核心：公共 UI、配置、MCP、AI 通信、WebView2/HTML 资源、共享服务应放这里。
- `ExcelAi/` 只放 Excel 特定逻辑：单元格、Range、Sheet、图表、ExcelDna 等。
- `WordAi/` 只放 Word 特定逻辑：文档、段落、翻译、续写、OpenXml 等。
- `PowerPointAi/` 只放 PowerPoint 特定逻辑：演示文稿、幻灯片、形状、续写、翻译等。
- `OfficeAgent/` 是安装包项目，`.vdproj` 慎改。
- `openspec/` 放需求、变更规格与归档，不是运行时代码。

### 共享与应用边界

- 可复用能力优先抽到 `ShareRibbon/`，不要复制到单个 Office 插件项目。
- 单个 Office 应用项目不要直接访问其他 Office 应用的对象模型。
- Office COM Interop 操作应保持应用边界清晰，并注意 UI 线程/宿主进程上下文。
- 配置/API Key/Prompt 相关逻辑优先查 `ShareRibbon/Config/`。
- MCP 协议相关逻辑优先查 `ShareRibbon/Mcp/`。

### 声明式 Office 对象操作硬规则

- 动态 API 能力必须接入现有主链：`BaseChatControl -> ChatRoutingOrchestrator -> AgentKernelService -> OfficeHarness -> AgentKernel -> LoopEngine -> ToolRegistry -> CodeExecutionService -> Host`。禁止复制或旁路其中任何一层。
- 只允许两个稳定的长尾入口：`DiscoverOfficeCapability`（只读发现）和 `OfficeObjectOperation`（声明式执行）。新增 SmartArt、艺术字、形状组合等能力时，默认扩展 Catalog/Resolver/Executor/Observer，不新增专用意图和 Tool。
- 共享 DTO/验证合同放 `ShareRibbon/Agent/OfficeOperations/`；具体 COM Catalog Provider、ObjectResolver、OperationExecutor、Observer 放对应 `WordAi/`、`ExcelAi/`、`PowerPointAi/` 项目。
- API Catalog 不得整体注入模型上下文；单次发现最多返回任务相关的少量成员。Catalog member 必须具有稳定 ID、参数 schema、风险等级和 executable 状态。
- 禁止任意反射执行。Executor 只能调用 Catalog 中 `Executable=True` 且通过参数绑定、宿主检查、SafetyGate 的成员；不得用 `CallByName`、VBA 或动态字符串路径绕过。
- 所有 Office COM 调用必须在宿主 UI 线程执行。Agent/Loop 可在后台运行，但进入 `ExecuteJsonCommandWithToolResult` 前必须统一 marshal；禁止各 Executor 私建线程调 COM。
- 通用 operation 成功必须返回 `ToolResult + Observation`，写操作至少包含 before/after、changed、targetRefs、warnings；原子调用成功不等于用户目标完成。
- Repair 继续使用现有 `LoopEngine`，最多按既定无进展阈值修复；不得在宿主 Executor 内嵌第二套模型调用或 repair loop。
- 现有 Tool 渐进迁移为 `OperationBatch` adapter，外部 Tool ID 保持兼容；迁移前后必须有 Golden 对比，禁止一次删除大批 Tool。
- 实施顺序和 Definition of Done 以 `docs/design/office-object-operation-integration.md` 为准。
- Chat UI、WebView2、HTTP/streaming 服务优先查 `ShareRibbon/Controls/`。

## Agent Working Protocol

- 修改前先阅读目标目录最近的 `AGENTS.md`。
- 多文件任务先计划再改，尤其是跨 `ShareRibbon` 和 Office 应用项目的改动。
- 新增 `.vb` 文件必须加入对应 `.vbproj` 的 `<Compile Include="...">`。
- 新增 `.js` / `.css` / `.html` 等前端资源必须加入对应 `.vbproj` 的 `<None Include="...">` 或 `<EmbeddedResource Include="...">`，并确认 HTML 引用路径能被 Office Virtual Server 访问。
- 多 Agent 并行开发时，先定义共享模型/接口/数据契约，再实现消费者；完成后做 API 一致性检查。
- 避免大范围格式化 `.vbproj`、`.vdproj`、生成文件、Designer 文件。
- 不要硬编码 API Key、用户配置或本机绝对路径。

## VB.NET / VSTO Pitfalls

AI Agent 容易混入 C# 写法。写 VB.NET 时必须主动避免：

| C# 写法（错误） | VB.NET 写法（正确） | 检测关键词 |
|---|---|---|
| `var x = ...` | `Dim x = ...` | `var ` |
| `List<T>` | `List(Of T)` | `<T>` |
| `Dictionary<K,V>` | `Dictionary(Of K, V)` | `<K,V>` |
| `x ?? y` | `If(x, y)` | `??` |
| `new()` | `New T()` | `new()` |
| `string` / `int` / `bool` / `void` | `String` / `Integer` / `Boolean` / `Sub` | 小写类型名 |
| `=>` lambda | `Function(x) ...` 或 `Sub(x) ...` | `=>` |
| `for (int i=0; i<n; i++)` | `For i = 0 To n - 1` | `for (` |

### VB.NET 关键字冲突

| 关键字 | 正确写法 | 场景 |
|---|---|---|
| `Error` | `[Error]` | 枚举值/属性名 |
| `Resume` | `[Resume]` | 枚举值/属性名 |
| `Structure` | `[Structure]` 或改名 | 属性名，建议改名 |

### Async 限制

- `Async Sub ... As Task` 不合法；返回 `Task` 必须写 `Async Function ... As Task`。
- `Async Function` 不能有 `ByRef` 参数。
- `Async Sub` 只适合事件处理器；普通异步流程应返回 `Task`。

### VSTO / COM 注意事项

- Office 对象模型调用要考虑宿主应用、活动文档/工作簿/演示文稿是否存在。
- UI 更新通常应在主线程/宿主 UI 上下文中进行。
- COM 对象生命周期要谨慎，避免长期持有无效 Office 对象引用。
- `Option Strict Off` 不代表可以忽略类型清晰度；新增代码应尽量显式、可读、便于编译检查。

### 事件绑定卫生

- `RemoveHandler` 只能摘掉 `AddHandler` 挂上去的那一份，**摘不掉 `Handles` 子句**的绑定。同一个处理器不要既有
  `Handles X.Event` 又有 `AddHandler X.AnotherEvent`——WebView2 上 `ChatBrowser.NavigationCompleted` 与
  `CoreWebView2.NavigationCompleted` 是两个事件，混用会让处理器重复执行，而「解绑」看似写了其实无效。
- 事件注入类初始化要用「文档粒度」闸门（见 `BaseChatControl._pageBootstrapDone`）：在装载新文档前复位、在处理器里
  **先判再占位、占位必须在第一个 `Await` 之前**，否则排队的第二次导航完成同样通过检查。
- 重复执行在这里的症状通常是「值被反复洗回」或「看起来没有副作用」，不是异常；不要用肉眼确认，用
  `scripts/smoke-event-binding.ps1` 的顺序/次数守卫。
- 读配置开关之前先确保配置文件被加载过（`ChatSettings.EnsureLoaded(...)`）；宿主在面板可见事件里判开关时，
  进程内可能还没有任何实例读过盘。

## Frontend Resource Rules

Chat 相关 HTML/JS/CSS 通过 Office Virtual Server 加载，是用户交互入口。新增或修改前端资源时必须检查：

1. 文件是否被主 HTML 正确引用。
2. 文件是否在对应 `.vbproj` 中注册。
3. Virtual Server 是否能访问该路径。
4. WebView2 控制台是否有 `ERR_FILE_NOT_FOUND` 或脚本错误。

常见错误表现：按钮点击无反应、前端没有日志、控制台资源 404、HTML 引用路径大小写/目录不匹配。

## Database / Migration Rules

- 新增 SQLite 字段必须设计升级迁移，例如 `ALTER TABLE ... ADD COLUMN`。
- 不要只改建表逻辑；老用户数据库升级路径必须可用。
- 新增表/字段应考虑版本号、迁移记录或幂等检查。
- 涉及配置/记忆/RAG/历史数据时，优先检查现有迁移与初始化逻辑。
- 用于过滤的枚举列（例如 `app_type`）必须有一份口径持有者，并且只在**写入边界**归一。读侧一般写成
  「等于当前取值 OR 空值豁免」，因此一个既不属于任何合法取值、又拿不到空值豁免的第三种值，会让那条数据在
  所有宿主下同时静默消失且不报错。收口顺序：先归一，再按证据强度回填，且**只填空、不覆盖已解析的值**，
  使整段可重复执行；同时同步 schema 版本断言、快照与门禁夹具。

## Installer / vdproj Rules

- `OfficeAgent/*.vdproj` 自动修改后容易加载失败；除非必要，避免大范围编辑。
- 如必须修改 `.vdproj`，保持最小 diff，不做格式化，不批量重排。
- 出现安装包加载失败时，优先回退最近的 `.vdproj` 改动。

## Verification Checklist

完成代码或文档改动前，按改动范围选择验证：

- VB.NET 语法没有 C# 污染。
- 新增 `.vb` / `.js` / `.css` / `.html` 已注册到对应 `.vbproj`。
- 前端资源已被 HTML 引用，并能通过 Virtual Server 访问。
- 共享 API 与消费者引用一致。
- 相关项目可构建，至少构建被修改项目及其依赖。
- 数据库字段/表变更有迁移路径。
- `.vdproj` 没有被无意改动或格式化。
- 文档没有过时日期、错误分支名或临时标记。
- 新增动态 Office API 能力没有创建平行 Harness/Loop/Safety/执行入口。
- `DiscoverOfficeCapability` 为只读且返回集合受限；`OfficeObjectOperation` 的每个 member 均来自 Catalog 并经过 Safety。
- 通用 COM Executor 在宿主 UI 线程执行，临时 COM 对象未跨 step/session 保存。
- 通用写操作有真实 Observation/Verifier，失败不会误报任务完成。
- 新增事件绑定只有一个绑定源、页面注入有文档粒度闸门，且读开关前配置已被加载过一次（`scripts/smoke-event-binding.ps1`）。
- 写过滤用枚举列（如 `app_type`）时都经过口径持有者归一，历史数据由「只填空、可重跑」的迁移收口，并且内联兜底与迁移文件逐表一致（`scripts/audit-p0-guardrails.ps1` E2-04 段）。

## CodeGraph Rules

本项目配置了 CodeGraph 语义索引时，请按以下规则使用：

### `.codegraph/` 存在时

- 主会话只直接使用轻量工具做定点查询：
  - `codegraph_search`：查符号位置
  - `codegraph_callers` / `codegraph_callees`：查调用关系
  - `codegraph_impact`：查修改影响面
  - `codegraph_node`：查单个符号详情
- 对于“解释某系统如何工作 / 探索某功能在哪里实现 / 大范围代码阅读”这类问题，主会话不要直接调用 `codegraph_context` 或 `codegraph_explore`；应派发 Explore agent，并要求其使用 `codegraph_explore` 作为主要工具。
- CodeGraph 结果来自 AST 索引；不要用 grep 重复验证同一结构性问题。

给 Explore agent 的推荐提示：

```text
This project has CodeGraph initialized (.codegraph/ exists). Use `codegraph_explore` as your PRIMARY tool — it returns full source code sections from all relevant files in one call.

Rules:
1. Follow the explore call budget in the `codegraph_explore` tool description.
2. Do NOT re-read files that codegraph_explore already returned source code for.
3. Only fall back to grep/glob/read for files listed under "Additional relevant files" if you need more detail, or if codegraph returned no results.
```

### `.codegraph/` 不存在时

如果项目没有 `.codegraph/`，先询问用户是否初始化：

```text
I notice this project doesn't have CodeGraph initialized. Would you like me to run `codegraph init -i` to build a code knowledge graph?
```
