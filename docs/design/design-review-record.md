# AI Native Harness 设计评审记录

| 项 | 内容 |
|---|---|
| 评审日期 | 2026-07-11 |
| 评审方式 | 按 `README.md` 序 0–12 统一评审 |
| 结论 | **有条件同意全套**（整改项已回写各专项，见 §3） |
| 编码 | 评审通过表示目标设计可作为实现约束；具体编码顺序以 [`../ai-native-harness-implementation-design.md`](../ai-native-harness-implementation-design.md) 为准 |

---

## 1. 分项评审结论

| 序 | 文档 | 结论 | 要点 |
|---:|---|---|---|
| 0 | ai-native-harness-design.md | **同意** | 唯一 Harness、对标矩阵、H0–H5 可执行 |
| 1 | office-harness-api.md | **同意** | 冻结模式 B、同 session 单 running、Adapter 一期 |
| 2 | context-pack-schema.md | **有条件同意→已整改** | 补决策摘要；阈值与 Safety/Excel 对齐 |
| 3 | skill-runtime-and-gates.md | **有条件同意→已整改** | 补决策摘要；S0–S2 时间语义澄清 |
| 4 | safety-policy.md | **同意** | VBA 默认关、T1/T2、硬门禁；超时=拒绝 |
| 5 | tool-result-observation.md | **有条件同意→已整改** | 补决策摘要；写工具 observation 硬要求 Debug 断言 |
| 6 | memory-runtime.md | **同意** | memory_item 单写；短期不进 RAG；隐私 P0 |
| 7 | model-gateway-streaming.md | **同意** | IModelClient；Pump 降级；ReAct 上移 |
| 8 | run-trace-storage.md | **有条件同意→已整改** | 补决策摘要；与 conversation_event 分工写死 |
| 9 | eval-golden-set.md | **有条件同意→已整改** | 补决策摘要；PR=L0、Nightly=L1 子集 |
| 10 | word-capability-map.md | **有条件同意→已整改** | 补决策摘要；快路径必须 Trace |
| 11 | excel-table-agent-runtime.md | **有条件同意→已整改** | 补决策摘要；阈值对齐 Safety；PQ 不做 |
| 12 | ppt-deck-agent-runtime.md | **有条件同意→已整改** | 补决策摘要；删页审批；禁死胡同文案 |

**总评**：无驳回项。阻塞 H0 的仅为「文档间默认决策未冻结」——本记录整改后视为解除。

---

## 2. 跨文档一致性问题（评审发现）

| ID | 问题 | 整改 |
|---|---|---|
| C1 | 多份文档缺「决策摘要」勾选，无法闭环评审 | 各专项补 §决策摘要（冻结默认） |
| C2 | Excel 大表阈值 vs Safety T1/T2 表述不完全一致 | 统一引用 Safety 阈值为权威；Excel 分块阈值交叉引用 |
| C3 | Harness 注入列表未列 `IMemoryRuntime` | office-harness-api 注入列表补齐 |
| C4 | ErrorCode 分散在 Observe/Safety/Gateway | 增加共享错误码索引表（本记录 §4 + 总纲附录引用） |
| C5 | 开放问题「建议默认」未提升为「已冻结」 | 关键默认提升为设计决策 |
| C6 | 版本号均为 v0.1，整改后无法区分 | 专项升为 **v0.2（评审修订）** |
| C7 | 审批模式 B 仅在 API 写清，其它文档偶发「阻塞等待」歧义 | 全局注明：审批不阻塞 `RunAsync` Task 挂起，用状态机挂起 run |
| C8 | 无 Skill 时 VisibleTools=全 app 工具 vs Safety 收紧 | 冻结：v1 全 app 工具 + Safety；写入 skill 专项决策 |

---

## 3. 整改清单（已落实到 md）

| # | 动作 | 目标文件 |
|---|---|---|
| R1 | 增加设计决策 + 决策摘要 | context / skill / observe / run-trace / golden / word / excel / ppt |
| R2 | 注入 `IMemoryRuntime`；全局审批模式 B 脚注 | office-harness-api.md |
| R3 | Excel 阈值节引用 safety-policy T1/T2 | excel-table-agent-runtime.md |
| R4 | 共享 ErrorCode 速查 | design/README.md + 本记录 §4 |
| R5 | 评审状态改为 v0.2 评审修订 | 各专项文头 + README |
| R6 | 总纲 §11/决策表 与评审结论对齐 | ai-native-harness-design.md |
| R7 | README 增加「评审结论」与设计决策总表 | design/README.md |

---

## 4. 共享 ErrorCode 速查（跨专项权威）

| Code | 主要来源 | 含义 |
|---|---|---|
| COM_ERROR | Observe/Classifier | Office COM 失败 |
| NETWORK_ERROR | Gateway/Classifier | HTTP/网络 |
| TIMEOUT | Gateway/Classifier | 超时/取消 |
| JSON_ERROR | Observe/Gateway | 解析失败 |
| ARGUMENT_ERROR | Observe | 参数非法 |
| NOT_FOUND | Observe | 对象/工具不存在 |
| TOOL_NOT_ALLOWED | Skill Gate | 超出 allowed-tools |
| SAFETY_BLOCKED | Safety | 拒绝执行 |
| SAFETY_NEEDS_APPROVAL | Safety | 待批准（非终态错误时可恢复） |
| HOST_UNSUPPORTED | Observe/Excel PQ | 宿主不支持 |
| NO_PROGRESS | Loop | 无进展熔断 |
| PARTIAL_APPLY | Observe | 部分成功 |
| VERIFY_FAILED | Observe | 成功标准未满足 |
| DOC_PROTECTED / DOC_MISSING | Safety/API | 保护或无文档 |
| VBA_DISABLED | Safety | 宏关闭 |
| RANGE_TOO_LARGE | Safety/Excel | 超阈值 |
| AUTH_FAILED / RATE_LIMITED / PROVIDER_ERROR / PARSE_ERROR | Gateway | 供应商侧 |
| USER_CANCELLED | API | 用户取消 |
| PROCESS_ORPHAN | RunTrace | 崩溃遗留 run |

新增 code 必须：Classifier 或专项登记 + 本表更新。

---

## 5. 设计决策（目标态，H0 起逐步落地）

> 本表是目标约束，不代表当前代码全部实现。已实现/未实现状态见 §7 与落地实施设计。

| ID | 决策 |
|---|---|
| D1 | 唯一入口 `IOfficeHarness.RunAsync`；Surface 禁止直调 Executor |
| D2 | 审批 **模式 B**：`RunAsync` 返回 `awaiting_approval`，不长期占用调用方同步阻塞 |
| D3 | 同 session 同时仅 1 个 running run |
| D4 | VBA / ExecuteVBA 默认关闭 |
| D5 | Skill `allowed-tools` 硬拒绝（`TOOL_NOT_ALLOWED`） |
| D6 | 新记忆只写 memory_item 管线；atomic 只读淘汰 |
| D7 | 短期记忆不进 RAG |
| D8 | 写工具必须 observation（含 diff 或明确 noop） |
| D9 | 快路径 Capability 必须写 Trace/ToolResult |
| D10 | HttpStreamService 目标态 = UI Pump；ReAct/MCP 循环在 Harness |
| D11 | Agent 默认 JSON Plan；toolCalling=auto 可升原生 tools |
| D12 | PowerQuery 完整执行 v1 **不做** |
| D13 | PR CI = L0 + smoke + build-code；L1 Nightly |
| D14 | 无 Skill 命中时 VisibleTools = 全 app 工具，仍过 Safety |
| D15 | Safety T1/T2 为影响面权威阈值；Excel 分块策略服从之 |

---

## 6. 签字栏（可选）

| 角色 | 姓名 | 日期 | 结论 |
|---|---|---|---|
| 架构 | （评审代理已完成文档整改） | 2026-07-11 | 有条件同意→整改完成 |
| 产品 | | | |
| 客户端 | | | |

---

## 7. 实现审计追加（2026-07-12）

| 项 | 结论 |
|---|---|
| H0 状态 | 已完成 adapter 版正式入口：`AgentKernelService.StartAgentAsync` 经 `OfficeHarness.RunAsync` 调用现有 `AgentKernel` |
| 已落地 | `ToolResult` 错误契约、原生 Office ToolResult 回灌、Word fast-path 结构化结果、VSTO shadow-copy 工具定位、0 迭代失败防误报、执行期 `allowed-tools` 硬拒绝、轻量 RunTrace、Word 基础写工具 before/after/diff Observation |
| 工程门禁 | `scripts/run-code-checks.ps1`、`scripts/audit-p0-guardrails.ps1` 已落地。本行原先还写了 `.github/workflows/code-checks.yml`，2026-10-25 复核仓库磁盘与 `git ls-files .github` 都没有这个文件（`.github/` 只有 `copilot-instructions.md`），所以它是当时就没落地的期望；带退出码的入口只有 `build/RunReleaseChecks.ps1` |
| Ralph | 主 HTML 不再加载 `ralph-loop.js`；后端 `startLoop` 仅保留兼容 shim |
| 仍未完成 | 完整 ContextHub、审批控制 API、`agent_run_event`/回放 UI、Excel/PPT 最小 Diff、Excel/PPT capability harness |

本追加不改变 D1–D15；它只记录设计与代码实态之间的同步状态。实现任务拆分见 [`../ai-native-harness-implementation-design.md`](../ai-native-harness-implementation-design.md)。

*本记录是专项文档的上级裁决；与专项冲突时以本记录设计决策为准，并应回写专项。*
