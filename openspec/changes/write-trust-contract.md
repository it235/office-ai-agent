# 写操作信任合同（E1-01）

> 适用范围：Word / Excel / PowerPoint 三端所有会修改用户文档的执行路径，包括宿主 JSON 命令、脚本/VBA 执行、声明式 `OfficeObjectOperation`。
> 实现方：宿主 `ExecuteJsonCommandWithToolResult`、`CodeExecutionService.ExecuteCodeWithToolResult`、各 `OfficeOperationExecutor`。
> 校验方：`scripts/audit-p0-guardrails.ps1`。

## 1. 成功语义

一次写操作只有在「宿主已执行且已观察到结果」时才能返回 `ToolResult.Success = True`。

- 预览被用户拒绝：`ErrorCode = CANCELLED`，`changed = false`，`recoverable = false`，不得报成功。
- 执行成功但快照未检测到变化：允许成功，但 Observation 必须带 `warnings`（如“未检测到变化”），不得静默当成已修改。
- 执行结果无法观察（快照采集失败）：`ErrorCode = OBSERVATION_FAILED`，`recoverable = false`，不得把“可能已执行”报成干净成功。
- 只弹提示、没有真实执行的操作（例如只有内置主题名称的 `ApplyTheme`）：必须返回失败并说明原因，禁止假成功。

## 2. Observation 合同

每个写路径的 `ToolResult.Observation` 必须是一个对象，至少包含：

| 字段 | 类型 | 含义 |
| --- | --- | --- |
| `kind` | string | `write`（宿主命令）或 `code_execution`（脚本/VBA） |
| `summary` | string | 本次操作给用户看的简短结果 |
| `changed` | bool | 是否真实观察到文档变化 |
| `targetRefs` | array | 变更目标引用（Range/Slide/Shape/Document 等） |
| `undoUnits` | int | 宿主可撤销单元数；0 表示该步不能安全撤销 |

失败、部分应用或结果不确定时，还必须包含：

```json
"rollback": { "succeeded": false, "attempted": false, "verified": false, "partialApply": true, "error": "..." }
```

规则：

- `changed = true` 且 `Success = false` ⇒ `ErrorCode = PARTIAL_APPLY`，`recoverable = false`（禁止自动重放，避免在已修改的文档上再执行一次）。
- 没有回滚手段的宿主必须如实声明 `attempted = false`，不得声称已还原。
- 截图、Base64、Data URL 只能作为当前轮次的内存证据，禁止进入 Observation、Data、Artifacts、History、Memory、RunTrace、日志或 UI 消息。

## 3. 撤销语义

- `ToolResult.UndoPointId` 只有在宿主确实能撤销“本次操作范围”时才允许填写。
- Word 的每条写命令使用独立 `UndoRecord`，失败时只回滚本轮记录；Observation 必须给出 `undoUnits`（一条自定义撤销记录 = 1 个单元）。
- Excel 普通命令与 PowerPoint 通用命令当前没有可安全限定范围的撤销实现：`undoAvailable = false`，UI 只能提示用户检查后手动撤销。
- `UndoManager` 仅作为提示与记账，不得在未验证范围时自动调用 `Application.Undo`。

### 3.1 本轮撤销（E1-03，Word）

`agent_run_change` 记录 `undo_units`、`doc_hash_before`、`doc_hash_after` 后，才允许对整轮写入执行一次撤销：

1. 只统计 `changed=1 AND rolled_back=0` 的残留记录。
2. 每条残留记录都必须 `undo_available=1 AND undo_units>0`，且有本轮开始与结束的文档文本指纹；否则不出现“撤销本轮”。
3. 撤销前比对当前文档指纹与 `doc_hash_after`：不一致说明用户之后又改过文档，直接拒绝，绝不盲目 `Undo`。
4. 按 `undo_units` 总数调用 `Document.Undo(n)`，再比对文档指纹与 `doc_hash_before`；不一致时如实报失败，不宣称已恢复。
5. 成功后把该 `run_id` 的记录标记 `rolled_back=1`，同一轮不会再次提供撤销。
6. Excel/PowerPoint 当前不提供该能力，UI 不渲染按钮。

## 4. 单轮变更日志（E1-02）

每完成一个写步骤，`LoopEngine` 从 Observation 生成一条 `AgentChangeRecord`，聚合为 `AgentRunChangeLog`：

- 只收录 `kind ∈ {write, code_execution}` 且带布尔 `changed` 的观察；读取、聊天、无观察依据的调用不进入日志。
- Word 批量命令的聚合观察同样使用 `kind=write`：批次级 `undoUnits` 是逐条汇总后的唯一权威值，`before`/`after` 取整批开始与结束时的文档指纹；`items` 只作明细。回滚成功或未改动文档的条目不得参与撤销单元与指纹推导，避免多撤一步或拿中途哈希当基线。
- 失败与取消也保留已落地记录，供“已写入 N 处，未完成 M 步”使用（E1-04）。
- 回滚成功的记录 `rolledBack=true`，不计入 `WrittenCount`，也不计入 `partialApply`。
- `AgentKernel` 把日志挂到 `AgentResult.ChangeLog`；`AgentKernelService` 在渲染终态前先持久化一次，`OfficeHarness` 随后以同一 `change_id` 幂等覆盖。
- 读取入口：`AgentChangeLogRepository.GetByRun(runId)`；撤销范围入口：`GetUndoScopeByRun(runId)`。

### 4.1 失败/取消摘要（E1-04）

取消与失败消息只根据真实残留计算：

- 有残留：`已按你的请求停止：已写入 N 处，未完成 M 步。`
- 无残留但有执行：`已按你的请求停止，尚未修改文档；未完成 M 步。`
- 异常退出：`执行异常: [...] ...（异常前已写入 N 处）`。

`N` 取 `AgentRunChangeLog.WrittenCount`（排除已回滚），`M` 取未完成/未跳过步骤数。禁止用“执行完成”掩盖失败，也不把已回滚的步骤算作已写入。

## 5. 静态校验

`scripts/audit-p0-guardrails.ps1` 必须持续校验：

1. 三端失败/部分应用仍携带 rollback 说明与非自动重放标记（P2-C、P3-B）。
2. `CodeExecutionService` 的脚本观察包含 `kind` 与 `changed`。
3. `LoopEngine` 调用 `session.ChangeLog.AddFromStep`，`AgentKernel` 把日志挂到结果。
4. `OfficeHarness` 的两条完成路径都持久化变更日志。
5. 变更日志实现中不出现截图/Base64 字段。
6. 迁移 `016` 存在，`AgentKernelService` 在终态前落库并传递撤销描述。
7. `agent:undoRun` 路由存在，Word `PerformUndoAgentRun` 读取撤销范围并核对首尾文档指纹，成功后标记回滚。
8. `LoopEngine` 的取消/失败摘要使用 `WrittenCount`。

## 6. 非目标

- 本期不实现跨文档撤销、历史版本管理、协作审阅合并。
- 本期不把变更日志做成新的 UI 面板；先保证数据真实、可查、可接 UI。
- 本期只有 Word 提供“撤销本轮”；Excel/PowerPoint 保持只提示、不自动撤销。
