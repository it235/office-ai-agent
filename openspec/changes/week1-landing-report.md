# Week 1 落地报告与真机验收清单

> 日期：2026-10-24
> 范围：E5-01 能力一致性、E5-03 PPT 假成功清理、E5-02 基线 Skill 结构操作、E2-01 会话按宿主过滤、E1-01/E1-02 写信任合同与单轮变更日志、E1-03 Word 本轮撤销、E1-04 取消/失败写入摘要。
> 结论：代码与自动化校验已通过；三端真机验收待执行。

## 1. 本轮改动

| 条目 | 落点 | 行为变化 |
| --- | --- | --- |
| E5-01 | `scripts/audit-p0-guardrails.ps1` P9/P10 | 校验 ToolRegistry、宿主白名单、Skill `allowed-tools` 三者一致；校验变更日志链路与隐私边界 |
| E5-02 | `ShareRibbon/Skills/excel-table-agent/SKILL.md`、`powerpoint-deck-agent/SKILL.md` | Excel 补齐建/删 Sheet、插入删除行列等结构工具；PPT 补齐删页/复制页/移动页/插形状 |
| E5-03 | `PowerPointAi/ChatControl.vb` | `SetSlideLayout` 先自定版式后内置版式并回读；`ApplyTheme` 仅接受真实 `.thmx` 文件；`AddSpeakerNotes` 回读校验，失败不再报成功 |
| E1-01 | `openspec/changes/write-trust-contract.md` | 固定成功、观察、回滚、撤销语义 |
| E1-02 | `ShareRibbon/Agent/Execution/AgentChangeLog.vb`、`ShareRibbon/Storage/AgentChangeLogRepository.vb`、`LoopEngine`、`AgentKernel`、`OfficeHarness` | 每个写步骤由宿主 Observation 聚合；轮次结束写入 `agent_run_change`；失败/取消保留已落地记录 |
| E2-01 | `ConversationRepository`、`ChatStateService`、`HistorySessionService`、`MemoryService`、`ChatContextBuilder`、`ExcelAi/ChatControl.vb` | 新消息写入 `app_type`；历史列表按宿主过滤；Excel 不再写 `Unknown`；旧行保持对三端可见 |
| E1-03 | `AgentChangeLog`、`AgentChangeLogRepository`、`AgentKernelService`、`BaseChatControl`、`WordAi/ChatControl.vb`、`agent-card.js` | 终态出现“撤销本轮”前必须存在完整撤销范围；Word 先核对当前文档指纹，再按撤销单元还原并复核；成功后标记 `rolled_back` |
| E1-04 | `LoopEngine`、`AgentKernel` | 取消/失败提示用真实残留数生成“已写入 N 处，未完成 M 步”；异常退出补充“异常前已写入 N 处” |
| 迁移 | `OfficeAiDatabase` 迁移 14/15/16 + `Storage/Migrations/014`、`015`、`016` | 14 回填 `conversation_event.app_type`；15 建 `agent_run_change`；16 增加 `rolled_back`、`undo_units`、`doc_hash_before`、`doc_hash_after`，并对重复加列迁移做幂等收敛 |

## 2. 自动化验证结果

- `build-code-projects.ps1 -Configuration Debug`：ShareRibbon / WordAi / ExcelAi / PowerPointAi 全部通过。
- `audit-p0-guardrails.ps1`：P0 守护通过（含 P9 能力一致性、P10 变更日志接线、迁移 016 与撤销链路）。
- `smoke-empty-db-initialization.ps1`：空库初始化 schema=16，13 张必需表与 `agent_run_change` 新增 4 列齐全。
- `smoke-db-schema-drift.ps1`：与 `OfficeAiDbSchema.current.sql` 无漂移（21 表 / 209 列 / 42 索引）。
- `smoke-change-log-and-session-scope.ps1`（新增）：
  - 仅 `kind=write/code_execution` 且带布尔 `changed` 的结果入日志；
  - 部分应用、回滚错误、撤销能力可持久化并按 `run_id` 顺序读回；
  - 批量观察以批次级 `undoUnits` 与整批首尾指纹为权威值，条目仅作明细；缺少批次级 `undoUnits` 时才回退为已提交条目求和；
  - 回滚成功或未改动文档的条目不计入残留，也不参与撤销单元与指纹推导；
  - `GetUndoScopeByRun` 能给出 2 单元撤销范围，`MarkRunRolledBack` 后不再提供撤销；
  - Excel/Word 历史互不串台，无 `app_type` 的旧会话三端可见；
  - 测试数据已清理。
- `run-code-checks.ps1 -SkipBuild`：全部冒烟、Golden L0 12 例、JS `node --check` 通过。

## 3. 三端真机验收

通用前置：关闭 Office；备份 `%USERPROFILE%\Documents\OfficeAiAppData\office_ai.db`（Release）或 `OfficeAiAppData-Debug`（Debug）；启动对应宿主后等待迁移完成。

### Word

1. 打开 Word，发起一次会话并让 Agent 执行一次写入（如插入文本）。
2. 打开历史列表：只应出现 Word 会话，不应出现此前 Excel/PPT 会话。
3. 执行一次会失败的写操作（如对不存在对象写入），确认提示不再报成功。
4. 查询 DB：
   ```sql
   SELECT run_id, seq, tool_id, changed, success, rolled_back, partial_apply,
          undo_available, undo_units, undo_point_id, doc_hash_before, doc_hash_after, summary
   FROM agent_run_change ORDER BY observed_at DESC LIMIT 20;
   ```
   失败步骤对应行应存在；`changed=1 AND success=0 AND rolled_back=0` 时 `partial_apply=1`；成功的 Word 写入应 `undo_available=1 AND undo_units>=1`。
5. 取消一次执行，确认提示为“已按你的请求停止：已写入 N 处，未完成 M 步”，且日志保留取消前已落地的写记录。
6. 完成一次包含多处写入的 Word 任务，点击终态卡片上的“撤销本轮”：文档应回到本轮开始前，按钮变为“↩ 已撤销本轮”，DB 中该 `run_id` 的 `rolled_back=1`。
7. 再次执行一轮写入，先手动编辑文档，再点“撤销本轮”：应明确拒绝并提示文档已有其他改动，文档不得被 `Undo` 改坏。
8. 没有可用撤销单元的写入（例如脚本路径）完成后不应出现“撤销本轮”。

### Excel

1. 发起会话，用 Skill 执行“新建 Sheet / 删除 Sheet / 插入行 / 删除列”各一次。
2. 确认工作表结构真实变化，而不只是回复文本。
3. 历史列表只显示 Excel 会话；打开 Word/PPT 时应看不到该会话。
4. 查询 `agent_run_change`，确认每次结构操作都有记录，`target_refs_json` 指向对应 Sheet/区域。

### PowerPoint

1. `SetSlideLayout` 使用一个自定义版式名：确认版式真实变化；使用不存在的版式名：应明确失败，不得报成功。
2. `ApplyTheme` 只给主题名不放 `.thmx` 文件：应返回失败并提示需要主题文件；给出真实 `.thmx` 才允许成功。
3. `AddSpeakerNotes`：确认备注页文本真实写入；对没有备注占位符的版式应失败并说明原因。
4. 结构操作删页/复制页/移动页/插形状各一次，确认页数与内容变化。
5. 历史列表只显示 PowerPoint 会话。

### 升级迁移

1. 用旧版本产生的 DB（schema 13 或更低）启动一次三端任一宿主。
2. 确认 `schema_version=16`，`agent_run_change` 存在且含 `rolled_back`、`undo_units`、`doc_hash_before`、`doc_hash_after`。
3. 旧会话仍能看到；能匹配到 `agent_run` 的会话已回填 `app_type`，回填后只出现在对应宿主。

## 4. 已知边界（本轮不处理）

- 变更日志暂不做 UI 面板，先保证数据真实可查（见 `write-trust-contract.md` 第 6 节）。
- 只有 Word 提供“撤销本轮”；Excel/PPT 的 `undo_available=0` 时只提示用户手动检查，不渲染撤销按钮。
- Word 撤销要求本轮结束后文档未再变化；用户中途编辑过文档时会拒绝撤销，这是有意的安全边界。
- 无 `app_type` 且无 `agent_run` 可回填的极老会话，按设计对三端可见，避免历史丢失。
- 混合会话（升级前开始、升级后继续）因旧行无宿主，仍可能被多端看到；新会话不受影响。

## 5. 失败处理

真机若发现某端历史仍串台，优先查该端是否重写了 `GetOfficeAppType`，以及消息是否绕过 `ChatStateService.AddMessage` 直接写库。若变更日志缺行，先确认该步骤的 `ToolResult.Observation` 是否带 `kind` 和布尔 `changed`，而不是在持久化层补猜。

若“撤销本轮”未出现，按顺序检查：该步 `undo_units` 是否为 0、`doc_hash_before/after` 是否为空、`AgentKernelService.BuildUndoDescriptor` 是否因宿主不是 Word 而跳过。若点击后提示指纹不一致，说明本轮结束后文档又被修改，属于预期保护；不要让用户绕过该检查。
