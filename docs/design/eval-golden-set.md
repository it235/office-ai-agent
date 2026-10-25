# 专项设计：黄金场景评测集（Golden Set）

| 项 | 内容 |
|---|---|
| 版本 | **v0.2（评审修订）** |
| 状态 | 目标设计已评审；Golden L0 已实现 |
| 实现状态 | **L0 已实现**：`tests/golden/l0/catalog.json` 与 `scripts/run-golden-l0.ps1` 已接入 code checks，覆盖 Skill gate、VBA 默认关闭、宿主隔离、审批恢复、ToolResult Observation、ContextPack 和 Harness 审批 API。**E11 v1 已实现**：`tests/golden/tasks/` 26 个真实任务 + JSON 夹具规格 + `scripts/New-GoldenFixtures.ps1` + 静态门禁 `scripts/check-golden-corpus.ps1` + 运行对比 `scripts/Compare-GoldenRuns.ps1` + L1 手册 `tests/golden/README.md`。L1 真宿主执行（E11-05）按当前指示延后，尚无任何通过率基线。 |
| 总纲 | [`../ai-native-harness-design.md`](../ai-native-harness-design.md) §7 / §9 / U1–U10 |
| 关联 | [`run-trace-storage.md`](./run-trace-storage.md)、Observe / Safety / Skill 专项 |
| 现有 | `scripts/smoke-*.ps1`（契约/DB/网关，**非**场景 Agent） |

---

## 1. 目标与非目标

### 1.1 目标

1. 用 **固定文档 + 固定用户话术 + 可机器判定的成功标准** 锁住 Harness 回归。  
2. 覆盖 Word / Excel / PPT 核心故事（U1–U10 及扩展）。  
3. 评分可在 CI 跑 **无 UI** 的子集；完整集可在本地 Office 宿主跑。  
4. 与 RunTrace 对齐：断言 step/tool/error/observation，而不是只看最终聊天字符串。  

### 1.2 非目标

- 不做在线 LLM 裁判为主分（成本与不稳）；LLM-as-judge 仅可选辅助。  
- 不要求每 PR 跑满含真实 API 的全集（可 nightly）。  
- 不替代单元测试与 schema smoke。  

### 1.3 原则

| 原则 | 说明 |
|---|---|
| 确定性优先 | 能用文档差分/结构断言就不用模型打分 |
| 最小文档 | 夹具小、可进仓库（注意许可证） |
| 可重复 | 同模型温度 0 或固定 mock |
| 分层 | L0 无网络 mock → L1 真 API 沙箱 → L2 手工 UX |
| 失败可诊断 | 输出 run_id / 期望 vs 实际 tool 序列 |

---

## 2. 评测分层

| 层 | 名称 | 网络 | Office | 何时跑 |
|---|---|---|---|---|
| **L0** | Contract + Mock Agent | 否 | 否 | 每 PR |
| **L1** | Host Golden | 是（API） | 是（真 COM） | Nightly / 发版前 |
| **L2** | UX 手测清单 | 是 | 是 | 发版前 |
| **L3** | 可选 LLM-judge | 是 | 是 | 抽检 |

L0 验证：门禁、Plan 校验、Diff 算法、Safety 矩阵（夹具 JSON）。  
L1 验证：真实改文档是否达标。  

---

## 3. 目录布局（建议）

```text
fixtures/
  golden/
    README.md
    catalog.json                 # 全场景目录
    word/
      U2-reformat-basic/
        input.docx
        case.json
        expect/
          structure.json         # 可选期望大纲
      U6-undo-round/
        ...
    excel/
      U3-summary-chart/
        input.xlsx
        case.json
    ppt/
      U4-deck-from-notes/
        input.pptx               # 或空演示 + 源 md
        case.json
    shared/
      mocks/
        plan_responses/          # L0 用
```

仓库策略：二进制夹具尽量小；可用脚本从 markdown 生成 docx（实现阶段）。

---

## 4. case.json Schema

```json
{
  "id": "U3-summary-chart",
  "version": "1.0",
  "appType": "Excel",
  "layer": ["L1"],
  "title": "选区数据生成汇总与柱状图",
  "userText": "根据当前表做各地区销售额汇总，并插入柱状图",
  "setup": {
    "openFile": "input.xlsx",
    "selection": "Sheet1!A1:C50",
    "locale": "zh-CN"
  },
  "harnessOptions": {
    "temperature": 0,
    "allowVba": false,
    "skillHint": "excel-table-agent"
  },
  "expect": {
    "status": "completed",
    "maxSteps": 12,
    "maxRepairs": 3,
    "mustUseToolsAnyOf": [["DataAnalysis", "CreatePivotTable"], ["CreateChart"]],
    "mustNotUseTools": ["ExecuteVBA"],
    "forbiddenErrorCodes": ["SAFETY_BLOCKED", "TOOL_NOT_ALLOWED"],
    "document": {
      "assertions": [
        { "type": "sheet_exists", "name": "汇总" },
        { "type": "chart_count_gte", "sheet": "汇总", "count": 1 },
        { "type": "range_numeric", "ref": "汇总!B2", "min": 0 }
      ]
    },
    "trace": {
      "minWriteSteps": 1,
      "everyWriteHasObservation": true,
      "skillPrimaryIn": ["excel-table-agent"]
    }
  },
  "tags": ["excel", "chart", "P0"],
  "owner": "excel-team"
}
```

### 4.1 断言类型目录（可扩展）

#### 通用

| type | 参数 | 含义 |
|---|---|---|
| run_status | status | run 最终状态 |
| error_code_not | codes[] | 不得出现 |
| tool_used | toolId | 至少一次 |
| tool_not_used | toolId | 从未 |
| tool_sequence_contains | [a,b] | 子序列 |
| max_steps | n | |
| write_has_observation | bool | |
| skill_primary | name | |

#### Word

| type | 含义 |
|---|---|
| paragraph_style | ref + styleName |
| heading_count_gte | level + n |
| text_contains | ref 或全文 + substr |
| text_not_contains | |
| numbering_continuous | 范围 |
| font_name | ref + name |

#### Excel

| type | 含义 |
|---|---|
| sheet_exists | name |
| range_equals | ref + value |
| range_numeric | ref + min/max |
| formula_error_count | sheet + max |
| chart_count_gte | sheet + n |
| table_exists | name |

#### PPT

| type | 含义 |
|---|---|
| slide_count_gte | n |
| slide_title_contains | index + text |
| shape_count_gte | slide + n |

---

## 5. 场景目录（U1–U10 落地）

| ID | 故事 | App | 层 | 核心断言方向 |
|---|---|---|---|---|
| **U1** | 只选中不说话，自动处理 | Word/Excel | L1+L2 | 有 plan；有读或写；无阻塞确认 |
| **U2** | 报告套模板样式 | Word | L1 | 标题样式/字体；write observation |
| **U3** | 表汇总 + 图 | Excel | L1 | sheet/chart；无 VBA |
| **U4** | 纪要变 8 页 PPT | PPT | L1 | slide_count；标题非空 |
| **U5** | 执行中可见步骤 | 任意 | L2 | UI 手测；L1 查 run_step≥1 |
| **U6** | 一键撤销本轮 | Word | L1+L2 | undo 后 fingerprint 恢复 |
| **U7** | 记住偏好字体 | Word | L1 | 第二轮默认字体（依赖 memory 夹具） |
| **U8** | MCP 查数写入表 | Excel | L1 可选 | mock MCP；WriteData |
| **U9** | 低置信先读后写 | Word | L0+L1 | tool 序列读在写前 |
| **U10** | 高风险需确认 | Excel/PPT | L0+L1 | Delete* → approval 或 SAFETY |

### 5.1 补充 P0 场景（建议）

| ID | 说明 |
|---|---|
| W-proofread-basic | 校对应用 N 条建议 |
| W-numbering-fix | 编号连续 |
| E-formula-repair | 坏公式修复后 error_count=0 |
| E-large-range-chunk | 超 T1 触发分块或 approval |
| P-beautify-align | 美化后关键形状仍在 |
| S-skill-gate | 越权 tool 被拒 |
| S-vba-disabled | ExecuteVBA deny |

---

## 6. 运行器设计（逻辑）

```text
GoldenRunner
  Load catalog / filter tags
  For each case:
    Prepare host doc (copy temp)
    Apply selection
    Invoke IOfficeHarness.RunAsync(turn)   ' L1
      or MockPlanner+MockTools            ' L0
    Collect run_id from TraceStore
    Evaluate assertions → CaseResult
  Write report.md + junit.xml
```

### 6.1 L0 Mock 策略

- 固定 `plan_responses/*.json` 驱动 Loop  
- Executor 用内存 FakeHost（不 COM）验证门禁与状态机  
- Diff 用合成 before/after Pack  

### 6.2 L1 真宿主

- 启动要求：本机 Office + API Key  
- 超时：单 case 120s  
- 隔离：每次 temp 副本，跑完删除  

### 6.3 模型波动

| 手段 | 说明 |
|---|---|
| temperature=0 | 默认 |
| soft assert tool | `mustUseToolsAnyOf` 而非严格单工具 |
| retry 1 次 | 仅网络错误 |
| 记录 flaky 标签 | 连续失败才红 |

---

## 7. 评分与报告

### 7.1 CaseResult

```json
{
  "id": "U3-summary-chart",
  "passed": false,
  "score": 0.6,
  "failedAssertions": [
    { "type": "chart_count_gte", "expected": 1, "actual": 0 }
  ],
  "runId": "run_...",
  "durationMs": 45000
}
```

### 7.2 分数（可选加权）

| 项 | 权重 |
|---|---|
| status completed | 0.3 |
| document assertions | 0.4 |
| trace/tool 约束 | 0.2 |
| 步数/repair 预算 | 0.1 |

P0 场景要求 **硬通过**（passed=true），不用分数蒙混。

### 7.3 报告

- `build/reports/golden/latest.md`  
- 失败附：期望断言、实际 tool 列表、export run 路径  

---

## 8. CI 集成建议

| 流水线 | 内容 |
|---|---|
| PR | L0 + 现有 smoke + build-code |
| Nightly | L1 子集（U2,U3,U9,S-*） |
| Release | L1 全 P0 + L2 手测签核 |

无 Office 的 CI agent 不得跑 L1。

---

## 9. 与手测清单（L2）关系

L2 不替代黄金集，只补 UI：

- 时间线是否展示  
- 审批卡片文案  
- 撤销按钮  
- 上下文控制台字段  

清单可放 `fixtures/golden/manual-checklist.md`（实现时）。

---

## 10. 数据与隐私

- 夹具禁止真实客户数据  
- 报告禁止打印 API Key  
- 失败 export 走 RunTrace 脱敏  

---

## 11. 验收标准（本评测体系本身）

1. catalog 中每个 P0 case 有 case.json 且 schema 合法。  
2. L0 在无网络环境 3 分钟内跑完。  
3. 故意破坏 Skill 门禁时 S-skill-gate 失败（证明断言有效）。  
4. U9 能检测「先写后读」错误序列。  
5. 文档说明如何新增 case（作者模板）。  

---

## 12. 新增 case 作者模板

```text
1. 复制最近似目录
2. 最小化 input 文档
3. 手写 userText（真实用户口吻）
4. 先手跑一遍 Harness，记录合理 tool 序列
5. 把可机器检查的结果写成 assertions
6. 本地 L1 跑通再提交
```

---

## 13. 开放问题（已冻结）

| # | 问题 | **冻结默认** |
|---|---|---|
| Q1 | 夹具 docx 是否用 OpenXml 生成？ | **是**（优先），减少二进制噪声 |
| Q2 | 是否快照整份 output 文档？ | **否**，只断言关键点 |
| Q3 | 多语言模型是否分集？ | **先中文** |
| Q4 | 失败是否自动建 issue？ | **否** |

### CI 冻结（D13）

| 流水线 | 跑什么 |
|---|---|
| **PR** | L0 + 现有 smoke + `build-code` |
| **Nightly** | L1 子集：U2, U3, U9, S-skill-gate, S-vba-disabled |
| **Release** | L1 全 P0 + L2 手测签核 |

---

## 14. 决策摘要（评审）

- [x] 同意 L0/L1/L2 分层  
- [x] 同意 case.json 断言驱动  
- [x] 同意 PR≠L1 真宿主（**D13**）  
- [x] 同意 P0 场景硬通过、不用分数糊弄  
- [x] 同意依赖 RunTrace 做序列断言  

---

## 15. 落地顺序（实现时）

1. catalog + case schema + L0 runner  
2. FakeHost 门禁/序列断言  
3. Word U2/U9 L1  
4. Excel U3、PPT U4  
5. Nightly 脚本  
6. 报告与 flaky 标记  

---

## 16. 与其它专项的依赖

| 依赖 | 用途 |
|---|---|
| ContextPack | U1 无选区推断 |
| Observation | everyWriteHasObservation |
| Safety | U10 / VBA |
| Skill gates | S-skill-gate |
| RunTrace | 序列与导出 |

**没有 RunTrace，L1 断言会退化成脆弱的字符串匹配——优先落地存储埋点再铺场景。**

---

*黄金集是产品「对标 Copilot 体验 + Cursor 执行力」的客观尺子；场景可先少而狠（P0 十来个），再横向扩展。*
