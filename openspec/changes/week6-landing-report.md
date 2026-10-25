# Week 6 落地记录：黄金评测集 v1（E11-01 / E11-02 / E11-03 / E11-04）

> 本轮只做「语料 + 静态门禁 + 运行记录对比 + L1 手册」。
> 真机首跑（E11-05）按当前指示延后，因此本文件不包含任何通过率数字，也不把 `not-run` 写成成绩。

## 1. E11-01：26 个真实文档任务

`tests/golden/tasks/catalog.json` 是唯一清单：

| 宿主 | 用例数 | 覆盖 |
|---|---|---|
| Word | 9 | 选区校对、会议纪要改写、表格校对、只读摘要、撤销本轮、编号检查（gap）、整篇清空负样例、含糊指令负样例、越权脚本负样例 |
| Excel | 9 | 同比增长列、销售表排版、只读洞察、异常值解读、补空值、预算公式修复、选区建图、VBA 透视表负样例、整本工作簿排版负样例 |
| PowerPoint | 8 | 本页要点、本页校对、选中文字改写、纪要成稿、整稿美化、品牌模板（gap）、批量删页负样例、整页清空负样例 |

验收对照：

- ≥20 个任务：26 个；
- 三端各 ≥4：9 / 9 / 8；
- ≥5 个安全负样例：7 个，覆盖越权（`S-skill-gate`）、VBA（`S-vba-disabled`）、高危删除（`destructive`）、超大范围（`large-range`）、含糊指令（`vague`）；
- ≥15 个 `ready`：24 个；
- gap ≥1 且指名持有者：2 个
  - `W-numbering-check` → `ShareRibbon/Agent/QuickActionCatalog.vb`，计划行 `E6-07`，探针 `word.selection.numbering` / `word.numbering-check`；
  - `P-brand-template-gap` → `PowerPointAi/Design`，计划行 `E6-07`，探针 `ppt.slide.brand-template` / `ppt.brand-template`。
- 每个任务都有中文用户话术、夹具与选区、可机器判定断言、步数与修复预算。

**没有新造入口、包、Skill 或工具 id**。用例只引用既有 `entryActionId` / `ribbonId` / `packageId` / `skill`；引用不到的两条如实标 `gap`。gap 的探针一旦可解析，门禁会变红，逼着把用例升级为 `ready`，避免留下永远跑不了的样例。

### 夹具

`tests/golden/tasks/fixtures/` 共 9 份 JSON 规格：

- Word：`word/report-zh`、`word/contract-zh`、`word/meeting-zh`；
- Excel：`excel/sales-zh`、`excel/messy-zh`、`excel/budget-zh`；
- PowerPoint：`ppt/meeting-notes-zh`、`ppt/brand-deck-zh`、`ppt/blank-zh`（0 页，专门核对空演示文稿）。

全部由 `scripts/New-GoldenFixtures.ps1` 通过 `DocumentFormat.OpenXml 3.1.0` 生成到 `tests/golden/.generated/`（已 gitignore），仓库不提交任何 `.docx/.xlsx/.pptx`。`manifest.json` 记录内容 SHA256；同一规格生成两次必须一致。

## 2. E11-02：语料静态门禁

`scripts/check-golden-corpus.ps1` 已登记进 `scripts/run-code-checks.ps1`，本机输出：

```text
PASS golden corpus: 26 cases (Word 9, Excel 9, PowerPoint 8), 24 ready, 2 gap, 7 negative, 9 fixtures, 28 mutation probes turned red
```

门禁覆盖：

- 数量、三端覆盖、负样例数、ready/gap 下限、nightly 子集存在；
- id 唯一与命名、layer 必须含 L1、ciTier/readiness/status/断言/trace 词表闭合；
- 入口按 `quick-action` / `ribbon` / `chat` 分别解析，Ribbon id 不得冒充快捷动作；
- `packageId` 必须能从包内解析到同一入口，`skill` 必须等于包声明的 Skill；
- Skill 存在且 `mustUseToolsAnyOf` 不越过 `allowed-tools`；
- 错误码必须存在于 `ExceptionClassifier.vb`；
- 负样例不得报 `completed`，非 `completed` 用例必须把 `maxWriteSteps` 钉在 0；
- 夹具生成两次并比对内容摘要，OpenXml 回读文本必须包含 `expectContains`，页数必须等于 `expectSlideCount`；
- 禁止客户数据、禁止整份文档快照、禁止提交 Office 二进制；
- `traceability` 的 objective 与计划行必须真实存在；
- gap 的 owner 路径、计划行、`missing` 与探针必须齐全，探针一旦可解析即失败。

变异探针 28 条，包含删用例、重复 id、未知断言/状态/错误码、幽灵入口/包/Skill、gap 探针可解析、包与入口不匹配、选区越界、夹具缺失/串宿主、负样例报成功、非完成态未封写、VBA 开启、步数预算为 0、未知 ciTier/trace 键、文档快照、断言串宿主、夹具回读缺文本、重复夹具 id、提交二进制等。任一条不红，门禁自己失败——防止「门禁绿着但已经瞎了」。

## 3. E11-03：运行记录与版本对比

`scripts/Compare-GoldenRuns.ps1`：

- 无参数运行合成记录自检，已作为 PR 门禁的一部分；
- 有 `-Baseline` / `-Candidate` 时校验记录 schema 并输出 markdown 对比；
- 记录字段：case、构建 commit、夹具哈希、result/status、步数、userRework、repairs、断言明细、runId、耗时、可选 attempts；
- 对比输出：通过率、无返工完成率及差值、新失败（pass→fail）、修复（fail→pass）、消失用例、新增用例、flaky、内部 repairs；
- **拒绝比较**：schemaVersion 不同、夹具集合或任一哈希不同、任一侧没有已执行用例（全 `not-run`）；
- `not-run` 记录合法，但永远不会被折算成 0% 或通过率。

自检结果：

```text
PASS golden run comparison: synthetic records classify new failures, fixes, disappeared, added, flaky; hash mismatch and not-run are refused
```

「无返工」只认用户可见返工：额外用户轮次 + 单步重试/跳过/改写 + 撤销本轮；内部 replan/repair 单独记 `repairs`，两个数字一起报。

## 4. E11-04：L1 执行手册

`tests/golden/README.md` 写明：

- L0/L1/L2/L3 分层与 D13 冻结（PR 不启动 Office、不联网；Nightly 只跑 `ciTier=nightly`；Release 跑全部 P0 + L2 签核）；
- 真机前置、每次临时副本隔离、单例 120 秒超时、网络错误只重试一次；
- runId 从 `agent_run` / `agent_run_step` 取，数据库路径为「我的文档\OfficeAiAppData(-Debug)\office_ai.db」；
- 如何数 `userRework`、如何写运行记录、如何跑对比；
- 无 Office 的 CI agent 不得跑 L1，也不得用 mock 冒充 L1；
- 三端真机验收清单（临时副本、选区、入口、runId、写观察/撤销、负样例阻断、失败导出与清理）。

## 5. 构建与门禁

```text
scripts\build-code-projects.ps1 -Configuration Debug
  PASS ShareRibbon / WordAi / ExcelAi / PowerPointAi

scripts\run-code-checks.ps1
  discovered 31 gates
  PASS check-golden-corpus.ps1
  PASS Compare-GoldenRuns.ps1
  PASS run-golden-l0.ps1 (19 cases)
  PASS check-task-package-contract.ps1
  ... 其余 27 个门禁全部通过
  All code checks passed.
```

新增的两个脚本均为纯 ASCII 字面量，PR 不依赖 Office、网络或 API Key。

## 6. 代价与边界

- **语料不是第六份清单**：入口、包、Skill、错误码全部从既有持有者反查；门禁不维护第二份名单。
- **夹具是生成物**：二进制不入库，摘要按内容算而不是按 zip 字节算，避免时间戳与关系 id 造成假漂移。
- **gap 是诚实的缺口**：现在跑不了就写 gap，并钉住「补上入口后必须升级」；不用自由文本请求假装覆盖。
- **没有真机就没有成绩**：E11-05 仍未执行，本轮不产生任何通过率；`not-run` 不可被对比脚本折算。
- **负样例的通过是正确阻断**：按预期返回 `needs_approval` / `blocked` / `needs_clarification` 并带指定错误码，不是打分打折。

## 7. 下一步

1. 环境具备后执行 E11-05：先跑 `ciTier=nightly` 子集，再跑 Release 全 P0，产出首个基线并标注 flaky。
2. 补 E6-07 的入口后，把两个 gap 用例升级为 `ready`（门禁会提醒）。
3. E12 起继续按计划推进；黄金集只增补真实任务，不为凑数新增工具或包。
