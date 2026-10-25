# 黄金评测 L1 执行手册（E11-04）

本文件说明如何在真实 Office 宿主上跑 [`tasks/`](./tasks/README.md) 里的语料、怎么留证据、以及 PR / Nightly / Release 三档各自该跑什么。

语料自身是否健康由 PR 门禁证明；**这份手册只解决「真机怎么跑、结果怎么记」**。没有真机时写 `not-run`，不允许补一份好看的成绩。

## 1. 分层与 D13 冻结

| 层 | 网络 | Office | 何时跑 | 证明什么 |
|---|---|---|---|---|
| **L0** 合同 + Mock | 否 | 否 | 每 PR | 门禁、状态机、包断言没坏 |
| **L1** 真宿主 | 是 | 是 | Nightly / Release | 真实任务能不能一次做成 |
| **L2** UX 手测 | 是 | 是 | Release | 时间线、审批卡、撤销按钮等 UI |
| **L3** LLM judge | 是 | 是 | 抽检 | 可选辅助，不作主分 |

D13 冻结：

| 流水线 | 跑什么 |
|---|---|
| **PR** | L0 + 现有 smoke + `build-code` + `scripts\check-golden-corpus.ps1`；不启动 Office、不联网 |
| **Nightly** | `ciTier = nightly` 的 L1 子集 |
| **Release** | 全部 `P0` 的 L1 + L2 手测签核 |

无 Office 的 CI agent **不得**跑 L1；也不得用 mock 结果冒充 L1。

## 2. 前置条件

- Windows + 已安装 Office（Word / Excel / PowerPoint 桌面版，能加载本项目 VSTO 插件）。
- 插件里已配置可用 API Key（经设置弹窗写入，不读环境变量、不写明文）。
- 仓库依赖已还原：`msbuild AiHelper.sln -t:Restore`。
- 四端代码可构建：`.\build-code.bat`。
- 当前分支与要记录的 `commit` 一致；工作区没有未提交的语料改动。

首次准备：

```powershell
powershell -File .\scripts\New-GoldenFixtures.ps1
```

生成物在 `tests\golden\.generated\`（已 gitignore），`manifest.json` 记录每个夹具的内容 SHA256。
每次正式跑之前重新生成一次；如果 manifest 与上次不同，说明语料变了，**旧记录不可直接对比**。

## 3. 一次 L1 运行

对 `tasks/catalog.json` 里筛选出的用例逐条执行：

1. **取临时副本**：把 `tests\golden\.generated\<fixture 文件>` 复制到临时目录，用临时副本打开宿主文档。禁止直接改夹具生成物。
2. **套用选区**：按 `setup.selection` 选中段落范围 / 表格 / 单元格区域 / 幻灯片 / 形状。
3. **从真实入口发起**：按 `entry` 使用命令面板条目、Ribbon 按钮或聊天框，把 `userText` 原样发给 Agent 主链。`entryActionId` / `ribbonId` 指向的入口必须真实可见；不可见就记 `fail` 并写明原因，不要改用别的入口代跑。
4. **固定选项**：`temperature = 0`、`allowVba = false`、`skillHint` 取自用例。
5. **超时**：单例 120 秒。超时记 `fail`，错误信息写 `timeout`。
6. **拿 runId**：run 结束后从 RunTrace 表取本轮 `run_id`。
   数据库路径：`我的文档\OfficeAiAppData\office_ai.db`（Debug 构建是 `OfficeAiAppData-Debug`）。
   表：`agent_run`（run 级状态）与 `agent_run_step`（step / tool / error / observation）。
7. **判定断言**：逐条对照 `expect.document.assertions` 与 `expect.trace`，把每条断言的 `passed` 写进记录。
8. **数用户返工** `userRework`：
   - 额外用户轮次（用户为完成同一任务又发了几轮）；
   - 单步重试 / 跳过 / 改写（`agent:stepAction`）；
   - 撤销本轮。
   内部 replan / repair 不计入 `userRework`，单独记 `repairs`。
9. **清理**：关闭文档、删除临时副本；失败也先导出证据再删。临时副本和导出文件都不入库。

网络错误允许重试一次，其余失败不重试。重试过的用例把每次结果写进 `attempts`；`attempts` 同时含 `pass` 与 `fail` 的用例由对比脚本标为 flaky。

## 4. 运行记录 schema

写到 `tests\golden\runs\<label>.json`（已 gitignore）。最小形状：

```json
{
  "schemaVersion": 1,
  "label": "nightly-2026-10-25",
  "commit": "75d44df",
  "generatedAt": "2026-10-25T12:00:00Z",
  "fixtures": {
    "word/report-zh": "<manifest.json 里的 contentSha256>"
  },
  "cases": [
    {
      "id": "W-proofread-selection",
      "result": "pass",
      "status": "completed",
      "steps": 4,
      "userRework": 0,
      "repairs": 1,
      "durationMs": 42000,
      "runId": "run_...",
      "assertions": [
        { "type": "text_not_contains", "passed": true }
      ],
      "attempts": ["pass"]
    }
  ]
}
```

规则：

- `result` 只能是 `pass` / `fail` / `not-run`；`status` 只能是 `completed` / `needs_approval` / `blocked` / `needs_clarification` / `failed` / `not-run`。
- `fixtures` 必须覆盖本次实际用到的全部夹具，值与 manifest 完全一致。
- `pass` / `fail` 必须带真实 `runId`；`not-run` 的 `status` 必须是 `not-run`，且步数、耗时、返工、修复、断言全部为空/零。
- 运行记录里不放截图、不放正文全文；只放断言的期望/实际与 runId。

## 5. 版本对比

```powershell
powershell -File .\scripts\Compare-GoldenRuns.ps1 `
  -Baseline .\tests\golden\runs\nightly-2026-10-20.json `
  -Candidate .\tests\golden\runs\nightly-2026-10-25.json `
  -OutFile .\tests\golden\runs\compare-2026-10-25.md
```

输出：通过率、无返工完成率及其差值，新失败（pass→fail）、修复（fail→pass）、消失用例、新增用例、flaky、内部 repairs。
以下情况脚本**拒绝比较**并以非零码退出：

- 两侧 `schemaVersion` 不同；
- 夹具集合或任一夹具哈希不同；
- 任一侧没有任何已执行用例（全是 `not-run`）。

无参数运行脚本会执行合成记录自检，这也是 PR 门禁的一部分。

## 6. 真机验收清单

每次 Release L1 至少签核下面几项；任意一项不过，不得宣布「黄金集通过」。

| 检查 | Word | Excel | PowerPoint |
|---|---|---|---|
| 临时副本能正常打开，原始夹具未被修改 | ☐ | ☐ | ☐ |
| `setup.selection` 能精确套用（段落 / 表格 / 区域 / 页 / 形状） | ☐ | ☐ | ☐ |
| 用例声明的入口真实可见、点击后进入同一条 Agent 主链 | ☐ | ☐ | ☐ |
| run 结束后能从 `agent_run` 取到 runId，step 序列可查 | ☐ | ☐ | ☐ |
| 写步骤都有 observation，撤销本轮能恢复指纹（若用例要求） | ☐ | ☐ | ☐ |
| 负样例按预期阻断，并带用例指定的错误码 | ☐ | ☐ | ☐ |
| 失败已导出 run、临时副本已清理 | ☐ | ☐ | ☐ |

## 7. 当前状态

- `tests/golden/tasks/` 语料与 `scripts/check-golden-corpus.ps1` 已接入 PR 门禁。
- `scripts/Compare-GoldenRuns.ps1` 已接入合成记录自检，可对真实记录做对比。
- **L1 尚未跑**：按当前指示延后（Office + API Key 环境具备后执行 E11-05）。在那之前，任何「通过率」都只能标 `not-run`，不能作为发布依据。
