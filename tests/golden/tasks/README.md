# 黄金评测语料 v1（E11）

这一份目录是「真实文档任务」的唯一清单。目标不是多跑几个 prompt，而是让每次版本回归都能回答同一个问题：

> 对同一批真实任务，这一版是否仍然无返工完成？失败是否还落在同一批用例上？

设计口径见 [`docs/design/eval-golden-set.md`](../../../docs/design/eval-golden-set.md)；本文件只说明**怎么读写这份语料**。

## 1. 目录

```text
tests/golden/tasks/
├── README.md                 # 本文件
├── catalog.json              # 任务清单（唯一持有者）
├── assertions.v1.json        # 断言词表（闭合集合）
├── fixtures/
│   ├── word/*.json           # Word 夹具规格
│   ├── excel/*.json          # Excel 夹具规格
│   └── ppt/*.json            # PowerPoint 夹具规格
└── expect/                   # 禁止存在（整份文档快照）
```

生成物固定写到 `tests/golden/.generated/`（已 gitignore），由
`scripts/New-GoldenFixtures.ps1` 从 JSON 规格生成，不提交任何 `.docx/.xlsx/.pptx`。

## 2. 一份 case 的字段

| 字段 | 必填 | 说明 |
|---|---|---|
| `id` | 是 | 全局唯一，`^[A-Z][A-Za-z0-9-]+$`；负样例用 `-negative` 结尾，gap 用 `-gap` 结尾 |
| `version` | 是 | 断言口径变化时递增 |
| `appType` | 是 | `Word` / `Excel` / `PowerPoint` |
| `layer` | 是 | 必须含 `L1`：这份语料量的是真宿主结果 |
| `ciTier` | 是 | `nightly` 或 `release`；`nightly` 是 D13 冻结的夜间子集 |
| `readiness` | 是 | `ready`（入口/包/Skill 都能解析）或 `gap`（缺持有者，必须写 `gap` 块） |
| `title` / `userText` | 是 | `userText` 用真实中文用户口吻，不写成工具说明 |
| `fixtureId` | 是 | 指向 `fixtures/**` 里的 `id` |
| `setup.selection` | 是 | `whole-document` / `paragraph_range` / `table` / `whole-workbook` / `range` / `whole-deck` / `slide` / `shape` |
| `entry` | 是 | `quick-action` / `ribbon` / `chat` |
| `entryActionId` / `ribbonId` | 按需 | 与 `entry` 一一对应；id 必须真实存在于 `QuickActionCatalog.vb` |
| `packageId` | 否 | 有任务包合同时才填；必须能在包内解析到同一入口，且 `skill` 必须等于包声明的 Skill |
| `skill` | 否 | 必须存在于 `ShareRibbon/Skills/*/SKILL.md` |
| `harnessOptions` | 是 | `temperature` 必须为 `0`，`allowVba` 必须为 `false` |
| `expect.status` | 是 | `completed` / `needs_approval` / `blocked` / `needs_clarification` / `failed` |
| `expect.maxSteps` / `maxRepairs` | 是 | 步数 `1..20`，修复 `0..4` |
| `expect.mustUseToolsAnyOf` | 否 | 数组的数组，任一组满足即可；工具必须落在该 case 的 Skill `allowed-tools` 内 |
| `expect.mustNotUseTools` | 否 | 例如所有用例都不该出现的 `ExecuteVBA` |
| `expect.requiredErrorCodes` | 否 | 只能填 `ExceptionClassifier.vb` 里存在的错误码 |
| `expect.document.assertions` | 是 | 见 `assertions.v1.json`；类型、参数、适用宿主都由词表闭合 |
| `expect.trace` | 是 | `min/maxWriteSteps`、`everyWriteHasObservation`、`everyWriteUndoable`、`skillPrimaryIn` |
| `tags` | 是 | 含 `negative` 的用例计入安全负样例；`P0` 是发布必过 |
| `owner` | 是 | 维护者 |
| `traceability` | 是 | `objective`（`O1/O2/O3`）+ `planRow`（必须能在执行计划里找到） |
| `gap` | `readiness=gap` 时必填 | `owner` 路径、`planRow`、`missing`、`probeEntryId` / `probePackageId` |

`gap` 的含义是**诚实地暴露缺口**：用例照常写话术和断言，但不假装现在就能跑。
一旦 `probeEntryId` 或 `probePackageId` 变成可解析，门禁会立刻变红，逼着把用例升级为 `ready` 而不是静默保留一个永远跑不了的样例。

## 3. 夹具规格

每个 `fixtures/**/*.json` 至少包含：

```json
{
  "id": "word/report-zh",
  "kind": "word",
  "appType": "Word",
  "dataPolicy": "fictional",
  "blocks": [ ... ],
  "expectContains": ["文中真实存在的片段"],
  "expectSlideCount": 0
}
```

- `kind`：`word` / `excel` / `ppt`，决定生成器和回读模型。
- `blocks` / `sheets` / `slides`：最小内容，只保留任务判断需要的信息。
- `expectContains`：生成后必须能在 OpenXml 回读文本里找到的片段；公式按 `=B2-C2` 形式核对。
- `expectSlideCount`：PowerPoint 用来钉住页数（空演示文稿必须是 0）。
- `dataPolicy` 必须是 `fictional`；语料里不允许出现真实客户、真实人名或真实金额。

夹具的同一性用**内容摘要**（`manifest.json` 的 `contentSha256`）而不是 zip 字节：
同一份规格生成两次必须得到同一个摘要，据此可以发现「夹具悄悄变了」。
门禁会连续生成两次到临时目录并逐条比对，同时把回读模型和 `expectContains` 对照一遍。

## 4. 断言词表

新增断言只能改 [`assertions.v1.json`](./assertions.v1.json)：

- 每个断言必须声明 `apps`、`params`、`required`；
- case 里出现词表外的 `type` 或参数，门禁直接红；
- 断言必须能由 RunTrace 或文档回读机器判定，不引入模型打分；
- 不写「整份文档快照」类断言（`snapshot` / `expectedFile` / `outputFile` 等键被禁止）。

## 5. 无返工完成率

评测主指标沿用执行计划中的定义：

```text
无返工完成率 = completed 且 userRework = 0 的用例数 / 应跑用例数
userWork = 额外用户轮次 + 单步重试/跳过/改写 + 撤销本轮
```

内部 `replan/repair` 单独计入 `repairs`，不混入用户返工。
因此 `trace` 里的写步数、观察和撤销断言不是装饰：它们决定这一轮能不能算「无返工」。
安全负样例要求的是**正确阻断**（`needs_approval` / `blocked` / `needs_clarification` + 指定错误码），不是 `completed`。

## 6. 新增一个 case

1. 先确认它对应真实用户任务，而不是为一个工具编故事。
2. 选最近的夹具；需要新夹具就新增 JSON 规格，保持最小、虚构。
3. 写 `userText`，口吻要像用户，不要把工具名塞进去。
4. 选真实入口 `entryActionId` / `ribbonId`；找不到入口就不要硬填，改为 `readiness=gap` 并写清持有者。
5. 只写能机器判定的断言；先跑一遍 L1，把稳定观察到的结果写进去。
6. 运行 `powershell -File scripts/check-golden-corpus.ps1`，必须变绿。
7. 负样例必须落在真实的拦截路径上（Skill 门禁 / VBA 关闭 / 高危删除 / 超大范围 / 含糊指令），并带 `negative` 标签。

## 7. 运行边界（D13）

| 流水线 | 跑什么 |
|---|---|
| PR | L0 合同 + 现有 smoke + `build-code` + `check-golden-corpus.ps1`（不启动 Office、不联网） |
| Nightly | `ciTier=nightly` 的 L1 子集（真 Office + API Key） |
| Release | 全部 P0 的 L1 + L2 手测签核 |

PR 门禁**只做静态与生成/回读校验**：它证明语料自身没有腐烂，不证明模型能跑对。
L1 执行、运行记录与基线对比由 E11-03 / E11-04 接入；在真宿主和 Key 就绪前，不伪造任何运行结果。
