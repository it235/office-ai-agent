# 产品 AI Native 执行计划（目标锁定版）

> **配套文档**：[product-ai-native-strategy-report.md](./product-ai-native-strategy-report.md)  
> **用途**：把战略报告拆成可执行、可验收、可追溯的工作项，防止执行过程中偏离原始目标。  
> **规则**：任何工作项必须映射到 O1 / O2 / O3 之一；映射不上，就不做。

---

## 0. 使用规则

1. 本文件是执行期间的**唯一 backlog 来源**。聊天里临时想到的功能，不直接进入开发。
2. 每周从“周计划”取 1–2 个工作项，完成后回填验收证据与指标。
3. 新需求先过“工作项准入清单”（第 8 节）；不能回答的问题，视为尚未想清楚。
4. 目标只能由用户明确确认后变更，并写入“决策日志”。
5. 不以工具数量、按钮数量、代码行数、编译通过作为产品进展。

---

## 1. 目标锚点（不可漂移）

### 1.1 一句话目标

> 让用户在 Word / Excel / PowerPoint 中，用一句话把真实文档任务交给 AI，并且**敢用、能退、越用越懂**。

### 1.2 三个用户结果（Outcome）

| ID | 用户结果 | 用户原话 | 反面案例 |
|---|---|---|---|
| **O1** | 一次做成 | “我说一遍，它就做对了。” | 反复改 prompt、做完还要手工修 |
| **O2** | 敢让它做 | “改错了我能退，知道它改了什么。” | 黑盒修改、取消后留半成品、假成功 |
| **O3** | 越用越懂 | “它记得我的文档、模板和习惯。” | 每次都重新解释背景与格式 |

### 1.3 北极星与护栏

| 类型 | 指标 | 3 个月目标 |
|---|---|---|
| 北极星 | 无返工任务完成率 | ≥ 60% |
| 激活 | 首次价值时间 | ≤ 60 秒 |
| 信任 | 用户撤销率 | ≤ 15% |
| 质量 | 平均修复次数 | ≤ 1.0 |
| 安全 | 未授权写入 / 假成功 | 0 |
| 留存 | 7 日任务留存 | 基线 +20% |

### 1.4 非目标（明确不做）

- 不继续增加工具数量，长尾能力走 `DiscoverOfficeCapability + OfficeObjectOperation`。
- 不新建第二套 Harness / Loop / Safety / 执行入口。
- 不追 Copilot 365 的全量企业能力。
- 不做无审批的删除、整表替换、VBA、取消保护。
- 不把所有交互都塞进聊天框。
- 不做“看起来智能”的假成功。

### 1.5 反跑偏规则

| 红旗 | 处理 |
|---|---|
| “这个功能挺酷，先加上” | 先问对应 O1/O2/O3 哪一个；答不上来进暂缓池 |
| “再加一个专用 Tool 更快” | 先评估现有任务包 + 动态对象操作能否承载 |
| “先做 UI，后端以后再接” | 不允许；UI 壳会制造假能力 |
| “编译过了就算完成” | 不算；必须在真实 Office 跑通验收脚本 |
| “模型能生成就说明能做” | 不算；必须有观察、验证、失败恢复 |
| “用户没投诉就先不动” | 不算；信任缺口本身就是投诉前风险 |

---

## 2. 追溯矩阵（战略 → 执行）

| Outcome | 任务包 | Epic | 领先指标 | 验收证据 |
|---|---|---|---|---|
| O2 | 全部写操作 | E1 信任层 | 撤销成功率、取消回滚率、假成功数 | 真机取消/撤销脚本 + 变更日志 |
| O3 | 全部任务包 | E2 会话与文档绑定 | 跨端串会话数、重复说明次数 | 三端历史隔离截图 + DB 查询 |
| O1 | 首次任务 | E3 首次 60 秒 | 首次价值时间、激活率 | 新用户录屏 + 埋点 |
| O1/O2 | 全部任务包 | E4 Office 原生入口 | 入口使用率、从选中到出结果时长 | 真机操作脚本 |
| O1 | 全部任务包 | E5 能力一致性 | 规划必失败率、TOOL_NOT_ALLOWED 数 | 一致性校验脚本 + 真机回归 |
| O1 | 12 个任务包 | E6 任务包框架 | 任务包完成率 | 每个任务包的黄金样例 |
| O1 | Excel 清洗/报表 | E7 Excel 大表安全 | 大表任务成功率、UI 冻结次数 | 10 万行真机脚本 |
| O2 | Excel 全部写操作 | E8 Excel 预览与回滚 | 预览准确率、回滚成功率 | old/new 对照 + 回滚脚本 |
| O1/O2 | PPT 成稿/美化 | E9 PPT 品牌与真实对象 | 套模板成功率、原生图表数 | 母版/图表真机验收 |
| O1 | 复杂任务 | E10 计划编辑与恢复 | 单步重试成功率、重跑步骤数 | 失败注入测试 |
| O1 | 全部任务包 | E11 黄金评测集 | 无返工完成率 | ≥20 个真实文档任务 |
| O3 | 全部任务包 | E12 记忆与上下文 | 重复说明次数、记忆命中率 | 同文档二次任务对比 |
| O2 | Agent 任务 | E13 运行时间线 | 任务可解释率、失败定位时长 | Trace UI + 回放截图 |
| O1 | 全部任务包 | E14 反馈与度量 | 反馈覆盖率、问题闭环率 | 看板 + 反馈标签 |
| O1 | 全部任务包 | E15 成本与模型路由 | 单任务成本、超预算率 | 费用预估对照 |
| O1 | 生态扩展 | E16 MCP / 模板市场 | 连接器使用率、模板复用率 | 试点案例 |
| O1/O2 | 长文档 / 大表 / 多目标任务 | E17 只读分诊与上下文隔离 | 主上下文峰值 Token、因截断返工次数、单任务成本 | 长文档对照跑 + Trace 子任务截图 |

---

## 3. 任务包规格模板（新能力必须填写）

任何新任务包或能力升级，先填完这张表，再进入开发：

| 字段 | 要求 |
|---|---|
| 任务包名称 | 用户能理解的工作名，不是技术名 |
| 触发入口 | 选区工具条 / 右键 / 快捷键 / 聊天 / Ribbon |
| 适用宿主 | Word / Excel / PowerPoint，是否三端一致 |
| 用户一句话 | 真实用户会怎么说 |
| 上下文读取 | 必须读取哪些文档结构、选区、样式、历史 |
| Skill | 使用哪个 Skill，是否需要新增 references |
| allowed-tools | 最小工具集；结构操作是否包含 |
| 风险级别 | safe / medium / risky / forbidden |
| 预览策略 | 不预览 / 摘要预览 / 文档内高亮 / 缩略图 |
| 验证器 | 如何判断“真的做成了” |
| 交付物 | 文档内结果 / 新 Sheet / 新 Deck / 报告 |
| 失败恢复 | 修复、换工具、回滚、降级说明 |
| 解释模板 | 用户看到的“改了什么、为什么、能不能退” |
| 成功标准 | 可观察、可测量的验收条件 |
| 黄金样例 | 至少 1 个成功样例 + 1 个失败样例 |
| 已有能力 | 仓库里已经能做什么、在哪个类；本任务只允许补“确实没有”的那部分，已有实现必须复用或跳转，不得在面板里再做一套（E3-01 首版就是栽在这里） |

---

## 4. Phase 1 详细拆解（第 1–4 周）：可信与一致

**Phase 目标**：让用户敢用、找得到、不串端。

### E1 信任层（对应 O2）

**Epic 目标**：任何写操作都能回答三个问题：改了什么、能不能退、为什么这么改。

| WI | 工作项 | 验收标准（用户可观察） | 证据 | 依赖 | 规模 |
|---|---|---|---|---|---|
| E1-01 | 写操作信任合同 | 所有写路径声明 preview/observe/undo/error 语义；缺失项被检查脚本列出 | 合同文档 + 检查脚本输出 | 无 | S |
| E1-02 | 单轮变更日志 | 一次任务结束后，能列出本轮所有变更：目标、前后摘要、是否成功 | 日志/DB 行 + 样例 | E1-01 | M |
| E1-03 | 一键撤销本轮（Word 优先） | 点击“撤销本轮”后，文档 hash/段落数回到任务前；0 残留；失败时明确报错 | 真机脚本 + 前后快照 | E1-02 | L |
| E1-04 | 取消与回滚语义 | 取消后显示“已写入 N 处，未完成 M 步”，提供撤销/保留；不静默留半成品 | 真机取消录屏 + 变更日志 | E1-02 | M |
| E1-05 | 审批粒度 | 支持允许一次 / 本会话允许该工具 / 拒绝；拒绝只影响当前步骤 | 真机审批脚本 | E1-02 | M |
| E1-06 | replan 幂等保护 | 注入第 3 步失败后 replan，第 1–2 步不重复执行；无重复写入 | 失败注入测试 + 日志 | E1-02 | M |
| E1-07 | Excel/PPT 接入撤销 | 两端取消/失败后可撤销本轮；行为与 Word 一致 | 三端对照脚本 | E1-03/04 | L |
| E1-08 | 信任层 UI | 计划卡显示范围/数量/风险；结束后显示变更摘要 + 撤销按钮；时间线状态正确 | UI 截图 + 交互脚本 | E1-03/04 | M |

**不在本期范围**：跨文档撤销、历史版本管理、协作审阅合并。

### E2 会话与文档绑定（对应 O3）

**Epic 目标**：历史找得到、不串端、可清理。

| WI | 工作项 | 验收标准 | 证据 | 依赖 | 规模 |
|---|---|---|---|---|---|
| E2-01 | 会话按宿主过滤 | Word 历史只显示 Word；Excel/PPT 同理 | 三端截图 + SQL 查询 | 无 | S |
| E2-02 | 会话绑定文档 | 打开某文档时，优先显示该文档会话；无绑定时明确显示“通用会话” | 真机脚本 | E2-01 | M |
| E2-03 | 自动标题 + 重命名 + 删除 | 首轮后生成可读标题；重命名/删除即时生效 | UI 脚本 + DB 校验 | E2-01 | M |
| E2-04 | 旧数据迁移/回填 | 老会话不丢失；无 app_type 的记录有明确归属策略 | 迁移脚本 + 升级测试 | E2-01 | M |
| E2-05 | 历史 UI 分组 | 按文档/时间分组；空状态有引导 | UI 截图 | E2-02/03 | S |

### E3 首次 60 秒（对应 O1）

**Epic 目标**：新用户从安装到完成第一个任务 ≤ 60 秒。

| WI | 工作项 | 验收标准 | 证据 | 依赖 | 规模 |
|---|---|---|---|---|---|
| E3-01 | 引导向导 | 隐私确认 → 示例任务；选服务商/填 Key/测连接一律复用既有配置窗口 | 新用户录屏 | 无 | M |
| E3-02 | 每端一个示例任务 | Word/Excel/PPT 各一个一键示例，使用当前文档或内置样例 | 三端录屏 | E3-01 | M |
| E3-03 | 空 Key 阻断 | 未配置时不能发送；给出明确 CTA，而不是状态栏报错 | 真机脚本 | E3-01 | S |
| E3-04 | 费用/额度提示 | 发送前显示预估或发送后显示实际 Token/费用；超预算有提示 | UI 截图 + 样例 | E3-01 | M |
| E3-05 | 隐私同意 UI | 首次云调用前记录同意；可查看、可撤回；文案说明数据边界 | UI 截图 + 配置记录 | E3-01 | S |

### E4 Office 原生入口（对应 O1/O2）

**Epic 目标**：AI 在用户工作的地方出现，而不是要求用户先打开面板。

| WI | 工作项 | 验收标准 | 证据 | 依赖 | 规模 |
|---|---|---|---|---|---|
| E4-01 | 选区工具条（Word 优先） | 选中文本后 300ms 内出现 3 个动作；点击后进入预览/执行 | 真机录屏 | E1-01 | M |
| E4-02 | 右键菜单（三端） | 右键出现“AI 处理选中内容”等 2–3 项；无选中时有合理降级 | 三端截图 | E4-01 | M |
| E4-03 | 全局快捷键 / 命令面板 | 任一宿主内可唤起；支持自然语言与命令 | 真机录屏 | E4-01 | M |
| E4-04 | 选区自动引用默认开启 | 默认携带选区上下文；UI 明确显示引用了什么 | 真机脚本 | E4-01 | S |
| E4-05 | 事件绑定卫生 | 不重复绑定；重复打开/关闭不产生双触发 | 回归脚本 | E4-01/02/03 | S |

### E5 能力一致性（对应 O1）

**Epic 目标**：不出现“说得到做不到”，常见结构操作不被自己的边界挡住。

| WI | 工作项 | 验收标准 | 证据 | 依赖 | 规模 |
|---|---|---|---|---|---|
| E5-01 | 单一能力清单 | Skill allowed-tools、Tools JSON、宿主白名单、Prompt、Ribbon 由同一清单校验 | 校验脚本 + 失败样例 | 无 | M |
| E5-02 | 基线 Skill 补结构操作 | Excel 可新建/删除 Sheet、插入行列；PPT 可删页/复制页/插形状；Word 常用结构操作可用 | 三端真机脚本 | E5-01 | M |
| E5-03 | 假成功清理 | `ApplyTheme` 等只弹提示的操作不再返回成功；统一返回 `UNSUPPORTED` 或真实执行 | 工具合同测试 | E5-01 | S |
| E5-04 | 不可达入口处理 | 模板库等隐藏入口要么接通，要么移除；不留“只有开发者知道”的能力 | 入口清单 + 真机检查 | E5-01 | S |
| E5-05 | CI 防漂移 | 四份能力清单不一致时构建失败；新增能力必须更新清单 | CI 输出 | E5-01 | S |

### Phase 1 周的排布建议

| 周 | 重点 | 交付 |
|---|---|---|
| W1 | E5-01/02/03 + E1-01/02 + E2-01 | 能力清单校验、变更日志、会话过滤 |
| W2 | E1-03/04/05 + E2-02/03 + E3-01/03 | Word 撤销本轮、取消语义、审批粒度、引导向导 |
| W3 | E1-06/08 + E4-01/02 + E3-02 | replan 幂等、信任 UI、选区工具条、示例任务 |
| W4 | E1-07 + E4-03/04/05 + E2-04/05 + E5-04/05 | Excel/PPT 撤销、快捷键、迁移、M1 闸门 |

---

## 5. Phase 2 / Phase 3（Epic 级）

### Phase 2（第 5–8 周）：任务闭环

| Epic | 目标 | 关键验收 | 退出条件 |
|---|---|---|---|
| E6 任务包框架 + 前 6 个任务包 | 用户知道能做什么，且能一次做成 | 每个任务包有入口、示例、预览、验证器、黄金样例 | 6 个任务包真机无返工完成率 ≥ 60% |
| E7 Excel 大表安全执行 | 十万行不假死 | 分块、进度、取消、UI 不冻结 | 10 万行脚本通过 |
| E8 Excel 真实预览与回滚 | 改前看得见，改坏能退 | 真实 old/new；失败/取消可回滚 | 预览准确率 ≥ 95%，回滚成功 |
| E9 PPT 品牌模板与真实对象 | 能套公司模板，图表可编辑 | 母版/版式复用；原生 Chart/Table；写入前预览 | 至少 1 套模板 + 1 个原生图表验收 |
| E10 计划编辑与单步恢复 | 用户能纠偏，不必重来 | 改步骤只重跑受影响步骤；失败步骤可重试 | 失败注入测试通过 |
| E11 黄金评测集 v1 | 可持续变聪明 | ≥20 个真实文档任务，含失败样例 | 每次版本回归可比较 |

#### E6 详细拆解（落地顺序即编号顺序）

E6 的坑是「再建一套与 QuickActionCatalog / Skill / ToolRegistry / 宿主 capability 平行的清单」。因此这一 Epic 的硬约束是：**任务包只允许承载目前没有任何持有者的字段**（预览策略、验证器、成功标准、交付物、解释模板、黄金样例引用），入口、示例、Skill、工具一律引用既有 id，由门禁逐级验证引用能解析——引用断裂必须构建失败，否则这张表三周后就是第五份腐烂清单（Phase 1 的教训见 §9 第 20/22 条）。

| WI | 工作项 | 验收标准（用户可观察） | 证据 | 依赖 | 规模 |
|---|---|---|---|---|---|
| E6-02 | 任务包合同数据化 | 一个任务包一份 JSON，字段填齐 §3 模板；包能被加载器读出来并拒绝未知字段/未知断言词，而不是靠人记 | `ShareRibbon/TaskPackages/*.json` + `TaskPackageCatalog` 加载输出 + 门禁 | 无 | M |
| E6-03 | 任务包一致性门禁 | 包→动作→capability→Skill→工具→黄金样例逐级可解析；没有入口的包不许声明；任一级断裂即失败 | `scripts/check-task-package-contract.ps1` 输出 + 变异探针 | E6-02 | S |
| E6-04 | 面板按任务包组织入口 | 命令面板里动作按任务包归组，用户看到的是工作目标而不是动词列表；零新消息 type、零新发送路径 | 面板真机 + `check-entry-reachability.js` | E6-02 | S |
| E6-01 | 入口话术收口到唯一清单 | 三端 Ribbon 里硬编码的 Agent 请求全部改为从 `QuickActionCatalog` 取（编译期引用，不做字符串查表），入口不再持有第二份话术 | 目录 diff + E6-03 门禁 | E6-02/03 | M |
| E6-05 | 验证器落成可执行断言 | 包声明的成功标准在 run 结束时被真的核对（复用 `ValidateObservedOutcome` 与 `agent_run_change` 事实），失败进入既有 repair loop 而不是直接宣称成功 | smoke + 真机 | E6-03 | L |
| E6-06 | 前 6 包黄金样例 | 每包 ≥1 成功 + ≥1 失败样例，`scripts/run-golden-l0.ps1` 跑得动 | `tests/golden/` + 门禁 | E6-03 | M |
| E6-07 | 补齐没有入口的包 | 报告 §3.2 前 6 包（Word 审阅润色/模板排版、Excel 清洗/洞察、PPT 成稿/品牌模板）全部有可解析入口；确实没有的必须显式记为未声明而不是硬编一个 | 门禁 + D9 | E6-01 | M |
| E6-08 | 六包真机完成率 | 每包 3 个真实文档样例，无返工完成率 ≥60%（M2 通过线） | 真机记录 | E6-05/06 | L |

#### E10 详细拆解（落地顺序即编号顺序）

E10 的坑是**为了「用户能纠偏」再造一套循环、一套终态或第二条裁决通道**。动手前先量现状（全部 `git show HEAD:` + 实读，不是印象）：

1. 步骤失败后用户没有任何单步手段。三处失败出口（解析失败 `LoopEngine.vb:150-158`、归一化失败 `:162-169`、ACT 失败 `:410-483`）都只改状态并前进索引；只有 `noProgressCount` 达 3 才让**模型**在 `retry/skip/replan` 三分支里替用户决定（`:444-477`），而且每次都要花掉一份 `replanAttempts` 预算。
2. 用户此刻能按的钮只有三个：`confirmAgentExecution`（简单任务 500ms 后自动点掉，`agent-card.js:199-201`）、`refineAgentPlan`、`abortAgent`。其中「修改计划」**不是改计划**：`HandleAgentRefinePlan` 把用户意见拼进原始请求后重新调 `HandleStartAgentCore`（`BaseChatControl.vb:2140-2160`），等于整轮重来——与验收标准「不必重来」正相反。
3. 面板↔宿主唯一的真人裁决通道是 harness 审批握手（`HandleKernelRequestApproval` / `SettleApprovalAsync` / `DecideAsync`，`OfficeHarness.Approvals.vb:12-102`），它带超时、放弃与 trace 语义，`DecideAsync`  settle 完还要 `Await pending.AgentTask`。它服务的是「循环停在安全点等人」，而单步纠偏发生在循环**并未**等人时，所以不能把它抢来当第二个用途（那会把 Deny/Once/Session/Abort 这套 E1-05 词表污染成两种含义）。
4. 「只重跑受影响步骤」这一半**已经有持有者**：E1-06 的写入台账按参数指纹只挡更早计划版本的写入（`AgentWriteLedger.vb:87-94`，同版重复是用户明确要求的语义），`retry` 分支的形状就是「状态复位 Pending + 索引不前移」（`:446-452`）。所以 E10 缺的是**用户的入口和一份收件箱**，不是新机制。

因此硬约束：**用户纠偏只允许在既有失败分支被消费**（循环自己拉取，UI 不往执行中段推）；`retry/skip` 词汇沿用反思分支那一份，只是来源从模型换成用户；终态出口仍只有 `AgentResult.SuccessResult` 一处；计划版本推进者仍只有 `Writes.BeginNextPlan()`；审批握手不为纠偏复用；一旦计划被整体替换（replan 或 run 末修复），旧索引上的用户决定必须作废——收件箱按「步号」寻址，新计划里同号不是同号。

| WI | 工作项 | 验收标准（用户可观察） | 证据 | 依赖 | 规模 |
|---|---|---|---|---|---|
| E10-01 | 失败步骤的用户纠偏 | 某一步失败后该步出现「重试这一步 / 跳过这一步」，点重试只重跑这一步（已成功步骤 0 次重跑）、不消耗重规划预算；点跳过该步标灰并继续后续步骤；无人点击时行为与今天逐条相同 | `scripts/smoke-step-intervention.ps1` 失败注入 + `check-agent-card-ui.js` | E1-06/E1-08 | M |
| E10-02 | 单步正文与参数编辑 | 用户改这一步的说明/参数后，卡片仍显示同一版计划（不出现「第 2 版」），只有这一步重跑；参数指纹变了所以台账不会把它当重复写入挡掉 | 门禁 + 真机 | E10-01 | M |
| E10-03 | 终态后的边界说明 | 本轮已结束时单步动作全部禁用并说明原因（此时「重试」等于重开一轮，就是「重来」），可用出口是撤销本轮或带意见重新发起 | 卡片真机 + UI 断言 | E1-07/E10-01 | S |
| E10-04 | 门禁与口径落地 | 失败注入门禁自动登记进 `run-code-checks.ps1`；P0 守卫钉住「三处失败分支都消费收件箱」「计划替换点必须清空」「不得新增第二套终态/裁决通道」；AGENTS 与落地记录写清边界与代价 | 门禁输出 + 变异探针 | E10-01 | S |

### Phase 3（第 9–12 周）：智能与生态

| Epic | 目标 | 关键验收 | 退出条件 |
|---|---|---|---|
| E12 记忆与上下文深度 | 越用越懂 | 文档级记忆、偏好记忆、上下文结构快照 | 同文档二次任务重复说明减少 ≥30% |
| E13 运行时间线与回放 | 看得懂、可复盘 | RunTrace UI 可查看；失败定位到步骤/错误码 | 失败定位时长下降 50% |
| E14 反馈与度量 | 知道哪里不好 | 接受/拒绝/原因标签；指标看板 | 反馈覆盖率 ≥ 30% |
| E15 成本与模型路由 | 花得明白 | 任务前预估；按复杂度选模型 | 单任务成本可解释 |
| E16 MCP / 模板市场试点 | 可扩展 | 1 个连接器 + 1 套可分享模板 | 真实试点案例 |
| E17 只读分诊与上下文隔离 | 长文档/大表任务不因为"读太多"而变笨 | 只读子任务可并行、写仍串行；主上下文峰值下降；子任务在时间线可见 | 见 5.1，长文档对照任务通过 |

### 5.1 E17 只读分诊与上下文隔离（Sub-task fan-out）

**为什么现在补**：Qoder / Trae / Claude Code / Codex 拆子任务的第一动机不是"看起来更智能"，而是**防止上下文污染**——把长检索、长阅读、大范围扫描的原始产物留在子任务里，主循环只接收压缩后的结构化结论。本插件在长文档校对、大表分析、多目标任务（PLN-7「校对并排版」）上正踩同一个坑：全文塞进主 Prompt → 截断 → 计划丢细节 → 返工。所以值得做，但形态必须受限。

**产品判断（结论）**：做"只读分诊"，不做通用多 Agent 框架。收益来自上下文隔离，不来自并行数量；写操作一律不并行。

| WI | 目标 | 验收 | 依赖 | 规模 |
|---|---|---|---|---|
| E17-01 | 只读分派步骤：LoopEngine 支持把**只读** Tool 按范围分片并发取回，结果只以结构化 Observation 汇总回主循环 | 分片读取上下文字节数下降 ≥50%，主 Prompt 不含原文片段；golden 样例通过 | E1-06 | M |
| E17-02 | 子任务上下文预算与压缩合同：每片有明确输入范围、输出 schema、最大 Token；超出即截断报错，不静默丢内容 | 契约测试 + 审计守卫（禁止子任务返回原文） | E17-01 | M |
| E17-03 | 时间线可见：计划卡/RunTrace 显示"分诊 N 片 → 汇总 → 执行"，失败定位到具体片 | Trace UI 截图；用户能看懂为什么慢 | E13、E1-08 | M |
| E17-04 | 成本闸门：分诊前预估额外 Token/耗时，超预算自动降级为单遍顺序读取 | 单任务成本可解释；无预算外暴涨 | E15 | M |

**硬约束（不可谈判）**：

1. 唯一主链不变：`OfficeHarness -> AgentKernel -> LoopEngine -> ToolRegistry -> CodeExecutionService -> 宿主`。禁止为分诊另建 Planner/Harness/Safety/终态；子任务只是 LoopEngine 内部的一种步骤形态，共用同一份 session 状态机和错误码合同。
2. **只读并行、写串行**（对齐 `docs/ai-native-harness-design.md` LOP-4）。Office COM 是单线程套间，宿主 UI 线程是唯一串行资源；子任务不得触碰 `ExecuteJsonCommand`、不得申请撤销点。
3. 子任务不得独立向用户弹审批。只读分片按低风险自动放行；一旦分片里混入写操作，必须回到主循环走 E1-05 三态审批。否则一次任务弹 N 个审批框，是体验灾难。
4. 截图/Base64 等视觉证据仍按现规则禁止进入 Observation/History/Trace；子任务回传只允许文本结构化数据。
5. 子任务失败必须能单独降级（该片标记缺失、主循环继续或 replan），不能因一片失败让整轮状态含糊。

**不做什么**：不做云端多 Agent 编排、不做 Agent 之间自由对话、不做"用户可自己创建子 Agent"。这三项在 M3 之前不进入开发。

**排期**：E17-01/02 提前到第 7–8 周做（依赖 E1-06 完成即可启动，收益直接体现在长文档任务完成率上）；E17-03/04 随 E13/E15 落在 Phase 3。

---

## 6. 里程碑闸门

### M1（第 4 周末）：可信与一致

**演示脚本**：

1. 新用户安装 → 向导配置 Key → 完成 Word 示例任务（≤60 秒）。
2. 发起一个 5 步 Word 任务，中途取消 → 看到“已写入 N 处” → 一键撤销 → 文档恢复。
3. 故意触发高风险工具 → 拒绝当前步骤 → 任务给出部分完成摘要，不静默终止。
4. Excel 打开历史 → 只看到 Excel 会话；Word 同理。
5. 抽查 10 个可见入口 → 全部有真实能力或已移除。

**通过线**：假成功 0；未授权写入 0；撤销脚本 100% 通过；跨端串会话 0。

### M2（第 8 周末）：任务闭环

**演示脚本**：

1. 6 个任务包各跑 3 个真实文档样例。
2. Excel 10 万行清洗 → 有进度、可取消、不冻结。
3. PPT 套用企业模板 → 图表为原生对象 → 写入前有缩略图预览。

**通过线**：6 个任务包无返工完成率 ≥60%；Excel 大表任务成功率 ≥90%；PPT 套模板成功率 ≥80%。

### M3（第 12 周末）：智能与生态

**演示脚本**：

1. 同一文档第二次任务 → 不再重复询问格式与背景。
2. 打开 RunTrace → 定位一次失败到具体步骤与错误码。
3. 通过 MCP 连接一个外部数据源 → 读取数据并写入表格。
4. 查看任务成功率、修复次数、撤销率、单任务成本看板。

**通过线**：重复说明减少 ≥30%；反馈覆盖率 ≥30%；至少 1 个生态试点。

---

## 7. 周执行节奏（防跑偏）

| 时间 | 动作 | 产出 |
|---|---|---|
| 周一 | 从 Phase 计划取 1–2 个 WI，确认 Outcome 映射 | 本周目标卡 |
| 每日 | 更新任务状态；记录证据链接 | 工作日志 |
| 周四 | 真机自测；发现偏差立即调整 | 测试记录 |
| 周五 | 演示 + 指标快照 + 更新决策日志 | 周报 |
| 每月 | 对着“目标锚点”复盘：哪些工作没有服务 O1/O2/O3 | 目标复盘记录 |

### 提交/PR 模板

```text
Outcome: O2
Job Pack: Word 审阅润色
Work Item: E1-03
Metric: undo_success_rate
Acceptance: 取消 5 步任务后一键撤销，文档 hash 恢复
Evidence: <真机录屏/日志/DB 行>
Risk: <回滚策略>
```

### 需求准入清单

任何一个新需求，开始前必须回答：

1. 对应 O1、O2、O3 中的哪一个？
2. 属于哪个任务包？
3. 哪个指标会变化？变化多少算成功？
4. 用户可观察的验收是什么？
5. 真机验证脚本是什么？
6. 失败/回滚策略是什么？
7. 是否新增 Tool 或入口？为什么现有任务包不能承载？

七个问题里超过两个答不上来，就不进入开发。

---

## 8. 决策日志（需要用户拍板）

| ID | 决策 | 选项 | 建议 | 影响 | 状态 |
|---|---|---|---|---|---|
| D1 | 撤销实现策略 | 文档快照 / 逆操作 / 混合 | **混合**：Word 快照 + UndoRecord；Excel 逆操作 + 范围快照；PPT 事务标记 + 单页快照 | 决定 E1 的工作量与可靠性 | 待确认 |
| D2 | 指标采集方式 | 仅本地 / 匿名上报 / 两者 | **本地优先 + 可选匿名聚合**，默认关闭 | 影响隐私文案与看板实现 | 待确认 |
| D3 | 免费额度与计费 | 纯 BYO Key / 赠送额度 / 订阅 | 先 BYO Key + 新用户试用额度 | 影响首次激活与商业模式 | 待确认 |
| D4 | 模板策略 | 内置模板 / 用户导入 / 企业目录 | 分阶段：内置 3 套 → 用户导入 → 企业目录 | 影响 PPT/Word 模板体验 | 待确认 |
| D5 | Deepseek/Doubao 独立入口 | 合并进主面板 / 保留 / 下线 | 合并模型配置与历史，移除独立壳 | 影响体验一致性与维护成本 | 待确认 |
| D6 | PPT 重生成粒度 | 整副 / 单页 / 两者 | 先单页重生成，再支持整副 | 影响 E9 的回滚与成本 | 待确认 |
| D7 | 是否引入子任务分诊（E17） | 不做 / 只做只读分诊 / 通用多 Agent | **只做只读分诊**：只读并行、写串行、子任务不弹审批、成本有闸门 | 长文档与大表任务质量↑，但Token成本与解释复杂度↑ | 待确认 |

---

## 9. 本周立即动作（Week 1）

1. **E5-01**：建立单一能力清单与校验脚本，先跑出当前所有漂移项。
2. **E5-03**：清理 PPT 假成功（`ApplyTheme` 等），统一返回真实结果。
3. **E1-01/E1-02**：写信任合同，先把“本轮变更日志”做出来。
4. **E2-01**：会话按宿主过滤，成本最低、感知最强。
5. **建立基线**：用 10 个真实文档任务记录当前完成率、修复次数、撤销率，作为 M1 对照。
6. **落地记录**：E5-01/02/03、E1-01/02/03/04、E2-01 已完成代码与自动化校验，真机验收步骤见 `week1-landing-report.md`。
7. **W2 落地记录**：E1-05 审批粒度、E2-02 会话绑定文档、E2-03 自动标题/重命名/删除已完成代码与自动化校验，落点与真机验收见 `week2-landing-report.md`。同轮修复门禁盲点：`smoke-memory-pipeline.ps1` 仍按旧契约插事件（`event_type='message'`）并用 assistant 事件 id 当 `TargetId`，与 E2 单源化后的生产契约不符，导致抽取任务“成功但零记忆”；已按 `chat_message` + 会话 id 修正，并补 TTL 检索断言与夹具自清理。
8. **W2 收尾落地记录**：E3-03 空 Key 阻断与 E3-01 引导向导已完成代码与自动化校验，落点与真机验收见 `week2-landing-report.md` §3/§4/§6.3/§6.4。重要取舍（真机验收后收敛）：E3-01 首版把“选服务商 / 填 Key / 测连接”也做成了面板卡片，复查确认这三步在 `ConfigApiForm` 里早已完整存在（填 Key 校验、模型选择、连接测试用的还是同一个 `ModelConnectivityTester`），属于重复实现，已删除；向导只保留此前确实没有的隐私确认与示例任务，配置动作走既有 `openApiConfigForm` 桥。连接测试从 `ConfigApiForm.ValidateApiAsync` 返回布尔改为返回 `ModelConnectivityResult`，失败弹窗原样给出 `detail/reason`，不再只说“请检查 API Key”。附带修掉两个真实缺陷：一是 `ConfigManager.SaveConfig` 此前把真实 Key 直接序列化进 `office_ai_config.json`（明文落在“我的文档”），现在逐个写入 DPAPI、文件里只留 `[encrypted]` 占位，并在 `Finally` 恢复内存中的真实 Key（MCP、意图识别、翻译直接读 `cfg.key`）；二是真机复现的旧快照问题——`ConfigSettings` 只在 `LoadConfig` 的 `If item.selected` 分支里赋值、`SaveConfig` 完全不刷新，导致清空 URL/Key 或删除当前服务商后面板仍拿着上一次的端点发请求、E3-03 闸门也会被旧 Key 放行；现在由 `ApplyConfigSnapshot` 单点无条件投射，`LoadConfig`/`SaveConfig` 都调用它，`ConfigApiForm` 两处手写快照随之删除。同轮新增两条通用守卫：新 JS 未同时登记 `Resources.resx` + `ResourceExtractor` 时构建变红；面板 post 的消息类型若没有宿主 `Register/Case` 处理函数也变红（已用假类型反向验证）。
9. **W3 起点**：按 E1-06 → E1-08 → E4-01/02 → E3-02 推进；E3-02（每端一个示例任务）直接复用 E3-01 已有的示例清单与 `appType` 映射，不另开一套。E17 只读分诊不在 W3 做，先补齐 E1-06 幂等再启动。W2 的真机验收（E1-05、E2-02/03、E3-01/03）仍未执行，需要 Word 2021 真机 + 重启宿主刷新前端副本。
10. **W3 落地记录（E1-06 replan 幂等保护）**：代码与自动化校验已完成，落点、约束与真机清单见 `week3-landing-report.md`。机制是新增 `Agent/Execution/AgentWriteLedger.vb` 作为本轮唯一写入台账（宿主确认提交才算数，参数指纹做键序归一），LoopEngine 在 ACT 之前查台账，跨计划版本命中就不进宿主、不建撤销点、不弹风险确认；只有换成全新计划才推进计划版本。重规划覆盖校验改为并入上一版已成功动作，避免把已完成工作误判成漏项。已落盘写入清单以系统级约束同一份文本下发给反思与重新规划。顺带修掉一个长期存在的产品缺陷：`PromptManager.BuildReflectionPrompt` 用三参数 `If` 把字符串观察当布尔，运行时抛转换异常且被 `Catch` 静默吞掉，replan 分支此前从未真正跑通——现在观察按长度收敛输出，异常改为脱敏 Warn 日志。门禁补强：`smoke-replan-idempotence.ps1`（已登记进 `run-code-checks.ps1`）用反射驱动 LoopEngine 做失败注入并断言重复写入被跳过；P0 守卫新增“每个 `.vb` 必须被本项目 `.vbproj` 编译”与一组台账/replan 版本推进守卫；为守住 800 行门禁，把 `LoopEngine.Planning.vb` 的覆盖校验职责拆到 `LoopEngine.Coverage.vb`，未走加白名单路线。
11. **W3 落地记录（E1-08 信任层 UI 收口）**：代码与自动化校验已完成，落点、约束与 12 条真机清单见 `week3-landing-report.md` §2/§3/§4。核心原则是**卡片只展示宿主算得出的事实**：范围取自 E1-04 已捕获的 `lastContextPack`（选区项数/地址、文档名），数量取自计划步数与已解析工具数，风险取自 `ToolRegistry.GetTool(id).RiskLevel` 的最大值并列出需确认工具名，全部由 LoopEngine 在发给 UI 前补进 `ExecutionPlan.Trust`，模型自述不参与；同时把步骤事件从 `Action(Of Integer, Boolean, String)` 改为状态令牌 `pending/running/completed/failed/skipped/aborted`，修掉“用户拒绝”和“审批放弃”都被画成红叉、且计划卡永远没有进行中态两个真实缺陷；replan 后重发计划卡复用同一容器并标记“第 2 版”，避免时间线停在上一版；终态变更摘要由 `AgentRunChangeLog.Entries` 单源生成（提交口径与 E1-06 一致），上限 8 条、单条 120 字，撤销按钮只在 Word + `FullyUndoable` + 有 runId 时出现，其余情况必须显示 `undoReason`，不让用户面对一个静默消失的按钮。门禁补强：新增 `scripts/check-agent-card-ui.js`（桩 DOM + `vm` 直接跑 `agent-card.js`，覆盖 trust chip、四种状态令牌、replan 不重复建卡、change summary 不把宿主文本当 markup、撤销按钮与 `undoReason` 互斥），并登记进 `run-code-checks.ps1`；P0 守卫补一组 E1-08 断言，含“`agent-card.js` 不得再声明 `getStepIcon`”——该函数与 `code-handler.js` 同名且加载顺序靠后，会覆盖代码卡的图标映射，属本轮附带修掉的既有缺陷。

12. **W3 落地记录（E4-01 Word 选区工具条）**：代码与自动化校验已完成，落点、约束与 18 条真机清单见 `week3-landing-report.md` §3/§4/§5。Epic 目标是「AI 在用户工作的地方出现，而不是要求用户先打开面板」，实现口径是**入口只说话、主链只有一条**：三个动作（改写/摘要/校对）与向导示例任务收进唯一清单 `ShareRibbon/Agent/QuickActionCatalog.vb`，`OnboardingService.BuildExamples()` 删掉自带的那份 `Select Case` 改读同一处，`CapabilityId` 必须能在 `WordCapabilityRegistry` 里查到真实 id，防止「入口说得到、能力做不到」；宿主侧只做「什么时候出现」，点击后一律交 `BaseChatControl.StartAgentFromHostEntry` -> `HandleStartAgent`，沿用同一个模型配置闸门、同一次上下文采集、同一条 `OfficeHarness -> AgentKernel -> LoopEngine` 主链和 E1-08 的预览/风险/摘要/撤销，不再由宿主拼 `sendMessageToServer` 的 JS 字符串。三条易踩的硬约束：**(a) 默认关闭** —— 浮层会碰到每一次选区操作，在面板设置 UI 补上开关之前只能是 `office_ai_chat_settings.json` 里的 opt-in，不能靠重新编译才关得掉；**(b) 不抢焦点** —— `WS_EX_NOACTIVATE` + `WS_EX_TOOLWINDOW` + 覆写 `ShowWithoutActivation`，否则点按钮先清掉了它正要作用的选区；**(c) 命名空间不得叫 `Agent.Entry`** —— VB 标识符大小写不敏感，该命名空间会劫持 `AgentChangeLog.ToJson()` 里的 `For Each entry`，整解决方案编不过（BC30112/BC30456）。窗口有效性按要求用 `ActiveWindow` 且必须是 `wdWindowDocument`，不用文档计数（WordAi 是 Application 级加载项，「有文档但没有活动文档窗口」是真实状态）；控制器在首次打开面板时才创建，不进 VSTO 启动关键路径，日志只记动作 id。门禁补强：新增 `scripts/smoke-selection-toolbar.ps1`（结构合同 + 用 `Add-Type` 现场编译纯几何模块跑落点/翻转/clamp 行为断言，不需要 Word）并登记进 `run-code-checks.ps1`，P0 守卫补一组 E4-01 断言（守卫文本一律 ASCII，PS 5.1 下中文串不能当锚点）。

13. **W3 落地记录（E4-02 三端右键菜单）**：代码与自动化校验已完成，落点、约束与真机清单见 `week3-landing-report.md` §4/§5/§6。同一份 `QuickActionCatalog` 直接扩到三端（Excel：解读选区/生成公式/整理格式，PowerPoint：改写选中文字/提炼本页要点/校对本页），入口层不许存在任何一张需要人工同步的本地清单；共享层 `ShareRibbon/Agent/HostContextMenu.vb` 只碰 OCOM 的 `Office.CommandBar`，三端适配器只回答「挂哪条菜栏」和「此刻上下文够不够」这两个宿主问题。三处判断值得记下来：**(a) 不做置灰** —— 仓库里那份 `MicrosoftOfficeCore.15.0.0` PIA 的 `CommandBar` 没有任何事件面（反射确认只有 `_IMsoOleAccDispObj`/`IAccessible`，`CommandBarEvents` 根本不存在），`BeforeRightClick` 编不过；即使将来能置灰，Attach 时算一次、之后不刷新的静止灰色也是在骗人，所以改成点击时探测 + 状态栏给出原因（`GlobalStatusStripAll.ShowProgress`，不弹窗、不抢焦点）。**(b) 命名空间一律 `HostEntry`** —— `Namespace Selection` 与互操作类型同名时 VB 会优先解析命名空间，直接把 `PowerPointAi` 里的 `sel As Selection` 打成 BC30182（应为类型）；这与 `Agent.Entry` 劫持 `For Each entry` 是同一类坑，比逻辑错更难查，因此三端入口层统一用不会撞名的 `HostEntry`。**(c) Excel/PPT 动作一律不声明 `CapabilityId`** —— 这两个宿主还没有 capability 注册表（开放待办 #36），宁可少声称，也不要「入口说得到、能力做不到」，smoke 直接断言这一点。可撤销性是改宿主菜单的底线：只按 `aiq:` Tag 前缀认领并倒序删除自己加的项、`Temporary:=True` 不落进用户配置、按钮 RCW 留在列表里否则 Click 回调被 GC 掉、Shutdown 必须 `Detach()`。默认同样关闭（`enableContextMenu`），因为它改的是用户本来就有的右键菜单。顺带把三端 `Ribbon1.StartAgentFromRibbonAsync` 从注入 `sendMessageToServer` JS 改为 `StartAgentFromHostEntry`：`router.Register("startAgent", ...)` 本来就直通 `HandleStartAgent`，因此行为等价，但少了一份注入式载荷和一处 WebView 时序依赖。门禁补强：新增 `scripts/smoke-host-context-menu.ps1`（**反射调用编译产物** `QuickActionCatalog.ActionsFor`，断言每端恰好 3 项、id 前缀正确、无重复 id、未接入宿主返回空表；再加结构合同：共享层不得引用宿主互操作类型、不得出现 `BeforeRightClick`，三端必须过开关、走单一入口、在面板可见路径挂载而不在 `*_Startup` 挂载；所有「禁止出现」类断言先剥掉单引号开头的注释行，否则会被解释性注释自己触发）并登记进 `run-code-checks.ps1`，P0 守卫补一组 E4-02 断言。

14. **W3 落地记录（E3-02 每端一个一键示例）**：代码与自动化校验已完成，落点、约束与真机清单见 `week3-landing-report.md` §5/§6/§7。验收要求是「Word/Excel/PPT 各一个一键示例，使用当前文档或内置样例」，落地口径是把 E3-01 已有的示例清单从 `String()` 升级为 `SampleTask(Text, RunRequestText?)`：**仍然是那一处清单、那一个 `appType` 映射**，只是让每端第一条带上可以直接发给主链的请求，因此向导、选区工具条、右键菜单和一键示例共用同一个话术来源，没有新开第二张表。三个判断：**(a) 空文档兜底写在请求话术里，而不是写在入口代码里** —— 每条一键示例同时说清「有内容怎么办」和「文档还空着就先补一段内置示例」，由能读上下文的主链自己决定，入口不猜用户有没有选区、也不为这一层判断去调 COM；这与根规则「不要靠关键词 if/else 绕过 harness」一致，代价是真机必须确认三端在空文档下确实没有走进死胡同。**(b) 一键示例仍然是入口** —— 面板只发 `startAgent`，于是配置闸门、上下文采集、计划卡、风险确认、变更摘要、撤销一条不落；`onboardingRunSample` 里禁止出现 `needsSetup`/`sendChatMessage`/`startAgentExecution`/`executeCode`，宿主侧禁止再注册 `runSample` 一类的第二个示例处理器，都有 P0 守卫。**(c) 卡片文案不能骗人** —— PowerPoint 那条在有内容的演示文稿里只做只读提炼、只在空演示文稿里生成示例页，因此文案写成「提炼本页要点（空文档就先做一页示例）」而不是「生成一页汇报初稿」。门禁补强：`smoke-onboarding-wizard.ps1` 扩 E3-02 段——反射取 `QuickActionCatalog.SampleTasks` 与面板实收载荷逐条比对 `text`/`request`（这是「向导不再自己拥有话术」的证据而非注释承诺），断言每端恰好 1 条可跑且排第一、三端请求互不相同、每条都带「预览/空文档分支/内置示例」三个语义锚点；锚点字符用 `[char]0xXXXX` 现场拼装以保持脚本 ASCII。P0 守卫新增 E3-02 段并已用变异探针验证非空跑（删掉 `item("request")`、把 JS 的 `startAgent` 换成自定义类型，审计立刻报红）。注意 `onboarding.js`/`styles.css` 属前端资源，真机验证前必须重启宿主刷新 `%LOCALAPPDATA%\OfficeAI\www` 副本。

15. **W4 落地记录（E1-07 Excel/PPT 接入本轮撤销）**：代码与自动化校验已完成，落点、约束与 12 条真机清单见 `week4-landing-report.md`。验收标准是「两端取消/失败后可撤销本轮；行为与 Word 一致」，而 Excel/PowerPoint 客观上没有「一条命令 = 一个撤销记录」的边界（`UndoManager.CreateExcelUndoPoint` 直接把 `CanUndo` 置 False，PPT 只有 CreateSlides 链的按标记删页），所以「一致」定在**用户看到的事实**上：结束时能说清能不能撤、依据是什么，能撤就一键回到执行前，不能撤就显示原因。合同层把可撤销判成两条互斥路径且都要求覆盖**全部**残留改动：`NativeUndoable`（Word：每步都有宿主撤销单元 + 首尾文档指纹）或「每步都有可重放还原快照」，半撤半留不算可撤销——这是本轮最重要的产品判断，宁可不出按钮，也不给用户一个只回退一半的按钮。快照走 `ToolResult.RestoreSnapshot`（`<JsonIgnore>`，与 `VisualEvidence` 同类）而不是 `Observation`：`LoopEngine.Planning` 会把 Observation 压成 8192 字符喂给 planner 并落 `agent_run_step.observation_json`，文档正文跟着走等于同时犯隐私、体积和提示词污染三条红线；权威落点只有迁移 018 新增的 `agent_run_change.restore_json`，上限 60000 字符，超限即视为无快照。宿主侧各自的闸门决定「这一步是否配得上一个撤销承诺」：Excel 只在区域值/公式/数字格式且行列数·图表数未变时产快照，写回时仅对以 `=` 开头的单元格恢复公式（否则 `"1-1"` 被 Excel 重新解释成日期）；PowerPoint 只在页数/形状数/动画数/格式签名全未变而文字或备注确实变了时产快照，重放按「序号 + 形状名」双重确认才写 `TextRange.Text`，文字一致的形状跳过以减少格式重置，备注只在执行前确有备注形状时才写。两端 `PerformUndoAgentRun` 同构：先校验快照覆盖条数，再按写入逆序重放，任一拒绝就如实报「已还原 N 处、剩余 M 处未还原」且**不**标记本轮已回滚，用户处理冲突后可再点一次。附带修掉 `ValidateObservedOutcome` 的一个真实缺口：把成功批次降级成 `CodePartialApply` 时原来会丢掉快照，而那正是最需要撤销的场景。UI 层删掉 `BuildTerminalDescriptor` 里的 Word 专属闸门，改由合同事实驱动 `undoReason`，按钮文案区分「N 个撤销单元」与「按快照还原 N 处」。门禁：`smoke-change-log-and-session-scope.ps1` 扩 E1-07 段（含「变更记录 JSON 不得含快照正文」「一步有快照一步没有必须整轮拒绝」「超上限快照被拒」），P0 守卫新增一组断言（含「终态描述里不得再出现 `_getOfficeAppType(), "Word"`」这类反向守卫），schema 基线与空库版本断言同步到 18。已声明的不作为：格式/动画/结构改动、`OfficeObjectOperation` 声明式批次都不产快照，因此这些写入本轮不可撤销——按原计划留给 E5 之后的能力收口，不在这里用假快照凑覆盖率。

16. **W4 落地记录（E4-03 全局快捷键 / 命令面板）**：代码与自动化校验已完成，落点、约束与 12 条真机清单见 `week4-landing-report.md` §2。验收标准是「任一宿主内可唤起；支持自然语言与命令」，实现口径是**状态只有一份、出口只有一句话**：快捷键轮询与命令面板状态都住在 `ShareRibbon/Controls/Services/`（`HostHotkeyWatcher.vb` + `CommandPaletteService.vb`），由 `BaseChatControl` 惰性持有、`Disposed` 时 `Shutdown()`，三个宿主各贡献且只贡献一个 `EnsureHostTaskPaneVisible` 覆写（`Globals.ThisAddIn.ShowChatTaskPane()`）。于是「任一宿主内可唤起」是结构性质，而不是三份需要人工同步的接线，改完组合键立即生效、不必重启宿主。四个必须记住的判断：**(a) 组合键安全合同** —— `HostHotkeyChord.IsSafe` 要求「有主键且带 Ctrl 或 Alt」，因为裸键会吃掉用户正常输入，而 Word 本身已占用 `Ctrl+K`（插入超链接），默认值 `Ctrl+Alt+K` 正是为此避开；修饰键按**精确集合**比对，多按一个 Shift 不算命中。**(b) 全局键盘状态必须配前台判定** —— `GetAsyncKeyState` 是系统级状态，用 `GetForegroundWindow` + `GetWindowThreadProcessId` 与宿主进程 id 比对，Word 与 Excel 同时打开、共用同一组合键时不会出现「看不见的窗口把面板弹出来」；轮询与下沿沿用 `WordCompletionManager` 的既有做法，回调期间停表、只在「未 Dispose / 未停监听 / 回调仍在」三条件成立后才重启，长按和输入法候选期间的重复按下不会变成反复唤起。**(c) 明确禁止** `SetWindowsHookEx`、`RegisterHotKey`、`keybd_event`、`SendInput`：前两个是全局钩子与系统级占用，后两个是伪造按键。**(d)「按了没反应」比没有快捷键更糟** —— 非法写法在 `LoadSettings` 归一化、`HandleSaveSettings` 校验并在状态栏说明、`Start` 返回 False 再提示三处拒绝，绝不落一个永远不生效的开关；页面未就绪时 `Open()` 排队而不是丢按键。话术仍只有 `QuickActionCatalog` 一处，服务代码里不存在 `startAgent`/`sendMessageToServer`，唯一起点在 `command-palette.js`；一度加上的 `router.Register("openCommandPalette")` 已删除——面板按钮本地就能开，多一条路等于多一套终态。本轮顺带修掉两个与「开关能不能关掉」直接相关的真实缺陷：`ChatSettings.SaveSettings` 此前用固定字典重建 JSON，凡是字典里没有的键都会在每次保存时被静默删除（`enableMcpList` 就是这样被抹掉的），现在改为读改写同一个 `JObject`；`HandleSaveSettings` 此前把 `EnableSelectionToolbar`/`EnableContextMenu` 从静态字段原样写回，等于这两个开关是只读的——上一轮「默认开启却没有 UI 开关」被驳回的正是这一点。现在一律从载荷读回（缺键保持当前值，旧版前端也洗不掉开关），面板设置里补齐三个入口开关与组合键输入框，**E4-04 的「选区自动引用默认开启」自此才有前提**。门禁：新增 `scripts/smoke-command-palette.ps1`（组合键解析与拒绝用例、监听器生命周期、反射驱动真实服务断言「就绪前一条都不发、就绪后恰好一次同步 + 一次展开、`ensurePaneVisible` 只调一次」、四处资源登记、剥掉注释后的反向守卫）并登记进 `run-code-checks.ps1`，`node --check` 名单加入 `command-palette.js` 与 `settings-manager.js`。已声明的不作为：不做跨应用系统级快捷键、不在入口层解析命令语法（自然语言与命令同一句话交回主链）、不做搜索历史与 frecency 排序。明确推迟到 E4-05 的两个缺陷已写入报告：宿主在 `ChatTaskPane_VisibleChanged` 里读 `ChatSettings.EnableContextMenu` 时可能尚未加载配置文件（进程内第一次打开面板，右键菜单可能不挂），以及 `RemoveHandler ChatBrowser.CoreWebView2.NavigationCompleted` 摘错了事件（处理器绑在 `ChatBrowser.NavigationCompleted`），导致 `InitializeSettings`/`PushState`/`ApplyHotkey` 每次导航都重跑。

17. **W4 落地记录（E4-04 选区自动引用默认开启）**：代码与自动化校验已完成，落点、约束与 12 条真机清单见 `week4-landing-report.md` §3。验收标准是「默认携带选区上下文；UI 明确显示引用了什么」，难点不在把默认值翻成 True，而在翻完之后仍然成立的三件事：开关必须真的能关掉、带出去了什么必须让用户看得见、摘要只能是显示而不能变成第 N 份持久化数据。实现口径是**决策点只有一个**：`BaseChatControl.BuildMessageWithAutoSelection` 先判 `selectedCellChecked` 再调用 `AppendCurrentSelectedContent`，因此三端 × {ChatControl, DoubaoChat, DeepseekControl} 共 9 个 override 从此只回答「怎么把选区写成提示」，不再回答「要不要写」（PowerPoint 里那份重复闸门已删除，Doubao/Deepseek 顺带收敛到同一开关下）；两条发送路径（聊天回合、Agent 回合）都改走这一个包装方法，共享层除包装本身不存在第二个调用点，门禁用出现次数把这句话钉死。显示侧的固有困难是 **Agent 回合不渲染用户气泡**，所以新增 `appendTurnSelectionReference` 纯文本通道复用既有 `.chat-message-references` 折叠样式，同回合重复同步先删旧条，两个参数一律经 `ToJsStringLiteral` 转义。摘要坚持从**真正附加上的那段增量**反推（`appended.Substring(message.Length)`），曾考虑改用 `CaptureOfficeContext`/`ContextPack.FromOfficeContext` 的 `Selection.Ref/Address/ItemCount/Preview` 生成标签，代价是选区被第二次注入提示、每回合多一次 COM 读取，因此否决。明确接受并写进设置说明的耦合：Word 的 override 同时记录 `PendingSelectionInfo`，共享层据此绑定本轮 `requestUuid` 供选区对比/写回，开关关掉后不再记录快照——「不引用也就不写回」是同一条语义，不是漏洞。本轮附带修掉一个**静默失效**的真实缺陷：`InitializeSettings` 注入脚本里两行 `getElementById('scrollChecked')`/`('selectedCell')` 指向面板中不存在的 id（实际是 `settings-scroll-checked`/`settings-selected-cell`），WebView2 注入块顺序执行、null 上取 `.checked` 抛 TypeError 后**其后每一行都不再执行**，也就是 E4-01/02/03 的三个入口开关与命令面板按钮同步一直是坏的；修法是删掉这两行而不是补 id，并在门禁里加通用守卫「注入块引用到的每个 id 都必须存在于面板 HTML」（本轮 14 个 id 全部核对通过），同时把这两个死 id 单独钉住。默认值仍遵守「缺键才用默认」：`LoadSettings` 只在 `ContainsKey("selectedCellChecked")` 时覆盖，显式存过 `false` 的用户不被翻回来——这正是 §9/16 那轮补齐面板开关后才允许的动作。门禁：新增 `scripts/smoke-selection-reference.ps1`（反射读编译后的 `ChatSettings.selectedCellChecked` 真实默认值 + 上述单点/次数/反向禁止 + 摘要必须增量反推与 160 字符折行截断 + `_autoSelectionReference` 不得出现在 `AppLogger`/`AddMessage`/`History`/`Memory`/`Observation` 行 + JS 通道禁止 `innerHTML`/`fetch`/`XMLHttpRequest`/`invokeMethod` + 开关可即时切换与载荷回写）并登记进 `run-code-checks.ps1`。已声明的不作为：不合并聊天路径既有的 `引用内容 (N)` 卡片、不做引用目的推断、不改三端采集逻辑本身、不为摘要新增持久化字段；`DescribeAutoSelection` 挂在 `MustInherit` 的 WinForms 控件类型上无法反射加载（连带解析 WebView2/interop 面），因此按源码不变量守卫。

18. **W4 落地记录（E4-05 事件绑定卫生）**：代码与自动化校验已完成，落点、约束与 10 条真机清单见 `week4-landing-report.md` §4。验收标准是「不重复绑定；重复打开/关闭不产生双触发」，这条最难的地方在于**重复执行在这里的表现不是报错而是「看起来没副作用」**，所以产出不是一个新机制，而是一个能被门禁钉住的形状。核心决定是把闸门放在**文档**而不是控件上：`BaseChatControl._pageBootstrapDone` 只在实际装载新文档处复位（`ChatBrowser.CoreWebView2.NavigateToString(htmlContent)` 之前那一行），处理器 `If Not e.IsSuccess Then Return` → `If _pageBootstrapDone Then Return` → **占位在第一个 `Await` 之前** → 才跑 `InitializeSettings()`/`InitializeMcpSettings()`。占位必须在 await 之前，否则 `Task.Delay` + `InvokeAsync` 期间排队的第二次导航完成会同样通过检查，闸门形同不存在。绑定唯一化的做法是**删除**而不是补解绑：原先 `NavigateToString` 后又 `AddHandler ChatBrowser.CoreWebView2.NavigationCompleted`，而处理器早就用 `Handles ChatBrowser.NavigationCompleted` 绑过一次，处理器顶部那句 `RemoveHandler` 摘的是另一个事件、永远摘不掉声明式那份，于是首次装载跑两遍、之后每次导航都再跑一遍；现在只留声明式一份，并在删除处留下「为什么不能再补回来」的注释。第二个真实缺陷是**读开关早于读文件**：三宿主在 `ChatTaskPane_VisibleChanged` 里判 `EnableContextMenu`/`EnableSelectionToolbar`，而那一刻进程内可能还没有任何 `ChatSettings` 实例读过盘，拿到的是代码默认值——用户开过的开关在第一次打开面板时不生效。修法是一个共享的 `ChatSettings.EnsureLoaded(applicationInfo)`（首次才构造、重复调用返回同一实例），三宿主在各自的入口方法**之前**各调一次；曾考虑在 `ThisAddIn_Startup` 统一加载，代价是把文件 IO 放进宿主启动关键路径且没有任何事件顺序保证，因此否决。顺带的口径统一：Excel 的 `MainForm_Load` 自己调过一次 `InitializeSettings()`，与基类装载路径重复，而 Word/PPT 本来就没有——删掉后三端都是「一次装载一次初始化」。E4-01/02/03 的生命周期代码经逐一核对**本来就对称**（`HostContextMenu.Attach` 的 `If _attached Then Detach()`、`aiq:` Tag 反序删除、按钮 Click 成对；`HostHotkeyWatcher` 的 Tick 成对增删；命令面板懒创建守卫；`EnsureWidthTimers` 的 `Is Nothing` 守卫与 `ThisAddIn_Shutdown` 的 Stop+Dispose/Detach），本轮没有为「有改动而改」，只把这些事实写进 `scripts/smoke-event-binding.ps1`——「已经对称」不属于任何人的记忆，下次新增入口最容易破坏的就是它。另有一条**看着像漏、实际不可达**：`LoadLocalHtmlFile()` → `WebViewBridge.LoadHtmlFileAsync(ChatHtmlFilePath)` 是唯一没配复位的导航点，但 `ChatHtmlFilePath` 是带时间戳的**保存目标**（读取那一刻必然不存在，`File.Exists` 直接 Return），宿主侧 `loadChatHtml` 又只在进程内首次显示面板时调用，因此归 E5-04 清理，本轮不为它写一条永不执行的复位分支。门禁：新增 `scripts/smoke-event-binding.ps1`（反射两次 `EnsureLoaded` 必须 `ReferenceEquals` + 禁止 `AddHandler/RemoveHandler ChatBrowser.CoreWebView2.NavigationCompleted` 回归 + `Handles` 必须存在 + 复位早于 `NavigateToString` + 处理器内「判→占→第一个 Await」顺序 + 标志位出现次数恒为 4 + 三宿主 `MainForm_Load` 不得出现 `InitializeSettings()` 且共享层裸调用恰好 1 处 + 三宿主 `VisibleChanged` 必须在 `EnsureContextMenu()` 前 `EnsureLoaded(New ApplicationInfo(...))` + 右键菜单/观察器/定时器的成对守卫）并登记进 `run-code-checks.ps1`。已声明的不作为：不给 `Handles` 绑定补手工解绑、不动 Doubao/Deepseek/爬虫三套各自加载远端 URL 的控件、不做配置热重载、不消除 §4.3 的不可达入口、不把三处 `EnsureLoaded` 收成一次 Startup 调用。

19. **W4 落地记录（E2-04 旧数据迁移/回填·会话归属）**：代码与自动化校验已完成，落点、归属策略、本机实测数据、门禁与 8 条真机清单见 `week4-landing-report.md` §5。验收标准是「老会话不丢失；无 app_type 的记录有明确归属策略」，根因是 E2-01 之后 `app_type` 事实上只剩两种合法值：三个规范宿主名，或空串（通用数据，三端都可见）；**第三种值既不等于任何宿主、又拿不到空串豁免，那条数据在 Word/Excel/PowerPoint 里同时查不到，而且不报错**。哨兵值确实进过库——`GetOfficeAppType()` 基类默认仍 `Return "Unknown"`，三端子类各自覆盖成规范名，所以新数据不再生成、老数据没人管。本轮先量后改（`Read Only=True` 探针，两库都没跑迁移）：安装版库 schema 17，281 条消息 = 269 空串 + 12 `Word`、白名单外 **0** 行，27 条会话档案 `app_type` 全空且只有 2 条带文档名（`report.docx`、`文档1`），3 条 `agent_run` 全是 `Word`，81 个会话里 **0 个**能连到轨迹；而同一库的 `atomic_memory` 231 行里 **123 行是 `Unknown`**——记忆链读侧用的是同一条「等于当前宿主 OR 空 OR NULL」规则，所以同一批哨兵值让老记忆也在三端同时失效。**这个测量结果改变了做法**：只修会话历史等于给同一个缺陷族留了一半，于是口径落在**列**上而不是落在「历史列表」上，凡带 `app_type` 的六张表（`conversation_event`/`agent_run`/`agent_run_change`/`session_summary`/`atomic_memory`/`memory_item`）一起收口。迁移 019 按证据强度顺序填、**每步只填空不覆盖已解析归属**（因此可重复执行）：先归一（必须先做，否则「最新一条 `Unknown` 盖住更早真实宿主」的数据永远落不回那个宿主）→ `agent_run` 最近一条带宿主轨迹（与迁移 14 同规则）→ `agent_run_change` 最近一条（按 `observed_at`）→ `session_summary.document_name` 扩展名分组（`.csv` 故意不推断、未保存名不猜）→ 档案下发事件 → 事件反填档案（E2-05 分组读的就是档案，不补则老会话永远显示成「通用会话」）→ 无证据者保持空串。写侧新增唯一口径持有者 `ShareRibbon/Storage/AppTypeContract.vb`（单独成文件而非塞进已 1000+ 行的 `OfficeAiDatabase.vb`），六处 `app_type` 绑定全部先过 `Normalize`；**读侧过滤参数故意不动**，因为那里 `''` 本就是豁免，归一读取值会把「按宿主精确查」变成「查全部」。另一个被否决的方案是放宽读过滤（「匹配不上也算」）——那会把别的宿主历史漏进当前端。本机净效果：1 个会话按 `.docx` 落 Word、3 个会话档案被事件反填、123 条记忆重新可召回，76 个完全无证据会话 + 1 个只有未保存文档名 `文档1` 的会话继续空串三端可见。门禁三重：`smoke-empty-db-initialization.ps1` 扩到 schema 19 + 四类 legacy 会话夹具 + 记忆表夹具，断言六张表白名单外 0 行、归一后的记忆必须被真实 `MemoryRepository.ListAtomicMemories` 在三端各召回一次、可见性走真实 `GetRecentSessions`，最后**重跑已部署的 019 文件**再断言归属不变（同时证明文件与内联兜底同义且可重跑）；`audit-p0-guardrails.ps1` 新增 E2-04 段，要求六个写入文件都真的引用 `AppTypeContract.Normalize`、`AgentMemoryRepository` 两个写入点分别归一（计数 ≥2）、019 文件与内联兜底**逐表**都含归一语句（防「只修文件不修兜底」），并已用独立探针验证这些正则在真实源码上逐表命中、不是空跑；`smoke-change-log-and-session-scope.ps1` 钉住写边界（`InsertMessage(..., "Unknown ")` 之后必须三端可查且两张表落库值为空串）。已声明的不作为：不把 `.csv` 归 Excel、不给未保存文档名猜宿主、不放宽读过滤、不改 `GetOfficeAppType()` 的 `"Unknown"` 默认值（那是「子类未实现」的信号，要收口应改成 `MustOverride`，归 E5-04）、不引入第四种合法值、不动无 `app_type` 列的 legacy `conversation` 表、不删那 123 行真实用户记忆。
20. **W4 落地记录（E2-05 历史 UI 分组 + 空状态引导）**：代码与自动化校验已完成，落点、分组规则、实测分布、门禁与 10 条真机清单见 `week4-landing-report.md` §6。验收标准是「按文档/时间分组；空状态有引导」，这条的陷阱不在画不出组，而在**谁拥有顺序**：`GetRecentSessions` 的 `ORDER BY` 是「当前文档整块在前 + `MAX(created_at)` 倒序」，前端一旦按时间重排就把 E2-02 那条合同冲掉了（上一轮已经在注释里钉死，本轮把它升级成门禁——JS 里出现 `sort(` 直接失败）。所以分组定成**纯呈现**：`buildGroups` 按「首次出现」切桶、桶内保持后端顺序，`scope=all` 分文档组（当前文档 → 其它文档 → 通用会话），`scope=document` 只剩一篇文档、改分时间段（今天/昨天/最近 7 天/最近 30 天/更早）。两个键是本轮真正补进去的合同：`lastActiveAt` 就是那条排序键（此前只有 `MIN(created_at)` 回传，按它分桶会让「上周开、今天还在聊」的会话排在最前却挂「更早」的标题，自己和自己矛盾），`documentId` 是 `SessionDocument` 早就准备好的路径哈希（注释写明「只用于分组」）——用文件名当组键会把两篇 `预算.xlsx` 并成一组把人带偏。日期只截 `YYYY-MM-DD` 自己算天数差，因为 `"2026-10-24 09:12:33"` 不是标准 ISO、`new Date(字符串)` 各引擎不一致，解析失败一律落「更早」而不是毁掉整列。条目 meta 交出文档名（组头承载）并把时间收到分钟。空态从两条文案变成三种带动作的情形：已保存文档 / 未保存文档（解释为什么只能先记成通用会话）/ 全局无会话（给一句可抄的具体任务示例），动作只有两种且都不另开发送路径——「直接开始」关面板 + 焦点回 `smart-input`（实测**不产生任何 post**），「新手引导」复用 E3-01 的 `window.showOnboarding` 且页面没有它就不出现。本机实测（只读探针，未跑迁移）：安装版库 81 个会话 / 27 条档案 / 只有 2 条带 `document_id`，按文档聚合 = 通用 79 + `文档1` + `report.docx`，按 `lastActiveAt` 分桶 = 昨天 4 / 最近 7 天 4 / 更早 73；首条与末条不同天的会话 1 个、恰好没跨桶边界，本机按 `createdAt` 分桶目前 0 条会落错组——**这条改动是给形状而不是给已发生的故事故底的**，报告里如实这么写；同名不同哈希本机 0 例，调试版库 0 会话（门禁夹具跑完即清），所以行为验证只能落在桩与夹具上。门禁三层：新增 `scripts/check-history-grouping-ui.js`（最小 DOM 桩加载真实 JS，断言 4 桶标签/计数/同名拆分、组序=首次出现、桶内保序、不丢不重、meta 只能是分钟精度时间戳、`lastActiveAt` 决定桶、脏时间戳落「更早」、三种空态文案与动作、`data-empty` 行为、XSS 口径、删除确认取消不落后端、两次 `init()` 仍只有一个委托监听）并登记进 `run-code-checks.ps1`；`audit-p0-guardrails.ps1` 双向钉字段合同（后端必须仍 SELECT `MAX(e.created_at)` 并映射、必须仍发 `lastActiveAt`/`documentId`，JS 必须仍有 `buildGroups`/`timeBucket`/`runEmptyAction`/`s.documentId`）加 **CSS↔JS 双向配对**（`styles.css` 里每条 `.history-group*`/`.empty-state-*` 规则必须仍被渲染，JS 渲染的 7 个新类必须仍有规则——本轮新增的样式若没有守卫，下一次删样式是静默的）；`smoke-change-log-and-session-scope.ps1` 在真库夹具上把首条事件回拨 2020 年后断言 `CreatedAt` 不变、`LastActiveAt` 不等于它、`DocumentId` 原样回传。6 个注入变异全部被断言拦下且报错逐条不同（桶序、桶数、组容器、空态动作、XSS、桶单调），CSS 侧删一条规则审计即如实失败——守卫非空跑已验证。已声明的不作为：不改后端排序、不做折叠/懒加载/搜索、不再按宿主二级拆组、不动那 79 条无证据会话的归属、不新增发送路径、不给未保存文档名猜路径、不加「共 N 条/已截断」页脚（payload 不带总数，写了就是假信息）。
21. **W4 落地记录（E5-04 不可达入口处理）**：代码与自动化校验已完成，入口清单、接通/移除两张表、门禁与 8 条真机清单见 `week4-landing-report.md` §7。这条的执行顺序被一个结构事实决定：`ChatCommandRouter.Dispatch` 对未注册 type 只 `Debug.WriteLine`，被注释掉的按钮、指向不存在函数的 `onclick`、查不到的 `getElementById` 全都静默通过——既没有异常也没有日志，所以**先把「入口」做成可执行清单再逐个处理**，否则「入口清单」只是一份会过期的文档。新门禁 `scripts/check-entry-reachability.js` 七条规则双向闭合（A 面板发的 type 必须注册 / B 注册的 type 必须有发送方 / C HTML 注释里不留可交互控件 / D 内联 `on*="fn("` 必须有定义 / E `getElementById` 字面量 id 必须有节点 / F `vsto` 桥方法必须有调用点 / G 三端 Ribbon 可见按钮必须有非空 Click 处理），语料必须含四模块全部 `.vb`：`SmartProofreadFocusMode.vb:361` 由 VB 拼出 `id="proofread-summary"`，`GetVstoBridgeScript()` 由 VB 注入发送方，只看 `Resources/` 会得出**相反结论**并误删正常功能。实测规模：面板 56 个发送点 / router 66 条注册（收口前 73）/ 内联事件 70 处 57 个函数名 / 字面量 id 103 处 / Ribbon 18 个按钮 Word 可见 14、Excel 13、PowerPoint 11。**接通**的那条正是计划点名的模板库：编辑器、列表、管理态、三端子类全都实现了，可见入口一条都没有——唯一调用 `enterReformatTemplateMode()` 的 PPT「排版」按钮在 `PowerPointAi/Ribbon1.Designer.vb:80` 被 `Visible = False`，Excel 同样隐藏，Word 的「排版」走智能分析，`TemplateFormatButton` 在共享基类 `BaseOfficeRibbon.Designer.vb:183` 就全局隐藏；做法是在输入条补一个 `#template-library-button`，复用已注册的 `getReformatTemplates` 与既有前端逻辑（**零新消息 type、零新发送路径**），顺带恢复自 `68e5fab` 起被注释的「新建排版模板」、给管理态卡片补「编辑/复制」（预置模板先拦住并提示复制）、并补上 `exitReformatMode()`——那个「退出排版」按钮自 `code-handler.js:951` 就指向一个从未定义的函数。**移除**的都是能静态证明走不到的：7 条 legacy 注册（Ralph 四个别名 + `startAgentExecution`/`refineAgentPlan`/`checkedChange`）连同 4 个专用 handler、§4.3 的 `LoadLocalHtmlFile` + `WebViewBridge.LoadHtmlFileAsync`（页面加载本就由 `InitializeWebView2` 的 `NavigateToString` 负责，每次桥调用前都 `_ensureInitialized()`，因此三端 `ShowChatTaskPane` 退回同步 `Sub`、Word 的 `Async` 变体与那条从未使用的 `loadDataCaptureHtml` 一并删掉）、`vsto.checkedChange`、`window.uploadStyleGuideDocument`。**同一个 bug 的第二层**是门禁自己的清单在腐烂：`run-code-checks.ps1` 里 `ralph-loop.js` 早已删除而「文件不存在就跳过」让它继续绿，纳入语法检查的只有 8 个文件而目录里实际有 20 个——清单改为**枚举真实目录**，缺失不再静默跳过。`audit-p0-guardrails.ps1` 负责「删掉的不得复活」（两个方法名、7 条注册、`GetOfficeAppType` 必须仍是 `MustOverride` 且三端各自 `Overrides`、模板库三点合同、门禁本身已登记），`smoke-selection-reference.ps1` 原来钉的是被删掉的 `HandleCheckedChange`，改为钉存活的那条：自动引用开关唯一的实时路径是 `HandleSaveSettings` 里的 `selectedCellChecked = jsonDoc("selectedCell")`，且 `checkedChange` 注册不得复活。8 个变异探针逐个破坏一条真实链路全部变红且报错各不相同（含 Word 隐藏按钮放开后「聚光灯可见但 handler 只剩注释」这条），探针脚本用 Buffer 原样读写以保住 BOM/CRLF 并逐文件比对字节。构建 PASS 四项目，`run-code-checks.ps1 -SkipBuild` 全绿。已声明的不作为：不改三端 Ribbon 可见性（放开哪个按钮是用户可见的产品决定）、不动挂在**可见**按钮上的 3290 行遗留面板栈 `BaseChat`/`BaseDeepseekChat`/`BaseDoubaoChat`（它可达，不属于「不可达入口」，但与 AGENTS.md「禁止另建第二套」直接冲突）、不删 `SaveFullWebPageAsync`（`saved_chat_*.html` 现在是写而不再读，属用户磁盘上的可见副作用）、不把规则 A–G 铺到遗留面板的内联 HTML。三项并入开放决策：D8（`saved_chat_*.html`）、D9（Ribbon 隐藏实现 + 遗留面板栈归属）、以及 E5-05 应吸收的方向——把 §7.4 这类静态合同闭合与 CSS↔JS 配对并成同一条防漂移门禁并进 CI。
22. **W4 落地记录（E5-05 CI 防漂移）**：代码与自动化校验已完成，改前漂移表、并轨决定、门禁补强与探针结果见 `week4-landing-report.md` §8。这条的验收标准是「四份能力清单不一致时构建失败 → CI 输出」，而 E5-01 早已把 P7/P8 两条一致性断言写进 `audit-p0-guardrails.ps1`——本轮量完才发现**它们从没在发布检查里跑过**，所以真正的产出不是第 N 条断言，而是把「谁在跑门禁」这一层收口。`git show HEAD:` 量到的四处漂移：`run-code-checks.ps1` 手写 16 条名单而磁盘有 17 条真门禁（漏的 `run-professional-design-smoke.ps1` 是会真 `throw` 的那条）；`build\RunReleaseChecks.ps1` 留着第二份私有名单、只有 5 条且不含 `audit-p0-guardrails.ps1`；它的 `Invoke-Step` 在子脚本 `exit 1` 时照样打印 PASS（实测：`LASTEXITCODE=1` 无人读，后续步骤继续跑）；JS 语法检查手写 7 个路径其中 `ralph-loop.js` 在 HEAD 已删除、而循环体只有 `if (Test-Path)` 没有 else，于是死名字长期为绿、20 个真实文件只有 6 个被检查，node 缺失更是 `Write-Warning` 后整段跳过仍报 all passed。做法是把清单变成**从磁盘推出的性质**：`scripts\*.ps1` 默认全是门禁，除非在 `$notGates` 里具名退出并写出理由（前缀规则 `smoke-*.ps1` 看着通用，实测正是它漏掉那条 PPT 门禁的），数量不写进代码、只留「发现规则自己坏了」的下限与「退出项必须仍在盘上」的反向检查；发布路径删掉私有名单改为委托 runner；两份 `Invoke-Step` 统一判退出码；JS 改目录枚举、缺 node 改 `throw`（只有 `-SkipNodeCheck` 能显式声明跳过）。**第二条并轨是同一合同的三个持有者**：`audit-p0-guardrails.ps1` 的 P4-C 片段、`smoke-selection-reference.ps1` §6 的 id 循环、E5-04 新建的 `check-entry-reachability.js` A–G 各查一半，合并成该门禁的规则 H（注入块字面量 id 必须有节点，语料四模块 309 个 `.vb` 剥注释全文扫，实测 16 处查询 16 个有节点）与规则 I（注入块调用的 JS 函数必须有定义，VB 词法读载荷 + 剥插值空洞，实测 124 段载荷、61 个调用名全部解析），两份副本删除并钉住「不许半并轨」（六个组成 token 必须齐）。被删的 P4-C 之所以长期零发现，是因为它把带 VB 插值的载荷整条跳过，而注入载荷绝大多数带插值。**第三条并轨把同一条「引用的目标必须存在于磁盘」用在文档地图上**（新增规则 J）：`AGENTS.md` 是 Agent 的地图，而地图坏掉不参与编译、也不是运行时资源，没有任何东西会响——HEAD 实测 11 份导航文档、24 条本地引用、**5 条指向从未存在的文件**（根 `AGENTS.md` 的 `docs/build-and-installer.md` 与 `docs/signing-and-certificates.md`、`CLAUDE.md` 的 `docs/VS2025-VSTO-Migration-Guide.md`、`README.md` 的前两条同名列），症状不是报错而是下一个读者重新考古一遍；修法一律改指向真实存在的地方（README 两节、`docs/VS2026-InstallerProjects.md`、`build/SignArtifacts.ps1`），不补空文件让门禁闭嘴，现在 38 条引用 0 条断链。范围边界同样是量出来的：第一版把裸文件名与模块内相对写法也算成引用会得到 13 条假阳性，而一条 noisy 的门禁最终只会被关掉，所以只核 markdown 链接和以顶层目录开头的反引号路径；下限 `< 30` 用来区分「地图干净」和「扫描自己失效」。扫描跑出两条真实缺陷，判定口径不同：**接通** `WordAi/ChatControl.vb:1748` 的 `showProofreadLoading()`（面板里从没这个函数，同一条注入块顺序执行、前一句生效后这一句抛 ReferenceError，于是校对轮次右上角动画从不出现）改为 `showLoadingIndicator()`，收起仍走 `changeSendButton()` 与宿主三条完成路径；**删除** `AgentKernelService.vb` 两处 `getElementById('planning-status-card').remove()`（该 id 不存在、又带 `if(...)` 守卫所以不报错，属于「声称复位规划卡、实际一条 DOM 都没动」的死码；卡片真实 id 是 `agent-card.js` 的 `'agent-plan-' + uuid`，终态复位由 `restoreAgentRequestUi()` 持有）。同样是坏注入，一条接一条删，判据不是偏好而是「谁拥有这块 DOM」。反空转：node 侧 5 个变异探针（含把 5 个注入 API 名整体改名让扫描自己变瞎、以及把根 `AGENTS.md` 指向一个不存在的 `docs/VS2026-InstallerProjectsX.md` 证明规则 J 不是只在新写文档时有效）全部变红且报错各不相同，runner 侧 3 个探针证明新门禁无需登记即被跑且退出码真能拦、发现规则坏掉时不会退化成「0 条门禁全绿」、去掉退出码判定后失败门禁会重新打出 PASS；另有不改代码只改环境的实测——把 `C:\Program Files\nodejs` 移出 `PATH` 跑套件，23 条 PowerShell 门禁跑完后在 `run-code-checks.ps1:83` 抛 `node not found`、退出码 1。最终规模：`discovered 23 gates`、46 个 PASS 步骤。已声明的不作为：**CSS↔JS 双向配对经实测没有推广**（`styles.css` 330 个类名，整词匹配后 13 个查不到，其中 12 个是 `agent-trust-risk-${riskLevel}`、`step-msg-${status}` 这类拼接前缀，只有 `.file-chip` 是真死样式；要跑成门禁就得为拼接开白名单，而白名单正是 E5-04 刚删掉的那类腐烂清单，E2-05 那条只钉本轮新增 7 个类的做法保持原样）、不给只报告的 `audit-installer-duplicates.ps1` 补失败条件（改为具名退出，让「退出」这件事可见）、不改任何一条门禁的内部实现、**规则 J 只收到导航层为止**（同一把尺子量 `docs/*.md` 的 21 份设计/审计文档：HEAD 120 条引用里 11 条断链，本轮修掉 3 条，剩 8 条写的是「计划要做」的东西——`.github/workflows/code-checks.yml` 两处、评估集产出 `build/reports/golden/latest.md`、以及 `HarnessEvents.vb`/`ComObjectHelper.vb`/`project-assessment-optimization-plan.md` 这类当时的设想；设计文档天生会描述尚未存在的文件，把它们算成坏引用会让门禁天天红，而这 8 条已作为已知账目记在 §8.4）、不假装已有 CI（`.github/` 只有 `copilot-instructions.md`，没有流水线定义，所以 `build\RunReleaseChecks.ps1` 是当前唯一有退出码的入口——这一条同时是 M1 闸门仍未闭合的原因）。
23. **W4 落地记录（E6-02/03/04 任务包合同、一致性门禁与面板归组）**：代码与自动化校验已完成，字段归属推导、三包口径、三层反空转与 6 条真机清单见 `week4-landing-report.md` §14。M1 需要用户坐在机器前并先拍板流水线承载，按「先不用做真机验收，直接接着实现」把主线推进到 Phase 2 的 E6 前三行。这条的成败由 E6 那个坑决定：**任务包不能成为第五份平行清单**（前四份是 `QuickActionCatalog`、Skill 的 `allowed-tools`、宿主 capability 注册表、`ToolRegistry`），所以落笔前逐个读回每个字段的当前持有者，只把没有主人的七项写进合同（预览策略、验证器、成功标准、交付物、解释模板、失败恢复、黄金样例引用、已有实现），schema 里连 `allowedTools`/`riskLevel`/`userPhrase` 字段都不设——不是漏了，是它们已有主人；这条刻在加载器类注释上，门禁反向断言注释仍指向 `QuickActionCatalog`，防止日后有人把话术偷偷抄回包里。落地的包只有三个：`word.review-polish`（`word.selection.proofread`+`.rewrite` → `word.proofread` → `word-document-agent`，断 `writtenCount gte 1`+`nativeUndoable isTrue`）、`excel.clean-and-fix`（`.formula`+`.format`，断 `writtenCount gte 1`+`changedCount gte 1`）、`excel.insight`（只读，断 `writtenCount eq 0`+`hasChanges isFalse`，`mode=none` 且 `requireBeforeWrite=false`）。报告 §3.2 前 6 包里剩下的 Word 模板排版、PPT 成稿、PPT 品牌模板**故意没有 JSON**：它们在 E6-01/07 之前没有可解析入口，而「没有入口的包不许声明」正是 E6-03 的验收标准，造占位壳就是 §1 明令禁止的假能力；门禁只设 ≥3 的下限，不维护「应该有几个包」的名单。**断言只能是代码真的会产生的事实**：事实表由反射 `AgentRunChangeLog` 的公开 Integer/Long/Boolean 只读属性推导（实测 7 项：4 数值 + 3 布尔），布尔只许 `isTrue/isFalse`、数值禁用它们，双向都拦；按宿主分开诚实口径——Word 断 `nativeUndoable`，Excel 写包**不断** `fullyUndoable`，因为 `ExcelAi/Undo/ExcelRunRestore.vb` 只快照 ≤4000 单元格的值/公式/数字格式并明确拒绝声称插行插列、新建图表可还原，把结构变更写成断言等于让合同替代码撒谎，「能不能撤」因此只留在 `successCriteria` 文案里。加载器沿用 `ToolRegistry.LoadFromRuntimeDirectories` 的向上走目录算法（避免出现第二套运行时视图），代价是同一次搜索会同时命中源码目录与 `bin\Debug` 输出、同一份合同被读两次，所以「id 已存在」不能直接算冲突，只有同一 `SourceDirectory` 内重复才拒绝，而那种情况已被「文件名必须等于包 id」天然排除；`Find/PackagesFor/FindByAction` 前挂 `EnsureLoaded()`，让面板不必记得先加载——不改 `BaseChatControl` 的惰性装配链，用第一次查询带一次磁盘扫描换掉「合同在磁盘上、运行时视图是空的」这类看不见的空表；被拒绝的包除 `Info` 外再写一条 `Warn` 带首条原因，因为「静默丢弃一个包」和「这个任务没有归组没有断言」是同一件事。新门禁 `scripts/check-task-package-contract.ps1`（475 行，纯静态）**不复制任何 schema**，12 张 `Allowed*` 词表和 `AllowedSchemaVersion` 用正则从 VB 源码反查、匹配不到即 throw（词表被改成看不懂的形态时门禁必须瞎在明处）；逐级解析为包→动作（`_wordActions/_excelActions/_powerPointActions`）、包→capability（Word 读 `WordCapabilityRegistry.vb`，无注册表宿主强制空数组，把 #36 变成显式约束）、包→Skill（目录必须有 `SKILL.md`、front matter 的 `name` 等于目录名、`allowed-tools` 非空）、包→黄金样例（`tests/golden/l0/catalog.json`）、包→已有实现（`reuse` 每条路径必须在磁盘上存在——这条让「先建壳」在经济上不划算，因为壳必须写它复用了谁）；外加三条注册不变式（三个 `.vb` 编译项 + `TaskPackages\**\*.json` 的 CopyToOutputDirectory 仍在 `.vbproj`、`ActionsFor` 交叉校验仍在、Models 注释仍指向目录），理由是合同再漂亮，没人加载就是装饰品。反空转做了三层，每层都能自己变红：14 个内存探针跑在两份干净基线（一写包一只读包）上且每个探针先证明基线干净、未知字段/幽灵入口/杜撰事实/布尔配 `gte`/失效 reuse/只读却要求写前确认/Skill 不在盘上/capability 填给无注册表宿主/杜撰黄金样例 id/杜撰占位符/id 与文件名不符/`schemaVersion` 漂移/`mode` 越词表全部如实变红；18 份坏样本喂给编译后的 DLL 反射跑 `LoadDirectory`，17 份当场被拒、产生 18 条拒绝消息（前缀不符那份同时命中两条规则），同批放了一份真包做对照，证明这批拒绝不是「什么都不认」——唯一没被加载器拒的是 reuse 路径失效那份，因为「引用能否落到磁盘」按分工属于门禁而不是运行时；1 次磁盘级变异把 `word.review-polish.json` 的 reuse 改成不存在的文件，整条门禁变红，随后按字节还原并核对 MD5。E6-04 的面板归组守住「零新消息 type、零新发送路径」：`CommandPaletteService.BuildState()` 把状态从一维动作表改成 `packages`（id/title/deliverable）+ `actions`（包内动作先发并携带 `packageId`，未被任何包收编的动作照旧发、不带 `packageId`），`command-palette.js` 只在 `packageId` 变化且后端给得出标题时插一行 `div.command-palette-group`，方向键游标仍是一条扁平列表（不引入「组」这个第二状态机），后端没下发 `packages` 时按全未归组渲染、标题只从后端读——前端另立包清单会产生最难发现的那份副本。`smoke-command-palette.ps1` 的 3b 段把归组钉成可失败断言：Word 宿主 ≥1 个包头、恰好 2 个动作带 `packageId`、首个归组包是 `word.review-polish`、`word.selection.summarize` 必须仍未归组（「没被包收编」和「被丢掉」是两回事）。过程中踩到的两条硬约束不是风格问题：`TaskPackageCatalog.vb` 初稿 524 行越过「单 class ≤500 行」，按仓库既有 `LoopEngine.Planning.vb` 的做法切成 `Partial Class`（目录/候选/查询 278 行，严格校验 263 行），并为此给门禁加一条「Parsing 那份必须仍声明 `Partial Class TaskPackageCatalog`」——否则一次重构就能把 schema 变成两个持有者；VB 大小写不敏感使参数名 `file`/`path` 遮蔽 `System.IO.File`/`System.IO.Path`，报出来的是 `ReadAllText is not a member of String` 这种误导性错误。最终规模（本轮重新实测）：四项目 `PASS`；`PASS task-package contract: 3 packages, 6 assertions, 15 schema keys, 7 facts, 14 probes turned red on 2 clean baselines, golden references=0`；反射跑加载器 `LOADED=3 ERRORS=0 WORD=1 EXCEL=2`；`Command palette smoke passed.`；`run-code-checks.ps1 -SkipBuild` 为 `discovered 24 gates`（新门禁被目录自动发现，未登记进 `$notGates`）、47 个 `PASS`、0 FAIL、`All code checks passed.`，Golden L0 仍 12 例。已声明的不作为：E6-05 未实现（断言目前只被静态校验合法性、还没有执行者，计划复用 `LoopEngine.Planning.vb` 的 `ValidateObservedOutcome` 而不是新建验证链）、E6-06 未做（三包 `goldenCases` 全为空数组，门禁输出里的 `golden references=0` 就是这个事实而非省略）、E6-01 未做（三端 Ribbon 仍有硬编码 Agent 请求，是下一行）、没有为 Excel/PPT 造 capability 清单（等 #36）、没有把 `riskLevel` 引入合同（风险闸门归 `SafetyGate`）、没有动 `BaseChatControl` 的运行时装配链。副作用记一笔：#36 的性质变了——两个 Excel 包的 `capabilityIds` 现在被门禁强制为空，定下来之后要同时改包和改那条空数组约束。
24. **W4 落地记录（E6-01 三端 Ribbon 话术迁入唯一清单）**：代码与自动化校验已完成，十处迁移、类型闸门的一次注入证伪、7 条探针与被顺手逮到的旧 smoke 假设见 `week4-landing-report.md` §15。这条的验收标准是「入口不再持有第二份话术」，而选区浮层/右键菜单在 E4 已经收口之后，剩下的第二个持有者正是三端 Ribbon 的按钮 Click：Word 1 处、Excel 6 处、PowerPoint 3 处，共 10 句 Agent 请求写死在宿主里（动手前按 9 处记，改完全仓 grep 才发现漏了 `ExcelAi/Ribbon1.vb:17` 抓取网页——「我记得有几处」在这类收口里不算证据）。十句原文一字未改进 `ShareRibbon/Agent/QuickActionCatalog.vb:111-176`，`Label` 取设计器里的真实按钮文字而不是重新拟的文案。**取的方式是编译期引用而不是字符串查表**：三宿主 `StartAgentFromRibbonAsync` 的形参由 `request As String` 改成 `action As Agent.QuickAction`（`WordAi/Ribbon1.vb:460`、`ExcelAi/Ribbon1.vb:229`、`PowerPointAi/Ribbon1.vb:416`），函数体只在一处用文本 `StartAgentFromHostEntry(action.RequestText)`；查表会把这条不变式降为约定，签名把它变成类型错误，并且做了一次注入证伪——把 `WordAi/Ribbon1.vb:27` 换回字面量后 MSBuild 报 `BC30311: Value of type 'String' cannot be converted to 'QuickAction'`，随后按 MD5 还原（前后同为 `79ac48da…`）并重建四项目，避免 §8 那次「DLL 比源码旧、白跑一轮」。**这十条没有进 `ActionsFor`**：那份清单喂的是选区浮层、右键菜单与命令面板，一并返回会当场改变面板行数，而 E6-01 的动词是「把话术搬走」不是「多给用户几行」，行集属于 E6-07 与开放决策 D9；这条边界在门禁里是规则而不是注释——ribbon id 出现在 palette 宇宙即失败。量到的一个副产品事实：10 条对应按钮有 9 条 `Visible = False`（8 条在宿主设计器，`TemplateFormatButton` 在 `ShareRibbon/Ribbon/BaseOfficeRibbon.Designer.vb:183`），唯一默认可见的是 Excel 数据分析，即本轮收口的是一批「已实现但看不见」的入口，D9 的代价从此更可判定（文案已只有一份）。门禁在 `check-task-package-contract.ps1:410-565` 补了围绕那条签名的闭包：签名必须仍是 `action As Agent.QuickAction`、任何 `StartAgentFromRibbonAsync("` 字面量判失败、引用成员必须已声明且前缀属于引用它的宿主、声明未用即装饰、id 必须 `<host>.ribbon.<button>` 形状、声明数必须等于引用数；7 条内存探针各自把干净状态改坏一次（字面量回来／签名退回 String／引用未声明成员／一条话术被两个按钮复用／引用别宿主话术／id 形状漂移／id 泄进 palette）。这套规则第一次跑全量门禁就逮到一条**本来就站不住的旧假设**：`smoke-selection-toolbar.ps1` 把 catalog 里所有 `word.` 前缀 id 当作选区清单，于是报「expects exactly 3, found 4」，多出的正是 `word.ribbon.analyze`；已改为按 `_wordActions` 块切片读取（与 E6-03 门禁同一个「下一个访问修饰符即块尾」的切法），恰好 3 条、必须 `word.selection.*` 这两条断言没有放松。明确的不作为：没改任何按钮可见性与点击行为（Ribbon 可见入口仍 Word 14/18、Excel 13/18、PPT 11/18），没给 ribbon 动作填 `CapabilityId`（等 #36），没把它们写进任务包 `entryActions`（包只能引用 palette 入口，这是 E6-03 定的），没做 E6-05/06（断言仍无执行者、三包 `goldenCases` 仍空）。本轮重新实测：`grep 'StartAgentFromRibbonAsync("'` 三端为 0，引用数 1/6/3 与声明数 10 相等；四项目 `PASS`；`PASS task-package contract: 3 packages, 6 assertions, 15 schema keys, 7 facts, 14 probes turned red on 2 clean baselines, golden references=0, ribbon utterances=10 at 10 button(s)`；`Command palette smoke passed.`；`run-code-checks.ps1 -SkipBuild` 为 `discovered 24 gates`、47 个 `PASS`、0 FAIL、`All code checks passed.`；`QuickActionCatalog.vb` 250 行、门禁脚本 634 行；导航文档本地引用 38→42 条（AGENTS.md 本轮加了两行导航）且全部解析到磁盘。
25. **W4 落地记录（E6-05 包断言在 run 收尾时真的核对）**：代码与自动化校验已完成，链路闭包、15 条运行时用例、14 条源码探针与 `onFail: repair` 的缺口见 `week4-landing-report.md` §16。E6-03 之后那 6 条断言只被静态校验「填得合法」，运行时没有任何一处把台账和它们比一遍——这不是差最后一步接线，而是这类合同最常见的死法：字段齐全、门禁全绿、无人读取。所以本轮验收写成「假成功会被拦下来」，并且拦下之后终路径仍只有一条。前置条件是此前根本不存在的一件事：**这一轮从哪个入口动作来**。断言属于包、包靠 `entryActions` 绑动作，一轮 run 若不知道自己属于哪个包，就没有任何东西能证明它该受哪个合同约束；从「帮我把这段润色一下」反查 `word.selection.proofread` 是一次 NLP 分类，猜错的后果不是「没验证」而是**把别人的成功标准扣到用户头上**（用户要洞察，系统按写包判「你一个字没改，未完成」）。沿用 E6-01 的「用类型而不是约定」：`StartAgentFromHostEntry` 的形参由 `request As String` 改成 `action As Agent.QuickAction`（`BaseChatControl.vb:1769`），话术和 id 同源同对象；payload 带上 `entryActionId`（`:1773`）并在 `:1806` 读回，下游 `HandleStartAgentCore`/`HandleStartAgentWithFilesAsync` 的该形参**不是 Optional**（`:1859/1982`），漏传是编译错误而不是靠审查，自由文本与聊天输入路径显式写 `""`（`:1731-1732`）表示「我知道我在问什么，这条没有合同」；三宿主 7 处调用点（Word 3/Excel 2/PPT 2）全部交对象，`grep 'StartAgentFromHostEntry(request'` 全仓 0 处；面板仍只发一条 `startAgent`（`command-palette.js:218`），`__natural_language__` 映射成空串（`:217`）否则占位符会被当真 id 去查包。链路 `AgentKernelService.vb:106`→`HarnessModels.vb:26`→`OfficeHarness.vb:127`→`AgentKernel.vb:208/218`→`AgentSession.vb:22`。两个坑单独记：UI 线程赋值与后台 `StartAgentAsync` 之间隔着线程切换，所以身份按每次调用显式携带而不是只放服务字段，`Nothing`（调用方忘传，写一条 Warn）与 `""`（自由文本本就无合同）是两种事实、合并会把「漏接的入口」悄悄洗成「合法的自由文本」；`FinalizeAgentUi` 的 Finally 与 `AbortAgent` 都复位该字段，否则下一轮继承上一轮的包。**变更事实仍只有台账能报**：验证器最容易长成第二份真相的地方就是「自己数记录」，所以读取能力放回持有者——`AgentRunChangeLog.TryFactValue`（`AgentChangeLog.vb:162`）反射读本实例公开、非索引、Integer/Long/Boolean 只读属性，词表外与 PascalCase 原名都读不出；命名规则也只有一个持有者，`ToFactName`（`:179`）同时供 `TaskPackageCatalog.BuildVerifiableFacts`（`:89`）使用、目录里那份私有 `ToCamel` 删除，否则门禁反向推导的词表与运行时能读的名字会各自漂移，漂移的表现正是「这条断言永远读不到事实」。新类 `TaskPackageOutcomeVerifier.vb`（104 行，进 `.vbproj` 编译项）判定一律 fail-closed：事实名空/读不出/比较符越词表/数值比较遇布尔/布尔判断遇数值/本轮无台账，每种都返回一条**点名是哪个事实**的原因，文案带声明值与包里的 `why`，让用户知道「包本来要求什么」而不是一句空泛的未完成。**位置**：`ValidateExecutionOutcome`（`LoopEngine.Coverage.vb:162`）本来就是「要求插图却没插、要求 N 页只做 M 页」的最后一道闸，包断言是同一类判断，接在它末尾（`:189-191`+`:199-213`）并共用 `LoopEngine.vb:479-485` 那条 `AgentResult.Failed` 出口，不新建第二套 verifier/loop/终态；返回值仍走 `AppendWriteSummary`，「没达成什么」和「已经改了什么」要一起给用户。反空转两层：**运行时**新门禁 `smoke-task-package-verifier.ps1`（172 行，被 `scripts\*.ps1` 自动发现）反射加载 `ShareRibbon.dll`，用 4 份真实 `AddFromStep` 载荷造台账，15 用例同时钉「满足→0 条原因」（正常写入/保持只读/写过且有原生撤销点）与「不满足→点名事实」（回滚后归零、只读包发生写入、无撤销栈、无变更、6 条 fail-closed），并把 7 条事实逐个 `TryFactValue`；只测能变红的验证器和只测能变绿的门禁一样不可信。**源码级**`check-task-package-contract.ps1:566-778` 对 12 个文件跑约 20 条规则、14 条变异探针，每条必须变红且必须被对应规则抓住（会话丢字段／harness 丢参数／内核从话术猜身份／宿主入口退回 String／聊天路径继承上一轮身份／面板不发 id／面板把占位当真 id／收尾跳过包／验证器自己数事实／数值分支永不满足／布尔分支永不满足／目录自带命名规则／`TryFactValue` 放开类型闸门／验证器不进编译项），外加 `entryActionId` 出现次数下限（≥6，实测 10）防「规则全对但链路已断」。这套规则第一次跑就自己红了一次：`identity probe 'verifier-never-satisfied' slipped through`——初版 `verifier -notmatch 'Return ""'` 在数值分支的 satisfied 被改掉后仍然绿，因为布尔分支还留着另一条 `Return ""`；一条 grep 挡不住「两个 satisfied 分支坏掉一个」，已分别钉住两条分支字面量并各配一条探针（13→14），与 §14.6 那次「加载器什么都拒」是同一类错误。**明确未达成**：`onFail` 盘上实测是 `excel.clean-and-fix`/`word.review-polish` = `repair`、`excel.insight` = `report`（写这段时我记成「三包都 repair」，那条错误在 E6-05b 的门禁前置检查里被真实包数据驳回后才更正，见 §18.4），而 run 收尾当时不能重进规划循环——`LoopEngine.vb` 步进循环约 360 行，end-of-run 之后再包一层 re-plan 会牵动 replan 幂等（E1-06）与写入摘要（E1-04），本轮没有真机在手，硬写等于拿用户文档试错；处置是拦下假成功并写一条明说 `run-end repair route is not wired` 的 Warn，接通路由记为 **E6-05b**。计划里「失败进入既有 repair loop 而不是直接宣称成功」的后半句达成、前半句没有，把它写成达成就是这份记录最不该犯的错。两条代价如实记下：自由文本发起的同一件事不受任何包合同约束（面板输入不核对、右键菜单核对），这是「不猜包」换来的确定性；`writtenCount gte 1` 会把「要求润色但确实无需改动」这类合法空写入判成未完成——合同没说谎，但这可能不是用户想要的答复，真机清单第 2 条专测其观感，结论决定 E6-06 失败样例怎么写。本轮重新实测：`PASS task package verifier: 3 real packages, 15 assertion cases, 7 facts readable at runtime`；`PASS task-package contract: 3 packages, 6 assertions, 15 schema keys, 7 facts, 14 probes turned red on 2 clean baselines, golden references=0, ribbon utterances=10 at 10 button(s), 14 entry-identity probes turned red at 10 panel site(s)`；`run-code-checks.ps1 -SkipBuild` 为 `discovered 25 gates`（24→25）、50 个 `PASS`、0 FAIL、`Golden L0 passed: 12 cases`、`All code checks passed.`；四项目 `PASS`、`Code projects built successfully.`；入口行集未动（Ribbon 可见入口仍 Word 14/18、Excel 13/18、PowerPoint 11/18），`BaseChatControl.vb` 3370 行的超限是历史欠账、本轮只加参数与注释。
26. **W4 落地记录（E6-06 黄金样例：每包一份满足、一份不满足）**：代码与自动化校验已完成，口径表、11 条静态探针、两次盘上真实变异与顺手逮到的旧合同缺陷见 `week4-landing-report.md` §17。这条的动词不是「多跑几个 case」而是「把包的语义钉到磁盘上」：E6-05 之后断言会执行，但执行器的正确性只由我为验证器造的那 15 条合成载荷证明——如果包自己的口径写错（例如把「发生过的写」当成「留在文档里的写」），验证器越准就越忠实地核对一个错误的合同。落 7 条 `task_package_outcome` 用例（`catalog.json` 12→19 例、92→247 行）覆盖**磁盘上真实存在的全部 3 个包**：`word.review-polish` 1 成功 + 2 失败、`excel.clean-and-fix` 1+1、`excel.insight` 1+1，每条带 `packageId/entryActionId/expectedOutcome/expectedUnmetCount/expectedUnmetFacts/canUndo/steps[]`，`observation` 用宿主 `ToolResult` 的原始形状（`kind`/`changed`/`undoUnits`/`rollback.succeeded`/`before|after.documentTextHash`），不新增验证器读不到的字段——那等于把 §16.7 禁的「包自己数事实」换个地方复活。三条失败样例各钉一种最容易糊过去的口径：`changed:false`+首尾指纹相同 ⇒ `writtenCount`+`nativeUndoable` 双未达成（§16.7 第 7 条代价的可执行版本）；写过但 `rollback.succeeded:true` ⇒ **只有** `writtenCount` 未达成而 `changedCount` 仍 1，这是本轮最关键的一条边界，若断言读前者，用户点过撤销的 run 会被合同判成功；写了但无 `undoUnits` ⇒ 只报 `nativeUndoable`，证明布尔事实与写次数不互相掩盖。只读包的通过样例只有 `kind:"read"`+`recorded:false` 一步，让验证器面对**空台账**而不是「有记录但都不算写」。**数据必须两边都真**：`run-golden-l0.ps1:25-50` 用真实 `AddFromStep` 造台账并双向核对收录口径（该收却被拒 ⇒「写观察合同漂移」；不该收却收了 ⇒「只读的一轮会谎称改过文档」），`:53-96` 由 `FindByAction(entryActionId)` 解析包、比对 `package.Id`，再对齐 `UnmetReasons` 的条数/极性/点名事实，`expectedOutcome` 第三种取值直接 throw。覆盖块 `:275-297`（每包 ≥1 pass + ≥1 fail 且包声明条数==catalog 归属条数）**第一次跑红就逮到一个真缺陷**：报 `excel.clean-and-fix declares 0 golden case(s) but the catalog holds 2`，根因是 `TaskPackageCatalog` 的 id 去重「先加载的那份赢」，而向上目录搜索同时命中源码与 `bin\Debug\TaskPackages` 的旧副本——一份我没读过也没写过的旧 JSON 正在冒充当前合同，所有基于合同的判定都在绿灯下核对错的东西。这不是「记得先 build」的纪律问题，修在持有者：`TaskPackageCatalog.vb:71/236/247-248/254-255/181` 记下每个 id 首次加载的规范化正文，跨目录同 id 正文不同即 `AddError`，`Reset()` 一并清空；门禁反向钉住这两条不变式（`check-task-package-contract.ps1:560-564` 匹配不到即 throw），防止下次重构把它当冗余删掉。反空转三层：包侧 `:326-347` 要求每包同时有 satisfied 与 unmet 用例（「没有东西能证明合同可能被满足／可能变红」这两种空转各拦一次）并核对引用与归属条数相等；用例侧 `:425-509` 判 12 类（重复 id／幽灵包／入口不属于该包／极性和事实互相矛盾／事实名词表外／缺 `canUndo`／steps 空／缺 `toolId` 或 `observation.kind`／`recorded` 与 `kind` 打架／收录步骤缺 boolean `changed`）；11 条内存探针各对应一条规则，另加 2 次**盘上**变异——只改源码包不重构建 → runner 红在跨目录内容不一致（MD5 还原后复绿），改一条 `expectedUnmetCount` → runner 红并点名两条事实。第二次变异暴露了我自己的错：还原用了「把第一个 1 换回 2」的反向字符串替换，误伤另一条本来就期望 1 的用例，又被 runner 当场抓出，最终按行号精确还原——**变异探针的还原必须按整文件 MD5 或行定位，反向 replace 不算还原**，这条 §14.7 早有规矩，我绕开它换来的是往测试数据里注入错误期望值。**如实未达成**：计划行叫「前 6 包」，磁盘上只有 3 包，Word 模板排版／PPT 成稿／PPT 品牌模板在 E6-07 与 D9 之前没有可解析入口，而「没有入口的包不许声明」是 E6-03 定的验收标准，先造壳再造样例等于把那条撤掉——这行的完整验收随 E6-07 收。也没有新增门禁文件（25 gates 不变，规则全加进已有两份脚本，避免门禁清单自己变成第二份真相）、没把样例落库、没为 `observation` 另造 schema 校验器。两条代价：`canUndo` 必填，写用例的人必须显式表态撤销栈有无，好处是非撤销宿主不会因默认 true 假绿；runner 现在拒绝加载任何跨目录漂移的包，即「改包不重建就跑 gates 会红」，若有人嫌烦去删加载器那道检查，漂移会退回静默——该改的是流程不是闸门。本轮重新实测：`Golden L0 passed: 19 cases`（12→19）；`PASS task-package contract: 3 packages, 6 assertions, 15 schema keys, 7 facts, 14 probes turned red on 2 clean baselines, golden references=7 (7 outcome cases, 11 golden probes turned red), ribbon utterances=10 at 10 button(s), 14 entry-identity probes turned red at 10 panel site(s)`；`PASS task package verifier: 3 real packages, 15 assertion cases, 7 facts readable at runtime`；`run-code-checks.ps1 -SkipBuild` 为 `discovered 25 gates`、50 个 `PASS`、0 FAIL、`All code checks passed.`；四项目 `PASS`、`Code projects built successfully.`，构建后三份包 JSON 的源码与 `bin\Debug` 副本 MD5 两两相等；入口可达性 56/66 消息 type、DOM id 103/103、注入调用 61/61、导航引用 46 条全部解析、`Oversized baseline files: 22`（未新增超限）。真机仍按「先不用做真机验收」暂缓，§16.8 那 8 条未跑，E6-08 依赖它。
27. **W4 落地记录（E6-05b 未达成的包真的回到既有 repair 路由）**：代码与自动化校验已完成，闸门位置、观察形状、6 条运行时用例、15 条新静态规则/探针与 1 次盘上变异见 `week4-landing-report.md` §18。这一行补的是 §16.6 如实记下的那半句缺口，而它的成败不是「会不会修」，是**修这条动作是否仍由既有的那一个循环、既有的那一份预算、既有的那一条终态出口持有**——最省事的写法（循环结束后再套一层重规划）正是仓库明令禁止的第二套 Loop/终态/幂等口径。做法是把 `LoopEngine.vb:103` 的 `While` 条件从 `stepIndex < Steps.Count` 扩成 `(stepIndex < Steps.Count OrElse Not runEndGated)`，计划跑完时循环**再进一圈**，那一圈只走合同闸门（`:110-134`）：四个前提缺一不可——本轮执行过、**没有任何未落地也未跳过的步骤**（失败步骤归 §14 的步内修复管，同一问题不得两个循环都来修）、包声明 `onFail=repair`、还剩 replan 预算（`replanAttempts < MaxReplanAttempts`，常量仍是 2，**与步内修复共用同一份计数器而不是第二个计数器**）；接受新计划才做四件事：`replanAttempts += 1`、换计划、`session.Writes.BeginNextPlan()`（E1-06 幂等：上一版已落盘的写入不在新版里重跑）、`stepIndex = 0` + 重新武装闸门，再 `RaisePlanGenerated` 把第二版交给面板；任何一条不成立就 `Exit While` 落回原收尾顺序，闸门与终态判定改用同一个 `IncompletePlanSteps`（内联 LINQ 已删，`:497`），`AgentResult.SuccessResult` 全文件计数仍为 1。判定与请求全部留在 `LoopEngine.Coverage.vb`：`MatchedTaskPackage`（自由文本/无包即 `Nothing`）、`UnmetTaskPackageReasons`、`PackageWantsRunEndRepair`（策略**读自包数据**）、`RequestRunEndRepairPlanAsync`（把未达成断言拼成带 ASCII 前缀 `RUN_END_CONTRACT_UNMET {包id}` 的观察文本，然后**整段交给既有 `ReflectAndReplanAsync`**，因此自动继承 E1-06「已落盘写入作为系统级约束下发」；只接受 `replan`，`retry`/`skip` 在 run 末没有当前失败步骤可落地，一律 Warn 并按未达成报告；新计划还要过 `ValidateReplannedCoverage`）、`IncompletePlanSteps`。一条 VB 语言约束改变了形状：**Async 函数不能有 `ByRef` 参数**，所以 helper 只返回 `ExecutionPlan`/`Nothing`，预算加减留在调用方，持有者仍是循环本身。反空转三层：**运行时**新门禁 `scripts/smoke-run-end-repair.ps1`（259 行，被 `scripts\*.ps1` 自动发现）反射加载真实 DLL，用真实 `LoopEngine` + 脚本化模型/执行器跑 6 条用例——未达成→恰好 1 次修复→成功且 `PlanGeneration=2`、`WrittenCount=1`；反复未达成停在共享预算且仍报未达成；自由文本 0 次反思；合同已满足 0 次反思；`onFail=report` 的包违例 **0 次反思**且失败点名 `excel.insight`；修复轮拿不到计划则 1 次反思后如实报未达成。**静态**`check-task-package-contract.ps1` 的身份语料加入 `LoopEngine.vb`（12→13 个文件），新增 15 条规则与 15 条一一对应的变异探针（闸门可达、失败步骤在场不触发、策略读自包、预算共用、必须推进计划版本、必须重新武装、无计划必须退回未达成、**终态出口计数必须为 1**、必须复用 `ReflectAndReplanAsync`、本文件不得出现 `SendAIRequest(`、只接受 `replan`、必须过覆盖校验、自由文本不绑包、rollback 未接必须仍被如实记录），入口身份探针 14→29 条且每条仍须被自己那条规则抓住。**盘上**把 `:118` 的预算判断改成恒假，四项目照样 `PASS`（这条路由编译器管不了），smoke 当场红 `unmet package must ask for repair exactly once, got 0`，按 MD5 `3d089d96…` 还原后复跑四项目与两条门禁全绿——证明路由真在被执行而非只被 grep 读到。两处我自己的错必须留档：(1) §16.6 与第 25 条写的「三个包的 `onFail` 都是 `repair`」是**假的**，盘上是 `excel.clean-and-fix`/`word.review-polish`=`repair`、`excel.insight`=`report`，发现方式是新 smoke 的前置检查按真实包数据驳回了我（门禁把我记错的事实拦在运行时用例之前），用例 5 顺势改成反向证明「策略读的是包不是硬编码」，两处文档与项目记忆已更正——这是「我记得有几处/是什么」第三次不是证据；(2) 复核 §17.8 发现我自己用 python 写文档时把三条路径打进了控制字符（`scripts`+CR、`ShareRibbon`+0x08），已按字节修复并全仓扫描归零，且**规则 J 没抓到它**（含控制字符的引用不匹配任何模式，断链检查直接看不见），这条盲区如实记录而不是宣称已覆盖。写 smoke 踩到的两个坑同样是合同级的：react 提示词发的是 `{"action":{"tool":…}}` 且不带工具清单，我按 `"toolId"` 派发就把 react 吞进步骤修复分支（`Reflections=0`），E1-06 那份用的是兜底分支；`ToolRegistry` 会校验必填参数，`DataAnalysis` 少 `sourceRange/type` 返回 `[ARGUMENT_ERROR]`，工具参数必须真实给出，否则台账收录口径是我编的。**明确未达成**：`onFail=rollback` 仍未接通（它需要 run 末把改动交回宿主还原，不属于这一行），本轮只保证它继续被如实报告并加了探针防「悄悄当成已处理」；自由文本仍不绑包也不享受修复；步内已耗尽 2 次 replan 的 run 在 run 末不再修复，直接报未达成。代价：修复轮多花一轮模型调用与一整圈重跑，`MaxIterations` 照旧消耗，用户会看到「一次任务两版计划」，观感由 §18.8 第 7 条真机判定；修复计划只按 spec 产出覆盖判定，包断言要到下一圈 run 末才再核对，所以「修了仍不满足」靠用例 2 而不是静态规则保证。本轮重新实测：`PASS run-end repair: 6 contract cases over the real LoopEngine`；`PASS task-package contract: 3 packages, 6 assertions, 15 schema keys, 7 facts, 14 probes turned red on 2 clean baselines, golden references=7 (7 outcome cases, 11 golden probes turned red), ribbon utterances=10 at 10 button(s), 29 entry-identity probes turned red at 10 panel site(s)`；`run-code-checks.ps1 -SkipBuild` 为 `discovered 26 gates`（25→26，新 smoke 自动发现）、52 个 `PASS`、0 FAIL、`Golden L0 passed: 19 cases`、`All code checks passed.`；四项目 `PASS`（变异还原后复跑一次仍 4 个 `PASS`）；规模 `LoopEngine.vb` 619→646、`LoopEngine.Coverage.vb` 325→382、门禁脚本 1033→1144 行，`Oversized baseline files: 22`（未新增超限）；入口可达性 56/66 消息 type、DOM id 103/103、注入调用 61/61、导航引用 46→48 条全解析、Ribbon 可见入口 Word 14/18、Excel 13/18、PowerPoint 11/18（未动行集）。至此 E6-05 那行的两半都有了实现，真机仍暂缓：§16.8 的 8 条与 §18.8 新增的 7 条未跑，E6-08 依赖它们，E6-07 仍卡 D9。

28. **W4 落地记录（E3-05 隐私同意：把「有个字段」变成「真有人拦」）**：代码与自动化校验已完成，出口清单、拒绝形状、9 组静态规则/22 条探针与 3 次真实离线拒绝见 `week4-landing-report.md` §19。开工先量 `HEAD`：`PrivacyPolicyAccepted` 全仓只有 6 处，3 处是存储本体、3 处是向导第 1 步，**没有任何发送链路读过它**，三个 `Allow*` 兄弟字段更是全仓 0 引用——这一行缺的从来不是界面，是 enforcement。设计由两条测量决定：`SendStreamRequestAsync` 只有 1 个调用方（`BaseChatControl.vb:2999`），拦在 `Send` 即够；而 `AiGateway.SendChatAsync` 有 5 个调用方（意图识别 ×2、补全、记忆沉淀，外加已在闸门内的 `AgentKernelService.vb:503`），它们绕开面板漏斗，所以**同意判定必须放在出口而不是入口**。做法是把出口先列出来：`HttpRequestMessage(HttpMethod.Post` + `.PostAsync(` 全量扫得 **16 个出口 / 12 个文件**，逐个定口径登记进 `scripts/smoke-privacy-consent.ps1` 的 `$outboundInventory`——9 个文件自己问（`AiGateway.vb:61`、`BaseChatControl.vb:2253` 的非流式发送器、`UtilsService.vb:265`、`LLMUtil.vb:96`、`AutocompleteService.vb:145`、`OfficeCompletionService.vb:121/:171`、`EmbeddingService.vb:94`、`DocumentTranslateService.vb:420` 及其 Excel 复制品 `ExcelDocumentTranslateService.vb:340`）；`HttpStreamService.vb` 的 3 个 POST 按 `coverage=funnel` 豁免，前提是它**只剩一个公开发送入口**，这条前提本身被写成计数断言并有探针；`ModelConnectivityTester.vb` 按 `coverage=probe` 豁免（请求体是字面量 `{"content", "Hi"}`、不带用户内容，且它是同意之前用来验证接口的唯一工具，拦它等于没法配置）——**这一条正是规则 9 上线第一次运行当场逮到的未登记出口**；`StreamJsonRpcMCPClient.vb` 的 2 个 POST 按 `coverage=mcp` 登记（用户自加的工具服务器，归 W2 的 MCP 授权面），写进清单而不是留白是为了让「以后有人把文档内容塞进 MCP 参数」必须回来重新拍板。清单的形状就是防漂移的形状：语料扫到的出口数与清单对不上即红（新出口未登记、登记了却不再发东西、某文件出口数变化），三条探针分别抓这三种。**拒绝形状沿用各出口既有的失败约定，不发明第三种终态**：面板两条漏斗推卡片（`IsBlockedByPrivacy`，`:1812` agent / `:2920` chat），`SendAndGetResponseAsync` 返回空串（正是它 `Catch` 的既有值），`UtilsService`/`LLMUtil` 用各自的空串与 `错误: ` 前缀，补全返回空列表、embedding 返回 `Nothing`（调用方本就降级），**翻译抛异常**——因为它的所有调用方按异常处理并保留原文，返回空串只会变成 `JObject.Parse` 的语法错误、把一次隐私拦截伪装成「模型返回坏了」；理由串与日志由 `PrivacyConsentGate.RefusalReason` / `RefuseQuietly(caller)` 单点持有。UI 侧坚持不建第二套：卡片复用 E3-03 的 `ModelConfigPrompt` + `showModelConfigRequired`，只加 `action` 判别（`open_config` / `accept_privacy`，`chat-manager.js:428` 分支），**零新增面板消息 type**（写回用既有 `onboardingSetPrivacy`，`accepted` 本就允许 false），`BuildShowScript` 因此改为 `Shared`——代价当场显现：`smoke-model-config-gate.ps1:69/:89` 用实例调它，PowerShell 报 `MethodNotFound`，改静态调用才复绿。可撤回走面板设置弹窗新增的一行「数据边界同意」（0 行新 CSS）：勾选框初始 `disabled`、只由宿主推送解锁，打开设置时若无状态则补问 `getOnboardingState`，取消勾选即撤回且下一条请求重新拦；**不做**「同意后自动重放那条请求」，因为卡片写的就是「这一轮没发出去」。文案持有者唯一：三条边界说明写在 `PrivacyConsentGate.BoundaryItems`、随 `privacyBoundary` 下发，门禁**从 `PrivacyConsentGate.vb` 正则读出这三句再全仓计数、要求恰好 1 次**，所以改文案不用改门禁、抄一份承诺才会红。本轮重新实测：新文件 62 行、新门禁 527 行（自动发现后 `discovered 27 gates`）、静态 9 组规则 + 语料 330 文件 + `.vbproj`、22 条探针全红、3 次真实离线拒绝 `RefusedMs=31`/`UtilsRefusedMs=7`/`LlmRefusedMs=4`、`ConsentMs=2034`；`OK: consent gates every registered cloud exit…`、`run-code-checks.ps1` 55 个 `PASS` / 0 FAIL / `Golden L0 passed: 19 cases` / `All code checks passed.`、四项目 `PASS`；入口可达性消息 type 56（未增）、router 66、DOM id 103→105/105、注入调用 61/61、导航引用 48/48、Ribbon 14/13/11，`Oversized baseline files: 22`（未新增）。写门禁自己踩的四个坑都是「测试先错」，留档：`return ,$problems` 把空数组读成 1 个缺陷；探针改 `Fail(...)` 文案没红（同一句话在上一行 `Debug.WriteLine` 里还在），规则 5 因此改成**位置合同**（守卫必须早于 `_httpClient.SendAsync`）；`BuildShowScript(Nothing)` 返回 `String.Empty` 而非 `Nothing`，`$null -ne` 断言写反；`$outboundInventory` 尾逗号在 PowerShell 5.1 直接语法错。**明确不作为**：默认值 `false` 意味着老用户升级后会被拦一次（这是同意的定义，但确是可见行为变化）；后台链路拒绝时没有卡片、只表现为「没有结果」+ 一行日志；三个 `Allow*` 装饰字段本轮没动（要先决定它们是不是产品能力）；翻译/ExcelDna 仍各自持有 HttpClient，把它们收进 `AiGateway` 是 `ShareRibbon/AGENTS.md:46` 的应有改造，另立一行；真机 8 条（§19.9）全部未跑，E3 的录屏证据仍欠。
29. **W4 落地记录（E3-04 费用/额度提示：先有唯一口径，再谈提示）**：代码与自动化校验已完成，五条 HEAD 测量、三层持有者、四键五处链接、10 组规则/18 条探针与门禁自己的六次红见 `week4-landing-report.md` §20。开工前逐文件 `git show HEAD:` 量回来的是四条事实加一条我自己的错：`ChatStateService.AddTokens` 在 HEAD **只有 1 个调用方**且拿的是 `LastTokenInfo.Value.TotalTokens`（输入/输出分项在进账那一刻就被丢掉，之后没人能按两种单价算钱）；`BaseChatControl` 里 `usage` **0 次**、`AiGatewayResponse` 根本没有 `Usage` 字段（非流式聊天与 Agent 一轮从来没记过用量）；面板只有 `processStreamComplete(uuid, totalTokens)` 两个参数，`消耗token：N` 是当时唯一用量出口且仅存在于流式路径；`pricePerMillion*`/`sessionTokenBudget` 全仓 **0 次**——费用与预算在产品里不存在。第五条是我原本的记忆：**HEAD 已经有两份各自漂移的 provider 解析**（`HttpStreamService.vb:387-391` 的 `CInt(jsonObj("usage")("prompt_tokens"))` 与 `:834` 的 Anthropic 读法），且服务端只回 Anthropic 键名时前一份会当场抛。这五条把这一行的重心从「加一块界面」改成「先让用量只有一份口径」——否则补上去的提示只是给两份会漂移的解析和一只只装得下整数的计数器各写一种说法。收口成三层，每层只有一个持有者：**解析** `ShareRibbon/Services/Ai/TokenUsage.vb`（新，76 行，全仓唯一允许拼 provider 键名的文件；三条规则各对应一次真实风险——读不到返回 `Nothing` 而不是 0，因为 0 会被读成免费；provider 的 `total_tokens` 优先于 `in+out`，因为推理模型把 reasoning token 计入 total、自己相加会低估；负数按 0 收但不丢弃整条用量）；**账本** `ChatStateService`（`_sessionUsage`/`_runUsage` 两个字段、`BeginRunUsage()` 开轮、`RecordUsage` 是唯一写入口，`AddTokens` 已删除；`CurrentSessionTotalTokens` 改为 `_sessionUsage.TotalTokens` 的只读投影，让读数与记账不再是两件事，有专门探针把 `Return _sessionUsage.TotalTokens` 改成 `Return 0` 验证）；**文案** `Controls/Services/UsageCostReporter.vb`（新，80 行，金额句/未配置句/无用量句/额度句/两条缺价提醒/输入框哨兵渲染各一处出处）。`AiGateway` 只多两行（`Usage As TokenUsage` + `FromResponse`），流式两处手写解析换成 `FromJson`，顺带填掉那个「只回 Anthropic 键名就抛」的旧缺陷。本轮最该被质疑的「第二套」是我自己刚造出来的：`Controls/Models/HistoryMessage.vb` 里既有的 `Public Structure TokenInfo`（三个 Integer）与新 `TokenUsage` 完全同义，留着它就变成「解析 TokenUsage → 拷进 TokenInfo → 结算再拷回」的往返搬运，所以直接把 `LastTokenInfo` 换成 `LastUsage`、删除 `TokenInfo`（全仓 grep 归零），并加一条规则 + 一条探针禁止这个名字回来。投递点仍然只有两个：Agent 终态卡片 `BuildUsageBlock`（`AgentKernelService.vb:499`，块挂在 `:468` 即**第一个 `Return descriptor` 之前**，所以只读的一轮也照计费）和流式脚注（沿用既有 `processStreamComplete` 通道**多带一个转义参数**，**零新增面板消息 type**，56 个不变）。钱只在用户配过单价时出现，三条硬口径：`-1` 是「没表过态」的哨兵、**半个价不算价**（`IsPriced` 要求两侧同时 ≥ 0，用 0 补另一半等于凭空造一个免费的一侧）、哨兵不显示成数字（四行经 `setNativeEntrySetting` 种子化后渲染成空输入框）、**清空一行 = 停止估算而不是沿用旧价**（`ReadNumberSetting` 带键非数字返回 -1，`HandleSaveSettings` 一律从载荷读回，延续 E4-03「开关必须真的能关掉」那条规矩）；`LoadSettings` 这四个键各自 `TryParse` 失败退回哨兵且**不进外层 Catch**，否则一个手抖填错的单价会把 `enableMcpList` 之后每一项读取一起带走——那是 E4-03 修过的同一类毛病。额度是提醒不是闸：门禁判 `UsageCostReporter` 里出现 `Throw`/`RefuseQuietly`/`IsBlockedByPrivacy` 为红，并把 `IsBlockedByPrivacy(ModelConfigGate.` 计数钉成 3，**能拦请求的只有同意（E3-05），不是价格**。§20.7 是本轮最值钱的一节，因为六次红全是「测试先错」：语料漏 `.vbproj` 让探针拿到 `null`；规则 2 用裸键名匹配，第一次基线就抓到 `HttpStreamService` 的**解释性注释**里写着那两个键名，说明规则钉错了对象——要防的是「再写一份解析」这个**形状**（`"<键名>")`）而不是词汇，那条注释同时改掉以留下教训；规则 4/5 原本只扫 `.js`，扩到 `.html` 才当场抓到用户可见的那句「不估算费用」在设置提示里（`chat-template-refactored.html:371`），即金额短语在盘上已有三份副本，只扫 JS 这条漂移永远不红；规则 8 要求 reporter 与 `processStreamComplete(` 同行而真实实现是两行；规则 9 把 `Protected Function IsBlockedByPrivacy(...)` 定义也算成拦截点（4≠3）。另两条**旧断言本身是空跑**，写第一版时就地废掉：「保存不抹掉既有设置」原本 grep `selectedCellChecked`——那个键是 `SaveSettings` 自己写的，永远存在，等于没测，改为先注入外来键 `smokeForeignSetting` 再要求保存后仍在；「坏值退回最后一次好值」原本拿静态属性当基线，`LoadSettings` 整个崩进 Catch 断言也必绿，改为写盘前清空四个静态、同一份文件里改坏一个键同时改好另一个（budget=777），坏键落哨兵而 777 必须真读回来。真实配置往返用 `%USERPROFILE%\Documents\OfficeAiAppData\office_ai_chat_settings.json`，按字节备份、`Finally` 还原后再比一次 MD5。**明确不作为**：只做「发送后实际」，**没有做发送前预估**——completion tokens 在发出前无法知道，预估出来的金额正是本节一直否决的那类数字（会被当成账单），验收标准写的是「预估或实际」所以这是取舍不是漏项；没有真实账单（单价手填、不抓取、不按模型分档）；**没有落库每轮用量**（`agent_run` 未新增 usage 列，因此跨重启累计与按日曲线不存在，理由是不为一行提示发明迁移和第二个真相源）；后台链路（补全/embedding/翻译/意图识别）仍各自持有 HttpClient 且**不进账**，并入 #41（E3-05b）；`BaseChatControl.vb` 3443 行的超限是历史欠账，本轮只 +51 行，`Oversized baseline files` 仍 22。本轮重新实测：新门禁 `scripts/smoke-usage-cost.ps1` 553 行、`ProbesRed : 18`、10 组静态规则、语料 336 个 `.vb/.js/.html/.vbproj`、四键 × 五处链接 24 条断言；真实 DLL 取值 `OpenAiTotal=150`/`AnthropicTotal=100`/`PricedCost=估算费用 $8`（1M/1M、单价 2 与 6）/`UnpricedCard=100 tokens · 未配置单价，只统计 token 用量`；`.RecordUsage(` 调用点 3、`AddTokens` 语料 0、`TokenInfo` 语料 0；改动 16 个跟踪文件 `+270/−60`；`run-code-checks.ps1 -SkipBuild` 为 `discovered 28 gates`（27→28，新门禁靠 `scripts\*.ps1` 自动发现、未进 `$notGates`）、54 个 `PASS`、0 FAIL、`Golden L0 passed: 19 cases`、`All code checks passed.`；四项目 `PASS`（删掉 `TokenInfo` 后复跑仍全绿）；DOM id 105→109/109、注入 id 20/20、注入调用 61/61、导航引用 48→55/55、Ribbon 14/13/11（未动行集）。文档侧把口径写进三处 `AGENTS.md`（根反跑偏清单第 20 条 + 导航表、`ShareRibbon/AGENTS.md` 表、`ShareRibbon/Controls/AGENTS.md`「三层各只有一个持有者」）。真机 10 条（§20.9）按「先不用做真机验收」未跑，其中第 2、5、7、10 条是这一行能否算达成的关键——分别验「没配价时屏幕上不出现金额」「清空即停止估算」「超额度不拦请求」「同意拦截不因本轮改动而漏出 token」。

30. **W5 落地记录（E10-01 失败步骤的单步纠偏 + E10-04 门禁）**：代码与自动化校验已完成，落点、5 条代价声明、24 条变异探针与 12 条真机清单见 `week5-landing-report.md` §1/§2/§4。口径是**给用户一个入口，而不是再造一套恢复循环**：新收件箱 `Agent/AgentStepInterventions.vb` 按步号寻址、同号覆盖、`Take` 即失效，词表沿用反思分支的 `retry/skip`（决定权从模型换回用户，不开第二套说法）；唯一的门是 `AgentKernel.QueueStepAction`（拒空会话 / 别的 run / 三种终态 / 未知动词），唯一的消费者是 `LoopEngine.ApplyUserInterventions`——全仓 `Take(` 恰好 1 次、`AgentKernel` 不写任何步骤状态，因为宿主或面板直接改 `PlanStep` 会造出第二套执行真相。「只重跑受影响步骤」是结构而非承诺：回退只把目标步骤复位 `Pending`，循环顶部新增的「已判定步骤直接前进」把其余步骤逐个走过且**一条状态都不发**（每个状态在写下时已发过一次）；`While` 因此加了第三子句 `Interventions.PendingCount > 0`（与 E6-05b 的 `Not runEndGated` 同形的「再进一圈」），两处中段终态（不可恢复失败、replan 预算耗尽）在返回错误前先让位给收件箱，两处计划替换点在 `BeginNextPlan()` 后立刻 `Clear()`（新计划里同号不是同号）。面板侧新 type `agent:stepAction` **刻意不复用** `agent:approveStep`（审批握手带超时/放弃语义，抢来当第二用途会把 E1-05 词表污染成两种含义），处理函数只收下并回执、被拒时把原因送回卡片；卡片只在 `failed` 步骤渲染两个按钮、点击即锁、终态统一 `lockAllStepActions` 并说明「重跑请重新发起」——E10-03 的卡片侧合同由此一并落地，不另起一轮。明确声明的代价：同版重试不按台账去重（E1-06 的同版重复是用户语义，安全网是撤销点/E1-07 快照）、反复重试只由 `MaxIterations=15` 兜住（落地决定会复位 `noProgressCount`，注入用例 7 证明停在 15 且 `PlanGeneration` 仍为 1）、重试不消耗 `replanAttempts`（用例 2/5 断言 `Reflections=0`）、replan 预算耗尽 + 收件箱非空那次延迟只做结构断言（构造需约 9 个连续失败步，成本高于信息量）、E10-02 单步参数编辑仍未做。证据：`scripts/smoke-step-intervention.ps1`（新，自动登记，2 合同 + 7 注入用例驱动真实 `LoopEngine`，数宿主调用证明 A=1/B=2/C=1）、`check-agent-card-ui.js` 第 5 节 18 条断言、`audit-p0-guardrails.ps1` 新增 32 条 E10 断言，24 条变异探针逐条按「必须出现该断言自己的消息」变红后整文件字节还原 + MD5 校验。副产品：`check-task-package-contract.ps1` 的 `run-end-gate-unreachable` 探针原本绑死旧 `While` 整行字面量，E10 换行后 `.Replace` 失效，门禁自己的「探针必须变红」自检当场报 `slipped through`；已改为锚在 `OrElse Not runEndGated` 这个 token 上——**门禁的探针也会因别处的正常改动而失效，失效必须报红**。四个代码项目 Debug 构建 PASS，`run-code-checks.ps1` 29 个门禁全绿。

---

## 10. 与原始目标的对齐检查（每次周会必问）

1. 本周工作让 **O1 一次做成** 更容易了吗？证据是什么？
2. 让 **O2 敢让它做** 更强了吗？证据是什么？
3. 让 **O3 越用越懂** 更明显了吗？证据是什么？
4. 有没有为了“功能看起来更多”而做的东西？如果有，停止或移除。
5. 有没有出现新的假成功、不可撤销写操作、不可达入口？如果有，优先级高于新功能。

> **记住**：这份计划的目的不是把功能做完，而是让用户敢把文档交给它。任何偏离这一点的“改进”，都不是改进。
