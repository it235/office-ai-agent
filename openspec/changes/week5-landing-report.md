# Week 5 落地记录（E10 计划编辑与单步恢复）

> 本文件只记录 W5 的 E10 落地事实：做了什么、边界在哪、怎么证明、还有哪些真机步骤待跑。
> 计划行与排布见 `product-ai-native-strategy-execution-plan.md` §9 第 30 条与 E10 拆解表。

## 1. E10-01 失败步骤的用户重试/跳过

验收标准（计划行 E10-01）：**某一步失败后该步出现「重试这一步 / 跳过这一步」，点重试只重跑这一步（已成功步骤 0 次重跑）、不消耗重规划预算；点跳过该步标灰并继续后续步骤；无人点击时行为与今天逐条相同。**

### 1.1 先量 HEAD：缺的不是机制，是「用户的入口」和「一份收件箱」

动手前逐项 `git show HEAD:` 读回，五条都是量出来的：

1. `LoopEngine.vb` 里 `Interventions` 出现 **0 次**：HEAD 没有任何「用户对某一步做决定」的通道。
2. `agent-card.js` 里 `renderStepActions` / `onStepAction` / `setStepActionResult` 各 **0 次**：卡片上根本没有单步按钮。
3. `agent-protocol.js` 里 `stepAction` / `STEP_ACTION` 各 **0 次**：面板也没有发得出去的 type。
4. `AgentKernel.vb` 里 `QueueStepAction` **0 次**、`BaseChatControl.vb` 里 `agent:stepAction` **0 次**：从点击到循环，整条链路一段都不在。
5. 已有的 `retry`/`skip` 词汇只活在**反思分支**（`LoopEngine.vb:446` / `:453`，HEAD 行号）：由模型在连续 3 次失败后决定，而且每轮要花掉一份 `replanAttempts` 预算（`MaxReplanAttempts = 2`）。用户此刻能按的只有 `confirmAgentExecution` / `refineAgentPlan` / `abortAgent`，其中「修改计划」是把意见拼回原始请求重开一轮——与验收标准「不必重来」正相反。

因此这一行不是新建一套恢复循环，而是：**给用户一个入口、一份按步号寻址的收件箱，并让循环在既有的步骤边界消费它**。词汇沿用反思分支那一份（`retry`/`skip`），只是决定权从模型换回用户；多一套说法就多一份需要人工同步的语义。

### 1.2 一个收件箱、一个消费者、一扇门

- **收件箱**：`ShareRibbon/Agent/AgentStepInterventions.vb`（新，72 行）只做四件事——按步号入队（同号覆盖：重复点击等于同一个决定）、`Take` 时**同瞬间失效**（一条决定只许作用于一次失败）、给出升序步号快照（回退总是退到最靠前的那一步）、`Clear`。`Parse` 是一份**闭合词表**：只认 `retry` / `skip`（大小写不敏感），认不出来就是 `None`，不猜成「用户想重试」。
- **挂载点**：`AgentSession.Interventions`（`AgentSession.vb:42`）随会话走，不另立全局单例。
- **唯一的门**：`AgentKernel.QueueStepAction`（`AgentKernel.vb:199`）拒绝四类情况——没有在跑的会话、会话 id 不是当前这一轮、本轮已 `Completed/Failed/Aborted`、词表之外的动作。全拒时返回 `False`，让界面能说出「这一步已不在当前这一轮里」。
- **唯一的消费者**：`LoopEngine.ApplyUserInterventions`（`LoopEngine.vb:574`）。全仓 `session.Interventions.Take(` 只出现 **1 次**；`AgentKernel` 一行 `.Status = StepStatus` 都没有——宿主或面板直接改 `PlanStep` 会造出第二套执行真相，而只有循环知道当前步号、能安全回退到更早的那一步。

### 1.3 「只重跑受影响步骤」是结构，不是注释里的承诺

循环顶部三步，顺序就是合同（`LoopEngine.vb:114-126`）：

1. 有收件箱就先落成步骤状态：`skip` → `Skipped` + 发一次 `skipped` 状态；`retry` → 复位 `Pending` + 发一次 `pending`，并把 `stepIndex` 退到 `Math.Min(stepIndex, idx)`。
2. **已判定过的步骤直接前进**：`Status <> StepStatus.Pending` 就 `stepIndex += 1` 继续，而且**一条状态都不发**——每个状态在它被写下的那一刻已经发过一次，再发一次只会让卡片重复出现同一句。
3. 只有 `Pending` 才进入执行。

于是「重试第 2 步」的实际效果是：回退到 2，3 及之后已完成/已跳过的步骤被逐个走过、不重放；「重试第 0 步」同理只重跑 0，后面全部走过。`scripts/smoke-step-intervention.ps1` 的注入用例直接数宿主调用：重试 B 时 A=1、C=1、B=2；回退到 A 时 A=2、B=1、C=1。

**循环必须再进一圈**：用户可能在某步失败后、循环已经走过最后一步时才点。`While` 条件因此加了第三个子句 `session.Interventions.PendingCount > 0`（与 E6-05b 的 `Not runEndGated` 同一种「再进一圈」形状），否则收件箱里的决定会一直等到下一轮——而下一轮永远不会来。

### 1.4 两处中段终态必须先让位

如果失败发生在「马上要收敛」的那一刻，用户点击还没发生，终态就把这一轮关掉了。两处中段终态因此在返回错误前先看收件箱：

- 不可恢复失败（`LoopEngine.vb:440`）：执行器判定 `Recoverable = False`，「不可恢复」是执行器的判断，不是「用户也只能看着」。
- 重规划预算耗尽（`LoopEngine.vb:463`）：这是「AI 想不出修法」，不是「用户也无话可说」。

两处都只做一件事：`Continue While`，回到步骤边界把决定落地。没有新增终态出口，成功终态仍只有 `AgentResult.SuccessResult` 一处。

### 1.5 计划被整体替换时，旧决定必须作废

收件箱按步号寻址，而 `replan` 与 run 末修复（E6-05b）都会换一份新计划——**新计划里的同号不是同号**。两处替换点（`LoopEngine.vb:148`、`:507`）都在 `session.Writes.BeginNextPlan()` 之后立刻 `session.Interventions.Clear()`；`audit-p0-guardrails.ps1` 的 E10 段钉住「恰好 2 个 Clear，且每个都紧跟 BeginNextPlan」。

### 1.6 面板侧：一条独立通道、一次诚实回执

- 新 type `agent:stepAction`（`agent-protocol.js` 的 `STEP_ACTION` + `requestStepAction`），**刻意不复用** `agent:approveStep`：审批握手带超时、放弃、trace 语义，`DecideAsync` 还要 `Await pending.AgentTask`；把它抢来当第二个用途会把 E1-05 那套 `Once/Session/Deny/Abort` 词表污染成两种含义。
- `BaseChatControl`：`router.Register("agent:stepAction", …)`（`:866`）→ `HandleAgentStepAction`（`:2169`）只收下并回执（`GlobalStatusStrip.ShowInfo("已收到，当前步骤结束后处理这一步")`），**不预测何时生效**；被拒绝时 `ReportStepActionRefused`（`:2197`）把原因送回同一张卡片。
- 卡片：`renderStepActions`（`agent-card.js:244`）只在 `status === 'failed'` 时出现两个按钮（沿用既有 `agent-btn-execute` / `agent-btn-abort` 样式，零新 CSS）；点击后**立即锁住两个按钮**并写提示，避免同一决定被连点；宿主拒绝时 `setStepActionResult` 解锁并写出原因——「按了没反应」比没有按钮更糟。
- **终态收口**：`abortAgent` 与 `completeAgent` 都调 `lockAllStepActions`，把这一轮自己的按钮置灰并说明原因（「本轮已经终止，这一步不能再单独处理」/「本轮已经结束，需要重跑这一步请重新发起任务」）。这同时覆盖了计划行 E10-03 的卡片侧合同：终态下单步动作不可用，且必须说明可用出口是「撤销本轮或带意见重新发起」。E10-03 因此不另起一轮实现，剩真机确认。

### 1.7 代价与已知边界（明确声明的不作为）

1. **同一计划版本内的重试不去重。** E1-06 的写入台账只挡「更早计划版本的已落地写入」，同版重复是用户明确要求的语义；用户点重试就是要求再执行一次。安全网是撤销点与 E1-07 的快照撤销，而不是替用户拒绝。
2. **反复重试只由 `MaxIterations = 15` 兜住。** 每次落地决定都会把 `noProgressCount` 复位（用户的决定不该被算作「模型没有进展」），所以 `MaxNoProgress = 3` 不再约束重试；注入用例 7 证明无界重试最终停在 15 次迭代、且 `PlanGeneration` 仍是 1（没有偷偷变成 replan）。
3. **重试不消耗 `replanAttempts`。** 用户的纠偏走的是步骤边界，不经过反思分支，`replanAttempts` 只服务模型 replan。用例 2/5 直接断言 `Reflections = 0`、`PlanGeneration = 1`。
4. **E10-02（单步正文/参数编辑）仍未实现。** 它需要一个新的「改这一步」入口与参数改写合同，不属于「恢复」这一行的最小闭环。
5. **唯一没有被行为测试覆盖的分支**：重规划预算已耗尽（`replanAttempts >= 2`）同时收件箱非空的那次延迟。构造它需要约 9 个连续失败步骤（每次重试/跳过都会复位 `noProgressCount`），成本高于信息量；该分支由 E10 守卫按形状钉住（两处 `Return AgentResult.Failed(...)` 前都有同一个 `If session.Interventions.PendingCount > 0 Then … Continue While`），并在此处如实记录为「结构断言、行为未测」。

## 2. E10-04 门禁：把「能纠偏」变成可重跑的计数

### 2.1 `scripts/smoke-step-intervention.ps1`（新，自动登记进门禁）

在 PowerShell 里反射加载编译后的 `ShareRibbon.dll`，用脚本化的 ReAct 回答驱动**真实 `LoopEngine`**，9 个用例：

- 2 个合同用例：收件箱（闭合词表、同号覆盖、Take 一次、升序、Clear）+ `AgentKernel` 的门（空会话 / 别的 run / 未知动词 / 三种终态 / 没有 run 在飞）。
- 7 个注入用例：无人点击无变化；重试只重跑该步且零 replan、版本不变；跳过只落一次且其余步骤照跑；不可恢复失败先让位给已排队的决定（没有点击时 C 步根本不执行）；回退到更早步骤只重跑该步；不存在的步号被丢弃而不是落到邻居；无界重试被 `MaxIterations` 兜住且版本仍为 1。

输出：`PASS step intervention: 9 cases (2 inbox/door contracts + 7 injected-failure runs over the real LoopEngine)`。

### 2.2 `scripts/check-agent-card-ui.js` 第 5 节（18 条断言）

同一个桩 DOM + `vm` 沙箱先加载 `agent-protocol.js` 再加载 `agent-card.js`，只断言代码形状能观察到的行为：`STEP_ACTION` 与 `APPROVE_STEP` 不同；只有失败步骤出现按钮；一次点击只发一条 `agent:stepAction` 且载荷是 `{sessionId, stepIndex, action}`；点击后两个按钮都锁住；锁定后再点不发消息；宿主拒绝会解锁并写原因（走 `textContent`，不写 markup）；`ok=true` 保持锁定；步骤进入 `pending` 时按钮消失、再次失败得到一组全新的可用按钮；终态禁用并解释「本轮已经结束」。

### 2.3 `scripts/audit-p0-guardrails.ps1` 新增 E10 段（32 条 `$failures.Add`）

钉住行为测试跑不到或不值得跑的结构性不变式：

- 收件箱：`Take` 必删、词表闭合、升序快照；`AgentSession` 仍持有它；`AgentStepInterventions.vb` 仍在 `ShareRibbon.vbproj`。
- 单持有者：`Take(` 全仓恰好 1 次；`AgentKernel` 不消费收件箱、不写步骤状态；门的四类拒绝仍在。
- 边界与顺序：`ApplyUserInterventions` 的位置（在「已判定步骤前进」之前、在执行之前）、已判定步骤的走向、回退用 `Math.Min`、跳过写 `Skipped`、`While` 的第三个子句与迭代上限。
- 两处中段终态的延迟、两处计划替换的清空与顺序。
- 面板：`agent:stepAction` 注册、处理函数不经过审批词表、拒绝必须有回执、`RequestStepAction` 仍返回字符串原因。
- 卡片与协议：五个函数、只对 failed 渲染、调用 `requestStepAction`、终态 `lockAllStepActions`。
- 门禁自身：`smoke-step-intervention.ps1` 存在且**没有**被 `$notGates` 摘掉；`check-agent-card-ui.js` 仍加载协议脚本。

### 2.4 变异探针：24 条，全部按预期变红

用一个临时探针脚本（跑完即删）逐条改坏源码/工程/门禁，再跑 P0 审计，要求**输出里出现那条断言自己的消息**（不是「只要不绿就算」），恢复用整文件字节还原 + MD5 校验：

| # | 改坏的东西 | 变红的断言 |
|---|---|---|
| 1 | 删掉边界处的 `ApplyUserInterventions` 调用 | 不再是唯一的落地点 / 顺序错 |
| 2 | 把「已判定步骤前进」的条件反过来 | 同上（位置锚点消失） |
| 3 | 回退改成 `stepIndex = idx` | 不再退到用户点的那一步 |
| 4 | 跳过写成 `Failed` | 跳过不再落成 skipped |
| 5 | 删掉 `While` 第三个子句 | 收尾后的点击被静默丢弃 |
| 6 | 删掉不可恢复失败的延迟 | 不可恢复失败仍然直接给终态 |
| 7 | 删掉重规划预算耗尽的延迟 | 同上 |
| 8 | 删掉一处 `Clear` | 清空点不再是两处 |
| 9 | 把第二处 `Clear` 挪到 `BeginNextPlan` 之前 | 新计划继承旧决定 |
| 10 | `Take` 调两次 | 循环不再是唯一消费者 |
| 11 | 内核里调用 `Take` | 内核消费收件箱 |
| 12 | 终态门改成 `Return True` | 已结束的一轮仍收决定 |
| 13 | 会话 id 比较改成忽略大小写 | 别的 run 的决定被收下 |
| 14 | 处理函数里引用 `ApprovalDecision` | 走了审批握手 |
| 15 | 吞掉拒绝回执 | 拒绝不告诉卡片 |
| 16 | 所有状态都渲染按钮 | 没失败的步骤也有动作 |
| 17 | `STEP_ACTION` 改名成审批 type | 协议不再导出单步消息 |
| 18 | 卡片改调 `requestStepActions` | 按钮不再走单步协议 |
| 19 | 从 `.vbproj` 删掉收件箱 | 收件箱未编进程序集 |
| 20 | 把 smoke 登记进 `$notGates` | 注入门禁被摘掉 |
| 21 | UI 检查不再加载协议脚本 | 卡片合同少了一半 |
| 22 | 删掉 `Take` 里的移除 | 决定不再 take-once |
| 23 | 重命名会话属性 | 会话不再持有收件箱 |
| 24 | 把 smoke 文件挪走 | 注入门禁不存在 |

### 2.5 触发的既有门禁修正（E6-05b 的探针曾绑死旧字面量）

`check-task-package-contract.ps1` 的 `run-end-gate-unreachable` 探针原本用 `Replace('While (stepIndex < session.Plan.Steps.Count OrElse Not runEndGated) AndAlso', …)` 改坏循环条件。E10-01 给 `While` 加了第三个子句并换行，探针的 `.Replace` 找不到目标——**状态没被改坏，检查自然还是绿的**，门禁自己的「探针必须变红」自检当场报 `identity probe 'run-end-gate-unreachable' slipped through`。修法是把探针锚在更小的语义单元上：删掉 `OrElse Not runEndGated` 这个 token，而不是整行字面量。这条修正是本轮最重要的副产品之一——**门禁的探针也是代码，也会因为别处的正常改动而失效**，失效时必须变红（它做到了），而不是静默放行。

## 3. 构建与门禁结果

- `scripts/build-code-projects.ps1 -Configuration Debug`：ShareRibbon / WordAi / ExcelAi / PowerPointAi 四个代码项目全部 PASS。
- `scripts/run-code-checks.ps1`：发现 29 个门禁 + 全部 JS 语法检查，最终 `All code checks passed.`（含新增的 `smoke-step-intervention.ps1` 与 `audit-p0-guardrails.ps1` 的 E10 段）。
- `scripts/smoke-step-intervention.ps1`：9 用例 PASS。
- `node scripts/check-agent-card-ui.js`：`agent-card UI contract checks passed`（56 条断言，其中第 5 节 18 条属 E10）。
- `scripts/audit-p0-guardrails.ps1`：`P0 guardrails passed`（新增 32 条 E10 断言，24 条变异探针逐条验证非空断言）。

## 4. 真机验收清单（E10，按指示暂缓；需要 Word 2021 真机 + 重启宿主刷新 `%LOCALAPPDATA%\OfficeAI\www`）

准备：准备一份**可以被写入失败的** Word 文档（最简单是只读打开，或受保护的文档），以及一个会产生 2–3 步写入的任务；确认 `%LOCALAPPDATA%\OfficeAI\www` 下的 `agent-card.js` / `agent-protocol.js` 时间戳是本次构建。

1. **按钮只在失败步骤出现**：让第 2 步写失败，卡片上只有第 2 步出现「重试这一步 / 跳过这一步」，第 1、3 步没有。
2. **重试只重跑这一步**：点「重试这一步」→ 第 2 步回到进行中并重新执行；第 1 步不重新执行（观察文档里第 1 步的写入没有二次痕迹/撤销栈没有新增），第 3 步不受影响。
3. **重试不花重规划预算**：连续点两次重试仍不触发「正在分析失败原因并重新规划…」；此时若真触发 replan，说明走了反思分支而不是收件箱。
4. **跳过落一次**：点「跳过这一步」→ 第 2 步变灰（skipped 样式，不是红叉），后续步骤继续；终态摘要写出跳过。
5. **无人点击时行为不变**：同样任务不点任何按钮，走到连续 3 次失败后仍由模型 replan——与今天逐条相同。
6. **早于当前步的重试**：第 2 步失败后循环已走到第 4 步时再点「重试第 2 步」→ 循环退回第 2 步，只重跑第 2 步，第 3 步不重放。
7. **收尾后的点击**：本轮已成功结束后再点旧卡片上的按钮（如果还能看到）→ 按钮已禁用并写明「本轮已经结束…」；成功/终止两条路径都要看。
8. **终止后的点击**：点「终止执行」→ 该轮所有单步按钮立刻禁用，提示「本轮已经终止…」。
9. **拒绝回执**：在已经结束的一轮上（或面板与宿主状态不同步时）点到按钮 → 按钮会解锁，并写出「这一步已经不在当前这一轮里了…」而不是静默置灰。
10. **计划被替换后旧决定作废**：失败后立即点重试，同时该轮恰好触发 replan → 新计划不会把旧步号的决定落在一个语义无关的新步骤上（观察新计划卡的第 2 版与执行步序）。
11. **不可恢复错误也收点击**：构造一个执行器判定不可恢复的失败（如宿主 COM 异常），点击重试仍能被处理，而不是立刻收到终态错误。
12. **三端一致**：Word / Excel / PowerPoint 各做第 1–4 条，确认按钮、状态令牌与回执在三个宿主都一致（逻辑在 `ShareRibbon`，本轮未改宿主专有代码）。

## 5. 下一步

- **E10-02**：单步正文与参数编辑（同一版计划、只重跑这一步；参数指纹变了所以台账不会把它当重复写入挡掉）。
- **E10-03** 只剩真机确认（卡片侧合同已随 E10-01 落地并有 UI 断言）。
- 仍未执行的旧真机验收：E1-05 / E2-02 / E2-03，以及 week3/week4 报告里登记的待跑清单。
- 等用户拍板的开放项：M1 流水线托管、#35、#36、D7、D8、D9，以及是否推送尚未推送的提交。
