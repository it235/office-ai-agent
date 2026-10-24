# Week 2 落地报告与真机验收清单

> 日期：2026-10-24
> 范围：E1-05 审批粒度、E2-02 会话绑定文档、E2-03 会话自动标题/重命名/删除、E3-03 空 Key 阻断、E3-01 首次使用引导向导（其余 W2 条目后续追加）。
> 结论：代码与自动化校验已通过；真机验收待执行。

## 1. E1-05 审批粒度

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 三态裁决 | `ShareRibbon/Agent/ApprovalModels.vb` | 新增 `ApprovalDecision { Deny, Once, Session, Abort }` 与 `ApprovalRequest` |
| 会话级放行 | `AgentSession`、`ToolExecutionContext`、`ToolRegistry` | `SessionApprovedTools` 只在本轮 Agent 会话内生效；命中后不再弹窗，且不会被一次性额度消耗 |
| 一次性放行 | `ToolExecutionContext.ApproveTool` + `ConsumeToolApproval` | 按“工具 + 规范化参数”放行一次，执行后即失效 |
| 拒绝只跳过当前步骤 | `LoopEngine` | `Deny` 生成 `decision=rejected` 的审批观察，当前步骤标记 `Skipped`，后续步骤继续；终态附带“已跳过 N 个你拒绝的操作” |
| 放弃立即停轮 | `LoopEngine`、`OfficeHarness` | `Abort` 生成 `decision=aborted`，本轮直接以失败收敛，不再执行后续步骤 |
| 超时/取消/取代 | `OfficeHarness.Approvals`、`OfficeHarness` | 审批超时、`CancelAsync`、新一轮对话取代挂起审批统一走 `Abort`，未确认的写操作绝不因为等待或切换话题而执行 |
| 前端三按钮 | `agent-card.js`、`agent-protocol.js`、`styles.css`、`BaseChatControl`、`AgentKernelService.RunControl` | “允许一次 / 本会话允许此工具 / 拒绝”分别路由到 `DecideAsync(Once/Session/Deny)` |

关键约束：

- 审批通道只有一条：`LoopEngine -> AgentKernel -> OfficeHarness PendingHarnessRun -> AgentKernelService UI -> 回传裁决`，没有新增平行状态机。
- `Abort` 不是“拒绝的另一种写法”：拒绝是用户对当前步骤的裁决，放弃是审批通道已经关闭（超时/停止/被打断），两者终态与后续行为必须不同。
- 会话级放行作用域是本轮 Agent 会话，不跨会话持久化。

## 2. E2-02 / E2-03 会话归属、标题与删除

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 会话档案表 | `Storage/Migrations/017_session_document_binding.sql`、`OfficeAiDatabase.vb`、`OfficeAiDbSchema.current.sql` | `session_summary` 改为“一行一会话”的档案：`app_type / document_id / document_name / custom_title / updated_at`，并建 `session_id` 唯一索引与文档维度索引 |
| 半途失败自愈 | `OfficeAiDatabase.EnsureSessionProfileSchema` | 迁移负责建形状，启动健康检查把去重和缺列补齐收敛回来，不让一次失败的 ALTER 卡死启动 |
| 文档标识 | `Controls/Models/SessionDocument.vb`、三宿主 `ChatControl.GetSessionDocument` | 由活动文档绝对路径做小写 SHA-256 前 8 字节；未保存文档（无目录）不绑定，路径本身不入库，只显示文件名 |
| 先写者胜 | `ConversationRepository.UpsertSessionProfile` | 会话归创建它的文档；后续轮次切换活动文档不得把已有会话搬走 |
| 作用域 | `ConversationRepository.GetRecentSessions(scope)`、`HistorySessionService.HandleGetSessionList` | `document` 只看当前文档；`all` 看全部但当前文档排最前；无档案的老会话仍可见、仍按“通用会话” |
| UI 线程约束 | `BaseChatControl.RefreshSessionDocumentContext` | 文档标识只在 UI 线程读取并缓存，后台落库线程不再回读 COM |
| 自动标题 | `ConversationRepository.AutoTitle` | 折叠空白，优先在句末标点收尾（小数点 `3.5` 不算句末），其次词边界截断加省略号；不调模型，成本闸门未落地前不为侧边栏多花一次请求 |
| 重命名 | `ConversationRepository.RenameSession`、路由 `renameSession`、`history-manager.js` | 自定义标题覆盖自动标题；空标题回到自动标题；老会话缺档案行时补行且不误造文档绑定 |
| 删除 | `HistorySessionService.HandleDeleteSession`、`history-manager.js` | 前端二次确认后级联清理消息/档案/记忆/队列/轨迹；删除当前会话则切新会话并清空聊天区；列表按当前作用域刷新 |

关键约束：

- 聊天历史的唯一权威源仍是 `conversation_event`，`session_summary` 只是归属与标题档案，列表聚合不依赖它是否存在。
- 历史顺序由后端决定（当前文档优先 → 时间倒序），前端不再二次排序，否则会把“本文档优先”冲掉。
- 历史条目动作走列表级事件委托 + `dataset.bound` 幂等标记，避免每次刷新重复挂同一批监听。

## 3. E3-03 未配置阻断与面板 CTA

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 唯一配置闸门 | `Controls/Services/ChatSendValidator.ValidateConfig` | 配置检查（Key/URL）与“有没有输入”分离，`Validate` 复用同一函数，聊天与 Agent 不可能各判一次 |
| 阻断载荷 | `Controls/Services/ModelConfigGate.vb` | `ModelConfigPrompt{Reason,Title,Message,ActionLabel,BlockedAction}`，原因码 `missing_api_key` / `missing_api_url`，序列化走 `JObject`，不手拼 JS 字符串 |
| 聊天入口 | `BaseChatControl.Send` | 未配置时不再发送，也不在状态栏打字报错；改推 `showModelConfigRequired({...})` 在面板内出卡片 |
| Agent 入口 | `BaseChatControl.HandleStartAgent` | 未配置时先闸门后 `RefreshSessionDocumentContext()`，避免半启动的 Agent 轮次留下空思考气泡 |
| CTA 可执行 | `Resources/js/chat-manager.js` | 卡片固定 id `model-config-cta`（重复阻断只刷新不叠卡），按钮复用既有 `openApiConfigForm` 桥；另有“查看使用引导”唤起 E3-01 卡片 |
| 陈旧卡片自愈 | `chat-manager.js`、`BaseChatControl.HandleOpenApiConfigForm` | 收到 AI 回复或配置窗口保存成功后主动 `hideModelConfigRequired()`，不会出现“已经配好但还挂着红字” |
| Key 落盘 | `Config/ConfigManager.SaveConfig` | 逐个把真实 Key 写入 DPAPI（当前用户），配置文件里只留 `[encrypted]` 占位；`Finally` 恢复内存中的真实 Key，因为 MCP、意图识别、翻译都直接读 `cfg.key` |
| 快照单点投射 | `Config/ConfigManager.ApplyConfigSnapshot` | `LoadConfig`/`SaveConfig` 之后统一把“选中服务商 + 选中对话模型”写进全局 `ConfigSettings`，且无条件赋值：清空 URL/Key 或删除当前服务商后，面板不会继续拿着上一次的端点发请求（真机复现过的缺陷，E3-03 闸门也会被旧 Key 放行） |

关键约束：

- 闸门只有 `Send` 这一条主链上的一个函数；没有为 Agent 另写第二套校验，也没有把检查下沉到 `AiGateway`（那里拿不到“用户点的是发送还是执行”的语义）。
- CTA 必须给可执行动作，不接受“请先配置 API Key”这种纯文案：文案区分“消息没有发送”和“任务没有开始执行，文档不会被修改”。

## 4. E3-01 首次使用引导向导

首版把“选服务商 → 填 Key → 测连接”也搬进了面板卡片，复查后判定为重复实现：这三步在 `ConfigApiForm` 里早就完整存在（填 Key 的校验、模型选择、连接测试甚至用的是同一个 `ModelConnectivityTester`）。现在向导只保留此前**确实没有**的两件事，配置动作直接跳既有配置窗口。

| 项 | 落点 | 行为变化 |
| --- | --- | --- |
| 向导状态 | `Controls/Services/OnboardingService.vb` | 只剩 `HandleGetState / HandleSetPrivacy / HandleFinish`；`BuildState()` 只给 `needsSetup / privacyAccepted / appType / step / examples`，**不再下发 providers 清单**，也不读写 `ConfigManager.ConfigData` |
| 步骤起点 | `OnboardingService.InitialStep` | 只有 privacy → example 两步，完全由宿主状态决定；前端不再保存步骤游标（原来的 `onboardingStepPinned` 会让卡片停在用户点过的那一步，和宿主的权威步骤长期矛盾） |
| 是否已配置 | `ChatSendValidator.ValidateConfig(ConfigSettings.ApiUrl, ConfigSettings.ApiKey)` | 向导用发送闸门的同一个函数判断“配好了没”，不会出现“引导说配好了、发送仍拦你” |
| 配置入口 | `onboarding.js` → `openModelConfig()` → `BaseChatControl.HandleOpenApiConfigForm` | 未配置时示例步骤给出“打开配置窗口”按钮，走的就是 E3-03 CTA 已有的那条桥；配置窗口关闭后宿主重推一次 `getOnboardingState`，卡片自动收敛到示例步 |
| 快照收口 | `ConfigApiForm.CloudSaveButton_Click` / `LocalSaveButton_Click` | 删除两处手写的 `ConfigSettings.ApiUrl/ApiKey/platform/ModelName/mcpable/fim*` 赋值，保存后由 `SaveConfig -> ApplyConfigSnapshot` 投射；全仓库只剩这一处会写这四个字段 |
| 连接测试 | `Config/ConfigApiForm.ValidateApiAsync` | 测试器返回 `ModelConnectivityResult` 而不是布尔；失败弹窗改用 `ValidationFailureText()`，把 `detail`/`reason`（401、超时、模型不存在…）原样告诉用户，不再只说“请检查 API Key” |
| 快照单点投射 | `Config/ConfigManager.ApplyConfigSnapshot` | `LoadConfig`/`SaveConfig` 之后统一把“选中服务商 + 选中对话模型”写进全局 `ConfigSettings`，且无条件赋值：清空 URL/Key 或删除当前服务商后，面板不会继续拿着上一次的端点发请求（真机复现过的缺陷，E3-03 闸门也会被旧 Key 放行） || 隐私确认 | `SecureConfigManager.PrivacySettings.PrivacyPolicyAccepted` | 沿用既有开关，不新增“同意过”字段；未同意时向导从 privacy 开始，可撤回 |
| 示例任务 | `OnboardingService.BuildExamples` | 按 `appType` 给 Excel / PowerPoint / Word 三套各 3 条，点一下填入输入框并收尾 |
| 前端渲染 | `Resources/js/onboarding.js`、`Resources/css/styles.css` | 两步卡片 + 两个进度点，可关闭、可“稍后再说”；服务商/Key/测试相关的 CSS 已随之删除，守卫会盯住残留死样式 |
| 桥接 | `BaseChatControl`（`getOnboardingState`、`onboardingSetPrivacy`、`onboardingFinish`） | 三个消息类型登记；已删除 `onboardingSelectProvider / onboardingSaveKey / onboardingTestConnection`，守卫反向禁止它们回来 |

关键约束：

- 向导不引入第二套配置实现、第二套存储、第二套闸门：它只补“隐私确认 + 示例任务”这两个配置窗口没有的环节。
- 配置的唯一写入方是 `ConfigApiForm` + `ConfigManager`；`OnboardingService` 里出现 `SaveConfig`、`ModelConnectivityTester`、`ConfigData` 或对 `ConfigSettings` 的赋值都算违规（守卫直接变红）。
- 前端不保存步骤状态：卡片显示哪一步只由宿主推来的 `step` 决定，配置窗口保存成功后卡片要么收敛到示例、要么自动收起。
- E3-03 的“查看使用引导”只是唤起卡片，配置本身仍然由同一个 `openApiConfigForm` 桥完成。

## 5. 自动化验证结果

- `build-code-projects.ps1 -Configuration Debug`：ShareRibbon / WordAi / ExcelAi / PowerPointAi 全部通过。
- `audit-p0-guardrails.ps1`：新增 E1-05/E1-05b 守卫，覆盖三态入口、拒绝限步、会话放行不被消费、`Abort` 仅用于停轮路径；新增 E2-01/E2-02/E2-03 守卫，覆盖宿主过滤、档案先写者胜、本文档优先排序、未保存文档不绑定、UI 线程刷新文档上下文、自动标题与重命名/删除路由、前端委托绑定与二次确认。
- `smoke-harness-approval.ps1`：
  - 超时与取代后会话可接受新一轮；
  - `Once` 放行后宿主执行器确实被调用；
  - 会话放行大小写不敏感、可重复查询、不被 `ConsumeToolApproval` 消耗；
  - `Abort` 不触发宿主执行、不报成功、并释放会话；
  - 等待审批期间宿主执行器调用次数为 0。
- `smoke-change-log-and-session-scope.ps1`：变更日志聚合/持久化/撤销作用域之外，新增 E2 断言——宿主隔离、本文档优先、文档作用域不串台、先绑定不被后写覆盖、自动标题（句末断句 + 小数点保护 + 词边界省略号）、重命名覆盖自动标题、老会话补档案、级联删除清掉档案行；断言失败时由 `finally` 统一清理 `smoke-` 前缀夹具，不再把测试会话留在真实库里。
- `smoke-empty-db-initialization.ps1`：空库期望版本升到 17，并断言档案列齐全、`session_id` 唯一、两轮消息只会有一行档案、升级后的老会话仍为未绑定。
- `smoke-memory-pipeline.ps1`（本轮修复的真实缺陷）：该冒烟此前按 `event_type='message'` 插入事件、并把 `TargetId` 填成 assistant 事件 id，与 `72522f4` 之后的生产契约（`chat_message` + `TargetId = sessionId`）不一致，于是任务“成功”却提取出 0 条记忆——冒烟变红才暴露出来。现按生产契约修正，并注入规则抽取器使门禁不依赖模型可用性；新增 TTL 断言（过期记忆不可被检索）与夹具自清理。
- `run-code-checks.ps1 -SkipBuild`：全部冒烟、Golden L0 12 例、JS `node --check`（已把 `history-manager.js`、`onboarding.js` 纳入检查清单）通过。
- E3-03 守卫（`audit-p0-guardrails.ps1`）：`ValidateConfig` 独立存在且被 `Validate` 复用、原因码与 JSON 序列化、聊天/Agent 两个入口都走闸门、配置保存后隐藏陈旧 CTA、CTA 固定 id 且复用 `openApiConfigForm`。
- `smoke-model-config-gate.ps1`：完整配置不拦；空 Key → `missing_api_key`，纯空格 Key 也算未配置；空 URL → `missing_api_url`；聊天与 Agent 文案必须不同；`Validate` 的判定顺序 Key→URL→输入；推送脚本是 `showModelConfigRequired({...})` 且载荷能被 `ConvertFrom-Json` 解析；引号/换行/反斜杠经 `ModelConfigPrompt` 往返后脚本仍是一条完整语句；无提示时不产生脚本；`BuildHideScript()` 与前端渲染函数同名。
- E3-01 守卫（收敛后）：只登记 `getOnboardingState / onboardingSetPrivacy / onboardingFinish` 三个消息类型；`onboardingSelectProvider / onboardingSaveKey / onboardingTestConnection` 一旦重新出现在宿主注册或 `onboarding.js` 里就变红；`OnboardingService` 内禁止 `ConfigManager.SaveConfig`、`ModelConnectivityTester`、`SecureConfigIntegration`、对 `ConfigSettings` 的赋值、直接摸 `ConfigManager.ConfigData`；`needsSetup` 必须来自 `ChatSendValidator.ValidateConfig(ConfigSettings.ApiUrl, ConfigSettings.ApiKey)`；`InitialStep` 只允许 privacy/example，出现 provider/key/test 返回值即失败；`onboarding.js` 必须复用 `openModelConfig`、且不得再自建步骤游标（`onboardingStepPinned` / `window.onboardingStep =`）；`styles.css` 里每一条 `.onboarding-*` 规则必须仍被 `onboarding.js` 渲染（防止删掉步骤后留下死样式，扫描到 13 条、死样式 0）；`HandleOpenApiConfigForm` 方法体内必须同时出现 `ModelConfigGate.BuildHideScript()` 与 `OnboardingSvc.HandleGetState()`（按方法体正则取，取不到就判失败）；`ConfigApiForm.ValidateApiAsync` 必须返回 `ModelConnectivityResult` 并用 `ValidationFailureText(validationResult)` 呈现，且不得再手写 `ConfigSettings.ApiUrl/ApiKey/platform/ModelName`；`SaveConfig` 必须经 `SecureConfigManager.SaveApiKey` 且配置文件只留 `[encrypted]`。
- 新增两条通用守卫：
  - 静态资源登记完整性——`Resources\js\*.js` 必须同时出现在 `Resources.resx` 和 `ResourceExtractor` 映射中，否则新 JS 编译能过但面板里只是“函数不存在”；
  - 桥接契约完整性——面板 `type: 'xxx'`  的每一个 post 消息类型都必须能在四个代码项目的 `Register("xxx")` / `Case "xxx"` 里找到宿主处理函数（58 个类型全覆盖，扫描数低于 40 视为扫描本身坏掉）。已用临时假类型 `zzzMissingHandlerProbe` 反向验证守卫确实会红。
- `smoke-onboarding-wizard.ps1`（反射加载编译产物，前后对 `office_ai_config.json` 与 `secure.config` 做字节级备份/还原）：载荷为 `needsSetup / privacyAccepted / appType / step / examples`，且断言 `providers` 字段不存在（一旦回来就判失败）；空 URL + 空 Key → `needsSetup`；未同意隐私时起点必为 privacy，同意后收敛到 example，撤回后回到 privacy（步骤只由宿主决定）；载荷里没有 accepted 时不得当作同意；示例任务在 Excel / PowerPoint / Word 三套必须互不相同且各不少于 2 条；`HandleFinish` 恰好触发一次宿主刷新；跑完整条向导后 `office_ai_config.json` 字节不变（证明向导不写配置）；尾部保留共享测试器的离线拒绝断言——空 URL / 空模型直接返回 `invalid_config`、给出 `detail`、3 秒内结束且不开套接字，`ToJson()` 里 `reason` 仍是机器可读的 `invalid_config`。
- E3-03b 守卫 + `smoke-config-snapshot.ps1`：钉住 `ApplyConfigSnapshot` 存在、`ConfigSettings.ApiUrl/ApiKey` 必须是无条件 `If(chosen?..., "")` 形式、`LoadConfig` 与 `SaveConfig` 方法体内都必须调用它、`smoke-config-snapshot.ps1` 必须登记在 `run-code-checks.ps1`；冒烟覆盖六组真值——正常投射（含空 `fimUrl` 回落服务商 URL）、**清空 URL/Key 后快照必须一起变空**（真机复现的回归）、没有选中项、选中模型被取消勾选、`ConfigData = Nothing` 不抛异常，以及快照变空后 `ModelConfigGate.Evaluate` 必须返回 `missing_api_key`。

## 6. 真机验收

通用前置：关闭 Office；备份 `%USERPROFILE%\Documents\OfficeAiAppData\office_ai.db`（Release）或 `OfficeAiAppData-Debug`（Debug）；启动宿主后等待迁移完成。

### 6.1 E1-05 审批粒度

1. 触发一个需要审批的高风险写操作，确认审批卡出现三个按钮。
2. 点“允许一次”：操作执行一次；同一工具再次触发时仍弹审批。
3. 点“本会话允许此工具”：当前会话内同工具后续操作不再弹审批；重启会话后重新弹审批。
4. 构造两步计划，拒绝第 1 步：第 2 步继续执行，终态明确提示已跳过 1 个操作，文档只有第 2 步的结果。
5. 审批卡停留至超时：本轮停止，提示“审批等待超时，本轮执行已停止”，文档无新增写入，会话可立即接受新任务。
6. 审批卡停留时点击“停止”或直接发送新需求：上一轮未确认的操作不执行，新任务正常开始。

查询 DB 复核：

```sql
SELECT run_id, seq, tool_id, status, error_code, observation_json
FROM agent_run_step
WHERE tool_id IN ('approval.request', 'approval.decision')
ORDER BY started_at DESC LIMIT 20;
```

`observation_json.decision` 应为 `approved_once` / `approved_for_session` / `rejected` / `aborted` / `timeout`，且 `observation_json.allowed` 仅在前两者为 `true`。

### 6.2 E2-02 / E2-03 历史侧边栏

前置：保存两篇不同的 Word 文档（例如 `report.docx`、`budget.docx`），并留一篇新建未保存文档。

1. 打开面板历史侧边栏：顶部出现“本文档 / 全部”两个页签，右侧显示当前文档名；未保存文档显示“通用会话”。
2. 在 `report.docx` 聊两轮、切到 `budget.docx` 再聊一轮：`全部` 里 `report.docx` 的会话排在最前，每条下方显示所属文档名。
3. 在 `report.docx` 中点“本文档”：只出现该文档的会话，`budget.docx` 与通用会话都不出现；空态给“这篇文档还没有会话 / 查看全部会话”。
4. 在 `report.docx` 的会话里继续发言后再切走活动文档：该会话仍归属 `report.docx`（先写者胜，不会被搬走）。
5. 用一句较长的需求（含句号）发起会话：列表标题在句末自然断开、可读；含 `3.5` 这类小数不被当成句末截断，超出长度以省略号结尾。
6. 点 ✎ 输入新标题回车：标题立即变为自定义名并在重新打开面板后保持；输入框 Esc 取消；把标题清空则回到自动标题。
7. 点 🗑：出现二次确认；确认后列表刷新且该会话消失。删除当前会话时聊天区被清空并切到新会话。
8. 在 Excel / PowerPoint 里打开历史：只看到各自宿主的会话，Word 会话不出现；升级前的老会话（无宿主标记）仍能看到。

查询 DB 复核：

```sql
SELECT s.session_id, s.app_type, s.document_name, s.custom_title, s.created_at
FROM session_summary s
ORDER BY s.created_at DESC LIMIT 20;

SELECT session_id, COUNT(*) AS turns, MIN(created_at) AS first_turn
FROM conversation_event
WHERE event_type = 'chat_message'
GROUP BY session_id
ORDER BY first_turn DESC LIMIT 20;
```

前者每会话一行、`custom_title` 为空表示仍在用自动标题；后者行数与列表条目一致（列表以 `conversation_event` 为唯一权威源，档案行缺失也不影响可见性）。

### 6.3 E3-03 未配置阻断

前置：把当前服务商的 API Key 清空（配置窗口里删掉保存即可），重启宿主让前端静态文件刷新。

1. 在输入框随便写一句话点发送：消息**不进入会话区**，聊天区出现琥珀色 CTA 卡片，标题“需要先配置大模型”，正文说明“消息没有发送，文档不会被修改”。
2. 重复点发送三次：仍然只有一张卡片（固定 id 刷新，不叠卡）。
3. 点卡片上的“配置 API”：直接打开既有配置窗口，不需要自己找 Ribbon 入口。
4. 在配置窗口保存成功后：CTA 卡片自动消失，面板顶部模型名刷新。
5. 切到 Agent 模式，在 Key 仍为空时点“开始执行”：任务不启动，卡片文案变成“任务没有开始执行，文档不会被修改”，会话区不出现空的思考气泡。
6. 点“查看使用引导”：CTA 消失并唤起 E3-01 卡片（此时向导停在示例步，因为还没配好模型）。
7. 只填 Key、把 URL 清空再发送：卡片原因变成缺 URL（提示指向正确的字段，不是笼统的“配置有误”）。
8. 配置齐全后先制造一次正常回复：确认上一轮遗留的 CTA 会被 AI 回复自愈清掉。

### 6.4 E3-01 首次使用引导（收敛为两步后）

前置：新建一篇 Word 文档，删掉 `%APPDATA%\AiHelper\secure.config` 里的 `PrivacyPolicyAccepted` 这一行（或整行改 false），重启宿主让前端副本刷新。

1. 面板打开约 1 秒后出现向导卡片，进度点 **2 个**，起点是“隐私确认”，三条数据边界说明（发给所选模型接口 / 本机会话与记忆 / Key 只进 DPAPI）。
2. 点“同意并继续”：卡片立刻由宿主重推为“试一个任务”，前端没有自己的步骤状态，不存在“卡片停在旧步骤”的情况。
3. 未配置模型时示例步骤顶部有提示“还没有可用的模型配置…”，并给出“打开配置窗口”按钮；点它直接弹出 `ConfigApiForm`，面板内不再有任何填 Key 的输入框。
4. 在配置窗口里选服务商、填 Key、选模型并保存：窗口关闭后面板顶部显示当前模型，向导卡片自动收敛到示例步（已配置且已同意时卡片自动收起）。
5. 配置窗口里故意填一个错误 Key 点“验证并保存”：弹窗必须带上原因（如“接口拒绝了当前 API Key（HTTP 401）…”），不能只剩“请检查 API Key 和模型名称是否正确”。
6. 点一条示例任务：文本填入输入框、向导收尾（`onboardingFinish` 只刷新模型显示，不写任何配置）。
7. 反向验证“向导不再写配置”：整条向导走完，`%USERPROFILE%\Documents\OfficeAiAppData\office_ai_config.json` 的字节必须和进入向导之前一致（冒烟里就是这条断言）。
8. 磁盘复核：配置文件中所有 `key` 只能是 `[encrypted]` 或空；真实 Key 只出现在 `%APPDATA%\AiHelper\secure.config`（DPAPI 密文）。
9. 关闭向导（×）或点“稍后再说”后再次发送消息：E3-03 的 CTA 仍可再次唤起向导，不会因为关过一次就找不回来。
10. 已配置且已同意隐私的机器上重启面板：向导**不出现**，不打扰老用户。
11. 配置窗口里把当前服务商的 URL 和 Key 都清空并保存：面板顶部模型名变空、发送被 E3-03 拦住（验证 `ApplyConfigSnapshot` 的无条件重置，不再拿着旧端点发请求）。
