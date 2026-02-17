# 实现任务清单

## 1. 数据库与存储层

- [x] 1.1 编写 SQLite 迁移脚本，新增 atomic_memory、user_profile、session_summary、conversation 表（字段 snake_case）
- [x] 1.2 若 prompt_template 表缺少 extra_json、is_skill 等字段，扩展表结构
- [x] 1.3 在 ShareRibbon 存储层提供上述表的 CRUD 访问接口

## 2. MemoryService

- [x] 2.1 创建 MemoryService 类，实现 GetRelevantMemories(query, topN, timeRange)（初版关键词+时间）
- [x] 2.2 实现 GetUserProfile() 与 GetRecentSessionSummaries(limit)
- [x] 2.3 实现 SaveAtomicMemoryAsync(userPrompt, assistantReply, sessionId) 及去重/冲突检测
- [x] 2.4 实现 SearchMemories(keyword, timeRange) 供可选 MCP 工具
- [x] 2.5 在 ConfigManager 或新建 MemoryConfig 中支持 memory.* 配置项读写

## 3. ChatStateService 与会话持久化

- [x] 3.1 为 ChatStateService 增加 CurrentSessionId 属性，新建会话时生成 GUID
- [x] 3.2 在「新建会话」流程中清空 ChatStateService 缓冲区并生成新 session_id
- [x] 3.3 实现 conversation 表写入：user 消息发送前、assistant 流式响应完成后
- [x] 3.4 确认 ManageHistorySize 按 context_limit 正确裁剪（与现有逻辑兼容）

## 4. 提示词与 Skills 加载

- [x] 4.1 实现按 scenario（excel/word/ppt/common）从 prompt_template 加载系统提示词
- [x] 4.2 实现按 scenario 与 supported_apps 加载已启用 Skills（is_skill=1）
- [x] 4.3 实现变量替换：{{选中内容}}、{{operation}} 等占位符由调用方传入并替换

## 5. ContextBuilder

- [x] 5.1 创建 ContextBuilder（或 ChatContextBuilder）类
- [x] 5.2 实现按 [0]～[6] 顺序组装消息：System、场景指令+Skills、Session Metadata、用户记忆 RAG、会话摘要、当前会话、本条消息
- [x] 5.3 支持 Memory 未启用时的降级（跳过 [3][4]）
- [x] 5.4 产出 List(Of Message) 供 HttpStreamService 使用

## 6. HttpStreamService 与 BaseChatControl 集成

- [x] 6.1 修改请求入口：接收「预组装的 messages」而非仅从 ChatStateService 读取
- [x] 6.2 在 BaseChatControl 发请求前调用 ContextBuilder 产出 messages，再传入 HttpStreamService
- [x] 6.3 流式响应完成后，Task.Run 调用 MemoryService.SaveAtomicMemoryAsync（fire-and-forget）
- [x] 6.4 实现「收藏回答」时更新 conversation.is_collected=1

## 7. Skills 配置窗口

- [x] 7.1 在 Ribbon「提示词配置」组添加/确认「提示词配置」按钮入口
- [x] 7.2 实现配置窗口：左侧场景筛选，主区域列表（系统提示词 + Skills）
- [x] 7.3 支持系统提示词与 Skills 的增删改；Skill 编辑时展示 parameters、supported_apps
- [x] 7.4 支持从 JSON/Markdown 文件导入 Skill
- [x] 7.5 配置持久化到 prompt_template 表

## 8. 记忆配置与可观测性（可选）

- [x] 8.1 提供记忆配置入口（Ribbon 或现有设置弹窗）：rag_top_n、enable_user_profile 等
- [x] 8.2 实现原子记忆、用户画像的 CRUD 管理界面（用户可查看、编辑、删除）
- [ ] 8.3 （可选）在 Pane 中展示「当前注入的记忆」供用户透明查看

## 9. 可选扩展

- [ ] 9.1 将记忆搜索暴露为 MCP 工具（memory.enable_agentic_search 为真时）
- [x] 9.2 会话摘要生成：新会话首条回复后写入 session_summary（snippet 可为首条 user 消息前 N 字符）

---

## 10. 阶段一：新会话 + 历史记录（Chat 内可见、可切换）

- [x] 10.1 ConversationRepository：GetMessagesBySession(sessionId)、会话列表（复用 session_summary 或 conversation 近期会话）
- [x] 10.2 ChatStateService：SwitchToSession(sessionId)，加载该会话消息并设为当前会话
- [x] 10.3 启动/打开侧边栏时拉取会话列表并展示；历史侧边栏增加「新会话」按钮
- [x] 10.4 点击某条历史会话时加载该会话消息并渲染到当前 Chat；点击「新会话」时清空当前会话并新建 session_id

## 11. 阶段二：场景/Skills 配置与记忆管理迁入 Chat 侧

- [x] 11.1 在 Chat 内（HTML/JS）增加「配置」入口（如设置面板或侧栏），替代或补充 WinForm 配置窗口
- [x] 11.2 场景与 Skills 的查看/编辑在 Chat 内完成（列表、编辑、导入），持久化仍用 prompt_template
- [x] 11.3 记忆管理（原子记忆、用户画像的查看/编辑/删除）在 Chat 内完成，不再依赖独立 WinForm

## 12. 阶段三：RAG 与意图在 UI 上的体现

- [x] 12.1 发请求前若使用了 RAG，在 Chat 中给出简短提示（如「已根据当前文档检索 N 条相关记忆」）
- [x] 12.2 若有意图识别结果，在 Chat 中展示（如「识别意图：摘要」），便于用户感知

---

## 13. 阶段四：统一智能体流程（意图 + 上下文 + Spec 规划与执行）

- [x] 13.1 发送前统一收集「内容区引用（选中/附件摘要）+ RAG 相关记忆」并注入意图/规划上下文（增强 GetContextSnapshot 或等效入口）
- [x] 13.2 意图明确后（用户确认或高置信度）：调用 LLM 产出 Spec 执行方案（JSON steps，与 Ralph 格式一致），解析并展示步骤清单
- [x] 13.3 复用 RalphLoopController 执行循环：将 Spec 步骤转为 RalphLoopStep，按步执行并在 Chat 中展示进度与结果
- [x] 13.4 Agent 模式下：意图确认后自动进入「规划 → 展示步骤 → 执行」流程；Chat 模式可保留当前「直接发送」或可选进入规划
