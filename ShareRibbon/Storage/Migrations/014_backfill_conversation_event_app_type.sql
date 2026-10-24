-- 会话按宿主隔离：conversation_event.app_type 历史上多数为空，导致 Word/Excel/PowerPoint
-- 的历史会话互相串台。Agent 运行轨迹 agent_run 已按会话记录了宿主类型，用它回填同一会话的
-- 聊天事件；无轨迹可依的旧会话保持空值，继续对三个宿主可见，避免升级后历史整体消失。
UPDATE conversation_event
SET app_type = IFNULL((
    SELECT r.app_type FROM agent_run r
    WHERE r.session_id = conversation_event.session_id
      AND IFNULL(r.app_type, '') <> ''
    ORDER BY r.started_at DESC
    LIMIT 1
), '')
WHERE IFNULL(app_type, '') = ''
  AND IFNULL(session_id, '') <> '';

UPDATE schema_version SET version = 14;
