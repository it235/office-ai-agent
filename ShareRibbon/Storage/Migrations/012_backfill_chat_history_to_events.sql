-- 聊天历史唯一化：conversation 曾是普通聊天路径的写入表，conversation_event 是唯一权威源。
-- 升级时把没有任何 chat_message 事件的旧会话整段回填到 conversation_event，
-- 这样历史列表与会话加载读的是同一批行，老会话不会因为换源而从列表中消失。
INSERT INTO conversation_event (event_id, session_id, event_type, role, content, created_at)
SELECT lower(hex(randomblob(16))), c.session_id, 'chat_message', c.role, c.content, c.create_time
FROM conversation c
WHERE IFNULL(c.session_id, '') <> ''
  AND NOT EXISTS (
    SELECT 1 FROM conversation_event e
    WHERE e.session_id = c.session_id AND e.event_type = 'chat_message'
  );

UPDATE schema_version SET version = 12;
