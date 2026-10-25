-- E2-04 会话归属收口：把 app_type 变成一份可以解释的合同，而不是一堆各版本各写一套的历史字符串。
--
-- 历史列表（ConversationRepository.GetRecentSessions）唯一的过滤条件是
-- 「e.app_type = 当前宿主」或「e.app_type 为空串（通用会话，三端都可见）」。
-- 于是任何第三种值都不等于任何宿主、又不是空串，那条会话在 Word/Excel/PowerPoint 里都查不到——
-- 用户看到的现象就是「升级后历史整段消失」。早期 Excel 版本写进库里的 "Unknown" 正是这种值。
--
-- 归属策略（与 Week1 已声明的口径一致，这里只是补齐证据来源并把它变成可校验的）：
--   1. 有宿主证据 → 归到那个宿主：agent_run（该会话最近一条有宿主的轨迹，与迁移 14 同一条规则）
--      → agent_run_change（同一条规则，覆盖只有变更日志没有轨迹头的会话）。
--   2. 没有轨迹但有会话档案的文档名 → 按扩展名推断宿主；.csv 故意不推断（Excel 能打开但不独占），
--      未保存文档（「文档1」没有扩展名）也不推断。推断不出来就不猜。
--   3. 两者都没有 → 保持空串，继续对三端可见。老会话宁可多端可见，也不能在升级后凭空消失。
--   4. 白名单之外的任何值（含大小写与首尾空白变体）一律归一为空串，绝不允许留在库里。
--      归一覆盖带 app_type 的全部六张表：conversation_event / agent_run / agent_run_change /
--      session_summary / atomic_memory / memory_item。后两张是记忆链，读侧豁免规则与历史链相同，
--      所以同一批 "Unknown" 也让老记忆在三端同时失效。
--
-- 写侧由 ShareRibbon.AppTypeContract.Normalize 收口（六处 @app_type 绑定全部过它），
-- 因此本迁移跑完之后不会再长回第三种值。
-- 每一步都只填空、不覆盖已解析出的归属，因此整段可重复执行。

-- 1) 归一：只承认三个规范名，其余（"Unknown"/"unknown "/"word"…）压成规范名或空串。
UPDATE conversation_event
SET app_type = CASE LOWER(TRIM(IFNULL(app_type, '')))
        WHEN 'word' THEN 'Word'
        WHEN 'excel' THEN 'Excel'
        WHEN 'powerpoint' THEN 'PowerPoint'
        ELSE ''
    END
WHERE IFNULL(app_type, '') <> ''
  AND IFNULL(app_type, '') NOT IN ('Word', 'Excel', 'PowerPoint');

UPDATE agent_run
SET app_type = CASE LOWER(TRIM(IFNULL(app_type, '')))
        WHEN 'word' THEN 'Word'
        WHEN 'excel' THEN 'Excel'
        WHEN 'powerpoint' THEN 'PowerPoint'
        ELSE ''
    END
WHERE IFNULL(app_type, '') <> ''
  AND IFNULL(app_type, '') NOT IN ('Word', 'Excel', 'PowerPoint');

UPDATE agent_run_change
SET app_type = CASE LOWER(TRIM(IFNULL(app_type, '')))
        WHEN 'word' THEN 'Word'
        WHEN 'excel' THEN 'Excel'
        WHEN 'powerpoint' THEN 'PowerPoint'
        ELSE ''
    END
WHERE IFNULL(app_type, '') <> ''
  AND IFNULL(app_type, '') NOT IN ('Word', 'Excel', 'PowerPoint');

UPDATE session_summary
SET app_type = CASE LOWER(TRIM(IFNULL(app_type, '')))
        WHEN 'word' THEN 'Word'
        WHEN 'excel' THEN 'Excel'
        WHEN 'powerpoint' THEN 'PowerPoint'
        ELSE ''
    END
WHERE IFNULL(app_type, '') <> ''
  AND IFNULL(app_type, '') NOT IN ('Word', 'Excel', 'PowerPoint');

-- 记忆表读的是同一份口径：MemoryRepository / AgentMemoryRepository 的过滤条件都是
-- 「app_type = 当前宿主 OR IS NULL OR = ''」，所以历史里那批 "Unknown" 记忆行在三端都召不回来。
-- 本机实测安装版库里 atomic_memory 231 行有 123 行是 "Unknown"，属于同一个「升级后静默丢失」缺陷族。
UPDATE atomic_memory
SET app_type = CASE LOWER(TRIM(IFNULL(app_type, '')))
        WHEN 'word' THEN 'Word'
        WHEN 'excel' THEN 'Excel'
        WHEN 'powerpoint' THEN 'PowerPoint'
        ELSE ''
    END
WHERE IFNULL(app_type, '') <> ''
  AND IFNULL(app_type, '') NOT IN ('Word', 'Excel', 'PowerPoint');

UPDATE memory_item
SET app_type = CASE LOWER(TRIM(IFNULL(app_type, '')))
        WHEN 'word' THEN 'Word'
        WHEN 'excel' THEN 'Excel'
        WHEN 'powerpoint' THEN 'PowerPoint'
        ELSE ''
    END
WHERE IFNULL(app_type, '') <> ''
  AND IFNULL(app_type, '') NOT IN ('Word', 'Excel', 'PowerPoint');

-- 2) 证据一：该会话最近一条带宿主的 Agent 轨迹。归一之后重跑，因此「最新一条是 Unknown、
--    更早一条是 Excel」的会话这次能落到 Excel，而不是继续背着一个谁都认不出的值。
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

-- 3) 证据二：变更日志里的宿主（部分老会话只有 agent_run_change，没有 agent_run 头）。
UPDATE conversation_event
SET app_type = IFNULL((
        SELECT c.app_type FROM agent_run_change c
        WHERE c.session_id = conversation_event.session_id
          AND IFNULL(c.app_type, '') <> ''
        ORDER BY c.observed_at DESC
        LIMIT 1
    ), '')
WHERE IFNULL(app_type, '') = ''
  AND IFNULL(session_id, '') <> '';

-- 4) 证据三：会话档案的文档扩展名先补档案行自己。
UPDATE session_summary
SET app_type = 'Word'
WHERE IFNULL(app_type, '') = ''
  AND TRIM(IFNULL(document_name, '')) <> ''
  AND (TRIM(IFNULL(document_name, '')) LIKE '%.doc'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.docx'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.docm'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.dot'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.dotx'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.rtf'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.wps');

UPDATE session_summary
SET app_type = 'Excel'
WHERE IFNULL(app_type, '') = ''
  AND TRIM(IFNULL(document_name, '')) <> ''
  AND (TRIM(IFNULL(document_name, '')) LIKE '%.xls'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.xlsx'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.xlsm'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.xlt'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.et');

UPDATE session_summary
SET app_type = 'PowerPoint'
WHERE IFNULL(app_type, '') = ''
  AND TRIM(IFNULL(document_name, '')) <> ''
  AND (TRIM(IFNULL(document_name, '')) LIKE '%.ppt'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.pptx'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.pptm'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.pps'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.ppsx'
    OR TRIM(IFNULL(document_name, '')) LIKE '%.dps');

-- 5) 档案行的归属下发给同一会话里仍未解析的事件行。
UPDATE conversation_event
SET app_type = (
        SELECT p.app_type FROM session_summary p
        WHERE p.session_id = conversation_event.session_id
          AND IFNULL(p.app_type, '') <> ''
        LIMIT 1
    )
WHERE IFNULL(conversation_event.app_type, '') = ''
  AND IFNULL(conversation_event.session_id, '') <> ''
  AND EXISTS (
        SELECT 1 FROM session_summary p
        WHERE p.session_id = conversation_event.session_id
          AND IFNULL(p.app_type, '') <> ''
    );

-- 6) 反向收口：事件行已由轨迹确定宿主、档案行还空着的会话，把归属补进档案行。
--    会话列表按文档/宿主分组读的是 session_summary（E2-05），这里不补就会显示成「通用会话」。
UPDATE session_summary
SET app_type = (
        SELECT e.app_type FROM conversation_event e
        WHERE e.session_id = session_summary.session_id
          AND IFNULL(e.app_type, '') <> ''
        ORDER BY e.created_at DESC
        LIMIT 1
    )
WHERE IFNULL(session_summary.app_type, '') = ''
  AND IFNULL(session_summary.session_id, '') <> ''
  AND EXISTS (
        SELECT 1 FROM conversation_event e
        WHERE e.session_id = session_summary.session_id
          AND IFNULL(e.app_type, '') <> ''
    );

UPDATE schema_version SET version = 19;
