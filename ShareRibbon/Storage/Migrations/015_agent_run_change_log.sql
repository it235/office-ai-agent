-- 单轮变更日志：一次 Agent 任务结束后，按 run_id 列出本轮所有写操作（目标、摘要、是否变更、
-- 是否成功、是否部分应用、撤销依据）。失败与取消同样写入已落地的记录，供“已写入 N 处”说明
-- 和后续“撤销本轮”使用。该表只存文本与引用，不存截图或 Base64 证据。
CREATE TABLE IF NOT EXISTS agent_run_change (
    change_id TEXT PRIMARY KEY,
    run_id TEXT NOT NULL,
    session_id TEXT,
    app_type TEXT,
    seq INTEGER NOT NULL,
    tool_id TEXT,
    step_index INTEGER,
    step_description TEXT,
    summary TEXT,
    target_refs_json TEXT,
    changed INTEGER NOT NULL DEFAULT 0,
    success INTEGER NOT NULL DEFAULT 0,
    partial_apply INTEGER NOT NULL DEFAULT 0,
    undo_available INTEGER NOT NULL DEFAULT 0,
    undo_point_id TEXT,
    rollback_error TEXT,
    observed_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_agent_run_change_run ON agent_run_change(run_id, seq);
CREATE INDEX IF NOT EXISTS idx_agent_run_change_session ON agent_run_change(session_id, observed_at);

UPDATE schema_version SET version = 15;
