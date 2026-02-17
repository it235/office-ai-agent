-- Office AI 记忆与会话数据库迁移脚本
-- 表名与字段名使用 snake_case
-- 首次运行创建所有表

-- 原子记忆表
CREATE TABLE IF NOT EXISTS atomic_memory (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp INTEGER NOT NULL,
    content TEXT NOT NULL,
    tags TEXT,
    session_id TEXT,
    create_time TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);

CREATE INDEX IF NOT EXISTS idx_atomic_memory_content ON atomic_memory(content);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_timestamp ON atomic_memory(timestamp);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_session ON atomic_memory(session_id);

-- 用户画像表
CREATE TABLE IF NOT EXISTS user_profile (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    content TEXT,
    updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);

-- 会话摘要表
CREATE TABLE IF NOT EXISTS session_summary (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL,
    title TEXT,
    snippet TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);

CREATE INDEX IF NOT EXISTS idx_session_summary_session ON session_summary(session_id);
CREATE INDEX IF NOT EXISTS idx_session_summary_created ON session_summary(created_at);

-- 会话消息表（conversation）
CREATE TABLE IF NOT EXISTS conversation (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL,
    role TEXT NOT NULL,
    content TEXT NOT NULL,
    create_time TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    is_collected INTEGER NOT NULL DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_conversation_session ON conversation(session_id);
CREATE INDEX IF NOT EXISTS idx_conversation_create_time ON conversation(create_time);

-- 可选：文件夹记忆表（本阶段可选，预留）
CREATE TABLE IF NOT EXISTS folder_memory (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_id TEXT NOT NULL,
    content TEXT,
    updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);
