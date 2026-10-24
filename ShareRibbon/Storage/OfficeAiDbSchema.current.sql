-- Office AI database current schema snapshot.
-- Current application schema version: 10.
-- Runtime migrations are implemented in OfficeAiDatabase.RunVersionedMigrations().

CREATE TABLE IF NOT EXISTS schema_version (
    version INTEGER NOT NULL DEFAULT 10
);

CREATE TABLE IF NOT EXISTS atomic_memory (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp INTEGER NOT NULL,
    content TEXT NOT NULL,
    tags TEXT,
    session_id TEXT,
    create_time TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    app_type TEXT DEFAULT '',
    embedding TEXT,
    memory_type TEXT DEFAULT 'short_term',
    importance REAL DEFAULT 0.5,
    access_count INTEGER DEFAULT 0,
    last_access TEXT,
    source_type TEXT DEFAULT 'general',
    linked_memories TEXT,
    embedding_model TEXT DEFAULT '',
    embedding_dim INTEGER DEFAULT 0,
    embedding_updated_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_atomic_memory_content ON atomic_memory(content);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_timestamp ON atomic_memory(timestamp);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_app_type ON atomic_memory(app_type);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_memory_type ON atomic_memory(memory_type);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_importance ON atomic_memory(importance);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_source_type ON atomic_memory(source_type);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_rag ON atomic_memory(memory_type, app_type, timestamp);
CREATE INDEX IF NOT EXISTS idx_atomic_memory_embedding_present
    ON atomic_memory(memory_type, app_type)
    WHERE embedding IS NOT NULL AND embedding != '';

CREATE TABLE IF NOT EXISTS memory_graph (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_id INTEGER NOT NULL,
    target_id INTEGER NOT NULL,
    relation_type TEXT NOT NULL DEFAULT 'similar',
    weight REAL DEFAULT 1.0,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);

CREATE INDEX IF NOT EXISTS idx_memory_graph_source ON memory_graph(source_id);
CREATE INDEX IF NOT EXISTS idx_memory_graph_target ON memory_graph(target_id);
CREATE INDEX IF NOT EXISTS idx_memory_graph_relation ON memory_graph(relation_type);

CREATE TABLE IF NOT EXISTS user_profile (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    key TEXT NOT NULL UNIQUE,
    value TEXT NOT NULL,
    category TEXT DEFAULT 'preference',
    confidence REAL DEFAULT 0.5,
    last_updated TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    observation_count INTEGER DEFAULT 1
);

CREATE INDEX IF NOT EXISTS idx_user_profile_category ON user_profile(category);
CREATE INDEX IF NOT EXISTS idx_user_profile_key ON user_profile(key);

CREATE TABLE IF NOT EXISTS session_summary (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL,
    title TEXT,
    snippet TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    app_type TEXT NOT NULL DEFAULT '',
    document_id TEXT NOT NULL DEFAULT '',
    document_name TEXT NOT NULL DEFAULT '',
    custom_title TEXT,
    updated_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_session_summary_session ON session_summary(session_id);
CREATE UNIQUE INDEX IF NOT EXISTS idx_session_summary_session_unique ON session_summary(session_id);
CREATE INDEX IF NOT EXISTS idx_session_summary_document ON session_summary(app_type, document_id);

CREATE TABLE IF NOT EXISTS conversation (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL,
    role TEXT NOT NULL,
    content TEXT NOT NULL,
    create_time TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    is_collected INTEGER NOT NULL DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_conversation_session ON conversation(session_id);

CREATE TABLE IF NOT EXISTS conversation_branch (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    conversation_id INTEGER NOT NULL,
    parent_message_id INTEGER,
    branch_name TEXT,
    is_active INTEGER DEFAULT 0,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);

CREATE INDEX IF NOT EXISTS idx_conversation_branch_conv ON conversation_branch(conversation_id);
CREATE INDEX IF NOT EXISTS idx_conversation_branch_active ON conversation_branch(is_active);

CREATE TABLE IF NOT EXISTS conversation_event (
    event_id TEXT PRIMARY KEY,
    session_id TEXT NOT NULL,
    app_type TEXT DEFAULT '',
    document_id TEXT DEFAULT '',
    event_type TEXT NOT NULL,
    role TEXT DEFAULT '',
    content TEXT,
    metadata_json TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    processed_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_conversation_event_session ON conversation_event(session_id, created_at);
CREATE INDEX IF NOT EXISTS idx_conversation_event_scope ON conversation_event(app_type, document_id, event_type);
CREATE INDEX IF NOT EXISTS idx_conversation_event_processed ON conversation_event(processed_at);

CREATE TABLE IF NOT EXISTS memory_item (
    memory_id TEXT PRIMARY KEY,
    source_event_id TEXT,
    scope TEXT NOT NULL DEFAULT 'user',
    app_type TEXT DEFAULT '',
    document_id TEXT DEFAULT '',
    project_id TEXT DEFAULT '',
    memory_type TEXT NOT NULL DEFAULT 'fact',
    content TEXT NOT NULL,
    summary TEXT,
    confidence REAL DEFAULT 0.5,
    importance REAL DEFAULT 0.5,
    status TEXT NOT NULL DEFAULT 'active',
    expires_at TEXT,
    last_verified_at TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    FOREIGN KEY (source_event_id) REFERENCES conversation_event(event_id)
);

CREATE INDEX IF NOT EXISTS idx_memory_item_scope ON memory_item(scope, app_type, document_id, status);
CREATE INDEX IF NOT EXISTS idx_memory_item_type ON memory_item(memory_type, status, importance);
CREATE INDEX IF NOT EXISTS idx_memory_item_source ON memory_item(source_event_id);

CREATE TABLE IF NOT EXISTS memory_embedding (
    embedding_id TEXT PRIMARY KEY,
    memory_id TEXT NOT NULL,
    embedding_model TEXT NOT NULL,
    embedding_dim INTEGER DEFAULT 0,
    embedding_json TEXT,
    vector_status TEXT NOT NULL DEFAULT 'pending',
    last_error TEXT,
    updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    FOREIGN KEY (memory_id) REFERENCES memory_item(memory_id) ON DELETE CASCADE
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_memory_embedding_unique ON memory_embedding(memory_id, embedding_model);
CREATE INDEX IF NOT EXISTS idx_memory_embedding_status ON memory_embedding(vector_status, embedding_model);

CREATE TABLE IF NOT EXISTS memory_job (
    job_id TEXT PRIMARY KEY,
    job_type TEXT NOT NULL,
    target_id TEXT,
    payload_json TEXT,
    status TEXT NOT NULL DEFAULT 'pending',
    attempt_count INTEGER DEFAULT 0,
    last_error TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    next_run_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_memory_job_pending ON memory_job(status, next_run_at, created_at);
CREATE INDEX IF NOT EXISTS idx_memory_job_type ON memory_job(job_type, status);

-- 同一 job_type 加 target_id 只允许一个未终结任务，重复入队由这里挡住
CREATE UNIQUE INDEX IF NOT EXISTS idx_memory_job_active_unique
    ON memory_job(job_type, target_id)
    WHERE status IN ('pending', 'processing');

CREATE TABLE IF NOT EXISTS skills_registry (
    skill_name TEXT PRIMARY KEY,
    file_path TEXT NOT NULL,
    app_scope TEXT DEFAULT '',
    intent_types TEXT DEFAULT '',
    trigger_keywords TEXT DEFAULT '',
    description TEXT,
    embedding_json TEXT,
    usage_count INTEGER DEFAULT 0,
    success_count INTEGER DEFAULT 0,
    last_indexed_at TEXT,
    enabled INTEGER NOT NULL DEFAULT 1
);

CREATE INDEX IF NOT EXISTS idx_skills_registry_enabled ON skills_registry(enabled, app_scope);
CREATE INDEX IF NOT EXISTS idx_skills_registry_usage ON skills_registry(usage_count, success_count);

CREATE TABLE IF NOT EXISTS data_migration_marker (
    migration_key TEXT PRIMARY KEY,
    applied_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);

CREATE TABLE IF NOT EXISTS prompt_template (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_name TEXT,
    scenario TEXT,
    content TEXT,
    is_skill INTEGER NOT NULL DEFAULT 0,
    extra_json TEXT,
    sort INTEGER DEFAULT 0,
    create_time TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    update_time TEXT,
    description TEXT DEFAULT '',
    keywords TEXT DEFAULT '',
    category TEXT DEFAULT '',
    priority INTEGER DEFAULT 50,
    enabled INTEGER DEFAULT 1,
    parameters_json TEXT DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_prompt_template_scenario ON prompt_template(scenario);

CREATE TABLE IF NOT EXISTS skills_usage (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    skill_name TEXT NOT NULL,
    usage_count INTEGER DEFAULT 0,
    success_count INTEGER DEFAULT 0,
    total_tokens INTEGER DEFAULT 0,
    last_used_at TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_skills_usage_name ON skills_usage(skill_name);

CREATE TABLE IF NOT EXISTS format_template (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id TEXT NOT NULL UNIQUE,
    name TEXT NOT NULL,
    description TEXT,
    category TEXT DEFAULT '通用',
    target_app TEXT DEFAULT 'Word',
    is_preset INTEGER NOT NULL DEFAULT 0,
    template_source TEXT DEFAULT 'manual',
    source_file_name TEXT,
    source_file_content TEXT,
    source_file_blob BLOB,
    layout_json TEXT,
    style_rules_json TEXT,
    page_settings_json TEXT,
    ai_guidance TEXT,
    thumbnail_base64 TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    last_modified TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);

CREATE INDEX IF NOT EXISTS idx_format_template_target_app ON format_template(target_app);
CREATE INDEX IF NOT EXISTS idx_format_template_category ON format_template(category);

CREATE TABLE IF NOT EXISTS format_element (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id TEXT NOT NULL,
    element_name TEXT NOT NULL,
    element_type TEXT NOT NULL,
    default_value TEXT,
    is_required INTEGER DEFAULT 1,
    sort_order INTEGER DEFAULT 0,
    font_config_json TEXT,
    paragraph_config_json TEXT,
    color_config_json TEXT,
    special_props_json TEXT,
    placeholder_content TEXT,
    is_enabled INTEGER DEFAULT 1,
    FOREIGN KEY (template_id) REFERENCES format_template(template_id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_format_element_template ON format_element(template_id);

CREATE TABLE IF NOT EXISTS format_style_rule (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id TEXT NOT NULL,
    rule_name TEXT NOT NULL,
    match_condition TEXT,
    sort_order INTEGER DEFAULT 0,
    font_config_json TEXT,
    paragraph_config_json TEXT,
    color_config_json TEXT,
    is_enabled INTEGER DEFAULT 1,
    FOREIGN KEY (template_id) REFERENCES format_template(template_id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_format_style_rule_template ON format_style_rule(template_id);

CREATE TABLE IF NOT EXISTS agent_run (
    run_id TEXT PRIMARY KEY,
    turn_id TEXT,
    session_id TEXT,
    app_type TEXT,
    status TEXT,
    user_text TEXT,
    started_at TEXT,
    finished_at TEXT,
    final_message TEXT,
    error_code TEXT
);

CREATE INDEX IF NOT EXISTS idx_agent_run_session ON agent_run(session_id, started_at);
CREATE INDEX IF NOT EXISTS idx_agent_run_status ON agent_run(status, started_at);

CREATE TABLE IF NOT EXISTS agent_run_step (
    step_id TEXT PRIMARY KEY,
    run_id TEXT NOT NULL,
    seq INTEGER NOT NULL,
    tool_id TEXT,
    status TEXT,
    message TEXT,
    error_code TEXT,
    observation_json TEXT,
    started_at TEXT,
    finished_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_agent_run_step_run ON agent_run_step(run_id, seq);
CREATE INDEX IF NOT EXISTS idx_agent_run_step_status ON agent_run_step(status, error_code);

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
    rolled_back INTEGER NOT NULL DEFAULT 0,
    partial_apply INTEGER NOT NULL DEFAULT 0,
    undo_available INTEGER NOT NULL DEFAULT 0,
    undo_units INTEGER NOT NULL DEFAULT 0,
    undo_point_id TEXT,
    rollback_error TEXT,
    doc_hash_before TEXT,
    doc_hash_after TEXT,
    observed_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_agent_run_change_run ON agent_run_change(run_id, seq);
CREATE INDEX IF NOT EXISTS idx_agent_run_change_session ON agent_run_change(session_id, observed_at);
