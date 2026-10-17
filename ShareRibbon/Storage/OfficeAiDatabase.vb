' ShareRibbon\Storage\OfficeAiDatabase.vb
' Office AI 数据库初始化与迁移

Imports System.Data.SQLite
Imports System.IO
Imports System.Linq
Imports Newtonsoft.Json

''' <summary>
''' Office AI SQLite 数据库初始化与迁移
''' </summary>
Public Class OfficeAiDatabase

    Private Shared _initialized As Boolean = False
    Private Shared ReadOnly _lockObj As New Object()

    ''' <summary>
    ''' 获取数据库文件路径。调试版使用 OfficeAiAppData-Debug 子目录，与安装版数据分离。
    ''' </summary>
    Public Shared Function GetDatabasePath() As String
        Dim folderName As String = ConfigSettings.OfficeAiAppDataFolder
        If IsDebugEnvironment() Then
            folderName = folderName & "-Debug"
        End If
        Dim baseDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            folderName)
        Return Path.Combine(baseDir, "office_ai.db")
    End Function

    ''' <summary>
    ''' 是否从本地调试目录运行（bin\Debug、bin\x64 等），与安装版区分
    ''' </summary>
    Private Shared Function IsDebugEnvironment() As Boolean
        Try
            Dim loc = GetType(OfficeAiDatabase).Assembly.Location
            If String.IsNullOrEmpty(loc) Then Return False
            Dim dir = Path.GetDirectoryName(loc)
            If String.IsNullOrEmpty(dir) Then Return False
            Dim lower = dir.ToLowerInvariant()
            Return lower.Contains("\bin\debug") OrElse
                   lower.Contains("\bin\x64") OrElse
                   lower.Contains("\bin\x86") OrElse
                   (lower.Contains("\bin\release") AndAlso Not lower.Contains("program files"))
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取连接字符串
    ''' </summary>
    Public Shared Function GetConnectionString() As String
        Dim dbPath = GetDatabasePath()
        Return $"Data Source={dbPath};Version=3;Pooling=True;Max Pool Size=10;Busy Timeout=5000;Journal Mode=WAL;Synchronous=NORMAL;Cache Size=-64000;"
    End Function

    ''' <summary>
    ''' 确保数据库已初始化并执行迁移
    ''' </summary>
    Public Shared Sub EnsureInitialized()
        If _initialized Then Return

        SyncLock _lockObj
            If _initialized Then Return

            Try
                SqliteAssemblyResolver.EnsureRegistered()
                SqliteNativeLoader.EnsureLoaded()
                Dim baseDir = Path.GetDirectoryName(GetDatabasePath())
                If Not String.IsNullOrEmpty(baseDir) AndAlso Not Directory.Exists(baseDir) Then
                    Directory.CreateDirectory(baseDir)
                End If

                Dim migrationSql = GetMigrationSql()
                Using conn As New SQLiteConnection(GetConnectionString())
                    conn.Open()
                    Using cmd As New SQLiteCommand(migrationSql, conn)
                        cmd.ExecuteNonQuery()
                    End Using
                    RunVersionedMigrations(conn)
                    RunSchemaHealthChecks(conn)
                End Using

                _initialized = True
            Catch ex As Exception
                Debug.WriteLine($"OfficeAiDatabase 初始化失败: {ex.Message}")
                Throw
            End Try
        End SyncLock
    End Sub

    Private Shared Function GetMigrationSql() As String
        ' 尝试从文件读取（开发时）
        Dim externalSql = TryReadMigrationSqlFile()
        If Not String.IsNullOrWhiteSpace(externalSql) Then Return externalSql

        Dim asmLoc = GetType(OfficeAiDatabase).Assembly.Location
        Dim dir = If(String.IsNullOrEmpty(asmLoc), "", Path.GetDirectoryName(asmLoc))
        Dim sqlPath = If(String.IsNullOrEmpty(dir), "OfficeAiDbMigration.sql", Path.Combine(dir, "OfficeAiDbMigration.sql"))
        If File.Exists(sqlPath) Then
            Try
                Return File.ReadAllText(sqlPath)
            Catch
            End Try
        End If

        ' 内联 SQL（基准 schema = 版本 1；新增字段通过 RunVersionedMigrations 的 ALTER 升级）
        Return "
CREATE TABLE IF NOT EXISTS schema_version (version INTEGER NOT NULL DEFAULT 1);
INSERT INTO schema_version (version) SELECT 1 WHERE NOT EXISTS (SELECT 1 FROM schema_version);
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
CREATE TABLE IF NOT EXISTS user_profile (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    content TEXT,
    updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);
CREATE TABLE IF NOT EXISTS session_summary (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL,
    title TEXT,
    snippet TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
);
CREATE INDEX IF NOT EXISTS idx_session_summary_session ON session_summary(session_id);
CREATE TABLE IF NOT EXISTS conversation (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL,
    role TEXT NOT NULL,
    content TEXT NOT NULL,
    create_time TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    is_collected INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_conversation_session ON conversation(session_id);
CREATE TABLE IF NOT EXISTS prompt_template (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_name TEXT,
    scenario TEXT,
    content TEXT,
    is_skill INTEGER NOT NULL DEFAULT 0,
    extra_json TEXT,
    sort INTEGER DEFAULT 0,
    create_time TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    update_time TEXT
);
CREATE INDEX IF NOT EXISTS idx_prompt_template_scenario ON prompt_template(scenario);
"
    End Function

    Private Shared Function TryReadMigrationSqlFile() As String
        For Each sqlPath In GetMigrationSqlCandidates()
            If String.IsNullOrWhiteSpace(sqlPath) OrElse Not File.Exists(sqlPath) Then Continue For

            Try
                Dim sql = File.ReadAllText(sqlPath)
                If Not String.IsNullOrWhiteSpace(sql) Then Return sql
            Catch ex As Exception
                Debug.WriteLine($"OfficeAiDatabase 读取迁移 SQL 文件失败: {sqlPath}, {ex.Message}")
            End Try
        Next

        Return Nothing
    End Function

    Private Shared Function GetMigrationSqlCandidates() As IEnumerable(Of String)
        Dim candidates As New List(Of String)()
        Dim asmLoc = GetType(OfficeAiDatabase).Assembly.Location
        Dim dir = If(String.IsNullOrEmpty(asmLoc), "", Path.GetDirectoryName(asmLoc))

        If Not String.IsNullOrWhiteSpace(dir) Then
            candidates.Add(Path.Combine(dir, "Storage", "Migrations", "001_baseline.sql"))
            candidates.Add(Path.Combine(dir, "OfficeAiDbMigration.sql"))
        End If

        candidates.Add("OfficeAiDbMigration.sql")
        Return candidates
    End Function

    Private Shared Function TryReadVersionedMigrationSql(version As Integer) As String
        For Each sqlPath In GetVersionedMigrationSqlCandidates(version)
            If String.IsNullOrWhiteSpace(sqlPath) OrElse Not File.Exists(sqlPath) Then Continue For

            Try
                Dim sql = File.ReadAllText(sqlPath)
                If Not String.IsNullOrWhiteSpace(sql) Then Return sql
            Catch ex As Exception
                Debug.WriteLine($"OfficeAiDatabase 读取版本迁移 SQL 文件失败: {sqlPath}, {ex.Message}")
            End Try
        Next

        Return Nothing
    End Function

    Private Shared Function GetVersionedMigrationSqlCandidates(version As Integer) As IEnumerable(Of String)
        Dim candidates As New List(Of String)()
        Dim asmLoc = GetType(OfficeAiDatabase).Assembly.Location
        Dim dir = If(String.IsNullOrEmpty(asmLoc), "", Path.GetDirectoryName(asmLoc))

        If Not String.IsNullOrWhiteSpace(dir) Then
            Dim migrationsDir = Path.Combine(dir, "Storage", "Migrations")
            If Directory.Exists(migrationsDir) Then
                candidates.AddRange(Directory.GetFiles(migrationsDir, version.ToString("000") & "_*.sql"))
                candidates.Add(Path.Combine(migrationsDir, version.ToString("000") & ".sql"))
            End If
        End If

        Return candidates
    End Function

    ''' <summary>
    ''' 按 schema_version 执行增量迁移（仅执行未应用过的版本），便于升级与版本控制。
    ''' </summary>
    Private Shared Sub RunVersionedMigrations(conn As SQLiteConnection)
        Dim currentVersion As Integer = 1
        Try
            Using cmd As New SQLiteCommand("SELECT version FROM schema_version LIMIT 1", conn)
                Dim obj = cmd.ExecuteScalar()
                If obj IsNot Nothing AndAlso Not IsDBNull(obj) Then
                    currentVersion = Convert.ToInt32(obj)
                End If
            End Using
        Catch
            ' 表不存在或为空时视为 1
        End Try

        ' 各版本迁移 SQL（仅 ALTER / CREATE INDEX / UPDATE version，不重复执行）
        Dim migrations As New Dictionary(Of Integer, String) From {
            {2, "ALTER TABLE atomic_memory ADD COLUMN app_type TEXT DEFAULT '';" &
             "CREATE INDEX IF NOT EXISTS idx_atomic_memory_app_type ON atomic_memory(app_type);" &
             "UPDATE schema_version SET version = 2;"},
            {3, "ALTER TABLE atomic_memory ADD COLUMN embedding TEXT DEFAULT NULL;" &
             "UPDATE schema_version SET version = 3;"},
            {4, "CREATE TABLE IF NOT EXISTS format_template (" &
             "id INTEGER PRIMARY KEY AUTOINCREMENT," &
             "template_id TEXT NOT NULL UNIQUE," &
             "name TEXT NOT NULL," &
             "description TEXT," &
             "category TEXT DEFAULT '通用'," &
             "target_app TEXT DEFAULT 'Word'," &
             "is_preset INTEGER NOT NULL DEFAULT 0," &
             "template_source TEXT DEFAULT 'manual'," &
             "source_file_name TEXT," &
             "source_file_content TEXT," &
             "source_file_blob BLOB," &
             "layout_json TEXT," &
             "style_rules_json TEXT," &
             "page_settings_json TEXT," &
             "ai_guidance TEXT," &
             "thumbnail_base64 TEXT," &
             "created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
             "last_modified TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))" &
             ");" &
             "CREATE INDEX IF NOT EXISTS idx_format_template_target_app ON format_template(target_app);" &
             "CREATE INDEX IF NOT EXISTS idx_format_template_category ON format_template(category);" &
             "CREATE TABLE IF NOT EXISTS format_element (" &
             "id INTEGER PRIMARY KEY AUTOINCREMENT," &
             "template_id TEXT NOT NULL," &
             "element_name TEXT NOT NULL," &
             "element_type TEXT NOT NULL," &
             "default_value TEXT," &
             "is_required INTEGER DEFAULT 1," &
             "sort_order INTEGER DEFAULT 0," &
             "font_config_json TEXT," &
             "paragraph_config_json TEXT," &
             "color_config_json TEXT," &
             "special_props_json TEXT," &
             "placeholder_content TEXT," &
             "is_enabled INTEGER DEFAULT 1," &
             "FOREIGN KEY (template_id) REFERENCES format_template(template_id) ON DELETE CASCADE" &
             ");" &
             "CREATE INDEX IF NOT EXISTS idx_format_element_template ON format_element(template_id);" &
             "CREATE TABLE IF NOT EXISTS format_style_rule (" &
             "id INTEGER PRIMARY KEY AUTOINCREMENT," &
             "template_id TEXT NOT NULL," &
             "rule_name TEXT NOT NULL," &
             "match_condition TEXT," &
             "sort_order INTEGER DEFAULT 0," &
             "font_config_json TEXT," &
             "paragraph_config_json TEXT," &
             "color_config_json TEXT," &
             "is_enabled INTEGER DEFAULT 1," &
             "FOREIGN KEY (template_id) REFERENCES format_template(template_id) ON DELETE CASCADE" &
             ");" &
             "CREATE INDEX IF NOT EXISTS idx_format_style_rule_template ON format_style_rule(template_id);" &
             "UPDATE schema_version SET version = 4;"},
            {5, "ALTER TABLE prompt_template ADD COLUMN description TEXT DEFAULT '';" &
             "ALTER TABLE prompt_template ADD COLUMN keywords TEXT DEFAULT '';" &
             "ALTER TABLE prompt_template ADD COLUMN category TEXT DEFAULT '';" &
             "ALTER TABLE prompt_template ADD COLUMN priority INTEGER DEFAULT 50;" &
             "ALTER TABLE prompt_template ADD COLUMN enabled INTEGER DEFAULT 1;" &
             "ALTER TABLE prompt_template ADD COLUMN parameters_json TEXT DEFAULT '';" &
             "UPDATE schema_version SET version = 5;"},
            {6, "ALTER TABLE atomic_memory ADD COLUMN memory_type TEXT DEFAULT 'short_term';" &
             "CREATE INDEX IF NOT EXISTS idx_atomic_memory_memory_type ON atomic_memory(memory_type);" &
             "UPDATE schema_version SET version = 6;"},
            {7, "CREATE TABLE IF NOT EXISTS skills_usage (" &
             "id INTEGER PRIMARY KEY AUTOINCREMENT," &
             "skill_name TEXT NOT NULL," &
             "usage_count INTEGER DEFAULT 0," &
             "success_count INTEGER DEFAULT 0," &
             "total_tokens INTEGER DEFAULT 0," &
             "last_used_at TEXT," &
             "created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
             "updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))" &
             ");" &
             "CREATE UNIQUE INDEX IF NOT EXISTS idx_skills_usage_name ON skills_usage(skill_name);" &
             "UPDATE schema_version SET version = 7;"},
            {8, "CREATE TABLE IF NOT EXISTS atomic_memory_new (" &
             "id INTEGER PRIMARY KEY AUTOINCREMENT," &
             "timestamp INTEGER NOT NULL," &
             "content TEXT NOT NULL," &
             "tags TEXT," &
             "session_id TEXT," &
             "create_time TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
             "app_type TEXT DEFAULT ''," &
             "embedding TEXT," &
             "memory_type TEXT DEFAULT 'short_term'," &
             "importance REAL DEFAULT 0.5," &
             "access_count INTEGER DEFAULT 0," &
             "last_access TEXT," &
             "source_type TEXT DEFAULT 'general'," &
             "linked_memories TEXT" &
             ");" &
             "INSERT OR IGNORE INTO atomic_memory_new (id, timestamp, content, tags, session_id, create_time, app_type, embedding, memory_type)" &
             "SELECT id, timestamp, content, tags, session_id, create_time, app_type, embedding, memory_type FROM atomic_memory;" &
             "DROP TABLE atomic_memory;" &
             "ALTER TABLE atomic_memory_new RENAME TO atomic_memory;" &
             "CREATE INDEX IF NOT EXISTS idx_atomic_memory_content ON atomic_memory(content);" &
             "CREATE INDEX IF NOT EXISTS idx_atomic_memory_timestamp ON atomic_memory(timestamp);" &
             "CREATE INDEX IF NOT EXISTS idx_atomic_memory_app_type ON atomic_memory(app_type);" &
             "CREATE INDEX IF NOT EXISTS idx_atomic_memory_memory_type ON atomic_memory(memory_type);" &
             "CREATE INDEX IF NOT EXISTS idx_atomic_memory_importance ON atomic_memory(importance);" &
             "CREATE INDEX IF NOT EXISTS idx_atomic_memory_source_type ON atomic_memory(source_type);" &
             "CREATE TABLE IF NOT EXISTS memory_graph (" &
             "id INTEGER PRIMARY KEY AUTOINCREMENT," &
             "source_id INTEGER NOT NULL," &
             "target_id INTEGER NOT NULL," &
             "relation_type TEXT NOT NULL DEFAULT 'similar'," &
             "weight REAL DEFAULT 1.0," &
             "created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))" &
             ");" &
             "CREATE INDEX IF NOT EXISTS idx_memory_graph_source ON memory_graph(source_id);" &
             "CREATE INDEX IF NOT EXISTS idx_memory_graph_target ON memory_graph(target_id);" &
             "CREATE INDEX IF NOT EXISTS idx_memory_graph_relation ON memory_graph(relation_type);" &
             "CREATE TABLE IF NOT EXISTS user_profile_new (" &
             "id INTEGER PRIMARY KEY AUTOINCREMENT," &
             "key TEXT NOT NULL UNIQUE," &
             "value TEXT NOT NULL," &
             "category TEXT DEFAULT 'preference'," &
             "confidence REAL DEFAULT 0.5," &
             "last_updated TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
             "observation_count INTEGER DEFAULT 1" &
             ");" &
             "INSERT OR IGNORE INTO user_profile_new (key, value) " &
             "SELECT 'legacy_data', content FROM user_profile WHERE content IS NOT NULL LIMIT 1;" &
             "DROP TABLE IF EXISTS user_profile;" &
             "ALTER TABLE user_profile_new RENAME TO user_profile;" &
             "CREATE INDEX IF NOT EXISTS idx_user_profile_category ON user_profile(category);" &
             "CREATE INDEX IF NOT EXISTS idx_user_profile_key ON user_profile(key);" &
             "CREATE TABLE IF NOT EXISTS conversation_branch (" &
             "id INTEGER PRIMARY KEY AUTOINCREMENT," &
             "conversation_id INTEGER NOT NULL," &
             "parent_message_id INTEGER," &
             "branch_name TEXT," &
             "is_active INTEGER DEFAULT 0," &
             "created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))" &
             ");" &
             "CREATE INDEX IF NOT EXISTS idx_conversation_branch_conv ON conversation_branch(conversation_id);" &
             "CREATE INDEX IF NOT EXISTS idx_conversation_branch_active ON conversation_branch(is_active);" &
             "UPDATE schema_version SET version = 8;"},
            {9, "ALTER TABLE atomic_memory ADD COLUMN embedding_model TEXT DEFAULT '';" &
             "ALTER TABLE atomic_memory ADD COLUMN embedding_dim INTEGER DEFAULT 0;" &
             "ALTER TABLE atomic_memory ADD COLUMN embedding_updated_at TEXT;" &
             "CREATE INDEX IF NOT EXISTS idx_atomic_memory_rag ON atomic_memory(memory_type, app_type, timestamp);" &
             "CREATE INDEX IF NOT EXISTS idx_atomic_memory_embedding_present ON atomic_memory(memory_type, app_type) WHERE embedding IS NOT NULL AND embedding != '';" &
             "UPDATE schema_version SET version = 9;"},
            {10, "CREATE TABLE IF NOT EXISTS conversation_event (" &
             "event_id TEXT PRIMARY KEY," &
             "session_id TEXT NOT NULL," &
             "app_type TEXT DEFAULT ''," &
             "document_id TEXT DEFAULT ''," &
             "event_type TEXT NOT NULL," &
             "role TEXT DEFAULT ''," &
             "content TEXT," &
             "metadata_json TEXT," &
             "created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
             "processed_at TEXT" &
             ");" &
             "CREATE INDEX IF NOT EXISTS idx_conversation_event_session ON conversation_event(session_id, created_at);" &
             "CREATE INDEX IF NOT EXISTS idx_conversation_event_scope ON conversation_event(app_type, document_id, event_type);" &
             "CREATE INDEX IF NOT EXISTS idx_conversation_event_processed ON conversation_event(processed_at);" &
             "CREATE TABLE IF NOT EXISTS memory_item (" &
             "memory_id TEXT PRIMARY KEY," &
             "source_event_id TEXT," &
             "scope TEXT NOT NULL DEFAULT 'user'," &
             "app_type TEXT DEFAULT ''," &
             "document_id TEXT DEFAULT ''," &
             "project_id TEXT DEFAULT ''," &
             "memory_type TEXT NOT NULL DEFAULT 'fact'," &
             "content TEXT NOT NULL," &
             "summary TEXT," &
             "confidence REAL DEFAULT 0.5," &
             "importance REAL DEFAULT 0.5," &
             "status TEXT NOT NULL DEFAULT 'active'," &
             "expires_at TEXT," &
             "last_verified_at TEXT," &
             "created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
             "updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
             "FOREIGN KEY (source_event_id) REFERENCES conversation_event(event_id)" &
             ");" &
             "CREATE INDEX IF NOT EXISTS idx_memory_item_scope ON memory_item(scope, app_type, document_id, status);" &
             "CREATE INDEX IF NOT EXISTS idx_memory_item_type ON memory_item(memory_type, status, importance);" &
             "CREATE INDEX IF NOT EXISTS idx_memory_item_source ON memory_item(source_event_id);" &
             "CREATE TABLE IF NOT EXISTS memory_embedding (" &
             "embedding_id TEXT PRIMARY KEY," &
             "memory_id TEXT NOT NULL," &
             "embedding_model TEXT NOT NULL," &
             "embedding_dim INTEGER DEFAULT 0," &
             "embedding_json TEXT," &
             "vector_status TEXT NOT NULL DEFAULT 'pending'," &
             "last_error TEXT," &
             "updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
             "FOREIGN KEY (memory_id) REFERENCES memory_item(memory_id) ON DELETE CASCADE" &
             ");" &
             "CREATE UNIQUE INDEX IF NOT EXISTS idx_memory_embedding_unique ON memory_embedding(memory_id, embedding_model);" &
             "CREATE INDEX IF NOT EXISTS idx_memory_embedding_status ON memory_embedding(vector_status, embedding_model);" &
             "CREATE TABLE IF NOT EXISTS memory_job (" &
             "job_id TEXT PRIMARY KEY," &
             "job_type TEXT NOT NULL," &
             "target_id TEXT," &
             "payload_json TEXT," &
             "status TEXT NOT NULL DEFAULT 'pending'," &
             "attempt_count INTEGER DEFAULT 0," &
             "last_error TEXT," &
             "created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
             "updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
             "next_run_at TEXT" &
             ");" &
             "CREATE INDEX IF NOT EXISTS idx_memory_job_pending ON memory_job(status, next_run_at, created_at);" &
             "CREATE INDEX IF NOT EXISTS idx_memory_job_type ON memory_job(job_type, status);" &
             "CREATE TABLE IF NOT EXISTS skills_registry (" &
             "skill_name TEXT PRIMARY KEY," &
             "file_path TEXT NOT NULL," &
             "app_scope TEXT DEFAULT ''," &
             "intent_types TEXT DEFAULT ''," &
             "trigger_keywords TEXT DEFAULT ''," &
             "description TEXT," &
             "embedding_json TEXT," &
             "usage_count INTEGER DEFAULT 0," &
             "success_count INTEGER DEFAULT 0," &
             "last_indexed_at TEXT," &
             "enabled INTEGER NOT NULL DEFAULT 1" &
             ");" &
             "CREATE INDEX IF NOT EXISTS idx_skills_registry_enabled ON skills_registry(enabled, app_scope);" &
             "CREATE INDEX IF NOT EXISTS idx_skills_registry_usage ON skills_registry(usage_count, success_count);" &
             "UPDATE schema_version SET version = 10;"},
            {11, "CREATE TABLE IF NOT EXISTS agent_run (" &
             "run_id TEXT PRIMARY KEY," &
             "turn_id TEXT," &
             "session_id TEXT," &
             "app_type TEXT," &
             "status TEXT," &
             "user_text TEXT," &
             "started_at TEXT," &
             "finished_at TEXT," &
             "final_message TEXT," &
             "error_code TEXT" &
             ");" &
             "CREATE INDEX IF NOT EXISTS idx_agent_run_session ON agent_run(session_id, started_at);" &
             "CREATE INDEX IF NOT EXISTS idx_agent_run_status ON agent_run(status, started_at);" &
             "CREATE TABLE IF NOT EXISTS agent_run_step (" &
             "step_id TEXT PRIMARY KEY," &
             "run_id TEXT NOT NULL," &
             "seq INTEGER NOT NULL," &
             "tool_id TEXT," &
             "status TEXT," &
             "message TEXT," &
             "error_code TEXT," &
             "observation_json TEXT," &
             "started_at TEXT," &
             "finished_at TEXT" &
             ");" &
             "CREATE INDEX IF NOT EXISTS idx_agent_run_step_run ON agent_run_step(run_id, seq);" &
             "CREATE INDEX IF NOT EXISTS idx_agent_run_step_status ON agent_run_step(status, error_code);" &
             "UPDATE schema_version SET version = 11;"}
        }

        For Each kvp In migrations.OrderBy(Function(x) x.Key)
            If kvp.Key <= currentVersion Then Continue For
            Dim migrationSql = TryReadVersionedMigrationSql(kvp.Key)
            If String.IsNullOrWhiteSpace(migrationSql) Then
                migrationSql = kvp.Value
            End If

            Dim transaction = conn.BeginTransaction()
            Try
                Using cmd As New SQLiteCommand(migrationSql, conn, transaction)
                    cmd.ExecuteNonQuery()
                End Using
                transaction.Commit()
                currentVersion = kvp.Key
                Debug.WriteLine($"OfficeAiDatabase 已应用迁移版本 {kvp.Key}")
            Catch ex As Exception
                Try
                    transaction.Rollback()
                Catch rollbackEx As Exception
                    Debug.WriteLine($"迁移回滚失败: {rollbackEx.Message}")
                End Try
                Debug.WriteLine($"迁移版本 {kvp.Key} 失败: {ex.Message}")
                Throw
            End Try
        Next
    End Sub

    Private Shared Sub RunSchemaHealthChecks(conn As SQLiteConnection)
        EnsureAtomicMemoryColumns(conn)
        EnsureUserProfileSchema(conn)
        EnsureAgentMemorySchema(conn)
        ' conversation_event is now the single source of truth for chat history; keep the
        ' collection flag on it too so the UI can star an assistant turn without a parallel table.
        EnsureColumn(conn, "conversation_event", "is_collected", "is_collected INTEGER NOT NULL DEFAULT 0")
        EnsureAgentRunTraceSchema(conn)
        EnsureDataMigrationMarkers(conn)
        MigrateLegacySkillUsage(conn)

        Using cmd As New SQLiteCommand("CREATE INDEX IF NOT EXISTS idx_atomic_memory_rag ON atomic_memory(memory_type, app_type, timestamp);" &
                                       "CREATE INDEX IF NOT EXISTS idx_atomic_memory_embedding_present ON atomic_memory(memory_type, app_type) WHERE embedding IS NOT NULL AND embedding != '';" &
                                       "CREATE INDEX IF NOT EXISTS idx_user_profile_category ON user_profile(category);" &
                                       "CREATE INDEX IF NOT EXISTS idx_user_profile_key ON user_profile(key);", conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Sub EnsureAgentRunTraceSchema(conn As SQLiteConnection)
        Using cmd As New SQLiteCommand(
            "CREATE TABLE IF NOT EXISTS agent_run (" &
            "run_id TEXT PRIMARY KEY," &
            "turn_id TEXT," &
            "session_id TEXT," &
            "app_type TEXT," &
            "status TEXT," &
            "user_text TEXT," &
            "started_at TEXT," &
            "finished_at TEXT," &
            "final_message TEXT," &
            "error_code TEXT" &
            ");" &
            "CREATE INDEX IF NOT EXISTS idx_agent_run_session ON agent_run(session_id, started_at);" &
            "CREATE INDEX IF NOT EXISTS idx_agent_run_status ON agent_run(status, started_at);" &
            "CREATE TABLE IF NOT EXISTS agent_run_step (" &
            "step_id TEXT PRIMARY KEY," &
            "run_id TEXT NOT NULL," &
            "seq INTEGER NOT NULL," &
            "tool_id TEXT," &
            "status TEXT," &
            "message TEXT," &
            "error_code TEXT," &
            "observation_json TEXT," &
            "started_at TEXT," &
            "finished_at TEXT" &
            ");" &
            "CREATE INDEX IF NOT EXISTS idx_agent_run_step_run ON agent_run_step(run_id, seq);" &
            "CREATE INDEX IF NOT EXISTS idx_agent_run_step_status ON agent_run_step(status, error_code);", conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Sub EnsureAgentMemorySchema(conn As SQLiteConnection)
        Using cmd As New SQLiteCommand(
            "CREATE TABLE IF NOT EXISTS conversation_event (" &
            "event_id TEXT PRIMARY KEY," &
            "session_id TEXT NOT NULL," &
            "app_type TEXT DEFAULT ''," &
            "document_id TEXT DEFAULT ''," &
            "event_type TEXT NOT NULL," &
            "role TEXT DEFAULT ''," &
            "content TEXT," &
            "metadata_json TEXT," &
            "created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
            "processed_at TEXT" &
            ");" &
            "CREATE INDEX IF NOT EXISTS idx_conversation_event_session ON conversation_event(session_id, created_at);" &
            "CREATE INDEX IF NOT EXISTS idx_conversation_event_scope ON conversation_event(app_type, document_id, event_type);" &
            "CREATE INDEX IF NOT EXISTS idx_conversation_event_processed ON conversation_event(processed_at);" &
            "CREATE TABLE IF NOT EXISTS memory_item (" &
            "memory_id TEXT PRIMARY KEY," &
            "source_event_id TEXT," &
            "scope TEXT NOT NULL DEFAULT 'user'," &
            "app_type TEXT DEFAULT ''," &
            "document_id TEXT DEFAULT ''," &
            "project_id TEXT DEFAULT ''," &
            "memory_type TEXT NOT NULL DEFAULT 'fact'," &
            "content TEXT NOT NULL," &
            "summary TEXT," &
            "confidence REAL DEFAULT 0.5," &
            "importance REAL DEFAULT 0.5," &
            "status TEXT NOT NULL DEFAULT 'active'," &
            "expires_at TEXT," &
            "last_verified_at TEXT," &
            "created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
            "updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
            "FOREIGN KEY (source_event_id) REFERENCES conversation_event(event_id)" &
            ");" &
            "CREATE INDEX IF NOT EXISTS idx_memory_item_scope ON memory_item(scope, app_type, document_id, status);" &
            "CREATE INDEX IF NOT EXISTS idx_memory_item_type ON memory_item(memory_type, status, importance);" &
            "CREATE INDEX IF NOT EXISTS idx_memory_item_source ON memory_item(source_event_id);" &
            "CREATE TABLE IF NOT EXISTS memory_embedding (" &
            "embedding_id TEXT PRIMARY KEY," &
            "memory_id TEXT NOT NULL," &
            "embedding_model TEXT NOT NULL," &
            "embedding_dim INTEGER DEFAULT 0," &
            "embedding_json TEXT," &
            "vector_status TEXT NOT NULL DEFAULT 'pending'," &
            "last_error TEXT," &
            "updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
            "FOREIGN KEY (memory_id) REFERENCES memory_item(memory_id) ON DELETE CASCADE" &
            ");" &
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_memory_embedding_unique ON memory_embedding(memory_id, embedding_model);" &
            "CREATE INDEX IF NOT EXISTS idx_memory_embedding_status ON memory_embedding(vector_status, embedding_model);" &
            "CREATE TABLE IF NOT EXISTS memory_job (" &
            "job_id TEXT PRIMARY KEY," &
            "job_type TEXT NOT NULL," &
            "target_id TEXT," &
            "payload_json TEXT," &
            "status TEXT NOT NULL DEFAULT 'pending'," &
            "attempt_count INTEGER DEFAULT 0," &
            "last_error TEXT," &
            "created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
            "updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
            "next_run_at TEXT" &
            ");" &
            "CREATE INDEX IF NOT EXISTS idx_memory_job_pending ON memory_job(status, next_run_at, created_at);" &
            "CREATE INDEX IF NOT EXISTS idx_memory_job_type ON memory_job(job_type, status);" &
            "CREATE TABLE IF NOT EXISTS skills_registry (" &
            "skill_name TEXT PRIMARY KEY," &
            "file_path TEXT NOT NULL," &
            "app_scope TEXT DEFAULT ''," &
            "intent_types TEXT DEFAULT ''," &
            "trigger_keywords TEXT DEFAULT ''," &
            "description TEXT," &
            "embedding_json TEXT," &
            "usage_count INTEGER DEFAULT 0," &
            "success_count INTEGER DEFAULT 0," &
            "last_indexed_at TEXT," &
            "enabled INTEGER NOT NULL DEFAULT 1" &
            ");" &
            "CREATE INDEX IF NOT EXISTS idx_skills_registry_enabled ON skills_registry(enabled, app_scope);" &
            "CREATE INDEX IF NOT EXISTS idx_skills_registry_usage ON skills_registry(usage_count, success_count);", conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Sub EnsureDataMigrationMarkers(conn As SQLiteConnection)
        Using cmd As New SQLiteCommand(
            "CREATE TABLE IF NOT EXISTS data_migration_marker (" &
            "migration_key TEXT PRIMARY KEY," &
            "applied_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))" &
            ");", conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Sub MigrateLegacySkillUsage(conn As SQLiteConnection)
        MigrateLegacySkillUsageTable(conn)
        MigrateLegacySkillUsageJson(conn)
    End Sub

    Private Shared Sub MigrateLegacySkillUsageTable(conn As SQLiteConnection)
        Const markerKey As String = "skills_usage_table_to_registry_v1"
        If HasMigrationMarker(conn, markerKey) Then Return
        If Not TableExists(conn, "skills_usage") Then Return

        Using tx = conn.BeginTransaction()
            Try
                Using cmd As New SQLiteCommand(
                    "INSERT INTO skills_registry " &
                    "(skill_name, file_path, app_scope, intent_types, trigger_keywords, description, embedding_json, usage_count, success_count, last_indexed_at, enabled) " &
                    "SELECT skill_name, '', '', '', '', NULL, NULL, " &
                    "COALESCE(usage_count, 0), COALESCE(success_count, 0), COALESCE(last_used_at, updated_at, datetime('now', 'localtime')), 1 " &
                    "FROM skills_usage WHERE skill_name IS NOT NULL AND trim(skill_name) != '' " &
                    "ON CONFLICT(skill_name) DO UPDATE SET " &
                    "usage_count = max(COALESCE(skills_registry.usage_count, 0), COALESCE(excluded.usage_count, 0)), " &
                    "success_count = max(COALESCE(skills_registry.success_count, 0), COALESCE(excluded.success_count, 0)), " &
                    "last_indexed_at = CASE " &
                    "WHEN skills_registry.last_indexed_at IS NULL OR skills_registry.last_indexed_at = '' THEN excluded.last_indexed_at " &
                    "WHEN excluded.last_indexed_at IS NULL OR excluded.last_indexed_at = '' THEN skills_registry.last_indexed_at " &
                    "WHEN excluded.last_indexed_at > skills_registry.last_indexed_at THEN excluded.last_indexed_at " &
                    "ELSE skills_registry.last_indexed_at END", conn, tx)
                    cmd.ExecuteNonQuery()
                End Using

                InsertMigrationMarker(conn, tx, markerKey)
                tx.Commit()
                Debug.WriteLine("OfficeAiDatabase 已迁移旧 skills_usage 统计到 skills_registry")
            Catch ex As Exception
                Try
                    tx.Rollback()
                Catch rollbackEx As Exception
                    Debug.WriteLine($"skills_usage 迁移回滚失败: {rollbackEx.Message}")
                End Try
                Debug.WriteLine($"skills_usage 迁移失败: {ex.Message}")
                Throw
            End Try
        End Using
    End Sub

    Private Shared Sub MigrateLegacySkillUsageJson(conn As SQLiteConnection)
        Const markerKey As String = "skills_usage_json_to_registry_v1"
        If HasMigrationMarker(conn, markerKey) Then Return

        Dim jsonPath = GetLegacySkillsUsageJsonPath()
        If String.IsNullOrWhiteSpace(jsonPath) OrElse Not File.Exists(jsonPath) Then Return

        Dim storage As SkillsUsageStorage = Nothing
        Try
            Dim json = File.ReadAllText(jsonPath)
            storage = JsonConvert.DeserializeObject(Of SkillsUsageStorage)(json)
        Catch ex As Exception
            Debug.WriteLine($"skills_usage.json 读取失败，跳过迁移: {ex.Message}")
            Return
        End Try

        If storage Is Nothing OrElse storage.Skills Is Nothing OrElse storage.Skills.Count = 0 Then
            MarkMigrationApplied(conn, markerKey)
            Return
        End If

        Using tx = conn.BeginTransaction()
            Try
                For Each kvp In storage.Skills
                    Dim stats = kvp.Value
                    If stats Is Nothing Then Continue For

                    Dim skillName = If(stats.SkillName, kvp.Key)
                    If String.IsNullOrWhiteSpace(skillName) Then Continue For

                    Using cmd As New SQLiteCommand(
                        "INSERT INTO skills_registry " &
                        "(skill_name, file_path, app_scope, intent_types, trigger_keywords, description, embedding_json, usage_count, success_count, last_indexed_at, enabled) " &
                        "VALUES (@skill_name, '', '', '', '', NULL, NULL, @usage_count, @success_count, @last_used_at, 1) " &
                        "ON CONFLICT(skill_name) DO UPDATE SET " &
                        "usage_count = max(COALESCE(skills_registry.usage_count, 0), COALESCE(excluded.usage_count, 0)), " &
                        "success_count = max(COALESCE(skills_registry.success_count, 0), COALESCE(excluded.success_count, 0)), " &
                        "last_indexed_at = CASE " &
                        "WHEN skills_registry.last_indexed_at IS NULL OR skills_registry.last_indexed_at = '' THEN excluded.last_indexed_at " &
                        "WHEN excluded.last_indexed_at IS NULL OR excluded.last_indexed_at = '' THEN skills_registry.last_indexed_at " &
                        "WHEN excluded.last_indexed_at > skills_registry.last_indexed_at THEN excluded.last_indexed_at " &
                        "ELSE skills_registry.last_indexed_at END", conn, tx)
                        cmd.Parameters.AddWithValue("@skill_name", skillName.Trim())
                        cmd.Parameters.AddWithValue("@usage_count", Math.Max(0, stats.UsageCount))
                        cmd.Parameters.AddWithValue("@success_count", Math.Max(0, stats.SuccessCount))
                        If stats.LastUsedAt.HasValue Then
                            cmd.Parameters.AddWithValue("@last_used_at", stats.LastUsedAt.Value.ToString("yyyy-MM-dd HH:mm:ss"))
                        Else
                            cmd.Parameters.AddWithValue("@last_used_at", DBNull.Value)
                        End If
                        cmd.ExecuteNonQuery()
                    End Using
                Next

                InsertMigrationMarker(conn, tx, markerKey)
                tx.Commit()
                Debug.WriteLine("OfficeAiDatabase 已迁移旧 skills_usage.json 统计到 skills_registry")
            Catch ex As Exception
                Try
                    tx.Rollback()
                Catch rollbackEx As Exception
                    Debug.WriteLine($"skills_usage.json 迁移回滚失败: {rollbackEx.Message}")
                End Try
                Debug.WriteLine($"skills_usage.json 迁移失败: {ex.Message}")
                Throw
            End Try
        End Using
    End Sub

    Private Shared Function GetLegacySkillsUsageJsonPath() As String
        Return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder,
            "skills_usage.json")
    End Function

    Private Shared Function HasMigrationMarker(conn As SQLiteConnection, migrationKey As String) As Boolean
        Using cmd As New SQLiteCommand("SELECT COUNT(1) FROM data_migration_marker WHERE migration_key = @key", conn)
            cmd.Parameters.AddWithValue("@key", migrationKey)
            Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
        End Using
    End Function

    Private Shared Sub MarkMigrationApplied(conn As SQLiteConnection, migrationKey As String)
        Using tx = conn.BeginTransaction()
            InsertMigrationMarker(conn, tx, migrationKey)
            tx.Commit()
        End Using
    End Sub

    Private Shared Sub InsertMigrationMarker(conn As SQLiteConnection, tx As SQLiteTransaction, migrationKey As String)
        Using cmd As New SQLiteCommand("INSERT OR IGNORE INTO data_migration_marker (migration_key) VALUES (@key)", conn, tx)
            cmd.Parameters.AddWithValue("@key", migrationKey)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Sub EnsureAtomicMemoryColumns(conn As SQLiteConnection)
        EnsureColumn(conn, "atomic_memory", "app_type", "app_type TEXT DEFAULT ''")
        EnsureColumn(conn, "atomic_memory", "embedding", "embedding TEXT DEFAULT NULL")
        EnsureColumn(conn, "atomic_memory", "memory_type", "memory_type TEXT DEFAULT 'short_term'")
        EnsureColumn(conn, "atomic_memory", "importance", "importance REAL DEFAULT 0.5")
        EnsureColumn(conn, "atomic_memory", "access_count", "access_count INTEGER DEFAULT 0")
        EnsureColumn(conn, "atomic_memory", "last_access", "last_access TEXT")
        EnsureColumn(conn, "atomic_memory", "source_type", "source_type TEXT DEFAULT 'general'")
        EnsureColumn(conn, "atomic_memory", "linked_memories", "linked_memories TEXT")
        EnsureColumn(conn, "atomic_memory", "embedding_model", "embedding_model TEXT DEFAULT ''")
        EnsureColumn(conn, "atomic_memory", "embedding_dim", "embedding_dim INTEGER DEFAULT 0")
        EnsureColumn(conn, "atomic_memory", "embedding_updated_at", "embedding_updated_at TEXT")
    End Sub

    Private Shared Sub EnsureUserProfileSchema(conn As SQLiteConnection)
        If TableHasColumn(conn, "user_profile", "key") AndAlso TableHasColumn(conn, "user_profile", "value") Then
            EnsureColumn(conn, "user_profile", "category", "category TEXT DEFAULT 'preference'")
            EnsureColumn(conn, "user_profile", "confidence", "confidence REAL DEFAULT 0.5")
            EnsureColumn(conn, "user_profile", "last_updated", "last_updated TEXT")
            EnsureColumn(conn, "user_profile", "observation_count", "observation_count INTEGER DEFAULT 1")
            Return
        End If

        Using tx = conn.BeginTransaction()
            Using cmd As New SQLiteCommand(
                "CREATE TABLE IF NOT EXISTS user_profile_new (" &
                "id INTEGER PRIMARY KEY AUTOINCREMENT," &
                "key TEXT NOT NULL UNIQUE," &
                "value TEXT NOT NULL," &
                "category TEXT DEFAULT 'preference'," &
                "confidence REAL DEFAULT 0.5," &
                "last_updated TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))," &
                "observation_count INTEGER DEFAULT 1" &
                ");", conn, tx)
                cmd.ExecuteNonQuery()
            End Using

            If TableHasColumn(conn, "user_profile", "content") Then
                Using cmd As New SQLiteCommand("INSERT OR IGNORE INTO user_profile_new (key, value, category) SELECT 'legacy_data', content, 'summary' FROM user_profile WHERE content IS NOT NULL AND content != '' LIMIT 1;", conn, tx)
                    cmd.ExecuteNonQuery()
                End Using
            End If

            Using cmd As New SQLiteCommand("DROP TABLE IF EXISTS user_profile; ALTER TABLE user_profile_new RENAME TO user_profile;", conn, tx)
                cmd.ExecuteNonQuery()
            End Using
            tx.Commit()
        End Using
    End Sub

    Private Shared Sub EnsureColumn(conn As SQLiteConnection, tableName As String, columnName As String, columnDefinition As String)
        If TableHasColumn(conn, tableName, columnName) Then Return
        Using cmd As New SQLiteCommand($"ALTER TABLE {tableName} ADD COLUMN {columnDefinition};", conn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Function TableHasColumn(conn As SQLiteConnection, tableName As String, columnName As String) As Boolean
        Using cmd As New SQLiteCommand($"PRAGMA table_info({tableName})", conn)
            Using rdr = cmd.ExecuteReader()
                While rdr.Read()
                    If String.Equals(rdr("name").ToString(), columnName, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                End While
            End Using
        End Using
        Return False
    End Function

    Private Shared Function TableExists(conn As SQLiteConnection, tableName As String) As Boolean
        Using cmd As New SQLiteCommand("SELECT COUNT(1) FROM sqlite_master WHERE type = 'table' AND name = @table_name", conn)
            cmd.Parameters.AddWithValue("@table_name", tableName)
            Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
        End Using
    End Function
End Class
