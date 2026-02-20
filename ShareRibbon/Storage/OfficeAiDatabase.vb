' ShareRibbon\Storage\OfficeAiDatabase.vb
' Office AI 数据库初始化与迁移

Imports System.Data.SQLite
Imports System.IO
Imports System.Linq

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
        Return $"Data Source={dbPath};Version=3;"
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
             "UPDATE schema_version SET version = 5;"}
        }

        For Each kvp In migrations.OrderBy(Function(x) x.Key)
            If kvp.Key <= currentVersion Then Continue For
            Try
                Using cmd As New SQLiteCommand(kvp.Value, conn)
                    cmd.ExecuteNonQuery()
                End Using
                currentVersion = kvp.Key
                Debug.WriteLine($"OfficeAiDatabase 已应用迁移版本 {kvp.Key}")
            Catch ex As Exception
                Debug.WriteLine($"迁移版本 {kvp.Key} 失败: {ex.Message}")
                Throw
            End Try
        Next
    End Sub
End Class
