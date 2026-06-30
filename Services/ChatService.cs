using System.IO;
using AIDocAssistant.Models;
using Microsoft.Data.Sqlite;

namespace AIDocAssistant.Services;

public class ChatService
{
    private readonly string _db;

    public ChatService()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AIDocAssistant");
        Directory.CreateDirectory(dir);
        _db = Path.Combine(dir, "history.db");
        EnsureSchema();
    }

    private SqliteConnection Open()
    {
        var conn = new SqliteConnection($"Data Source={_db}");
        conn.Open();
        return conn;
    }

    private void EnsureSchema()
    {
        using var conn = Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            CREATE TABLE IF NOT EXISTS chat_sessions (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                title         TEXT    NOT NULL DEFAULT '',
                document_name TEXT    NOT NULL DEFAULT '',
                created_at    TEXT    NOT NULL,
                updated_at    TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS chat_messages (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id INTEGER NOT NULL,
                timestamp  TEXT    NOT NULL,
                role       TEXT    NOT NULL,
                content    TEXT    NOT NULL DEFAULT '',
                FOREIGN KEY(session_id) REFERENCES chat_sessions(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS chat_resources (
                session_id INTEGER NOT NULL,
                file_path  TEXT    NOT NULL,
                added_at   TEXT    NOT NULL,
                PRIMARY KEY(session_id, file_path),
                FOREIGN KEY(session_id) REFERENCES chat_sessions(id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_chat_messages_session_id
                ON chat_messages(session_id, id);
            CREATE INDEX IF NOT EXISTS idx_chat_sessions_updated_at
                ON chat_sessions(updated_at DESC);";
        cmd.ExecuteNonQuery();
    }

    public ChatSessionInfo CreateSession(string title, string documentName)
    {
        var now = DateTime.Now;
        using var conn = Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO chat_sessions (title, document_name, created_at, updated_at)
            VALUES (@title, @doc, @created, @updated);
            SELECT last_insert_rowid();";
        cmd.Parameters.AddWithValue("@title", title);
        cmd.Parameters.AddWithValue("@doc", documentName);
        cmd.Parameters.AddWithValue("@created", ToDbDate(now));
        cmd.Parameters.AddWithValue("@updated", ToDbDate(now));
        var id = (long)(cmd.ExecuteScalar() ?? 0L);

        return new ChatSessionInfo
        {
            Id = id,
            Title = title,
            DocumentName = documentName,
            CreatedAt = now,
            UpdatedAt = now
        };
    }

    public List<ChatSessionInfo> GetSessions(int limit = 100)
    {
        using var conn = Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT id, title, document_name, created_at, updated_at
            FROM chat_sessions
            ORDER BY updated_at DESC, id DESC
            LIMIT @limit";
        cmd.Parameters.AddWithValue("@limit", limit);
        using var reader = cmd.ExecuteReader();

        var result = new List<ChatSessionInfo>();
        while (reader.Read())
        {
            result.Add(new ChatSessionInfo
            {
                Id = reader.GetInt64(0),
                Title = reader.GetString(1),
                DocumentName = reader.GetString(2),
                CreatedAt = ParseDbDate(reader.GetString(3)),
                UpdatedAt = ParseDbDate(reader.GetString(4))
            });
        }
        return result;
    }

    public List<StoredChatMessage> GetMessages(long sessionId)
    {
        using var conn = Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT id, session_id, timestamp, role, content
            FROM chat_messages
            WHERE session_id = @sessionId
            ORDER BY id ASC";
        cmd.Parameters.AddWithValue("@sessionId", sessionId);
        using var reader = cmd.ExecuteReader();

        var result = new List<StoredChatMessage>();
        while (reader.Read())
        {
            result.Add(new StoredChatMessage
            {
                Id = reader.GetInt64(0),
                SessionId = reader.GetInt64(1),
                Timestamp = ParseDbDate(reader.GetString(2)),
                Role = reader.GetString(3),
                Content = reader.GetString(4)
            });
        }
        return result;
    }

    public void AddMessage(long sessionId, string role, string content)
    {
        var now = DateTime.Now;
        using var conn = Open();
        using var tx = conn.BeginTransaction();

        using (var cmd = conn.CreateCommand())
        {
            cmd.Transaction = tx;
            cmd.CommandText = @"
                INSERT INTO chat_messages (session_id, timestamp, role, content)
                VALUES (@sessionId, @timestamp, @role, @content)";
            cmd.Parameters.AddWithValue("@sessionId", sessionId);
            cmd.Parameters.AddWithValue("@timestamp", ToDbDate(now));
            cmd.Parameters.AddWithValue("@role", role);
            cmd.Parameters.AddWithValue("@content", content);
            cmd.ExecuteNonQuery();
        }

        using (var cmd = conn.CreateCommand())
        {
            cmd.Transaction = tx;
            cmd.CommandText = "UPDATE chat_sessions SET updated_at = @updated WHERE id = @sessionId";
            cmd.Parameters.AddWithValue("@updated", ToDbDate(now));
            cmd.Parameters.AddWithValue("@sessionId", sessionId);
            cmd.ExecuteNonQuery();
        }

        tx.Commit();
    }

    public List<string> GetResourcePaths(long sessionId)
    {
        using var conn = Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
            SELECT file_path
            FROM chat_resources
            WHERE session_id = @sessionId
            ORDER BY added_at ASC";
        cmd.Parameters.AddWithValue("@sessionId", sessionId);
        using var reader = cmd.ExecuteReader();

        var result = new List<string>();
        while (reader.Read())
            result.Add(reader.GetString(0));
        return result;
    }

    public void SetResourcePaths(long sessionId, IEnumerable<string> filePaths)
    {
        var now = ToDbDate(DateTime.Now);
        using var conn = Open();
        using var tx = conn.BeginTransaction();

        using (var cmd = conn.CreateCommand())
        {
            cmd.Transaction = tx;
            cmd.CommandText = "DELETE FROM chat_resources WHERE session_id = @sessionId";
            cmd.Parameters.AddWithValue("@sessionId", sessionId);
            cmd.ExecuteNonQuery();
        }

        foreach (var path in filePaths.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            using var cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = @"
                INSERT OR REPLACE INTO chat_resources (session_id, file_path, added_at)
                VALUES (@sessionId, @path, @added)";
            cmd.Parameters.AddWithValue("@sessionId", sessionId);
            cmd.Parameters.AddWithValue("@path", path);
            cmd.Parameters.AddWithValue("@added", now);
            cmd.ExecuteNonQuery();
        }

        using (var cmd = conn.CreateCommand())
        {
            cmd.Transaction = tx;
            cmd.CommandText = "UPDATE chat_sessions SET updated_at = @updated WHERE id = @sessionId";
            cmd.Parameters.AddWithValue("@updated", now);
            cmd.Parameters.AddWithValue("@sessionId", sessionId);
            cmd.ExecuteNonQuery();
        }

        tx.Commit();
    }

    private static string ToDbDate(DateTime value) => value.ToString("yyyy-MM-dd HH:mm:ss");

    private static DateTime ParseDbDate(string value) =>
        DateTime.TryParse(value, out var parsed) ? parsed : DateTime.Now;
}
