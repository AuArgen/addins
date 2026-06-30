using System.IO;
using System.Text.Json;
using AIDocAssistant.Models;
using Microsoft.Data.Sqlite;

namespace AIDocAssistant.Services;

public class HistoryService
{
    private readonly string _db;

    public HistoryService()
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
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = @"
            CREATE TABLE IF NOT EXISTS history (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp        TEXT    NOT NULL,
                document_name    TEXT    NOT NULL DEFAULT '',
                command          TEXT    NOT NULL DEFAULT '',
                phase1_plan      TEXT    NOT NULL DEFAULT '',
                operations_count INTEGER NOT NULL DEFAULT 0,
                applied_count    INTEGER NOT NULL DEFAULT 0,
                was_direct       INTEGER NOT NULL DEFAULT 0,
                operations_json  TEXT    NOT NULL DEFAULT ''
            )";
        cmd.ExecuteNonQuery();
    }

    public void Add(HistoryEntry e)
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO history
                (timestamp, document_name, command, phase1_plan,
                 operations_count, applied_count, was_direct, operations_json)
            VALUES
                (@ts, @doc, @cmd, @plan, @opCnt, @appCnt, @direct, @ops)";
        cmd.Parameters.AddWithValue("@ts",     e.Timestamp.ToString("yyyy-MM-dd HH:mm:ss"));
        cmd.Parameters.AddWithValue("@doc",    e.DocumentName);
        cmd.Parameters.AddWithValue("@cmd",    e.Command);
        cmd.Parameters.AddWithValue("@plan",   e.Phase1Plan);
        cmd.Parameters.AddWithValue("@opCnt",  e.OperationsCount);
        cmd.Parameters.AddWithValue("@appCnt", e.AppliedCount);
        cmd.Parameters.AddWithValue("@direct", e.WasDirect ? 1 : 0);
        cmd.Parameters.AddWithValue("@ops",    e.OperationsJson);
        cmd.ExecuteNonQuery();
    }

    public long AddAndGetId(HistoryEntry e)
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = @"
            INSERT INTO history
                (timestamp, document_name, command, phase1_plan,
                 operations_count, applied_count, was_direct, operations_json)
            VALUES
                (@ts, @doc, @cmd, @plan, @opCnt, @appCnt, @direct, @ops);
            SELECT last_insert_rowid();";
        cmd.Parameters.AddWithValue("@ts",     e.Timestamp.ToString("yyyy-MM-dd HH:mm:ss"));
        cmd.Parameters.AddWithValue("@doc",    e.DocumentName);
        cmd.Parameters.AddWithValue("@cmd",    e.Command);
        cmd.Parameters.AddWithValue("@plan",   e.Phase1Plan);
        cmd.Parameters.AddWithValue("@opCnt",  e.OperationsCount);
        cmd.Parameters.AddWithValue("@appCnt", e.AppliedCount);
        cmd.Parameters.AddWithValue("@direct", e.WasDirect ? 1 : 0);
        cmd.Parameters.AddWithValue("@ops",    e.OperationsJson);
        return (long)(cmd.ExecuteScalar() ?? 0L);
    }

    public void UpdateApplied(long id, int appliedCount)
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = "UPDATE history SET applied_count=@cnt WHERE id=@id";
        cmd.Parameters.AddWithValue("@cnt", appliedCount);
        cmd.Parameters.AddWithValue("@id",  id);
        cmd.ExecuteNonQuery();
    }

    public List<HistoryEntry> GetAll(int limit = 500)
    {
        using var conn   = Open();
        using var cmd    = conn.CreateCommand();
        cmd.CommandText  = $"SELECT * FROM history ORDER BY id DESC LIMIT {limit}";
        using var reader = cmd.ExecuteReader();

        var result = new List<HistoryEntry>();
        while (reader.Read())
        {
            result.Add(new HistoryEntry
            {
                Id             = reader.GetInt32(0),
                Timestamp      = DateTime.Parse(reader.GetString(1)),
                DocumentName   = reader.GetString(2),
                Command        = reader.GetString(3),
                Phase1Plan     = reader.GetString(4),
                OperationsCount= reader.GetInt32(5),
                AppliedCount   = reader.GetInt32(6),
                WasDirect      = reader.GetInt32(7) == 1,
                OperationsJson = reader.GetString(8)
            });
        }
        return result;
    }

    public void Clear()
    {
        using var conn = Open();
        using var cmd  = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM history";
        cmd.ExecuteNonQuery();
    }

    public string DbPath => _db;
}
