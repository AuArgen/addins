using System.IO;

namespace AIDocAssistant.Services;

public class BackupService
{
    private readonly string _logDir;

    public BackupService()
    {
        _logDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AIDocAssistant", "logs");
        Directory.CreateDirectory(_logDir);
    }

    public string CreateBackup(dynamic doc)
    {
        var dir   = Path.GetDirectoryName((string)doc.FullName) ?? string.Empty;
        var name  = Path.GetFileNameWithoutExtension((string)doc.Name);
        var stamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm");
        var path  = Path.Combine(dir, $"{name}_backup_{stamp}.docx");

        // SaveCopyAs — сохраняет КОПИЮ, оригинал остаётся открытым и активным
        // SaveAs2 НЕЛЬЗЯ — оно перемещает документ на новый путь (оригинал "исчезает")
        doc.SaveCopyAs(path);
        LogAction($"Backup создан: {path}");
        return path;
    }

    public void LogOperation(string operationType, int paragraphIndex, bool success, string? error = null)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {operationType} p{paragraphIndex} " +
                   (success ? "OK" : $"FAIL: {error}");
        LogAction(line);
    }

    public void LogAction(string message)
    {
        var file = Path.Combine(_logDir, $"log_{DateTime.Now:yyyy-MM-dd}.txt");
        File.AppendAllText(file, message + Environment.NewLine, System.Text.Encoding.UTF8);
    }

    // Проверяет — не является ли документ уже backup-файлом
    public static bool IsBackupFile(string docName) =>
        docName.Contains("_backup_");

    public string LogDirectory => _logDir;
}
