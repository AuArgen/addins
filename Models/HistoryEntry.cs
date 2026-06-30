namespace AIDocAssistant.Models;

public class HistoryEntry
{
    public int      Id             { get; set; }
    public DateTime Timestamp      { get; set; } = DateTime.Now;
    public string   DocumentName   { get; set; } = "";
    public string   Command        { get; set; } = "";
    public string   Phase1Plan     { get; set; } = ""; // DataQuery JSON
    public int      OperationsCount{ get; set; }
    public int      AppliedCount   { get; set; }
    public bool     WasDirect      { get; set; } // Phase 2 пропущена
    public string   OperationsJson { get; set; } = "";

    // Для отображения в таблице
    public string TimestampStr    => Timestamp.ToString("dd.MM.yyyy HH:mm:ss");
    public string StatusSummary   => WasDirect
        ? $"⚡ Прямое · {OperationsCount} оп."
        : $"Фаза 2 · {OperationsCount} оп. · применено {AppliedCount}";
}
