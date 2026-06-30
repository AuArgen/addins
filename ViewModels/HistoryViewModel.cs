using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using AIDocAssistant.Models;
using AIDocAssistant.Services;

namespace AIDocAssistant.ViewModels;

public partial class HistoryViewModel : ObservableObject
{
    private readonly HistoryService _svc;
    private readonly List<HistoryEntry> _all = [];

    public HistoryViewModel(HistoryService svc)
    {
        _svc = svc;
        Reload();
    }

    [ObservableProperty] private ObservableCollection<HistoryEntry> _entries = [];
    [ObservableProperty] private HistoryEntry? _selected;
    [ObservableProperty] private string _filterText = string.Empty;
    [ObservableProperty] private string _statusText = string.Empty;

    partial void OnFilterTextChanged(string value) => ApplyFilter();

    partial void OnSelectedChanged(HistoryEntry? value)
    {
        // Форматируем детали для просмотра
        if (value is null) { StatusText = ""; return; }

        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"[{value.TimestampStr}] {value.DocumentName}");
        sb.AppendLine();
        sb.AppendLine($"Команда: {value.Command}");
        sb.AppendLine();
        sb.AppendLine("Фаза 1 (план):");
        sb.AppendLine(string.IsNullOrEmpty(value.Phase1Plan) ? "  (нет данных)" : "  " + value.Phase1Plan);
        sb.AppendLine();
        sb.AppendLine($"Операции ({value.OperationsCount}, применено {value.AppliedCount}):");

        if (!string.IsNullOrEmpty(value.OperationsJson))
        {
            try
            {
                var ops = System.Text.Json.JsonSerializer.Deserialize<List<ChangeOperation>>(value.OperationsJson,
                    new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                if (ops is not null)
                    foreach (var op in ops)
                        sb.AppendLine($"  • {op.DisplayText}");
            }
            catch
            {
                sb.AppendLine("  " + value.OperationsJson);
            }
        }
        else
        {
            sb.AppendLine("  (нет данных)");
        }

        StatusText = sb.ToString();
    }

    [RelayCommand]
    private void Reload()
    {
        _all.Clear();
        _all.AddRange(_svc.GetAll());
        ApplyFilter();
    }

    [RelayCommand]
    private void ClearHistory()
    {
        _svc.Clear();
        _all.Clear();
        Entries.Clear();
        StatusText = "История очищена.";
    }

    private void ApplyFilter()
    {
        var q = FilterText.Trim().ToLower();
        var filtered = string.IsNullOrEmpty(q)
            ? _all
            : _all.Where(e =>
                e.Command.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                e.DocumentName.Contains(q, StringComparison.OrdinalIgnoreCase)).ToList();

        Entries = new ObservableCollection<HistoryEntry>(filtered);
    }
}
