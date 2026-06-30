using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using AIDocAssistant.Models;

namespace AIDocAssistant.ViewModels;

public partial class PreviewViewModel : ObservableObject
{
    public event Action<IEnumerable<ChangeOperation>>? Applied;

    [ObservableProperty]
    private ObservableCollection<ChangeOperationItem> _items = [];

    [ObservableProperty]
    private int _selectedCount;

    public PreviewViewModel(List<ChangeOperation> operations)
    {
        foreach (var op in operations)
        {
            var item = new ChangeOperationItem(op);
            item.PropertyChanged += (_, _) => UpdateCount();
            Items.Add(item);
        }
        UpdateCount();
    }

    [RelayCommand]
    private void ApplySelected()
    {
        var selected = Items.Where(i => i.IsSelected).Select(i => i.Operation).ToList();
        Applied?.Invoke(selected);
    }

    [RelayCommand]
    private void CancelAll()
    {
        foreach (var item in Items) item.IsSelected = false;
    }

    [RelayCommand]
    private void SelectAll()
    {
        foreach (var item in Items) item.IsSelected = true;
    }

    private void UpdateCount() =>
        SelectedCount = Items.Count(i => i.IsSelected);
}

public partial class ChangeOperationItem : ObservableObject
{
    [ObservableProperty]
    private bool _isSelected = true;

    public ChangeOperation Operation { get; }
    public string DisplayText => Operation.DisplayText;
    public string ScriptText  => Operation.ScriptText;
    public string GroupKey    => Operation.GroupKey;

    public ChangeOperationItem(ChangeOperation op) => Operation = op;
}
