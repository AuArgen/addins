using System.Windows;
using System.Windows.Controls;
using AIDocAssistant.ViewModels;

namespace AIDocAssistant.Views;

public partial class HistoryWindow : Window
{
    public HistoryWindow()
    {
        InitializeComponent();
    }

    private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Handled by binding SelectedItem → HistoryViewModel.Selected
    }
}
