using System.Windows;
using System.Windows.Controls;
using AIDocAssistant.ViewModels;

namespace AIDocAssistant.Views;

public partial class SettingsWindow : Window
{
    public SettingsWindow() => InitializeComponent();

    private void ApiKeyBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (DataContext is SettingsViewModel vm)
            vm.ApiKey = ApiKeyBox.Password;
    }
}
