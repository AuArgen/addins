using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using AIDocAssistant.ViewModels;

namespace AIDocAssistant.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        var vm = new MainViewModel();
        DataContext = vm;

        // Авто-скролл вниз при новых сообщениях
        vm.ChatMessages.CollectionChanged += OnChatChanged;
    }

    private void OnChatChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.Action == NotifyCollectionChangedAction.Add)
            Dispatcher.BeginInvoke(() => ChatScroll.ScrollToBottom(), System.Windows.Threading.DispatcherPriority.Background);
    }

    private void InputBox_KeyDown(object sender, KeyEventArgs e)
    {
        // Ctrl+Enter или просто Enter — отправить
        if (e.Key == Key.Enter && !Keyboard.IsKeyDown(Key.LeftShift))
        {
            if (DataContext is MainViewModel vm)
                vm.SendMessageCommand.Execute(null);
            e.Handled = true;
        }
    }

    private void DocList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (DataContext is MainViewModel vm)
            vm.SelectDocumentCommand.Execute(null);
    }
}
