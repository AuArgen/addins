using System.Windows;
using System.Windows.Controls;
using AIDocAssistant.ViewModels;

namespace AIDocAssistant.Views;

public class ChatTemplateSelector : DataTemplateSelector
{
    public DataTemplate? UserTemplate   { get; set; }
    public DataTemplate? AiTemplate     { get; set; }
    public DataTemplate? SystemTemplate { get; set; }

    public override DataTemplate? SelectTemplate(object item, DependencyObject container) =>
        item is ChatMessage msg
            ? msg.Role switch
            {
                MessageRole.User      => UserTemplate,
                MessageRole.Assistant => AiTemplate,
                _                     => SystemTemplate
            }
            : base.SelectTemplate(item, container);
}
