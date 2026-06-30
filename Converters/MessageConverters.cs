using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using AIDocAssistant.ViewModels;

namespace AIDocAssistant.Converters;

public class MessageRoleToColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        value is MessageRole role ? role switch
        {
            MessageRole.User      => new SolidColorBrush(Color.FromRgb(124, 58, 237)),  // фиолетовый
            MessageRole.Assistant => new SolidColorBrush(Color.FromRgb(55, 65, 81)),    // тёмно-серый
            MessageRole.System    => new SolidColorBrush(Color.FromRgb(31, 41, 55)),    // очень тёмный
            _                     => Brushes.Gray
        } : Brushes.Gray;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
        throw new NotImplementedException();
}

public class MessageRoleToAlignConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        value is MessageRole role ? role switch
        {
            MessageRole.User  => HorizontalAlignment.Right,
            _                 => HorizontalAlignment.Left
        } : HorizontalAlignment.Left;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
        throw new NotImplementedException();
}
