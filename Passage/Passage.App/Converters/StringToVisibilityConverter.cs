using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Passage.App.Converters;

public class StringToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        bool hasValue = !string.IsNullOrWhiteSpace(value as string);
        if (parameter as string == "Inverse")
        {
            return hasValue ? Visibility.Collapsed : Visibility.Visible;
        }
        return hasValue ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
