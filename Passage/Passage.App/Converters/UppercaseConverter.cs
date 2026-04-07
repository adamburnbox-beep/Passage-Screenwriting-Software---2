using System.Globalization;
using System.Windows.Data;

namespace Passage.App.Converters;

public sealed class UppercaseConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        return value is string text
            ? text.ToUpper(culture)
            : string.Empty;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
