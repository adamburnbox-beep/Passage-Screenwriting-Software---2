using System.Globalization;
using System.Text;
using System.Windows.Data;

namespace Passage.App.Converters;

public sealed class GoalTypeLabelConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var text = value?.ToString() ?? string.Empty;
        if (text.Length == 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder(text.Length + 4);
        builder.Append(text[0]);

        for (var i = 1; i < text.Length; i++)
        {
            var current = text[i];
            var previous = text[i - 1];

            if (char.IsUpper(current) && previous != ' ')
            {
                builder.Append(' ');
            }

            builder.Append(current);
        }

        return builder.ToString();
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
