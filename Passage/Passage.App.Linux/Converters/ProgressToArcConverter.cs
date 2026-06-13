using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;

namespace Passage.App.Converters;

public sealed class ProgressToArcConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        double percent = 0;
        if (value is double d) percent = d;
        else if (value is float f) percent = f;
        else if (value is int i) percent = i;
        else if (value is decimal dec) percent = (double)dec;

        percent = Math.Clamp(percent, 0, 100);
        if (percent >= 100)
        {
            return Geometry.Parse("M 55,5 A 50,50 0 1 1 54.99,5 Z");
        }
        if (percent <= 0)
        {
            return Geometry.Parse("M 55,5");
        }

        double angle = (percent / 100.0) * 360.0;
        double rad = (angle - 90.0) * Math.PI / 180.0;
        double endX = 55.0 + 50.0 * Math.Cos(rad);
        double endY = 55.0 + 50.0 * Math.Sin(rad);

        int isLargeArc = percent > 50 ? 1 : 0;
        string pathData = string.Format(CultureInfo.InvariantCulture, "M 55,5 A 50,50 0 {0} 1 {1:F2},{2:F2}", isLargeArc, endX, endY);
        return Geometry.Parse(pathData);
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
