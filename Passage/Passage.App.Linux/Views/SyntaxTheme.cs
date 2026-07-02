using System.Collections.Generic;
using Avalonia;
using Avalonia.Media;

namespace Passage.App.Views;

/// <summary>
/// Resolves editor syntax brushes from the application's theme resources so the
/// Fountain and Markdown colorizers follow light/dark theme switches. Falls back
/// to a given hex value (the dark-theme colour) when a key isn't defined.
/// </summary>
internal static class SyntaxTheme
{
    private static readonly Dictionary<string, IBrush> FallbackCache = new();

    public static IBrush Brush(string resourceKey, string fallbackHex)
    {
        var app = Application.Current;
        if (app is not null &&
            app.Resources.TryGetResource(resourceKey, null, out var value) &&
            value is IBrush themed)
        {
            return themed;
        }

        if (!FallbackCache.TryGetValue(fallbackHex, out var fallback))
        {
            fallback = new SolidColorBrush(Color.Parse(fallbackHex));
            FallbackCache[fallbackHex] = fallback;
        }

        return fallback;
    }
}
