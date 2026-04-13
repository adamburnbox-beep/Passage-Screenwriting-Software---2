using System.Security;
using Microsoft.Win32;
using System.Windows;
using System.Windows.Threading;

namespace Passage.App.Services;

public static class ThemeManager
{
    public const string LightThemeName = "Light";
    public const string DarkThemeName = "Dark";
    public const string EReaderThemeName = "E-Reader";
    public const string EReaderDarkThemeName = "E-Reader Dark";
    public const string SystemThemeName = "System Default";

    private const string ThemeFolder = "Themes";
    private const string LightThemeFile = "LightTheme.xaml";
    private const string DarkThemeFile = "DarkTheme.xaml";
    private const string EReaderThemeFile = "EReaderTheme.xaml";
    private const string EReaderDarkThemeFile = "EReaderDarkTheme.xaml";
    private const string WindowsThemeKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
    private const string AppsUseLightThemeValueName = "AppsUseLightTheme";

    private static readonly DispatcherTimer SystemThemeTimer = new()
    {
        Interval = TimeSpan.FromSeconds(1)
    };

    private static string _requestedThemeName = DarkThemeName;
    private static string _appliedThemeName = DarkThemeName;
    private static string _lastObservedSystemTheme = DarkThemeName;
    private static bool _followSystemTheme;

    static ThemeManager()
    {
        SystemThemeTimer.Tick += SystemThemeTimer_Tick;
    }

    public static event EventHandler? ThemeChanged;

    public static string CurrentThemeName => _requestedThemeName;

    public static string AppliedThemeName => _appliedThemeName;

    public static void SetTheme(string themeName)
    {
        var normalizedThemeName = NormalizeThemeName(themeName);
        var resolvedThemeName = normalizedThemeName == SystemThemeName
            ? GetWindowsSystemTheme()
            : normalizedThemeName;

        RunOnDispatcher(() =>
        {
            _requestedThemeName = normalizedThemeName;
            _appliedThemeName = resolvedThemeName;
            _followSystemTheme = normalizedThemeName == SystemThemeName;
            ApplyThemeDictionary(resolvedThemeName);
            UpdateSystemThemeWatcher();
        });
    }

    public static string GetWindowsSystemTheme()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(WindowsThemeKeyPath, writable: false);
            var value = key?.GetValue(AppsUseLightThemeValueName);
            return value is int appsUseLightTheme && appsUseLightTheme == 0
                ? DarkThemeName
                : LightThemeName;
        }
        catch (SecurityException)
        {
            return LightThemeName;
        }
        catch (UnauthorizedAccessException)
        {
            return LightThemeName;
        }
    }

    public static void Shutdown()
    {
        RunOnDispatcher(() =>
        {
            _followSystemTheme = false;
            SystemThemeTimer.Stop();
        });
    }

    private static void ApplyThemeDictionary(string themeName)
    {
        var app = Application.Current;
        if (app is null)
        {
            return;
        }

        var themeDictionary = new ResourceDictionary
        {
            Source = new Uri(
                $"pack://application:,,,/Passage.App;component/{ThemeFolder}/{GetThemeFileName(themeName)}",
                UriKind.Absolute)
        };

        app.Resources.MergedDictionaries.Clear();
        app.Resources.MergedDictionaries.Add(themeDictionary);
        ThemeChanged?.Invoke(null, EventArgs.Empty);
    }

    private static string GetThemeFileName(string themeName)
    {
        return NormalizeThemeName(themeName) switch
        {
            DarkThemeName => DarkThemeFile,
            EReaderThemeName => EReaderThemeFile,
            EReaderDarkThemeName => EReaderDarkThemeFile,
            _ => LightThemeFile
        };
    }

    private static string NormalizeThemeName(string themeName)
    {
        if (string.IsNullOrWhiteSpace(themeName))
        {
            return DarkThemeName;
        }

        if (string.Equals(themeName, DarkThemeName, StringComparison.OrdinalIgnoreCase))
        {
            return DarkThemeName;
        }

        if (string.Equals(themeName, EReaderThemeName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(themeName, "EReader", StringComparison.OrdinalIgnoreCase))
        {
            return EReaderThemeName;
        }

        if (string.Equals(themeName, EReaderDarkThemeName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(themeName, "EReaderDark", StringComparison.OrdinalIgnoreCase))
        {
            return EReaderDarkThemeName;
        }

        if (string.Equals(themeName, SystemThemeName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(themeName, "System", StringComparison.OrdinalIgnoreCase))
        {
            return SystemThemeName;
        }

        return LightThemeName;
    }

    private static void UpdateSystemThemeWatcher()
    {
        if (_followSystemTheme)
        {
            _lastObservedSystemTheme = GetWindowsSystemTheme();
            if (!SystemThemeTimer.IsEnabled)
            {
                SystemThemeTimer.Start();
            }

            return;
        }

        SystemThemeTimer.Stop();
    }

    private static void SystemThemeTimer_Tick(object? sender, EventArgs e)
    {
        if (!_followSystemTheme)
        {
            SystemThemeTimer.Stop();
            return;
        }

        var currentSystemTheme = GetWindowsSystemTheme();
        if (string.Equals(currentSystemTheme, _lastObservedSystemTheme, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        _lastObservedSystemTheme = currentSystemTheme;
        _appliedThemeName = currentSystemTheme;
        ApplyThemeDictionary(currentSystemTheme);
    }

    private static void RunOnDispatcher(Action action)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is null || dispatcher.CheckAccess())
        {
            action();
            return;
        }

        dispatcher.Invoke(action);
    }
}
