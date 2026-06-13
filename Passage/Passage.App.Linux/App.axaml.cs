using System.Diagnostics;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Styling;
using Avalonia.Threading;
using Passage.App.Views;

namespace Passage.App;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.MainWindow = new MainWindow();

            // Detect system theme on Linux
            DetectAndApplySystemTheme();
        }

        base.OnFrameworkInitializationCompleted();
    }

    private void DetectAndApplySystemTheme()
    {
        // Try D-Bus first, then gsettings, then fallback to dark
        var theme = DetectLinuxTheme();
        RequestedThemeVariant = theme;

        // Load the appropriate theme resources
        LoadThemeResources(theme == ThemeVariant.Light);
    }

    private ThemeVariant DetectLinuxTheme()
    {
        // Try D-Bus first (most reliable for GNOME/COSMIC)
        if (TryDbusThemeDetection(out var dbusTheme))
            return dbusTheme;

        // Fallback to gsettings command
        if (TryGsettingsThemeDetection(out var gsettingsTheme))
            return gsettingsTheme;

        // Fallback to environment variable
        if (TryEnvironmentThemeDetection(out var envTheme))
            return envTheme;

        // Default to dark (matches original app)
        return ThemeVariant.Dark;
    }

    private bool TryDbusThemeDetection(out ThemeVariant theme)
    {
        theme = ThemeVariant.Dark;
        try
        {
            // Use gsettings via process - D-Bus requires additional libraries
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "gsettings",
                    Arguments = "get org.gnome.desktop.interface color-scheme",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };
            process.Start();
            var result = process.StandardOutput.ReadToEnd().Trim().TrimEnd('\''); // Trim trailing quote
            process.WaitForExit();

            // gsettings returns 'prefer-dark', 'prefer-light', or 'default'
            if (result.Contains("prefer-dark"))
            {
                theme = ThemeVariant.Dark;
                return true;
            }
            if (result.Contains("prefer-light"))
            {
                theme = ThemeVariant.Light;
                return true;
            }
        }
        catch
        {
            // Ignore errors and try next method
        }
        return false;
    }

    private bool TryGsettingsThemeDetection(out ThemeVariant theme)
    {
        theme = ThemeVariant.Dark;
        try
        {
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "gsettings",
                    Arguments = "get org.gnome.desktop.interface gtk-theme",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };
            process.Start();
            var result = process.StandardOutput.ReadToEnd().Trim().ToLower();
            process.WaitForExit();

            // Check if theme name contains 'dark'
            if (result.Contains("dark"))
            {
                theme = ThemeVariant.Dark;
                return true;
            }
            if (result.Contains("light"))
            {
                theme = ThemeVariant.Light;
                return true;
            }
        }
        catch
        {
            // Ignore errors and try next method
        }
        return false;
    }

    private bool TryEnvironmentThemeDetection(out ThemeVariant theme)
    {
        theme = ThemeVariant.Dark;
        // Check common environment variables
        var gtkTheme = Environment.GetEnvironmentVariable("GTK_THEME");
        if (!string.IsNullOrEmpty(gtkTheme))
        {
            if (gtkTheme.Contains("dark", StringComparison.OrdinalIgnoreCase))
            {
                theme = ThemeVariant.Dark;
                return true;
            }
            if (gtkTheme.Contains("light", StringComparison.OrdinalIgnoreCase))
            {
                theme = ThemeVariant.Light;
                return true;
            }
        }

        // Check XDG current theme
        var xdgTheme = Environment.GetEnvironmentVariable("XDG_CURRENT_DESKTOP");
        if (!string.IsNullOrEmpty(xdgTheme))
        {
            // COSMIC/Unity typically uses dark mode by default
            if (xdgTheme.Contains("COSMIC", StringComparison.OrdinalIgnoreCase) ||
                xdgTheme.Contains("unity", StringComparison.OrdinalIgnoreCase))
            {
                theme = ThemeVariant.Dark;
                return true;
            }
        }

        return false;
    }

    private void LoadThemeResources(bool isLight)
    {
        if (Resources is null) return;

        if (isLight)
        {
            // Light theme colors
            Resources["ThemeBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#F1F1F1"));
            Resources["WindowBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#F6F6F6"));
            Resources["SurfaceBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#F6F6F6"));
            Resources["SurfaceMutedBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#EAEAEA"));
            Resources["SurfaceRaisedBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#FFFFFF"));
            Resources["SurfaceBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#D3D3D3"));
            Resources["ControlBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#EEEEEE"));
            Resources["ControlForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#2E2E2E"));
            Resources["ControlBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#D3D3D3"));
            Resources["ControlAccent"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#3584E4"));
            Resources["ControlPressedBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#DCDCDB"));
            Resources["ControlPressedForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#2E2E2E"));
            Resources["HeaderText"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#1E1E1E"));
            Resources["SecondaryText"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#5E5E5E"));
            Resources["MutedText"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#8E8E8E"));
            Resources["EditorForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#2E2E2E"));
            Resources["EditorPageBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#D3D3D3"));
            Resources["WindowForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#2E2E2E"));
            Resources["CardBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#FFFFFF"));
            Resources["CardBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#E0E0E0"));
            Resources["CardAccent"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#3584E4"));
        }
        else
        {
            // Dark theme colors (original)
            Resources["ThemeBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#1E1E1E"));
            Resources["WindowBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#242424"));
            Resources["SurfaceBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#242424"));
            Resources["SurfaceMutedBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#1E1E1E"));
            Resources["SurfaceRaisedBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#2D2D2D"));
            Resources["SurfaceBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#353535"));
            Resources["ControlBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#303030"));
            Resources["ControlForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#DEDEDE"));
            Resources["ControlBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#353535"));
            Resources["ControlAccent"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#3584E4"));
            Resources["ControlPressedBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#3A3A3A"));
            Resources["ControlPressedForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#FFFFFF"));
            Resources["HeaderText"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#FFFFFF"));
            Resources["SecondaryText"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#B0B0B0"));
            Resources["MutedText"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#787878"));
            Resources["EditorForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#DEDEDE"));
            Resources["EditorPageBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#353535"));
            Resources["WindowForeground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#DEDEDE"));
            Resources["CardBackground"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#303030"));
            Resources["CardBorder"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#3D3D3D"));
            Resources["CardAccent"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#3584E4"));
        }
    }
}
