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

    private static Avalonia.Media.SolidColorBrush Brush(string hex) =>
        new(Avalonia.Media.Color.Parse(hex));

    private void LoadThemeResources(bool isLight)
    {
        if (Resources is null) return;

        if (isLight)
        {
            // Light theme colors (e-reader cream/paper)
            Resources["ThemeBackground"] = Brush("#EFE7D6");
            Resources["WindowBackground"] = Brush("#EFE7D6");
            Resources["SurfaceBackground"] = Brush("#F2ECDD");
            Resources["SurfaceMutedBackground"] = Brush("#E6DDC8");
            Resources["SurfaceRaisedBackground"] = Brush("#FBF6EA");
            Resources["SurfaceBorder"] = Brush("#D8CDB5");
            Resources["ControlBackground"] = Brush("#ECE3D0");
            Resources["ControlForeground"] = Brush("#2B2620");
            Resources["ControlBorder"] = Brush("#CFC1A8");
            Resources["ControlAccent"] = Brush("#2B2620");
            Resources["ControlPressedBackground"] = Brush("#2B2620");
            Resources["ControlPressedForeground"] = Brush("#FBF6EA");
            Resources["HeaderText"] = Brush("#1E1A14");
            Resources["SecondaryText"] = Brush("#6B6353");
            Resources["MutedText"] = Brush("#938A76");
            Resources["EditorForeground"] = Brush("#2B2620");
            Resources["EditorPageBorder"] = Brush("#D8CDB5");
            Resources["WindowForeground"] = Brush("#2B2620");
            Resources["HierarchyIndicatorBrush"] = Brush("#2B2620");
            Resources["DragOverBackground"] = Brush("#252B2620");
            Resources["CardBackground"] = Brush("#FBF6EA");
            Resources["CardBorder"] = Brush("#D8CDB5");
            Resources["CardAccent"] = Brush("#2B2620");
        }
        else
        {
            // Dark theme colors (e-reader near-black; inverse of light)
            Resources["ThemeBackground"] = Brush("#15140F");
            Resources["WindowBackground"] = Brush("#15140F");
            Resources["SurfaceBackground"] = Brush("#1B1A15");
            Resources["SurfaceMutedBackground"] = Brush("#232118");
            Resources["SurfaceRaisedBackground"] = Brush("#100F0B");
            Resources["SurfaceBorder"] = Brush("#2E2B22");
            Resources["ControlBackground"] = Brush("#232118");
            Resources["ControlForeground"] = Brush("#ECE3D0");
            Resources["ControlBorder"] = Brush("#3A352A");
            Resources["ControlAccent"] = Brush("#F2ECDD");
            Resources["ControlPressedBackground"] = Brush("#F2ECDD");
            Resources["ControlPressedForeground"] = Brush("#15140F");
            Resources["HeaderText"] = Brush("#F7F2E6");
            Resources["SecondaryText"] = Brush("#A89E89");
            Resources["MutedText"] = Brush("#756C5A");
            Resources["EditorForeground"] = Brush("#ECE3D0");
            Resources["EditorPageBorder"] = Brush("#2E2B22");
            Resources["WindowForeground"] = Brush("#ECE3D0");
            Resources["HierarchyIndicatorBrush"] = Brush("#F2ECDD");
            Resources["DragOverBackground"] = Brush("#25F2ECDD");
            Resources["CardBackground"] = Brush("#100F0B");
            Resources["CardBorder"] = Brush("#2E2B22");
            Resources["CardAccent"] = Brush("#F2ECDD");
        }
    }
}
