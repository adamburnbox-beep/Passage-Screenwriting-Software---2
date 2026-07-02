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

    // The single runtime source of truth for both palettes — a modern e-ink
    // monochrome scheme: paper-white surfaces with true-ink black in light, and
    // its direct inverse in dark. Called at startup (system detection above)
    // and by the View > Theme commands. Keep App.axaml's startup defaults in sync.
    public void LoadThemeResources(bool isLight)
    {
        if (Resources is null) return;

        if (isLight)
        {
            // Light theme (paper white / ink black)
            Resources["ThemeBackground"] = Brush("#EDEDEB");
            Resources["WindowBackground"] = Brush("#EDEDEB");
            Resources["SurfaceBackground"] = Brush("#F4F4F2");
            Resources["SurfaceMutedBackground"] = Brush("#E7E7E4");
            Resources["SurfaceRaisedBackground"] = Brush("#FFFFFF");
            Resources["SurfaceBorder"] = Brush("#DEDEDA");
            Resources["ControlBackground"] = Brush("#FFFFFF");
            Resources["ControlForeground"] = Brush("#111111");
            Resources["ControlBorder"] = Brush("#D6D6D2");
            Resources["ControlAccent"] = Brush("#111111");
            Resources["ControlPressedBackground"] = Brush("#111111");
            Resources["ControlPressedForeground"] = Brush("#FFFFFF");
            Resources["HeaderText"] = Brush("#0A0A0A");
            Resources["SecondaryText"] = Brush("#4A4A47");
            Resources["MutedText"] = Brush("#8E8E8A");
            Resources["EditorForeground"] = Brush("#161616");
            Resources["EditorPageBorder"] = Brush("#E3E3DF");
            Resources["WindowForeground"] = Brush("#161616");
            Resources["HierarchyIndicatorBrush"] = Brush("#111111");
            Resources["DragOverBackground"] = Brush("#22111111");
            Resources["CardBackground"] = Brush("#FFFFFF");
            Resources["CardBorder"] = Brush("#E3E3DF");
            Resources["CardAccent"] = Brush("#111111");

            // Editor syntax colours, darkened for the paper-white page
            Resources["SyntaxSceneHeading"] = Brush("#2F55B8");
            Resources["SyntaxCharacter"] = Brush("#94306A");
            Resources["SyntaxDialogue"] = Brush("#3C3C39");
            Resources["SyntaxParenthetical"] = Brush("#757570");
            Resources["SyntaxTransition"] = Brush("#5B4392");
            Resources["SyntaxSection"] = Brush("#1D6FA5");
            Resources["SyntaxSynopsis"] = Brush("#9A6A1C");
            Resources["SyntaxNote"] = Brush("#3E7D44");
        }
        else
        {
            // Dark theme (ink black / paper white; inverse of light)
            Resources["ThemeBackground"] = Brush("#0B0B0B");
            Resources["WindowBackground"] = Brush("#0B0B0B");
            Resources["SurfaceBackground"] = Brush("#111111");
            Resources["SurfaceMutedBackground"] = Brush("#1E1E1E");
            Resources["SurfaceRaisedBackground"] = Brush("#161616");
            Resources["SurfaceBorder"] = Brush("#242424");
            Resources["ControlBackground"] = Brush("#161616");
            Resources["ControlForeground"] = Brush("#F2F2F0");
            Resources["ControlBorder"] = Brush("#2E2E2E");
            Resources["ControlAccent"] = Brush("#F5F5F3");
            Resources["ControlPressedBackground"] = Brush("#F5F5F3");
            Resources["ControlPressedForeground"] = Brush("#0B0B0B");
            Resources["HeaderText"] = Brush("#FAFAF8");
            Resources["SecondaryText"] = Brush("#B4B4B0");
            Resources["MutedText"] = Brush("#70706C");
            Resources["EditorForeground"] = Brush("#EDEDEB");
            Resources["EditorPageBorder"] = Brush("#242424");
            Resources["WindowForeground"] = Brush("#EDEDEB");
            Resources["HierarchyIndicatorBrush"] = Brush("#F5F5F3");
            Resources["DragOverBackground"] = Brush("#25F5F5F3");
            Resources["CardBackground"] = Brush("#161616");
            Resources["CardBorder"] = Brush("#262626");
            Resources["CardAccent"] = Brush("#F5F5F3");

            // Editor syntax colours for the ink-black page
            Resources["SyntaxSceneHeading"] = Brush("#6797FF");
            Resources["SyntaxCharacter"] = Brush("#C05587");
            Resources["SyntaxDialogue"] = Brush("#CEB2C9");
            Resources["SyntaxParenthetical"] = Brush("#8A8A85");
            Resources["SyntaxTransition"] = Brush("#8A6FC9");
            Resources["SyntaxSection"] = Brush("#4FC3F7");
            Resources["SyntaxSynopsis"] = Brush("#FFB74D");
            Resources["SyntaxNote"] = Brush("#81C784");
        }
    }
}
