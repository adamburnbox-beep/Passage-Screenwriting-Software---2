using System.IO;
using System.Security;
using System.Text.Json;
using Passage.Core.Goals;

namespace Passage.App.Services;

public sealed record RecoveryDocument
{
    public string Text { get; init; } = string.Empty;

    public string? FilePath { get; init; }

    public DateTimeOffset SavedAtUtc { get; init; }

    public GoalConfiguration GoalConfiguration { get; init; } = new();

    public SessionGoalConfiguration SessionGoalConfiguration { get; init; } = new();

    public double EditorZoomPercent { get; init; } = 100.0;
}

public static class RecoveryStorage
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true
    };

    public static string RecoveryFilePath
    {
        get
        {
            var directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Passage");

            return Path.Combine(directory, "recovery.json");
        }
    }

    public static bool TryReadRecovery(out RecoveryDocument? document)
    {
        document = null;

        try
        {
            if (!File.Exists(RecoveryFilePath))
            {
                return false;
            }

            var json = File.ReadAllText(RecoveryFilePath);
            document = JsonSerializer.Deserialize<RecoveryDocument>(json, SerializerOptions);
            return document is not null;
        }
        catch (IOException)
        {
            ClearRecoveryFile();
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            ClearRecoveryFile();
            return false;
        }
        catch (SecurityException)
        {
            ClearRecoveryFile();
            return false;
        }
    }

    public static void SaveRecoveryFile(RecoveryDocument document)
    {
        try
        {
            var directory = Path.GetDirectoryName(RecoveryFilePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var json = JsonSerializer.Serialize(document, SerializerOptions);
            File.WriteAllText(RecoveryFilePath, json);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
        catch (SecurityException)
        {
        }
    }

    public static void ClearRecoveryFile()
    {
        try
        {
            if (!File.Exists(RecoveryFilePath))
            {
                return;
            }

            File.Delete(RecoveryFilePath);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
        catch (SecurityException)
        {
        }
    }
}
