using System.Text;

namespace Passage.Web.Services;

public sealed record ScriptFileEntry(string Name, DateTime LastModifiedUtc, long SizeBytes);

/// <summary>
/// Server-side script storage backed by a single flat directory (the Docker
/// volume on a NAS deployment). File names are validated so browser input can
/// never escape the library root.
/// </summary>
public sealed class ScriptLibrary
{
    private static readonly string[] AllowedExtensions = [".fountain", ".md", ".txt"];
    private readonly object _gate = new();

    public ScriptLibrary(IConfiguration configuration)
    {
        var configured = configuration["Passage:DataDir"]
            ?? Environment.GetEnvironmentVariable("PASSAGE_DATA_DIR");

        RootPath = string.IsNullOrWhiteSpace(configured)
            ? (Directory.Exists("/data") ? "/data" : Path.Combine(AppContext.BaseDirectory, "data"))
            : configured;

        Directory.CreateDirectory(RootPath);
    }

    public string RootPath { get; }

    public IReadOnlyList<ScriptFileEntry> List()
    {
        lock (_gate)
        {
            return Directory.EnumerateFiles(RootPath)
                .Where(path => AllowedExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase))
                .Select(path => new FileInfo(path))
                .OrderByDescending(info => info.LastWriteTimeUtc)
                .Select(info => new ScriptFileEntry(info.Name, info.LastWriteTimeUtc, info.Length))
                .ToList();
        }
    }

    public string Load(string name)
    {
        var path = ResolvePath(name);
        lock (_gate)
        {
            return File.ReadAllText(path, Encoding.UTF8);
        }
    }

    public void Save(string name, string content)
    {
        var path = ResolvePath(name);
        lock (_gate)
        {
            File.WriteAllText(path, content, Encoding.UTF8);
        }
    }

    public bool Exists(string name)
    {
        if (!TryValidateName(name, out var validated))
        {
            return false;
        }

        lock (_gate)
        {
            return File.Exists(Path.Combine(RootPath, validated));
        }
    }

    public void Delete(string name)
    {
        var path = ResolvePath(name);
        lock (_gate)
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// Normalizes user input into a safe library file name: strips path
    /// segments, rejects reserved characters, and appends .fountain when no
    /// allowed extension was given. Returns false when nothing usable remains.
    /// </summary>
    public static bool TryValidateName(string? name, out string validated)
    {
        validated = string.Empty;
        var trimmed = (name ?? string.Empty).Trim();
        if (trimmed.Length == 0)
        {
            return false;
        }

        if (trimmed != Path.GetFileName(trimmed))
        {
            return false;
        }

        if (trimmed.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || trimmed.StartsWith('.'))
        {
            return false;
        }

        if (!AllowedExtensions.Contains(Path.GetExtension(trimmed), StringComparer.OrdinalIgnoreCase))
        {
            trimmed += ".fountain";
        }

        validated = trimmed;
        return true;
    }

    private string ResolvePath(string name)
    {
        if (!TryValidateName(name, out var validated))
        {
            throw new ArgumentException($"Invalid script name '{name}'.", nameof(name));
        }

        return Path.Combine(RootPath, validated);
    }
}
