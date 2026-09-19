using System.Text.Json;

namespace Passage.Web.Services;

/// <summary>
/// The per-script Slate sidecar. Worksheet working-out lives here, never in
/// the fountain file, so exports, page counts and the desktop apps' parse are
/// untouched. The sidecar is a staging area: once a runner promotes a result
/// into the script, the script is the truth (docs/SLATE-PLAN.md, decision B).
/// </summary>
public sealed class SlateDocument
{
    public int SlateVersion { get; set; } = SlateStore.CurrentVersion;
}

/// <summary>
/// Reads and writes <c>&lt;root&gt;/.slate/&lt;script name&gt;.json</c>. Keys
/// go through <see cref="ScriptLibrary.TryValidateName"/> so browser input can
/// no more escape this directory than it can the library root. The
/// dot-directory has no script extension, so <see cref="ScriptLibrary.List"/>
/// never surfaces it.
/// </summary>
public sealed class SlateStore
{
    public const int CurrentVersion = 1;

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _root;
    private readonly object _gate = new();

    public SlateStore(ScriptLibrary library)
    {
        _root = Path.Combine(library.RootPath, ".slate");
    }

    /// <summary>
    /// Returns the sidecar for a script, or null when it has none. A corrupt
    /// file throws <see cref="JsonException"/> rather than being replaced with
    /// an empty document: the caller decides how to surface lost working-out.
    /// </summary>
    public SlateDocument? Load(string scriptName)
    {
        var path = ResolvePath(scriptName);
        lock (_gate)
        {
            if (!File.Exists(path))
            {
                return null;
            }

            return JsonSerializer.Deserialize<SlateDocument>(File.ReadAllText(path), JsonOptions);
        }
    }

    public void Save(string scriptName, SlateDocument document)
    {
        var path = ResolvePath(scriptName);
        document.SlateVersion = CurrentVersion;
        lock (_gate)
        {
            Directory.CreateDirectory(_root);
            File.WriteAllText(path, JsonSerializer.Serialize(document, JsonOptions));
        }
    }

    public bool Exists(string scriptName)
    {
        if (!ScriptLibrary.TryValidateName(scriptName, out var validated))
        {
            return false;
        }

        lock (_gate)
        {
            return File.Exists(PathFor(validated));
        }
    }

    /// <summary>Cascade for the library's delete. No-op when there is no sidecar.</summary>
    public void Delete(string scriptName)
    {
        var path = ResolvePath(scriptName);
        lock (_gate)
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// Deletes every sidecar whose script is not in <paramref name="scriptNames"/>
    /// and returns the script names they belonged to, so the caller can say
    /// so rather than pruning silently. Scripts can be removed from the volume
    /// outside the app, so this runs on load, like the recent-files list.
    /// </summary>
    public IReadOnlyList<string> PruneOrphans(IEnumerable<string> scriptNames)
    {
        var present = scriptNames.ToHashSet(StringComparer.Ordinal);
        var pruned = new List<string>();
        lock (_gate)
        {
            if (!Directory.Exists(_root))
            {
                return pruned;
            }

            foreach (var path in Directory.EnumerateFiles(_root, "*.json"))
            {
                var scriptName = Path.GetFileNameWithoutExtension(path);
                if (present.Contains(scriptName))
                {
                    continue;
                }

                File.Delete(path);
                pruned.Add(scriptName);
            }
        }

        return pruned;
    }

    private string ResolvePath(string scriptName)
    {
        if (!ScriptLibrary.TryValidateName(scriptName, out var validated))
        {
            throw new ArgumentException($"Invalid script name '{scriptName}'.", nameof(scriptName));
        }

        return PathFor(validated);
    }

    private string PathFor(string validatedName) => Path.Combine(_root, validatedName + ".json");
}
