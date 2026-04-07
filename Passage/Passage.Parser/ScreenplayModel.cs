using System.Collections.ObjectModel;

namespace Passage.Parser;

public sealed class ParsedScreenplay
{
    public ParsedScreenplay(string rawText, IReadOnlyList<ScreenplayElement> elements)
        : this(rawText, TitlePageData.Empty, elements, null)
    {
    }

    public ParsedScreenplay(string rawText, TitlePageData titlePage, IReadOnlyList<ScreenplayElement> elements)
        : this(rawText, titlePage, elements, null)
    {
    }

    public ParsedScreenplay(
        string rawText,
        TitlePageData titlePage,
        IReadOnlyList<ScreenplayElement> elements,
        IReadOnlyDictionary<int, ScreenplayElementType>? lineTypeOverrides)
    {
        RawText = rawText ?? throw new ArgumentNullException(nameof(rawText));
        TitlePage = titlePage ?? throw new ArgumentNullException(nameof(titlePage));
        Elements = elements ?? throw new ArgumentNullException(nameof(elements));
        LineTypeOverrides = new ReadOnlyDictionary<int, ScreenplayElementType>(
            lineTypeOverrides is null
                ? new Dictionary<int, ScreenplayElementType>()
                : new Dictionary<int, ScreenplayElementType>(lineTypeOverrides));
    }

    public string RawText { get; }

    public TitlePageData TitlePage { get; }

    public IReadOnlyList<ScreenplayElement> Elements { get; }

    public IReadOnlyDictionary<int, ScreenplayElementType> LineTypeOverrides { get; }
}

public enum TitlePageFieldType
{
    Title,
    Episode,
    Author,
    Credit,
    DraftDate,
    Contact,
    Source,
    Revision,
    Notes,
    Custom
}

public sealed class TitlePageEntry
{
    public TitlePageEntry(TitlePageFieldType fieldType, string rawText, string value, int lineIndex, string customLabel = "")
    {
        FieldType = fieldType;
        RawText = rawText ?? throw new ArgumentNullException(nameof(rawText));
        Value = value ?? throw new ArgumentNullException(nameof(value));
        LineIndex = lineIndex;
        CustomLabel = customLabel ?? string.Empty;
    }

    public TitlePageFieldType FieldType { get; }
    public string CustomLabel { get; }

    public string Label => FieldType switch
    {
        TitlePageFieldType.Title => "Title",
        TitlePageFieldType.Episode => "Episode",
        TitlePageFieldType.Author => "Author",
        TitlePageFieldType.Credit => "Credit",
        TitlePageFieldType.DraftDate => "Draft date",
        TitlePageFieldType.Contact => "Contact",
        TitlePageFieldType.Source => "Source",
        TitlePageFieldType.Revision => "Revision",
        TitlePageFieldType.Notes => "Notes",
        TitlePageFieldType.Custom => CustomLabel,
        _ => FieldType.ToString()
    };

    public string RawText { get; }

    public string Value { get; }

    public int LineIndex { get; }

    public int LineNumber => LineIndex + 1;
}

public sealed class TitlePageData
{
    public static TitlePageData Empty { get; } = new(Array.Empty<TitlePageEntry>(), 0);

    public TitlePageData(IReadOnlyList<TitlePageEntry> entries, int bodyStartLineIndex)
    {
        Entries = entries ?? throw new ArgumentNullException(nameof(entries));
        BodyStartLineIndex = bodyStartLineIndex < 0 ? 0 : bodyStartLineIndex;
    }

    public IReadOnlyList<TitlePageEntry> Entries { get; }

    public int BodyStartLineIndex { get; }
}

public enum ScreenplayElementType
{
    Section,
    Synopsis,
    Note,
    Boneyard,
    CenteredText,
    Lyrics,
    SceneHeading,
    Action,
    Transition,
    Character,
    Parenthetical,
    Dialogue,
    TitlePageCentered,
    TitlePageContact
}

public abstract class ScreenplayElement
{
    protected ScreenplayElement(
        ScreenplayElementType type,
        string text,
        string rawText,
        int lineIndex,
        int endLineIndex,
        int level = 3)
    {
        if (lineIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(lineIndex));
        }

        if (endLineIndex < lineIndex)
        {
            throw new ArgumentOutOfRangeException(nameof(endLineIndex));
        }

        Type = type;
        Text = text ?? throw new ArgumentNullException(nameof(text));
        RawText = rawText ?? throw new ArgumentNullException(nameof(rawText));
        LineIndex = lineIndex;
        EndLineIndex = endLineIndex;
        Level = level < 0 ? throw new ArgumentOutOfRangeException(nameof(level)) : level;
    }

    public Guid Id { get; set; } = Guid.NewGuid();

    public ScreenplayElementType Type { get; }

    public string Text { get; }

    public string RawText { get; }

    public int Level { get; set; }

    public bool IsCollapsed { get; set; }

    public string? BodyText { get; set; }

    public bool IsSuppressed { get; set; }

    public bool IsDraft { get; set; }

    public virtual string Heading => Type switch
    {
        ScreenplayElementType.SceneHeading => FirstNonEmptyLineOrFallback(Text, "Scene"),
        ScreenplayElementType.Section => FirstNonEmptyLineOrFallback(Text, "Section"),
        ScreenplayElementType.Synopsis => "Synopsis",
        ScreenplayElementType.Note => "Note",
        ScreenplayElementType.Boneyard => "Boneyard",
        ScreenplayElementType.Character => FirstNonEmptyLineOrFallback(Text, "Character"),
        ScreenplayElementType.Transition => FirstNonEmptyLineOrFallback(Text, "Transition"),
        _ => KindLabel
    };

    public virtual string ScriptHeading => Type == ScreenplayElementType.SceneHeading
        ? FirstNonEmptyLineOrFallback(Text, Heading)
        : string.Empty;

    public virtual string Description => NormalizeCardText(Text);

    public virtual string BoardDescription => Type switch
    {
        ScreenplayElementType.SceneHeading => NormalizeCardText(BodyText),
        ScreenplayElementType.Section => NormalizeCardText(BodyText),
        ScreenplayElementType.Note => NormalizeCardText(Text),
        _ => NormalizeCardText(BodyText)
    };

    public string PreviewText
    {
        get
        {
            var preview = BuildPreviewText(Description);
            return preview.Length > 0 ? preview : Heading;
        }
    }

    public string KindLabel => Type switch
    {
        ScreenplayElementType.SceneHeading => "Scene",
        ScreenplayElementType.Note => "Note",
        ScreenplayElementType.Section => Level switch
        {
            0 => "Act",
            1 => "Sequence",
            _ => "Beat"
        },
        ScreenplayElementType.Synopsis => "Synopsis",
        ScreenplayElementType.Action => "Beat",
        ScreenplayElementType.Dialogue => "Dialogue",
        ScreenplayElementType.Parenthetical => "Parenthetical",
        ScreenplayElementType.Character => "Character",
        ScreenplayElementType.Transition => "Transition",
        ScreenplayElementType.CenteredText => "Centered",
        ScreenplayElementType.Lyrics => "Lyrics",
        ScreenplayElementType.Boneyard => "Boneyard",
        _ => Type.ToString()
    };

    public int LineIndex { get; }

    public int LineNumber => LineIndex + 1;

    public int EndLineIndex { get; }

    public int EndLineNumber => EndLineIndex + 1;

    public int StartLine => LineNumber;

    public int EndLine => EndLineNumber;

    private static string NormalizeCardText(string? value)
    {
        return (value ?? string.Empty)
            .ReplaceLineEndings("\n")
            .Trim();
    }

    private static string FirstNonEmptyLineOrFallback(string value, string fallback)
    {
        var normalized = NormalizeCardText(value);
        if (normalized.Length == 0)
        {
            return fallback;
        }

        foreach (var line in normalized.Split('\n'))
        {
            var trimmedLine = line.Trim();
            if (trimmedLine.Length > 0)
            {
                return trimmedLine;
            }
        }

        return fallback;
    }

    private static string BuildPreviewText(string value)
    {
        var normalized = NormalizeCardText(value);
        if (normalized.Length == 0)
        {
            return string.Empty;
        }

        var parts = normalized
            .Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(line => line.Trim())
            .Where(line => line.Length > 0);

        return string.Join(" ", parts);
    }
}

public sealed class SectionElement : ScreenplayElement
{
    public SectionElement(string text, string rawText, int lineIndex, int sectionDepth)
        : base(ScreenplayElementType.Section, text, rawText, lineIndex, lineIndex, Math.Max(0, sectionDepth - 1))
    {
        if (sectionDepth < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(sectionDepth));
        }

        SectionDepth = sectionDepth;
    }

    public int SectionDepth { get; }
}

public sealed class SynopsisElement : ScreenplayElement
{
    public SynopsisElement(string text, string rawText, int lineIndex, int? sectionLevel)
        : base(ScreenplayElementType.Synopsis, text, rawText, lineIndex, lineIndex)
    {
        SectionLevel = sectionLevel;
    }

    public int? SectionLevel { get; }
}

public sealed class NoteElement : ScreenplayElement
{
    public NoteElement(string text, string rawText, int lineIndex, int endLineIndex, bool isClosed)
        : base(ScreenplayElementType.Note, text, rawText, lineIndex, endLineIndex)
    {
        IsClosed = isClosed;
    }

    public bool IsClosed { get; }
}

public sealed class BoneyardElement : ScreenplayElement
{
    public BoneyardElement(string text, string rawText, int lineIndex, int endLineIndex, bool isClosed)
        : base(ScreenplayElementType.Boneyard, text, rawText, lineIndex, endLineIndex)
    {
        IsClosed = isClosed;
    }

    public bool IsClosed { get; }
}

public sealed class CenteredTextElement : ScreenplayElement
{
    public CenteredTextElement(string text, string rawText, int lineIndex)
        : base(ScreenplayElementType.CenteredText, text, rawText, lineIndex, lineIndex)
    {
    }
}

public sealed class LyricsElement : ScreenplayElement
{
    public LyricsElement(string text, string rawText, int lineIndex)
        : base(ScreenplayElementType.Lyrics, text, rawText, lineIndex, lineIndex)
    {
    }
}

public sealed class SceneHeadingElement : ScreenplayElement
{
    public SceneHeadingElement(string text, string rawText, int lineIndex, bool isForced)
        : base(ScreenplayElementType.SceneHeading, text, rawText, lineIndex, lineIndex, 2)
    {
        IsForced = isForced;
    }

    public bool IsForced { get; }
}

public sealed class ActionElement : ScreenplayElement
{
    public ActionElement(string text, string rawText, int lineIndex, int endLineIndex)
        : base(ScreenplayElementType.Action, text, rawText, lineIndex, endLineIndex)
    {
    }
}

public sealed class TransitionElement : ScreenplayElement
{
    public TransitionElement(string text, string rawText, int lineIndex, bool isForced)
        : base(ScreenplayElementType.Transition, text, rawText, lineIndex, lineIndex)
    {
        IsForced = isForced;
    }

    public bool IsForced { get; }
}

public sealed class CharacterElement : ScreenplayElement
{
    public CharacterElement(string characterName, string rawText, int lineIndex, bool isDualDialogue)
        : base(ScreenplayElementType.Character, characterName, rawText, lineIndex, lineIndex)
    {
        CharacterName = characterName;
        IsDualDialogue = isDualDialogue;
    }

    public string CharacterName { get; }

    public bool IsDualDialogue { get; }
}

public sealed class ParentheticalElement : ScreenplayElement
{
    public ParentheticalElement(
        string text,
        string rawText,
        int lineIndex,
        string? characterName,
        bool isDualDialogue)
        : base(ScreenplayElementType.Parenthetical, text, rawText, lineIndex, lineIndex)
    {
        CharacterName = characterName;
        IsDualDialogue = isDualDialogue;
    }

    public string? CharacterName { get; }

    public bool IsDualDialogue { get; }
}

public sealed class DialogueElement : ScreenplayElement
{
    public DialogueElement(
        string characterName,
        IReadOnlyList<string> parentheticals,
        IReadOnlyList<string> lines,
        string text,
        string rawText,
        int lineIndex,
        int endLineIndex,
        bool isDualDialogue)
        : base(ScreenplayElementType.Dialogue, text, rawText, lineIndex, endLineIndex)
    {
        CharacterName = characterName;
        Parentheticals = parentheticals ?? throw new ArgumentNullException(nameof(parentheticals));
        Lines = lines ?? throw new ArgumentNullException(nameof(lines));
        IsDualDialogue = isDualDialogue;
    }

    public string CharacterName { get; }

    public IReadOnlyList<string> Parentheticals { get; }

    public IReadOnlyList<string> Lines { get; }

    public bool IsDualDialogue { get; }
}

public sealed class ScratchpadCardElement : ScreenplayElement
{
    public ScratchpadCardElement(
        ScreenplayElementType type,
        string heading,
        string description,
        string text,
        string rawText,
        int level,
        int? sourceLineIndex = null,
        int? sourceEndLineIndex = null,
        string? scriptHeading = null)
        : base(
            type,
            text,
            rawText,
            sourceLineIndex ?? 0,
            sourceEndLineIndex ?? sourceLineIndex ?? 0,
            level)
    {
        if (sourceLineIndex.HasValue != sourceEndLineIndex.HasValue)
        {
            throw new ArgumentException("Scratchpad source line range must provide both start and end values.");
        }

        Heading = string.IsNullOrWhiteSpace(heading) ? KindLabel : heading.Trim();
        Description = (description ?? string.Empty).ReplaceLineEndings("\n").Trim();
        ScriptHeading = type == ScreenplayElementType.SceneHeading
            ? NormalizeSceneHeading(scriptHeading, Heading)
            : string.Empty;
        SourceLineIndex = sourceLineIndex;
        SourceEndLineIndex = sourceEndLineIndex;
        IsDraft = true;
        IsCollapsed = true;
    }

    public override string Heading { get; }

    public override string ScriptHeading { get; }

    public override string Description { get; }

    public override string BoardDescription => Description;

    public int? SourceLineIndex { get; }

    public int? SourceEndLineIndex { get; }

    private static string NormalizeSceneHeading(string? scriptHeading, string fallback)
    {
        var trimmedHeading = (scriptHeading ?? string.Empty).Trim();
        return trimmedHeading.Length > 0 ? trimmedHeading : fallback;
    }
}
