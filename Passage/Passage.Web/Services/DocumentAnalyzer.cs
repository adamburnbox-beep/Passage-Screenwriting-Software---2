using Passage.Core;
using Passage.Export;
using Passage.Parser;

namespace Passage.Web.Services;

public sealed record OutlineNode(string Label, string Kind, int LineNumber, List<OutlineNode> Children, Guid Id = default);

public sealed record NoteEntry(string Text, int LineNumber);

public sealed record PreviewLine(string Text, LayoutTextStyle Style, double IndentChars);

public sealed record PreviewPage(IReadOnlyList<PreviewLine> Lines, int PageNumber, bool IsTitlePage);

public sealed record BoardCard(string Heading, string Kind, string Description, int LineNumber, List<BoardCard> Children, Guid Id = default);

/// <summary>A Sequence and the cards beneath it. SequenceCard is null for cards
/// that appear before any Sequence heading.</summary>
public sealed record BoardGroup(BoardCard? SequenceCard, List<BoardCard> Cards);

/// <summary>An Act and the Sequence groups beneath it. ActCard is null for
/// groups that appear before any Act heading.</summary>
public sealed record BoardLane(BoardCard? ActCard, List<BoardGroup> Groups);

public sealed record DocumentAnalysis(
    string[] LineClasses,
    IReadOnlyList<OutlineNode> Outline,
    IReadOnlyList<NoteEntry> Notes,
    IReadOnlyList<PreviewPage> PreviewPages,
    IReadOnlyList<BoardLane> BoardLanes,
    int WordCount,
    int PageCount,
    IReadOnlyList<TitlePageEntry> TitlePage,
    // First line of the script proper. The title-page editor splices lines
    // 0..BodyStart, so it needs the boundary the parser already worked out.
    int TitlePageBodyStart,
    // The parsed elements and the document's line count, which
    // BeatBoardText.GetCardLineRange needs to resolve a card to a line range.
    IReadOnlyList<ScreenplayElement> Elements,
    int LineCount,
    // Autocomplete sources, pushed to the client after each parse so typing
    // never round-trips (trap 4).
    IReadOnlyList<string> SceneHeadingSuggestions,
    IReadOnlyList<string> CharacterSuggestions,
    // Per line: "scene", "character" or "". Which list (if any) autocomplete
    // should offer there. Decided here so the shared TextAnalysis helpers stay
    // the single implementation; the client only does the prefix matching.
    string[] SuggestionKinds);

/// <summary>
/// Runs the shared Fountain pipeline (parser + layout builder) over the editor
/// text and projects the result into the shapes the web UI renders: per-line
/// syntax classes for the editor overlay, the outline/notes trees, paginated
/// preview pages, and beat-board cards.
/// </summary>
public sealed class DocumentAnalyzer
{
    private readonly FountainParser _parser = new();

    public DocumentAnalysis Analyze(string text) => Analyze(text, null);

    /// <summary>
    /// <paramref name="lineTypeOverrides"/> is the "Classify As" map, keyed by
    /// 1-based line number — the same shape the parser takes on the desktop.
    /// </summary>
    public DocumentAnalysis Analyze(
        string text, IReadOnlyDictionary<int, ScreenplayElementType>? lineTypeOverrides)
    {
        var parsed = _parser.Parse(text, lineTypeOverrides);
        var lineCount = CountLines(text);

        SuppressCardSynopses(parsed.Elements);

        var lineClasses = BuildLineClasses(parsed, lineCount);
        var outline = BuildOutline(parsed.Elements);
        var notes = parsed.Elements
            .OfType<NoteElement>()
            .Select(note => new NoteEntry(note.Description, note.LineNumber))
            .ToList();
        var previewPages = BuildPreviewPages(parsed);
        var boardLanes = BuildBoardLanes(outline);
        CollectSuggestions(parsed.Elements, out var sceneHeadings, out var characters);
        var wordCount = TextAnalysis.CountWords(text);
        var pageCount = previewPages.Count(page => !page.IsTitlePage);

        return new DocumentAnalysis(
            lineClasses,
            outline,
            notes,
            previewPages,
            boardLanes,
            wordCount,
            pageCount,
            parsed.TitlePage.Entries,
            parsed.TitlePage.BodyStartLineIndex,
            parsed.Elements,
            lineCount,
            sceneHeadings,
            characters,
            BuildSuggestionKinds(parsed, text, lineCount));
    }

    public static DocumentAnalysis AnalyzeMarkdown(string text)
    {
        var lines = SplitLines(text);
        var lineClasses = new string[lines.Length];
        var outline = new List<OutlineNode>();
        var stack = new List<(int Level, OutlineNode Node)>();

        for (var index = 0; index < lines.Length; index++)
        {
            var line = lines[index].TrimStart();
            if (!line.StartsWith('#'))
            {
                lineClasses[index] = string.Empty;
                continue;
            }

            var depth = 0;
            while (depth < line.Length && line[depth] == '#')
            {
                depth++;
            }

            lineClasses[index] = "md-heading";
            var label = line[depth..].Trim();
            if (label.Length == 0)
            {
                continue;
            }

            var node = new OutlineNode(label, $"H{Math.Min(depth, 6)}", index + 1, new List<OutlineNode>());
            while (stack.Count > 0 && stack[^1].Level >= depth)
            {
                stack.RemoveAt(stack.Count - 1);
            }

            if (stack.Count == 0)
            {
                outline.Add(node);
            }
            else
            {
                stack[^1].Node.Children.Add(node);
            }

            stack.Add((depth, node));
        }

        return new DocumentAnalysis(
            lineClasses,
            outline,
            Array.Empty<NoteEntry>(),
            Array.Empty<PreviewPage>(),
            Array.Empty<BoardLane>(),
            TextAnalysis.CountWords(text),
            0,
            Array.Empty<TitlePageEntry>(),
            0,
            Array.Empty<ScreenplayElement>(),
            CountLines(text),
            Array.Empty<string>(),
            Array.Empty<string>(),
            Array.Empty<string>());
    }

    private static string[] BuildLineClasses(ParsedScreenplay parsed, int lineCount)
    {
        var classes = new string[lineCount];
        Array.Fill(classes, string.Empty);

        void Mark(int startIndex, int endIndex, string cssClass)
        {
            for (var index = startIndex; index <= endIndex && index < lineCount; index++)
            {
                if (index >= 0)
                {
                    classes[index] = cssClass;
                }
            }
        }

        foreach (var entry in parsed.TitlePage.Entries)
        {
            var endIndex = entry.LineIndex + Math.Max(0, entry.RawText.Split('\n').Length - 1);
            Mark(entry.LineIndex, endIndex, "sx-titlepage");
        }

        foreach (var element in parsed.Elements)
        {
            var cssClass = element.Type switch
            {
                ScreenplayElementType.SceneHeading => "sx-scene",
                ScreenplayElementType.Character => "sx-character",
                ScreenplayElementType.Dialogue => "sx-dialogue",
                ScreenplayElementType.Parenthetical => "sx-paren",
                ScreenplayElementType.Transition => "sx-transition",
                ScreenplayElementType.Section => "sx-section",
                ScreenplayElementType.Synopsis => "sx-synopsis",
                ScreenplayElementType.Note => "sx-note",
                ScreenplayElementType.Boneyard => "sx-boneyard",
                ScreenplayElementType.CenteredText => "sx-centered",
                ScreenplayElementType.Lyrics => "sx-lyrics",
                _ => string.Empty
            };

            if (cssClass.Length > 0)
            {
                Mark(element.LineIndex, element.EndLineIndex, cssClass);
            }
        }

        return classes;
    }

    private static List<OutlineNode> BuildOutline(IReadOnlyList<ScreenplayElement> elements)
    {
        const int SceneStackLevel = 99;
        var roots = new List<OutlineNode>();
        var stack = new List<(int Level, OutlineNode Node)>();

        void PopTo(int level)
        {
            while (stack.Count > 0 && stack[^1].Level >= level)
            {
                stack.RemoveAt(stack.Count - 1);
            }
        }

        void Attach(OutlineNode node)
        {
            if (stack.Count == 0)
            {
                roots.Add(node);
            }
            else
            {
                stack[^1].Node.Children.Add(node);
            }
        }

        foreach (var element in elements)
        {
            switch (element)
            {
                case SectionElement section:
                {
                    var node = new OutlineNode(section.Heading, section.KindLabel, section.LineNumber, new List<OutlineNode>(), section.Id);
                    PopTo(section.SectionDepth);
                    Attach(node);
                    stack.Add((section.SectionDepth, node));
                    break;
                }
                case SceneHeadingElement scene:
                {
                    var node = new OutlineNode(scene.Heading, "Scene", scene.LineNumber, new List<OutlineNode>(), scene.Id);
                    PopTo(SceneStackLevel);
                    Attach(node);
                    stack.Add((SceneStackLevel, node));
                    break;
                }
                case SynopsisElement synopsis:
                {
                    Attach(new OutlineNode(synopsis.Description, "Synopsis", synopsis.LineNumber, new List<OutlineNode>()));
                    break;
                }
            }
        }

        return roots;
    }

    private static List<PreviewPage> BuildPreviewPages(ParsedScreenplay parsed)
    {
        var (titlePages, bodyPages) = ScreenplayLayoutBuilder.BuildPages(parsed);
        var pages = new List<PreviewPage>(titlePages.Count + bodyPages.Count);

        foreach (var page in titlePages.Concat(bodyPages))
        {
            var lines = page.Lines
                .Select(line => new PreviewLine(
                    line.Text,
                    line.Style,
                    Math.Max(0, (line.X - ScreenplayLayoutBuilder.MarginLeft) / ScreenplayLayoutBuilder.CharWidth)))
                .ToList();
            pages.Add(new PreviewPage(lines, page.PageNumber, page.IsTitlePage));
        }

        return pages;
    }

    /// <summary>
    /// Groups the outline into Act lanes containing Sequence groups containing
    /// cards, mirroring the Linux RebuildBeatBoardLanes. Cards that appear
    /// before any Act or Sequence heading get an implicit lane or group, so a
    /// script with no structure markers still renders.
    /// </summary>
    /// <summary>
    /// Marks the synopsis lines that follow a heading as suppressed, because the
    /// board shows them as that card's description rather than as cards of their
    /// own. Mirrors what the Avalonia board build does, and is what lets
    /// BeatBoardText.GetCardLineRange include them in the card's own range —
    /// without it, editing a description would leave the old "= " lines behind.
    /// </summary>
    /// <summary>
    /// The unique scene headings and character names in the script, uppercased
    /// and sorted. Ports UpdateUniqueScreenplayElements, including taking
    /// character names from Dialogue elements as well as Character ones.
    /// </summary>
    /// <summary>
    /// Which suggestion list belongs on each line. Ports
    /// GetLatestEffectiveLineType: the parse wins where it has an opinion, and
    /// otherwise the same live-cue fallback applies, so a half-typed character
    /// name is offered completions before the parser can commit to it.
    /// </summary>
    private static string[] BuildSuggestionKinds(ParsedScreenplay parsed, string text, int lineCount)
    {
        var kinds = new string[lineCount];
        for (var i = 0; i < kinds.Length; i++)
        {
            kinds[i] = string.Empty;
        }

        foreach (var element in parsed.Elements)
        {
            var kind = element.Type switch
            {
                ScreenplayElementType.SceneHeading => "scene",
                ScreenplayElementType.Character => "character",
                _ => string.Empty
            };

            if (kind.Length == 0)
            {
                continue;
            }

            for (var line = element.LineIndex; line <= element.EndLineIndex && line < kinds.Length; line++)
            {
                if (line >= 0)
                {
                    kinds[line] = kind;
                }
            }
        }

        // The fallback, for lines the parse has no element for — a name being
        // typed on a fresh line is the case that matters.
        var lines = text.Replace("\r\n", "\n").Split('\n');
        for (var i = 0; i < kinds.Length && i < lines.Length; i++)
        {
            if (kinds[i].Length > 0)
            {
                continue;
            }

            var trimmed = lines[i].Trim();
            if (trimmed.Length == 0)
            {
                continue;
            }

            if (TextAnalysis.LooksLikeSceneHeadingStart(trimmed.AsSpan()))
            {
                kinds[i] = "scene";
            }
            else if (TextAnalysis.IsLiveCharacterCueCandidate(lines[i].AsSpan(), 45, 6))
            {
                kinds[i] = "character";
            }
        }

        return kinds;
    }

    private static void CollectSuggestions(
        IReadOnlyList<ScreenplayElement> elements,
        out IReadOnlyList<string> sceneHeadings,
        out IReadOnlyList<string> characters)
    {
        var headingSet = new HashSet<string>(StringComparer.Ordinal);
        var characterSet = new HashSet<string>(StringComparer.Ordinal);

        foreach (var element in elements)
        {
            switch (element)
            {
                case SceneHeadingElement:
                {
                    var heading = element.Text.Trim();
                    if (heading.Length > 0)
                    {
                        headingSet.Add(heading.ToUpperInvariant());
                    }

                    break;
                }

                case CharacterElement character:
                {
                    var name = character.CharacterName.Trim();
                    if (name.Length > 0)
                    {
                        characterSet.Add(name.ToUpperInvariant());
                    }

                    break;
                }

                case DialogueElement dialogue:
                {
                    var name = dialogue.CharacterName.Trim();
                    if (name.Length > 0)
                    {
                        characterSet.Add(name.ToUpperInvariant());
                    }

                    break;
                }
            }
        }

        sceneHeadings = headingSet.OrderBy(entry => entry, StringComparer.Ordinal).ToList();
        characters = characterSet.OrderBy(entry => entry, StringComparer.Ordinal).ToList();
    }

    private static void SuppressCardSynopses(IReadOnlyList<ScreenplayElement> elements)
    {
        for (var i = 0; i < elements.Count; i++)
        {
            if (elements[i].Type is not (ScreenplayElementType.Section
                or ScreenplayElementType.SceneHeading
                or ScreenplayElementType.Note))
            {
                continue;
            }

            for (var j = i + 1; j < elements.Count; j++)
            {
                if (elements[j].Type != ScreenplayElementType.Synopsis)
                {
                    break;
                }

                elements[j].IsSuppressed = true;
            }
        }
    }

    private static List<BoardLane> BuildBoardLanes(IReadOnlyList<OutlineNode> outline)
    {
        var lanes = new List<BoardLane>();
        BoardLane? currentLane = null;
        BoardGroup? currentGroup = null;

        foreach (var card in FlattenBoardCards(outline))
        {
            if (card.Kind == "Act")
            {
                currentLane = new BoardLane(card, new List<BoardGroup>());
                currentGroup = null;
                lanes.Add(currentLane);
                continue;
            }

            if (currentLane is null)
            {
                currentLane = new BoardLane(null, new List<BoardGroup>());
                lanes.Add(currentLane);
            }

            if (card.Kind == "Sequence")
            {
                currentGroup = new BoardGroup(card, new List<BoardCard>());
                currentLane.Groups.Add(currentGroup);
                continue;
            }

            if (currentGroup is null)
            {
                currentGroup = new BoardGroup(null, new List<BoardCard>());
                currentLane.Groups.Add(currentGroup);
            }

            currentGroup.Cards.Add(card);
        }

        return lanes;
    }

    /// <summary>
    /// Walks the outline in document order and yields one flat card per node,
    /// which is the shape RebuildBeatBoardLanes consumes. Synopsis children
    /// become their parent's description rather than cards of their own.
    /// </summary>
    private static IEnumerable<BoardCard> FlattenBoardCards(IEnumerable<OutlineNode> nodes)
    {
        foreach (var node in nodes)
        {
            if (node.Kind == "Synopsis")
            {
                continue;
            }

            var synopsis = string.Join("\n", node.Children
                .Where(child => child.Kind == "Synopsis")
                .Select(child => child.Label));

            yield return new BoardCard(
                node.Label, node.Kind, synopsis, node.LineNumber, new List<BoardCard>(), node.Id);

            foreach (var child in FlattenBoardCards(node.Children))
            {
                yield return child;
            }
        }
    }


    private static int CountLines(string text)
    {
        return SplitLines(text).Length;
    }

    private static string[] SplitLines(string text)
    {
        return (text ?? string.Empty).ReplaceLineEndings("\n").Split('\n');
    }
}
