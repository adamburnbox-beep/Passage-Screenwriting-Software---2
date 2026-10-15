using System.Text;
using System.Xml;
using Passage.Parser;

namespace Passage.Export;

/// <summary>
/// Writes Final Draft <c>.fdx</c> XML. Notes, boneyard, sections and synopses
/// are non-printing and are left out, as in the PDF.
/// </summary>
public sealed class FdxExporter : IExporter
{
    public string DisplayName => "Final Draft";

    public string DefaultExtension => ".fdx";

    public void Export(ParsedScreenplay screenplay, string filePath)
    {
        ArgumentNullException.ThrowIfNull(filePath);

        File.WriteAllText(filePath, GetExportText(screenplay), new UTF8Encoding(false));
    }

    public string GetExportText(ParsedScreenplay screenplay)
    {
        ArgumentNullException.ThrowIfNull(screenplay);

        var builder = new StringBuilder();
        var settings = new XmlWriterSettings { Indent = true, Encoding = new UTF8Encoding(false) };
        using (var writer = XmlWriter.Create(new StringWriterWithEncoding(builder, settings.Encoding), settings))
        {
            writer.WriteStartDocument(standalone: false);
            writer.WriteStartElement("FinalDraft");
            writer.WriteAttributeString("DocumentType", "Script");
            writer.WriteAttributeString("Template", "No");
            writer.WriteAttributeString("Version", "5");

            writer.WriteStartElement("Content");
            foreach (var item in BuildItems(screenplay.Elements))
            {
                item.Write(writer);
            }
            writer.WriteEndElement();

            WriteTitlePage(writer, screenplay.TitlePage);

            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        return builder.ToString();
    }

    // A body item is one paragraph, a speaker's block, or two blocks side by side.
    private abstract record Item
    {
        public abstract void Write(XmlWriter writer);
    }

    private sealed record ParagraphItem(string Type, string Text, string? Number = null, string? Alignment = null) : Item
    {
        public override void Write(XmlWriter writer) => WriteParagraph(writer, Type, Text, Number, Alignment);
    }

    private sealed record DialogueBlockItem(IReadOnlyList<ParagraphItem> Paragraphs) : Item
    {
        public override void Write(XmlWriter writer)
        {
            foreach (var paragraph in Paragraphs)
            {
                paragraph.Write(writer);
            }
        }
    }

    private sealed record DualDialogueItem(DialogueBlockItem Left, DialogueBlockItem Right) : Item
    {
        public override void Write(XmlWriter writer)
        {
            writer.WriteStartElement("Paragraph");
            writer.WriteStartElement("DualDialogue");
            Left.Write(writer);
            Right.Write(writer);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }
    }

    private static List<Item> BuildItems(IReadOnlyList<ScreenplayElement> elements)
    {
        var items = new List<Item>();
        for (var index = 0; index < elements.Count; index++)
        {
            var element = elements[index];
            switch (element)
            {
                case NoteElement or BoneyardElement or SectionElement or SynopsisElement:
                    break;
                case CharacterElement character:
                    var block = BuildDialogueBlock(character, elements, ref index);
                    if (character.IsDualDialogue && items.Count > 0 && items[^1] is DialogueBlockItem previous)
                    {
                        items[^1] = new DualDialogueItem(previous, block);
                    }
                    else
                    {
                        items.Add(block);
                    }
                    break;
                case SceneHeadingElement sceneHeading:
                    AddParagraph(items, "Scene Heading", Clean(sceneHeading.Text).ToUpperInvariant(), sceneHeading.SceneNumber);
                    break;
                case ParentheticalElement parenthetical:
                    AddParagraph(items, "Parenthetical", Clean($"({parenthetical.Text})"));
                    break;
                case DialogueElement dialogue:
                    AddLines(items, "Dialogue", dialogue.Text);
                    break;
                case TransitionElement transition:
                    AddParagraph(items, "Transition", Clean(transition.Text).ToUpperInvariant());
                    break;
                case CenteredTextElement centered:
                    AddLines(items, "Action", centered.Text, "Center");
                    break;
                case LyricsElement lyrics:
                    AddLines(items, "General", lyrics.Text);
                    break;
                default:
                    AddLines(items, "Action", element.Text);
                    break;
            }
        }

        return items;
    }

    // Takes the cue and the parentheticals and dialogue that follow it, advancing index past them.
    // The parser emits each parenthetical before one combined DialogueElement, so line order is
    // rebuilt from that element's source lines and the parentheticals it covers are skipped.
    private static DialogueBlockItem BuildDialogueBlock(CharacterElement character, IReadOnlyList<ScreenplayElement> elements, ref int index)
    {
        var paragraphs = new List<ParagraphItem>
        {
            new("Character", Clean(character.CharacterName).TrimStart('@').Trim().ToUpperInvariant())
        };

        var end = index + 1;
        while (end < elements.Count && elements[end] is ParentheticalElement or DialogueElement)
        {
            end++;
        }

        var dialogue = elements.Skip(index + 1).Take(end - index - 1).OfType<DialogueElement>().FirstOrDefault();
        for (var cursor = index + 1; cursor < end; cursor++)
        {
            switch (elements[cursor])
            {
                case ParentheticalElement parenthetical
                    when dialogue is null || parenthetical.LineIndex < dialogue.LineIndex || parenthetical.LineIndex > dialogue.EndLineIndex:
                    paragraphs.Add(new ParagraphItem("Parenthetical", Clean($"({parenthetical.Text})")));
                    break;
                case DialogueElement spoken:
                    foreach (var line in SplitLines(Clean(spoken.RawText)))
                    {
                        paragraphs.Add(line.StartsWith('(') && line.EndsWith(')')
                            ? new ParagraphItem("Parenthetical", line)
                            : new ParagraphItem("Dialogue", line));
                    }
                    break;
            }
        }

        index = end - 1;
        return new DialogueBlockItem(paragraphs);
    }

    private static void AddParagraph(List<Item> items, string type, string text, string? number = null)
    {
        if (!string.IsNullOrWhiteSpace(text))
        {
            items.Add(new ParagraphItem(type, text.Trim(), number));
        }
    }

    // Final Draft does not reliably honour line breaks inside <Text>, so each line is its own paragraph.
    private static void AddLines(List<Item> items, string type, string text, string? alignment = null)
    {
        foreach (var line in SplitLines(Clean(text)))
        {
            items.Add(new ParagraphItem(type, line, Alignment: alignment));
        }
    }

    private static string Clean(string text) => FountainMarkup.StripOmissions(text);

    private static IEnumerable<string> SplitLines(string text) =>
        text.Replace("\r", string.Empty).Split('\n').Select(line => line.Trim()).Where(line => line.Length > 0);

    private static void WriteTitlePage(XmlWriter writer, TitlePageData titlePage)
    {
        if (titlePage.Entries.Count == 0)
        {
            return;
        }

        writer.WriteStartElement("TitlePage");
        writer.WriteStartElement("Content");
        foreach (var entry in titlePage.Entries)
        {
            var alignment = entry.FieldType is TitlePageFieldType.Title or TitlePageFieldType.Episode
                or TitlePageFieldType.Credit or TitlePageFieldType.Author or TitlePageFieldType.Source
                ? "Center"
                : "Left";
            foreach (var line in SplitLines(Clean(entry.Value)))
            {
                WriteParagraph(writer, "General", line, null, alignment);
            }
        }
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteParagraph(XmlWriter writer, string type, string text, string? number, string? alignment)
    {
        writer.WriteStartElement("Paragraph");
        if (alignment is not null)
        {
            writer.WriteAttributeString("Alignment", alignment);
        }
        if (number is not null)
        {
            writer.WriteAttributeString("Number", number);
        }
        writer.WriteAttributeString("Type", type);

        foreach (var (runText, style) in ParseEmphasis(text))
        {
            writer.WriteStartElement("Text");
            if (style.Length > 0)
            {
                writer.WriteAttributeString("Style", style);
            }
            writer.WriteString(runText);
            writer.WriteEndElement();
        }

        writer.WriteEndElement();
    }

    private static readonly (string Marker, bool Bold, bool Italic, bool Underline)[] EmphasisMarkers =
    {
        ("***", true, true, false),
        ("**", true, false, false),
        ("*", false, true, false),
        ("_", false, false, true)
    };

    // Splits Fountain *italic*, **bold**, ***both*** and _underline_ into styled runs. A marker
    // only opens when the same marker closes it later on the line; '\' escapes the next character.
    private static List<(string Text, string Style)> ParseEmphasis(string text)
    {
        var runs = new List<(string Text, string Style)>();
        var current = new StringBuilder();
        bool bold = false, italic = false, underline = false;

        void Flush()
        {
            if (current.Length == 0) return;
            // A whitespace-only <Text> is dropped by readers that ignore insignificant whitespace,
            // so a lone space between two styled words joins the run before it.
            if (runs.Count > 0 && string.IsNullOrWhiteSpace(current.ToString()))
            {
                runs[^1] = (runs[^1].Text + current, runs[^1].Style);
                current.Clear();
                return;
            }
            var parts = new List<string>(3);
            if (bold) parts.Add("Bold");
            if (italic) parts.Add("Italic");
            if (underline) parts.Add("Underline");
            runs.Add((current.ToString(), string.Join("+", parts)));
            current.Clear();
        }

        for (var index = 0; index < text.Length; index++)
        {
            if (text[index] == '\\' && index + 1 < text.Length)
            {
                current.Append(text[++index]);
                continue;
            }

            var matched = false;
            foreach (var (marker, markerBold, markerItalic, markerUnderline) in EmphasisMarkers)
            {
                if (string.CompareOrdinal(text, index, marker, 0, marker.Length) != 0)
                {
                    continue;
                }

                var isOpen = (markerBold && markerItalic) ? bold && italic
                    : markerBold ? bold
                    : markerItalic ? italic
                    : underline;
                if (!isOpen && FindUnescaped(text, marker, index + marker.Length) < 0)
                {
                    break;
                }

                Flush();
                if (markerBold) bold = !isOpen;
                if (markerItalic) italic = !isOpen;
                if (markerUnderline) underline = !isOpen;
                index += marker.Length - 1;
                matched = true;
                break;
            }

            if (!matched)
            {
                current.Append(text[index]);
            }
        }

        Flush();
        if (runs.Count == 0)
        {
            runs.Add((string.Empty, string.Empty));
        }

        return runs;
    }

    private static int FindUnescaped(string text, string marker, int start)
    {
        for (var index = start; index <= text.Length - marker.Length; index++)
        {
            if (text[index] == '\\')
            {
                index++;
                continue;
            }

            if (string.CompareOrdinal(text, index, marker, 0, marker.Length) == 0)
            {
                return index;
            }
        }

        return -1;
    }

    private sealed class StringWriterWithEncoding(StringBuilder builder, Encoding encoding) : StringWriter(builder)
    {
        public override Encoding Encoding { get; } = encoding;
    }
}
