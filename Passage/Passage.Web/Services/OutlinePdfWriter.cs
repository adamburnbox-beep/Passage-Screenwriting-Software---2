using System.Globalization;
using System.Text;

namespace Passage.Web.Services;

/// <summary>
/// Writes the Board outline as a PDF (docs/STORY-PLAN.md S3.3): the same cards
/// and fields as <see cref="OutlineExport"/>, Act/Sequence headings in bold, a
/// storyline colour swatch beside each card and page numbers. Hand-rolled with
/// the built-in Courier fonts like the screenplay exporter, but web-only so the
/// shared <c>Passage.Export</c> library is untouched. Pure, so it is testable
/// without a browser.
/// </summary>
public static class OutlinePdfWriter
{
    public const double PageWidth = 612;
    public const double PageHeight = 792;
    public const double MarginLeft = 72;
    public const double MarginRight = 54;
    public const double MarginTop = 60;
    public const double MarginBottom = 60;
    public const double BodySize = 10;
    private const double CharRatio = 0.6; // Courier advance width per point of font size
    private const double IndentStep = 14;
    private const double SwatchSize = 7;
    private const double CardGap = 7;

    private sealed record Row(string Text, bool Bold, double Size, double X, double[][] Swatches, bool KeepWithNext, double GapAfter);

    private sealed class Page
    {
        public readonly StringBuilder Content = new();
    }

    public static byte[] Build(
        IReadOnlyList<BoardLane> lanes,
        StoryData story,
        IReadOnlyList<Beat> beats,
        int targetPages,
        OutlineExportOptions options)
    {
        var rows = new List<Row>();
        OutlineExport.Walk(lanes, story, options, (card, depth, hashes) => AddCard(rows, story, options, card, depth, hashes));
        if (options.BeatSheet)
        {
            AddBeatSheet(rows, story, beats, targetPages, options);
        }

        return Render(Paginate(rows));
    }

    /// <summary>Greedy word wrap to <paramref name="width"/> characters. Words
    /// longer than the column are split so nothing runs past the margin.</summary>
    public static List<string> Wrap(string text, int width)
    {
        width = Math.Max(1, width);
        var lines = new List<string>();
        var current = new StringBuilder();
        foreach (var word in text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            var rest = word;
            while (rest.Length > width)
            {
                if (current.Length > 0)
                {
                    lines.Add(current.ToString());
                    current.Clear();
                }

                lines.Add(rest[..width]);
                rest = rest[width..];
            }

            if (current.Length > 0 && current.Length + 1 + rest.Length > width)
            {
                lines.Add(current.ToString());
                current.Clear();
            }

            if (current.Length > 0)
            {
                current.Append(' ');
            }

            current.Append(rest);
        }

        if (current.Length > 0)
        {
            lines.Add(current.ToString());
        }

        return lines;
    }

    /// <summary>Escapes a string for a PDF literal: backslash and parentheses are
    /// backslashed, control characters become spaces, anything beyond Latin-1
    /// becomes '?'.</summary>
    public static string Escape(string text)
    {
        var builder = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            if (ch is '\\' or '(' or ')')
            {
                builder.Append('\\').Append(ch);
            }
            else if (char.IsControl(ch))
            {
                builder.Append(' ');
            }
            else
            {
                builder.Append(ch <= 'ÿ' ? ch : '?');
            }
        }

        return builder.ToString();
    }

    private static void AddCard(List<Row> rows, StoryData story, OutlineExportOptions options, BoardCard card, int depth, int hashes)
    {
        var tags = card.Id == Guid.Empty ? null : story.Items.FirstOrDefault(item => item.TargetId == card.Id);
        var x = MarginLeft + depth * IndentStep;
        var size = hashes switch { 1 => 14.0, 2 => 12.0, _ => BodySize };
        var swatches = options.Storylines && tags is not null
            ? story.Storylines.Where(line => tags.Storylines.Contains(line.Id)).Select(line => Rgb(StoryItems.StorylineColour(line.Colour))).ToArray()
            : Array.Empty<double[]>();

        var heading = card.Heading.Trim();
        var headingLines = Wrap(heading.Length == 0 ? "(untitled)" : heading, Columns(x, size));
        var fields = Fields(story, options, options.Synopses ? card.Description : string.Empty, tags, x + IndentStep);
        for (var i = 0; i < headingLines.Count; i++)
        {
            var last = i == headingLines.Count - 1 && fields.Count == 0;
            // Headings stay with the line after them; a card's last line carries the gap.
            rows.Add(new Row(headingLines[i], hashes > 0, size, x, i == 0 ? swatches : Array.Empty<double[]>(),
                !last, last ? CardGap : 0));
        }

        for (var i = 0; i < fields.Count; i++)
        {
            rows.Add(fields[i] with { GapAfter = i == fields.Count - 1 ? CardGap : 0 });
        }
    }

    private static List<Row> Fields(StoryData story, OutlineExportOptions options, string synopsis, ItemTags? tags, double x)
    {
        var rows = new List<Row>();
        void Add(string text)
        {
            foreach (var line in Wrap(text, Columns(x, BodySize)))
            {
                rows.Add(new Row(line, false, BodySize, x, Array.Empty<double[]>(), false, 0));
            }
        }

        foreach (var line in synopsis.ReplaceLineEndings("\n").Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            Add(line.Trim());
        }

        if (tags is null)
        {
            return rows;
        }

        if (options.Storylines)
        {
            var names = story.Storylines.Where(line => tags.Storylines.Contains(line.Id)).Select(line => line.Name).ToList();
            if (names.Count > 0)
            {
                Add("Storylines: " + string.Join(", ", names));
            }
        }

        if (options.Characters && tags.Characters.Count > 0)
        {
            Add("Characters: " + string.Join(", ", tags.Characters));
        }

        if (options.Locations && tags.Locations.Count > 0)
        {
            Add("Locations: " + string.Join(", ", tags.Locations));
        }

        return rows;
    }

    private static void AddBeatSheet(List<Row> rows, StoryData story, IReadOnlyList<Beat> beats, int targetPages, OutlineExportOptions options)
    {
        var chosen = options.KeyOnly ? beats.Where(beat => StoryItems.IsKey(story, beat.Id)).ToList() : beats.ToList();
        var placed = chosen.Where(beat => beat.Page > 0).OrderBy(beat => beat.Page).ToList();
        var inbox = chosen.Where(beat => beat.Page == 0).ToList();

        rows.Add(new Row("BEAT SHEET" + (targetPages > 0 ? $" ({targetPages} pages)" : string.Empty), true, 14, MarginLeft,
            Array.Empty<double[]>(), true, CardGap));
        void Beat(Beat beat, string? pages)
        {
            var title = beat.Title.Trim().Length > 0 ? beat.Title.Trim() : "Untitled beat";
            var tags = story.Items.FirstOrDefault(item => item.Kind == StoryItems.BeatKind && item.TargetId == beat.Id);
            var x = MarginLeft + IndentStep;
            var swatches = options.Storylines && tags is not null
                ? story.Storylines.Where(line => tags.Storylines.Contains(line.Id)).Select(line => Rgb(StoryItems.StorylineColour(line.Colour))).ToArray()
                : Array.Empty<double[]>();
            var fields = Fields(story, options, options.Synopses ? beat.Description : string.Empty, tags, x + IndentStep);
            var head = Wrap(pages is null ? title : pages + "  " + title, Columns(x, BodySize));
            for (var i = 0; i < head.Count; i++)
            {
                var last = i == head.Count - 1 && fields.Count == 0;
                rows.Add(new Row(head[i], true, BodySize, x, i == 0 ? swatches : Array.Empty<double[]>(), !last, last ? CardGap : 0));
            }

            for (var i = 0; i < fields.Count; i++)
            {
                rows.Add(fields[i] with { GapAfter = i == fields.Count - 1 ? CardGap : 0 });
            }
        }

        foreach (var beat in placed)
        {
            var span = Math.Max(1, beat.Span);
            Beat(beat, span == 1 ? $"p. {beat.Page}" : $"pp. {beat.Page}-{beat.Page + span - 1}");
        }

        if (inbox.Count > 0)
        {
            rows.Add(new Row("INBOX", true, BodySize, MarginLeft, Array.Empty<double[]>(), true, CardGap));
            foreach (var beat in inbox)
            {
                Beat(beat, null);
            }
        }

        if (placed.Count == 0 && inbox.Count == 0)
        {
            rows.Add(new Row("(no beats)", false, BodySize, MarginLeft + IndentStep, Array.Empty<double[]>(), false, CardGap));
        }
    }

    private static int Columns(double x, double size) =>
        (int)Math.Floor((PageWidth - MarginRight - x) / (size * CharRatio));

    private static double LineHeight(Row row) => row.Size * 1.3;

    /// <summary>Lays rows top to bottom. A row marked KeepWithNext moves to the
    /// next page with the row after it rather than stranding a heading.</summary>
    private static List<Page> Paginate(List<Row> rows)
    {
        var pages = new List<Page>();
        Page? page = null;
        var y = 0.0;
        var top = PageHeight - MarginTop;

        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var needed = LineHeight(row) + (row.KeepWithNext && i + 1 < rows.Count ? LineHeight(rows[i + 1]) : 0);
            if (page is null || y - needed < MarginBottom)
            {
                page = new Page();
                pages.Add(page);
                y = top;
            }

            y -= LineHeight(row);
            foreach (var colour in row.Swatches.Select((rgb, index) => (rgb, index)))
            {
                page.Content.Append(string.Format(CultureInfo.InvariantCulture,
                    "{0:0.###} {1:0.###} {2:0.###} rg {3:0.###} {4:0.###} {5} {5} re f\n",
                    colour.rgb[0], colour.rgb[1], colour.rgb[2],
                    row.X - 11 - colour.index * (SwatchSize + 2), y + 1, SwatchSize));
            }

            if (row.Swatches.Length > 0)
            {
                page.Content.Append("0 g\n");
            }

            page.Content.Append(string.Format(CultureInfo.InvariantCulture,
                "BT /{0} {1:0.###} Tf 1 0 0 1 {2:0.###} {3:0.###} Tm ({4}) Tj ET\n",
                row.Bold ? "F2" : "F1", row.Size, row.X, y, Escape(row.Text)));
            y -= row.GapAfter;
        }

        if (pages.Count == 0)
        {
            pages.Add(new Page());
        }

        return pages;
    }

    private static byte[] Render(List<Page> pages)
    {
        for (var i = 0; i < pages.Count; i++)
        {
            var number = (i + 1).ToString(CultureInfo.InvariantCulture);
            var x = PageWidth / 2 - number.Length * BodySize * CharRatio / 2;
            pages[i].Content.Append(string.Format(CultureInfo.InvariantCulture,
                "BT /F1 {0:0.###} Tf 1 0 0 1 {1:0.###} {2:0.###} Tm ({3}) Tj ET\n", BodySize, x, MarginBottom / 2, number));
        }

        // Objects: 1 catalog, 2 pages, 3 Courier, 4 Courier-Bold, then page/content pairs.
        var objects = new List<string>
        {
            "<< /Type /Catalog /Pages 2 0 R >>",
            $"<< /Type /Pages /Kids [{string.Join(" ", pages.Select((_, i) => $"{5 + i * 2} 0 R"))}] /Count {pages.Count} >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Courier-Bold >>"
        };
        for (var i = 0; i < pages.Count; i++)
        {
            var content = pages[i].Content.ToString();
            objects.Add(string.Format(CultureInfo.InvariantCulture,
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {0} {1}] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents {2} 0 R >>",
                PageWidth, PageHeight, 6 + i * 2));
            objects.Add($"<< /Length {Encoding.Latin1.GetByteCount(content)} >>\nstream\n{content}endstream");
        }

        using var stream = new MemoryStream();
        using var writer = new StreamWriter(stream, Encoding.Latin1, leaveOpen: true) { NewLine = "\n" };
        writer.Write("%PDF-1.4\n");
        writer.Flush();
        var offsets = new List<long>();
        for (var i = 0; i < objects.Count; i++)
        {
            offsets.Add(stream.Position);
            writer.WriteLine($"{i + 1} 0 obj");
            writer.WriteLine(objects[i]);
            writer.WriteLine("endobj");
            writer.Flush();
        }

        var xref = stream.Position;
        writer.WriteLine("xref");
        writer.WriteLine($"0 {objects.Count + 1}");
        writer.WriteLine("0000000000 65535 f ");
        foreach (var offset in offsets)
        {
            writer.WriteLine($"{offset:0000000000} 00000 n ");
        }

        writer.WriteLine("trailer");
        writer.WriteLine($"<< /Size {objects.Count + 1} /Root 1 0 R >>");
        writer.WriteLine("startxref");
        writer.WriteLine(xref.ToString(CultureInfo.InvariantCulture));
        writer.WriteLine("%%EOF");
        writer.Flush();
        return stream.ToArray();
    }

    /// <summary>A palette hex as 0–1 RGB; mid grey for a storyline with no colour.</summary>
    private static double[] Rgb(string hex)
    {
        if (hex.Length == 7 && hex[0] == '#'
            && int.TryParse(hex.AsSpan(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var value))
        {
            return new[] { ((value >> 16) & 255) / 255.0, ((value >> 8) & 255) / 255.0, (value & 255) / 255.0 };
        }

        return new[] { 0.6, 0.6, 0.6 };
    }
}
