using System.Globalization;
using System.Text;
using Passage.Core;
using Passage.Parser;

namespace Passage.Export;

public sealed class ScreenplayPdfExporter : IExporter
{
    private static readonly CultureInfo InvariantCulture = CultureInfo.InvariantCulture;

    public string DisplayName => "PDF";

    public string DefaultExtension => ".pdf";

    public void Export(ParsedScreenplay screenplay, string filePath)
    {
        ArgumentNullException.ThrowIfNull(screenplay);
        ArgumentNullException.ThrowIfNull(filePath);

        var (titlePages, bodyPages) = ScreenplayLayoutBuilder.BuildPages(screenplay);
        WritePdf(titlePages, bodyPages, filePath);
    }

    private static void WritePdf(IReadOnlyList<LayoutPage> titlePages, IReadOnlyList<LayoutPage> bodyPages, string filePath)
    {
        var contentStreams = new List<string>();
        foreach (var titlePage in titlePages)
        {
            contentStreams.Add(BuildContentStream(titlePage.Lines, null));
        }

        for (var i = 0; i < bodyPages.Count; i++)
        {
            contentStreams.Add(BuildContentStream(bodyPages[i].Lines, bodyPages[i].PageNumber));
        }

        var totalPages = titlePages.Count + bodyPages.Count;
        var objectStrings = new List<string>();

        objectStrings.Add("<< /Type /Catalog /Pages 2 0 R >>");

        var pageObjectStart = 5;
        var pageRefs = new List<string>();
        for (var i = 0; i < totalPages; i++)
        {
            var pageObjectNumber = pageObjectStart + (i * 2);
            var contentObjectNumber = pageObjectNumber + 1;
            pageRefs.Add($"{pageObjectNumber} 0 R");

            objectStrings.Add(
                $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {ScreenplayLayoutBuilder.PageWidth.ToString("0.###", InvariantCulture)} {ScreenplayLayoutBuilder.PageHeight.ToString("0.###", InvariantCulture)}] " +
                $"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents {contentObjectNumber} 0 R >>");

            objectStrings.Add(BuildStreamObject(contentStreams[i]));
        }

        objectStrings.Insert(1, $"<< /Type /Pages /Kids [{string.Join(" ", pageRefs)}] /Count {totalPages} >>");
        objectStrings.Insert(2, "<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>");
        objectStrings.Insert(3, "<< /Type /Font /Subtype /Type1 /BaseFont /Courier-Bold >>");

        using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
        using var writer = new StreamWriter(stream, Encoding.Latin1, leaveOpen: true);

        writer.Write("%PDF-1.4\n");
        writer.Flush();

        var offsets = new List<long>();
        for (var i = 0; i < objectStrings.Count; i++)
        {
            offsets.Add(stream.Position);
            writer.WriteLine($"{i + 1} 0 obj");
            writer.WriteLine(objectStrings[i]);
            writer.WriteLine("endobj");
            writer.Flush();
        }

        var xrefStart = stream.Position;
        writer.WriteLine("xref");
        writer.WriteLine($"0 {objectStrings.Count + 1}");
        writer.WriteLine("0000000000 65535 f ");
        foreach (var offset in offsets)
        {
            writer.WriteLine($"{offset:0000000000} 00000 n ");
        }

        writer.WriteLine("trailer");
        writer.WriteLine($"<< /Size {objectStrings.Count + 1} /Root 1 0 R >>");
        writer.WriteLine("startxref");
        writer.WriteLine(xrefStart.ToString(InvariantCulture));
        writer.WriteLine("%%EOF");
        writer.Flush();
    }

    private static string BuildContentStream(IReadOnlyList<LayoutLine> lines, int? pageNumber)
    {
        var builder = new StringBuilder();
        var y = ScreenplayLayoutBuilder.PageHeight - ScreenplayLayoutBuilder.MarginTop;

        if (pageNumber.HasValue && pageNumber.Value > 1)
        {
            AppendTextRun(builder, $"{pageNumber.Value}.", LayoutTextStyle.RightWithinBody, 0.0, ScreenplayLayoutBuilder.PageHeight - ScreenplayLayoutBuilder.PageNumberTopMargin);
        }

        foreach (var line in lines)
        {
            if (line.IsBlank)
            {
                y -= ScreenplayLayoutBuilder.LineHeight;
                continue;
            }

            AppendTextRun(builder, line.Text, line.Style, line.X, y);

            y -= ScreenplayLayoutBuilder.LineHeight;
        }

        var data = Encoding.Latin1.GetBytes(builder.ToString());
        return $"<< /Length {data.Length} >>\nstream\n{builder}endstream";
    }

    private static string BuildStreamObject(string contentStream)
    {
        return contentStream;
    }

    private static void AppendTextRun(StringBuilder builder, string text, LayoutTextStyle style, double x, double y)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        var sanitizedText = SanitizePdfText(text);
        var resolvedX = ComputeX(new LayoutLine(sanitizedText, style, x));
        var fontResource = GetFontResource(style);

        builder.Append("BT ");
        builder.Append('/');
        builder.Append(fontResource);
        builder.Append(' ');
        builder.Append(ScreenplayLayoutBuilder.FontSize.ToString("0.###", InvariantCulture));
        builder.Append(" Tf ");
        builder.Append("1 0 0 1 ");
        builder.Append(resolvedX.ToString("0.###", InvariantCulture));
        builder.Append(' ');
        builder.Append(y.ToString("0.###", InvariantCulture));
        builder.Append(" Tm (");
        builder.Append(sanitizedText);
        builder.AppendLine(") Tj ET");
    }

    private static double ComputeX(LayoutLine line)
    {
        var usableWidth = ScreenplayLayoutBuilder.PageWidth - ScreenplayLayoutBuilder.MarginLeft - ScreenplayLayoutBuilder.MarginRight;
        var lineWidth = Math.Min(usableWidth, Math.Max(1, line.Text.Length) * ScreenplayLayoutBuilder.CharWidth);

        return line.Style switch
        {
            LayoutTextStyle.CenterWithinBody or LayoutTextStyle.CenterWithinBodyBold => ScreenplayLayoutBuilder.MarginLeft + Math.Max(0, (usableWidth - lineWidth) / 2),
            LayoutTextStyle.RightWithinBody => ScreenplayLayoutBuilder.MarginLeft + Math.Max(0, usableWidth - lineWidth),
            _ => line.X
        };
    }

    private static string GetFontResource(LayoutTextStyle style)
    {
        return style == LayoutTextStyle.LeftBold || style == LayoutTextStyle.CenterWithinBodyBold ? "F2" : "F1";
    }

    private static string SanitizePdfText(string text)
    {
        var builder = new StringBuilder(text.Length);

        foreach (var ch in text)
        {
            if (ch is '\\' or '(' or ')')
            {
                builder.Append('\\').Append(ch);
                continue;
            }

            builder.Append(ch <= '\u00FF' ? ch : '?');
        }

        return builder.ToString();
    }
}
