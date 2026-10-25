using System.Text;
using System.Text.RegularExpressions;

namespace Passage.Web.Services;

public enum OutlineFormat
{
    Fountain,
    Text
}

/// <summary>Which fields "Export outline…" writes (docs/STORY-PLAN.md S3.2).
/// Headings are always written.</summary>
public sealed record OutlineExportOptions(
    bool Synopses = true,
    bool Storylines = true,
    bool Characters = true,
    bool Locations = true,
    bool KeyOnly = false,
    bool BeatSheet = false);

/// <summary>
/// Writes the Board as an outline: Fountain that re-opens in Passage as the
/// same lanes and cards (tags ride along as [[notes]]), or an indented plain
/// text outline. Web-only, so it is deliberately not an <c>IExporter</c> in
/// the shared catalog that feeds the desktop menus. Pure, so the output is
/// testable without a browser.
/// </summary>
public static class OutlineExport
{
    private static readonly Regex StandardHeading = new(
        @"^(INT|EXT|EST|INT\./EXT|INT/EXT|I/E)[\. /]", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static string Extension(OutlineFormat format) => format == OutlineFormat.Fountain ? ".fountain" : ".txt";

    public static string Build(
        IReadOnlyList<BoardLane> lanes,
        StoryData story,
        IReadOnlyList<Beat> beats,
        int targetPages,
        OutlineExportOptions options,
        OutlineFormat format)
    {
        var writer = new Writer(format == OutlineFormat.Fountain, story, options);

        foreach (var lane in lanes)
        {
            var laneKept = !options.KeyOnly || IsKey(story, lane.ActCard)
                || lane.Groups.Any(group => IsKey(story, group.SequenceCard) || group.Cards.Any(card => IsKey(story, card)));
            if (!laneKept)
            {
                continue;
            }

            var depth = 0;
            if (lane.ActCard is not null)
            {
                writer.Card(lane.ActCard, depth, 1);
                depth++;
            }

            foreach (var group in lane.Groups)
            {
                var cards = options.KeyOnly ? group.Cards.Where(card => IsKey(story, card)).ToList() : group.Cards;
                if (options.KeyOnly && cards.Count == 0 && !IsKey(story, group.SequenceCard))
                {
                    continue;
                }

                var cardDepth = depth;
                if (group.SequenceCard is not null)
                {
                    writer.Card(group.SequenceCard, depth, 2);
                    cardDepth++;
                }

                foreach (var card in cards)
                {
                    writer.Card(card, cardDepth, card.Kind == "Scene" ? 0 : 3);
                }
            }
        }

        if (options.BeatSheet)
        {
            writer.BeatSheet(beats, targetPages);
        }

        return writer.ToString();
    }

    private static bool IsKey(StoryData story, BoardCard? card) =>
        card is not null && card.Id != Guid.Empty && StoryItems.IsKey(story, card.Id);

    private sealed class Writer
    {
        private readonly bool _fountain;
        private readonly StoryData _story;
        private readonly OutlineExportOptions _options;
        private readonly StringBuilder _out = new();

        public Writer(bool fountain, StoryData story, OutlineExportOptions options)
        {
            _fountain = fountain;
            _story = story;
            _options = options;
        }

        public override string ToString() => _out.ToString();

        /// <summary>One heading and its fields. Hashes is the Fountain section
        /// depth (1–3), or 0 for a scene heading.</summary>
        public void Card(BoardCard card, int depth, int hashes)
        {
            var tags = card.Id == Guid.Empty ? null : _story.Items.FirstOrDefault(item => item.TargetId == card.Id);
            var heading = card.Heading.Trim();
            if (_fountain)
            {
                _out.Append(hashes > 0
                    ? new string('#', hashes) + " " + heading
                    : StandardHeading.IsMatch(heading) ? heading : "." + heading).Append('\n');
            }
            else
            {
                _out.Append(Indent(depth)).Append(heading).Append('\n');
            }

            Fields(depth + 1, _options.Synopses ? card.Description : string.Empty, tags);
            _out.Append('\n');
        }

        public void BeatSheet(IReadOnlyList<Beat> beats, int targetPages)
        {
            var chosen = _options.KeyOnly
                ? beats.Where(beat => StoryItems.IsKey(_story, beat.Id)).ToList()
                : beats.ToList();
            var placed = chosen.Where(beat => beat.Page > 0).OrderBy(beat => beat.Page).ToList();
            var inbox = chosen.Where(beat => beat.Page == 0).ToList();

            // A boneyard: the parser ignores it, so the outline's lanes are unchanged on re-open.
            _out.Append(_fountain ? "/*\n" : string.Empty);
            _out.Append("BEAT SHEET").Append(targetPages > 0 ? $" ({targetPages} pages)" : string.Empty).Append('\n');
            foreach (var beat in placed)
            {
                var span = Math.Max(1, beat.Span);
                var pages = span == 1 ? $"p. {beat.Page}" : $"pp. {beat.Page}-{beat.Page + span - 1}";
                Beat(beat, pages);
            }

            if (inbox.Count > 0)
            {
                _out.Append("\nINBOX\n");
                foreach (var beat in inbox)
                {
                    Beat(beat, null);
                }
            }

            if (placed.Count == 0 && inbox.Count == 0)
            {
                _out.Append("(no beats)\n");
            }

            _out.Append(_fountain ? "*/\n" : string.Empty);
        }

        private void Beat(Beat beat, string? pages)
        {
            var title = beat.Title.Trim().Length > 0 ? beat.Title.Trim() : "Untitled beat";
            _out.Append(Indent(1)).Append(pages is null ? title : pages + "  " + title).Append('\n');
            var tags = _story.Items.FirstOrDefault(item => item.Kind == StoryItems.BeatKind && item.TargetId == beat.Id);
            Fields(2, _options.Synopses ? beat.Description : string.Empty, tags, boneyard: true);
        }

        private void Fields(int depth, string synopsis, ItemTags? tags, bool boneyard = false)
        {
            foreach (var line in synopsis.ReplaceLineEndings("\n").Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                var text = line.Trim();
                if (text.Length == 0)
                {
                    continue;
                }

                // Inside the beat sheet's boneyard "*/" would end the comment early.
                text = boneyard ? text.Replace("*/", "* /") : text;
                _out.Append(_fountain && !boneyard ? "= " + text : Indent(depth) + text).Append('\n');
            }

            if (tags is null)
            {
                return;
            }

            if (_options.Storylines)
            {
                var names = _story.Storylines.Where(line => tags.Storylines.Contains(line.Id)).Select(line => line.Name);
                Tag("Storylines", names, depth, boneyard);
            }

            if (_options.Characters)
            {
                Tag("Characters", tags.Characters, depth, boneyard);
            }

            if (_options.Locations)
            {
                Tag("Locations", tags.Locations, depth, boneyard);
            }
        }

        private void Tag(string label, IEnumerable<string> values, int depth, bool boneyard)
        {
            var list = values.Select(value => value.Replace("]]", "] ]").Replace("*/", "* /")).ToList();
            if (list.Count == 0)
            {
                return;
            }

            var text = label + ": " + string.Join(", ", list);
            _out.Append(_fountain && !boneyard ? "[[" + text + "]]" : Indent(depth) + text).Append('\n');
        }

        private static string Indent(int depth) => new(' ', depth * 2);
    }
}
