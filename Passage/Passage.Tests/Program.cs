using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Passage.Parser;
using Passage.Core;
using Passage.Web.Services;

namespace Passage.Tests;

class Program
{
    static int Main(string[] args)
    {
        Console.WriteLine("=== Running Passage Screenwriting Software Unit Tests ===");
        var failures = 0;

        failures += RunTest("Test Parse Simple Scene Heading", TestParseSimpleSceneHeading);
        failures += RunTest("Test Parse Character and Dialogue", TestParseCharacterAndDialogue);
        failures += RunTest("Test Parse Parenthetical", TestParseParenthetical);
        failures += RunTest("Test Parse Explicit Line Overrides", TestParseExplicitLineOverrides);
        failures += RunTest("Test TextAnalysis Helper Methods", TestTextAnalysisHelperMethods);
        failures += RunTest("Test BeatBoard Card Own Line Range", TestBeatBoardCardOwnLineRange);
        failures += RunTest("Test BeatBoard Nested Section Block", TestBeatBoardNestedSectionBlock);
        failures += RunTest("Test BeatBoard Scene Block Extent", TestBeatBoardSceneBlockExtent);
        failures += RunTest("Test BeatBoard Range Out Of Range Input", TestBeatBoardRangeOutOfRangeInput);
        failures += RunTest("Test BeatBoard Build And Splice Card Lines", TestBeatBoardBuildAndSpliceCardLines);
        failures += RunTest("Test BeatBoard Plan Move", TestBeatBoardPlanMove);
        failures += RunTest("Test BeatBoard Plan Move Rejections", TestBeatBoardPlanMoveRejections);
        failures += RunTest("Test BracketScanner Finds And Ranks", TestBracketScannerFindsAndRanks);
        failures += RunTest("Test BracketScanner Ignores Notes And Boneyard", TestBracketScannerIgnoresNotesAndBoneyard);
        failures += RunTest("Test BracketScanner Malformed Input", TestBracketScannerMalformedInput);
        failures += RunTest("Test SlateStore Round Trip", TestSlateStoreRoundTrip);
        failures += RunTest("Test SlateStore Rejects Invalid Names", TestSlateStoreRejectsInvalidNames);
        failures += RunTest("Test SlateStore Prunes Orphans", TestSlateStorePrunesOrphans);
        failures += RunTest("Test SlateStore Corrupt Sidecar Fails Visibly", TestSlateStoreCorruptSidecarFailsVisibly);
        failures += RunTest("Test SlateStore Chain Round Trip", TestSlateStoreChainRoundTrip);
        failures += RunTest("Test Synopsis Placement", TestSynopsisPlacement);
        failures += RunTest("Test SplitScript Lines And Parse", TestSplitScriptLinesAndParse);
        failures += RunTest("Test ShapeLine Derives From Lanes", TestShapeLineDerivesFromLanes);
        failures += RunTest("Test SlateStore Split Family Round Trip", TestSlateStoreSplitFamilyRoundTrip);
        failures += RunTest("Test SlateStore Bridge And Position Round Trip", TestSlateStoreBridgeAndPositionRoundTrip);

        Console.WriteLine("\n=== Test Run Completed ===");
        if (failures == 0)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("ALL TESTS PASSED SUCCESSFULLY!");
            Console.ResetColor();
            return 0;
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"{failures} TEST(S) FAILED.");
            Console.ResetColor();
            return 1;
        }
    }

    static int RunTest(string testName, Action testAction)
    {
        Console.Write($"Running: {testName}... ");
        try
        {
            testAction();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("PASSED");
            Console.ResetColor();
            return 0;
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("FAILED");
            Console.WriteLine(ex.ToString());
            Console.ResetColor();
            return 1;
        }
    }


    // ---- BeatBoardText (Passage.Parser) ----
    //
    // These cover the line-range and splicing logic that used to live in the
    // Avalonia view model, where nothing could reach it.

    const string BoardScript =
        "# ACT ONE\n" +          // 0
        "= Act synopsis.\n" +    // 1
        "\n" +                   // 2
        "## Setup\n" +           // 3
        "\n" +                   // 4
        "INT. KITCHEN - DAY\n" + // 5
        "= Scene synopsis.\n" +  // 6
        "\n" +                   // 7
        "She burns the toast.\n" + // 8
        "\n" +                   // 9
        "EXT. GARDEN - LATER\n" + // 10
        "\n" +                   // 11
        "# ACT TWO\n" +          // 12
        "\n" +                   // 13
        "INT. HALLWAY - NIGHT";   // 14

    static (ParsedScreenplay parsed, int lineCount) ParseBoardScript()
    {
        var parsed = new FountainParser().Parse(BoardScript);
        return (parsed, BoardScript.Split('\n').Length);
    }

    static ScreenplayElement FindElement(ParsedScreenplay parsed, int lineIndex)
    {
        foreach (var element in parsed.Elements)
        {
            if (element.LineIndex == lineIndex) return element;
        }

        throw new Exception($"No element starts at line index {lineIndex}");
    }

    static void TestBeatBoardCardOwnLineRange()
    {
        var (parsed, lineCount) = ParseBoardScript();
        var scene = FindElement(parsed, 5);

        // A freshly parsed synopsis is not suppressed, so it is still a card of
        // its own and the scene does not swallow it.
        var beforeClaim = BeatBoardText.GetCardLineRange(parsed.Elements, scene.Id, lineCount, includeNestedBlock: false);
        Assert(beforeClaim.EndLineIndex == 5,
            $"An unclaimed synopsis should stay outside the range, got end {beforeClaim.EndLineIndex}");

        // The board marks a synopsis suppressed once a card shows it as its
        // description; from then on the card owns that line too.
        FindElement(parsed, 6).IsSuppressed = true;

        var own = BeatBoardText.GetCardLineRange(parsed.Elements, scene.Id, lineCount, includeNestedBlock: false);
        Assert(own.IsFound, "Scene own range should be found");
        Assert(own.StartLineIndex == 5, $"Expected own start 5, got {own.StartLineIndex}");
        Assert(own.EndLineIndex == 6, $"Expected own end 6 (claimed synopsis), got {own.EndLineIndex}");
        Assert(own.LineCount == 2, $"Expected 2 lines, got {own.LineCount}");
    }

    static void TestBeatBoardNestedSectionBlock()
    {
        var (parsed, lineCount) = ParseBoardScript();

        // ACT ONE owns everything up to the line before ACT TWO.
        var actOne = FindElement(parsed, 0);
        var actRange = BeatBoardText.GetCardLineRange(parsed.Elements, actOne.Id, lineCount, includeNestedBlock: true);
        Assert(actRange.StartLineIndex == 0, $"Expected act start 0, got {actRange.StartLineIndex}");
        Assert(actRange.EndLineIndex == 11, $"Expected act block to end at 11, got {actRange.EndLineIndex}");

        // A deeper section stops at the next section of the same or higher level,
        // so "## Setup" also ends just before ACT TWO.
        var setup = FindElement(parsed, 3);
        var setupRange = BeatBoardText.GetCardLineRange(parsed.Elements, setup.Id, lineCount, includeNestedBlock: true);
        Assert(setupRange.StartLineIndex == 3, $"Expected sequence start 3, got {setupRange.StartLineIndex}");
        Assert(setupRange.EndLineIndex == 11, $"Expected sequence block to end at 11, got {setupRange.EndLineIndex}");

        // The last act runs to the end of the document.
        var actTwo = FindElement(parsed, 12);
        var lastRange = BeatBoardText.GetCardLineRange(parsed.Elements, actTwo.Id, lineCount, includeNestedBlock: true);
        Assert(lastRange.EndLineIndex == lineCount - 1,
            $"Expected trailing act to end at {lineCount - 1}, got {lastRange.EndLineIndex}");
    }

    static void TestBeatBoardSceneBlockExtent()
    {
        var (parsed, lineCount) = ParseBoardScript();
        var scene = FindElement(parsed, 5);

        // With the block included, a scene runs to just before the next scene
        // heading or section — here the blank line before EXT. GARDEN.
        var block = BeatBoardText.GetCardLineRange(parsed.Elements, scene.Id, lineCount, includeNestedBlock: true);
        Assert(block.StartLineIndex == 5, $"Expected block start 5, got {block.StartLineIndex}");
        Assert(block.EndLineIndex == 9, $"Expected block end 9, got {block.EndLineIndex}");
    }

    static void TestBeatBoardRangeOutOfRangeInput()
    {
        var (parsed, lineCount) = ParseBoardScript();

        var missing = BeatBoardText.GetCardLineRange(parsed.Elements, Guid.NewGuid(), lineCount, includeNestedBlock: true);
        Assert(!missing.IsFound, "An unknown card id should not resolve to a range");
        Assert(missing.StartLineIndex == -1 && missing.EndLineIndex == -1, "Missing range should be (-1, -1)");
        Assert(missing.LineCount == 0, "A missing range should count zero lines");

        var noElements = BeatBoardText.GetCardLineRange(new List<ScreenplayElement>(), Guid.NewGuid(), 0, includeNestedBlock: true);
        Assert(!noElements.IsFound, "An empty element list should not resolve to a range");

        // Splicing outside the document leaves the text alone rather than throwing.
        const string text = "one\ntwo\nthree";
        Assert(BeatBoardText.ReplaceLines(text, -1, 0, new[] { "x" }) == text, "Negative start should be a no-op");
        Assert(BeatBoardText.ReplaceLines(text, 99, 100, new[] { "x" }) == text, "Start past the end should be a no-op");

        // An end index past the document clamps instead of overrunning.
        Assert(BeatBoardText.ReplaceLines(text, 1, 99, new[] { "X" }) == "one\nX",
            "An overrunning end index should clamp to the last line");
    }

    static void TestBeatBoardBuildAndSpliceCardLines()
    {
        var id = Guid.NewGuid();

        var act = BeatBoardText.BuildCardLines("Act", "  ACT ONE  ", "First half.", id);
        Assert(act.Count == 2, $"Expected heading plus synopsis, got {act.Count}");
        Assert(act[0] == $"# ACT ONE [[id:{id}]]", $"Unexpected act heading: '{act[0]}'");
        Assert(act[1] == "= First half.", $"Unexpected synopsis line: '{act[1]}'");

        Assert(BeatBoardText.BuildCardLines("Sequence", "Setup", "", id)[0] == $"## Setup [[id:{id}]]",
            "Sequence should use two hashes");

        // A heading that already reads as a scene is left alone; anything else
        // gets the forcing dot.
        Assert(BeatBoardText.BuildCardLines("Scene", "INT. KITCHEN - DAY", "", id)[0] == $"INT. KITCHEN - DAY [[id:{id}]]",
            "A real scene heading should not be forced");
        Assert(BeatBoardText.BuildCardLines("Scene", "Somewhere else", "", id)[0] == $". Somewhere else [[id:{id}]]",
            "A non-scene heading should be forced with a dot");
        Assert(BeatBoardText.BuildCardLines("Note", "Remember this", "", id)[0] == $"[[Remember this id:{id}]]",
            "Note should use double-bracket syntax");

        // Blank description lines are dropped, not written as empty synopses.
        var multi = BeatBoardText.BuildCardLines("Act", "ACT", "one\n\n  two  ", id);
        Assert(multi.Count == 3, $"Expected heading plus two synopsis lines, got {multi.Count}");
        Assert(multi[1] == "= one" && multi[2] == "= two", "Description lines should be trimmed and prefixed");

        // And the splice puts them back in place of the old range.
        var spliced = BeatBoardText.ReplaceLines("a\nb\nc\nd", 1, 2, new[] { "B" });
        Assert(spliced == "a\nB\nd", $"Unexpected splice result: '{spliced}'");
    }


    static void TestBeatBoardPlanMove()
    {
        var lines = new List<string> { "A", "B", "C", "D", "E" };

        // Move A (0) to after C (2): expect B, C, A in the affected span.
        Assert(BeatBoardText.TryPlanMove(lines,
            new BeatBoardText.LineRange(0, 0), new BeatBoardText.LineRange(2, 2),
            insertAfter: true, out var splice, out var replacement),
            "Moving forwards should be planned");
        Assert(splice.StartLineIndex == 0 && splice.EndLineIndex == 2,
            $"Expected splice 0..2, got {splice.StartLineIndex}..{splice.EndLineIndex}");
        Assert(string.Join(",", replacement) == "B,C,A", $"Unexpected replacement: {string.Join(",", replacement)}");

        // Move E (4) to before B (1).
        Assert(BeatBoardText.TryPlanMove(lines,
            new BeatBoardText.LineRange(4, 4), new BeatBoardText.LineRange(1, 1),
            insertAfter: false, out splice, out replacement),
            "Moving backwards should be planned");
        Assert(splice.StartLineIndex == 1 && splice.EndLineIndex == 4,
            $"Expected splice 1..4, got {splice.StartLineIndex}..{splice.EndLineIndex}");
        Assert(string.Join(",", replacement) == "E,B,C,D", $"Unexpected replacement: {string.Join(",", replacement)}");

        // A multi-line block moves as a unit.
        var block = new List<string> { "h1", "b1", "b2", "X", "Y" };
        Assert(BeatBoardText.TryPlanMove(block,
            new BeatBoardText.LineRange(0, 2), new BeatBoardText.LineRange(4, 4),
            insertAfter: true, out splice, out replacement),
            "A block move should be planned");
        Assert(string.Join(",", replacement) == "X,Y,h1,b1,b2",
            $"Unexpected block move: {string.Join(",", replacement)}");

        // Applying the splice reproduces the whole document.
        var applied = BeatBoardText.ReplaceLines(string.Join("\n", block),
            splice.StartLineIndex, splice.EndLineIndex, replacement);
        Assert(applied == "X,Y,h1,b1,b2".Replace(",", "\n"), $"Unexpected applied text: '{applied}'");
    }

    static void TestBeatBoardPlanMoveRejections()
    {
        var lines = new List<string> { "A", "B", "C", "D" };

        Assert(!BeatBoardText.TryPlanMove(lines,
            new BeatBoardText.LineRange(1, 1), new BeatBoardText.LineRange(1, 1),
            insertAfter: true, out _, out _),
            "Dropping a card on itself should be refused");

        // An act (0..2) dropped onto a card nested inside it (1..1).
        Assert(!BeatBoardText.TryPlanMove(lines,
            new BeatBoardText.LineRange(0, 2), new BeatBoardText.LineRange(1, 1),
            insertAfter: true, out _, out _),
            "Dropping a block onto its own nested card should be refused");

        Assert(!BeatBoardText.TryPlanMove(lines,
            BeatBoardText.LineRange.NotFound, new BeatBoardText.LineRange(1, 1),
            insertAfter: true, out _, out _),
            "A missing source range should be refused");

        Assert(!BeatBoardText.TryPlanMove(lines,
            new BeatBoardText.LineRange(0, 0), BeatBoardText.LineRange.NotFound,
            insertAfter: true, out _, out _),
            "A missing target range should be refused");

        Assert(!BeatBoardText.TryPlanMove(lines,
            new BeatBoardText.LineRange(2, 99), new BeatBoardText.LineRange(0, 0),
            insertAfter: false, out _, out _),
            "A source range running past the document should be refused");
    }

    // ---- BracketScanner (Passage.Parser) ----

    static void TestBracketScannerFindsAndRanks()
    {
        const string script =
            "INT. KITCHEN - DAY\n" +                       // 0
            "She reads the letter [sic] twice.\n" +        // 1
            "\n" +                                         // 2
            "[SOMETHING forces her to stay] She sits.\n" + // 3
            "= [todo: name the neighbour] arrives\n" +     // 4
            "He says [nothing] and [NOTHING].";             // 5

        var brackets = BracketScanner.Scan(script);

        Assert(brackets.Count == 5, $"Five brackets found, got {brackets.Count}");

        // Rank 0 first, in document order; then rank 1 in document order.
        Assert(brackets[0].Text == "[SOMETHING forces her to stay]" && brackets[0].LineIndex == 3 && brackets[0].Column == 0,
            "Keyword bracket sorts first with its position");
        Assert(brackets[1].Text == "[todo: name the neighbour]" && brackets[1].LineNumber == 5,
            "Keyword match is case-insensitive and stops at punctuation");
        Assert(brackets[2].Text == "[NOTHING]" && brackets[2].Rank == 0, "All-caps content ranks 0");
        Assert(brackets[3].Text == "[sic]" && brackets[3].Rank == 1, "Prose bracket is listed, ranked 1");
        Assert(brackets[4].Text == "[nothing]" && brackets[4].Column == 8, "Second prose bracket keeps its column");
        Assert(brackets[1].Content == "todo: name the neighbour", "Content is the text inside the brackets");
    }

    static void TestBracketScannerIgnoresNotesAndBoneyard()
    {
        const string script =
            "[[a note with [brackets] in it]]\n" +
            "/* boneyard [HIDDEN]\n" +
            "still boneyard [HIDDEN] */ [VISIBLE]\n" +
            "Action [[note]] then [KEPT].";

        var brackets = BracketScanner.Scan(script);

        Assert(brackets.Count == 2, $"Only the two brackets outside omissions are found, got {brackets.Count}");
        Assert(brackets[0].Text == "[VISIBLE]" && brackets[0].LineIndex == 2 && brackets[0].Column == 27,
            "A bracket after a multi-line boneyard keeps its real line and column");
        Assert(brackets[1].Text == "[KEPT]" && brackets[1].LineIndex == 3 && brackets[1].Column == 21,
            "Masking a note on the same line does not shift later columns");
    }

    static void TestBracketScannerMalformedInput()
    {
        Assert(BracketScanner.Scan(null).Count == 0, "Null text scans to nothing");
        Assert(BracketScanner.Scan("").Count == 0, "Empty text scans to nothing");
        Assert(BracketScanner.Scan("No brackets here.").Count == 0, "Plain prose scans to nothing");
        Assert(BracketScanner.Scan("Unclosed [bracket runs\nonto the next line]").Count == 0,
            "A bracket does not span lines");
        Assert(BracketScanner.Scan("Empty [] and blank [   ] spans").Count == 0, "Empty brackets are not placeholders");

        var nested = BracketScanner.Scan("Outer [a [INNER] c] end");
        Assert(nested.Count == 1 && nested[0].Text == "[INNER]", "Nested brackets yield the innermost span only");

        var spaced = BracketScanner.Scan("[ SOMETHING padded ]");
        Assert(spaced.Count == 1 && spaced[0].Text == "[ SOMETHING padded ]" && spaced[0].Rank == 0,
            "Text keeps the padding as written so a replacement can match it; ranking uses the trimmed content");
    }

    // ---- SlateStore (Passage.Web) ----
    //
    // The sidecar store is pure file logic with no circuit behind it, so it is
    // exercised here against a throwaway library root.

    static (ScriptLibrary library, SlateStore store, string root) NewSlateFixture()
    {
        var root = Path.Combine(Path.GetTempPath(), "passage-tests-" + Guid.NewGuid().ToString("N"));
        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?> { ["Passage:DataDir"] = root })
            .Build();
        var library = new ScriptLibrary(configuration);
        return (library, new SlateStore(library), root);
    }

    static void TestSlateStoreRoundTrip()
    {
        var (library, store, root) = NewSlateFixture();
        try
        {
            Assert(store.Load("draft") is null, "A script with no sidecar loads as null");
            Assert(!store.Exists("draft"), "Exists is false before the first save");

            var document = new SlateDocument { SlateVersion = 0 };
            document.IgnoredBrackets.Add("[sic]");
            document.Fills.Add(new FillRun
            {
                Bracket = "[SOMETHING forces her to stay]",
                Options = { [0] = new FillOption { Text = "the storm", WhyWrong = "weather is never a reason" } },
                PointsTo = "she wants to be made to stay",
                Answer = string.Empty
            });
            store.Save("draft", document);

            var path = Path.Combine(root, ".slate", "draft.fountain.json");
            Assert(File.Exists(path), "Sidecar lands in .slate under the validated script name");
            Assert(store.Exists("draft.fountain"), "The bare and extended names resolve to the same sidecar");

            var loaded = store.Load("draft");
            Assert(loaded is not null, "Sidecar loads back");
            Assert(loaded!.SlateVersion == SlateStore.CurrentVersion, "Save stamps the current schema version");
            Assert(loaded.IgnoredBrackets.SequenceEqual(new[] { "[sic]" }), "Ignore list round-trips");
            Assert(loaded.Fills.Count == 1 && loaded.Fills[0].Bracket == "[SOMETHING forces her to stay]", "Fill run round-trips by bracket text");
            Assert(loaded.Fills[0].Options.Count == 3 && loaded.Fills[0].Options[0].WhyWrong == "weather is never a reason"
                && loaded.Fills[0].Options[2].Text == string.Empty, "All three option rows round-trip, empty ones included");
            Assert(loaded.Fills[0].PointsTo == "she wants to be made to stay", "Free-text fields round-trip");

            Assert(library.List().Count == 0, "The .slate directory never appears in the script list");

            store.Delete("draft");
            Assert(!store.Exists("draft"), "Delete removes the sidecar");
            store.Delete("draft"); // deleting twice is not an error
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void TestSlateStoreRejectsInvalidNames()
    {
        var (_, store, root) = NewSlateFixture();
        try
        {
            foreach (var bad in new[] { "", "   ", "../escape", "sub/dir", ".hidden", "bad\0name" })
            {
                var threw = false;
                try
                {
                    store.Save(bad, new SlateDocument());
                }
                catch (ArgumentException)
                {
                    threw = true;
                }

                Assert(threw, $"Save rejects '{bad}'");
                Assert(!store.Exists(bad), $"Exists is false for '{bad}' rather than throwing");
            }

            Assert(!Directory.Exists(Path.Combine(root, ".slate")), "Nothing was written for a rejected name");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void TestSlateStorePrunesOrphans()
    {
        var (library, store, root) = NewSlateFixture();
        try
        {
            Assert(store.PruneOrphans(new[] { "a.fountain" }).Count == 0, "Pruning with no .slate directory is a no-op");

            library.Save("kept", "INT. ROOM - DAY");
            store.Save("kept", new SlateDocument());
            store.Save("gone", new SlateDocument());
            store.Save("also-gone.md", new SlateDocument());

            var pruned = store.PruneOrphans(library.List().Select(entry => entry.Name));

            Assert(pruned.Count == 2, $"Two orphans reported, got {pruned.Count}");
            Assert(pruned.Contains("gone.fountain") && pruned.Contains("also-gone.md"), "Pruned names are the script names, not file paths");
            Assert(store.Exists("kept"), "A sidecar whose script exists survives");
            Assert(!store.Exists("gone") && !store.Exists("also-gone.md"), "Orphaned sidecars are deleted");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void TestSlateStoreCorruptSidecarFailsVisibly()
    {
        var (_, store, root) = NewSlateFixture();
        try
        {
            var dir = Path.Combine(root, ".slate");
            Directory.CreateDirectory(dir);
            File.WriteAllText(Path.Combine(dir, "broken.fountain.json"), "{ not json");

            var threw = false;
            try
            {
                store.Load("broken");
            }
            catch (JsonException)
            {
                threw = true;
            }

            Assert(threw, "A corrupt sidecar throws rather than being read as empty");
            Assert(store.Exists("broken"), "The corrupt file is left in place for the user to recover");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void TestSlateStoreChainRoundTrip()
    {
        var (_, store, root) = NewSlateFixture();
        try
        {
            var document = new SlateDocument();
            var woac = new ChainRun { Path = ChainPath.Woac, Seed = "MARA wants the key" };
            woac.Rounds[0].Want = "the key";
            woac.Rounds[0].Consequence = "the door is open but the dog is loose";
            woac.Rounds.Add(new ChainRound { Obstacle = "the dog" });
            document.Chains.Add(woac);
            var flaw = new ChainRun { Path = ChainPath.Flaw, Seed = "MARA — never asks for help" };
            flaw.Answers[2] = "she carries it alone and drops it";
            document.Chains.Add(flaw);
            document.Burst.Enabled = true;
            document.Burst.Seconds = 10;
            store.Save("draft", document);

            var json = File.ReadAllText(Path.Combine(root, ".slate", "draft.fountain.json"));
            Assert(json.Contains("\"Woac\"") && json.Contains("\"Flaw\""), "The path is stored by name, not by enum number");

            var loaded = store.Load("draft")!;
            Assert(loaded.Chains.Count == 2, "Both chains round-trip");
            Assert(loaded.Chains[0].Path == ChainPath.Woac && loaded.Chains[0].Rounds.Count == 2
                && loaded.Chains[0].Rounds[1].Obstacle == "the dog", "WOAC rounds round-trip in order");
            Assert(loaded.Chains[0].Rounds[1].Want == string.Empty, "A later round stores no Want: it is read from the consequence above");
            Assert(loaded.Chains[1].Path == ChainPath.Flaw && loaded.Chains[1].Answers.Count == ChainRun.FlawQuestionCount
                && loaded.Chains[1].Answers[2] == "she carries it alone and drops it", "All eight flaw answers round-trip, empty ones included");
            Assert(loaded.Burst.Enabled && loaded.Burst.Seconds == 10, "Burst settings round-trip");

            Assert(new ChainRun().IsEmpty, "A fresh chain is empty");
            Assert(!new ChainRun { ReadBack = "x" }.IsEmpty, "A read-back line alone makes a chain worth keeping");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void TestSynopsisPlacement()
    {
        // 0 "# Act 1", 1 "= old synopsis", 2 "", 3 "INT. KITCHEN", 4 "Action", 5 "## Seq", 6 "Action"
        var classes = new[] { "sx-section", "sx-synopsis", "", "sx-scene", "", "sx-section", "" };

        Assert(SynopsisPlacement.Find(classes, 4) == (4, 3), "Under the scene heading above the caret");
        Assert(SynopsisPlacement.Find(classes, 2) == (2, 0), "After the synopsis lines already under the section");
        Assert(SynopsisPlacement.Find(classes, 0) == (2, 0), "The caret on the heading itself counts as under it");
        Assert(SynopsisPlacement.Find(classes, 6) == (6, 5), "The nearest heading wins, not the first");
        Assert(SynopsisPlacement.Find(classes, 99) == (6, 5), "A caret past the end clamps to the last line");

        var noHeading = new[] { "", "sx-character", "sx-dialogue" };
        Assert(SynopsisPlacement.Find(noHeading, 2) == (2, -1), "No heading above: insert at the caret and say so");
        Assert(SynopsisPlacement.Find(Array.Empty<string>(), 0) == (0, -1), "An empty document inserts at line 0");
    }

    static void TestSplitScriptLinesAndParse()
    {
        var run = new SplitRun { A = "She arrives at the lighthouse", Z = "She leaves it burning", Midpoint = "The keeper\nconfesses" };
        var lines = SplitScript.BuildLines(run);
        var text = string.Join("\n", lines);

        Assert(lines.Count(line => line.StartsWith("# ")) == 4, "Four acts");
        Assert(lines.Count(line => line.StartsWith("## ")) == 8, "Eight sequences");
        Assert(lines.Count(line => line.StartsWith("= ")) == 9, "Nine slot lines: the two ends and seven turns");
        Assert(text.Contains("= Starts: She arrives at the lighthouse"), "A goes on sequence 1");
        Assert(text.Contains("= Midpoint: The keeper confesses"), "A multi-line value is written as one line");
        Assert(text.Contains("= Plot point 1: [PLOT POINT 1]"), "An unknown turn is written as a bracket for Fill");
        Assert(lines.IndexOf("= Midpoint: The keeper confesses") > lines.FindIndex(line => line.StartsWith("## Sequence 4")), "The midpoint sits on sequence 4");
        Assert(lines.IndexOf("= Ends: She leaves it burning") > lines.FindIndex(line => line.StartsWith("## Sequence 8")), "Z sits on sequence 8");

        // The Beat Board hands descriptions back without the "= ".
        Assert(SplitScript.TryParse("Midpoint: The keeper confesses", out var label, out var value)
            && label == "Midpoint" && value == "The keeper confesses", "A description line parses back to its slot");
        Assert(SplitScript.TryParse("=   pinch 1 :  [PINCH 1]", out label, out value)
            && label == "Pinch 1" && value == "[PINCH 1]", "Case and spacing do not matter");
        Assert(!SplitScript.TryParse("= Some other synopsis: with a colon", out _, out _), "Only the nine labels parse");
        Assert(!SplitScript.IsKnown("[PINCH 1]") && !SplitScript.IsKnown("") && SplitScript.IsKnown("the storm"), "Known means real text");
        Assert(SplitScript.IsDropped("–") && !SplitScript.IsKnown("–"), "A dash drops the turn");

        var scriptLines = text.Split('\n');
        var midpointLine = SplitScript.FindLine(scriptLines, SplitScript.Slots[4]);
        Assert(midpointLine >= 0 && scriptLines[midpointLine].StartsWith("= Midpoint:"), "FindLine locates a slot in the script");
        Assert(SplitScript.FindLine(new[] { "INT. HOUSE", "Action." }, SplitScript.Slots[4]) == -1, "FindLine is -1 when absent");
    }

    static void TestShapeLineDerivesFromLanes()
    {
        Assert(ShapeLine.Derive(new List<BoardLane>()) is null, "No sequences: no shape line");

        static BoardCard Sequence(string description) => new("Sequence", "Sequence", description, 1, new List<BoardCard>());
        static BoardCard Scene() => new("INT. ROOM", "Scene", string.Empty, 1, new List<BoardCard>());

        var lanes = new List<BoardLane>
        {
            new(null, new List<BoardGroup>
            {
                new(Sequence("Starts: she arrives\nInciting incident: the letter"), new List<BoardCard> { Scene() }),
                new(Sequence("Plot point 1: [PLOT POINT 1]"), new List<BoardCard>()),
                new(Sequence("Pinch 1: –"), new List<BoardCard>()),
                new(Sequence("Midpoint: the keeper confesses"), new List<BoardCard>())
            })
        };

        Assert(ShapeLine.Derive(lanes) == "▪●▫·▫–▫●▫·▫·▫·▫", "Known, unknown, dropped and missing sequences each get their glyph, in split order");

        var prose = new List<BoardLane> { new(null, new List<BoardGroup> { new(Sequence("Just a synopsis"), new List<BoardCard>()) }) };
        Assert(ShapeLine.Derive(prose) == "▫·▫·▫·▫·▫·▫·▫·▫", "A sequence with no slot lines shows everything unknown, not dropped");
    }

    static void TestSlateStoreSplitFamilyRoundTrip()
    {
        var (_, store, root) = NewSlateFixture();
        try
        {
            var document = new SlateDocument();
            document.Split.A = "arrives";
            document.Split.Midpoint = "confesses";
            document.Split.MidpointAlive = "alive";
            document.Split.BeliefLayer = true;
            document.Belief.Shape = "Fall";
            document.Belief.Cut(50).Text = "she thinks she can leave";
            document.Belief.Cut(50).Alive = "flat";
            document.Extend.Z = "the lighthouse burns";
            document.Extend.Links.Add(new ExtendLink { Text = "she lit it", Mystery = true });
            store.Save("draft", document);

            var loaded = store.Load("draft")!;
            Assert(loaded.Split.Midpoint == "confesses" && loaded.Split.MidpointAlive == "alive" && loaded.Split.BeliefLayer, "Split run round-trips");
            Assert(loaded.Split.Candidates.Count == 3, "Three candidate rows round-trip, empty included");
            Assert(loaded.Belief.Shape == "Fall" && loaded.Belief.Cut(50).Text == "she thinks she can leave" && loaded.Belief.Cut(50).Alive == "flat", "Belief cuts round-trip by ratio");
            Assert(loaded.Belief.Cuts.Count == 5, "All five named cuts, no more");
            Assert(loaded.Extend.Links.Count == 2 && loaded.Extend.Links[1].Mystery, "Extend links round-trip with the mystery lens");
            Assert(!loaded.Split.HasRung2 && !loaded.Belief.HasRung3, "Rung flags read the content, nothing stored");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void TestSlateStoreBridgeAndPositionRoundTrip()
    {
        var (_, store, root) = NewSlateFixture();
        try
        {
            var document = new SlateDocument();
            var bridge = new BridgeRun { A = "the ring is planted", Z = "the ring is found" };
            bridge.Rounds[0].Candidates[1] = "she pawns it";
            bridge.Rounds[0].Picked = "she pawns it";
            bridge.Rounds.Add(new BridgeRound());
            document.Bridges.Add(bridge);
            document.Bridges.Add(new BridgeRun());
            var position = new PositionRun { Moment = "the dog on the roof" };
            position.Turns[0].Slot = "Midpoint";
            position.Turns[0].Alive = "alive";
            position.Shortlist[0] = "midpoint";
            document.Positions.Add(position);
            store.Save("draft", document);

            var loaded = store.Load("draft")!;
            Assert(loaded.Bridges.Count == 2, "Bridges round-trip; pruning empties is the runner's job, not the store's");
            Assert(loaded.Bridges[0].Rounds.Count == 2 && loaded.Bridges[0].Rounds[0].Picked == "she pawns it"
                && loaded.Bridges[0].Rounds[0].Candidates.Count == 3, "Rounds, their three candidates and the picked line round-trip");
            Assert(loaded.Bridges[1].IsEmpty && !loaded.Bridges[0].IsEmpty, "IsEmpty reads every field");
            Assert(loaded.Positions.Count == 1 && loaded.Positions[0].Turns[0].Slot == "Midpoint"
                && loaded.Positions[0].Turns[0].Alive == "alive" && loaded.Positions[0].Shortlist.Count == 2, "Position turns and shortlist round-trip");
            Assert(new PositionRun().IsEmpty && !position.IsEmpty, "A fresh run is empty; one with a moment is not");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void Assert(bool condition, string message)
    {
        if (!condition)
        {
            throw new Exception($"Assertion Failed: {message}");
        }
    }

    static void TestParseSimpleSceneHeading()
    {
        var text = "INT. COFFEE SHOP - DAY\n\nThis is action text.";
        var parser = new FountainParser();
        var screenplay = parser.Parse(text);

        Assert(screenplay.Elements.Count == 2, $"Expected 2 elements, got {screenplay.Elements.Count}");
        Assert(screenplay.Elements[0] is SceneHeadingElement, "First element should be SceneHeadingElement");
        var heading = (SceneHeadingElement)screenplay.Elements[0];
        Assert(heading.Text == "INT. COFFEE SHOP - DAY", $"Incorrect heading text: '{heading.Text}'");

        Assert(screenplay.Elements[1] is ActionElement, "Second element should be ActionElement");
        var action = (ActionElement)screenplay.Elements[1];
        Assert(action.Text == "This is action text.", $"Incorrect action text: '{action.Text}'");
    }

    static void TestParseCharacterAndDialogue()
    {
        var text = "INT. COFFEE SHOP - DAY\n\nJOHN\nWhat are you doing?";
        var parser = new FountainParser();
        var screenplay = parser.Parse(text);

        Assert(screenplay.Elements.Count == 3, $"Expected 3 elements, got {screenplay.Elements.Count}");
        Assert(screenplay.Elements[1] is CharacterElement, "Second element should be CharacterElement");
        var character = (CharacterElement)screenplay.Elements[1];
        Assert(character.CharacterName == "JOHN", $"Incorrect character name: '{character.CharacterName}'");

        Assert(screenplay.Elements[2] is DialogueElement, "Third element should be DialogueElement");
        var dialogue = (DialogueElement)screenplay.Elements[2];
        Assert(dialogue.Text == "What are you doing?", $"Incorrect dialogue text: '{dialogue.Text}'");
        Assert(dialogue.CharacterName == "JOHN", $"Incorrect character on dialogue: '{dialogue.CharacterName}'");
    }

    static void TestParseParenthetical()
    {
        var text = "JOHN\n(smiling)\nI am writing a script.";
        var parser = new FountainParser();
        var screenplay = parser.Parse(text);

        Assert(screenplay.Elements.Count == 3, $"Expected 3 elements, got {screenplay.Elements.Count}");
        Assert(screenplay.Elements[0] is CharacterElement, "First element should be CharacterElement");
        
        Assert(screenplay.Elements[1] is ParentheticalElement, "Second element should be ParentheticalElement");
        var paren = (ParentheticalElement)screenplay.Elements[1];
        Assert(paren.Text == "smiling", $"Incorrect parenthetical text: '{paren.Text}'");

        Assert(screenplay.Elements[2] is DialogueElement, "Third element should be DialogueElement");
        var dialogue = (DialogueElement)screenplay.Elements[2];
        Assert(dialogue.Text == "I am writing a script.", $"Incorrect dialogue text: '{dialogue.Text}'");
    }

    static void TestParseExplicitLineOverrides()
    {
        // Line 1 is Character, but we override it as Scene Heading via type overrides.
        // Remember line numbers are 1-based in lineTypeOverrides
        var text = "JOHN\n(smiling)\nI am writing a script.";
        var parser = new FountainParser();
        var overrides = new Dictionary<int, ScreenplayElementType>
        {
            { 1, ScreenplayElementType.SceneHeading }
        };
        var screenplay = parser.Parse(text, overrides);

        Assert(screenplay.Elements.Count > 0, "Expected parsed elements");
        Assert(screenplay.Elements[0] is SceneHeadingElement, "First element should be overridden to SceneHeadingElement");
        var heading = (SceneHeadingElement)screenplay.Elements[0];
        Assert(heading.Text == "JOHN", $"Expected overridden heading text 'JOHN', got '{heading.Text}'");
    }

    static void TestTextAnalysisHelperMethods()
    {
        var headingCandidate = "INT. ROOM - NIGHT";
        Assert(TextAnalysis.LooksLikeSceneHeadingStart(headingCandidate.AsSpan()), "INT. should be recognized as scene heading start");
        
        var actionCandidate = "The room is dark.";
        Assert(!TextAnalysis.LooksLikeSceneHeadingStart(actionCandidate.AsSpan()), "Action text should not look like scene heading start");

        var uppercaseName = "MARY";
        Assert(TextAnalysis.IsUppercaseLike(uppercaseName), "MARY should be uppercase like");

        var mixedcaseName = "Mary";
        Assert(!TextAnalysis.IsUppercaseLike(mixedcaseName), "Mary should not be uppercase like");
    }
}
