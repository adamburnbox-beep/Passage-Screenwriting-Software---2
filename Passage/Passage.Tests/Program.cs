using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Microsoft.Extensions.Configuration;
using Passage.Parser;
using Passage.Core;
using Passage.Core.Extensibility;
using Passage.Export;
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
        failures += RunTest("Test Parse Headings Without Scene Numbers Match Golden", TestParseHeadingsWithoutSceneNumbersMatchGolden);
        failures += RunTest("Test Parse Scene Numbers", TestParseSceneNumbers);
        failures += RunTest("Test Layout Carries Scene Numbers", TestLayoutCarriesSceneNumbers);
        failures += RunTest("Test SceneNumbers Remove", TestSceneNumbersRemove);
        failures += RunTest("Test SceneNumbers Add Keeps Existing", TestSceneNumbersAddKeepsExisting);
        failures += RunTest("Test Characters Bare Name", TestCharactersBareName);
        failures += RunTest("Test Characters Stats And Speakers", TestCharactersStatsAndSpeakers);
        failures += RunTest("Test Characters Rename Cues Only", TestCharactersRenameCuesOnly);
        failures += RunTest("Test Characters Settings Round Trip", TestCharactersSettingsRoundTrip);
        failures += RunTest("Test Gender Report Reconciles", TestGenderReportReconciles);
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
        failures += RunTest("Test Structure Templates Print Nothing", TestStructureTemplatesPrintNothing);
        failures += RunTest("Test Structure Template Passage Is The Shape", TestStructureTemplatePassageIsTheShape);
        failures += RunTest("Test Structure Template Page Scaling", TestStructureTemplatePageScaling);
        failures += RunTest("Test Structure Template Seed Skips Duplicates", TestStructureTemplateSeedSkipsDuplicates);
        failures += RunTest("Test Structure Template Save Round Trip", TestStructureTemplateSaveRoundTrip);
        failures += RunTest("Test SlateStore Keeps Templates File", TestSlateStoreKeepsTemplatesFile);
        failures += RunTest("Test ShapeLine Derives From Lanes", TestShapeLineDerivesFromLanes);
        failures += RunTest("Test Story Stamp Keeps Id And Scene Number", TestStoryStampKeepsIdAndSceneNumber);
        failures += RunTest("Test Story Duplicate Id Flags Second Only", TestStoryDuplicateIdFlagsSecondOnly);
        failures += RunTest("Test Story Detached Is Computed Never Removed", TestStoryDetachedIsComputedNeverRemoved);
        failures += RunTest("Test Story Delete Beat Removes Tags", TestStoryDeleteBeatRemovesTags);
        failures += RunTest("Test Tag Toggle Adds Then Removes", TestTagToggleAddsThenRemoves);
        failures += RunTest("Test Tag Stamps Unstamped Card Once", TestTagStampsUnstampedCardOnce);
        failures += RunTest("Test Tag Rename Rewrites And Merges", TestTagRenameRewritesAndMerges);
        failures += RunTest("Test Tag Characters Are Bare Names", TestTagCharactersAreBareNames);
        failures += RunTest("Test Locations Normalise Prefixes", TestLocationsNormalisePrefixes);
        failures += RunTest("Test Locations Normalise Times Of Day", TestLocationsNormaliseTimesOfDay);
        failures += RunTest("Test Locations Normalise Edge Cases", TestLocationsNormaliseEdgeCases);
        failures += RunTest("Test Locations Build Groups Scenes", TestLocationsBuildGroupsScenes);
        failures += RunTest("Test Story Filter Match Table", TestStoryFilterMatchTable);
        failures += RunTest("Test Story Filter Scene Speakers", TestStoryFilterSceneSpeakers);
        failures += RunTest("Test Story Old Sidecar Loads", TestStoryOldSidecarLoads);
        failures += RunTest("Test Storylines Round Trip", TestStorylinesRoundTrip);
        failures += RunTest("Test Storyline Delete Untags Items", TestStorylineDeleteUntagsItems);
        failures += RunTest("Test Storyline Colour Rejects Unknown", TestStorylineColourRejectsUnknown);
        failures += RunTest("Test SlateStore Split Family Round Trip", TestSlateStoreSplitFamilyRoundTrip);
        failures += RunTest("Test SlateStore Bridge And Position Round Trip", TestSlateStoreBridgeAndPositionRoundTrip);
        failures += RunTest("Test SlateStore Revise Round Trip", TestSlateStoreReviseRoundTrip);
        failures += RunTest("Test SlateStore Ideation Is Kept Apart", TestSlateStoreIdeationIsKeptApart);
        failures += RunTest("Test IdeationDealer Deals Every Lane", TestIdeationDealerDealsEveryLane);
        failures += RunTest("Test Suggestions Ask For The Next Field", TestSuggestionsAskForTheNextField);
        failures += RunTest("Test NullStoryPartner Offers Nothing", TestNullStoryPartnerOffersNothing);
        failures += RunTest("Test SyntaxSchemes Presets Are Complete", TestSyntaxSchemesPresetsAreComplete);
        failures += RunTest("Test SyntaxSchemes Custom Falls Back And Round Trips", TestSyntaxSchemesCustomFallsBackAndRoundTrips);
        failures += RunTest("Test Pagination Null Options Matches Golden", TestPaginationNullOptionsMatchesGolden);
        failures += RunTest("Test Pagination Keep Together", TestPaginationKeepTogether);
        failures += RunTest("Test Pagination Overflow Tolerance", TestPaginationOverflowTolerance);
        failures += RunTest("Test Pagination More And Continued", TestPaginationMoreAndContinued);
        failures += RunTest("Test Pagination More Falls Back To Line Ends", TestPaginationMoreFallsBackToLineEnds);
        failures += RunTest("Test Highland Export Import Round Trip", TestHighlandExportImportRoundTrip);
        failures += RunTest("Test Highland Hand Built Bundle Imports", TestHighlandHandBuiltBundleImports);
        failures += RunTest("Test Highland Without Text Fails Visibly", TestHighlandWithoutTextFailsVisibly);
        failures += RunTest("Test FDX Export Paragraph Types", TestFdxExportParagraphTypes);
        failures += RunTest("Test Page Directives Parse", TestPageDirectivesParse);
        failures += RunTest("Test Layout Header Footer Directives", TestLayoutHeaderFooterDirectives);
        failures += RunTest("Test Layout Strips Markers", TestLayoutStripsMarkers);
        failures += RunTest("Test Outline Lists Named Markers", TestOutlineListsNamedMarkers);
        failures += RunTest("Test Outline Filters", TestOutlineFilters);
        failures += RunTest("Test BeatSheet Sidecar Round Trip", TestBeatSheetSidecarRoundTrip);
        failures += RunTest("Test BeatSheet Clamping", TestBeatSheetClamping);
        failures += RunTest("Test BeatSheet Stacking", TestBeatSheetStacking);
        failures += RunTest("Test BeatSheet Scene Pages Match Preview", TestBeatSheetScenePagesMatchPreview);
        failures += RunTest("Test Beside List Order Matches Board", TestBesideListOrderMatchesBoard);
        failures += RunTest("Test Beside Caret Lookup", TestBesideCaretLookup);
        failures += RunTest("Test Inbox Conversion Lines", TestInboxConversionLines);
        failures += RunTest("Test Inbox Tag Transfer", TestInboxTagTransfer);
        failures += RunTest("Test Inbox Insert Points", TestInboxInsertPoints);
        failures += RunTest("Test Inbox Missing Target Is Reported", TestInboxMissingTarget);

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

    static void TestLocationsNormalisePrefixes()
    {
        // Every prefix in decision C, with and without the period, in either case.
        foreach (var prefix in new[] { "INT.", "EXT.", "INT./EXT.", "EXT./INT.", "INT/EXT", "I/E.", "I/E", "EST.", "INT", "EXT", "int.", "Ext" })
            Assert(Locations.Normalise($"{prefix} KITCHEN - DAY") == "KITCHEN", $"Prefix stripped: {prefix}");

        Assert(Locations.Normalise(".KITCHEN") == "KITCHEN", "A forced heading's leading . is stripped");
        Assert(Locations.Normalise("INT. KITCHEN - DAY") == Locations.Normalise("EXT. KITCHEN - NIGHT"), "INT and EXT are one place");
        Assert(Locations.Normalise("INTERIOR HALL") == "INTERIOR HALL", "A word that merely starts with INT is not a prefix");
        Assert(Locations.Normalise("EXTRA ROOM - DAY") == "EXTRA ROOM", "EXTRA is not the EXT prefix");
    }

    static void TestLocationsNormaliseTimesOfDay()
    {
        foreach (var time in new[] { "DAY", "NIGHT", "MORNING", "AFTERNOON", "EVENING", "DAWN", "DUSK", "SUNRISE", "SUNSET",
            "LATER", "MOMENTS LATER", "CONTINUOUS", "SAME", "SAME TIME", "MAGIC HOUR" })
            Assert(Locations.Normalise($"INT. KITCHEN - {time}") == "KITCHEN", $"Time of day dropped: {time}");

        Assert(Locations.Normalise("INT. KITCHEN \u2014 NIGHT") == "KITCHEN", "An em dash separates too");
        Assert(Locations.Normalise("int. kitchen - day") == "KITCHEN", "Result is uppercased");
        Assert(Locations.Normalise("INT.   KITCHEN   -   DAY") == "KITCHEN", "Whitespace is collapsed");

        // Sub-locations stay separate; only trailing times are dropped.
        Assert(Locations.Normalise("INT. HOUSE - KITCHEN - NIGHT") == "HOUSE - KITCHEN", "A multi-dash sub-location keeps its parent");
        Assert(Locations.Normalise("INT. HOUSE - KITCHEN - NIGHT") != Locations.Normalise("INT. HOUSE - DAY"), "HOUSE - KITCHEN is not HOUSE");
        Assert(Locations.Normalise("INT. KITCHEN - DAY - LATER") == "KITCHEN", "Repeated trailing times are all dropped");
        Assert(Locations.Normalise("INT. KITCHEN - LATER - FLASHBACK") == "KITCHEN - LATER - FLASHBACK", "Only the last segments are dropped, and only while they are times");
        Assert(Locations.Normalise("INT. KITCHEN - FLASHBACK") == "KITCHEN - FLASHBACK", "An unknown trailing segment stays");
        Assert(Locations.Normalise("INT. DAYCARE - NIGHT") == "DAYCARE", "A time word inside a name is not a time");
    }

    static void TestLocationsNormaliseEdgeCases()
    {
        Assert(Locations.Normalise(".FLASHBACK") == "FLASHBACK", "Forced heading with content is a location");
        foreach (var empty in new[] { "", "   ", ".", "INT.", "INT. - DAY", "EXT. NIGHT", "DAY", "INT. - " })
            Assert(Locations.Normalise(empty) == "", $"Empty after stripping is not a location: '{empty}'");

        Assert(Locations.Normalise("INT. KITCHEN - DAY #12#") == "KITCHEN", "A scene number is ignored");
        Assert(Locations.Normalise($"INT. KITCHEN - DAY [[id:{Guid.NewGuid()}]]") == "KITCHEN", "An id marker is ignored");
        Assert(Locations.Normalise($"INT. KITCHEN - DAY #12# [[id:{Guid.NewGuid()}]]") == "KITCHEN", "Number and id together are ignored");
        Assert(Locations.Normalise("INT. ROOM #5 - DAY") == "ROOM #5", "A # inside the name is kept");
        Assert(Locations.Normalise(". NEW SCENE") == "NEW SCENE", "A Board-created scene normalises to NEW SCENE");
    }

    static void TestLocationsBuildGroupsScenes()
    {
        var script = "INT. KITCHEN - DAY\n\nA.\n\nEXT. GARDEN - DAY\n\nB.\n\nINT. KITCHEN - NIGHT\n\nC.\n\nEXT. KITCHEN - DAY\n\nD.\n\nINT. HOUSE - KITCHEN - NIGHT\n\nE.";
        var places = Locations.Build(new FountainParser().Parse(script).Elements);
        Assert(places.Select(place => place.Name).SequenceEqual(new[] { "KITCHEN", "GARDEN", "HOUSE - KITCHEN" }), "Locations are listed by first appearance");
        Assert(places[0].Scenes.Count == 3 && places[0].FirstLine == 1, "Three kitchen scenes are one location");
        Assert(places[0].Scenes.Select(scene => scene.LineNumber).SequenceEqual(new[] { 1, 9, 13 }), "Each scene keeps its line for the jump");
        Assert(places[2].Scenes.Count == 1, "The sub-location is its own entry");
    }

    static void TestStoryStampKeepsIdAndSceneNumber()
    {
        foreach (var (line, number) in new[] { ("INT. HOUSE - DAY", (string?)null), ("INT. HOUSE - DAY #12#", "12") })
        {
            var id = Guid.NewGuid();
            var stamped = StoryItems.StampLine(line, id);
            var scene = new FountainParser().Parse(stamped + "\n\nAction.").Elements.OfType<SceneHeadingElement>().Single();
            Assert(scene.Id == id, $"Stamped id survives the parse: {line}");
            Assert(scene.SceneNumber == number, $"Scene number survives stamping: {line}");
            Assert(scene.Heading == "INT. HOUSE - DAY", $"Heading text is unchanged: {line}");
        }

        var heading = "# ACT ONE";
        var actId = Guid.NewGuid();
        Assert(new FountainParser().Parse(StoryItems.StampLine(heading, actId)).Elements.OfType<SectionElement>().Single().Id == actId,
            "A section heading stamps too");

        var swapped = StoryItems.StampLine($"INT. A - DAY [[id:{Guid.NewGuid()}]]", actId);
        Assert(swapped.Contains(actId.ToString()) && swapped.Split("[[id:").Length == 2, "Stamping a duplicate replaces its marker, not adds a second");
    }

    static void TestStoryDuplicateIdFlagsSecondOnly()
    {
        var id = Guid.NewGuid();
        var text = $"INT. A - DAY [[id:{id}]]\n\nOne.\n\nINT. B - DAY\n\nTwo.\n\nINT. A - DAY [[id:{id}]]\n\nThree.\n";
        var analysis = new DocumentAnalyzer().Analyze(text);
        Assert(analysis.DuplicateIdLines is { Count: 1 } && analysis.DuplicateIdLines.Contains(9),
            "Only the second occurrence (line 9) is flagged");
    }

    static void TestStoryDetachedIsComputedNeverRemoved()
    {
        var (_, store, root) = NewSlateFixture();
        try
        {
            var id = Guid.NewGuid();
            var document = new SlateDocument();
            document.Story.Items.Add(new ItemTags { TargetId = id, Kind = "card", Label = "INT. A - DAY", Key = true });
            document.Story.Items.Add(new ItemTags { TargetId = Guid.NewGuid(), Kind = "beat", Key = true });
            store.Save("draft", document);
            var loaded = store.Load("draft")!;
            Assert(loaded.Story.Items.Count == 2, "Load keeps items whose card is gone");

            var without = new FountainParser().Parse("INT. B - DAY\n\nOne.\n").Elements;
            var detached = StoryItems.Detached(loaded.Story.Items, without);
            Assert(detached.Count == 1 && detached[0].TargetId == id, "Only the card item detaches; beats never do");
            store.Save("draft", loaded);
            Assert(store.Load("draft")!.Story.Items.Count == 2, "Save keeps detached items");

            var back = new FountainParser().Parse($"INT. A - DAY [[id:{id}]]\n\nOne.\n").Elements;
            Assert(StoryItems.Detached(loaded.Story.Items, back).Count == 0, "The item re-attaches when the id parses again");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void TestStoryDeleteBeatRemovesTags()
    {
        var story = new StoryData();
        var beat = Guid.NewGuid();
        var card = Guid.NewGuid();
        StoryItems.SetKey(story, beat, "beat", "Beat", true);
        StoryItems.SetKey(story, card, "card", "Scene", true);
        StoryItems.RemoveBeat(story, beat);
        Assert(story.Items.Count == 1 && story.Items[0].TargetId == card, "Deleting a beat removes its tags and leaves cards alone");
        StoryItems.SetKey(story, card, "card", "Scene", false);
        Assert(story.Items.Count == 0, "Un-starring an item that carries nothing removes it");
    }

    static void TestTagToggleAddsThenRemoves()
    {
        var story = new StoryData();
        var target = Guid.NewGuid();
        var line = Guid.NewGuid();
        var storyKey = "storyline:" + line;
        Assert(StoryItems.ToggleTag(story, target, "card", "Scene", storyKey), "First toggle turns the tag on");
        Assert(story.Items.Single().Storylines.Single() == line, "The storyline is on the item");
        Assert(StoryItems.ToggleTag(story, target, "card", "Scene", "location:KITCHEN"), "A location toggles on");
        Assert(!StoryItems.ToggleTag(story, target, "card", "Scene", storyKey), "Second toggle turns the tag off");
        Assert(story.Items.Single().Storylines.Count == 0 && story.Items.Single().Locations.Single() == "KITCHEN", "Only that tag is removed");
        StoryItems.ToggleTag(story, target, "card", "Scene", "location:KITCHEN");
        Assert(story.Items.Count == 0, "An item left carrying nothing is dropped");
        Assert(!StoryItems.ToggleTag(story, target, "card", "Scene", "storyline:not-a-guid") && story.Items.Count == 0,
            "A malformed key is rejected and creates nothing");
        StoryItems.SetKey(story, target, "card", "Scene", true);
        StoryItems.ToggleTag(story, target, "card", "Scene", "character:BOB");
        StoryItems.ToggleTag(story, target, "card", "Scene", "character:BOB");
        Assert(story.Items.Single().Key, "Untagging the last tag keeps a starred item");
    }

    static void TestTagStampsUnstampedCardOnce()
    {
        var line = "INT. HOUSE - DAY #12#";
        Assert(StoryItems.NeedsStamp(line, false), "An unmarked heading needs a stamp");
        var id = Guid.NewGuid();
        var stamped = StoryItems.StampLine(line, id);
        Assert(!StoryItems.NeedsStamp(stamped, false), "A stamped heading is not stamped again");
        Assert(StoryItems.NeedsStamp(stamped, true), "A copied marker counts as unstamped");
        var scene = new FountainParser().Parse(stamped + "\n\nOne.\n").Elements.OfType<SceneHeadingElement>().Single();
        Assert(scene.Id == id, "The stamp parses back to the same id");
        Assert(StoryItems.StampLine(stamped, Guid.NewGuid()).Split("[[id:").Length == 2, "Restamping never leaves two markers");
    }

    static void TestTagRenameRewritesAndMerges()
    {
        var story = new StoryData();
        var a = Guid.NewGuid();
        var b = Guid.NewGuid();
        var c = Guid.NewGuid();
        StoryItems.ToggleTag(story, a, "card", "A", "character:BOB");
        StoryItems.ToggleTag(story, b, "beat", "B", "character:BOB");
        StoryItems.ToggleTag(story, b, "beat", "B", "character:ROBERT");
        StoryItems.ToggleTag(story, c, "beat", "C", "character:MARA");
        var changed = StoryItems.RenameCharacter(story, "BOB", "ROBERT");
        Assert(changed == 2, "Two items carried BOB");
        Assert(story.Items.First(i => i.TargetId == a).Characters.Single() == "ROBERT", "BOB is rewritten to ROBERT");
        Assert(story.Items.First(i => i.TargetId == b).Characters.Single() == "ROBERT", "An item with both names merges to one");
        Assert(story.Items.First(i => i.TargetId == c).Characters.Single() == "MARA", "Other characters are untouched");
    }

    static void TestTagCharactersAreBareNames()
    {
        var story = new StoryData();
        var target = Guid.NewGuid();
        StoryItems.ToggleTag(story, target, "card", "Scene", "character:BOB (V.O.)");
        Assert(story.Items.Single().Characters.Single() == "BOB", "BOB (V.O.) is stored as BOB");
        Assert(!StoryItems.ToggleTag(story, target, "card", "Scene", "character:BOB (O.S.)") && story.Items.Count == 0,
            "A different extension toggles the same character off");
    }

    static void TestStorylinesRoundTrip()
    {
        var (_, store, root) = NewSlateFixture();
        try
        {
            var document = new SlateDocument();
            var line = new Storyline { Name = "A Story", Colour = Characters.Palette[4], Description = "Mara and the key" };
            document.Story.Storylines.Add(line);
            document.Story.Items.Add(new ItemTags { TargetId = Guid.NewGuid(), Kind = "beat", Storylines = { line.Id } });
            store.Save("draft", document);
            var loaded = store.Load("draft")!;
            var back = loaded.Story.Storylines.Single();
            Assert(back.Id == line.Id && back.Name == "A Story" && back.Colour == Characters.Palette[4] && back.Description == "Mara and the key",
                "A storyline keeps its id, name, colour and description");
            Assert(loaded.Story.Items.Single().Storylines.Single() == line.Id, "Item tags keep the storyline id");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void TestStorylineDeleteUntagsItems()
    {
        var story = new StoryData();
        var a = new Storyline { Name = "A Story" };
        var b = new Storyline { Name = "B Story" };
        story.Storylines.AddRange(new[] { a, b });
        story.Items.Add(new ItemTags { TargetId = Guid.NewGuid(), Kind = "card", Storylines = { a.Id, b.Id } });
        story.Items.Add(new ItemTags { TargetId = Guid.NewGuid(), Kind = "beat", Storylines = { a.Id } });
        story.Items.Add(new ItemTags { TargetId = Guid.NewGuid(), Kind = "beat", Key = true, Storylines = { a.Id } });
        story.Items.Add(new ItemTags { TargetId = Guid.NewGuid(), Kind = "beat", Storylines = { b.Id } });
        Assert(StoryItems.StorylineUseCount(story, a.Id) == 3, "The count matches the items tagged");
        var untagged = StoryItems.DeleteStoryline(story, a.Id);
        Assert(untagged == 3, "Delete reports how many items it untagged");
        Assert(story.Storylines.Single() == b, "Only the deleted storyline goes");
        Assert(story.Items.All(item => !item.Storylines.Contains(a.Id)), "No item still carries the deleted id");
        Assert(story.Items.Count == 3 && story.Items.Count(item => item.Key) == 1 && story.Items.Count(item => item.Storylines.Contains(b.Id)) == 2,
            "Items still starred or tagged stay; the one left empty is dropped");
    }

    static void TestStorylineColourRejectsUnknown()
    {
        Assert(StoryItems.StorylineColour(Characters.Palette[2]) == Characters.Palette[2], "A palette colour is kept");
        Assert(StoryItems.StorylineColour("#123456").Length == 0, "An unknown colour is rejected");
        Assert(StoryItems.StorylineColour(null).Length == 0, "Null reads as no colour");
    }

    static void TestStoryFilterMatchTable()
    {
        var a = Guid.NewGuid();
        var b = Guid.NewGuid();
        var none = Array.Empty<string>();
        var noLines = Array.Empty<Guid>();

        Assert(new StoryFilter().Matches(false, noLines, none, none), "An empty filter matches everything");
        Assert(!new StoryFilter().IsActive, "An empty filter is inactive");

        var lines = new StoryFilter { Storylines = { a, b } };
        Assert(lines.Matches(false, new[] { b }, none, none), "OR within a kind: one storyline is enough");
        Assert(!lines.Matches(false, noLines, none, none), "An untagged item fails a storyline filter");

        var both = new StoryFilter { Storylines = { a }, Locations = { "KITCHEN" } };
        Assert(both.Matches(false, new[] { a }, none, new[] { "KITCHEN" }), "AND across kinds: both satisfied");
        Assert(!both.Matches(false, new[] { a }, none, new[] { "HALL" }), "AND across kinds: location missing");
        Assert(!both.Matches(false, noLines, none, new[] { "KITCHEN" }), "AND across kinds: storyline missing");

        Assert(!new StoryFilter { KeyOnly = true }.Matches(false, noLines, none, none), "Key only dims a non-key item");
        Assert(new StoryFilter { KeyOnly = true, Characters = { "BOB" } }.Matches(true, noLines, new[] { "BOB" }, none), "Key only ANDs with the other kinds");
        Assert(!new StoryFilter { KeyOnly = true, Characters = { "BOB" } }.Matches(true, noLines, new[] { "ANN" }, none), "Key does not rescue a character miss");

        Assert(StoryFilter.LaneMatches(false, new[] { false, true }), "A lane matches via a child card");
        Assert(!StoryFilter.LaneMatches(false, new[] { false, false }), "A lane with no matching card dims");
        Assert(StoryFilter.LaneMatches(true, Array.Empty<bool>()), "A lane matches by its own tags");

        var toggled = new StoryFilter().ToggleLocation("KITCHEN").ToggleLocation("HALL").ToggleLocation("KITCHEN");
        Assert(toggled.Locations.SequenceEqual(new[] { "HALL" }), "Toggling twice removes a choice");

        // Implicit location and derived speaker, as the Editor builds them.
        var elements = new FountainParser().Parse("INT. KITCHEN - DAY\n\nBOB (V.O.)\nHello.\n\nINT. HALL - NIGHT\n\nANN\nHi.\n").Elements;
        var speakers = StoryFilter.SceneSpeakers(elements);
        var scenes = elements.OfType<SceneHeadingElement>().ToList();
        var kitchen = new HashSet<string>(speakers[scenes[0].LineNumber]);
        var filter = new StoryFilter { Characters = { "BOB" }, Locations = { Locations.Normalise(scenes[0].Text) } };
        Assert(filter.Matches(false, noLines, kitchen, new[] { Locations.Normalise(scenes[0].Text) }), "A scene matches by its heading and a (V.O.) speaker");
        Assert(!filter.Matches(false, noLines, speakers[scenes[1].LineNumber], new[] { Locations.Normalise(scenes[1].Text) }), "The other scene dims");
    }

    static void TestStoryFilterSceneSpeakers()
    {
        var text = "INT. KITCHEN - DAY\n\nBOB (V.O.)\nOne.\n\nANN LEE\nTwo.\n\nEXT. YARD - DAY\n\nBOB\nThree.\n\nINT. EMPTY - DAY\n\nNobody speaks.\n";
        var elements = new FountainParser().Parse(text).Elements;
        var scenes = elements.OfType<SceneHeadingElement>().ToList();
        var speakers = StoryFilter.SceneSpeakers(elements);
        Assert(speakers[scenes[0].LineNumber].SetEquals(new[] { "BOB", "ANN LEE" }), "Cues are bare names: (V.O.) stripped");
        Assert(speakers[scenes[1].LineNumber].SetEquals(new[] { "BOB" }), "Speakers do not leak into the next scene");
        Assert(speakers[scenes[2].LineNumber].Count == 0, "A scene with no cues has no speakers");
    }

    static void TestStoryOldSidecarLoads()
    {
        var (_, store, root) = NewSlateFixture();
        try
        {
            var dir = Path.Combine(root, ".slate");
            Directory.CreateDirectory(dir);
            File.WriteAllText(Path.Combine(dir, "old.fountain.json"), "{ \"SlateVersion\": 1, \"Beats\": [] }");
            var loaded = store.Load("old");
            Assert(loaded is not null && loaded.Story.Items.Count == 0, "A sidecar with no Story loads with an empty one");
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

    static string StripIds(string line) => System.Text.RegularExpressions.Regex.Replace(line, @"\s*\[\[id:[0-9a-fA-F-]+\]\]", string.Empty);

    static void TestStructureTemplatesPrintNothing()
    {
        var expected = new Dictionary<string, (int Acts, int Sequences, int Sections)>
        {
            ["8 sequences / 7 turns (Passage)"] = (4, 8, 0),
            ["Paradigm (Syd Field)"] = (3, 0, 5),
            ["Save the Cat (Blake Snyder)"] = (3, 0, 15),
            ["Five-act TV"] = (6, 0, 4)
        };
        Assert(StructureTemplates.BuiltIns.Count == 4, "Four built-in templates");
        foreach (var template in StructureTemplates.BuiltIns)
        {
            var text = string.Join("\n", StructureTemplates.BuildLines(template));
            var lines = text.Split('\n');
            Assert(lines.All(line => line.Length == 0 || line.StartsWith("# ") || line.StartsWith("## ") || line.StartsWith("### ") || line.StartsWith("= ")),
                $"{template.Name} writes only #, ##, ### and = lines");
            Assert(template.Description.Length > 0, $"{template.Name} credits its source in the description");

            var analysis = new DocumentAnalyzer().Analyze(text);
            Assert(analysis.Elements.All(element => element.Type is ScreenplayElementType.Section or ScreenplayElementType.Synopsis),
                $"{template.Name} has no printable element");
            Assert(analysis.PageCount <= 1, $"{template.Name} adds no pages, got {analysis.PageCount}");
            var (acts, sequences, sections) = expected[template.Name];
            Assert(analysis.Elements.Count(e => e.Type == ScreenplayElementType.Section && e.Level == 0) == acts, $"{template.Name}: {acts} acts");
            Assert(analysis.Elements.Count(e => e.Type == ScreenplayElementType.Section && e.Level == 1) == sequences, $"{template.Name}: {sequences} sequences");
            Assert(analysis.Elements.Count(e => e.Type == ScreenplayElementType.Section && e.Level == 2) == sections, $"{template.Name}: {sections} sections");
        }
    }

    static void TestStructureTemplatePassageIsTheShape()
    {
        var passage = StructureTemplates.BuiltIns[0];
        var fromTemplate = StructureTemplates.BuildLines(passage).Select(StripIds).ToList();
        var fromShape = SplitScript.BuildLines(new SplitRun()).Select(StripIds).ToList();
        Assert(fromTemplate.SequenceEqual(fromShape), "Passage's template is SplitScript.BuildLines of an empty run");
    }

    static void TestStructureTemplatePageScaling()
    {
        Assert(StructureTemplates.ScalePage(55, 110) == 55, "On the reference length a page is itself");
        Assert(StructureTemplates.ScalePage(55, 90) == 45, "Page 55 of 110 is 45 of 90 (45.0)");
        Assert(StructureTemplates.ScalePage(1, 90) == 1, "Scaling never goes below page 1");
        Assert(StructureTemplates.ScalePage(110, 90) == 90, "The last reference page is the last target page");
        Assert(StructureTemplates.ScalePage(110, 300) == 300, "A longer target scales up");
        Assert(StructureTemplates.ScalePage(5, 3) is >= 1 and <= 3, "A tiny target clamps onto the grid");
    }

    static void TestStructureTemplateSeedSkipsDuplicates()
    {
        var template = StructureTemplates.BuiltIns.Single(entry => entry.Name.StartsWith("Save the Cat"));
        var withPage = template.Nodes.Count(node => node.Page > 0);
        var beats = new List<Beat>();
        var (added, skipped) = StructureTemplates.SeedBeats(template, beats, 90, 90);
        Assert(added == withPage && skipped == 0 && beats.Count == withPage, "Every node with a page seeds a beat");
        var midpoint = beats.Single(beat => beat.Title == "Midpoint");
        Assert(midpoint.Page == 45, $"Midpoint lands at page 45 of 90, got {midpoint.Page}");
        Assert(beats.All(beat => beat.Page is >= 1 and <= 90), "Every seeded beat is on the grid");

        (added, skipped) = StructureTemplates.SeedBeats(template, beats, 90, 90);
        Assert(added == 0 && skipped == withPage && beats.Count == withPage, "A second run adds nothing and counts what was there");

        var existing = new List<Beat> { new() { Title = "  catalyst " } };
        (added, skipped) = StructureTemplates.SeedBeats(template, existing, 90, 90);
        Assert(skipped == 1 && added == withPage - 1, "A title match ignores case and spacing");
    }

    static void TestStructureTemplateSaveRoundTrip()
    {
        var script = string.Join("\n", new[]
        {
            "# Act One [[id:" + Guid.NewGuid() + "]]",
            "= The setup.",
            "",
            "## Opening [[id:" + Guid.NewGuid() + "]]",
            "",
            ".INT. KITCHEN - DAY [[id:" + Guid.NewGuid() + "]]",
            "= A scene synopsis that is not the template's.",
            "",
            "ANNA",
            "Hello.",
            "",
            "### Catalyst [[id:" + Guid.NewGuid() + "]]",
            "= Something breaks.",
            "= Twice.",
        });
        var template = StructureTemplates.FromScript("Mine", script.Split('\n'));
        Assert(template.Nodes.Select(node => node.Kind).SequenceEqual(new[] { "Act", "Sequence", "Section" }), "Headings only, scenes excluded");
        Assert(template.Nodes[0].Prompt == "The setup." && template.Nodes[1].Prompt.Length == 0, "A heading keeps its own = lines");
        Assert(template.Nodes[2].Prompt == "Something breaks.\nTwice.", "Several = lines stay several");
        Assert(!template.Nodes.Any(node => node.Heading.Contains("id:")), "Ids are not kept");

        var written = StructureTemplates.BuildLines(template);
        var again = StructureTemplates.FromScript("Mine", written);
        Assert(again.Nodes.Select(node => (node.Kind, node.Heading, node.Prompt)).SequenceEqual(template.Nodes.Select(node => (node.Kind, node.Heading, node.Prompt))),
            "Writing and reading a template gives it back");
        Assert(!string.Join("\n", written).Contains("KITCHEN"), "No scene text in the written template");
        Assert(!StructureTemplates.HasStructure(new[] { "### Only a section", "= x" }) && StructureTemplates.HasStructure(new[] { "text", "## Seq" }),
            "Templates are offered unless there is a # or ## heading");

        var (library, store, root) = NewSlateFixture();
        try
        {
            store.SaveTemplates(new TemplatesDocument { Templates = { template } });
            var loaded = store.LoadTemplates()!.Templates.Single();
            Assert(loaded.Name == "Mine" && loaded.Nodes.Count == 3 && loaded.Nodes[2].Prompt == "Something breaks.\nTwice.",
                "A saved template reloads, as it would for another script");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void TestSlateStoreKeepsTemplatesFile()
    {
        var (library, store, root) = NewSlateFixture();
        try
        {
            store.SaveTemplates(new TemplatesDocument { Templates = { new StructureTemplate { Name = "Kept" } } });
            store.Save("gone", new SlateDocument());
            var pruned = store.PruneOrphans(Array.Empty<string>());
            Assert(pruned.Count == 1 && pruned[0] == "gone.fountain", "Only the orphan sidecar is pruned");
            Assert(store.LoadTemplates() is { Templates.Count: 1 }, "The orphan sweep leaves templates.json alone");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
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

    static void TestSuggestionsAskForTheNextField()
    {
        // Fill: the worksheet so far, in order, empties skipped.
        var fill = new FillRun { Bracket = "[SOMETHING forces her to stay]" };
        fill.Options[0].Text = "Her sister calls";
        fill.Options[0].WhyWrong = "Too convenient";
        var fillAsk = Suggestions.ForFill(fill, "[SOMETHING forces her to stay] She sits.");
        Assert(fillAsk.Tool == "Fill" && fillAsk.Count == 3, "Fill asks for three under its own name");
        Assert(fillAsk.Context.Select(line => line.Label).SequenceEqual(new[] { "Bracket", "The line it's on", "Candidate 1", "Why candidate 1 is probably wrong" }),
            "Fill's context is the bracket, its line, and only the written cells, in reading order");
        Assert(fillAsk.Context[1].Text == "[SOMETHING forces her to stay] She sits.", "The script line rides along for the partner");

        // WOAC: the next empty answer of the last round, in the chain's order; null when all written.
        var chain = new ChainRun { Path = ChainPath.Woac, Seed = "MARA wants to leave" };
        Assert(Suggestions.ForChain(chain, Array.Empty<string>())?.Ask.StartsWith("Want") == true, "Round 1 starts by asking the want");
        chain.Rounds[0].Want = "to get out the door";
        chain.Rounds[0].Obstacle = "the key is in his coat";
        var woac = Suggestions.ForChain(chain, Array.Empty<string>());
        Assert(woac?.Ask.StartsWith("Action") == true, "Then the first empty of obstacle, action, consequence");
        Assert(woac!.Context.Select(line => line.Text).SequenceEqual(new[] { "MARA wants to leave", "to get out the door", "the key is in his coat" }),
            "The chain so far is the context");
        chain.Rounds[0].Action = "she goes through his pockets";
        chain.Rounds[0].Consequence = "he stirs";
        chain.Rounds.Add(new ChainRound());
        var round2 = Suggestions.ForChain(chain, Array.Empty<string>());
        Assert(round2?.Ask.StartsWith("Obstacle") == true, "Round 2 reads its want from the consequence above and asks the obstacle");
        Assert(round2!.Context.Any(line => line.Label == "Round 2 — want" && line.Text == "he stirs"), "The carried want is in the context");
        chain.Rounds[1].Obstacle = "o"; chain.Rounds[1].Action = "a"; chain.Rounds[1].Consequence = "c";
        Assert(Suggestions.ForChain(chain, Array.Empty<string>()) is null, "Nothing to ask when every answer is written");

        // Flaw: the first unanswered question, asked in the worksheet's words.
        var questions = new[] { "Worst situation the flaw causes", "The character's first (wrong) response" };
        var flaw = new ChainRun { Path = ChainPath.Flaw, Seed = "MARA — never asks for help" };
        flaw.Answers[0] = "She loses the flat";
        var flawAsk = Suggestions.ForChain(flaw, questions);
        Assert(flawAsk?.Ask == "2. The character's first (wrong) response", "Flaw asks the next question by its number and words");
        Assert(flawAsk!.Context.Last().Label == "1. Worst situation the flaw causes", "Answered questions carry their question as the label");

        // Extend: the open link, mystery lens changing the ask.
        var extend = new ExtendRun { Z = "The lighthouse burns" };
        extend.Links[0].Text = "Someone knocked over the lamp";
        extend.Links.Add(new ExtendLink { Mystery = true });
        var extendAsk = Suggestions.ForExtendLink(extend);
        Assert(extendAsk.Ask.StartsWith("Mystery lens"), "The open link's mystery lens changes the question");
        Assert(extendAsk.Context.Select(line => line.Text).SequenceEqual(new[] { "The lighthouse burns", "Someone knocked over the lamp" }), "Z and the links so far, in order");

        // Bridge: the ends, earlier rounds' picks, this round's candidates.
        var bridge = new BridgeRun { A = "The ring is planted", Z = "The ring turns up at the pawnshop" };
        bridge.Rounds[0].Picked = "She finds it (half)";
        bridge.Rounds.Add(new BridgeRound());
        bridge.Rounds[1].Candidates[0] = "She pawns it herself";
        var bridgeAsk = Suggestions.ForBridgeRound(bridge);
        Assert(bridgeAsk.Context.Select(line => line.Label).SequenceEqual(new[] { "A — the point you're starting from", "Z — the point you're building to", "Round 1, picked / half-picked", "Round 2, candidate 1" }),
            "Bridge's context is the ends, the picks so far and this round's candidates");

        // Split: the ends and the candidates so far.
        var split = new SplitRun { A = "A stowaway is found", Z = "The keeper rows away" };
        split.Candidates[1] = "The supply boat comes early";
        var splitAsk = Suggestions.ForSplitMidpoint(split);
        Assert(splitAsk.Context.Count == 3 && splitAsk.Context[2].Label == "Candidate 2", "Split skips the empty candidate slot and keeps the written one's number");
    }

    static void TestNullStoryPartnerOffersNothing()
    {
        IStoryPartner partner = new NullStoryPartner();
        var request = new SuggestionRequest("Fill", "Candidate", Array.Empty<SuggestionLine>(), 3);
        var offered = partner.SuggestAsync(request, CancellationToken.None).GetAwaiter().GetResult();
        Assert(offered.Count == 0, "The null partner offers nothing, so the dock has nothing to show");
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

    static void TestSlateStoreReviseRoundTrip()
    {
        var (_, store, root) = NewSlateFixture();
        try
        {
            var document = new SlateDocument();
            var revision = new ReviseRun { Scene = "INT. KITCHEN - DAY", Purpose = "she decides to leave" };
            revision.Check("Polarity").Answer = "starts warm, ends warm";
            revision.Check("Polarity").Flag = "Y";
            revision.Check("Polarity").Lens = "Tone swap, minimum edit";
            revision.Check("Polarity").Fragment = "line one\nline two";
            revision.Check("Polarity").Moved = "Y";
            revision.Gsu[0].Character = "MARA";
            revision.Gsu[0].Goal = "get the keys back";
            revision.Gsu.Add(new GsuRow { Character = "TOM", Urgency = "the bus leaves at six" });
            revision.GsuFlag = "Y";
            document.Revisions.Add(revision);
            document.Burst.LensMinutes = 3;
            store.Save("draft", document);

            var loaded = store.Load("draft")!;
            Assert(loaded.Revisions.Count == 1 && loaded.Revisions[0].Scene == "INT. KITCHEN - DAY", "Revise run round-trips by scene");
            Assert(loaded.Revisions[0].Checks.Count == ReviseRun.CheckNames.Length, "All six checks are stored, unanswered ones included");
            var polarity = loaded.Revisions[0].Check("Polarity");
            Assert(polarity.Flag == "Y" && polarity.Lens == "Tone swap, minimum edit" && polarity.Fragment == "line one\nline two" && polarity.Moved == "Y", "A flagged check keeps its Lens, multi-line fragment and verdict");
            Assert(loaded.Revisions[0].Check("Linkage").IsEmpty, "An untouched check is empty");
            Assert(loaded.Burst.LensMinutes == 3, "Lens minutes round-trip");
            var gsu = loaded.Revisions[0].Gsu;
            Assert(gsu.Count == 3 && gsu[0].Character == "MARA" && gsu[0].Goal == "get the keys back" && gsu[1].IsEmpty && gsu[2].Urgency == "the bus leaves at six", "GSU rows round-trip in order, an empty row included");
            Assert(loaded.Revisions[0].GsuFlag == "Y", "The GSU flag round-trips");
            Assert(new ReviseRun().Gsu.Count == 2, "A new run starts with two GSU rows");
            Assert(!new ReviseRun { GsuFlag = "N" }.IsEmpty && !new ReviseRun { Gsu = { new GsuRow { Stakes = "x" } } }.IsEmpty, "GSU content makes a run non-empty");
            Assert(new ReviseRun().IsEmpty && !revision.IsEmpty && !revision.HasDiagnostic && !revision.HasReadBack, "Emptiness and the deeper sections read the content");
            Assert(new ReviseRun().Check("Made up").Name == "Made up", "Check() adds a missing check rather than throwing");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void TestIdeationDealerDealsEveryLane()
    {
        var rng = new Random(7);
        for (var lane = 1; lane <= 10; lane++)
        {
            var dealt = Enumerable.Range(0, 12).Select(_ => IdeationDealer.Deal(lane, rng)).ToList();
            Assert(dealt.All(d => !string.IsNullOrWhiteSpace(d)), $"Lane {lane} never deals an empty input");
            Assert(dealt.Distinct().Count() > 1, $"Lane {lane} deals more than one thing");
        }

        // The colliding lanes never collide a thing with itself.
        for (var i = 0; i < 200; i++)
        {
            var halves = IdeationDealer.Deal(4, rng).Split("  ×  ");
            Assert(halves.Length == 2 && halves[0] != halves[1], "Situation collision deals two different fragments");
        }

        Assert(!string.IsNullOrWhiteSpace(IdeationDealer.Deal(99, rng)), "An out-of-range lane still deals something");
    }

    static void TestSlateStoreIdeationIsKeptApart()
    {
        var (_, store, root) = NewSlateFixture();
        try
        {
            Assert(store.LoadIdeation() is null, "No ideation file loads as null");

            var document = new IdeationDocument { RoundSeconds = 60, Current = new IdeationSitting { Lane = 4 } };
            document.Current.Rounds[0].Input = "a kettle and a eulogy";
            document.Current.Rounds[0].Output = "line one\nline two";
            document.Current.KeptLine = "line two";
            document.Kept.Add("an earlier sitting's line");
            store.SaveIdeation(document);

            Assert(File.Exists(Path.Combine(root, ".slate", SlateStore.IdeationFile)), "Ideation lives in its own file under .slate");

            var loaded = store.LoadIdeation()!;
            Assert(loaded.RoundSeconds == 60 && loaded.Current is { Lane: 4 } && loaded.Current.Rounds[0].Output == "line one\nline two"
                && loaded.Current.KeptLine == "line two", "The sitting in progress round-trips");
            Assert(loaded.Kept.SequenceEqual(new[] { "an earlier sitting's line" }), "Kept lines round-trip as plain lines");

            var pruned = store.PruneOrphans(Array.Empty<string>());
            Assert(pruned.Count == 0 && store.LoadIdeation() is not null, "The orphan sweep leaves the ideation file alone — it belongs to no script");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void TestHighlandExportImportRoundTrip()
    {
        // CRLF, a trailing newline, non-ASCII and a leading blank line must all survive.
        var text = "\r\nTitle: Café Noir\n\nINT. CAFÉ - NIGHT\r\n\nJOSÉ\n(quietly)\nÇa va?\n\n> FADE OUT.\n";
        var result = new ExportService().Export(text, "highland", "Noir.fountain");
        Assert(result is not null, "highland export should be recognised");
        Assert(result!.Value.FileName == "Noir.highland", $"file name, got {result.Value.FileName}");
        Assert(result.Value.ContentType == "application/zip", "content type");

        using (var zip = new ZipArchive(new MemoryStream(result.Value.Bytes)))
        {
            var names = zip.Entries.Select(entry => entry.FullName).OrderBy(name => name).ToArray();
            var expected = new[]
            {
                "Noir.textbundle/characters.json", "Noir.textbundle/info.json",
                "Noir.textbundle/sprints.json", "Noir.textbundle/text.fountain"
            };
            Assert(names.SequenceEqual(expected), $"bundle layout, got {string.Join(", ", names)}");

            using var info = JsonDocument.Parse(zip.GetEntry("Noir.textbundle/info.json")!.Open());
            Assert(info.RootElement.GetProperty("version").GetInt32() == 2, "info.json version");
            Assert(info.RootElement.GetProperty("type").GetString() == "com.quoteunquoteapps.fountain", "info.json type");
            Assert(info.RootElement.TryGetProperty("app.passage", out _), "info.json namespaces Passage state");
        }

        var imported = HighlandBundle.Read(new MemoryStream(result.Value.Bytes));
        Assert(imported.Text == text, "round-tripped text should be identical");
        Assert(!imported.IsMarkdown, "exported bundle is Fountain");
        Assert(imported.Skipped.Count == 3, $"the three sidecars are reported, got {imported.Skipped.Count}");
    }

    static void TestFdxExportParagraphTypes()
    {
        var text = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "H2Fdx.fountain"));
        var result = new ExportService().Export(text, "fdx", "Test.fountain");
        Assert(result is not null, "fdx export should be recognised");
        Assert(result!.Value.FileName == "Test.fdx", $"file name, got {result.Value.FileName}");
        Assert(result.Value.ContentType == "application/xml", "content type");

        // XDocument.Parse throws on malformed XML.
        var document = XDocument.Parse(Encoding.UTF8.GetString(result.Value.Bytes));
        var root = document.Root!;
        Assert(root.Name == "FinalDraft" && (string?)root.Attribute("DocumentType") == "Script", "FinalDraft Script root");

        static string Describe(XElement paragraph) =>
            $"{(string?)paragraph.Attribute("Type")}|{string.Concat(paragraph.Elements("Text").Select(run => run.Value))}";

        var body = root.Element("Content")!.Elements("Paragraph").ToList();
        var actual = body.Select(paragraph => paragraph.Element("DualDialogue") is null ? Describe(paragraph) : "DualDialogue").ToArray();
        var expected = new[]
        {
            "Scene Heading|INT. HOUSE - DAY",
            "Action|Jo enters with great care and speed.",
            "Action|A second line of action.",
            "Character|JO",
            "Parenthetical|(quietly)",
            "Dialogue|Hi.",
            "Parenthetical|(beat)",
            "Dialogue|Bye.",
            "DualDialogue",
            "Transition|CUT TO:",
            "Scene Heading|EXT. GARDEN - NIGHT",
            "Action|THE END",
            "General|Sing a song",
            "Scene Heading|FORCED SCENE"
        };
        Assert(actual.SequenceEqual(expected), $"body paragraphs, got:\n{string.Join("\n", actual)}");

        Assert((string?)body[0].Attribute("Number") == "1", "scene number carried as Number");
        Assert(body[10].Attribute("Number") is null, "unnumbered heading has no Number");
        Assert((string?)body[11].Attribute("Alignment") == "Center", "centered text is centre-aligned Action");

        var dual = body[8].Element("DualDialogue")!.Elements("Paragraph").Select(Describe).ToArray();
        Assert(dual.SequenceEqual(new[] { "Character|BRAD", "Dialogue|Me first.", "Character|MAX", "Dialogue|Me too." }),
            $"dual dialogue pair, got {string.Join(", ", dual)}");

        var styles = body[1].Elements("Text").Select(run => $"{run.Value}={(string?)run.Attribute("Style")}").ToArray();
        Assert(styles.SequenceEqual(new[] { "Jo enters with =", "great =Italic", "care=Bold", " and =", "speed=Underline", ".=" }),
            $"emphasis runs, got {string.Join(", ", styles)}");

        var title = root.Element("TitlePage")!.Element("Content")!.Elements("Paragraph").ToList();
        var titleLines = title.Select(paragraph => $"{(string?)paragraph.Attribute("Alignment")}|{Describe(paragraph)}").ToArray();
        var expectedTitle = new[]
        {
            "Center|General|THE FDX TEST", "Center|General|written by", "Center|General|Jo Writer",
            "Left|General|1 May 2026", "Left|General|jo@example.com", "Left|General|555 0100"
        };
        Assert(titleLines.SequenceEqual(expectedTitle), $"title page, got:\n{string.Join("\n", titleLines)}");
        Assert((string?)title[0].Element("Text")!.Attribute("Style") == "Bold+Underline", "title keeps its emphasis");
    }

    static void TestHighlandHandBuiltBundleImports()
    {
        // Shaped like a real Highland 2 save: BOM-less UTF-8, Highland's own
        // info.json keys, extra sidecars, an assets directory and a Finder
        // directory entry for the bundle itself.
        var text = "INT. HOUSE - DAY\n\nMARY\nHello.\n";
        using var buffer = new MemoryStream();
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            void Add(string path, string content)
            {
                using var writer = new StreamWriter(zip.CreateEntry(path).Open(), new UTF8Encoding(false));
                writer.Write(content);
            }

            zip.CreateEntry("Draft.textbundle/");
            Add("Draft.textbundle/info.json",
                "{\"com.quoteunquoteapps.highland2\":{\"templateName\":\"Screenplay\"},\"transient\":false,\"type\":\"com.quoteunquoteapps.fountain\",\"version\":2}");
            Add("Draft.textbundle/text.fountain", text);
            Add("Draft.textbundle/characters.json", "{\"MARY\":{\"gender\":\"female\"}}");
            Add("Draft.textbundle/navigatorFilters.json", "{\"notes\":true}");
            zip.CreateEntry("Draft.textbundle/assets/");
            Add("Draft.textbundle/assets/still.png", "not really a png");
        }

        buffer.Position = 0;
        var imported = HighlandBundle.Read(buffer);
        Assert(imported.Text == text, "text should be read verbatim");
        Assert(!imported.IsMarkdown, "text.fountain is Fountain");
        var expectedSkipped = new[]
        {
            "Draft.textbundle/info.json", "Draft.textbundle/characters.json",
            "Draft.textbundle/navigatorFilters.json", "Draft.textbundle/assets/still.png"
        };
        Assert(imported.Skipped.SequenceEqual(expectedSkipped),
            $"every non-text file is reported, got {string.Join(", ", imported.Skipped)}");

        // A Markdown bundle zipped from its parent folder.
        using var markdown = new MemoryStream();
        using (var zip = new ZipArchive(markdown, ZipArchiveMode.Create, leaveOpen: true))
        {
            using var writer = new StreamWriter(zip.CreateEntry("Notes/Notes.textbundle/text.md").Open());
            writer.Write("# Notes\n");
        }

        markdown.Position = 0;
        var notes = HighlandBundle.Read(markdown);
        Assert(notes.IsMarkdown && notes.Text == "# Notes\n", "text.md imports as Markdown");
        Assert(notes.Skipped.Count == 0, "nothing else to skip");
    }

    static void TestHighlandWithoutTextFailsVisibly()
    {
        using var buffer = new MemoryStream();
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            using var writer = new StreamWriter(zip.CreateEntry("Draft.textbundle/info.json").Open());
            writer.Write("{}");
        }

        foreach (var bytes in new[] { buffer.ToArray(), Encoding.UTF8.GetBytes("not a zip") })
        {
            var threw = false;
            try
            {
                HighlandBundle.Read(new MemoryStream(bytes));
            }
            catch (InvalidDataException)
            {
                threw = true;
            }

            Assert(threw, "a bundle with no script text must throw, not return empty text");
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

    static string SerializeElements(ParsedScreenplay screenplay)
    {
        var builder = new StringBuilder();
        foreach (var element in screenplay.Elements)
        {
            builder.Append(element.Type).Append('|')
                .Append(element.LineIndex.ToString(CultureInfo.InvariantCulture)).Append('|')
                .Append(element is SceneHeadingElement { IsForced: true } ? "forced|" : "|")
                .Append(element.Text).Append('|')
                .Append(element.RawText).Append('\n');
        }
        return builder.ToString();
    }

    static void TestParseHeadingsWithoutSceneNumbersMatchGolden()
    {
        // H1.3 shared-code guard: headings without #…# must parse exactly as before the
        // scene-number change. The golden was captured from the pre-H1.3 parser.
        var golden = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "H1SceneHeadings.golden.txt"));
        var parsed = ParseFixture("H1SceneHeadings.fountain");
        var actual = SerializeElements(parsed);
        Assert(actual == golden, $"Headings without scene numbers must parse as before H1.3. Got:\n{actual}");
    }

    static SceneHeadingElement OnlyHeading(string text) =>
        new FountainParser().Parse(text).Elements.OfType<SceneHeadingElement>().Single();

    static void TestParseSceneNumbers()
    {
        var plain = OnlyHeading("INT. HOUSE - DAY #12#");
        Assert(plain.Text == "INT. HOUSE - DAY", $"Number must be stripped from the text, got '{plain.Text}'");
        Assert(plain.SceneNumber == "12", $"Expected scene number 12, got '{plain.SceneNumber}'");
        Assert(plain.RawText == "INT. HOUSE - DAY #12#", "RawText keeps the number");

        var forced = OnlyHeading(".FLASHBACK #4A#");
        Assert(forced.IsForced && forced.Text == "FLASHBACK" && forced.SceneNumber == "4A", $"Forced heading: '{forced.Text}' #{forced.SceneNumber}#");

        var withId = OnlyHeading("EXT. MEADOW - NIGHT #3.1# [[id:11111111-2222-3333-4444-555555555555]]");
        Assert(withId.Text == "EXT. MEADOW - NIGHT" && withId.SceneNumber == "3.1", $"With id: '{withId.Text}' #{withId.SceneNumber}#");
        Assert(withId.Id == Guid.Parse("11111111-2222-3333-4444-555555555555"), "Id still parsed alongside the number");

        var overridden = new FountainParser().Parse("JOHN #7#\nHello.", new Dictionary<int, ScreenplayElementType> { { 1, ScreenplayElementType.SceneHeading } })
            .Elements.OfType<SceneHeadingElement>().Single();
        Assert(overridden.Text == "JOHN" && overridden.SceneNumber == "7", $"Override: '{overridden.Text}' #{overridden.SceneNumber}#");

        Assert(OnlyHeading("INT. ROOM #2 - DAY").SceneNumber is null, "A # inside the heading is not a scene number");
        Assert(OnlyHeading("INT. HOUSE - DAY").SceneNumber is null, "Unnumbered heading has no number");
    }

    static void TestLayoutCarriesSceneNumbers()
    {
        var parsed = new FountainParser().Parse("INT. HOUSE - DAY #12#\n\nAction.\n\nEXT. YARD - DAY\n\nMore.");
        var lines = ScreenplayLayoutBuilder.BuildPages(parsed).bodyPages.SelectMany(p => p.Lines).Where(l => !l.IsBlank).ToList();
        Assert(lines[0].Text == "INT. HOUSE - DAY" && lines[0].SceneNumber == "12", $"Heading line: '{lines[0].Text}' #{lines[0].SceneNumber}#");
        Assert(lines.Skip(1).All(l => l.SceneNumber is null), "Only the numbered heading carries a number");

        var text = new ScreenplayExporter().GetExportText(parsed);
        Assert(text.Contains("12  INT. HOUSE - DAY  12"), $"Plain text keeps the number in both margins, got:\n{text}");

        var pdfPath = Path.Combine(Path.GetTempPath(), $"passage-h13-{Guid.NewGuid():N}.pdf");
        try
        {
            new ScreenplayPdfExporter().Export(parsed, pdfPath);
            var pdf = Encoding.Latin1.GetString(File.ReadAllBytes(pdfPath));
            var left = $"1 0 0 1 {ScreenplayLayoutBuilder.LeftSceneNumberX("12").ToString("0.###", CultureInfo.InvariantCulture)} ";
            var right = $"1 0 0 1 {ScreenplayLayoutBuilder.RightSceneNumberX.ToString("0.###", CultureInfo.InvariantCulture)} ";
            Assert(pdf.Contains(left) && pdf.Contains(right), "PDF draws the scene number in both margins");
            Assert(pdf.Split("(12) Tj").Length - 1 == 2, "PDF draws the number exactly twice");
        }
        finally
        {
            File.Delete(pdfPath);
        }
    }

    static string ApplyEdits(string content, IReadOnlyList<SceneNumbers.LineEdit> edits)
    {
        var lines = content.Split('\n');
        foreach (var edit in edits)
        {
            Assert(lines[edit.LineIndex] == edit.Expected, $"Edit expects '{edit.Expected}' on line {edit.LineIndex}, found '{lines[edit.LineIndex]}'");
            lines[edit.LineIndex] = edit.Replacement;
        }
        return string.Join('\n', lines);
    }

    // ---- Character panel (H3.1) ----

    static void TestCharactersBareName()
    {
        Assert(Characters.BareName("SARAH (V.O.)") == "SARAH", "Extension stripped");
        Assert(Characters.BareName("@McClane (CONT'D)") == "MCCLANE", "@ stripped, uppercased");
        Assert(Characters.BareName("  BOB (O.S.) (CONT'D) ^") == "BOB", "Stacked extensions and ^ stripped");
        Assert(Characters.BareName("MAN (OLD) TWO") == "MAN (OLD) TWO", "A mid-name parenthesis is part of the name");
    }

    static void TestCharactersStatsAndSpeakers()
    {
        var content = "INT. ROOM - DAY\n\nSARAH\nHello there.\n\nBOB (V.O.)\n(quietly)\nHi.\n\nSARAH (CONT'D)\nThree more words.\n\nSarah waves.";
        var analysis = new DocumentAnalyzer().Analyze(content);
        var sarah = analysis.Characters.Single(stat => stat.Name == "SARAH");
        var bob = analysis.Characters.Single(stat => stat.Name == "BOB");
        Assert(analysis.Characters.Count == 2, $"Extensions do not make new characters, got {string.Join(", ", analysis.Characters.Select(c => c.Name))}");
        Assert(sarah.Lines == 2 && sarah.Words == 5, $"SARAH: 2 speeches, 5 words; got {sarah.Lines}/{sarah.Words}");
        Assert(bob.Lines == 1 && bob.Words == 1, $"BOB: parenthetical not counted; got {bob.Lines}/{bob.Words}");
        Assert(analysis.Characters[0].Name == "SARAH", "Most speeches first");

        var speakers = analysis.LineSpeakers;
        Assert(speakers[2] == "SARAH" && speakers[3] == "SARAH", "Cue and dialogue lines belong to the speaker");
        Assert(speakers[5] == "BOB" && speakers[6] == "BOB" && speakers[7] == "BOB", "Parenthetical is inside the block");
        Assert(speakers[0] == "" && speakers[4] == "" && speakers[12] == "", "Heading, blank and action lines have no speaker");
    }

    static void TestCharactersRenameCuesOnly()
    {
        var content = "SARAH\nBob, wait.\n\nBOB (V.O.)\nSARAH is here.\n\n@BOB (CONT'D) ^\nStill me.\n\nBob leaves.";
        var analysis = new DocumentAnalyzer().Analyze(content);
        var result = ApplyEdits(content, Characters.Rename(content, analysis.Elements, "BOB", "Robert"));
        Assert(result == "SARAH\nBob, wait.\n\nROBERT (V.O.)\nSARAH is here.\n\n@ROBERT (CONT'D) ^\nStill me.\n\nBob leaves.",
            $"Rename rewrites cues only and keeps @, extensions and ^, got:\n{result}");
        Assert(Characters.Rename(content, analysis.Elements, "NOBODY", "X").Count == 0, "Unknown name renames nothing");
        Assert(Characters.Rename(content, analysis.Elements, "BOB", "  ").Count == 0, "Blank name renames nothing");
    }

    static void TestCharactersSettingsRoundTrip()
    {
        var (_, store, root) = NewSlateFixture();
        try
        {
            var document = new SlateDocument();
            document.Characters.Add(new CharacterSettings { Name = "SARAH", Colour = Characters.Palette[0], Gender = "Female" });
            store.Save("draft", document);
            var loaded = store.Load("draft")!;
            Assert(loaded.Characters.Count == 1 && loaded.Characters[0].Name == "SARAH" && loaded.Characters[0].Colour == Characters.Palette[0],
                "Character colour survives a save and load");
            Assert(loaded.Characters[0].Gender == "Female", "Gender tag survives a save and load");
        }
        finally
        {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    // ---- Dialogue balance report (H3.2) ----

    static void TestGenderReportReconciles()
    {
        var content = "INT. ROOM - DAY\n\nSARAH\nHello there.\n\nBOB (V.O.)\nHi.\n\nCLERK\nNext.\n\n"
            + "EXT. STREET - NIGHT\n\nSARAH\nThree more words.\n\nANNE\nOne two three four.\n\n"
            + "INT. CAR - NIGHT\n\nBOB\nAlone again.";
        var analysis = new DocumentAnalyzer().Analyze(content);
        // "Other" is not a tag and reads as Unspecified; CLERK has no entry.
        var genders = new Dictionary<string, string> { ["SARAH"] = "Female", ["ANNE"] = "Female", ["BOB"] = "Male", ["CLERK"] = "Other" };
        var report = Characters.Report(analysis.Elements, analysis.Characters, genders);

        Assert(report.Rows.Count == analysis.Characters.Count, "One row per character in the panel");
        foreach (var total in report.Totals)
        {
            var rows = report.Rows.Where(row => row.Tag == total.Tag).ToList();
            Assert(total.Characters == rows.Count, $"{total.Tag}: character count {total.Characters} vs {rows.Count} rows");
            Assert(total.Lines == rows.Sum(row => row.Lines), $"{total.Tag}: lines {total.Lines} vs rows {rows.Sum(row => row.Lines)}");
            Assert(total.Words == rows.Sum(row => row.Words), $"{total.Tag}: words {total.Words} vs rows {rows.Sum(row => row.Words)}");
        }

        Assert(report.Totals.Sum(t => t.Lines) == report.Rows.Sum(r => r.Lines) && report.Totals.Sum(t => t.Words) == report.Rows.Sum(r => r.Words),
            "Grand totals match the table");
        var female = report.Totals.Single(t => t.Tag == "Female");
        var male = report.Totals.Single(t => t.Tag == "Male");
        var other = report.Totals.Single(t => t.Tag == "Unspecified");
        Assert(female.Characters == 2 && female.Lines == 3 && female.Words == 9, $"Female 2/3/9, got {female.Characters}/{female.Lines}/{female.Words}");
        Assert(male.Characters == 1 && male.Lines == 2 && male.Words == 3, $"Male 1/2/3, got {male.Characters}/{male.Lines}/{male.Words}");
        Assert(other.Characters == 1 && other.Lines == 1 && other.Words == 1, $"Unspecified 1/1/1, got {other.Characters}/{other.Lines}/{other.Words}");

        // Scene 1: SARAH+BOB (F/M), SARAH+CLERK (F/U), BOB+CLERK (M/U).
        // Scene 2: SARAH+ANNE (F/F). Scene 3: BOB alone, no pair.
        int Pair(string a, string b) => report.Pairs.Single(p => p.First == a && p.Second == b).Scenes;
        Assert(report.Pairs.Count == 6, "Every tag pair is listed once");
        Assert(Pair("Female", "Female") == 1 && Pair("Female", "Male") == 1 && Pair("Female", "Unspecified") == 1
            && Pair("Male", "Unspecified") == 1 && Pair("Male", "Male") == 0 && Pair("Unspecified", "Unspecified") == 0,
            $"Pair counts, got {string.Join(", ", report.Pairs.Select(p => $"{p.First}/{p.Second}={p.Scenes}"))}");
    }

    static void TestSceneNumbersRemove()
    {
        var content = "INT. HOUSE - DAY #12#\n\nAction #1# stays.\n\n.FLASHBACK #4A# [[id:11111111-2222-3333-4444-555555555555]]\n\nINT. ROOM #2 - DAY";
        var result = ApplyEdits(content, SceneNumbers.Remove(content));
        Assert(result == "INT. HOUSE - DAY\n\nAction #1# stays.\n\n.FLASHBACK [[id:11111111-2222-3333-4444-555555555555]]\n\nINT. ROOM #2 - DAY",
            $"Remove strips heading numbers only, got:\n{result}");
        Assert(SceneNumbers.Remove(result).Count == 0, "Remove on an unnumbered script is a no-op");
    }

    static void TestSceneNumbersAddKeepsExisting()
    {
        var fresh = "INT. A - DAY\n\nEXT. B - DAY\n\nINT. C - DAY [[id:11111111-2222-3333-4444-555555555555]]";
        Assert(ApplyEdits(fresh, SceneNumbers.Add(fresh)) == "INT. A - DAY #1#\n\nEXT. B - DAY #2#\n\nINT. C - DAY #3# [[id:11111111-2222-3333-4444-555555555555]]",
            $"Fresh script numbers 1..3, got:\n{ApplyEdits(fresh, SceneNumbers.Add(fresh))}");

        // A new scene between 5 and 6 becomes 5A; one between 5A and 6 becomes 5B;
        // a gap after the last number continues on; existing numbers never move.
        var locked = "INT. A - DAY #5#\n\nINT. NEW1 - DAY\n\nINT. NEW2 - DAY\n\nINT. B - DAY #6#\n\nINT. NEW3 - DAY";
        var result = ApplyEdits(locked, SceneNumbers.Add(locked));
        Assert(result == "INT. A - DAY #5#\n\nINT. NEW1 - DAY #5A#\n\nINT. NEW2 - DAY #5B#\n\nINT. B - DAY #6#\n\nINT. NEW3 - DAY #7#",
            $"Locked numbering, got:\n{result}");
        Assert(SceneNumbers.Add(result).Count == 0, "Add on a fully numbered script is a no-op");

        // New scenes before scene 1 take prefix letters: A1, B1.
        var opening = "INT. NEW1 - DAY\n\nINT. NEW2 - DAY\n\nINT. A - DAY #1#";
        var prefixed = ApplyEdits(opening, SceneNumbers.Add(opening));
        Assert(prefixed == "INT. NEW1 - DAY #A1#\n\nINT. NEW2 - DAY #B1#\n\nINT. A - DAY #1#", $"Before scene 1, got:\n{prefixed}");

        // Room below the first number is still used plainly: new, then 3 → 1, 3.
        var room = "INT. NEW - DAY\n\nINT. A - DAY #3#";
        Assert(ApplyEdits(room, SceneNumbers.Add(room)) == "INT. NEW - DAY #1#\n\nINT. A - DAY #3#", "Room before the first number takes 1");
    }

    static void TestSyntaxSchemesPresetsAreComplete()
    {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var preset in SyntaxSchemes.Presets)
        {
            Assert(ids.Add(preset.Id), $"Duplicate preset id '{preset.Id}'");
            Assert(preset.Id != SyntaxSchemes.CustomId, "A preset must not use the custom id");
            foreach (var (token, _, _) in SyntaxSchemes.Tokens)
            {
                Assert(preset.Dark.TryGetValue(token, out var darkColour) && SyntaxSchemes.IsValidColour(darkColour),
                    $"Preset '{preset.Id}' has no valid dark colour for '{token}'");
                Assert(preset.Light.TryGetValue(token, out var lightColour) && SyntaxSchemes.IsValidColour(lightColour),
                    $"Preset '{preset.Id}' has no valid light colour for '{token}'");
            }
        }

        // Classic must match app.css, which is the palette painted before any
        // scheme is chosen.
        Assert(SyntaxSchemes.Classic.Dark["section"] == "#4FC3F7" && SyntaxSchemes.Classic.Light["section"] == "#1D6FA5",
            "Classic section colours should mirror app.css");

        var (dark, light) = SyntaxSchemes.Resolve(new SyntaxSchemeState { Preset = "ember" });
        Assert(dark["scene"] == "#F0A868" && light["scene"] == "#B35A1E", "A preset id should resolve to that preset");

        var (unknownDark, _) = SyntaxSchemes.Resolve(new SyntaxSchemeState { Preset = "no-such-preset" });
        Assert(unknownDark["scene"] == SyntaxSchemes.Classic.Dark["scene"], "An unknown preset should resolve to Classic");
    }

    static void TestSyntaxSchemesCustomFallsBackAndRoundTrips()
    {
        var state = new SyntaxSchemeState
        {
            Preset = SyntaxSchemes.CustomId,
            CustomDark = new Dictionary<string, string>
            {
                ["scene"] = "#123456",
                ["note"] = "not a colour"
            }
        };

        var (dark, light) = SyntaxSchemes.Resolve(state);
        Assert(dark["scene"] == "#123456", "A custom colour should be used");
        Assert(dark["note"] == SyntaxSchemes.Classic.Dark["note"], "An invalid custom colour should fall back to Classic");
        Assert(dark["character"] == SyntaxSchemes.Classic.Dark["character"], "An unset custom token should fall back to Classic");
        Assert(dark.Count == SyntaxSchemes.Tokens.Count, "A resolved palette should cover every token");
        Assert(light["scene"] == SyntaxSchemes.Classic.Light["scene"], "A missing custom light set should fall back to Classic");

        var restored = SyntaxSchemes.Parse(SyntaxSchemes.Serialize(state));
        Assert(restored.Preset == SyntaxSchemes.CustomId, "Preset should survive the round trip");
        Assert(restored.CustomDark is not null && restored.CustomDark["scene"] == "#123456",
            "Custom colours should survive the round trip");
        Assert(restored.CustomLight is null, "An unset light palette should stay unset");

        Assert(SyntaxSchemes.Parse(null).Preset == SyntaxSchemes.Classic.Id, "No stored blob should mean Classic");
        Assert(SyntaxSchemes.Parse("{not json").Preset == SyntaxSchemes.Classic.Id, "A corrupt blob should mean Classic");
        Assert(SyntaxSchemes.Parse("{\"preset\":\"gone\"}").Preset == SyntaxSchemes.Classic.Id,
            "A preset that no longer exists should mean Classic");
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

        Assert(TextAnalysis.EstimateReadingMinutes(0) == 0, "No words should take no time");
        Assert(TextAnalysis.EstimateReadingMinutes(1) == 1, "Any words should round up to a minute");
        Assert(TextAnalysis.EstimateReadingMinutes(230) == 1, "230 words at the default 230 wpm is one minute");
        Assert(TextAnalysis.EstimateReadingMinutes(231) == 2, "231 words at 230 wpm rounds up to two minutes");
        Assert(TextAnalysis.EstimateReadingMinutes(1000, 100) == 10, "1000 words at 100 wpm is ten minutes");
        Assert(TextAnalysis.EstimateReadingMinutes(500, 0) == 0, "A non-positive rate should not divide by zero");
    }

    // ---- ScreenplayLayoutBuilder pagination (Passage.Export, HIGHLAND-PLAN H1.1) ----
    //
    // Fixtures/H1Pagination.fountain is tuned so the fixed 54-line cut ends pages on a scene
    // heading, a bare character cue and a cue + parenthetical, and opens one with a transition.
    // The golden file was written from the fixed paginator before keep-together existed.

    const double CueX = ScreenplayLayoutBuilder.MarginLeft + 144.0;
    const double ParentheticalX = ScreenplayLayoutBuilder.MarginLeft + 108.0;

    static ParsedScreenplay ParseFixture(string name) =>
        new FountainParser().Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", name)));

    static string SerializePages(IReadOnlyList<LayoutPage> pages)
    {
        var builder = new StringBuilder();
        foreach (var page in pages)
        {
            builder.Append("== PAGE ").Append(page.PageNumber.ToString(CultureInfo.InvariantCulture)).Append('\n');
            foreach (var line in page.Lines)
            {
                builder.Append(line.Style).Append('|').Append(line.X.ToString("R", CultureInfo.InvariantCulture)).Append('|').Append(line.Text).Append('\n');
            }
        }
        return builder.ToString();
    }

    // Why a page's bottom is wrong, or null when it is fine.
    static string? BadPageEnd(LayoutPage page)
    {
        var index = page.Lines.Count - 1;
        while (index >= 0 && page.Lines[index].IsBlank) index--;
        if (index < 0) return null;
        if (page.Lines[index].Style == LayoutTextStyle.LeftBold) return $"scene heading '{page.Lines[index].Text}'";
        while (index >= 0 && page.Lines[index].X == ParentheticalX) index--;
        if (index >= 0 && page.Lines[index].X == CueX) return $"bare character cue '{page.Lines[index].Text}'";
        return null;
    }

    static List<string> ContentLines(IEnumerable<LayoutPage> pages) =>
        pages.SelectMany(page => page.Lines).Where(line => !line.IsBlank).Select(line => line.Text).ToList();

    static void TestPaginationNullOptionsMatchesGolden()
    {
        var golden = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "H1Pagination.null-options.golden.txt"));
        var parsed = ParseFixture("H1Pagination.fountain");

        Assert(SerializePages(ScreenplayLayoutBuilder.BuildPages(parsed).bodyPages) == golden,
            "BuildPages with default options must reproduce the pre-H1.1 output byte-for-byte");
        Assert(SerializePages(ScreenplayLayoutBuilder.BuildPages(parsed, null).bodyPages) == golden,
            "BuildPages with explicit null options must reproduce the pre-H1.1 output byte-for-byte");
    }

    static void TestPaginationKeepTogether()
    {
        var parsed = ParseFixture("H1Pagination.fountain");
        var fixedPages = ScreenplayLayoutBuilder.BuildPages(parsed).bodyPages;
        var keptPages = ScreenplayLayoutBuilder.BuildPages(parsed, new PaginationOptions()).bodyPages;

        // The fixture must actually exercise the rules, or the checks below prove nothing.
        var fixedFaults = fixedPages.Take(fixedPages.Count - 1).Select(BadPageEnd).OfType<string>().ToList();
        Assert(fixedFaults.Any(fault => fault.StartsWith("scene heading")), "Fixture should end a fixed page on a scene heading");
        Assert(fixedFaults.Any(fault => fault.StartsWith("bare character cue")), "Fixture should end a fixed page on a bare character cue");
        Assert(fixedPages.Skip(1).Any(page => page.Lines.Count > 0 && page.Lines[0].Style == LayoutTextStyle.RightWithinBody),
            "Fixture should open a fixed page with a transition");

        for (var index = 0; index < keptPages.Count - 1; index++)
        {
            var fault = BadPageEnd(keptPages[index]);
            Assert(fault is null, $"Page {keptPages[index].PageNumber} ends on a {fault}");
        }

        foreach (var page in keptPages.Skip(1))
        {
            Assert(page.Lines.Count > 0 && !page.Lines[0].IsBlank, $"Page {page.PageNumber} should not open with a blank line");
            Assert(page.Lines[0].Style != LayoutTextStyle.RightWithinBody, $"Page {page.PageNumber} should not open with a transition");
        }

        Assert(keptPages.All(page => page.Lines.Count <= 55), "No page may run more than one line past the bottom margin");
        Assert(ContentLines(keptPages).SequenceEqual(ContentLines(fixedPages)), "Keep-together must not drop or reorder any text");
        Assert(keptPages.Select(page => page.PageNumber).SequenceEqual(Enumerable.Range(1, keptPages.Count)), "Pages should number from 1");
    }

    static void TestPaginationOverflowTolerance()
    {
        // Heading on line 3, 24 one-line actions ending on line 51, then a three-line action
        // on lines 53-55: its last line is a one-line orphan unless it may hang.
        var script = new StringBuilder("INT. ROOM - DAY\n\n");
        for (var index = 0; index < 24; index++) script.Append($"Short action {index}.\n\n");
        script.Append("She crosses the room slowly, counting every board that creaks beneath her feet, and stops only when the window is close enough to touch with one hand.\n");
        var parsed = new FountainParser().Parse(script.ToString());

        var hanging = ScreenplayLayoutBuilder.BuildPages(parsed, new PaginationOptions(OverflowLines: 1)).bodyPages;
        Assert(hanging[0].Lines.Count == 55 && hanging[0].Lines[54].Text == "close enough to touch with one hand.",
            "With one line of tolerance the orphan should hang on page 1");

        var strict = ScreenplayLayoutBuilder.BuildPages(parsed, new PaginationOptions(OverflowLines: 0)).bodyPages;
        Assert(strict[0].Lines.Count == 54 && strict[1].Lines[0].Text == "close enough to touch with one hand.",
            "With no tolerance the orphan should move to page 2");
    }

    // ---- (MORE)/(CONT'D) dialogue splitting (HIGHLAND-PLAN H1.2) ----
    //
    // Fixtures/H1MoreContd.fountain holds a 70-plus-line speech that starts mid-page, so it must
    // split twice, then a dual-dialogue pair placed across a page break.

    const double DialogueX = ScreenplayLayoutBuilder.MarginLeft + 72.0;

    static bool EndsSentence(string text) => text.TrimEnd('"', '\'', ')').EndsWith('.') || text.EndsWith('!') || text.EndsWith('?');

    static List<LayoutLine> Content(LayoutPage page) => page.Lines.Where(line => !line.IsBlank).ToList();

    static void TestPaginationMoreAndContinued()
    {
        var parsed = ParseFixture("H1MoreContd.fountain");
        var speech = parsed.Elements.OfType<DialogueElement>().First();
        var fixedPages = ScreenplayLayoutBuilder.BuildPages(parsed).bodyPages;
        Assert(fixedPages.SelectMany(page => page.Lines).Count(line => line.X == DialogueX) >= 70 + 2,
            "Fixture should hold a speech of at least 70 lines");
        Assert(fixedPages.All(page => page.Lines.All(line => line.Text != "(MORE)")), "Null options must never add (MORE)");

        var pages = ScreenplayLayoutBuilder.BuildPages(parsed, new PaginationOptions()).bodyPages;
        var splits = 0;
        for (var index = 0; index < pages.Count - 1; index++)
        {
            var content = Content(pages[index]);
            if (content[^1].Text != "(MORE)") continue;
            splits++;
            Assert(content[^1].X == CueX, $"(MORE) on page {pages[index].PageNumber} should sit at the cue indent");
            Assert(content[^2].X == DialogueX && EndsSentence(content[^2].Text),
                $"Page {pages[index].PageNumber} should break the speech at a sentence end, not '{content[^2].Text}'");
            var next = pages[index + 1].Lines[0];
            Assert(next.Text == "BOB (CONT'D)" && next.X == CueX, $"Page {pages[index + 1].PageNumber} should open with BOB (CONT'D), not '{next.Text}'");
        }
        Assert(splits == 2, $"The speech should split across three pages, split {splits} times");
        Assert(pages.All(page => page.Lines.Count <= 55), "No page may run more than one line past the bottom margin");

        var spoken = pages.SelectMany(page => page.Lines).TakeWhile(line => line.Text != "MARY").Where(line => line.X == DialogueX).SelectMany(line => line.Text.Split(' '));
        Assert(spoken.SequenceEqual(speech.Text.Split(' ', StringSplitOptions.RemoveEmptyEntries)), "Splitting must not drop or reorder any dialogue");
        Assert(pages.SelectMany(page => page.Lines).Count(line => line.Text == "(quietly)") == 1, "The parenthetical is not repeated after (CONT'D)");

        var dual = pages.Single(page => page.Lines.Any(line => line.Text == "MARY"));
        Assert(dual.Lines.Count(line => line.Text is "MARY" or "BOB") == 2 && dual.Lines.Count(line => line.Text == "Then we row.") == 2,
            "The dual-dialogue pair should stay on one page");

        var plain = ScreenplayLayoutBuilder.BuildPages(parsed, new PaginationOptions(MoreAndContinued: false)).bodyPages;
        Assert(plain.All(page => page.Lines.All(line => line.Text != "(MORE)" && !line.Text.Contains("(CONT'D)"))), "MoreAndContinued: false must not split speeches");

        var custom = ScreenplayLayoutBuilder.BuildPages(parsed, new PaginationOptions(MoreText: "(more)", ContinuedText: "(CONT\u2019D)")).bodyPages;
        Assert(custom.Count(page => Content(page)[^1].Text == "(more)") == 2 && custom[1].Lines[0].Text == "BOB (CONT\u2019D)", "Custom MORE/CONT'D strings should be used");

        // The PDF draws the same pages: each page's stream holds its lines, in order.
        var pdfPath = Path.Combine(Path.GetTempPath(), $"passage-h1-2-{Guid.NewGuid():N}.pdf");
        try
        {
            new ScreenplayPdfExporter(new PaginationOptions()).Export(parsed, pdfPath);
            var streams = File.ReadAllText(pdfPath, Encoding.Latin1).Split("endstream")[..^1]
                .Select(segment => segment[(segment.LastIndexOf("stream", StringComparison.Ordinal) + 6)..]).ToList();
            var titleCount = ScreenplayLayoutBuilder.BuildPages(parsed, new PaginationOptions()).titlePages.Count;
            Assert(streams.Count == titleCount + pages.Count, $"PDF should have {titleCount + pages.Count} pages, has {streams.Count}");
            for (var index = 0; index < pages.Count; index++)
            {
                var stream = streams[titleCount + index];
                var at = 0;
                foreach (var line in Content(pages[index]))
                {
                    var drawn = "(" + line.Text.Replace("(", "\\(").Replace(")", "\\)") + ")";
                    var found = stream.IndexOf(drawn, at, StringComparison.Ordinal);
                    Assert(found >= 0, $"PDF page {pages[index].PageNumber} is missing or misorders '{line.Text}'");
                    at = found + drawn.Length;
                }
            }
        }
        finally
        {
            File.Delete(pdfPath);
        }
    }

    static void TestPaginationMoreFallsBackToLineEnds()
    {
        // Heading on line 3, 20 one-line actions ending on line 43, then a speech that is one
        // sentence of about 20 lines: no sentence ends in the space, so it breaks at a line end.
        var script = new StringBuilder("INT. ROOM - DAY\n\n");
        for (var index = 0; index < 20; index++) script.Append($"Short action {index}.\n\n");
        script.Append("ANNA\n");
        for (var index = 0; index < 60; index++) script.Append($"and then word{index} ");
        script.Append("and done.\n");
        var pages = ScreenplayLayoutBuilder.BuildPages(new FountainParser().Parse(script.ToString()), new PaginationOptions()).bodyPages;

        var first = Content(pages[0]);
        Assert(first[^1].Text == "(MORE)" && first[^2].X == DialogueX && !EndsSentence(first[^2].Text),
            "A sentence longer than the space should break at a line end with (MORE)");
        Assert(pages[0].Lines.Count == 54, "The line-end break should fill the page exactly");
        Assert(pages[1].Lines[0].Text == "ANNA (CONT'D)", "The next page should open with ANNA (CONT'D)");
    }

    // ---- Header / footer directives (H4.1) ----

    static void TestPageDirectivesParse()
    {
        Assert(PageDirectives.TryParse("{{HEADER: DRAFT}}", out var header) && !header.IsFooter
            && header.Slots == new PageSlots("", "", "DRAFT"), "One header slot goes right");
        Assert(PageDirectives.TryParse("  {{footer: Confidential }}", out var footer) && footer.IsFooter
            && footer.Slots == new PageSlots("", "Confidential", ""), "One footer slot goes centre; case-insensitive");
        Assert(PageDirectives.TryParse("{{HEADER: Title | %p.}}", out var two) && two.Slots == new PageSlots("Title", "", "%p."), "Two slots are left and right");
        Assert(PageDirectives.TryParse("{{HEADER: a | b | c}}", out var three) && three.Slots == new PageSlots("a", "b", "c"), "Three slots");
        Assert(PageDirectives.TryParse("{{HEADER: %none}}", out var off) && off.Slots.IsEmpty, "%none turns it off");
        Assert(!PageDirectives.TryParse("Before {{HEADER: x}}", out _), "Only whole lines are directives");
        Assert(!PageDirectives.TryParse("{{%m: Midpoint}}", out _), "Markers are not header/footer directives");
        Assert(new PageSlots("%p of %pp", "", "").Resolve(3, 12).Left == "3 of 12", "%p and %pp resolve");
    }

    static void TestLayoutHeaderFooterDirectives()
    {
        var script = new StringBuilder("INT. HOUSE - DAY\n\n{{HEADER: Title | DRAFT | %p of %pp}}\n\n{{FOOTER: CONFIDENTIAL}}\n\n");
        for (var i = 0; i < 60; i++) script.Append($"Action paragraph {i} goes on for a while.\n\n");
        script.Append("Last words.\n{{HEADER: %none}}\n");
        var parsed = new FountainParser().Parse(script.ToString());

        var legacy = ScreenplayLayoutBuilder.BuildPages(parsed).bodyPages;
        Assert(legacy.SelectMany(p => p.Lines).Count(l => l.Text.Contains("{{")) == 3, "Without options the directive lines still print (desktop unchanged)");
        Assert(legacy.All(p => p.Header is null && p.Footer is null), "Without options pages carry no header/footer");

        var pages = ScreenplayLayoutBuilder.BuildPages(parsed, new PaginationOptions()).bodyPages;
        Assert(pages.Count >= 3, $"Fixture should run to at least 3 pages, got {pages.Count}");
        Assert(pages.SelectMany(p => p.Lines).All(l => !l.Text.Contains("{{")), "Directive lines never print");
        Assert(pages.SelectMany(p => p.Lines).Any(l => l.Text == "Last words."), "Text sharing a paragraph with a directive still prints");
        Assert(pages[0].Header == new PageSlots("Title", "DRAFT", $"1 of {pages.Count}"), $"Page 1 header: {pages[0].Header}");
        Assert(pages[1].Header == new PageSlots("Title", "DRAFT", $"2 of {pages.Count}"), $"Page 2 header: {pages[1].Header}");
        Assert(pages.All(p => p.Footer == new PageSlots("", "CONFIDENTIAL", "")), "Footer is centred on every page");
        Assert(pages[^1].Header is { IsEmpty: true }, "%none turns the header off from its page");

        var unset = ScreenplayLayoutBuilder.BuildPages(new FountainParser().Parse("INT. HOUSE - DAY\n\nAction."), new PaginationOptions()).bodyPages;
        Assert(unset[0].Header is null && unset[0].Footer is null, "No directive keeps the default page number");

        var pdfPath = Path.Combine(Path.GetTempPath(), $"passage-h41-{Guid.NewGuid():N}.pdf");
        try
        {
            new ScreenplayPdfExporter(new PaginationOptions()).Export(parsed, pdfPath);
            var pdf = Encoding.Latin1.GetString(File.ReadAllBytes(pdfPath));
            Assert(pdf.Split("(DRAFT) Tj").Length - 1 == pages.Count - 1, "PDF draws the header on every page until %none");
            Assert(pdf.Split("(CONFIDENTIAL) Tj").Length - 1 == pages.Count, "PDF draws the footer on every page");
            Assert(pdf.Contains($"(1 of {pages.Count}) Tj"), "PDF resolves %p of %pp");
            Assert(!pdf.Contains("(2.) Tj"), "A header replaces the default page number");
            Assert(!pdf.Contains("{{"), "PDF never prints a directive");
        }
        finally
        {
            File.Delete(pdfPath);
        }
    }

    // ---- Markers (H4.2) ----

    static void TestLayoutStripsMarkers()
    {
        var parsed = new FountainParser().Parse("INT. HOUSE - DAY\n\n{{%m: Midpoint}}\n\nShe waits {{%m}} by the door.\n{{%m}}\nThen leaves.\n\n{{%m}} Last.");
        var legacy = ScreenplayLayoutBuilder.BuildPages(parsed).bodyPages.SelectMany(p => p.Lines).ToList();
        Assert(legacy.Any(l => l.Text.Contains("{{%m")), "Without options markers still print (desktop unchanged)");

        var lines = ScreenplayLayoutBuilder.BuildPages(parsed, new PaginationOptions()).bodyPages
            .SelectMany(p => p.Lines).Where(l => !l.IsBlank).Select(l => l.Text).ToList();
        Assert(lines.All(l => !l.Contains("{{")), "Markers never print: " + string.Join(" | ", lines));
        Assert(lines.SequenceEqual(new[] { "INT. HOUSE - DAY", "She waits by the door.", "Then leaves.", "Last." }),
            "Marker-only lines vanish and inline markers leave no gap: " + string.Join(" | ", lines));
        Assert(PageDirectives.MarkerNames("a {{%m: One}} b {{%m}} {{ %m : Two }}").SequenceEqual(new[] { "One", "Two" }), "Named markers only");
    }

    static void TestOutlineListsNamedMarkers()
    {
        var text = "{{%m: Cold open}}\n\nINT. HOUSE - DAY\n\nAction.\n\n{{%m: Midpoint}}\n\n{{%m}}\n\n/*\n{{%m: Cut}}\n*/\n\nEXT. YARD - DAY\n\nMore. {{%m: Inline}}";
        var analysis = new DocumentAnalyzer().Analyze(text);
        var outline = analysis.Outline;
        Assert(outline.Count == 3, $"Expected marker + 2 scenes at the root, got {outline.Count}");
        Assert(outline[0] is { Kind: "Marker", Label: "Cold open", LineNumber: 1 }, $"First root: {outline[0]}");
        Assert(outline[1].Children.Count == 1 && outline[1].Children[0] is { Kind: "Marker", Label: "Midpoint", LineNumber: 7 },
            "Midpoint sits under the first scene; the unnamed and boneyard markers are not listed");
        Assert(outline[2].Children.Count == 1 && outline[2].Children[0].Label == "Inline", "An inline named marker is listed");
        var cards = analysis.BoardLanes.SelectMany(l => l.Groups).SelectMany(g => g.Cards).ToList();
        Assert(cards.Count == 2 && cards.All(c => c.Kind == "Scene"), "Markers never become board cards");
    }

    // ---- Outline filters (H4.3) ----

    static void TestOutlineFilters()
    {
        var text = "# Act One\n\n## Setup\n\nINT. HOUSE - DAY\n\n= She arrives.\n\n[[Check the timeline]]\n\n{{%m: Midpoint}}\n\nEXT. YARD - NIGHT";
        var analysis = new DocumentAnalyzer().Analyze(text);

        static List<string> Flatten(IEnumerable<OutlineNode> nodes, int depth = 0) =>
            nodes.SelectMany(n => new[] { $"{depth}:{n.Kind}:{n.Label}" }.Concat(Flatten(n.Children, depth + 1))).ToList();

        var all = Flatten(new OutlineFilters { Notes = true }.Apply(analysis.Outline));
        Assert(all.Any(n => n.EndsWith(":Note:Check the timeline")), "Notes are in the outline: " + string.Join(", ", all));

        var defaults = Flatten(new OutlineFilters().Apply(analysis.Outline));
        Assert(!defaults.Any(n => n.Contains(":Note:")), "Notes are hidden by default");
        Assert(defaults.Count == all.Count - 1, "Defaults hide only the note");

        var noSections = Flatten(new OutlineFilters { Sections = false }.Apply(analysis.Outline));
        Assert(noSections.Contains("0:Scene:INT. HOUSE - DAY") && noSections.Contains("1:Synopsis:She arrives."),
            "Hiding sections promotes their scenes to the root, keeping scene children: " + string.Join(", ", noSections));
        Assert(!noSections.Any(n => n.Contains(":Act:") || n.Contains(":Sequence:")), "No sections remain");

        var noScenes = Flatten(new OutlineFilters { Scenes = false, Synopses = false, Markers = false }.Apply(analysis.Outline));
        Assert(noScenes.SequenceEqual(new[] { "0:Act:Act One", "1:Sequence:Setup" }), "Only sections remain: " + string.Join(", ", noScenes));

        var cards = analysis.BoardLanes.SelectMany(l => l.Groups).SelectMany(g => g.Cards).ToList();
        Assert(cards.All(c => c.Kind == "Scene"), "Notes never become board cards");

        var restored = JsonSerializer.Deserialize<OutlineFilters>("{\"scenes\":false}", new JsonSerializerOptions(JsonSerializerDefaults.Web))!;
        Assert(!restored.Scenes && restored.Sections && !restored.Notes && restored.Markers, "Missing keys in the stored session fall back to defaults");
    }

    // ---- Beat Sheet (Passage.Web) ----

    static void TestBeatSheetSidecarRoundTrip()
    {
        var (_, store, root) = NewSlateFixture();
        try
        {
            var document = new SlateDocument { BeatSheetPages = 95 };
            var midpoint = new Beat { Title = "Midpoint", Colour = "#f2c46d", Description = "False victory.\nShe wins the wrong fight.", Page = 55, Span = 3 };
            var loose = new Beat { Title = "Maybe a dog?", Page = 0, Span = 2 };
            document.Beats.Add(midpoint);
            document.Beats.Add(loose);
            store.Save("draft", document);

            var loaded = store.Load("draft")!;
            Assert(loaded.BeatSheetPages == 95, "The target length round-trips");
            Assert(loaded.Beats.Count == 2, "Both beats round-trip");
            var first = loaded.Beats[0];
            Assert(first.Id == midpoint.Id && first.Title == "Midpoint" && first.Colour == "#f2c46d"
                && first.Description == midpoint.Description && first.Page == 55 && first.Span == 3,
                "A placed beat keeps its id, title, colour, description, page and span");
            Assert(loaded.Beats[1].Id == loose.Id && loaded.Beats[1].Page == 0 && loaded.Beats[1].Span == 2,
                "An unassigned beat stays in the tray and keeps its span, in list order");

            // A sidecar written before the Beat Sheet existed has neither key.
            File.WriteAllText(Path.Combine(root, ".slate", "old.fountain.json"), "{ \"SlateVersion\": 1, \"IgnoredBrackets\": [\"[x]\"] }");
            var old = store.Load("old")!;
            Assert(old.Beats.Count == 0 && old.BeatSheetPages == BeatSheet.DefaultTargetPages && old.IgnoredBrackets.Count == 1,
                "An older sidecar loads with no beats and the default target, keeping its other data");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void TestBeatSheetClamping()
    {
        Assert(BeatSheet.Columns(21, 110) == 110, "A short script gets the target length of columns");
        Assert(BeatSheet.Columns(130, 110) == 130, "A long script widens the grid past the target");
        Assert(BeatSheet.Columns(0, 0) == 1, "The grid always has at least one column");

        Assert(BeatSheet.Clamp(0, 1, 110) == (1, 1), "No beat starts before page 1");
        Assert(BeatSheet.Clamp(-5, 3, 110) == (1, 3), "A negative start clamps to page 1 and keeps its span");
        Assert(BeatSheet.Clamp(200, 1, 110) == (110, 1), "No beat starts past the last column");
        Assert(BeatSheet.Clamp(3, 0, 110) == (3, 1), "A span is at least one page");
        Assert(BeatSheet.Clamp(108, 10, 110) == (108, 3), "A resize never runs past the last column");
        Assert(BeatSheet.Clamp(110, 4, 110) == (110, 1), "A beat on the last column spans only that column");

        Assert(BeatSheet.ClampMove(110, 3, 110) == (108, 3), "A moved beat dropped at the end slides back and keeps its length");
        Assert(BeatSheet.ClampMove(0, 2, 110) == (1, 2), "A moved beat never starts before page 1");
        Assert(BeatSheet.ClampMove(5, 20, 10) == (1, 10), "A beat longer than the grid is cut to the grid");

        var beats = new List<Beat> { new() { Page = 0, Span = 9 }, new() { Page = 40, Span = 5 }, new() { Page = 10, Span = 1 } };
        Assert(BeatSheet.LastBeatPage(beats) == 44, "The last beat page ignores tray beats and counts the span");
        Assert(BeatSheet.LastBeatPage(new List<Beat>()) == 0, "No placed beats means no last page");
    }

    static void TestBeatSheetStacking()
    {
        var a = new Beat { Title = "a", Page = 3, Span = 2 };   // 3–4
        var b = new Beat { Title = "b", Page = 4, Span = 3 };   // 4–6, overlaps a
        var c = new Beat { Title = "c", Page = 5, Span = 1 };   // 5, free on a's row
        var d = new Beat { Title = "d", Page = 2, Span = 1 };   // 2, earliest page
        var tray = new Beat { Title = "tray", Page = 0 };
        var rows = BeatSheet.Stack(new List<Beat> { a, b, c, tray, d }).ToDictionary(item => item.Beat.Title, item => item.Row);

        Assert(!rows.ContainsKey("tray"), "Tray beats are not on the grid");
        Assert(rows["d"] == 0 && rows["a"] == 0, "Beats that do not overlap share the first row");
        Assert(rows["b"] == 1, "An overlapping beat stacks below");
        Assert(rows["c"] == 0, "A beat goes back up to the first row that is free at its page");

        var same = BeatSheet.Stack(new List<Beat> { new() { Title = "first", Page = 7 }, new() { Title = "second", Page = 7 } });
        Assert(same[0].Beat.Title == "first" && same[0].Row == 0 && same[1].Row == 1, "Beats on the same page stack in list order");
    }

    static void TestInboxConversionLines()
    {
        var id = Guid.NewGuid();

        var titled = BeatInbox.SceneLines(new Beat { Title = "  Dinner goes wrong ", Description = "She leaves.\n\nHe stays." }, id);
        Assert(titled.Count == 3, $"Title plus two description lines, got {titled.Count}");
        Assert(titled[0] == $". Dinner goes wrong [[id:{id}]]", $"Scene heading carries the id, got '{titled[0]}'");
        Assert(titled[1] == "= She leaves." && titled[2] == "= He stays.", "Each non-blank description line becomes a synopsis line");

        var bare = BeatInbox.SceneLines(new Beat { Title = "Dinner" }, id);
        Assert(bare.Count == 1 && bare[0] == $". Dinner [[id:{id}]]", "A beat without a description is one heading line");

        var untitled = BeatInbox.SceneLines(new Beat { Description = "Something." }, id);
        Assert(untitled[0] == $". UNTITLED BEAT [[id:{id}]]", $"An untitled beat still gets a real heading, got '{untitled[0]}'");

        var sluglined = BeatInbox.SceneLines(new Beat { Title = "INT. CELLAR - NIGHT" }, id);
        Assert(sluglined[0] == $"INT. CELLAR - NIGHT [[id:{id}]]", "A title that is already a slugline is not prefixed with a dot");
    }

    static void TestInboxTagTransfer()
    {
        var story = new StoryData();
        var beatId = Guid.NewGuid();
        var cardId = Guid.NewGuid();
        var line = Guid.NewGuid();
        var other = Guid.NewGuid();
        story.Items.Add(new ItemTags { TargetId = beatId, Kind = StoryItems.BeatKind, Label = "old", Key = true, Storylines = { line }, Characters = { "ANNA" }, Locations = { "kitchen" } });
        story.Items.Add(new ItemTags { TargetId = other, Kind = StoryItems.BeatKind, Label = "other", Key = true });

        Assert(BeatInbox.MoveTags(story, beatId, cardId, "NEW"), "A tagged beat reports that its tags moved");
        var moved = story.Items.Single(item => item.TargetId == cardId);
        Assert(moved.Kind == StoryItems.CardKind, "The moved item is re-keyed to kind card");
        Assert(moved.Label == "NEW" && moved.Key && moved.Storylines.Single() == line && moved.Characters.Single() == "ANNA" && moved.Locations.Single() == "kitchen",
            "Star, storyline, character and location survive the move");
        Assert(story.Items.All(item => item.TargetId != beatId), "Nothing is left keyed to the beat");
        Assert(story.Items.Single(item => item.TargetId == other).Kind == StoryItems.BeatKind, "Another beat's item is untouched");
        Assert(!BeatInbox.MoveTags(story, beatId, Guid.NewGuid(), "x"), "A beat with no tags moves nothing");
    }

    static void TestInboxInsertPoints()
    {
        // BoardScript: ACT ONE 0-11 (whole block), Setup 3-11, KITCHEN 5-9, GARDEN 10-11, ACT TWO 12-14.
        var (parsed, lineCount) = ParseBoardScript();
        var scriptLines = BoardScript.Split('\n');
        var cardLines = new List<string> { ". NEW [[id:x]]" };

        int At(int lineIndex, bool after)
        {
            var target = FindElement(parsed, lineIndex);
            Assert(BeatInbox.TryPlan(parsed.Elements, scriptLines, target.Id, after, cardLines, out var insertAt, out _),
                $"Target at line {lineIndex} should be found");
            return insertAt;
        }

        Assert(At(0, false) == 0, "Before an act is the act's heading line");
        Assert(At(0, true) == 12, "After an act is past its whole block, the next act's heading");
        Assert(At(3, false) == 3, "Before a sequence is the sequence's heading line");
        Assert(At(3, true) == 12, "After a sequence is past its scenes");
        Assert(At(5, false) == 5, "Before a scene is its heading line");
        Assert(At(5, true) == 10, "After a scene is past its action, the next scene's heading");
        Assert(At(12, true) == lineCount, "After the last act is the end of the script");

        string TextFor(int lineIndex, bool after)
        {
            BeatInbox.TryPlan(parsed.Elements, scriptLines, FindElement(parsed, lineIndex).Id, after, cardLines, out _, out var text);
            return text;
        }

        // The blocks end in their own blank line, so nothing doubles up.
        Assert(TextFor(5, true) == ". NEW [[id:x]]\n", $"After a scene: no leading blank, trailing blank, got '{TextFor(5, true)}'");
        Assert(TextFor(5, false) == ". NEW [[id:x]]\n", $"Before a scene: the line above is blank, got '{TextFor(5, false)}'");
        // The script ends on text, so a blank line goes above and nothing below.
        Assert(TextFor(12, true) == "\n. NEW [[id:x]]", $"After the last act pads above only, got '{TextFor(12, true)}'");
        Assert(BeatInbox.TextAtEnd(scriptLines, cardLines) == "\n. NEW [[id:x]]", "Send to Board at the end pads like after the last block");
        Assert(TextFor(0, false) == ". NEW [[id:x]]\n", "Before the first line needs no leading blank");
        Assert(BeatInbox.TextAtEnd(new[] { "text", string.Empty }, cardLines) == ". NEW [[id:x]]", "A script already ending in a blank line is not padded again");
    }

    static void TestInboxMissingTarget()
    {
        var (parsed, lineCount) = ParseBoardScript();
        var placed = BeatInbox.TryPlan(parsed.Elements, BoardScript.Split('\n'), Guid.NewGuid(), true, new List<string> { ". X" }, out var insertAt, out var text);
        Assert(!placed, "A target that is not in the script is reported, not placed");
        Assert(insertAt == 0 && text.Length == 0, "Nothing is planned for a missing target");
    }

    static void TestBeatSheetScenePagesMatchPreview()
    {
        // Enough distinct scenes to run several pages, with one heading long
        // enough to wrap and acts so the act breaks have pages to land on.
        var text = new StringBuilder("Title: Beats\n\n");
        var expected = new List<string>();
        for (var scene = 1; scene <= 30; scene++)
        {
            if (scene is 1 or 12 or 24)
            {
                text.Append($"# ACT {scene}\n\n");
            }

            var heading = scene == 7
                ? "INT. OFFICE OF THE VERY LONG NAMED DEPARTMENT OF MUNICIPAL WATER AND SEWERAGE SERVICES - NIGHT"
                : $"INT. ROOM {scene} - DAY";
            expected.Add(heading);
            text.Append(heading).Append("\n\n");
            for (var beat = 0; beat < 2 + scene % 5; beat++)
            {
                text.Append("She crosses the room and looks at the thing on the table for a long while, thinking.\n\n");
                text.Append("MARA\nI don't know what you want from me.\n\n");
            }
        }

        var analysis = new DocumentAnalyzer().Analyze(text.ToString());
        var scenes = BeatSheet.MapScenes(analysis.Elements, analysis.PreviewPages);
        Assert(analysis.PageCount > 5, "The fixture runs several pages");
        Assert(scenes.Count == 30 && scenes.Select(scene => scene.Heading).SequenceEqual(expected), "Every scene is listed in script order");

        // The truth is the preview: the page whose lines carry the heading.
        foreach (var scene in scenes)
        {
            var firstLine = scene.Heading.ToUpperInvariant()[..Math.Min(20, scene.Heading.Length)];
            var page = analysis.PreviewPages.First(candidate => !candidate.IsTitlePage
                && candidate.Lines.Any(line => line.Text.StartsWith(firstLine, StringComparison.Ordinal)
                    && (line.Text.Length >= scene.Heading.Length || scene.Heading.ToUpperInvariant().StartsWith(line.Text, StringComparison.Ordinal))
                    && line.Style == LayoutTextStyle.LeftBold));
            Assert(scene.Page == page.PageNumber, $"{scene.Heading} is on page {page.PageNumber}, mapped to {scene.Page}");
        }

        Assert(scenes.Select(scene => scene.Page).Zip(scenes.Skip(1).Select(scene => scene.Page)).All(pair => pair.First <= pair.Second),
            "Scene pages never go backwards");
        var lines = text.ToString().Split('\n');
        Assert(scenes.All(scene => lines[scene.LineNumber - 1] == scene.Heading), "Each scene's line number points at its heading");

        var acts = BeatSheet.ActBreaks(analysis.Elements, scenes);
        Assert(acts.Count == 3 && acts[0].Page == scenes[0].Page && acts[1].Page == scenes[11].Page && acts[2].Page == scenes[23].Page,
            "Each act breaks at the page of its first scene");

        // A heading the layout cannot place is reported as page 0, not guessed.
        var orphan = new List<ScreenplayElement>(analysis.Elements) { new SceneHeadingElement("EXT. NOWHERE - DAY", "EXT. NOWHERE - DAY", 9999, false) };
        var withOrphan = BeatSheet.MapScenes(orphan, analysis.PreviewPages);
        Assert(withOrphan[^1].Page == 0 && withOrphan.Take(30).Select(scene => scene.Page).SequenceEqual(scenes.Select(scene => scene.Page)),
            "An unplaceable heading gets page 0 and does not disturb the others");
    }

    const string BesideScript = "Title: T\n\n# ACT ONE\n\n= Setup.\n\n## RUN\n\n.FIRST\n\nAction one.\n\n.SECOND\n\nAction two.\n\n.THIRD\n\nAction three.\n\n# ACT TWO\n\n.FOURTH\n\nAction.\n";

    static void TestBesideListOrderMatchesBoard()
    {
        var analysis = new DocumentAnalyzer().Analyze(BesideScript);
        var rows = Passage.Web.Services.BesideList.Rows(analysis.BoardLanes);
        var board = new List<string>();
        foreach (var lane in analysis.BoardLanes)
        {
            board.Add(lane.ActCard?.Heading ?? "Script");
            foreach (var group in lane.Groups)
            {
                if (group.SequenceCard is { } sequence) board.Add(sequence.Heading);
                foreach (var card in group.Cards) board.Add(card.Heading);
            }
        }

        Assert(rows.Count > 4 && rows.Select(row => row.Text).SequenceEqual(board), "The list is the Board's lanes, groups and cards in order");
        var lines = rows.Where(row => row.Line > 0).Select(row => row.Line).ToList();
        Assert(lines.SequenceEqual(lines.OrderBy(line => line)), "Rows are in document order");
        Assert(rows.Count(row => row.Role == "card") >= 4, "The scenes are listed as cards");
    }

    static void TestBesideCaretLookup()
    {
        var analysis = new DocumentAnalyzer().Analyze(BesideScript);
        var rows = Passage.Web.Services.BesideList.Rows(analysis.BoardLanes);
        var script = BesideScript.Split('\n');
        int RowOf(string heading) => rows.FindIndex(row => row.Text == heading);
        int LineOf(string text) => Array.FindIndex(script, line => line.Contains(text)) + 1;

        Assert(Passage.Web.Services.BesideList.RowAt(rows, 1) == -1, "Before any card the lookup finds none");
        var second = RowOf("SECOND");
        Assert(second >= 0, "Fixture has SECOND");
        Assert(Passage.Web.Services.BesideList.RowAt(rows, LineOf(".SECOND")) == second, "On a heading it finds that card");
        Assert(Passage.Web.Services.BesideList.RowAt(rows, LineOf("Action two")) == second, "Inside a scene it finds that scene");
        Assert(Passage.Web.Services.BesideList.RowAt(rows, 9999) == rows.Count - 1, "Past the end it finds the last card");

        var scenes = new List<Passage.Web.Services.ScenePlacement> { new("A", 3, 1), new("B", 10, 4), new("C", 20, 9) };
        var beats = Passage.Web.Services.BesideList.Beats(new List<Passage.Web.Services.Beat>
        {
            new() { Title = "late", Page = 9 }, new() { Title = "loose", Page = 0 }, new() { Title = "mid", Page = 3, Span = 3 }
        }, scenes);
        Assert(beats.Select(item => item.Beat.Title).SequenceEqual(new[] { "loose", "mid", "late" }), "Beats are in page order, unassigned first");
        Assert(beats.Select(item => item.Line).SequenceEqual(new[] { 0, 3, 20 }), "A beat points at the scene on its page, else the one before");
        Assert(Passage.Web.Services.BesideList.BeatAt(beats, scenes, 12) == 1, "A caret in a scene on a spanned page finds the covering beat");
        Assert(Passage.Web.Services.BesideList.BeatAt(beats, scenes, 1) == -1, "Before the first scene no beat is current");
    }
}
