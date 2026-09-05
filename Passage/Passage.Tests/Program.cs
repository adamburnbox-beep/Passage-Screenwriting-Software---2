using System;
using System.Collections.Generic;
using System.Diagnostics;
using Passage.Parser;
using Passage.Core;

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
