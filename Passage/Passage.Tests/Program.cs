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
