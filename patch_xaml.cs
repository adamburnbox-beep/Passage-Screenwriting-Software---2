using System;
using System.IO;
using System.Text.RegularExpressions;

var path = @"Passage\Passage.App\MainWindow.xaml";
var content = File.ReadAllText(path);

// Pattern to match the DescriptionTextBlock block
var pattern = @"<TextBlock x:Name=""DescriptionTextBlock""\s+Text=""{Binding BoardDescription}"">[\s\S]*?</TextBlock>";

var replacement = @"<Grid Margin=""0,10,0,0"">
              <TextBlock x:Name=""DescriptionTextBlock""
                         Text=""{Binding BoardDescription}""
                         Style=""{StaticResource BeatBoardDescriptionStyle}""
                         Margin=""0"" />
              <TextBlock x:Name=""DescriptionPlaceholder""
                         Text=""Click the pencil to edit. Double-click to locate in script.""
                         Style=""{StaticResource BeatBoardDescriptionStyle}""
                         Margin=""0""
                         Opacity=""0.5""
                         FontStyle=""Italic""
                         Visibility=""Collapsed"" />
            </Grid>";

var newContent = Regex.Replace(content, pattern, replacement);

// Also fix the DataTemplate.Triggers
var triggerPattern = @"<Setter TargetName=""DescriptionTextBlock"" Property=""Visibility"" Value=""Collapsed"" />";
var triggerReplacement = @"<Setter TargetName=""DescriptionTextBlock"" Property=""Visibility"" Value=""Collapsed"" />
          <Setter TargetName=""DescriptionPlaceholder"" Property=""Visibility"" Value=""Collapsed"" />";

newContent = newContent.Replace(triggerPattern, triggerReplacement);

File.WriteAllText(path, newContent);
Console.WriteLine("Successfully patched MainWindow.xaml");
