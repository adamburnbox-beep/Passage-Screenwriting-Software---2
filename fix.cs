using System.IO;
using System.Text.RegularExpressions;

class Program {
    static void Main() {
        var path = @"Passage\Passage.App\MainWindow.xaml";
        var content = File.ReadAllText(path);
        
        // Remove Visibility="Collapsed" from InlineTextBox elements
        content = Regex.Replace(content, @"(<TextBox x:Name="Inline(?:Heading|Description|SceneHeading)TextBox"\s+)Visibility="Collapsed"\s*", "$1");
        
        File.WriteAllText(path, content);
    }
}
