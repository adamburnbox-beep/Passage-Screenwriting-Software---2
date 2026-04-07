using System.Windows;
using System.Windows.Controls;
using Passage.Parser;

namespace Passage.App.Views;

public sealed class BeatBoardCardTemplateSelector : DataTemplateSelector
{
    public DataTemplate? ActTemplate { get; set; }

    public DataTemplate? SequenceTemplate { get; set; }

    public DataTemplate? SceneTemplate { get; set; }

    public DataTemplate? NoteTemplate { get; set; }

    public DataTemplate? DefaultTemplate { get; set; }

    public override DataTemplate? SelectTemplate(object item, DependencyObject container)
    {
        if (item is not ScreenplayElement element)
        {
            return base.SelectTemplate(item, container);
        }

        if (element.Type == ScreenplayElementType.Note)
        {
            return NoteTemplate ?? DefaultTemplate ?? base.SelectTemplate(item, container);
        }

        if (element.Type == ScreenplayElementType.Section && element.Level == 0)
        {
            return ActTemplate ?? DefaultTemplate ?? base.SelectTemplate(item, container);
        }

        if (element.Type == ScreenplayElementType.Section && element.Level == 1)
        {
            return SequenceTemplate ?? DefaultTemplate ?? base.SelectTemplate(item, container);
        }

        if (element.Type == ScreenplayElementType.SceneHeading || element.Level == 2)
        {
            return SceneTemplate ?? DefaultTemplate ?? base.SelectTemplate(item, container);
        }

        return DefaultTemplate ?? SceneTemplate ?? base.SelectTemplate(item, container);
    }
}
