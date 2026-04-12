using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using Passage.App.ViewModels;

namespace Passage.App.Converters;

public class LaneViewModel
{
    public string Title { get; set; } = string.Empty;
    public OutlineNodeViewModel? SequenceNode { get; set; }
    public IEnumerable<OutlineNodeViewModel> Items { get; set; } = Array.Empty<OutlineNodeViewModel>();
    public bool IsGeneralLane => SequenceNode == null;
}

public class LaneGroupingConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is not IEnumerable<OutlineNodeViewModel> nodes)
        {
            return Array.Empty<LaneViewModel>();
        }

        var lanes = new List<LaneViewModel>();
        var currentGeneralItems = new List<OutlineNodeViewModel>();

        foreach (var node in nodes)
        {
            // A node is a sequence if it's a section at level 2.
            if (node.Kind == OutlineNodeKind.Section && node.SectionLevel == 2)
            {
                if (currentGeneralItems.Count > 0)
                {
                    lanes.Add(new LaneViewModel
                    {
                        Title = "General",
                        Items = currentGeneralItems.ToArray()
                    });
                    currentGeneralItems.Clear();
                }

                lanes.Add(new LaneViewModel
                {
                    Title = node.Text,
                    SequenceNode = node,
                    Items = node.Children
                });
            }
            else
            {
                currentGeneralItems.Add(node);
            }
        }

        if (currentGeneralItems.Count > 0)
        {
            lanes.Add(new LaneViewModel
            {
                Title = "General",
                Items = currentGeneralItems.ToArray()
            });
        }

        return lanes;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
