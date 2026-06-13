using System.Collections.Generic;
using Avalonia.Media;

namespace Passage.App.ViewModels;

public class PreviewLineViewModel
{
    public string Text { get; set; } = string.Empty;
    public double X { get; set; }
    public double Y { get; set; }
    public bool IsBold { get; set; }

    public FontWeight Weight => IsBold ? FontWeight.Bold : FontWeight.Normal;
}

public class PreviewPageViewModel
{
    public List<PreviewLineViewModel> Lines { get; set; } = new();
    public string PageNumberLabel { get; set; } = string.Empty;
}
