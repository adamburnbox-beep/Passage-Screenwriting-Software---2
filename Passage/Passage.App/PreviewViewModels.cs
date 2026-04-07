using System.Collections.Generic;

namespace Passage.App;

public class PreviewLineViewModel
{
    public string Text { get; set; } = string.Empty;
    public double X { get; set; }
    public double Y { get; set; }
    public bool IsBold { get; set; }
}

public class PreviewPageViewModel
{
    public List<PreviewLineViewModel> Lines { get; set; } = new();
    public string PageNumberLabel { get; set; } = string.Empty;
}
