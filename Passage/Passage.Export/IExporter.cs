using Passage.Parser;
using Passage.Core.Extensibility;

namespace Passage.Export;

public interface IExporter : INamedAction
{
    string DefaultExtension { get; }

    void Export(ParsedScreenplay screenplay, string filePath);
}
