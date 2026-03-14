using DocumentFormat.OpenXml;
using OfficeTalk.Ast;

namespace OfficeTalkEngine.Addressing;

/// <summary>
/// Resolves OfficeTalk addresses against a specific document format,
/// returning the matched OpenXML elements.
/// </summary>
public interface IAddressResolver
{
    IReadOnlyList<OpenXmlElement> Resolve(Address address);
}
