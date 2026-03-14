using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeTalk.Ast;

namespace OfficeTalkEngine.Addressing;

/// <summary>
/// Resolves OfficeTalk addresses against a WordprocessingDocument.
/// Supports paragraph, heading, table, row, and cell segments with
/// positional and key-value predicates.
/// </summary>
public class WordAddressResolver : IAddressResolver
{
    private readonly WordprocessingDocument _document;

    public WordAddressResolver(WordprocessingDocument document)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
    }

    public IReadOnlyList<OpenXmlElement> Resolve(Address address)
    {
        if (address.Segments.Count == 0)
            return Array.Empty<OpenXmlElement>();

        IReadOnlyList<OpenXmlElement> current = _document.MainDocumentPart?.Document?.Body?.ChildElements
            .OfType<OpenXmlElement>()
            .ToList() ?? new List<OpenXmlElement>();

        for (int i = 0; i < address.Segments.Count; i++)
        {
            var segment = address.Segments[i];
            current = ResolveSegment(segment, current, isRoot: i == 0);

            if (current.Count == 0)
                return current;
        }

        return current;
    }

    private IReadOnlyList<OpenXmlElement> ResolveSegment(
        AddressSegment segment, IReadOnlyList<OpenXmlElement> context, bool isRoot)
    {
        return segment.Identifier.ToLowerInvariant() switch
        {
            "paragraph" => ResolveParagraphs(segment, context, isRoot),
            "heading" => ResolveHeadings(segment, context, isRoot),
            "table" => ResolveTables(segment, context, isRoot),
            "row" => ResolveRows(segment, context),
            "cell" => ResolveCells(segment, context),
            _ => throw new NotImplementedException(
                $"Address segment '{segment.Identifier}' is not yet supported for Word documents.")
        };
    }

    private IReadOnlyList<OpenXmlElement> ResolveParagraphs(
        AddressSegment segment, IReadOnlyList<OpenXmlElement> context, bool isRoot)
    {
        var paragraphs = isRoot
            ? context.OfType<Paragraph>().ToList()
            : context.SelectMany(e => e.Descendants<Paragraph>()).ToList();

        // Exclude paragraphs that are headings (have an OutlineLevel)
        paragraphs = paragraphs
            .Where(p => !IsHeading(p))
            .ToList();

        return ApplyPredicates(paragraphs.Cast<OpenXmlElement>().ToList(), segment.Predicates);
    }

    private IReadOnlyList<OpenXmlElement> ResolveHeadings(
        AddressSegment segment, IReadOnlyList<OpenXmlElement> context, bool isRoot)
    {
        var paragraphs = isRoot
            ? context.OfType<Paragraph>().ToList()
            : context.SelectMany(e => e.Descendants<Paragraph>()).ToList();

        var headings = paragraphs.Where(p => IsHeading(p)).ToList();

        // Apply level predicate first if present
        var levelPredicate = segment.Predicates
            .OfType<KeyValuePredicate>()
            .FirstOrDefault(p => p.Key.Equals("level", StringComparison.OrdinalIgnoreCase));

        if (levelPredicate != null)
        {
            if (int.TryParse(levelPredicate.Value, out int level))
            {
                headings = headings.Where(p => GetHeadingLevel(p) == level).ToList();
            }
        }

        // Apply remaining predicates (text, positional)
        var remainingPredicates = segment.Predicates
            .Where(p => p is not KeyValuePredicate kvp ||
                        !kvp.Key.Equals("level", StringComparison.OrdinalIgnoreCase))
            .ToList();

        return ApplyPredicates(headings.Cast<OpenXmlElement>().ToList(), remainingPredicates);
    }

    private IReadOnlyList<OpenXmlElement> ResolveTables(
        AddressSegment segment, IReadOnlyList<OpenXmlElement> context, bool isRoot)
    {
        var tables = isRoot
            ? context.OfType<Table>().ToList()
            : context.SelectMany(e => e.Descendants<Table>()).ToList();

        return ApplyPredicates(tables.Cast<OpenXmlElement>().ToList(), segment.Predicates);
    }

    private IReadOnlyList<OpenXmlElement> ResolveRows(
        AddressSegment segment, IReadOnlyList<OpenXmlElement> context)
    {
        var rows = context.SelectMany(e => e is Table t
            ? t.Elements<TableRow>()
            : e.Descendants<TableRow>()).ToList();

        return ApplyPredicates(rows.Cast<OpenXmlElement>().ToList(), segment.Predicates);
    }

    private IReadOnlyList<OpenXmlElement> ResolveCells(
        AddressSegment segment, IReadOnlyList<OpenXmlElement> context)
    {
        var cells = context.SelectMany(e => e is TableRow tr
            ? tr.Elements<TableCell>()
            : e.Descendants<TableCell>()).ToList();

        return ApplyPredicates(cells.Cast<OpenXmlElement>().ToList(), segment.Predicates);
    }

    private static IReadOnlyList<OpenXmlElement> ApplyPredicates(
        List<OpenXmlElement> elements, List<Predicate> predicates)
    {
        var result = elements;

        foreach (var predicate in predicates)
        {
            result = predicate switch
            {
                PositionalPredicate pos => ApplyPositional(result, pos),
                KeyValuePredicate kvp => ApplyKeyValue(result, kvp),
                _ => result
            };
        }

        return result;
    }

    private static List<OpenXmlElement> ApplyPositional(
        List<OpenXmlElement> elements, PositionalPredicate predicate)
    {
        // 1-based indexing
        int index = predicate.Position - 1;
        if (index >= 0 && index < elements.Count)
            return new List<OpenXmlElement> { elements[index] };

        return new List<OpenXmlElement>();
    }

    private static List<OpenXmlElement> ApplyKeyValue(
        List<OpenXmlElement> elements, KeyValuePredicate predicate)
    {
        if (predicate.Key.Equals("text", StringComparison.OrdinalIgnoreCase))
        {
            return elements.Where(e => MatchesText(e, predicate)).ToList();
        }

        return elements;
    }

    private static bool MatchesText(OpenXmlElement element, KeyValuePredicate predicate)
    {
        string text = GetElementText(element);

        return predicate.Operator switch
        {
            PredicateOperator.Equals => text.Equals(predicate.Value, StringComparison.Ordinal),
            PredicateOperator.AsteriskEquals => text.Contains(predicate.Value, StringComparison.Ordinal),
            PredicateOperator.CaretEquals => text.StartsWith(predicate.Value, StringComparison.Ordinal),
            PredicateOperator.DollarEquals => text.EndsWith(predicate.Value, StringComparison.Ordinal),
            _ => false
        };
    }

    private static string GetElementText(OpenXmlElement element)
    {
        return element.InnerText;
    }

    private static bool IsHeading(Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId != null && styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
            return true;

        var outlineLevel = paragraph.ParagraphProperties?.OutlineLevel?.Val;
        return outlineLevel != null;
    }

    private static int GetHeadingLevel(Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId != null && styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
        {
            if (int.TryParse(styleId.AsSpan("Heading".Length), out int level))
                return level;
        }

        var outlineLevel = paragraph.ParagraphProperties?.OutlineLevel?.Val;
        if (outlineLevel != null)
            return outlineLevel.Value + 1; // OutlineLevel is 0-based

        return 0;
    }
}
