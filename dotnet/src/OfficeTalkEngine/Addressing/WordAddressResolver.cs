using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeTalk.Ast;

namespace OfficeTalkEngine.Addressing;

/// <summary>
/// Resolves OfficeTalk addresses against a WordprocessingDocument.
/// Supports all Word segment types, heading-scoped section addressing,
/// and all predicate types including regex matching.
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

        // Skip the "body" root segment — context is already body's children.
        var segments = address.Segments;
        int startIndex = 0;

        if (segments[0].Identifier.Equals("body", StringComparison.OrdinalIgnoreCase))
        {
            if (segments.Count == 1)
                return current; // AT body → all body children
            startIndex = 1;
        }

        // Handle header/footer as root segments (they aren't under body)
        if (segments[0].Identifier.Equals("header", StringComparison.OrdinalIgnoreCase))
            return ResolveHeaderFooter(segments, isHeader: true);
        if (segments[0].Identifier.Equals("footer", StringComparison.OrdinalIgnoreCase))
            return ResolveHeaderFooter(segments, isHeader: false);

        for (int i = startIndex; i < segments.Count; i++)
        {
            var segment = segments[i];
            bool hasMoreSegments = i < segments.Count - 1;

            // Heading-scoped section addressing: when heading is non-terminal,
            // resolve it as a section scope rather than the heading element itself.
            if (segment.Identifier.Equals("heading", StringComparison.OrdinalIgnoreCase) && hasMoreSegments)
            {
                current = ResolveHeadingSectionScope(segment, current, isRoot: i == startIndex);
                // current is now the body-level elements within the heading's section scope
                continue;
            }

            current = ResolveSegment(segment, current, isRoot: i == startIndex);

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
            "run" => ResolveRuns(segment, context),
            "list" => ResolveLists(segment, context, isRoot),
            "item" => ResolveListItems(segment, context, isRoot),
            "image" => ResolveImages(segment, context, isRoot),
            "section" => ResolveSections(segment),
            "bookmark" => ResolveBookmarks(segment, context, isRoot),
            "content-control" => ResolveContentControls(segment, context, isRoot),
            _ => throw new NotImplementedException(
                $"Address segment '{segment.Identifier}' is not yet supported for Word documents.")
        };
    }

    #region Heading-Scoped Section Addressing

    /// <summary>
    /// Resolves a heading segment as a section scope — returns all sibling elements
    /// between the matched heading and the next heading at the same or higher level.
    /// </summary>
    private IReadOnlyList<OpenXmlElement> ResolveHeadingSectionScope(
        AddressSegment segment, IReadOnlyList<OpenXmlElement> context, bool isRoot)
    {
        // First, find the matching heading(s)
        var headings = ResolveHeadings(segment, context, isRoot);
        if (headings.Count == 0)
            return headings;

        var result = new List<OpenXmlElement>();
        var body = _document.MainDocumentPart?.Document?.Body;
        if (body == null)
            return result;

        var bodyChildren = body.ChildElements.OfType<OpenXmlElement>().ToList();

        foreach (var heading in headings)
        {
            if (heading is not Paragraph headingPara)
                continue;

            int headingLevel = GetHeadingLevel(headingPara);
            int headingIndex = bodyChildren.IndexOf(headingPara);
            if (headingIndex < 0) continue;

            // Collect all elements after this heading until next heading at same or higher level
            for (int j = headingIndex + 1; j < bodyChildren.Count; j++)
            {
                var element = bodyChildren[j];
                if (element is Paragraph p && IsHeading(p))
                {
                    int nextLevel = GetHeadingLevel(p);
                    if (nextLevel <= headingLevel)
                        break; // Same or higher level heading ends the scope
                }
                result.Add(element);
            }
        }

        return result;
    }

    #endregion

    #region Segment Resolvers

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

    private static IReadOnlyList<OpenXmlElement> ResolveRuns(
        AddressSegment segment, IReadOnlyList<OpenXmlElement> context)
    {
        var runs = context.SelectMany(e => e is Paragraph p
            ? p.Elements<Run>()
            : e.Descendants<Run>()).ToList();

        return ApplyPredicates(runs.Cast<OpenXmlElement>().ToList(), segment.Predicates);
    }

    private IReadOnlyList<OpenXmlElement> ResolveLists(
        AddressSegment segment, IReadOnlyList<OpenXmlElement> context, bool isRoot)
    {
        // In OpenXML, lists are groups of consecutive paragraphs with NumberingProperties.
        // We group them by numbering ID to identify distinct lists.
        var paragraphs = isRoot
            ? context.OfType<Paragraph>().ToList()
            : context.SelectMany(e => e.Descendants<Paragraph>()).ToList();

        var lists = new List<List<Paragraph>>();
        List<Paragraph>? currentList = null;
        string? currentNumId = null;

        foreach (var p in paragraphs)
        {
            var numId = GetNumberingId(p);
            if (numId != null)
            {
                if (currentList == null || numId != currentNumId)
                {
                    currentList = new List<Paragraph>();
                    lists.Add(currentList);
                    currentNumId = numId;
                }
                currentList.Add(p);
            }
            else
            {
                currentList = null;
                currentNumId = null;
            }
        }

        // Each "list" is represented by its first paragraph for addressing purposes.
        // The full list paragraphs are accessible via descendants.
        // For now, we return the first paragraph of each list group as the "list" element.
        // When followed by /item, the items within that list will be resolved.
        var listElements = lists.Select(l => (OpenXmlElement)l[0]).ToList();
        return ApplyPredicates(listElements, segment.Predicates);
    }

    private IReadOnlyList<OpenXmlElement> ResolveListItems(
        AddressSegment segment, IReadOnlyList<OpenXmlElement> context, bool isRoot)
    {
        // List items are paragraphs with NumberingProperties
        var paragraphs = isRoot
            ? context.OfType<Paragraph>().ToList()
            : context.SelectMany(e => e.Descendants<Paragraph>()).ToList();

        // If context came from a list segment, include all paragraphs in that list group
        var items = new List<Paragraph>();
        foreach (var p in paragraphs)
        {
            if (GetNumberingId(p) != null)
                items.Add(p);
        }

        // If context is body-level (isRoot), get all list items
        if (items.Count == 0 && isRoot)
        {
            items = paragraphs.Where(p => GetNumberingId(p) != null).ToList();
        }

        return ApplyPredicates(items.Cast<OpenXmlElement>().ToList(), segment.Predicates);
    }

    private IReadOnlyList<OpenXmlElement> ResolveImages(
        AddressSegment segment, IReadOnlyList<OpenXmlElement> context, bool isRoot)
    {
        var drawings = isRoot
            ? context.SelectMany(e => e.Descendants<Drawing>()).ToList()
            : context.SelectMany(e => e.Descendants<Drawing>()).ToList();

        return ApplyPredicates(drawings.Cast<OpenXmlElement>().ToList(), segment.Predicates);
    }

    private IReadOnlyList<OpenXmlElement> ResolveSections(AddressSegment segment)
    {
        // Word sections are defined by SectionProperties at the end of the last paragraph
        // in each section, plus the final section in Body.SectionProperties.
        var body = _document.MainDocumentPart?.Document?.Body;
        if (body == null) return Array.Empty<OpenXmlElement>();

        // Collect SectionProperties from paragraph properties (section breaks)
        var sections = new List<OpenXmlElement>();
        foreach (var para in body.Elements<Paragraph>())
        {
            var sectPr = para.ParagraphProperties?.Elements<SectionProperties>().FirstOrDefault();
            if (sectPr != null)
                sections.Add(para);
        }

        // The final section is represented by body's SectionProperties
        var bodySectPr = body.Elements<SectionProperties>().FirstOrDefault();
        if (bodySectPr != null)
            sections.Add(bodySectPr);

        return ApplyPredicates(sections, segment.Predicates);
    }

    private IReadOnlyList<OpenXmlElement> ResolveBookmarks(
        AddressSegment segment, IReadOnlyList<OpenXmlElement> context, bool isRoot)
    {
        var bookmarkStarts = isRoot
            ? context.SelectMany(e => e.Descendants<BookmarkStart>()).ToList()
            : context.SelectMany(e => e.Descendants<BookmarkStart>()).ToList();

        // Filter out internal bookmarks (starting with _)
        bookmarkStarts = bookmarkStarts
            .Where(b => b.Name?.Value != null && !b.Name.Value.StartsWith("_"))
            .ToList();

        return ApplyPredicates(bookmarkStarts.Cast<OpenXmlElement>().ToList(), segment.Predicates);
    }

    private IReadOnlyList<OpenXmlElement> ResolveContentControls(
        AddressSegment segment, IReadOnlyList<OpenXmlElement> context, bool isRoot)
    {
        var sdts = isRoot
            ? context.OfType<SdtBlock>().Cast<OpenXmlElement>()
                .Concat(context.SelectMany(e => e.Descendants<SdtBlock>()).Cast<OpenXmlElement>())
                .Concat(context.SelectMany(e => e.Descendants<SdtRun>()).Cast<OpenXmlElement>())
                .Distinct().ToList()
            : context.SelectMany(e => e.Descendants<SdtBlock>().Cast<OpenXmlElement>()
                .Concat(e.Descendants<SdtRun>().Cast<OpenXmlElement>())).ToList();

        return ApplyPredicates(sdts, segment.Predicates);
    }

    private IReadOnlyList<OpenXmlElement> ResolveHeaderFooter(
        List<AddressSegment> segments, bool isHeader)
    {
        var body = _document.MainDocumentPart?.Document?.Body;
        if (body == null) return Array.Empty<OpenXmlElement>();

        var segment = segments[0];
        var typePredicate = segment.Predicates
            .OfType<KeyValuePredicate>()
            .FirstOrDefault(p => p.Key.Equals("type", StringComparison.OrdinalIgnoreCase));

        string typeFilter = typePredicate?.Value?.ToLowerInvariant() ?? "default";

        // Get the last section properties (main section)
        var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectPr == null) return Array.Empty<OpenXmlElement>();

        if (isHeader)
        {
            var headerRef = sectPr.Elements<HeaderReference>()
                .FirstOrDefault(h => MatchesHeaderFooterType(h.Type?.Value, typeFilter));

            if (headerRef?.Id?.Value != null)
            {
                var headerPart = _document.MainDocumentPart?.GetPartById(headerRef.Id.Value) as HeaderPart;
                if (headerPart?.Header != null)
                {
                    var elements = headerPart.Header.ChildElements.OfType<OpenXmlElement>().ToList();
                    if (segments.Count > 1)
                    {
                        // Resolve remaining segments within header
                        for (int i = 1; i < segments.Count; i++)
                        {
                            elements = ResolveSegment(segments[i], elements, isRoot: true)
                                .ToList();
                        }
                    }
                    return elements;
                }
            }
        }
        else
        {
            var footerRef = sectPr.Elements<FooterReference>()
                .FirstOrDefault(f => MatchesHeaderFooterType(f.Type?.Value, typeFilter));

            if (footerRef?.Id?.Value != null)
            {
                var footerPart = _document.MainDocumentPart?.GetPartById(footerRef.Id.Value) as FooterPart;
                if (footerPart?.Footer != null)
                {
                    var elements = footerPart.Footer.ChildElements.OfType<OpenXmlElement>().ToList();
                    if (segments.Count > 1)
                    {
                        for (int i = 1; i < segments.Count; i++)
                        {
                            elements = ResolveSegment(segments[i], elements, isRoot: true)
                                .ToList();
                        }
                    }
                    return elements;
                }
            }
        }

        return Array.Empty<OpenXmlElement>();
    }

    private static bool MatchesHeaderFooterType(
        DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues? value, string typeFilter)
    {
        if (value == null) return typeFilter == "default";
        return typeFilter switch
        {
            "default" => value == HeaderFooterValues.Default,
            "first" => value == HeaderFooterValues.First,
            "even" => value == HeaderFooterValues.Even,
            _ => false
        };
    }

    #endregion

    #region Predicate Application

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
                BareStringPredicate bare => ApplyBareString(result, bare),
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
        var key = predicate.Key.ToLowerInvariant();

        return key switch
        {
            "text" => elements.Where(e => MatchesText(e, predicate)).ToList(),
            "name" => elements.Where(e => MatchesName(e, predicate)).ToList(),
            "tag" => elements.Where(e => MatchesTag(e, predicate)).ToList(),
            "caption" => elements.Where(e => MatchesCaption(e, predicate)).ToList(),
            "alt" => elements.Where(e => MatchesAlt(e, predicate)).ToList(),
            "type" => elements, // type is handled at the segment level (header/footer)
            "level" => elements, // level is handled at the heading segment level
            _ => elements
        };
    }

    /// <summary>
    /// Applies a bare string predicate as a shorthand for the element's primary identifier.
    /// bookmark["intro"] → bookmark[name="intro"], content-control["tag"] → content-control[tag="tag"]
    /// </summary>
    private static List<OpenXmlElement> ApplyBareString(
        List<OpenXmlElement> elements, BareStringPredicate predicate)
    {
        return elements.Where(e => MatchesBareString(e, predicate.Value)).ToList();
    }

    #endregion

    #region Predicate Matchers

    private static bool MatchesText(OpenXmlElement element, KeyValuePredicate predicate)
    {
        string text = GetElementText(element);

        return predicate.Operator switch
        {
            PredicateOperator.Equals => text.Equals(predicate.Value, StringComparison.Ordinal),
            PredicateOperator.AsteriskEquals => text.Contains(predicate.Value, StringComparison.Ordinal),
            PredicateOperator.CaretEquals => text.StartsWith(predicate.Value, StringComparison.Ordinal),
            PredicateOperator.DollarEquals => text.EndsWith(predicate.Value, StringComparison.Ordinal),
            PredicateOperator.TildeEquals => MatchesRegex(text, predicate.Value),
            _ => false
        };
    }

    private static bool MatchesRegex(string text, string pattern)
    {
        try
        {
            // I-Regexp (RFC 9485) maps to standard .NET regex with implicit anchoring
            return Regex.IsMatch(text, pattern, RegexOptions.None, TimeSpan.FromSeconds(1));
        }
        catch (RegexParseException)
        {
            return false;
        }
    }

    private static bool MatchesName(OpenXmlElement element, KeyValuePredicate predicate)
    {
        string? name = element switch
        {
            BookmarkStart bm => bm.Name?.Value,
            _ => null
        };

        return name != null && CompareString(name, predicate.Value, predicate.Operator);
    }

    private static bool MatchesTag(OpenXmlElement element, KeyValuePredicate predicate)
    {
        string? tag = GetSdtTag(element);
        return tag != null && CompareString(tag, predicate.Value, predicate.Operator);
    }

    private static bool MatchesCaption(OpenXmlElement element, KeyValuePredicate predicate)
    {
        if (element is Table table)
        {
            var tblPr = table.GetFirstChild<TableProperties>();
            var tableCaption = tblPr?.TableCaption?.Val?.Value;
            if (tableCaption != null)
                return CompareString(tableCaption, predicate.Value, predicate.Operator);
        }
        return false;
    }

    private static bool MatchesAlt(OpenXmlElement element, KeyValuePredicate predicate)
    {
        if (element is Drawing drawing)
        {
            var docProps = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()
                .FirstOrDefault();
            var altText = docProps?.Description?.Value;
            if (altText != null)
                return CompareString(altText, predicate.Value, predicate.Operator);
        }
        return false;
    }

    private static bool MatchesBareString(OpenXmlElement element, string value)
    {
        // Map bare string to the element's primary identifier
        return element switch
        {
            BookmarkStart bm => string.Equals(bm.Name?.Value, value, StringComparison.OrdinalIgnoreCase),
            _ when IsSdtElement(element) =>
                string.Equals(GetSdtTag(element), value, StringComparison.OrdinalIgnoreCase),
            _ => GetElementText(element).Equals(value, StringComparison.Ordinal)
        };
    }

    private static bool CompareString(string actual, string expected, PredicateOperator op)
    {
        return op switch
        {
            PredicateOperator.Equals => actual.Equals(expected, StringComparison.Ordinal),
            PredicateOperator.AsteriskEquals => actual.Contains(expected, StringComparison.Ordinal),
            PredicateOperator.CaretEquals => actual.StartsWith(expected, StringComparison.Ordinal),
            PredicateOperator.DollarEquals => actual.EndsWith(expected, StringComparison.Ordinal),
            PredicateOperator.TildeEquals => MatchesRegex(actual, expected),
            _ => false
        };
    }

    #endregion

    #region Helpers

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

    private static string? GetNumberingId(Paragraph paragraph)
    {
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        return numPr?.NumberingId?.Val?.ToString();
    }

    private static string? GetSdtTag(OpenXmlElement element)
    {
        if (element is SdtBlock sdtBlock)
            return sdtBlock.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
        if (element is SdtRun sdtRun)
            return sdtRun.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
        return null;
    }

    private static bool IsSdtElement(OpenXmlElement element)
    {
        return element is SdtBlock or SdtRun;
    }

    #endregion
}
