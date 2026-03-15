using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeTalk.Ast;
using D = DocumentFormat.OpenXml.Drawing;

namespace OfficeTalkEngine.Addressing;

/// <summary>
/// Resolves OfficeTalk addresses against PowerPoint (PresentationDocument).
/// Supports slide[n], slide[n]/title, slide[n]/body, slide[n]/notes, slide[n]/shape[name].
/// </summary>
public class PowerPointAddressResolver : IAddressResolver
{
    private readonly PresentationDocument _doc;

    public PowerPointAddressResolver(PresentationDocument doc)
    {
        _doc = doc;
    }

    public IReadOnlyList<OpenXmlElement> Resolve(Address address)
    {
        if (address.Segments.Count == 0)
            throw new InvalidOperationException("Address has no segments.");

        var firstSeg = address.Segments[0];

        if (firstSeg.Identifier.Equals("slide", StringComparison.OrdinalIgnoreCase))
        {
            var slideParts = ResolveSlides(firstSeg);

            if (address.Segments.Count == 1)
            {
                // Return the slide's CommonSlideData for each matched slide
                return slideParts.Select(sp => (OpenXmlElement)sp.Slide).ToList();
            }

            // Resolve further segments within each slide
            var results = new List<OpenXmlElement>();
            foreach (var slidePart in slideParts)
            {
                var innerResults = ResolveWithinSlide(
                    slidePart, address.Segments.Skip(1).ToList());
                results.AddRange(innerResults);
            }
            return results;
        }

        throw new InvalidOperationException(
            $"PowerPoint addresses must start with 'slide', got '{firstSeg.Identifier}'.");
    }

    private IReadOnlyList<SlidePart> ResolveSlides(AddressSegment segment)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation has no presentation part.");
        var presentation = presentationPart.Presentation;
        var slideIdList = presentation.SlideIdList
            ?? throw new InvalidOperationException("Presentation has no slides.");

        var slideIds = slideIdList.Elements<SlideId>().ToList();
        var slideParts = slideIds
            .Select(sid => (SlidePart)presentationPart.GetPartById(sid.RelationshipId!))
            .ToList();

        // Apply predicates
        var indices = Enumerable.Range(0, slideParts.Count).ToList();

        foreach (var pred in segment.Predicates)
        {
            if (pred is PositionalPredicate pos)
            {
                if (pos.Position >= 1 && pos.Position <= slideParts.Count)
                    indices = new List<int> { pos.Position - 1 };
                else
                    indices = new List<int>();
            }
            else if (pred is KeyValuePredicate kv)
            {
                if (kv.Key.Equals("text", StringComparison.OrdinalIgnoreCase))
                {
                    var val = kv.Value?.ToString() ?? "";
                    indices = indices.Where(i =>
                        MatchText(GetSlideText(slideParts[i]), val, kv.Operator)).ToList();
                }
            }
            else if (pred is BareStringPredicate bare)
            {
                indices = indices.Where(i =>
                    GetSlideText(slideParts[i]).Contains(bare.Value)).ToList();
            }
        }

        return indices.Select(i => slideParts[i]).ToList();
    }

    private static string GetSlideText(SlidePart slidePart)
    {
        return slidePart.Slide.CommonSlideData?.ShapeTree?.InnerText ?? "";
    }

    private static IReadOnlyList<OpenXmlElement> ResolveWithinSlide(
        SlidePart slidePart, List<AddressSegment> segments)
    {
        var seg = segments[0];
        var id = seg.Identifier.ToLowerInvariant();

        IReadOnlyList<OpenXmlElement> results;

        switch (id)
        {
            case "title":
                results = FindPlaceholderShape(slidePart, PlaceholderValues.Title,
                    PlaceholderValues.CenteredTitle);
                break;

            case "subtitle":
                results = FindPlaceholderShape(slidePart, PlaceholderValues.SubTitle);
                break;

            case "body":
                results = FindPlaceholderShape(slidePart, PlaceholderValues.Body,
                    PlaceholderValues.Object);
                break;

            case "notes":
                results = ResolveNotes(slidePart);
                break;

            case "shape":
                results = ResolveShapes(slidePart, seg.Predicates);
                break;

            case "table":
                results = ResolveTables(slidePart, seg.Predicates);
                break;

            case "image":
                results = ResolveImages(slidePart, seg.Predicates);
                break;

            default:
                throw new InvalidOperationException(
                    $"Unsupported PowerPoint slide segment '{seg.Identifier}'.");
        }

        // Resolve deeper segments (e.g., table/row/cell)
        if (segments.Count > 1)
        {
            return ResolveDeeper(results, segments.Skip(1).ToList());
        }

        return results;
    }

    private static IReadOnlyList<OpenXmlElement> FindPlaceholderShape(
        SlidePart slidePart, params PlaceholderValues[] types)
    {
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (shapeTree == null) return Array.Empty<OpenXmlElement>();

        foreach (var shape in shapeTree.Elements<Shape>())
        {
            var ph = shape.NonVisualShapeProperties
                ?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            if (ph?.Type?.Value != null && types.Contains(ph.Type.Value))
            {
                return new OpenXmlElement[] { shape };
            }
        }

        // Fallback: if no placeholder type is set, try matching by index
        // (some themes use index-based placeholders without explicit type)
        if (types.Contains(PlaceholderValues.Title))
        {
            foreach (var shape in shapeTree.Elements<Shape>())
            {
                var ph = shape.NonVisualShapeProperties
                    ?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<PlaceholderShape>();
                if (ph != null && ph.Index?.Value == 0)
                    return new OpenXmlElement[] { shape };
            }
        }

        return Array.Empty<OpenXmlElement>();
    }

    private static IReadOnlyList<OpenXmlElement> ResolveNotes(SlidePart slidePart)
    {
        var notesPart = slidePart.NotesSlidePart;
        if (notesPart?.NotesSlide?.CommonSlideData?.ShapeTree == null)
            return Array.Empty<OpenXmlElement>();

        // Find the notes body placeholder
        foreach (var shape in notesPart.NotesSlide.CommonSlideData.ShapeTree.Elements<Shape>())
        {
            var ph = shape.NonVisualShapeProperties
                ?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            if (ph?.Type?.Value == PlaceholderValues.Body)
                return new OpenXmlElement[] { shape };
        }

        return Array.Empty<OpenXmlElement>();
    }

    private static IReadOnlyList<OpenXmlElement> ResolveShapes(
        SlidePart slidePart, List<Predicate> predicates)
    {
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (shapeTree == null) return Array.Empty<OpenXmlElement>();

        var shapes = shapeTree.Elements<Shape>().ToList();
        IEnumerable<Shape> filtered = shapes;

        foreach (var pred in predicates)
        {
            if (pred is PositionalPredicate pos)
            {
                filtered = pos.Position >= 1 && pos.Position <= shapes.Count
                    ? new[] { shapes[pos.Position - 1] }
                    : Enumerable.Empty<Shape>();
            }
            else if (pred is KeyValuePredicate kv)
            {
                if (kv.Key.Equals("name", StringComparison.OrdinalIgnoreCase))
                {
                    var val = kv.Value?.ToString() ?? "";
                    filtered = filtered.Where(s =>
                    {
                        var name = s.NonVisualShapeProperties
                            ?.NonVisualDrawingProperties?.Name?.Value ?? "";
                        return MatchText(name, val, kv.Operator);
                    });
                }
                else if (kv.Key.Equals("text", StringComparison.OrdinalIgnoreCase))
                {
                    var val = kv.Value?.ToString() ?? "";
                    filtered = filtered.Where(s => MatchText(s.InnerText, val, kv.Operator));
                }
            }
            else if (pred is BareStringPredicate bare)
            {
                filtered = filtered.Where(s =>
                {
                    var name = s.NonVisualShapeProperties
                        ?.NonVisualDrawingProperties?.Name?.Value ?? "";
                    return name == bare.Value;
                });
            }
        }

        return filtered.Cast<OpenXmlElement>().ToList();
    }

    private static IReadOnlyList<OpenXmlElement> ResolveTables(
        SlidePart slidePart, List<Predicate> predicates)
    {
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (shapeTree == null) return Array.Empty<OpenXmlElement>();

        // Tables in PowerPoint are inside GraphicFrame elements
        var graphicFrames = shapeTree.Elements<GraphicFrame>().ToList();
        var tables = new List<OpenXmlElement>();

        foreach (var gf in graphicFrames)
        {
            var table = gf.Descendants<D.Table>().FirstOrDefault();
            if (table != null) tables.Add(table);
        }

        // Apply positional predicates
        foreach (var pred in predicates)
        {
            if (pred is PositionalPredicate pos)
            {
                if (pos.Position >= 1 && pos.Position <= tables.Count)
                    tables = new List<OpenXmlElement> { tables[pos.Position - 1] };
                else
                    tables.Clear();
            }
        }

        return tables;
    }

    private static IReadOnlyList<OpenXmlElement> ResolveImages(
        SlidePart slidePart, List<Predicate> predicates)
    {
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (shapeTree == null) return Array.Empty<OpenXmlElement>();

        var pictures = shapeTree.Elements<Picture>().Cast<OpenXmlElement>().ToList();

        foreach (var pred in predicates)
        {
            if (pred is PositionalPredicate pos)
            {
                if (pos.Position >= 1 && pos.Position <= pictures.Count)
                    pictures = new List<OpenXmlElement> { pictures[pos.Position - 1] };
                else
                    pictures.Clear();
            }
        }

        return pictures;
    }

    private static IReadOnlyList<OpenXmlElement> ResolveDeeper(
        IReadOnlyList<OpenXmlElement> parents, List<AddressSegment> segments)
    {
        // For table/row/cell navigation
        var seg = segments[0];
        var id = seg.Identifier.ToLowerInvariant();
        var results = new List<OpenXmlElement>();

        foreach (var parent in parents)
        {
            if (id == "row" && parent is D.Table table)
            {
                var rows = table.Elements<D.TableRow>().Cast<OpenXmlElement>().ToList();
                results.AddRange(ApplyPositionalPredicate(rows, seg.Predicates));
            }
            else if (id == "cell" && parent is D.TableRow row)
            {
                var cells = row.Elements<D.TableCell>().Cast<OpenXmlElement>().ToList();
                results.AddRange(ApplyPositionalPredicate(cells, seg.Predicates));
            }
        }

        if (segments.Count > 1)
            return ResolveDeeper(results, segments.Skip(1).ToList());

        return results;
    }

    private static IReadOnlyList<OpenXmlElement> ApplyPositionalPredicate(
        List<OpenXmlElement> elements, List<Predicate> predicates)
    {
        foreach (var pred in predicates)
        {
            if (pred is PositionalPredicate pos)
            {
                if (pos.Position >= 1 && pos.Position <= elements.Count)
                    return new[] { elements[pos.Position - 1] };
                return Array.Empty<OpenXmlElement>();
            }
        }
        return elements;
    }

    private static bool MatchText(string actual, string expected, PredicateOperator op)
    {
        return op switch
        {
            PredicateOperator.Equals => actual.Equals(expected, StringComparison.Ordinal),
            PredicateOperator.AsteriskEquals => actual.Contains(expected, StringComparison.Ordinal),
            PredicateOperator.CaretEquals => actual.StartsWith(expected, StringComparison.Ordinal),
            PredicateOperator.DollarEquals => actual.EndsWith(expected, StringComparison.Ordinal),
            _ => actual.Equals(expected, StringComparison.Ordinal)
        };
    }
}
