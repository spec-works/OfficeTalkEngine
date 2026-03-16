using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using OfficeTalk.Ast;

namespace OfficeTalkEngine.Execution;

/// <summary>
/// Executes OfficeTalk operations against a live PowerPoint instance via dynamic COM.
/// Uses late-bound COM to avoid dependency on Office PIAs.
/// Requires PowerPoint to be running with the target presentation open.
/// </summary>
[SupportedOSPlatform("windows")]
public class PowerPointComExecutor : IOfficeTalkExecutor
{
    // PowerPoint placeholder type constants (PpPlaceholderType)
    private const int PpPlaceholderTitle = 1;
    private const int PpPlaceholderBody = 2;
    private const int PpPlaceholderCenterTitle = 3;
    private const int PpPlaceholderSubtitle = 4;

    public void Execute(OfficeTalkDocument document, string targetPath, string? outputPath = null)
    {
        if (outputPath != null)
            throw new NotSupportedException(
                "Output path is not supported when editing via PowerPoint COM. The live presentation is modified in place.");

        dynamic pptApp = GetRunningPowerPointInstance();
        dynamic presentation = FindOpenPresentation(pptApp, targetPath);

        foreach (var block in document.OperationBlocks)
        {
            var targets = ResolveAddress(presentation, block.Address);
            foreach (var target in targets)
            {
                foreach (var operation in block.Operations)
                {
                    ExecuteOperation(presentation, target, operation);
                }
            }
        }
    }

    /// <summary>
    /// Checks whether PowerPoint is running and has the specified presentation open.
    /// </summary>
    public static bool IsAvailable(string targetPath)
    {
        try
        {
            dynamic pptApp = GetRunningPowerPointInstance();
            FindOpenPresentation(pptApp, targetPath);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static object GetRunningPowerPointInstance()
    {
        try
        {
            return ComInteropHelper.GetActiveObject("PowerPoint.Application");
        }
        catch (COMException)
        {
            throw new InvalidOperationException(
                "PowerPoint is not running. The COM executor requires an active PowerPoint instance.");
        }
    }

    private static object FindOpenPresentation(dynamic pptApp, string targetPath)
    {
        string fullPath = Path.GetFullPath(targetPath);

        foreach (dynamic pres in pptApp.Presentations)
        {
            if (string.Equals((string)pres.FullName, fullPath, StringComparison.OrdinalIgnoreCase))
                return pres;
        }

        throw new FileNotFoundException(
            $"Presentation is not open in PowerPoint: {fullPath}");
    }

    #region Address Resolution

    private static List<dynamic> ResolveAddress(dynamic presentation, Address address)
    {
        if (address.Segments.Count == 0)
            throw new InvalidOperationException("Address has no segments.");

        var first = address.Segments[0];

        if (!first.Identifier.Equals("slide", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                $"PowerPoint addresses must start with 'slide', got '{first.Identifier}'.");

        var slides = ResolveSlides(presentation, first);
        if (address.Segments.Count == 1)
            return slides;

        var results = new List<dynamic>();
        foreach (var slide in slides)
        {
            var inner = ResolveWithinSlide(slide, address.Segments.Skip(1).ToList());
            results.AddRange(inner);
        }
        return results;
    }

    private static List<dynamic> ResolveSlides(dynamic presentation, AddressSegment segment)
    {
        var slides = new List<dynamic>();
        int count = (int)presentation.Slides.Count;

        for (int i = 1; i <= count; i++)
            slides.Add(presentation.Slides[i]);

        return ApplyPredicates(slides, segment.Predicates,
            getName: s =>
            {
                try
                {
                    // Try to get the slide's title text
                    return (string)s.Shapes.Title.TextFrame.TextRange.Text;
                }
                catch
                {
                    return "";
                }
            });
    }

    private static List<dynamic> ResolveWithinSlide(
        dynamic slide, List<AddressSegment> segments)
    {
        var seg = segments[0];
        var id = seg.Identifier.ToLowerInvariant();

        List<dynamic> results;

        switch (id)
        {
            case "title":
                results = FindPlaceholder(slide, PpPlaceholderTitle, PpPlaceholderCenterTitle);
                break;

            case "subtitle":
                results = FindPlaceholder(slide, PpPlaceholderSubtitle);
                break;

            case "body":
                results = FindPlaceholder(slide, PpPlaceholderBody);
                break;

            case "notes":
                results = ResolveNotes(slide);
                break;

            case "shape":
                results = ResolveShapes(slide, seg.Predicates);
                break;

            default:
                throw new InvalidOperationException(
                    $"Unsupported PowerPoint slide segment '{seg.Identifier}'.");
        }

        return results;
    }

    private static List<dynamic> FindPlaceholder(dynamic slide, params int[] placeholderTypes)
    {
        int shapeCount = (int)slide.Shapes.Count;
        for (int i = 1; i <= shapeCount; i++)
        {
            dynamic shape = slide.Shapes[i];
            try
            {
                int phType = (int)shape.PlaceholderFormat.Type;
                if (placeholderTypes.Contains(phType))
                    return new List<dynamic> { shape };
            }
            catch
            {
                // Shape doesn't have a placeholder format — skip
            }
        }
        return new List<dynamic>();
    }

    private static List<dynamic> ResolveNotes(dynamic slide)
    {
        try
        {
            dynamic notesPage = slide.NotesPage;
            // The notes text is in the second placeholder (body) of the notes page
            int shapeCount = (int)notesPage.Shapes.Count;
            for (int i = 1; i <= shapeCount; i++)
            {
                dynamic shape = notesPage.Shapes[i];
                try
                {
                    int phType = (int)shape.PlaceholderFormat.Type;
                    if (phType == PpPlaceholderBody)
                        return new List<dynamic> { shape };
                }
                catch { }
            }
        }
        catch { }
        return new List<dynamic>();
    }

    private static List<dynamic> ResolveShapes(dynamic slide, List<Predicate> predicates)
    {
        var shapes = new List<dynamic>();
        int count = (int)slide.Shapes.Count;
        for (int i = 1; i <= count; i++)
            shapes.Add(slide.Shapes[i]);

        return ApplyPredicates(shapes, predicates, getName: s => (string)s.Name);
    }

    #endregion

    #region Operation Execution

    #pragma warning disable CA1416
    private static void ExecuteOperation(dynamic presentation, dynamic target, Operation operation)
    {
        switch (operation)
        {
            case SetOperation set:
                ExecuteSet(target, set);
                break;
            case DeleteOperation:
                ExecuteDelete(target);
                break;
            case CommentOperation comment:
                ExecuteComment(target, comment);
                break;
            default:
                throw new NotSupportedException(
                    $"Operation type '{operation.GetType().Name}' is not supported by the PowerPoint COM executor.");
        }
    }
    #pragma warning restore CA1416

    private static void ExecuteSet(dynamic target, SetOperation operation)
    {
        // target is a Shape — set its text
        target.TextFrame.TextRange.Text = operation.Content.Text;
    }

    private static void ExecuteDelete(dynamic target)
    {
        target.Delete();
    }

    private static void ExecuteComment(dynamic target, CommentOperation operation)
    {
        // In PowerPoint COM, comments are added to slides, not shapes.
        // Navigate to the parent slide, then add a comment.
        dynamic slide;
        try
        {
            // If target is a slide
            slide = target.SlideIndex != null ? target : throw new Exception();
        }
        catch
        {
            // If target is a shape, get its parent slide
            try
            {
                slide = target.Parent;
                // Verify it's a slide by accessing SlideIndex
                _ = (int)slide.SlideIndex;
            }
            catch
            {
                throw new InvalidOperationException(
                    "Could not determine the parent slide for the comment target.");
            }
        }

        // Slide.Comments.Add(Left, Top, Width, Height, Text)
        // Position at top-left of slide
        slide.Comments.Add(0, 0, "OfficeTalk", "OT", operation.Content.Text);
    }

    #endregion

    #region Helpers

    private static List<dynamic> ApplyPredicates(
        List<dynamic> elements, List<Predicate> predicates, Func<dynamic, string> getName)
    {
        IEnumerable<dynamic> filtered = elements;

        foreach (var pred in predicates)
        {
            if (pred is PositionalPredicate pos)
            {
                if (pos.Position >= 1 && pos.Position <= elements.Count)
                    filtered = new[] { elements[pos.Position - 1] };
                else
                    filtered = Enumerable.Empty<dynamic>();
            }
            else if (pred is BareStringPredicate bare)
            {
                filtered = filtered.Where(e => getName(e) == bare.Value);
            }
            else if (pred is KeyValuePredicate kv)
            {
                var val = kv.Value?.ToString() ?? "";
                filtered = filtered.Where(e => MatchText(getName(e), val, kv.Operator));
            }
        }

        return filtered.ToList();
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

    #endregion
}
