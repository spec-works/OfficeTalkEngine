using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using OfficeTalk.Ast;

namespace OfficeTalkEngine.Execution;

/// <summary>
/// Executes OfficeTalk operations against a live Word instance via COM Interop.
/// Uses dynamic COM to avoid dependency on Office Primary Interop Assemblies.
/// Requires Word to be running with the target document open.
/// Changes appear instantly in the Word UI.
/// </summary>
[SupportedOSPlatform("windows")]
public class WordComExecutor : IOfficeTalkExecutor
{
    // Word WdOutlineLevel constants
    private const int WdOutlineLevelBodyText = 10;
    private const int WdOutlineLevel1 = 1;
    private const int WdOutlineLevel9 = 9;

    // Word WdFindWrap constants
    private const int WdFindStop = 0;

    // Word WdReplace constants
    private const int WdReplaceNone = 0;
    private const int WdReplaceOne = 1;
    private const int WdReplaceAll = 2;

    /// <summary>
    /// Executes operations against a document that is open in a running Word instance.
    /// The outputPath parameter is not supported — COM always modifies the live document.
    /// </summary>
    public void Execute(OfficeTalkDocument document, string targetPath, string? outputPath = null)
    {
        if (outputPath != null)
            throw new NotSupportedException(
                "Output path is not supported when editing via Word COM. The live document is modified in place.");

        dynamic wordApp = GetRunningWordInstance();
        dynamic doc = FindOpenDocument(wordApp, targetPath);

        // Phase 1: Snapshot — resolve all addresses and capture range positions
        var resolvedBlocks = new List<(OperationBlock Block, List<(int Start, int End)> Positions)>();
        foreach (var block in document.OperationBlocks)
        {
            var ranges = ResolveAddress(doc, block.Address);
            var positions = new List<(int Start, int End)>();
            foreach (var r in ranges)
            {
                positions.Add(((int)r.Start, (int)r.End));
            }
            resolvedBlocks.Add((block, positions));
        }

        // Phase 2: Execute operations in reverse order to preserve earlier positions
        for (int i = resolvedBlocks.Count - 1; i >= 0; i--)
        {
            var (block, positions) = resolvedBlocks[i];
            for (int j = positions.Count - 1; j >= 0; j--)
            {
                var (start, end) = positions[j];
                dynamic range = doc.Range(start, end);
                foreach (var operation in block.Operations)
                {
                    ExecuteOperation(doc, range, operation);
                }
            }
        }

        // Apply document-level property settings
        foreach (var prop in document.PropertySettings)
        {
            ApplyProperty(doc, prop.Name, prop.Value);
        }
    }

    /// <summary>
    /// Checks whether Word is running and has the specified document open.
    /// </summary>
    public static bool IsAvailable(string targetPath)
    {
        try
        {
            dynamic wordApp = GetRunningWordInstance();
            FindOpenDocument(wordApp, targetPath);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static object GetRunningWordInstance()
    {
        try
        {
            return ComInteropHelper.GetActiveObject("Word.Application");
        }
        catch (COMException)
        {
            throw new InvalidOperationException(
                "Word is not running. The COM executor requires an active Word instance.");
        }
    }

    private static object FindOpenDocument(dynamic wordApp, string targetPath)
    {
        string fullPath = Path.GetFullPath(targetPath);

        foreach (dynamic doc in wordApp.Documents)
        {
            string docPath = Path.Combine((string)doc.Path, (string)doc.Name);
            if (string.Equals(docPath, fullPath, StringComparison.OrdinalIgnoreCase))
                return doc;
        }

        throw new FileNotFoundException(
            $"Document is not open in Word: {fullPath}");
    }

    #region Address Resolution

    private static List<dynamic> ResolveAddress(dynamic doc, Address address)
    {
        if (address.Segments.Count == 0)
            return new List<dynamic>();

        var segments = address.Segments;
        int startIndex = 0;

        if (segments[0].Identifier.Equals("body", StringComparison.OrdinalIgnoreCase))
        {
            if (segments.Count == 1)
                return new List<dynamic> { doc.Content };
            startIndex = 1;
        }

        var segment = segments[startIndex];
        List<dynamic> current = segment.Identifier.ToLowerInvariant() switch
        {
            "paragraph" => ResolveParagraphs(doc, segment),
            "heading" => ResolveHeadings(doc, segment),
            "table" => ResolveTables(doc, segment),
            _ => throw new NotImplementedException(
                $"Address segment '{segment.Identifier}' is not yet supported for Word COM.")
        };

        for (int i = startIndex + 1; i < segments.Count; i++)
        {
            segment = segments[i];
            current = segment.Identifier.ToLowerInvariant() switch
            {
                "row" => ResolveRows(doc, segment, current),
                "cell" => ResolveCells(doc, segment, current),
                _ => throw new NotImplementedException(
                    $"Address segment '{segment.Identifier}' is not yet supported for Word COM.")
            };
        }

        return current;
    }

    private static List<dynamic> ResolveParagraphs(dynamic doc, AddressSegment segment)
    {
        var paragraphs = new List<dynamic>();
        foreach (dynamic para in doc.Paragraphs)
        {
            if (!IsHeading(para))
                paragraphs.Add(para.Range);
        }

        return ApplyRangePredicates(doc, paragraphs, segment.Predicates);
    }

    private static List<dynamic> ResolveHeadings(dynamic doc, AddressSegment segment)
    {
        var headings = new List<(dynamic Range, int Level)>();
        foreach (dynamic para in doc.Paragraphs)
        {
            if (IsHeading(para))
            {
                int level = GetHeadingLevel(para);
                headings.Add((para.Range, level));
            }
        }

        // Apply level predicate
        var levelPredicate = segment.Predicates
            .OfType<KeyValuePredicate>()
            .FirstOrDefault(p => p.Key.Equals("level", StringComparison.OrdinalIgnoreCase));

        if (levelPredicate != null && int.TryParse(levelPredicate.Value, out int targetLevel))
        {
            headings = headings.Where(h => h.Level == targetLevel).ToList();
        }

        var ranges = headings.Select(h => h.Range).ToList();

        var remainingPredicates = segment.Predicates
            .Where(p => p is not KeyValuePredicate kvp ||
                        !kvp.Key.Equals("level", StringComparison.OrdinalIgnoreCase))
            .ToList();

        return ApplyRangePredicates(doc, ranges, remainingPredicates);
    }

    private static List<dynamic> ResolveTables(dynamic doc, AddressSegment segment)
    {
        var tables = new List<dynamic>();
        foreach (dynamic table in doc.Tables)
        {
            tables.Add(table.Range);
        }

        var positional = segment.Predicates.OfType<PositionalPredicate>().FirstOrDefault();
        if (positional != null)
        {
            int index = positional.Position - 1;
            if (index >= 0 && index < tables.Count)
                return new List<dynamic> { tables[index] };
            return new List<dynamic>();
        }

        return tables;
    }

    private static List<dynamic> ResolveRows(dynamic doc, AddressSegment segment, List<dynamic> tableRanges)
    {
        var rows = new List<dynamic>();
        foreach (var tableRange in tableRanges)
        {
            foreach (dynamic table in doc.Tables)
            {
                if ((int)table.Range.Start == (int)tableRange.Start)
                {
                    foreach (dynamic row in table.Rows)
                    {
                        rows.Add(row.Range);
                    }
                    break;
                }
            }
        }

        var positional = segment.Predicates.OfType<PositionalPredicate>().FirstOrDefault();
        if (positional != null)
        {
            int index = positional.Position - 1;
            if (index >= 0 && index < rows.Count)
                return new List<dynamic> { rows[index] };
            return new List<dynamic>();
        }

        return rows;
    }

    private static List<dynamic> ResolveCells(dynamic doc, AddressSegment segment, List<dynamic> rowRanges)
    {
        var cells = new List<dynamic>();
        foreach (var rowRange in rowRanges)
        {
            foreach (dynamic table in doc.Tables)
            {
                foreach (dynamic row in table.Rows)
                {
                    if ((int)row.Range.Start == (int)rowRange.Start)
                    {
                        foreach (dynamic cell in row.Cells)
                        {
                            cells.Add(cell.Range);
                        }
                    }
                }
            }
        }

        var positional = segment.Predicates.OfType<PositionalPredicate>().FirstOrDefault();
        if (positional != null)
        {
            int index = positional.Position - 1;
            if (index >= 0 && index < cells.Count)
                return new List<dynamic> { cells[index] };
            return new List<dynamic>();
        }

        return cells;
    }

    private static List<dynamic> ApplyRangePredicates(
        dynamic doc, List<dynamic> ranges, List<Predicate> predicates)
    {
        foreach (var predicate in predicates)
        {
            ranges = predicate switch
            {
                PositionalPredicate pos => ApplyPositional(ranges, pos),
                KeyValuePredicate kvp => ApplyKeyValue(ranges, kvp),
                _ => ranges
            };
        }

        return ranges;
    }

    private static List<dynamic> ApplyPositional(List<dynamic> ranges, PositionalPredicate predicate)
    {
        int index = predicate.Position - 1;
        if (index >= 0 && index < ranges.Count)
            return new List<dynamic> { ranges[index] };
        return new List<dynamic>();
    }

    private static List<dynamic> ApplyKeyValue(List<dynamic> ranges, KeyValuePredicate predicate)
    {
        if (predicate.Key.Equals("text", StringComparison.OrdinalIgnoreCase))
        {
            return ranges.Where(r => MatchesText(r, predicate)).ToList();
        }

        if (predicate.Key.Equals("style", StringComparison.OrdinalIgnoreCase))
        {
            return ranges.Where(r =>
            {
                try
                {
                    string styleName = (string)r.get_Style().NameLocal;
                    string styleId = styleName.Replace(" ", "");
                    return string.Equals(styleId, predicate.Value, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(styleName, predicate.Value, StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    return false;
                }
            }).ToList();
        }

        return ranges;
    }

    private static bool MatchesText(dynamic range, KeyValuePredicate predicate)
    {
        string text = ((string)(range.Text ?? "")).TrimEnd('\r', '\a');

        return predicate.Operator switch
        {
            PredicateOperator.Equals => text.Equals(predicate.Value, StringComparison.Ordinal),
            PredicateOperator.AsteriskEquals => text.Contains(predicate.Value, StringComparison.Ordinal),
            PredicateOperator.CaretEquals => text.StartsWith(predicate.Value, StringComparison.Ordinal),
            PredicateOperator.DollarEquals => text.EndsWith(predicate.Value, StringComparison.Ordinal),
            _ => false
        };
    }

    private static bool IsHeading(dynamic paragraph)
    {
        try
        {
            int outlineLevel = (int)paragraph.OutlineLevel;
            return outlineLevel >= WdOutlineLevel1 && outlineLevel <= WdOutlineLevel9;
        }
        catch
        {
            return false;
        }
    }

    private static int GetHeadingLevel(dynamic paragraph)
    {
        try
        {
            int outlineLevel = (int)paragraph.OutlineLevel;
            if (outlineLevel >= WdOutlineLevel1 && outlineLevel <= WdOutlineLevel9)
                return outlineLevel;
        }
        catch
        {
            // Ignore
        }
        return 0;
    }

    #endregion

    #region Operation Execution

    private static void ExecuteOperation(dynamic doc, dynamic range, Operation operation)
    {
        switch (operation)
        {
            case SetOperation set:
                ExecuteSet(range, set);
                break;
            case ReplaceOperation replace:
                ExecuteReplace(range, replace);
                break;
            case DeleteOperation:
                ExecuteDelete(range);
                break;
            case AppendOperation append:
                ExecuteAppend(range, append);
                break;
            case PrependOperation prepend:
                ExecutePrepend(range, prepend);
                break;
            case StyleOperation style:
                ExecuteStyle(range, style);
                break;
            default:
                throw new NotSupportedException(
                    $"Operation type '{operation.GetType().Name}' is not supported by the COM executor.");
        }
    }

    private static void ExecuteSet(dynamic range, SetOperation operation)
    {
        // Duplicate range and trim trailing paragraph mark
        dynamic textRange = range.Duplicate;
        string currentText = (string)(textRange.Text ?? "");
        if (currentText.EndsWith("\r"))
        {
            textRange.End = (int)textRange.End - 1;
        }
        textRange.Text = operation.Content.Text;
    }

    private static void ExecuteReplace(dynamic range, ReplaceOperation operation)
    {
        dynamic find = range.Find;
        find.ClearFormatting();
        find.Replacement.ClearFormatting();
        find.Text = operation.Search;
        find.Replacement.Text = operation.Replacement;
        find.Forward = true;
        find.Wrap = WdFindStop;
        find.MatchCase = true;
        find.MatchWholeWord = false;

        int replaceType = operation.IsAll ? WdReplaceAll : WdReplaceOne;
        find.Execute(Replace: replaceType);
    }

    private static void ExecuteDelete(dynamic range)
    {
        range.Delete();
    }

    private static void ExecuteAppend(dynamic range, AppendOperation operation)
    {
        dynamic insertRange = range.Duplicate;
        string currentText = (string)(insertRange.Text ?? "");
        if (currentText.EndsWith("\r"))
        {
            insertRange.Start = (int)insertRange.End - 1;
            insertRange.End = (int)insertRange.Start;
        }
        else
        {
            insertRange.Start = (int)insertRange.End;
        }
        insertRange.InsertAfter(operation.Content.Text);
    }

    private static void ExecutePrepend(dynamic range, PrependOperation operation)
    {
        dynamic insertRange = range.Duplicate;
        insertRange.End = (int)insertRange.Start;
        insertRange.InsertBefore(operation.Content.Text);
    }

    private static void ExecuteStyle(dynamic range, StyleOperation operation)
    {
        try
        {
            range.set_Style(operation.StyleName);
        }
        catch (COMException)
        {
            try
            {
                range.set_Style(operation.StyleName.Replace(" ", ""));
            }
            catch (COMException ex)
            {
                throw new NotSupportedException(
                    $"Style '{operation.StyleName}' not found in document.", ex);
            }
        }
    }

    #endregion

    #region Document Properties

    private static void ApplyProperty(dynamic doc, string name, string value)
    {
        // WdBuiltInProperty enum values
        const int wdPropertyTitle = 1;
        const int wdPropertySubject = 2;
        const int wdPropertyAuthor = 3;
        const int wdPropertyKeywords = 4;
        const int wdPropertyCategory = 18;

        int propertyIndex = name.ToLowerInvariant() switch
        {
            "title" => wdPropertyTitle,
            "author" => wdPropertyAuthor,
            "subject" => wdPropertySubject,
            "keywords" => wdPropertyKeywords,
            "category" => wdPropertyCategory,
            _ => throw new NotSupportedException($"Document property '{name}' is not supported.")
        };

        dynamic properties = doc.BuiltInDocumentProperties;
        dynamic prop = properties[propertyIndex];
        prop.Value = value;
    }

    #endregion
}
