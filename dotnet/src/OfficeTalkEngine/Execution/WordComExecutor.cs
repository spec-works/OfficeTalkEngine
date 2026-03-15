using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text.RegularExpressions;
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

    // Word WdUnits constants
    private const int WdCharacter = 1;
    private const int WdParagraph = 4;

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

        bool hasStructuralOps = document.OperationBlocks.Any(b =>
            b.Operations.Any(op => op is InsertBeforeOperation or InsertAfterOperation));

        if (hasStructuralOps)
        {
            // Sequential mode: resolve and execute each block in order.
            // Required when operations change document structure (insert/delete paragraphs)
            // so that subsequent addresses reflect the updated document.
            foreach (var block in document.OperationBlocks)
            {
                var ranges = ResolveAddress(doc, block.Address);
                foreach (var range in ranges)
                {
                    foreach (var operation in block.Operations)
                    {
                        ExecuteOperation(doc, range, operation);
                    }
                }
            }
        }
        else
        {
            // Snapshot mode: resolve all addresses upfront, execute in reverse.
            // Safe when operations only modify content within existing ranges.
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

        // Collect all paragraphs once for repeated use
        var allParas = GetAllParagraphs(doc);

        List<dynamic> current = new();
        bool afterHeadingScope = false;

        for (int i = startIndex; i < segments.Count; i++)
        {
            var segment = segments[i];
            bool hasMore = i < segments.Count - 1;
            bool isFirst = (i == startIndex) || afterHeadingScope;

            // Heading-scoped section addressing
            if (segment.Identifier.Equals("heading", StringComparison.OrdinalIgnoreCase) && hasMore)
            {
                var headings = isFirst
                    ? ResolveHeadingsFromAll(doc, segment, allParas)
                    : ResolveHeadingsInScope(doc, segment, current, allParas);

                current = ResolveHeadingScope(doc, headings, allParas);
                afterHeadingScope = true;
                continue;
            }

            if (isFirst)
            {
                current = ResolveRootSegment(doc, segment, allParas);
            }
            else
            {
                current = ResolveChildSegment(doc, segment, current);
            }

            afterHeadingScope = false;

            if (current.Count == 0)
                return current;
        }

        return current;
    }

    /// <summary>
    /// Resolve a segment at root level (directly from the document).
    /// </summary>
    private static List<dynamic> ResolveRootSegment(
        dynamic doc, AddressSegment segment, List<(dynamic Para, int Index)> allParas)
    {
        return segment.Identifier.ToLowerInvariant() switch
        {
            "paragraph" => ResolveParagraphs(doc, segment, allParas),
            "heading" => ResolveHeadingsFromAll(doc, segment, allParas),
            "table" => ResolveTables(doc, segment),
            "run" => ResolveRuns(doc, segment, null),
            "list" or "item" => ResolveListItems(doc, segment, allParas),
            "image" => ResolveImages(doc, segment),
            "bookmark" => ResolveBookmarks(doc, segment),
            _ => throw new NotImplementedException(
                $"Address segment '{segment.Identifier}' is not yet supported for Word COM.")
        };
    }

    /// <summary>
    /// Resolve a segment within a parent context (e.g., row within table, cell within row).
    /// </summary>
    private static List<dynamic> ResolveChildSegment(
        dynamic doc, AddressSegment segment, List<dynamic> parentRanges)
    {
        return segment.Identifier.ToLowerInvariant() switch
        {
            "paragraph" => ResolveParagraphsInScope(doc, segment, parentRanges),
            "row" => ResolveRows(doc, segment, parentRanges),
            "cell" => ResolveCells(doc, segment, parentRanges),
            "run" => ResolveRuns(doc, segment, parentRanges),
            "table" => ResolveTablesInScope(doc, segment, parentRanges),
            _ => throw new NotImplementedException(
                $"Address segment '{segment.Identifier}' is not yet supported as child for Word COM.")
        };
    }

    private static List<(dynamic Para, int Index)> GetAllParagraphs(dynamic doc)
    {
        var result = new List<(dynamic Para, int Index)>();
        int i = 0;
        foreach (dynamic para in doc.Paragraphs)
        {
            result.Add((para, i++));
        }
        return result;
    }

    private static List<dynamic> ResolveParagraphs(
        dynamic doc, AddressSegment segment, List<(dynamic Para, int Index)> allParas)
    {
        var paragraphs = new List<dynamic>();
        foreach (var (para, _) in allParas)
        {
            if (!IsHeading(para))
                paragraphs.Add(para.Range);
        }
        return ApplyRangePredicates(doc, paragraphs, segment.Predicates);
    }

    private static List<dynamic> ResolveParagraphsInScope(
        dynamic doc, AddressSegment segment, List<dynamic> scopeRanges)
    {
        // Find paragraphs whose range falls within any scope range
        var paragraphs = new List<dynamic>();
        foreach (dynamic para in doc.Paragraphs)
        {
            if (IsHeading(para)) continue;
            int paraStart = (int)para.Range.Start;
            foreach (var scope in scopeRanges)
            {
                int scopeStart = (int)scope.Start;
                int scopeEnd = (int)scope.End;
                if (paraStart >= scopeStart && paraStart < scopeEnd)
                {
                    paragraphs.Add(para.Range);
                    break;
                }
            }
        }
        return ApplyRangePredicates(doc, paragraphs, segment.Predicates);
    }

    private static List<dynamic> ResolveHeadingsFromAll(
        dynamic doc, AddressSegment segment, List<(dynamic Para, int Index)> allParas)
    {
        var headings = new List<(dynamic Range, int Level)>();
        foreach (var (para, _) in allParas)
        {
            if (IsHeading(para))
                headings.Add((para.Range, GetHeadingLevel(para)));
        }
        return FilterHeadings(doc, headings, segment);
    }

    private static List<dynamic> ResolveHeadingsInScope(
        dynamic doc, AddressSegment segment, List<dynamic> scopeRanges, List<(dynamic Para, int Index)> allParas)
    {
        var headings = new List<(dynamic Range, int Level)>();
        foreach (var (para, _) in allParas)
        {
            if (!IsHeading(para)) continue;
            int paraStart = (int)para.Range.Start;
            foreach (var scope in scopeRanges)
            {
                if (paraStart >= (int)scope.Start && paraStart < (int)scope.End)
                {
                    headings.Add((para.Range, GetHeadingLevel(para)));
                    break;
                }
            }
        }
        return FilterHeadings(doc, headings, segment);
    }

    private static List<dynamic> FilterHeadings(
        dynamic doc, List<(dynamic Range, int Level)> headings, AddressSegment segment)
    {
        var levelPred = segment.Predicates
            .OfType<KeyValuePredicate>()
            .FirstOrDefault(p => p.Key.Equals("level", StringComparison.OrdinalIgnoreCase));

        if (levelPred != null && int.TryParse(levelPred.Value, out int targetLevel))
            headings = headings.Where(h => h.Level == targetLevel).ToList();

        var ranges = headings.Select(h => h.Range).ToList();
        var remaining = segment.Predicates
            .Where(p => p is not KeyValuePredicate kvp ||
                        !kvp.Key.Equals("level", StringComparison.OrdinalIgnoreCase))
            .ToList();

        return ApplyRangePredicates(doc, ranges, remaining);
    }

    /// <summary>
    /// For each matched heading, build a scope range covering all content
    /// between the heading and the next heading at same or higher level.
    /// </summary>
    private static List<dynamic> ResolveHeadingScope(
        dynamic doc, List<dynamic> headingRanges, List<(dynamic Para, int Index)> allParas)
    {
        var scopes = new List<dynamic>();

        foreach (var headingRange in headingRanges)
        {
            int headingStart = (int)headingRange.Start;

            // Find this heading in allParas to get its level and index
            int headingLevel = 0;
            int headingIdx = -1;
            for (int i = 0; i < allParas.Count; i++)
            {
                if ((int)allParas[i].Para.Range.Start == headingStart)
                {
                    headingLevel = GetHeadingLevel(allParas[i].Para);
                    headingIdx = i;
                    break;
                }
            }

            if (headingIdx < 0) continue;

            // Scope starts after the heading paragraph
            int scopeStart = (int)headingRange.End;
            int scopeEnd = (int)doc.Content.End;

            // Find the next heading at same or higher level
            for (int j = headingIdx + 1; j < allParas.Count; j++)
            {
                var (para, _) = allParas[j];
                if (IsHeading(para))
                {
                    int nextLevel = GetHeadingLevel(para);
                    if (nextLevel <= headingLevel)
                    {
                        scopeEnd = (int)para.Range.Start;
                        break;
                    }
                }
            }

            if (scopeEnd > scopeStart)
                scopes.Add(doc.Range(scopeStart, scopeEnd));
        }

        return scopes;
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

    private static List<dynamic> ResolveTablesInScope(
        dynamic doc, AddressSegment segment, List<dynamic> scopeRanges)
    {
        var tables = new List<dynamic>();
        foreach (dynamic table in doc.Tables)
        {
            int tableStart = (int)table.Range.Start;
            foreach (var scope in scopeRanges)
            {
                if (tableStart >= (int)scope.Start && tableStart < (int)scope.End)
                {
                    tables.Add(table.Range);
                    break;
                }
            }
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

    private static List<dynamic> ResolveRuns(dynamic doc, AddressSegment segment, List<dynamic>? parentRanges)
    {
        // Check if there's a text predicate we can use with Find for precise matching
        var textPred = segment.Predicates.OfType<KeyValuePredicate>()
            .FirstOrDefault(p => p.Key.Equals("text", StringComparison.OrdinalIgnoreCase)
                              && p.Operator == PredicateOperator.Equals);

        if (textPred != null)
        {
            // Use Word's Find to locate exact text — returns precise ranges
            var found = new List<dynamic>();
            var searchRanges = parentRanges ?? new List<dynamic> { doc.Content };
            foreach (var searchRange in searchRanges)
            {
                dynamic range = searchRange.Duplicate;
                dynamic find = range.Find;
                find.ClearFormatting();
                find.Text = textPred.Value;
                find.Forward = true;
                find.Wrap = 0; // wdFindStop
                find.MatchCase = true;
                find.MatchWholeWord = false;

                while ((bool)find.Execute())
                {
                    found.Add(range.Duplicate);
                    // Move past the found text to find next occurrence
                    range.Start = (int)range.End;
                    range.End = (int)searchRange.End;
                    find = range.Find;
                    find.ClearFormatting();
                    find.Text = textPred.Value;
                    find.Forward = true;
                    find.Wrap = 0;
                    find.MatchCase = true;
                    find.MatchWholeWord = false;
                }
            }
            // Apply remaining predicates (positional, etc.) excluding the text one we already used
            var remaining = segment.Predicates
                .Where(p => !ReferenceEquals(p, textPred)).ToList();
            return ApplyRangePredicates(doc, found, remaining);
        }

        // Fallback: iterate Words collection (less precise)
        var runs = new List<dynamic>();
        if (parentRanges == null)
        {
            foreach (dynamic para in doc.Paragraphs)
            {
                foreach (dynamic word in para.Range.Words)
                    runs.Add(word);
            }
        }
        else
        {
            foreach (var parentRange in parentRanges)
            {
                foreach (dynamic word in parentRange.Words)
                    runs.Add(word);
            }
        }

        return ApplyRangePredicates(doc, runs, segment.Predicates);
    }

    private static List<dynamic> ResolveListItems(
        dynamic doc, AddressSegment segment, List<(dynamic Para, int Index)> allParas)
    {
        var items = new List<dynamic>();
        foreach (var (para, _) in allParas)
        {
            try
            {
                dynamic listFormat = para.Range.ListFormat;
                if ((int)listFormat.ListType != 0) // wdListNoNumbering = 0
                    items.Add(para.Range);
            }
            catch { /* not a list item */ }
        }
        return ApplyRangePredicates(doc, items, segment.Predicates);
    }

    private static List<dynamic> ResolveImages(dynamic doc, AddressSegment segment)
    {
        var images = new List<dynamic>();
        foreach (dynamic shape in doc.InlineShapes)
        {
            images.Add(shape.Range);
        }
        return ApplyRangePredicates(doc, images, segment.Predicates);
    }

    private static List<dynamic> ResolveBookmarks(dynamic doc, AddressSegment segment)
    {
        var bookmarks = new List<dynamic>();
        foreach (dynamic bm in doc.Bookmarks)
        {
            string name = (string)bm.Name;
            if (!name.StartsWith("_"))
                bookmarks.Add(bm.Range);
        }

        // Handle bare string predicates as name match
        var barePred = segment.Predicates.OfType<BareStringPredicate>().FirstOrDefault();
        if (barePred != null)
        {
            try
            {
                dynamic bm = doc.Bookmarks[barePred.Value];
                return new List<dynamic> { bm.Range };
            }
            catch
            {
                return new List<dynamic>();
            }
        }

        var namePred = segment.Predicates.OfType<KeyValuePredicate>()
            .FirstOrDefault(p => p.Key.Equals("name", StringComparison.OrdinalIgnoreCase));
        if (namePred != null)
        {
            try
            {
                dynamic bm = doc.Bookmarks[namePred.Value];
                return new List<dynamic> { bm.Range };
            }
            catch
            {
                return new List<dynamic>();
            }
        }

        return ApplyRangePredicates(doc, bookmarks, segment.Predicates);
    }

    #endregion

    #region Predicate Application

    private static List<dynamic> ApplyRangePredicates(
        dynamic doc, List<dynamic> ranges, List<Predicate> predicates)
    {
        foreach (var predicate in predicates)
        {
            ranges = predicate switch
            {
                PositionalPredicate pos => ApplyPositional(ranges, pos),
                KeyValuePredicate kvp => ApplyKeyValue(ranges, kvp),
                BareStringPredicate bare => ApplyBareString(ranges, bare),
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
            return ranges.Where(r => MatchesText(r, predicate)).ToList();

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
                catch { return false; }
            }).ToList();
        }

        return ranges;
    }

    private static List<dynamic> ApplyBareString(List<dynamic> ranges, BareStringPredicate predicate)
    {
        return ranges.Where(r =>
        {
            string text = ((string)(r.Text ?? "")).TrimEnd('\r', '\a');
            return text.Equals(predicate.Value, StringComparison.OrdinalIgnoreCase);
        }).ToList();
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
            PredicateOperator.TildeEquals => MatchesRegex(text, predicate.Value),
            _ => false
        };
    }

    private static bool MatchesRegex(string text, string pattern)
    {
        try
        {
            return Regex.IsMatch(text, pattern, RegexOptions.None, TimeSpan.FromSeconds(1));
        }
        catch (RegexParseException) { return false; }
    }

    private static bool IsHeading(dynamic paragraph)
    {
        try
        {
            int outlineLevel = (int)paragraph.OutlineLevel;
            return outlineLevel >= WdOutlineLevel1 && outlineLevel <= WdOutlineLevel9;
        }
        catch { return false; }
    }

    private static int GetHeadingLevel(dynamic paragraph)
    {
        try
        {
            int outlineLevel = (int)paragraph.OutlineLevel;
            if (outlineLevel >= WdOutlineLevel1 && outlineLevel <= WdOutlineLevel9)
                return outlineLevel;
        }
        catch { }
        return 0;
    }

    #endregion

    #region Operation Execution

    #pragma warning disable CA1416 // Platform compatibility
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
            case DeleteOperation delete:
                ExecuteDelete(doc, range, delete);
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
            case FormatOperation format:
                ExecuteFormat(range, format);
                break;
            case InsertBeforeOperation insertBefore:
                ExecuteInsertBefore(range, insertBefore);
                break;
            case InsertAfterOperation insertAfter:
                ExecuteInsertAfter(range, insertAfter);
                break;
            case InsertRowOperation insertRow:
                ExecuteInsertRow(doc, range, insertRow);
                break;
            case InsertColumnOperation insertColumn:
                ExecuteInsertColumn(doc, range, insertColumn);
                break;
            case MergeCellsOperation merge:
                ExecuteMergeCells(doc, range, merge);
                break;
            case SetCellsOperation setCells:
                ExecuteSetCells(doc, range, setCells);
                break;
            default:
                throw new NotSupportedException(
                    $"Operation type '{operation.GetType().Name}' is not supported by the COM executor.");
        }
    }
    #pragma warning restore CA1416

    private static void ExecuteSet(dynamic range, SetOperation operation)
    {
        dynamic textRange = range.Duplicate;
        string currentText = (string)(textRange.Text ?? "");
        if (currentText.EndsWith("\r"))
            textRange.End = (int)textRange.End - 1;
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

    private static void ExecuteDelete(dynamic doc, dynamic range, DeleteOperation operation)
    {
        switch (operation.Target)
        {
            case DeleteTarget.Element:
                range.Delete();
                break;
            case DeleteTarget.Row:
                // Find the row containing this range and delete it
                foreach (dynamic table in doc.Tables)
                {
                    foreach (dynamic row in table.Rows)
                    {
                        if ((int)row.Range.Start <= (int)range.Start && (int)row.Range.End >= (int)range.End)
                        {
                            row.Delete();
                            return;
                        }
                    }
                }
                range.Delete();
                break;
            case DeleteTarget.Column:
                // Find the cell and delete the entire column
                foreach (dynamic table in doc.Tables)
                {
                    foreach (dynamic row in table.Rows)
                    {
                        foreach (dynamic cell in row.Cells)
                        {
                            if ((int)cell.Range.Start == (int)range.Start)
                            {
                                dynamic col = table.Columns[cell.ColumnIndex];
                                col.Delete();
                                return;
                            }
                        }
                    }
                }
                break;
        }
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
        // Try the style name as given, then with spaces inserted before capitals
        // (e.g., "Heading2" → "Heading 2"), then with spaces removed.
        // Word COM expects display names, but OTK may use OpenXML style IDs.
        var candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            operation.StyleName,
            InsertSpacesBeforeCapitalsAndDigits(operation.StyleName),
            operation.StyleName.Replace(" ", ""),
        };

        foreach (var candidate in candidates)
        {
            try
            {
                range.Style = candidate;
                return;
            }
            catch (COMException) { }
        }

        throw new NotSupportedException(
            $"Style '{operation.StyleName}' not found in document.");
    }

    private static string InsertSpacesBeforeCapitalsAndDigits(string name)
    {
        // "Heading2" → "Heading 2", "ListBullet" → "List Bullet"
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < name.Length; i++)
        {
            if (i > 0 && (char.IsUpper(name[i]) || char.IsDigit(name[i]))
                      && !char.IsUpper(name[i - 1]) && !char.IsDigit(name[i - 1]))
                sb.Append(' ');
            sb.Append(name[i]);
        }
        return sb.ToString();
    }

    private static void ExecuteFormat(dynamic range, FormatOperation operation)
    {
        foreach (var (key, value) in operation.Properties)
        {
            var strValue = value?.ToString() ?? "";
            switch (key.ToLowerInvariant())
            {
                case "bold":
                    range.Bold = strValue.Equals("true", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
                    break;
                case "italic":
                    range.Italic = strValue.Equals("true", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
                    break;
                case "underline":
                    // wdUnderlineSingle = 1, wdUnderlineNone = 0
                    range.Underline = strValue.Equals("true", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
                    break;
                case "strikethrough":
                    range.Font.StrikeThrough = strValue.Equals("true", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
                    break;
                case "font-name":
                    range.Font.Name = strValue;
                    break;
                case "font-size":
                    if (TryParsePoints(strValue, out double pts))
                        range.Font.Size = (float)pts;
                    break;
                case "color":
                    range.Font.Color = ParseColorToRgb(strValue);
                    break;
                case "alignment":
                case "align":
                    // wdAlignParagraphLeft=0, Center=1, Right=2, Justify=3
                    range.ParagraphFormat.Alignment = strValue.ToLowerInvariant() switch
                    {
                        "left" => 0,
                        "center" => 1,
                        "right" => 2,
                        "justify" => 3,
                        _ => 0
                    };
                    break;
                case "spacing-before":
                    if (TryParsePoints(strValue, out double before))
                        range.ParagraphFormat.SpaceBefore = (float)before;
                    break;
                case "spacing-after":
                    if (TryParsePoints(strValue, out double after))
                        range.ParagraphFormat.SpaceAfter = (float)after;
                    break;
                case "line-spacing":
                    if (TryParsePoints(strValue, out double line))
                        range.ParagraphFormat.LineSpacing = (float)line;
                    break;
                case "style":
                    try { range.Style = strValue; } catch { }
                    break;
            }
        }
    }

    private static void ExecuteInsertBefore(dynamic range, InsertBeforeOperation operation)
    {
        dynamic insertRange = range.Duplicate;
        insertRange.End = (int)insertRange.Start;
        insertRange.InsertBefore(operation.Content.Text + "\r");
    }

    private static void ExecuteInsertAfter(dynamic range, InsertAfterOperation operation)
    {
        int originalEnd = (int)range.End;

        range.InsertParagraphAfter();

        // Select the new paragraph's content range
        dynamic doc = range.Document;
        dynamic newRange = doc.Range(originalEnd, (int)range.End - 1);
        newRange.InsertBefore(operation.Content.Text);

        // Reset to Normal style — InsertParagraphAfter inherits the source
        // paragraph's style, which may be a heading. Subsequent STYLE operations
        // will set the correct style if needed.
        try { newRange.Style = "Normal"; } catch { }
    }

    private static void ExecuteInsertRow(dynamic doc, dynamic range, InsertRowOperation operation)
    {
        // Find the row containing this range
        foreach (dynamic table in doc.Tables)
        {
            foreach (dynamic row in table.Rows)
            {
                if ((int)row.Range.Start <= (int)range.Start && (int)row.Range.End >= (int)range.End)
                {
                    if (operation.Position == InsertPosition.Before)
                        table.Rows.Add(row);
                    else
                    {
                        // Add after: add before next row, or add at end
                        try
                        {
                            int rowIdx = row.Index;
                            if (rowIdx < table.Rows.Count)
                                table.Rows.Add(table.Rows[rowIdx + 1]);
                            else
                                table.Rows.Add();
                        }
                        catch
                        {
                            table.Rows.Add();
                        }
                    }
                    return;
                }
            }
        }
    }

    private static void ExecuteInsertColumn(dynamic doc, dynamic range, InsertColumnOperation operation)
    {
        foreach (dynamic table in doc.Tables)
        {
            foreach (dynamic row in table.Rows)
            {
                foreach (dynamic cell in row.Cells)
                {
                    if ((int)cell.Range.Start == (int)range.Start)
                    {
                        dynamic col = table.Columns[cell.ColumnIndex];
                        if (operation.Position == InsertPosition.Before)
                            table.Columns.Add(col);
                        else
                        {
                            try
                            {
                                int colIdx = cell.ColumnIndex;
                                if (colIdx < table.Columns.Count)
                                    table.Columns.Add(table.Columns[colIdx + 1]);
                                else
                                    table.Columns.Add();
                            }
                            catch
                            {
                                table.Columns.Add();
                            }
                        }
                        return;
                    }
                }
            }
        }
    }

    private static void ExecuteMergeCells(dynamic doc, dynamic range, MergeCellsOperation operation)
    {
        // Find start cell
        foreach (dynamic table in doc.Tables)
        {
            foreach (dynamic row in table.Rows)
            {
                foreach (dynamic cell in row.Cells)
                {
                    if ((int)cell.Range.Start == (int)range.Start)
                    {
                        // Parse target to find end cell position
                        var cellSeg = operation.TargetAddress.Segments
                            .FirstOrDefault(s => s.Identifier.Equals("cell", StringComparison.OrdinalIgnoreCase));
                        if (cellSeg == null) return;

                        var posPred = cellSeg.Predicates.OfType<PositionalPredicate>().FirstOrDefault();
                        if (posPred == null) return;

                        dynamic endCell = row.Cells[posPred.Position];
                        cell.Merge(endCell);
                        return;
                    }
                }
            }
        }
    }

    private static void ExecuteSetCells(dynamic doc, dynamic range, SetCellsOperation operation)
    {
        // Find the row containing this range
        foreach (dynamic table in doc.Tables)
        {
            foreach (dynamic row in table.Rows)
            {
                if ((int)row.Range.Start <= (int)range.Start && (int)row.Range.End >= (int)range.End)
                {
                    int cellIdx = 0;
                    foreach (dynamic cell in row.Cells)
                    {
                        if (cellIdx < operation.Values.Count)
                        {
                            dynamic cellRange = cell.Range;
                            // Trim trailing cell marker
                            cellRange.End = (int)cellRange.End - 1;
                            cellRange.Text = operation.Values[cellIdx];
                        }
                        cellIdx++;
                    }
                    return;
                }
            }
        }
    }

    #endregion

    #region Document Properties

    private static void ApplyProperty(dynamic doc, string name, string value)
    {
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

    #region Helpers

    private static bool TryParsePoints(string value, out double points)
    {
        var trimmed = value.TrimEnd();
        if (trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed[..^2];
        else if (trimmed.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            if (double.TryParse(trimmed[..^2], out double inches))
            {
                points = inches * 72;
                return true;
            }
        }
        else if (trimmed.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            if (double.TryParse(trimmed[..^2], out double cm))
            {
                points = cm * 28.3465;
                return true;
            }
        }
        return double.TryParse(trimmed, out points);
    }

    private static int ParseColorToRgb(string color)
    {
        var hex = color.TrimStart('#');
        if (hex.Length >= 6 && int.TryParse(hex[..6], System.Globalization.NumberStyles.HexNumber, null, out int rgb))
        {
            // Word COM uses BGR format
            int r = (rgb >> 16) & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = rgb & 0xFF;
            return r | (g << 8) | (b << 16);
        }
        return 0;
    }

    #endregion
}
