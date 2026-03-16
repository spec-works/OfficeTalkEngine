using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using OfficeTalk.Ast;

namespace OfficeTalkEngine.Execution;

/// <summary>
/// Executes OfficeTalk operations against a live Excel instance via dynamic COM.
/// Uses late-bound COM to avoid dependency on Office PIAs.
/// Requires Excel to be running with the target workbook open.
/// </summary>
[SupportedOSPlatform("windows")]
public class ExcelComExecutor : IOfficeTalkExecutor
{
    public void Execute(OfficeTalkDocument document, string targetPath, string? outputPath = null)
    {
        if (outputPath != null)
            throw new NotSupportedException(
                "Output path is not supported when editing via Excel COM. The live workbook is modified in place.");

        dynamic excelApp = GetRunningExcelInstance();
        dynamic workbook = FindOpenWorkbook(excelApp, targetPath);

        foreach (var block in document.OperationBlocks)
        {
            var targets = ResolveAddress(workbook, block.Address);
            foreach (var target in targets)
            {
                foreach (var operation in block.Operations)
                {
                    ExecuteOperation(target, operation);
                }
            }
        }
    }

    /// <summary>
    /// Checks whether Excel is running and has the specified workbook open.
    /// </summary>
    public static bool IsAvailable(string targetPath)
    {
        try
        {
            dynamic excelApp = GetRunningExcelInstance();
            FindOpenWorkbook(excelApp, targetPath);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static object GetRunningExcelInstance()
    {
        try
        {
            return ComInteropHelper.GetActiveObject("Excel.Application");
        }
        catch (COMException)
        {
            throw new InvalidOperationException(
                "Excel is not running. The COM executor requires an active Excel instance.");
        }
    }

    private static object FindOpenWorkbook(dynamic excelApp, string targetPath)
    {
        string fullPath = Path.GetFullPath(targetPath);

        foreach (dynamic wb in excelApp.Workbooks)
        {
            if (string.Equals((string)wb.FullName, fullPath, StringComparison.OrdinalIgnoreCase))
                return wb;
        }

        throw new FileNotFoundException(
            $"Workbook is not open in Excel: {fullPath}");
    }

    #region Address Resolution

    private static List<dynamic> ResolveAddress(dynamic workbook, Address address)
    {
        if (address.Segments.Count == 0)
            throw new InvalidOperationException("Address has no segments.");

        var first = address.Segments[0];

        if (!first.Identifier.Equals("sheet", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                $"Excel addresses must start with 'sheet', got '{first.Identifier}'.");

        var sheets = ResolveSheets(workbook, first);
        if (address.Segments.Count == 1)
            return sheets;

        var results = new List<dynamic>();
        foreach (var sheet in sheets)
        {
            var inner = ResolveWithinSheet(sheet, address.Segments.Skip(1).ToList());
            results.AddRange(inner);
        }
        return results;
    }

    private static List<dynamic> ResolveSheets(dynamic workbook, AddressSegment segment)
    {
        var sheets = new List<dynamic>();
        int count = (int)workbook.Worksheets.Count;

        for (int i = 1; i <= count; i++)
            sheets.Add(workbook.Worksheets[i]);

        return ApplyPredicates(sheets, segment.Predicates, getName: s => (string)s.Name);
    }

    private static List<dynamic> ResolveWithinSheet(
        dynamic sheet, List<AddressSegment> segments)
    {
        var seg = segments[0];
        var id = seg.Identifier.ToLowerInvariant();

        List<dynamic> results;

        switch (id)
        {
            case "row":
                results = ResolveRows(sheet, seg);
                break;

            default:
                if (IsCellReference(seg.Identifier))
                {
                    // Direct cell reference like A1, B7
                    dynamic cell = sheet.Range[seg.Identifier.ToUpperInvariant()];
                    results = new List<dynamic> { cell };
                }
                else
                {
                    throw new InvalidOperationException(
                        $"Unsupported Excel segment '{seg.Identifier}'.");
                }
                break;
        }

        // Deeper segments (e.g., row[n]/cell)
        if (segments.Count > 1)
        {
            var deeper = new List<dynamic>();
            var nextSeg = segments[1];
            if (nextSeg.Identifier.Equals("cell", StringComparison.OrdinalIgnoreCase))
            {
                foreach (var row in results)
                {
                    // row is an Excel Range representing the row
                    // Apply positional predicate to get specific cell in the row
                    foreach (var pred in nextSeg.Predicates)
                    {
                        if (pred is PositionalPredicate pos)
                        {
                            deeper.Add(row.Cells[1, pos.Position]);
                        }
                    }
                }
            }
            return deeper;
        }

        return results;
    }

    private static List<dynamic> ResolveRows(dynamic sheet, AddressSegment segment)
    {
        var rows = new List<dynamic>();
        foreach (var pred in segment.Predicates)
        {
            if (pred is PositionalPredicate pos)
            {
                rows.Add(sheet.Rows[pos.Position]);
            }
        }
        return rows;
    }

    #endregion

    #region Operation Execution

    #pragma warning disable CA1416
    private static void ExecuteOperation(dynamic target, Operation operation)
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
                    $"Operation type '{operation.GetType().Name}' is not supported by the Excel COM executor.");
        }
    }
    #pragma warning restore CA1416

    private static void ExecuteSet(dynamic target, SetOperation operation)
    {
        // target is an Excel Range (cell or range of cells)
        target.Value = operation.Content.Text;
    }

    private static void ExecuteDelete(dynamic target)
    {
        // Works for rows, columns, cells, or sheets
        target.Delete();
    }

    private static void ExecuteComment(dynamic target, CommentOperation operation)
    {
        // Clear any existing comment first
        try { target.ClearComments(); } catch { /* no existing comment */ }
        // AddComment adds a comment to the range (cell)
        target.AddComment(operation.Content.Text);
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

    private static bool IsCellReference(string s)
    {
        if (string.IsNullOrEmpty(s)) return false;
        int i = 0;
        while (i < s.Length && char.IsLetter(s[i])) i++;
        if (i == 0 || i == s.Length) return false;
        while (i < s.Length && char.IsDigit(s[i])) i++;
        return i == s.Length;
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
