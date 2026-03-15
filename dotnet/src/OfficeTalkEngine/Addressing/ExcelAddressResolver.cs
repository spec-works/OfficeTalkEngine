using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeTalk.Ast;

namespace OfficeTalkEngine.Addressing;

/// <summary>
/// Resolves OfficeTalk addresses against Excel (SpreadsheetDocument).
/// Supports sheet[name]/cellRef, sheet[n]/row[n], and range references.
/// </summary>
public class ExcelAddressResolver : IAddressResolver
{
    private readonly SpreadsheetDocument _doc;

    public ExcelAddressResolver(SpreadsheetDocument doc)
    {
        _doc = doc;
    }

    public IReadOnlyList<OpenXmlElement> Resolve(Address address)
    {
        if (address.Segments.Count == 0)
            throw new InvalidOperationException("Address has no segments.");

        var firstSeg = address.Segments[0];

        if (firstSeg.Identifier.Equals("sheet", StringComparison.OrdinalIgnoreCase))
        {
            var sheets = ResolveSheets(firstSeg);
            if (address.Segments.Count == 1)
                return sheets;

            // Resolve further segments within the sheet
            var results = new List<OpenXmlElement>();
            foreach (var sheetElement in sheets)
            {
                if (sheetElement is Sheet sheet)
                {
                    var worksheetPart = GetWorksheetPart(sheet);
                    if (worksheetPart == null) continue;

                    var innerResults = ResolveWithinSheet(
                        worksheetPart, address.Segments.Skip(1).ToList());
                    results.AddRange(innerResults);
                }
            }
            return results;
        }

        throw new InvalidOperationException(
            $"Excel addresses must start with 'sheet', got '{firstSeg.Identifier}'.");
    }

    private IReadOnlyList<OpenXmlElement> ResolveSheets(AddressSegment segment)
    {
        var workbookPart = _doc.WorkbookPart
            ?? throw new InvalidOperationException("Spreadsheet has no workbook part.");
        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList()
            ?? new List<Sheet>();

        return ApplyPredicates(sheets, segment.Predicates);
    }

    private WorksheetPart? GetWorksheetPart(Sheet sheet)
    {
        if (sheet.Id?.Value == null) return null;
        var workbookPart = _doc.WorkbookPart!;
        return workbookPart.GetPartById(sheet.Id!.Value!) as WorksheetPart;
    }

    private IReadOnlyList<OpenXmlElement> ResolveWithinSheet(
        WorksheetPart worksheetPart, List<AddressSegment> segments)
    {
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return Array.Empty<OpenXmlElement>();

        var seg = segments[0];
        var id = seg.Identifier.ToLowerInvariant();

        IReadOnlyList<OpenXmlElement> results;

        switch (id)
        {
            case "row":
                var rows = sheetData.Elements<Row>().ToList();
                results = ApplyPredicates(rows, seg.Predicates, isRow: true);
                break;

            default:
                // Try as cell reference (e.g., A1, B7)
                if (IsCellReference(seg.Identifier))
                {
                    results = ResolveCellReference(sheetData, seg.Identifier);
                }
                else
                {
                    throw new InvalidOperationException(
                        $"Unsupported Excel segment '{seg.Identifier}'.");
                }
                break;
        }

        // If more segments remain, resolve deeper
        if (segments.Count > 1)
        {
            var deeper = new List<OpenXmlElement>();
            foreach (var el in results)
            {
                if (el is Row row && segments[1].Identifier.Equals("cell", StringComparison.OrdinalIgnoreCase))
                {
                    var cells = row.Elements<Cell>().ToList();
                    deeper.AddRange(ApplyPredicates(cells, segments[1].Predicates));
                }
            }
            return deeper;
        }

        return results;
    }

    private static IReadOnlyList<OpenXmlElement> ResolveCellReference(SheetData sheetData, string cellRef)
    {
        // Parse cell reference like "A1", "B7"
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value != null &&
                    cell.CellReference.Value.Equals(cellRef, StringComparison.OrdinalIgnoreCase))
                {
                    return new[] { cell };
                }
            }
        }

        // Cell doesn't exist yet — create it in the appropriate row
        var (_, rowNum) = ParseCellReference(cellRef);
        var targetRow = sheetData.Elements<Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == (uint)rowNum);

        if (targetRow == null)
        {
            targetRow = new Row { RowIndex = (uint)rowNum };
            sheetData.AppendChild(targetRow);
        }

        var newCell = new Cell { CellReference = cellRef.ToUpperInvariant() };
        targetRow.AppendChild(newCell);
        return new[] { newCell };
    }

    private static bool IsCellReference(string s)
    {
        if (string.IsNullOrEmpty(s)) return false;
        int i = 0;
        // Must start with one or more letters
        while (i < s.Length && char.IsLetter(s[i])) i++;
        if (i == 0 || i == s.Length) return false;
        // Must end with one or more digits
        while (i < s.Length && char.IsDigit(s[i])) i++;
        return i == s.Length;
    }

    private static (string Column, int Row) ParseCellReference(string cellRef)
    {
        int i = 0;
        while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
        var col = cellRef[..i];
        var row = int.Parse(cellRef[i..]);
        return (col, row);
    }

    private static IReadOnlyList<OpenXmlElement> ApplyPredicates<T>(
        List<T> elements, List<Predicate> predicates, bool isRow = false)
        where T : OpenXmlElement
    {
        IEnumerable<T> filtered = elements;

        foreach (var pred in predicates)
        {
            if (pred is PositionalPredicate pos)
            {
                if (pos.Position >= 1 && pos.Position <= elements.Count)
                    filtered = new[] { elements[pos.Position - 1] };
                else
                    filtered = Enumerable.Empty<T>();
            }
            else if (pred is KeyValuePredicate kv)
            {
                if (kv.Key.Equals("name", StringComparison.OrdinalIgnoreCase) ||
                    kv.Key.Equals("text", StringComparison.OrdinalIgnoreCase))
                {
                    var val = kv.Value?.ToString() ?? "";
                    filtered = filtered.Where(e =>
                    {
                        if (e is Sheet sheet)
                            return MatchText(sheet.Name?.Value ?? "", val, kv.Operator);
                        return MatchText(e.InnerText, val, kv.Operator);
                    });
                }
            }
            else if (pred is BareStringPredicate bare)
            {
                filtered = filtered.Where(e =>
                {
                    if (e is Sheet sheet)
                        return sheet.Name?.Value == bare.Value;
                    return e.InnerText == bare.Value;
                });
            }
        }

        return filtered.Cast<OpenXmlElement>().ToList();
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