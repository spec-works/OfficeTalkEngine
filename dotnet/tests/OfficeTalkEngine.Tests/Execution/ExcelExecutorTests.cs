using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using OfficeTalk.Ast;
using OfficeTalkEngine.Execution;
using Xunit;

namespace OfficeTalkEngine.Tests.Execution;

public class ExcelExecutorTests
{
    #region Helpers

    private static SpreadsheetDocument CreateInMemorySpreadsheet(string sheetName = "Sheet1", Action<SheetData>? configure = null)
    {
        var stream = new MemoryStream();
        var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = sheetName
        });

        configure?.Invoke(sheetData);

        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return doc;
    }

    private static void AddRow(SheetData sheetData, uint rowIndex, params (string CellRef, string Value)[] cells)
    {
        var row = new Row { RowIndex = rowIndex };
        foreach (var (cellRef, value) in cells)
        {
            var cell = new Cell
            {
                CellReference = cellRef,
                CellValue = new CellValue(value),
                DataType = CellValues.String
            };
            row.AppendChild(cell);
        }
        sheetData.AppendChild(row);
    }

    private static OfficeTalkDocument MakeDocument(params OperationBlock[] blocks)
    {
        return new OfficeTalkDocument
        {
            Version = "1.0",
            DocType = DocType.Excel,
            OperationBlocks = blocks.ToList()
        };
    }

    private static OperationBlock MakeBlock(Address address, params Operation[] operations)
    {
        return new OperationBlock
        {
            Address = address,
            Operations = operations.ToList()
        };
    }

    private static Address MakeAddress(params AddressSegment[] segments)
    {
        return new Address { Segments = segments.ToList() };
    }

    private static AddressSegment Seg(string id, params Predicate[] predicates)
    {
        return new AddressSegment { Identifier = id, Predicates = predicates.ToList() };
    }

    #endregion

    [Fact]
    public void Comment_adds_comment_to_cell()
    {
        using var doc = CreateInMemorySpreadsheet("Revenue", sheetData =>
        {
            AddRow(sheetData, 1, ("A1", "Revenue"), ("B1", "Amount"));
            AddRow(sheetData, 2, ("A2", "Q1"), ("B2", "1250.00"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(
                    Seg("sheet", new BareStringPredicate("Revenue")),
                    Seg("B2")),
                new CommentOperation { Content = new ContentValue("Verify this figure.") }));

        new ExcelExecutor().Execute(otDoc, doc);

        // Find the worksheet comments part
        var worksheetPart = doc.WorkbookPart!.WorksheetParts.First();
        var commentsPart = worksheetPart.WorksheetCommentsPart;
        commentsPart.Should().NotBeNull("a comments part should exist");

        var comments = commentsPart!.Comments
            .GetFirstChild<CommentList>()!
            .Elements<Comment>().ToList();
        comments.Should().HaveCount(1);
        comments[0].Reference!.Value.Should().Be("B2");
        comments[0].InnerText.Should().Be("Verify this figure.");
    }

    [Fact]
    public void Set_updates_cell_value()
    {
        using var doc = CreateInMemorySpreadsheet("Data", sheetData =>
        {
            AddRow(sheetData, 1, ("A1", "Old Value"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(
                    Seg("sheet", new BareStringPredicate("Data")),
                    Seg("A1")),
                new SetOperation { Content = new ContentValue("New Value") }));

        new ExcelExecutor().Execute(otDoc, doc);

        var worksheetPart = doc.WorkbookPart!.WorksheetParts.First();
        var cell = worksheetPart.Worksheet
            .GetFirstChild<SheetData>()!
            .Elements<Row>().First()
            .Elements<Cell>().First();

        cell.InlineString!.InnerText.Should().Be("New Value");
    }
}
