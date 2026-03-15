using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeTalk.Ast;
using OfficeTalkEngine.Execution;
using Xunit;

namespace OfficeTalkEngine.Tests.Execution;

public class WordExecutorTests
{
    #region Helpers

    private static WordprocessingDocument CreateInMemoryDocument(Action<Body> configure)
    {
        var stream = new MemoryStream();
        var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();
        mainPart.Document.Body = body;

        configure(body);

        mainPart.Document.Save();
        return doc;
    }

    private static Paragraph MakeParagraph(string text)
    {
        return new Paragraph(
            new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Paragraph MakeHeading(string text, int level)
    {
        return new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId { Val = $"Heading{level}" }),
            new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static OfficeTalkDocument MakeDocument(params OperationBlock[] blocks)
    {
        return new OfficeTalkDocument
        {
            Version = "1.0",
            DocType = DocType.Word,
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

    private static PositionalPredicate Pos(int n) => new() { Position = n };

    private static Table MakeTable(int rows, int cols)
    {
        var table = new Table();
        for (int r = 0; r < rows; r++)
        {
            var row = new TableRow();
            for (int c = 0; c < cols; c++)
            {
                row.AppendChild(new TableCell(
                    new Paragraph(new Run(
                        new Text($"R{r + 1}C{c + 1}") { Space = SpaceProcessingModeValues.Preserve }))));
            }
            table.AppendChild(row);
        }
        return table;
    }

    #endregion

    #region Existing Tests — SET, REPLACE, DELETE, APPEND, PREPEND, STYLE, PROPERTY

    [Fact]
    public void Set_replaces_paragraph_text()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Original"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new SetOperation { Content = new ContentValue { Text = "Replaced" } }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var paragraph = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        paragraph.InnerText.Should().Be("Replaced");
    }

    [Fact]
    public void Replace_finds_and_replaces_text()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Hello World"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new ReplaceOperation { Search = "World", Replacement = "Universe" }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var paragraph = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        paragraph.InnerText.Should().Be("Hello Universe");
    }

    [Fact]
    public void Delete_removes_element()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Keep"));
            body.AppendChild(MakeParagraph("Remove"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(2))),
                new DeleteOperation()));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var paragraphs = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToList();
        paragraphs.Should().HaveCount(1);
        paragraphs[0].InnerText.Should().Be("Keep");
    }

    [Fact]
    public void Append_adds_text_after_existing()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Hello"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new AppendOperation { Content = new ContentValue { Text = " World" } }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var paragraph = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        paragraph.InnerText.Should().Be("Hello World");
    }

    [Fact]
    public void Prepend_adds_text_before_existing()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("World"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new PrependOperation { Content = new ContentValue { Text = "Hello " } }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var paragraph = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        paragraph.InnerText.Should().Be("Hello World");
    }

    [Fact]
    public void Style_sets_paragraph_style()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Text"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new StyleOperation { StyleName = "Heading1" }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var paragraph = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value.Should().Be("Heading1");
    }

    [Fact]
    public void Property_sets_document_title()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Content"));
        });

        var otDoc = new OfficeTalkDocument
        {
            Version = "1.0",
            DocType = DocType.Word,
            PropertySettings = new List<PropertySetting>
            {
                new() { Name = "title", Value = "My Document" }
            }
        };

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        doc.PackageProperties.Title.Should().Be("My Document");
    }

    [Fact]
    public void Replace_all_replaces_every_occurrence()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("foo bar foo baz foo"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new ReplaceOperation { Search = "foo", Replacement = "qux", IsAll = true }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var paragraph = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        paragraph.InnerText.Should().Be("qux bar qux baz qux");
    }

    #endregion

    #region INSERT BEFORE / INSERT AFTER

    [Fact]
    public void InsertBefore_inserts_paragraph_before_element()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Original"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new InsertBeforeOperation { Content = new ContentValue("Before text") }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var paras = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToList();
        paras.Should().HaveCount(2);
        paras[0].InnerText.Should().Be("Before text");
        paras[1].InnerText.Should().Be("Original");
    }

    [Fact]
    public void InsertAfter_inserts_paragraph_after_element()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Original"));
            body.AppendChild(MakeParagraph("Last"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new InsertAfterOperation { Content = new ContentValue("After text") }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var paras = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToList();
        paras.Should().HaveCount(3);
        paras[0].InnerText.Should().Be("Original");
        paras[1].InnerText.Should().Be("After text");
        paras[2].InnerText.Should().Be("Last");
    }

    [Fact]
    public void InsertBefore_with_heading_scope()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Conclusion", 1));
            body.AppendChild(MakeParagraph("Conclusion text"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(
                    Seg("body"),
                    Seg("heading", new KeyValuePredicate("text", PredicateOperator.Equals, "Conclusion"))),
                new InsertBeforeOperation
                {
                    Content = new ContentValue("New section before conclusion")
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var allElements = doc.MainDocumentPart!.Document.Body!.ChildElements.OfType<Paragraph>().ToList();
        allElements.Should().HaveCount(3);
        allElements[0].InnerText.Should().Be("New section before conclusion");
    }

    #endregion

    #region FORMAT

    [Fact]
    public void Format_sets_bold_on_paragraph_runs()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Bold text"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new FormatOperation
                {
                    Properties = new Dictionary<string, object> { ["bold"] = "true" }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var run = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First()
            .Elements<Run>().First();
        run.RunProperties.Should().NotBeNull();
        run.RunProperties!.Bold.Should().NotBeNull();
    }

    [Fact]
    public void Format_sets_font_size()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Big text"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new FormatOperation
                {
                    Properties = new Dictionary<string, object> { ["font-size"] = "14pt" }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var run = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First()
            .Elements<Run>().First();
        run.RunProperties.Should().NotBeNull();
        // 14pt = 28 half-points
        run.RunProperties!.FontSize!.Val!.Value.Should().Be("28");
    }

    [Fact]
    public void Format_sets_color()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Colored text"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new FormatOperation
                {
                    Properties = new Dictionary<string, object> { ["color"] = "#2B579A" }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var run = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First()
            .Elements<Run>().First();
        run.RunProperties!.Color!.Val!.Value.Should().Be("2B579A");
    }

    [Fact]
    public void Format_sets_alignment()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Centered"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new FormatOperation
                {
                    Properties = new Dictionary<string, object> { ["alignment"] = "center" }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var para = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        para.ParagraphProperties!.Justification!.Val!.Value.Should().Be(JustificationValues.Center);
    }

    #endregion

    #region Table Operations — INSERT ROW, INSERT COLUMN, DELETE ROW/COLUMN, SET CELLS, MERGE CELLS

    [Fact]
    public void InsertRow_after_inserts_empty_row()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(2, 3));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("table", Pos(1)), Seg("row", Pos(1))),
                new InsertRowOperation { Position = InsertPosition.After }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var table = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var rows = table.Elements<TableRow>().ToList();
        rows.Should().HaveCount(3);
        // New row should be at index 1 (after first row)
        var newRowCells = rows[1].Elements<TableCell>().ToList();
        newRowCells.Should().HaveCount(3);
        // Cells should be empty
        newRowCells[0].InnerText.Should().BeEmpty();
    }

    [Fact]
    public void InsertRow_before_inserts_empty_row()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(2, 2));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("table", Pos(1)), Seg("row", Pos(1))),
                new InsertRowOperation { Position = InsertPosition.Before }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var table = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var rows = table.Elements<TableRow>().ToList();
        rows.Should().HaveCount(3);
        // New row at index 0, original first row at index 1
        rows[1].InnerText.Should().Contain("R1C1");
    }

    [Fact]
    public void InsertColumn_after_adds_cell_to_every_row()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(2, 2));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("table", Pos(1)), Seg("row", Pos(1)), Seg("cell", Pos(1))),
                new InsertColumnOperation { Position = InsertPosition.After }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var table = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        foreach (var row in table.Elements<TableRow>())
        {
            row.Elements<TableCell>().Should().HaveCount(3);
        }
    }

    [Fact]
    public void Delete_row_removes_table_row()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(3, 2));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("table", Pos(1)), Seg("row", Pos(2))),
                new DeleteOperation { Target = DeleteTarget.Row }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var table = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var rows = table.Elements<TableRow>().ToList();
        rows.Should().HaveCount(2);
        rows[0].InnerText.Should().Contain("R1C1");
        rows[1].InnerText.Should().Contain("R3C1");
    }

    [Fact]
    public void Delete_column_removes_cell_from_each_row()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(2, 3));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("table", Pos(1)), Seg("row", Pos(1)), Seg("cell", Pos(2))),
                new DeleteOperation { Target = DeleteTarget.Column }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var table = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        foreach (var row in table.Elements<TableRow>())
        {
            row.Elements<TableCell>().Should().HaveCount(2);
        }
        // Verify column 2 was removed — remaining should be C1 and C3
        var firstRow = table.Elements<TableRow>().First();
        firstRow.Elements<TableCell>().First().InnerText.Should().Contain("R1C1");
        firstRow.Elements<TableCell>().Last().InnerText.Should().Contain("R1C3");
    }

    [Fact]
    public void SetCells_populates_row_cells()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(2, 3));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("table", Pos(1)), Seg("row", Pos(1))),
                new SetCellsOperation { Values = new List<string> { "A", "B", "C" } }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var row = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First()
            .Elements<TableRow>().First();
        var cells = row.Elements<TableCell>().ToList();
        cells[0].InnerText.Should().Be("A");
        cells[1].InnerText.Should().Be("B");
        cells[2].InnerText.Should().Be("C");
    }

    [Fact]
    public void MergeCells_applies_horizontal_merge()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(2, 4));
        });

        var targetAddress = new Address
        {
            Segments = new List<AddressSegment>
            {
                new() { Identifier = "row", Predicates = new List<Predicate> { Pos(1) } },
                new() { Identifier = "cell", Predicates = new List<Predicate> { Pos(3) } }
            }
        };

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("table", Pos(1)), Seg("row", Pos(1)), Seg("cell", Pos(1))),
                new MergeCellsOperation { TargetAddress = targetAddress }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var row = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First()
            .Elements<TableRow>().First();
        var cells = row.Elements<TableCell>().ToList();

        cells[0].TableCellProperties!.HorizontalMerge!.Val!.Value.Should().Be(MergedCellValues.Restart);
        cells[1].TableCellProperties!.HorizontalMerge!.Val!.Value.Should().Be(MergedCellValues.Continue);
        cells[2].TableCellProperties!.HorizontalMerge!.Val!.Value.Should().Be(MergedCellValues.Continue);
    }

    #endregion

    #region COMMENT Tests

    [Fact]
    public void Comment_adds_comment_to_paragraph()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("This needs review."));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("body"), Seg("paragraph", Pos(1))),
                new CommentOperation { Content = new ContentValue("Please verify this claim.") }));

        new WordExecutor().Execute(otDoc, doc);

        // Verify comment was created
        var commentsPart = doc.MainDocumentPart!.WordprocessingCommentsPart;
        commentsPart.Should().NotBeNull("a comments part should exist");

        var comments = commentsPart!.Comments.Elements<Comment>().ToList();
        comments.Should().HaveCount(1);
        comments[0].InnerText.Should().Be("Please verify this claim.");
        comments[0].Author!.Value.Should().Be("OfficeTalk");

        // Verify comment range markers in paragraph
        var para = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        para.Descendants<CommentRangeStart>().Should().HaveCount(1);
        para.Descendants<CommentRangeEnd>().Should().HaveCount(1);
        para.Descendants<CommentReference>().Should().HaveCount(1);
    }

    [Fact]
    public void Comment_multiple_on_same_paragraph_creates_distinct_comments()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Revenue grew 29% year over year."));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("body"), Seg("paragraph", Pos(1))),
                new CommentOperation { Content = new ContentValue("Verify this number.") },
                new CommentOperation { Content = new ContentValue("Finance team to confirm.") }));

        new WordExecutor().Execute(otDoc, doc);

        var comments = doc.MainDocumentPart!.WordprocessingCommentsPart!
            .Comments.Elements<Comment>().ToList();
        comments.Should().HaveCount(2);
        comments[0].InnerText.Should().Be("Verify this number.");
        comments[1].InnerText.Should().Be("Finance team to confirm.");

        // IDs should be unique
        comments[0].Id!.Value.Should().NotBe(comments[1].Id!.Value);
    }

    [Fact]
    public void Comment_with_content_block_preserves_multiline()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("The plan calls for 20% expansion."));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("body"), Seg("paragraph", Pos(1))),
                new CommentOperation
                {
                    Content = new ContentValue("Line one\nLine two\nLine three", true)
                }));

        new WordExecutor().Execute(otDoc, doc);

        var comment = doc.MainDocumentPart!.WordprocessingCommentsPart!
            .Comments.Elements<Comment>().Single();

        // Each line becomes a paragraph in the comment
        var paras = comment.Elements<Paragraph>().ToList();
        paras.Should().HaveCount(3);
        paras[0].InnerText.Should().Be("Line one");
        paras[1].InnerText.Should().Be("Line two");
        paras[2].InnerText.Should().Be("Line three");
    }

    #endregion
}
