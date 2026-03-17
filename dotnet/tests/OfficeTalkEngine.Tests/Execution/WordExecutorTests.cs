using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeTalk.Ast;
using OfficeTalkEngine.Execution;
using Xunit;

using AstListItem = OfficeTalk.Ast.ListItem;

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

    #region INSERT TABLE Tests

    [Fact]
    public void InsertTable_creates_table_with_dimensions()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Before table"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new InsertTableOperation
                {
                    Position = InsertPosition.After,
                    Rows = 3,
                    Columns = 4
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var body = doc.MainDocumentPart!.Document.Body!;
        var tables = body.Elements<Table>().ToList();
        tables.Should().HaveCount(1);

        var rows = tables[0].Elements<TableRow>().ToList();
        rows.Should().HaveCount(3);

        foreach (var row in rows)
        {
            row.Elements<TableCell>().Should().HaveCount(4);
        }
    }

    [Fact]
    public void InsertTable_before_inserts_before_element()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("After table"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new InsertTableOperation
                {
                    Position = InsertPosition.Before,
                    Rows = 2,
                    Columns = 2
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var body = doc.MainDocumentPart!.Document.Body!;
        var children = body.ChildElements.ToList();
        children[0].Should().BeOfType<Table>();
    }

    [Fact]
    public void InsertTable_followed_by_SetCells_populates_first_row()
    {
        // Snapshot semantics: the table must already exist for SET CELLS to find it.
        // This test verifies that INSERT TABLE + SET CELLS works when the table
        // already exists (SET CELLS is in the same block as the table's row).
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Anchor"));
            body.AppendChild(MakeTable(2, 3));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("table", Pos(1)), Seg("row", Pos(1))),
                new SetCellsOperation { Values = new List<string> { "A", "B", "C" } }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var table = doc.MainDocumentPart!.Document.Body!.Elements<Table>().First();
        var firstRow = table.Elements<TableRow>().First();
        var cells = firstRow.Elements<TableCell>().ToList();
        cells[0].InnerText.Should().Be("A");
        cells[1].InnerText.Should().Be("B");
        cells[2].InnerText.Should().Be("C");
    }

    #endregion

    #region LINK Tests

    [Fact]
    public void Link_creates_hyperlink_on_paragraph()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Click me"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new LinkOperation { Url = "https://example.com" }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var para = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        var hyperlinks = para.Elements<Hyperlink>().ToList();
        hyperlinks.Should().HaveCount(1);
        hyperlinks[0].InnerText.Should().Be("Click me");
    }

    [Fact]
    public void Link_creates_hyperlink_relationship()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Link text"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new LinkOperation { Url = "https://example.com/page" }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var hyperlink = doc.MainDocumentPart!.Document.Body!
            .Descendants<Hyperlink>().First();
        var relId = hyperlink.Id!.Value;
        var rel = doc.MainDocumentPart!.HyperlinkRelationships
            .First(r => r.Id == relId);
        rel.Uri.ToString().Should().Be("https://example.com/page");
    }

    #endregion

    #region INSERT LIST Tests

    [Fact]
    public void InsertList_creates_numbered_paragraphs()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Before list"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new InsertListOperation
                {
                    Position = InsertPosition.After,
                    ListType = ListType.Unordered,
                    Items = new List<AstListItem>
                    {
                        new() { Content = new ContentValue("Item 1") },
                        new() { Content = new ContentValue("Item 2") },
                        new() { Content = new ContentValue("Item 3") }
                    }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var body = doc.MainDocumentPart!.Document.Body!;
        var paras = body.Elements<Paragraph>().ToList();
        paras.Should().HaveCount(4); // original + 3 items
        paras[1].InnerText.Should().Be("Item 1");
        paras[2].InnerText.Should().Be("Item 2");
        paras[3].InnerText.Should().Be("Item 3");

        // Each item paragraph should have numbering properties
        for (int i = 1; i <= 3; i++)
        {
            paras[i].ParagraphProperties.Should().NotBeNull();
            paras[i].ParagraphProperties!.NumberingProperties.Should().NotBeNull();
        }
    }

    [Fact]
    public void InsertList_nested_items_have_higher_level()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Anchor"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new InsertListOperation
                {
                    Position = InsertPosition.After,
                    ListType = ListType.Unordered,
                    Items = new List<AstListItem>
                    {
                        new() { Content = new ContentValue("Parent") },
                        new() { Content = new ContentValue("Child"), IsNested = true }
                    }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var body = doc.MainDocumentPart!.Document.Body!;
        var paras = body.Elements<Paragraph>().ToList();
        paras.Should().HaveCount(3);

        // Parent is level 0, child is level 1
        var parentLevel = paras[1].ParagraphProperties!.NumberingProperties!
            .NumberingLevelReference!.Val!.Value;
        var childLevel = paras[2].ParagraphProperties!.NumberingProperties!
            .NumberingLevelReference!.Val!.Value;

        parentLevel.Should().Be(0);
        childLevel.Should().Be(1);
    }

    [Fact]
    public void InsertList_ordered_creates_numbering_definition()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Anchor"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new InsertListOperation
                {
                    Position = InsertPosition.After,
                    ListType = ListType.Ordered,
                    Items = new List<AstListItem>
                    {
                        new() { Content = new ContentValue("Step 1") }
                    }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var numberingPart = doc.MainDocumentPart!.NumberingDefinitionsPart;
        numberingPart.Should().NotBeNull();
        numberingPart!.Numbering.Elements<AbstractNum>().Should().NotBeEmpty();
        numberingPart!.Numbering.Elements<NumberingInstance>().Should().NotBeEmpty();
    }

    #endregion

    #region SET RUNS Tests

    [Fact]
    public void SetRuns_creates_multiple_formatted_runs()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Original text"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new SetRunsOperation
                {
                    Runs = new List<RunDefinition>
                    {
                        new() { Content = new ContentValue("Normal ") },
                        new()
                        {
                            Content = new ContentValue("Bold"),
                            Properties = new Dictionary<string, object> { ["bold"] = "true" }
                        },
                        new() { Content = new ContentValue(" text") }
                    }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var para = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        var runs = para.Elements<Run>().ToList();
        runs.Should().HaveCount(3);
        runs[0].InnerText.Should().Be("Normal ");
        runs[1].InnerText.Should().Be("Bold");
        runs[2].InnerText.Should().Be(" text");

        // Second run should be bold
        runs[1].RunProperties.Should().NotBeNull();
        runs[1].RunProperties!.Bold.Should().NotBeNull();
    }

    [Fact]
    public void SetRuns_with_href_creates_hyperlink()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Original"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new SetRunsOperation
                {
                    Runs = new List<RunDefinition>
                    {
                        new() { Content = new ContentValue("Visit ") },
                        new()
                        {
                            Content = new ContentValue("our site"),
                            Properties = new Dictionary<string, object>
                            {
                                ["href"] = "https://example.com"
                            }
                        },
                        new() { Content = new ContentValue(" today") }
                    }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var para = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();

        // Should have: Run("Visit "), Hyperlink(Run("our site")), Run(" today")
        var directRuns = para.Elements<Run>().ToList();
        directRuns.Should().HaveCount(2);
        directRuns[0].InnerText.Should().Be("Visit ");
        directRuns[1].InnerText.Should().Be(" today");

        var hyperlinks = para.Elements<Hyperlink>().ToList();
        hyperlinks.Should().HaveCount(1);
        hyperlinks[0].InnerText.Should().Be("our site");
    }

    [Fact]
    public void SetRuns_with_font_properties()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Original"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new SetRunsOperation
                {
                    Runs = new List<RunDefinition>
                    {
                        new()
                        {
                            Content = new ContentValue("code"),
                            Properties = new Dictionary<string, object>
                            {
                                ["font-name"] = "Consolas",
                                ["font-size"] = "10pt",
                                ["color"] = "#FF0000"
                            }
                        }
                    }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var para = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        var run = para.Elements<Run>().First();
        run.InnerText.Should().Be("code");
        run.RunProperties.Should().NotBeNull();
        run.RunProperties!.RunFonts!.Ascii!.Value.Should().Be("Consolas");
        run.RunProperties!.Color!.Val!.Value.Should().Be("FF0000");
    }

    [Fact]
    public void SetRuns_replaces_all_existing_content()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            var para = new Paragraph(
                new Run(new Text("Run 1")),
                new Run(new Text("Run 2")),
                new Run(new Text("Run 3")));
            body.AppendChild(para);
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new SetRunsOperation
                {
                    Runs = new List<RunDefinition>
                    {
                        new() { Content = new ContentValue("Only run") }
                    }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var para = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        var runs = para.Elements<Run>().ToList();
        runs.Should().HaveCount(1);
        runs[0].InnerText.Should().Be("Only run");
    }

    #endregion

    #region INSERT IMAGE Tests

    [Fact]
    public void InsertImage_inserts_paragraph_with_drawing()
    {
        // Create a minimal test image file
        var tempDir = Path.Combine(Path.GetTempPath(), $"otk-test-{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);
        var imgPath = Path.Combine(tempDir, "test.png");

        try
        {
            // Write a minimal 1x1 PNG
            File.WriteAllBytes(imgPath, CreateMinimalPng());

            using var doc = CreateInMemoryDocument(body =>
            {
                body.AppendChild(MakeParagraph("Before image"));
            });

            var otDoc = MakeDocument(
                MakeBlock(
                    MakeAddress(Seg("paragraph", Pos(1))),
                    new InsertImageOperation
                    {
                        Position = InsertPosition.After,
                        Source = imgPath,
                        Properties = new Dictionary<string, object>
                        {
                            ["alt"] = "Test image",
                            ["width"] = "4in"
                        }
                    }));

            var executor = new WordExecutor();
            executor.Execute(otDoc, doc);

            var body = doc.MainDocumentPart!.Document.Body!;
            var paras = body.Elements<Paragraph>().ToList();
            paras.Should().HaveCount(2);

            // Second paragraph should contain a drawing
            var drawings = paras[1].Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().ToList();
            drawings.Should().HaveCount(1);

            // Image part should exist
            doc.MainDocumentPart!.ImageParts.Should().NotBeEmpty();
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Fact]
    public void InsertImage_before_inserts_before_element()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), $"otk-test-{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDir);
        var imgPath = Path.Combine(tempDir, "test.png");

        try
        {
            File.WriteAllBytes(imgPath, CreateMinimalPng());

            using var doc = CreateInMemoryDocument(body =>
            {
                body.AppendChild(MakeParagraph("After image"));
            });

            var otDoc = MakeDocument(
                MakeBlock(
                    MakeAddress(Seg("paragraph", Pos(1))),
                    new InsertImageOperation
                    {
                        Position = InsertPosition.Before,
                        Source = imgPath
                    }));

            var executor = new WordExecutor();
            executor.Execute(otDoc, doc);

            var body = doc.MainDocumentPart!.Document.Body!;
            var paras = body.Elements<Paragraph>().ToList();
            paras.Should().HaveCount(2);

            // First paragraph should contain the drawing
            paras[0].Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>()
                .Should().NotBeEmpty();
            paras[1].InnerText.Should().Be("After image");
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    /// <summary>Creates a minimal valid 1x1 pixel PNG file.</summary>
    private static byte[] CreateMinimalPng()
    {
        // Minimal 1x1 white PNG
        return new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, // IDAT chunk
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, // IEND chunk
            0x44, 0xAE, 0x42, 0x60, 0x82
        };
    }

    #endregion

    #region FORMAT — Background Color & Borders

    [Fact]
    public void Format_BackgroundColor_AppliesShadingToRuns()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Code text"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new FormatOperation
                {
                    Properties = new Dictionary<string, object> { ["background-color"] = "#F0F0F0" }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var run = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First()
            .Elements<Run>().First();
        run.RunProperties.Should().NotBeNull();
        run.RunProperties!.Shading.Should().NotBeNull();
        run.RunProperties!.Shading!.Fill!.Value.Should().Be("F0F0F0");
    }

    [Fact]
    public void Format_Highlight_AppliesHighlightToRuns()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Highlighted"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new FormatOperation
                {
                    Properties = new Dictionary<string, object> { ["highlight"] = "yellow" }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var run = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First()
            .Elements<Run>().First();
        run.RunProperties.Should().NotBeNull();
        run.RunProperties!.Highlight.Should().NotBeNull();
        run.RunProperties!.Highlight!.Val!.Value.Should().Be(HighlightColorValues.Yellow);
    }

    [Fact]
    public void Format_BorderBottom_CreatesParagraphBorder()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph(""));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new FormatOperation
                {
                    Properties = new Dictionary<string, object>
                    {
                        ["border-bottom"] = "single",
                        ["border-color"] = "#CCCCCC"
                    }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var para = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        para.ParagraphProperties.Should().NotBeNull();
        var borders = para.ParagraphProperties!.ParagraphBorders;
        borders.Should().NotBeNull();
        borders!.BottomBorder.Should().NotBeNull();
        borders.BottomBorder!.Val!.Value.Should().Be(BorderValues.Single);
        borders.BottomBorder!.Color!.Value.Should().Be("CCCCCC");
    }

    [Fact]
    public void Format_AllBorderSides_CreatesAllBorders()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Boxed"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new FormatOperation
                {
                    Properties = new Dictionary<string, object>
                    {
                        ["border-top"] = "single",
                        ["border-bottom"] = "single",
                        ["border-left"] = "single",
                        ["border-right"] = "single"
                    }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var para = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        var borders = para.ParagraphProperties!.ParagraphBorders;
        borders.Should().NotBeNull();
        borders!.TopBorder.Should().NotBeNull();
        borders!.BottomBorder.Should().NotBeNull();
        borders!.LeftBorder.Should().NotBeNull();
        borders!.RightBorder.Should().NotBeNull();
    }

    [Fact]
    public void SetRuns_BackgroundColor_AppliesShadingToRun()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("old"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new SetRunsOperation
                {
                    Runs = new List<RunDefinition>
                    {
                        new() { Content = new ContentValue("code"),
                            Properties = new Dictionary<string, object>
                            {
                                ["font-name"] = "Consolas",
                                ["background-color"] = "#F5F5F5"
                            }
                        }
                    }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var run = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First()
            .Elements<Run>().First();
        run.RunProperties.Should().NotBeNull();
        run.RunProperties!.Shading.Should().NotBeNull();
        run.RunProperties!.Shading!.Fill!.Value.Should().Be("F5F5F5");
        run.RunProperties!.RunFonts!.Ascii!.Value.Should().Be("Consolas");
    }

    [Fact]
    public void Format_BorderBottomDouble_UsesDoubleStyle()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph(""));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", Pos(1))),
                new FormatOperation
                {
                    Properties = new Dictionary<string, object>
                    {
                        ["border-bottom"] = "double"
                    }
                }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var para = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        var borders = para.ParagraphProperties!.ParagraphBorders;
        borders!.BottomBorder!.Val!.Value.Should().Be(BorderValues.Double);
    }

    #endregion
}
