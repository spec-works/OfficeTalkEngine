using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeTalk.Ast;
using OfficeTalk.Parsing;
using OfficeTalkEngine.Inspection;
using OfficeTalkEngine.Responses;
using Xunit;

namespace OfficeTalkEngine.Tests.Inspection;

public class WordInspectorTests
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

    private static Paragraph MakeStyledParagraph(string text, string styleId)
    {
        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = styleId }),
            new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Table MakeTable(params string[][] rows)
    {
        var table = new Table();
        foreach (var row in rows)
        {
            var tr = new TableRow();
            foreach (var cellText in row)
            {
                tr.AppendChild(new TableCell(MakeParagraph(cellText)));
            }
            table.AppendChild(tr);
        }
        return table;
    }

    private static OfficeTalkDocument ParseOtk(string input)
    {
        var lexer = new OfficeTalkLexer(input);
        var tokens = lexer.Tokenize();
        var parser = new OfficeTalkParser(tokens);
        return parser.Parse();
    }

    #endregion

    // ─── Basic addressing ────────────────────────────────────────

    [Fact]
    public void Inspect_Heading_ReturnsAllHeadings()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Introduction", 1));
            body.AppendChild(MakeParagraph("Some text."));
            body.AppendChild(MakeHeading("Details", 1));
            body.AppendChild(MakeParagraph("More text."));
            body.AppendChild(MakeHeading("Conclusion", 1));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses.Should().HaveCount(1);
        var response = responses[0];
        response.Op.Should().Be("inspect");
        response.Matched.Should().Be(3);
        response.Elements.Should().HaveCount(3);
        response.Elements[0].Type.Should().Be("heading");
        response.Elements[0].Level.Should().Be(1);
        response.Error.Should().BeNull();
    }

    [Fact]
    public void Inspect_Paragraph_ReturnsAllParagraphs()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("First"));
            body.AppendChild(MakeParagraph("Second"));
            body.AppendChild(MakeParagraph("Third"));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Matched.Should().Be(3);
        responses[0].Elements.Should().AllSatisfy(e => e.Type.Should().Be("paragraph"));
    }

    [Fact]
    public void Inspect_PositionalPredicate_ReturnsSingleElement()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("First", 1));
            body.AppendChild(MakeHeading("Second", 1));
            body.AppendChild(MakeHeading("Third", 1));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading[2]\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Matched.Should().Be(1);
    }

    [Fact]
    public void Inspect_Table_ReturnsTableInfo()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Before table"));
            body.AppendChild(MakeTable(
                new[] { "A", "B" },
                new[] { "1", "2" }
            ));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/table[1]\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Matched.Should().Be(1);
        responses[0].Elements[0].Type.Should().Be("table");
        responses[0].Elements[0].Index.Should().Be(1);
    }

    // ─── INCLUDE content ─────────────────────────────────────────

    [Fact]
    public void Inspect_IncludeContent_ReturnsTextForHeadings()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Introduction", 1));
            body.AppendChild(MakeParagraph("Body text"));
            body.AppendChild(MakeHeading("Conclusion", 1));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n  INCLUDE content\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Content.Should().NotBeNull();
        responses[0].Elements[0].Content!.Text.Should().Be("Introduction");
        responses[0].Elements[1].Content!.Text.Should().Be("Conclusion");
    }

    [Fact]
    public void Inspect_NoInclude_ContentIsNull()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Test", 1));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Content.Should().BeNull();
    }

    [Fact]
    public void Inspect_IncludeContent_TableRowReturnsCells()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(
                new[] { "Header1", "Header2" },
                new[] { "Val1", "Val2" }
            ));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/table[1]/row[1]\n  INCLUDE content\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Matched.Should().Be(1);
        var row = responses[0].Elements[0];
        row.Type.Should().Be("row");
        row.Content!.Cells.Should().BeEquivalentTo(new[] { "Header1", "Header2" });
    }

    // ─── INCLUDE properties ──────────────────────────────────────

    [Fact]
    public void Inspect_IncludeProperties_ReturnsBoldForFormattedRun()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(new Paragraph(
                new Run(
                    new RunProperties(new Bold()),
                    new Text("Bold text") { Space = SpaceProcessingModeValues.Preserve })));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph[1]\n  INCLUDE properties\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Properties.Should().NotBeNull();
        responses[0].Elements[0].Properties!.Should().ContainKey("bold");
    }

    [Fact]
    public void Inspect_IncludeProperties_ReturnsStyleForStyledParagraph()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeStyledParagraph("Title text", "Title"));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph[1]\n  INCLUDE properties\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Properties!["style"].Should().Be("Title");
    }

    [Fact]
    public void Inspect_NoIncludeProperties_PropertiesIsNull()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Plain text"));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph[1]\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Properties.Should().BeNull();
    }

    // ─── DEPTH ───────────────────────────────────────────────────

    [Fact]
    public void Inspect_Depth1_ReturnsTableRows()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(
                new[] { "A1", "B1" },
                new[] { "A2", "B2" },
                new[] { "A3", "B3" }
            ));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/table[1]\n  DEPTH 1\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        var table = responses[0].Elements[0];
        table.Children.Should().NotBeNull();
        table.Children!.Should().HaveCount(3);
        table.Children![0].Type.Should().Be("row");
    }

    [Fact]
    public void Inspect_Depth2_ReturnsTableRowsAndCells()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(
                new[] { "A1", "B1" },
                new[] { "A2", "B2" }
            ));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/table[1]\n  DEPTH 2\n  INCLUDE content\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        var table = responses[0].Elements[0];
        table.Children.Should().HaveCount(2);
        table.Children![0].Children.Should().NotBeNull();
        table.Children![0].Children!.Should().HaveCount(2);
        table.Children![0].Children![0].Type.Should().Be("cell");
        table.Children![0].Children![0].Content!.Text.Should().Be("A1");
    }

    [Fact]
    public void Inspect_Depth0_NoChildren()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(new[] { "A", "B" }));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/table[1]\n  DEPTH 0\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Children.Should().BeNull();
    }

    // ─── CONTEXT ─────────────────────────────────────────────────

    [Fact]
    public void Inspect_Context_ReturnsSurroundingElements()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Before 1"));
            body.AppendChild(MakeParagraph("Before 2"));
            body.AppendChild(MakeHeading("Target", 1));
            body.AppendChild(MakeParagraph("After 1"));
            body.AppendChild(MakeParagraph("After 2"));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading[1]\n  INCLUDE content\n  CONTEXT 2\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        var element = responses[0].Elements[0];
        element.Context.Should().NotBeNull();
        element.Context!.Before.Should().HaveCount(2);
        element.Context!.After.Should().HaveCount(2);
        element.Context!.Before[0].Content!.Text.Should().Be("Before 1");
        element.Context!.After[1].Content!.Text.Should().Be("After 2");
    }

    [Fact]
    public void Inspect_Context_ClampedAtDocumentEdges()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("First", 1));
            body.AppendChild(MakeParagraph("Only after"));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading[1]\n  INCLUDE content\n  CONTEXT 5\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        var ctx = responses[0].Elements[0].Context!;
        ctx.Before.Should().BeEmpty();
        ctx.After.Should().HaveCount(1);
    }

    [Fact]
    public void Inspect_NoContext_ContextIsNull()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Test", 1));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Context.Should().BeNull();
    }

    // ─── Comments ────────────────────────────────────────────────

    [Fact]
    public void Inspect_Comments_IncludedWhenPresent()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            var para = new Paragraph(
                new CommentRangeStart { Id = "1" },
                new Run(new Text("Commented text") { Space = SpaceProcessingModeValues.Preserve }),
                new CommentRangeEnd { Id = "1" }
            );
            body.AppendChild(para);
        });

        // Add comment part
        var commentsPart = doc.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments = new Comments(
            new Comment
            {
                Id = "1",
                Author = "TestUser",
                Date = DateTime.Parse("2024-01-15T10:30:00Z"),
                InnerXml = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Fix this</w:t></w:r></w:p>"
            }
        );
        doc.MainDocumentPart!.Document.Save();

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph[1]\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        var element = responses[0].Elements[0];
        element.Comments.Should().NotBeNull();
        element.Comments!.Should().HaveCount(1);
        element.Comments![0].Author.Should().Be("TestUser");
        element.Comments![0].Text.Should().Contain("Fix this");
    }

    [Fact]
    public void Inspect_NoComments_CommentsIsNull()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("No comments"));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph[1]\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Comments.Should().BeNull();
    }

    // ─── Multiple INSPECT blocks ─────────────────────────────────

    [Fact]
    public void Inspect_MultipleBlocks_ReturnsMultipleResponses()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Title", 1));
            body.AppendChild(MakeParagraph("Body text"));
            body.AppendChild(MakeTable(new[] { "A", "B" }));
        });

        var otk = ParseOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading
  INCLUDE content

INSPECT body/paragraph
  INCLUDE content

INSPECT body/table[1]
");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses.Should().HaveCount(3);
        responses[0].Elements[0].Content!.Text.Should().Be("Title");
        responses[1].Elements[0].Content!.Text.Should().Be("Body text");
        responses[2].Elements[0].Type.Should().Be("table");
    }

    // ─── All modifiers combined ──────────────────────────────────

    [Fact]
    public void Inspect_AllModifiers_ProducesCompleteResponse()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Before"));
            body.AppendChild(MakeTable(
                new[] { "H1", "H2" },
                new[] { "V1", "V2" }
            ));
            body.AppendChild(MakeParagraph("After"));
        });

        var otk = ParseOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/table[1]
  DEPTH 2
  INCLUDE content, properties
  CONTEXT 1
");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        var table = responses[0].Elements[0];
        table.Type.Should().Be("table");
        table.Content.Should().NotBeNull();
        table.Properties.Should().NotBeNull();
        table.Children.Should().NotBeNull();
        table.Context.Should().NotBeNull();
        table.Context!.Before.Should().HaveCount(1);
        table.Context!.After.Should().HaveCount(1);
    }

    // ─── Error handling ──────────────────────────────────────────

    [Fact]
    public void Inspect_InvalidAddress_ReturnsErrorResponse()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Test"));
        });

        var block = new InspectBlock
        {
            Address = new Address
            {
                Segments = new List<AddressSegment>
                {
                    new() { Identifier = "nonexistent" }
                }
            },
            Line = 4
        };

        var inspector = new WordInspector(doc);
        var response = inspector.InspectBlock(block);

        response.Matched.Should().Be(0);
        response.Error.Should().NotBeNull();
    }

    [Fact]
    public void Inspect_NoMatches_ReturnsEmptyElements()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("No headings"));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Matched.Should().Be(0);
        responses[0].Elements.Should().BeEmpty();
    }

    // ─── Position tracking ───────────────────────────────────────

    [Fact]
    public void Inspect_HeadingPosition_ReflectsBodyPosition()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Para 1"));
            body.AppendChild(MakeHeading("Heading 1", 1));
            body.AppendChild(MakeParagraph("Para 2"));
            body.AppendChild(MakeHeading("Heading 2", 1));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Index.Should().Be(2);
        responses[0].Elements[0].Of.Should().Be(4);
        responses[0].Elements[1].Index.Should().Be(4);
    }

    // ─── JSONL serialization ─────────────────────────────────────

    [Fact]
    public void JsonlWriter_ProducesValidJsonl()
    {
        var response = new InspectResponse
        {
            Address = "body/heading",
            Matched = 2,
            Elements = new List<ElementInfo>
            {
                new() { Type = "heading", Level = 1, Index = 1, Of = 5 },
                new() { Type = "heading", Level = 1, Index = 3, Of = 5 }
            }
        };

        using var sw = new StringWriter();
        var writer = new JsonlResponseWriter(sw);
        writer.WriteInspectResponse(response);
        writer.Flush();

        var json = sw.ToString().Trim();
        var parsed = JsonDocument.Parse(json);

        parsed.RootElement.GetProperty("op").GetString().Should().Be("inspect");
        parsed.RootElement.GetProperty("matched").GetInt32().Should().Be(2);
        parsed.RootElement.GetProperty("elements").GetArrayLength().Should().Be(2);
    }

    [Fact]
    public void JsonlWriter_OmitsNullFields()
    {
        var response = new InspectResponse
        {
            Address = "body/paragraph",
            Matched = 1,
            Elements = new List<ElementInfo>
            {
                new() { Type = "paragraph", Index = 1, Of = 3 }
            }
        };

        using var sw = new StringWriter();
        var writer = new JsonlResponseWriter(sw);
        writer.WriteInspectResponse(response);
        writer.Flush();

        var json = sw.ToString().Trim();
        json.Should().NotContain("\"error\"");
        json.Should().NotContain("\"content\"");
        json.Should().NotContain("\"properties\"");
        json.Should().NotContain("\"children\"");
        json.Should().NotContain("\"context\"");
        json.Should().NotContain("\"comments\"");
        json.Should().NotContain("\"level\"");
    }

    [Fact]
    public void JsonlWriter_OperationResponse_SerializesCorrectly()
    {
        var response = new OperationResponse
        {
            Op = "set",
            Address = "body/paragraph[1]",
            Status = "ok"
        };

        using var sw = new StringWriter();
        var writer = new JsonlResponseWriter(sw);
        writer.WriteOperationResponse(response);
        writer.Flush();

        var json = sw.ToString().Trim();
        var parsed = JsonDocument.Parse(json);
        parsed.RootElement.GetProperty("op").GetString().Should().Be("set");
        parsed.RootElement.GetProperty("status").GetString().Should().Be("ok");
        json.Should().NotContain("\"message\"");
    }

    // ─── Full spec example: Word document outline ────────────────

    [Fact]
    public void Inspect_SpecExample_WordDocumentOutline()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Introduction", 1));
            body.AppendChild(MakeParagraph("Intro text"));
            body.AppendChild(MakeParagraph("More intro"));
            body.AppendChild(MakeHeading("Specifications", 1));
            body.AppendChild(MakeParagraph("Spec text"));
            body.AppendChild(MakeParagraph("More spec"));
            body.AppendChild(MakeHeading("Getting Started", 1));
            body.AppendChild(MakeParagraph("Getting started text"));
            body.AppendChild(MakeParagraph("More getting started"));
            body.AppendChild(MakeParagraph("Final paragraph"));
            body.AppendChild(MakeParagraph("Last paragraph"));
        });

        var otk = ParseOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading
  INCLUDE content
");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        var resp = responses[0];
        resp.Matched.Should().Be(3);
        resp.Elements[0].Type.Should().Be("heading");
        resp.Elements[0].Level.Should().Be(1);
        resp.Elements[0].Style.Should().Be("Heading1");
        resp.Elements[0].Content!.Text.Should().Be("Introduction");
        resp.Elements[1].Content!.Text.Should().Be("Specifications");
        resp.Elements[2].Content!.Text.Should().Be("Getting Started");

        resp.Elements[0].Index.Should().Be(1);
        resp.Elements[0].Of.Should().Be(11);
    }

    // ─── Heading with context spec example ───────────────────────

    [Fact]
    public void Inspect_SpecExample_HeadingWithContext()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Summary of findings"));
            body.AppendChild(MakeParagraph("Key point A"));
            body.AppendChild(MakeHeading("Conclusion", 1));
            body.AppendChild(MakeParagraph("In conclusion..."));
            body.AppendChild(MakeParagraph("Future work"));
            body.AppendChild(MakeParagraph("References"));
        });

        var otk = ParseOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading[text=""Conclusion""]
  INCLUDE content
  CONTEXT 3
");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        var resp = responses[0];
        resp.Matched.Should().Be(1);
        resp.Elements[0].Content!.Text.Should().Be("Conclusion");
        resp.Elements[0].Context.Should().NotBeNull();
        resp.Elements[0].Context!.Before.Should().HaveCount(2);
        resp.Elements[0].Context!.After.Should().HaveCount(3);
    }

    // ─── Table with depth + content + properties ─────────────────

    [Fact]
    public void Inspect_SpecExample_TableWithDepthAndProperties()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(
                new[] { "Department", "Budget", "Actual" },
                new[] { "Engineering", "150000", "162000" },
                new[] { "Marketing", "80000", "75000" }
            ));
        });

        var otk = ParseOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/table[1]
  DEPTH 2
  INCLUDE content, properties
");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        var table = responses[0].Elements[0];
        table.Type.Should().Be("table");
        table.Children.Should().HaveCount(3);
        table.Children![0].Type.Should().Be("row");
        table.Children![0].Children.Should().HaveCount(3);
        table.Children![0].Children![0].Type.Should().Be("cell");
        table.Children![0].Children![0].Content!.Text.Should().Be("Department");
    }

    // ─── Content control ─────────────────────────────────────────

    [Fact]
    public void Inspect_ContentControl_ReturnsTag()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            var sdt = new SdtBlock(
                new SdtProperties(new Tag { Val = "my-control" }),
                new SdtContentBlock(MakeParagraph("Control content"))
            );
            body.AppendChild(sdt);
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/content-control[1]\n  INCLUDE content\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Type.Should().Be("content-control");
        responses[0].Elements[0].Tag.Should().Be("my-control");
        responses[0].Elements[0].Content!.Text.Should().Be("Control content");
    }

    // ─── Mixed heading levels ────────────────────────────────────

    [Fact]
    public void Inspect_MixedHeadingLevels_ReturnsCorrectLevels()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Chapter 1", 1));
            body.AppendChild(MakeHeading("Section 1.1", 2));
            body.AppendChild(MakeHeading("Section 1.2", 2));
            body.AppendChild(MakeHeading("Chapter 2", 1));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n  INCLUDE content\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Matched.Should().Be(4);
        responses[0].Elements[0].Level.Should().Be(1);
        responses[0].Elements[0].Content!.Text.Should().Be("Chapter 1");
        responses[0].Elements[1].Level.Should().Be(2);
        responses[0].Elements[2].Level.Should().Be(2);
        responses[0].Elements[3].Level.Should().Be(1);
    }

    // ─── Edge cases ─────────────────────────────────────────────

    [Fact]
    public void Inspect_EmptyDocument_ReturnsZeroMatches()
    {
        using var doc = CreateInMemoryDocument(body => { });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses.Should().HaveCount(1);
        responses[0].Matched.Should().Be(0);
        responses[0].Elements.Should().BeEmpty();
    }

    [Fact]
    public void Inspect_SingleElementDocument_ReturnsOneMatch()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Only paragraph"));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph\n  INCLUDE content\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Matched.Should().Be(1);
        responses[0].Elements[0].Index.Should().Be(1);
        responses[0].Elements[0].Of.Should().Be(1);
        responses[0].Elements[0].Content!.Text.Should().Be("Only paragraph");
    }

    [Fact]
    public void Inspect_ContextOnSingleElement_ClampsCorrectly()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Only paragraph"));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph[1]\n  CONTEXT 5\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Context.Should().NotBeNull();
        responses[0].Elements[0].Context!.Before.Should().BeEmpty();
        responses[0].Elements[0].Context!.After.Should().BeEmpty();
    }

    [Fact]
    public void Inspect_DepthBeyondStructure_ReturnsAvailableLevels()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(
                new[] { "A", "B" }
            ));
        });

        // DEPTH 5 on a simple table — should stop at cells (depth 2)
        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/table[1]\n  DEPTH 5\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        var table = responses[0].Elements[0];
        table.Children.Should().HaveCount(1); // 1 row
        table.Children![0].Children.Should().HaveCount(2); // 2 cells
    }

    [Fact]
    public void Inspect_LargeContext_ReturnsAllAvailable()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Para 1"));
            body.AppendChild(MakeHeading("Middle", 1));
            body.AppendChild(MakeParagraph("Para 2"));
        });

        // CONTEXT 100 — should clamp to available elements
        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading[1]\n  CONTEXT 100\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Context.Should().NotBeNull();
        responses[0].Elements[0].Context!.Before.Should().HaveCount(1);
        responses[0].Elements[0].Context!.After.Should().HaveCount(1);
    }

    [Fact]
    public void Inspect_OverlappingAddresses_ReturnsSeparateResponses()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Title", 1));
            body.AppendChild(MakeParagraph("Body"));
        });

        var otk = ParseOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading[1]
  INCLUDE content

INSPECT body/heading
  INCLUDE content
");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses.Should().HaveCount(2);
        responses[0].Matched.Should().Be(1);
        responses[0].Elements[0].Content!.Text.Should().Be("Title");
        responses[1].Matched.Should().Be(1);
        responses[1].Elements[0].Content!.Text.Should().Be("Title");
    }

    [Fact]
    public void Inspect_LargeDocument_HandlesCorrectly()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            for (int i = 0; i < 200; i++)
            {
                if (i % 10 == 0)
                    body.AppendChild(MakeHeading($"Section {i / 10 + 1}", 1));
                else
                    body.AppendChild(MakeParagraph($"Paragraph {i}"));
            }
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n  INCLUDE content\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Matched.Should().Be(20);
        responses[0].Elements.Should().HaveCount(20);
        responses[0].Elements[0].Content!.Text.Should().Be("Section 1");
        responses[0].Elements[19].Content!.Text.Should().Be("Section 20");
    }

    [Fact]
    public void Inspect_TableCells_DirectAddressing()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(
                new[] { "Name", "Value" },
                new[] { "Alpha", "100" },
                new[] { "Beta", "200" }
            ));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/table[1]/row[2]\n  DEPTH 1\n  INCLUDE content\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Matched.Should().Be(1);
        var row = responses[0].Elements[0];
        row.Type.Should().Be("row");
        row.Children.Should().HaveCount(2);
        row.Children![0].Content!.Text.Should().Be("Alpha");
        row.Children![1].Content!.Text.Should().Be("100");
    }

    [Fact]
    public void Inspect_MultipleTablesInDocument_PositionalIndexing()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(new[] { "Table1-A", "Table1-B" }));
            body.AppendChild(MakeParagraph("Separator"));
            body.AppendChild(MakeTable(new[] { "Table2-A", "Table2-B" }));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/table\n  DEPTH 2\n  INCLUDE content\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Matched.Should().Be(2);
        responses[0].Elements[0].Index.Should().Be(1);
        responses[0].Elements[0].Of.Should().Be(2);
        responses[0].Elements[1].Index.Should().Be(2);
        responses[0].Elements[1].Of.Should().Be(2);
        responses[0].Elements[0].Children![0].Children![0].Content!.Text.Should().Be("Table1-A");
        responses[0].Elements[1].Children![0].Children![0].Content!.Text.Should().Be("Table2-A");
    }

    [Fact]
    public void Inspect_DepthZeroExplicit_NoChildren()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeTable(
                new[] { "A", "B" },
                new[] { "C", "D" }
            ));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/table[1]\n  DEPTH 0\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Children.Should().BeNull();
    }

    [Fact]
    public void Inspect_IncludeContentWithoutProperties_PropertiesNull()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Title", 1));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading[1]\n  INCLUDE content\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Content.Should().NotBeNull();
        responses[0].Elements[0].Properties.Should().BeNull();
    }

    [Fact]
    public void Inspect_IncludePropertiesWithoutContent_ContentNull()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Title", 1));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading[1]\n  INCLUDE properties\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Content.Should().BeNull();
        responses[0].Elements[0].Properties.Should().NotBeNull();
    }

    [Fact]
    public void Inspect_NoInclude_BothContentAndPropertiesNull()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Title", 1));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading[1]\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Elements[0].Content.Should().BeNull();
        responses[0].Elements[0].Properties.Should().BeNull();
        // Should still have addressing info
        responses[0].Elements[0].Type.Should().Be("heading");
        responses[0].Elements[0].Level.Should().Be(1);
    }

    [Fact]
    public void Inspect_ResponseAddressMatchesInput()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Title", 1));
            body.AppendChild(MakeParagraph("Text"));
        });

        var otk = ParseOtk("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading[1]\n  INCLUDE content\n");
        var inspector = new WordInspector(doc);
        var responses = inspector.Inspect(otk);

        responses[0].Address.Should().Contain("heading");
    }

    // ─── JSONL serialization edge cases ─────────────────────────

    [Fact]
    public void JsonlWriter_MultipleResponses_WritesOneLine​PerResponse()
    {
        var writer = new StringWriter();
        var jsonl = new JsonlResponseWriter(writer);

        jsonl.WriteInspectResponse(new InspectResponse
        {
            Address = "body/heading",
            Matched = 1,
            Elements = new List<ElementInfo>
            {
                new() { Type = "heading", Index = 1, Of = 1, Level = 1 }
            }
        });

        jsonl.WriteInspectResponse(new InspectResponse
        {
            Address = "body/paragraph",
            Matched = 0,
            Elements = new List<ElementInfo>()
        });

        jsonl.Flush();
        var lines = writer.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries);
        lines.Should().HaveCount(2);

        // Each line should be valid JSON
        foreach (var line in lines)
        {
            var action = () => JsonDocument.Parse(line);
            action.Should().NotThrow();
        }
    }

    [Fact]
    public void JsonlWriter_EmptyElements_SerializesAsEmptyArray()
    {
        var writer = new StringWriter();
        var jsonl = new JsonlResponseWriter(writer);

        jsonl.WriteInspectResponse(new InspectResponse
        {
            Address = "body/heading",
            Matched = 0,
            Elements = new List<ElementInfo>()
        });
        jsonl.Flush();

        var json = JsonDocument.Parse(writer.ToString().Trim());
        json.RootElement.GetProperty("elements").GetArrayLength().Should().Be(0);
        json.RootElement.GetProperty("matched").GetInt32().Should().Be(0);
    }

    [Fact]
    public void JsonlWriter_ErrorResponse_IncludesErrorField()
    {
        var writer = new StringWriter();
        var jsonl = new JsonlResponseWriter(writer);

        jsonl.WriteInspectResponse(new InspectResponse
        {
            Address = "body/nonexistent",
            Matched = 0,
            Elements = new List<ElementInfo>(),
            Error = "Address resolution failed"
        });
        jsonl.Flush();

        var json = JsonDocument.Parse(writer.ToString().Trim());
        json.RootElement.GetProperty("error").GetString().Should().Be("Address resolution failed");
    }

    [Fact]
    public void JsonlWriter_ContextWithEmptyBeforeAfter_Serializes()
    {
        var writer = new StringWriter();
        var jsonl = new JsonlResponseWriter(writer);

        jsonl.WriteInspectResponse(new InspectResponse
        {
            Address = "body/paragraph[1]",
            Matched = 1,
            Elements = new List<ElementInfo>
            {
                new()
                {
                    Type = "paragraph",
                    Index = 1,
                    Of = 1,
                    Context = new ContextInfo
                    {
                        Before = new List<ElementInfo>(),
                        After = new List<ElementInfo>()
                    }
                }
            }
        });
        jsonl.Flush();

        var json = JsonDocument.Parse(writer.ToString().Trim());
        var element = json.RootElement.GetProperty("elements")[0];
        element.GetProperty("context").GetProperty("before").GetArrayLength().Should().Be(0);
        element.GetProperty("context").GetProperty("after").GetArrayLength().Should().Be(0);
    }
}
