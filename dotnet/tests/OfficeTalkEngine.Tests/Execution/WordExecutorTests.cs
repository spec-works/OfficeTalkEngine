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

    #endregion

    [Fact]
    public void Set_replaces_paragraph_text()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Original"));
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("paragraph", new PositionalPredicate { Position = 1 })),
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
                MakeAddress(Seg("paragraph", new PositionalPredicate { Position = 1 })),
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
                MakeAddress(Seg("paragraph", new PositionalPredicate { Position = 2 })),
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
                MakeAddress(Seg("paragraph", new PositionalPredicate { Position = 1 })),
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
                MakeAddress(Seg("paragraph", new PositionalPredicate { Position = 1 })),
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
                MakeAddress(Seg("paragraph", new PositionalPredicate { Position = 1 })),
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
                MakeAddress(Seg("paragraph", new PositionalPredicate { Position = 1 })),
                new ReplaceOperation { Search = "foo", Replacement = "qux", IsAll = true }));

        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        var paragraph = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First();
        paragraph.InnerText.Should().Be("qux bar qux baz qux");
    }
}
