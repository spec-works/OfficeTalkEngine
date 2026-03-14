using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeTalk.Ast;
using OfficeTalkEngine.Addressing;
using Xunit;

namespace OfficeTalkEngine.Tests.Addressing;

public class WordAddressResolverTests
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
    public void Paragraph_positional_resolves_first()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("First"));
            body.AppendChild(MakeParagraph("Second"));
            body.AppendChild(MakeParagraph("Third"));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("paragraph", new PositionalPredicate { Position = 1 }));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
        result[0].InnerText.Should().Be("First");
    }

    [Fact]
    public void Paragraph_text_equals_resolves_exact_match()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Alpha"));
            body.AppendChild(MakeParagraph("Beta"));
            body.AppendChild(MakeParagraph("Gamma"));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("paragraph",
            new KeyValuePredicate { Key = "text", Operator = PredicateOperator.Equals, Value = "Beta" }));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
        result[0].InnerText.Should().Be("Beta");
    }

    [Fact]
    public void Paragraph_text_contains_resolves_partial_match()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Hello World"));
            body.AppendChild(MakeParagraph("Goodbye World"));
            body.AppendChild(MakeParagraph("Hello Again"));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("paragraph",
            new KeyValuePredicate { Key = "text", Operator = PredicateOperator.AsteriskEquals, Value = "World" }));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(2);
        result[0].InnerText.Should().Be("Hello World");
        result[1].InnerText.Should().Be("Goodbye World");
    }

    [Fact]
    public void Heading_level1_resolves()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Title", 1));
            body.AppendChild(MakeParagraph("Body text"));
            body.AppendChild(MakeHeading("Subtitle", 2));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("heading",
            new KeyValuePredicate { Key = "level", Operator = PredicateOperator.Equals, Value = "1" }));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
        result[0].InnerText.Should().Be("Title");
    }

    [Fact]
    public void Heading_level2_with_text_resolves()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Section A", 2));
            body.AppendChild(MakeHeading("Section B", 2));
            body.AppendChild(MakeHeading("Chapter", 1));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("heading",
            new KeyValuePredicate { Key = "level", Operator = PredicateOperator.Equals, Value = "2" },
            new KeyValuePredicate { Key = "text", Operator = PredicateOperator.Equals, Value = "Section B" }));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
        result[0].InnerText.Should().Be("Section B");
    }

    [Fact]
    public void Table_row_cell_resolves()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            var table = new Table(
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("R1C1")))),
                    new TableCell(new Paragraph(new Run(new Text("R1C2"))))),
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("R2C1")))),
                    new TableCell(new Paragraph(new Run(new Text("R2C2"))))));
            body.AppendChild(table);
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(
            Seg("table", new PositionalPredicate { Position = 1 }),
            Seg("row", new PositionalPredicate { Position = 1 }),
            Seg("cell", new PositionalPredicate { Position = 2 }));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
        result[0].InnerText.Should().Be("R1C2");
    }

    [Fact]
    public void Address_not_found_returns_empty()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Only paragraph"));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("paragraph", new PositionalPredicate { Position = 5 }));

        var result = resolver.Resolve(address);

        result.Should().BeEmpty();
    }

    [Fact]
    public void Ambiguous_address_returns_multiple_elements()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Same"));
            body.AppendChild(MakeParagraph("Same"));
            body.AppendChild(MakeParagraph("Same"));
        });

        var resolver = new WordAddressResolver(doc);
        // No positional predicate — matches all paragraphs
        var address = MakeAddress(Seg("paragraph",
            new KeyValuePredicate { Key = "text", Operator = PredicateOperator.Equals, Value = "Same" }));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(3);
    }

    [Fact]
    public void Paragraphs_exclude_headings()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Heading", 1));
            body.AppendChild(MakeParagraph("Body"));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("paragraph"));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
        result[0].InnerText.Should().Be("Body");
    }
}
