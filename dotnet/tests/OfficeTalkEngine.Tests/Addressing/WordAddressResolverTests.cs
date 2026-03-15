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

    private static Paragraph MakeListItem(string text, int numId, int level = 0)
    {
        return new Paragraph(
            new ParagraphProperties(
                new NumberingProperties(
                    new NumberingId { Val = numId },
                    new NumberingLevelReference { Val = level })),
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

    private static KeyValuePredicate KV(string key, PredicateOperator op, string value) =>
        new() { Key = key, Operator = op, Value = value };

    private static PositionalPredicate Pos(int n) => new() { Position = n };

    #endregion

    #region Existing Tests — Paragraphs & Headings

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
        var address = MakeAddress(Seg("paragraph", Pos(1)));

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
        var address = MakeAddress(Seg("paragraph", KV("text", PredicateOperator.Equals, "Beta")));

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
        var address = MakeAddress(Seg("paragraph", KV("text", PredicateOperator.AsteriskEquals, "World")));

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
        var address = MakeAddress(Seg("heading", KV("level", PredicateOperator.Equals, "1")));

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
            KV("level", PredicateOperator.Equals, "2"),
            KV("text", PredicateOperator.Equals, "Section B")));

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
        var address = MakeAddress(Seg("table", Pos(1)), Seg("row", Pos(1)), Seg("cell", Pos(2)));

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
        var address = MakeAddress(Seg("paragraph", Pos(5)));

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
        var address = MakeAddress(Seg("paragraph", KV("text", PredicateOperator.Equals, "Same")));

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

    #endregion

    #region Heading-Scoped Section Addressing

    [Fact]
    public void Heading_scope_resolves_paragraph_within_section()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Introduction", 1));
            body.AppendChild(MakeParagraph("Intro text"));
            body.AppendChild(MakeHeading("Methods", 1));
            body.AppendChild(MakeParagraph("Methods text"));
            body.AppendChild(MakeParagraph("More methods"));
            body.AppendChild(MakeHeading("Results", 1));
            body.AppendChild(MakeParagraph("Results text"));
        });

        var resolver = new WordAddressResolver(doc);
        // body/heading[text="Methods"]/paragraph[1]
        var address = MakeAddress(
            Seg("body"),
            Seg("heading", KV("text", PredicateOperator.Equals, "Methods")),
            Seg("paragraph", Pos(1)));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
        result[0].InnerText.Should().Be("Methods text");
    }

    [Fact]
    public void Heading_scope_stops_at_same_level_heading()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Section A", 2));
            body.AppendChild(MakeParagraph("A paragraph 1"));
            body.AppendChild(MakeParagraph("A paragraph 2"));
            body.AppendChild(MakeHeading("Section B", 2));
            body.AppendChild(MakeParagraph("B paragraph 1"));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(
            Seg("heading", KV("text", PredicateOperator.Equals, "Section A")),
            Seg("paragraph"));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(2);
        result[0].InnerText.Should().Be("A paragraph 1");
        result[1].InnerText.Should().Be("A paragraph 2");
    }

    [Fact]
    public void Heading_scope_stops_at_higher_level_heading()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Chapter 1", 1));
            body.AppendChild(MakeHeading("Sub A", 2));
            body.AppendChild(MakeParagraph("Sub A content"));
            body.AppendChild(MakeHeading("Chapter 2", 1));
            body.AppendChild(MakeParagraph("Chapter 2 content"));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(
            Seg("heading", KV("text", PredicateOperator.Equals, "Sub A")),
            Seg("paragraph"));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
        result[0].InnerText.Should().Be("Sub A content");
    }

    [Fact]
    public void Heading_scope_extends_to_end_if_last_heading()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("First", 1));
            body.AppendChild(MakeParagraph("Para 1"));
            body.AppendChild(MakeHeading("Last", 1));
            body.AppendChild(MakeParagraph("Last para 1"));
            body.AppendChild(MakeParagraph("Last para 2"));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(
            Seg("heading", KV("text", PredicateOperator.Equals, "Last")),
            Seg("paragraph"));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(2);
        result[0].InnerText.Should().Be("Last para 1");
        result[1].InnerText.Should().Be("Last para 2");
    }

    [Fact]
    public void Heading_scope_nested_resolves_subsection()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Chapter 1", 1));
            body.AppendChild(MakeHeading("Analysis", 2));
            body.AppendChild(MakeParagraph("Analysis content"));
            body.AppendChild(MakeHeading("Discussion", 2));
            body.AppendChild(MakeParagraph("Discussion content"));
            body.AppendChild(MakeHeading("Chapter 2", 1));
        });

        var resolver = new WordAddressResolver(doc);
        // body/heading[level=1, text="Chapter 1"]/heading[level=2, text="Analysis"]/paragraph[1]
        var address = MakeAddress(
            Seg("body"),
            Seg("heading",
                KV("level", PredicateOperator.Equals, "1"),
                KV("text", PredicateOperator.Equals, "Chapter 1")),
            Seg("heading",
                KV("level", PredicateOperator.Equals, "2"),
                KV("text", PredicateOperator.Equals, "Analysis")),
            Seg("paragraph", Pos(1)));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
        result[0].InnerText.Should().Be("Analysis content");
    }

    [Fact]
    public void Heading_scope_includes_table_within_section()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeHeading("Data", 1));
            body.AppendChild(new Table(
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("Cell1")))))));
            body.AppendChild(MakeHeading("Next", 1));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(
            Seg("heading", KV("text", PredicateOperator.Equals, "Data")),
            Seg("table", Pos(1)));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
    }

    #endregion

    #region Text Regex Predicate (text~=)

    [Fact]
    public void Text_tilde_equals_matches_regex()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Chapter 1: Introduction"));
            body.AppendChild(MakeParagraph("Some other text"));
            body.AppendChild(MakeParagraph("Chapter 2: Methods"));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("paragraph",
            KV("text", PredicateOperator.TildeEquals, "^Chapter \\d+")));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(2);
        result[0].InnerText.Should().StartWith("Chapter 1");
        result[1].InnerText.Should().StartWith("Chapter 2");
    }

    [Fact]
    public void Text_tilde_equals_no_match_returns_empty()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Hello World"));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("paragraph",
            KV("text", PredicateOperator.TildeEquals, "^Goodbye")));

        var result = resolver.Resolve(address);

        result.Should().BeEmpty();
    }

    [Fact]
    public void Text_starts_with_resolves()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("In conclusion, we found"));
            body.AppendChild(MakeParagraph("Not a conclusion"));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("paragraph",
            KV("text", PredicateOperator.CaretEquals, "In conclusion")));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
        result[0].InnerText.Should().StartWith("In conclusion");
    }

    [Fact]
    public void Text_ends_with_resolves()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Results respectively."));
            body.AppendChild(MakeParagraph("Other text."));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("paragraph",
            KV("text", PredicateOperator.DollarEquals, "respectively.")));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
    }

    #endregion

    #region Run Segment

    [Fact]
    public void Run_resolves_within_paragraph()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(new Paragraph(
                new Run(new Text("First run")),
                new Run(new Text("Second run"))));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("paragraph", Pos(1)), Seg("run", Pos(2)));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
        result[0].InnerText.Should().Be("Second run");
    }

    #endregion

    #region List / Item Segments

    [Fact]
    public void List_item_resolves_numbered_paragraphs()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            body.AppendChild(MakeParagraph("Not a list"));
            body.AppendChild(MakeListItem("Item one", 1));
            body.AppendChild(MakeListItem("Item two", 1));
            body.AppendChild(MakeListItem("Item three", 1));
            body.AppendChild(MakeParagraph("After list"));
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("item", Pos(2)));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
        result[0].InnerText.Should().Be("Item two");
    }

    #endregion

    #region Bookmark Segment

    [Fact]
    public void Bookmark_resolves_by_name_predicate()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            var para = MakeParagraph("Referenced text");
            para.PrependChild(new BookmarkStart { Id = "1", Name = "references" });
            para.AppendChild(new BookmarkEnd { Id = "1" });
            body.AppendChild(para);
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("bookmark",
            KV("name", PredicateOperator.Equals, "references")));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
    }

    [Fact]
    public void Bookmark_bare_string_resolves()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            var para = MakeParagraph("Intro text");
            para.PrependChild(new BookmarkStart { Id = "1", Name = "intro" });
            para.AppendChild(new BookmarkEnd { Id = "1" });
            body.AppendChild(para);
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("bookmark", new BareStringPredicate("intro")));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
    }

    [Fact]
    public void Bookmark_excludes_internal_bookmarks()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            var para = MakeParagraph("Text");
            para.PrependChild(new BookmarkStart { Id = "1", Name = "_GoBack" });
            para.AppendChild(new BookmarkEnd { Id = "1" });
            para.PrependChild(new BookmarkStart { Id = "2", Name = "myBookmark" });
            para.AppendChild(new BookmarkEnd { Id = "2" });
            body.AppendChild(para);
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("bookmark"));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1); // Only "myBookmark", not "_GoBack"
    }

    #endregion

    #region Content Control Segment

    [Fact]
    public void Content_control_resolves_by_tag()
    {
        using var doc = CreateInMemoryDocument(body =>
        {
            var sdt = new SdtBlock(
                new SdtProperties(new Tag { Val = "abstract" }),
                new SdtContentBlock(new Paragraph(new Run(new Text("Abstract content")))));
            body.AppendChild(sdt);
        });

        var resolver = new WordAddressResolver(doc);
        var address = MakeAddress(Seg("content-control",
            KV("tag", PredicateOperator.Equals, "abstract")));

        var result = resolver.Resolve(address);

        result.Should().HaveCount(1);
    }

    #endregion
}
