using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;

namespace OfficeTalkEngine.Tests.TestCases;

/// <summary>
/// Reads a WordprocessingDocument and asserts it matches the expected test fixture output.
/// Comparison is driven by the expected element types: "heading" elements are matched by
/// checking the paragraph's HeadingN style, while "paragraph" elements may optionally
/// carry a style attribute.
/// </summary>
public static class TestDocumentReader
{
    public static void AssertExpected(WordprocessingDocument doc, TestExpectedOutput expected)
    {
        if (expected.Body != null)
        {
            AssertBody(doc, expected.Body);
        }

        if (expected.Properties != null)
        {
            AssertProperties(doc, expected.Properties);
        }
    }

    private static void AssertBody(WordprocessingDocument doc, List<BodyElement> expectedBody)
    {
        var body = doc.MainDocumentPart!.Document.Body!;

        // Collect actual body children as OpenXmlElements (paragraphs + tables + content controls)
        var actualChildren = body.ChildElements
            .Where(c => c is Paragraph or Table or SdtBlock)
            .ToList();

        actualChildren.Should().HaveCount(expectedBody.Count,
            "document body should have {0} elements", expectedBody.Count);

        for (int i = 0; i < expectedBody.Count; i++)
        {
            var exp = expectedBody[i];
            var actual = actualChildren[i];

            switch (exp.Type)
            {
                case "paragraph":
                    if (actual is SdtBlock sdtBlock)
                    {
                        var sdtPara = sdtBlock.SdtContentBlock?.GetFirstChild<Paragraph>();
                        sdtPara.Should().NotBeNull(
                            "element {0} should contain a paragraph inside content control", i);
                        AssertParagraph(sdtPara!, exp, i);
                    }
                    else
                    {
                        actual.Should().BeOfType<Paragraph>(
                            "element {0} should be a paragraph", i);
                        AssertParagraph((Paragraph)actual, exp, i);
                    }
                    break;
                case "heading":
                    actual.Should().BeOfType<Paragraph>(
                        "element {0} should be a paragraph (heading)", i);
                    AssertHeading((Paragraph)actual, exp, i);
                    break;
                case "table":
                    actual.Should().BeOfType<Table>(
                        "element {0} should be a table", i);
                    AssertTable((Table)actual, exp.Rows!, i);
                    break;
            }
        }
    }

    private static void AssertParagraph(Paragraph para, BodyElement expected, int index)
    {
        para.InnerText.Should().Be(expected.Text,
            "paragraph {0} text should match", index);

        if (expected.Style != null)
        {
            var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            styleId.Should().Be(expected.Style,
                "paragraph {0} style should match", index);
        }
    }

    private static void AssertHeading(Paragraph para, BodyElement expected, int index)
    {
        para.InnerText.Should().Be(expected.Text,
            "heading {0} text should match", index);

        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        styleId.Should().NotBeNull("heading {0} should have a style", index);
        styleId.Should().StartWith("Heading",
            "heading {0} should have a Heading style", index);

        if (expected.Level.HasValue)
        {
            int.TryParse(styleId!.AsSpan("Heading".Length), out int actualLevel)
                .Should().BeTrue("heading {0} style should contain a level number", index);
            actualLevel.Should().Be(expected.Level.Value,
                "heading {0} level should match", index);
        }
    }

    private static void AssertTable(Table table, List<TableRowSpec> expectedRows, int tableIndex)
    {
        var actualRows = table.Elements<TableRow>().ToList();
        actualRows.Should().HaveCount(expectedRows.Count,
            "table {0} should have {1} rows", tableIndex, expectedRows.Count);

        for (int r = 0; r < expectedRows.Count; r++)
        {
            var actualCells = actualRows[r].Elements<TableCell>()
                .Select(c => c.InnerText)
                .ToList();
            actualCells.Should().Equal(expectedRows[r].Cells,
                "table {0} row {1} cells should match", tableIndex, r);
        }
    }

    private static void AssertProperties(
        WordprocessingDocument doc, Dictionary<string, string> expected)
    {
        var props = doc.PackageProperties;
        foreach (var (key, value) in expected)
        {
            switch (key.ToLowerInvariant())
            {
                case "title":
                    props.Title.Should().Be(value, "document title should match");
                    break;
                case "author":
                    props.Creator.Should().Be(value, "document author should match");
                    break;
                case "subject":
                    props.Subject.Should().Be(value, "document subject should match");
                    break;
                case "description":
                    props.Description.Should().Be(value, "document description should match");
                    break;
                case "keywords":
                    props.Keywords.Should().Be(value, "document keywords should match");
                    break;
                case "category":
                    props.Category.Should().Be(value, "document category should match");
                    break;
            }
        }
    }
}
