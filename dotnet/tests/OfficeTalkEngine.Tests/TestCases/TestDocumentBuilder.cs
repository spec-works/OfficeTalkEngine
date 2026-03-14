using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeTalkEngine.Tests.TestCases;

/// <summary>
/// Builds an in-memory WordprocessingDocument from a test fixture's input specification.
/// </summary>
public static class TestDocumentBuilder
{
    public static (MemoryStream Stream, WordprocessingDocument Document) Build(TestInputSpec input)
    {
        var stream = new MemoryStream();
        var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();
        mainPart.Document.Body = body;

        foreach (var element in input.Body)
        {
            switch (element.Type)
            {
                case "paragraph":
                    body.AppendChild(MakeParagraph(element.Text ?? "", element.Style));
                    break;
                case "heading":
                    body.AppendChild(MakeHeading(element.Text ?? "", element.Level ?? 1));
                    break;
                case "table":
                    body.AppendChild(MakeTable(element.Rows ?? new()));
                    break;
            }
        }

        mainPart.Document.Save();
        return (stream, doc);
    }

    private static Paragraph MakeParagraph(string text, string? style = null)
    {
        var paragraph = new Paragraph(
            new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

        if (style != null)
        {
            paragraph.PrependChild(new ParagraphProperties(
                new ParagraphStyleId { Val = style }));
        }

        return paragraph;
    }

    private static Paragraph MakeHeading(string text, int level)
    {
        return new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId { Val = $"Heading{level}" }),
            new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Table MakeTable(List<TableRowSpec> rows)
    {
        var table = new Table();
        foreach (var row in rows)
        {
            var tableRow = new TableRow();
            foreach (var cellText in row.Cells)
            {
                var cell = new TableCell(
                    new Paragraph(
                        new Run(new Text(cellText) { Space = SpaceProcessingModeValues.Preserve })));
                tableRow.AppendChild(cell);
            }
            table.AppendChild(tableRow);
        }
        return table;
    }
}
