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

        // Add style definitions so Word COM recognizes heading outline levels
        AddStyleDefinitions(mainPart, input.Body);

        // Add numbering definitions if list styles are used
        if (input.Body.Any(e => e.Style == "ListBullet"))
            AddNumberingDefinitions(mainPart);

        int nextBookmarkId = 0;

        foreach (var element in input.Body)
        {
            switch (element.Type)
            {
                case "paragraph":
                    var para = MakeParagraph(element.Text ?? "", element.Style);
                    if (element.BookmarkName != null)
                    {
                        var bmId = (nextBookmarkId++).ToString();
                        para.PrependChild(new BookmarkStart { Name = element.BookmarkName, Id = bmId });
                        para.AppendChild(new BookmarkEnd { Id = bmId });
                    }
                    body.AppendChild(para);
                    break;
                case "heading":
                    body.AppendChild(MakeHeading(element.Text ?? "", element.Level ?? 1));
                    break;
                case "table":
                    body.AppendChild(MakeTable(element.Rows ?? new()));
                    break;
                case "content-control":
                    var sdtBlock = new SdtBlock();
                    var sdtPr = new SdtProperties();
                    if (element.Tag != null)
                        sdtPr.AppendChild(new Tag { Val = element.Tag });
                    sdtBlock.AppendChild(sdtPr);
                    var sdtContent = new SdtContentBlock();
                    sdtContent.AppendChild(MakeParagraph(element.Text ?? ""));
                    sdtBlock.AppendChild(sdtContent);
                    body.AppendChild(sdtBlock);
                    break;
                case "section-break":
                    var sbPara = new Paragraph(
                        new ParagraphProperties(
                            new SectionProperties()));
                    body.AppendChild(sbPara);
                    break;
            }
        }

        mainPart.Document.Save();
        return (stream, doc);
    }

    private static void AddStyleDefinitions(MainDocumentPart mainPart, List<BodyElement> elements)
    {
        var headingLevels = elements
            .Where(e => e.Type == "heading")
            .Select(e => e.Level ?? 1)
            .Distinct()
            .ToList();

        var styles = elements
            .Where(e => e.Style != null)
            .Select(e => e.Style!)
            .Distinct()
            .ToList();

        if (headingLevels.Count == 0 && styles.Count == 0)
            return;

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var stylesRoot = new Styles();

        // Heading styles with outline levels — required for Word COM to recognize headings
        foreach (var level in headingLevels)
        {
            var outlineLevel = new OutlineLevel { Val = level - 1 }; // 0-based
            var style = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = $"Heading{level}",
                StyleName = new StyleName { Val = $"heading {level}" },
                PrimaryStyle = new PrimaryStyle(),
            };
            style.AppendChild(new StyleParagraphProperties(outlineLevel));
            stylesRoot.AppendChild(style);
        }

        // Named paragraph styles (e.g., ListBullet)
        foreach (var styleName in styles)
        {
            // Skip if already added as a heading
            if (headingLevels.Any(l => $"Heading{l}" == styleName))
                continue;

            stylesRoot.AppendChild(new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = styleName,
                StyleName = new StyleName { Val = styleName },
            });
        }

        stylesPart.Styles = stylesRoot;
        stylesPart.Styles.Save();
    }

    private static Paragraph MakeParagraph(string text, string? style = null)
    {
        var paragraph = new Paragraph(
            new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

        if (style != null)
        {
            var paraProps = new ParagraphProperties(
                new ParagraphStyleId { Val = style });

            if (style == "ListBullet")
            {
                paraProps.AppendChild(new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 1 }));
            }

            paragraph.PrependChild(paraProps);
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

    private static void AddNumberingDefinitions(MainDocumentPart mainPart)
    {
        var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
        var numbering = new Numbering();

        var abstractNum = new AbstractNum(
            new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.Bullet },
                new LevelText { Val = "\u00B7" })
            { LevelIndex = 0 })
        { AbstractNumberId = 1 };

        numbering.AppendChild(abstractNum);
        numbering.AppendChild(new NumberingInstance(
            new AbstractNumId { Val = 1 })
        { NumberID = 1 });

        numberingPart.Numbering = numbering;
        numberingPart.Numbering.Save();
    }
}
