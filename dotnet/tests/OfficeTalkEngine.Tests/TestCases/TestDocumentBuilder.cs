using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

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
                case "image":
                    body.AppendChild(MakeImageParagraph(mainPart, element.Alt ?? "image", ref nextBookmarkId));
                    break;
            }
        }

        // Add header/footer parts if any elements define them
        var headerElements = input.Body.Where(e => e.Type == "header").ToList();
        if (headerElements.Count > 0)
        {
            AddHeaderPart(mainPart, body, headerElements[0].Text ?? "");
        }
        var footerElements = input.Body.Where(e => e.Type == "footer").ToList();
        if (footerElements.Count > 0)
        {
            AddFooterPart(mainPart, body, footerElements[0].Text ?? "");
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

    private static Paragraph MakeImageParagraph(MainDocumentPart mainPart, string altText, ref int nextId)
    {
        // 1x1 white PNG
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVQI12NgAAIABQAB" +
            "Nl7BcQAAAABJRU5ErkJggg==");

        var imagePart = mainPart.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(pngBytes))
            imagePart.FeedData(ms);

        string relationshipId = mainPart.GetIdOfPart(imagePart);
        int id = ++nextId;

        var drawing = new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = 914400, Cy = 914400 }, // 1 inch
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DW.DocProperties { Id = (uint)id, Name = $"Image{id}", Description = altText },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = (uint)id, Name = $"Image{id}" },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0, Y = 0 },
                                    new A.Extents { Cx = 914400, Cy = 914400 }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                    ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            ) { DistanceFromTop = 0, DistanceFromBottom = 0, DistanceFromLeft = 0, DistanceFromRight = 0 });

        return new Paragraph(new Run(drawing));
    }

    private static void AddHeaderPart(MainDocumentPart mainPart, Body body, string headerText)
    {
        var headerPart = mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = new Header(
            new Paragraph(
                new Run(new Text(headerText) { Space = SpaceProcessingModeValues.Preserve })));
        headerPart.Header.Save();

        string headerPartId = mainPart.GetIdOfPart(headerPart);

        // Ensure body has SectionProperties with HeaderReference
        var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectPr == null)
        {
            sectPr = new SectionProperties();
            body.AppendChild(sectPr);
        }
        sectPr.PrependChild(new HeaderReference
        {
            Type = HeaderFooterValues.Default,
            Id = headerPartId
        });
    }

    private static void AddFooterPart(MainDocumentPart mainPart, Body body, string footerText)
    {
        var footerPart = mainPart.AddNewPart<FooterPart>();
        footerPart.Footer = new Footer(
            new Paragraph(
                new Run(new Text(footerText) { Space = SpaceProcessingModeValues.Preserve })));
        footerPart.Footer.Save();

        string footerPartId = mainPart.GetIdOfPart(footerPart);

        var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectPr == null)
        {
            sectPr = new SectionProperties();
            body.AppendChild(sectPr);
        }
        sectPr.PrependChild(new FooterReference
        {
            Type = HeaderFooterValues.Default,
            Id = footerPartId
        });
    }
}
