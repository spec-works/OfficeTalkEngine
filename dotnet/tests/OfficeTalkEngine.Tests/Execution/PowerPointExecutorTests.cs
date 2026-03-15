using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using OfficeTalk.Ast;
using OfficeTalkEngine.Execution;
using Xunit;
using D = DocumentFormat.OpenXml.Drawing;

namespace OfficeTalkEngine.Tests.Execution;

public class PowerPointExecutorTests
{
    #region Helpers

    private static PresentationDocument CreateInMemoryPresentation(Action<PresentationPart>? configure = null)
    {
        var stream = new MemoryStream();
        var doc = PresentationDocument.Create(stream, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        presentationPart.Presentation = new Presentation
        {
            SlideIdList = new SlideIdList(),
            SlideSize = new SlideSize { Cx = 9144000, Cy = 6858000 }
        };

        configure?.Invoke(presentationPart);

        presentationPart.Presentation.Save();
        return doc;
    }

    private static SlidePart AddSlide(PresentationPart presentationPart, uint slideId,
        string? titleText = null, string? bodyText = null)
    {
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        var slide = new Slide(new CommonSlideData(new ShapeTree()));
        slidePart.Slide = slide;

        var slideIdList = presentationPart.Presentation.SlideIdList!;
        slideIdList.AppendChild(new SlideId
        {
            Id = slideId,
            RelationshipId = presentationPart.GetIdOfPart(slidePart)
        });

        if (titleText != null)
        {
            AddShape(slide, titleText, PlaceholderValues.Title);
        }
        if (bodyText != null)
        {
            AddShape(slide, bodyText, PlaceholderValues.Body);
        }

        return slidePart;
    }

    private static void AddShape(Slide slide, string text, PlaceholderValues phType)
    {
        var shape = new Shape();
        shape.NonVisualShapeProperties = new NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = (uint)(phType == PlaceholderValues.Title ? 2 : 3), Name = phType.ToString() },
            new NonVisualShapeDrawingProperties(),
            new ApplicationNonVisualDrawingProperties(
                new PlaceholderShape { Type = phType }));

        shape.TextBody = new TextBody(
            new D.Paragraph(
                new D.Run(
                    new D.Text(text))));

        slide.CommonSlideData!.ShapeTree!.AppendChild(shape);
    }

    private static OfficeTalkDocument MakeDocument(params OperationBlock[] blocks)
    {
        return new OfficeTalkDocument
        {
            Version = "1.0",
            DocType = DocType.PowerPoint,
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
    public void Comment_adds_comment_to_slide()
    {
        using var doc = CreateInMemoryPresentation(pp =>
        {
            AddSlide(pp, 256, "Overview", "Some content");
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("slide", new PositionalPredicate(1))),
                new CommentOperation { Content = new ContentValue("Review slide layout.") }));

        new PowerPointExecutor().Execute(otDoc, doc);

        var presentationPart = doc.PresentationPart!;

        // Check comment authors
        var authorsPart = presentationPart.CommentAuthorsPart;
        authorsPart.Should().NotBeNull();
        var author = authorsPart!.CommentAuthorList
            .Elements<CommentAuthor>().First();
        author.Name!.Value.Should().Be("OfficeTalk");

        // Check slide comment
        var slidePart = presentationPart.SlideParts.First();
        var commentsPart = slidePart.SlideCommentsPart;
        commentsPart.Should().NotBeNull();

        var comments = commentsPart!.CommentList
            .Elements<Comment>().ToList();
        comments.Should().HaveCount(1);
        comments[0].InnerText.Should().Be("Review slide layout.");
    }

    [Fact]
    public void Set_updates_shape_text()
    {
        using var doc = CreateInMemoryPresentation(pp =>
        {
            AddSlide(pp, 256, "Old Title", "Old Body");
        });

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(
                    Seg("slide", new PositionalPredicate(1)),
                    Seg("title")),
                new SetOperation { Content = new ContentValue("New Title") }));

        new PowerPointExecutor().Execute(otDoc, doc);

        var slidePart = doc.PresentationPart!.SlideParts.First();
        var shapes = slidePart.Slide.Descendants<Shape>().ToList();
        var titleShape = shapes.FirstOrDefault(s =>
            s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<PlaceholderShape>()?.Type?.Value == PlaceholderValues.Title);

        titleShape.Should().NotBeNull();
        titleShape!.TextBody!.InnerText.Should().Be("New Title");
    }

    [Fact]
    public void Delete_removes_slide()
    {
        using var doc = CreateInMemoryPresentation(pp =>
        {
            AddSlide(pp, 256, "Slide 1");
            AddSlide(pp, 257, "Slide 2");
        });

        doc.PresentationPart!.Presentation.SlideIdList!
            .Elements<SlideId>().Should().HaveCount(2);

        var otDoc = MakeDocument(
            MakeBlock(
                MakeAddress(Seg("slide", new PositionalPredicate(2))),
                new DeleteOperation()));

        new PowerPointExecutor().Execute(otDoc, doc);

        doc.PresentationPart.Presentation.SlideIdList!
            .Elements<SlideId>().Should().HaveCount(1);
    }
}
