using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeTalk.Ast;
using OfficeTalkEngine.Addressing;
using D = DocumentFormat.OpenXml.Drawing;

namespace OfficeTalkEngine.Execution;

/// <summary>
/// Executes OfficeTalk operations on PowerPoint documents using snapshot semantics.
/// </summary>
public class PowerPointExecutor : IOfficeTalkExecutor
{
    public void Execute(OfficeTalkDocument document, string targetPath, string? outputPath = null)
    {
        string workingPath = outputPath ?? targetPath;
        if (outputPath != null && outputPath != targetPath)
            File.Copy(targetPath, outputPath, overwrite: true);

        using var memoryStream = new MemoryStream();
        using (var fs = new FileStream(workingPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            fs.CopyTo(memoryStream);
        }
        memoryStream.Position = 0;

        using (var presentationDoc = PresentationDocument.Open(memoryStream, true))
        {
            Execute(document, presentationDoc);
            presentationDoc.Save();
        }

        var data = memoryStream.ToArray();
        using (var fs = new FileStream(workingPath, FileMode.Open, FileAccess.Write, FileShare.None))
        {
            fs.Write(data, 0, data.Length);
            fs.SetLength(data.Length);
        }
    }

    /// <summary>
    /// Executes operations against an already-open PresentationDocument (for testing).
    /// </summary>
    public void Execute(OfficeTalkDocument document, PresentationDocument presentationDoc)
    {
        var resolver = new PowerPointAddressResolver(presentationDoc);

        var resolvedBlocks = new List<(OperationBlock Block, IReadOnlyList<OpenXmlElement> Elements)>();
        foreach (var block in document.OperationBlocks)
        {
            var elements = resolver.Resolve(block.Address);
            resolvedBlocks.Add((block, elements));
        }

        foreach (var (block, elements) in resolvedBlocks)
        {
            foreach (var element in elements)
            {
                foreach (var operation in block.Operations)
                {
                    ExecuteOperation(presentationDoc, element, operation);
                }
            }
        }
    }

    private static void ExecuteOperation(
        PresentationDocument doc, OpenXmlElement element, Operation operation)
    {
        switch (operation)
        {
            case SetOperation set:
                ExecuteSet(element, set);
                break;
            case DeleteOperation:
                ExecuteDelete(element);
                break;
            case CommentOperation comment:
                ExecuteComment(doc, element, comment);
                break;
            default:
                throw new NotSupportedException(
                    $"Operation type '{operation.GetType().Name}' is not yet supported for PowerPoint documents.");
        }
    }

    private static void ExecuteSet(OpenXmlElement element, SetOperation operation)
    {
        if (element is Shape shape)
        {
            // Clear existing text body and set new text
            var textBody = shape.TextBody;
            if (textBody == null)
            {
                textBody = new TextBody();
                shape.AppendChild(textBody);
            }

            // Remove all existing paragraphs
            foreach (var existing in textBody.Elements<D.Paragraph>().ToList())
                existing.Remove();

            // Add new paragraphs from content
            var lines = operation.Content.Text.Split('\n');
            foreach (var line in lines)
            {
                var para = new D.Paragraph(
                    new D.Run(
                        new D.Text(line.TrimEnd('\r'))));
                textBody.AppendChild(para);
            }
        }
        else if (element is Slide slide)
        {
            throw new NotSupportedException(
                "SET on a slide element requires targeting a specific shape (title, subtitle, body).");
        }
        else
        {
            throw new NotSupportedException(
                $"SET is not supported on {element.GetType().Name} in PowerPoint.");
        }
    }

    private static void ExecuteDelete(OpenXmlElement element)
    {
        if (element is Slide slide)
        {
            // Delete the slide from the presentation
            var slidePart = slide.SlidePart;
            if (slidePart == null) return;

            var presentationPart = slidePart.GetParentParts()
                .OfType<PresentationPart>().FirstOrDefault();
            if (presentationPart == null) return;

            var slideIdList = presentationPart.Presentation.SlideIdList;
            if (slideIdList == null) return;

            var slideId = slideIdList.Elements<SlideId>()
                .FirstOrDefault(sid =>
                    presentationPart.GetPartById(sid.RelationshipId!) == slidePart);
            slideId?.Remove();
            presentationPart.DeletePart(slidePart);
        }
        else
        {
            element.Remove();
        }
    }

    private static void ExecuteComment(
        PresentationDocument doc, OpenXmlElement element, CommentOperation operation)
    {
        // Find the slide part that contains this element
        SlidePart? slidePart;
        if (element is Slide slide)
        {
            slidePart = slide.SlidePart;
        }
        else
        {
            slidePart = element.Ancestors<Slide>().FirstOrDefault()?.SlidePart;
            if (slidePart == null)
            {
                // Try walking up to find any part
                var current = element;
                while (current != null)
                {
                    if (current is Slide s)
                    {
                        slidePart = s.SlidePart;
                        break;
                    }
                    current = current.Parent;
                }
            }
        }

        if (slidePart == null)
            throw new InvalidOperationException("Could not find slide for the addressed element.");

        var presentationPart = doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation has no presentation part.");

        // Ensure comment authors part exists
        var authorsPart = presentationPart.CommentAuthorsPart;
        if (authorsPart == null)
        {
            authorsPart = presentationPart.AddNewPart<CommentAuthorsPart>();
            authorsPart.CommentAuthorList = new CommentAuthorList();
        }

        var authorList = authorsPart.CommentAuthorList;

        // Find or create "OfficeTalk" author
        uint authorId = 0;
        uint lastIdx = 0;
        var existingAuthor = authorList.Elements<CommentAuthor>()
            .FirstOrDefault(a => a.Name?.Value == "OfficeTalk");

        if (existingAuthor != null)
        {
            authorId = existingAuthor.Id!.Value;
            lastIdx = existingAuthor.LastIndex!.Value;
        }
        else
        {
            var allAuthors = authorList.Elements<CommentAuthor>().ToList();
            authorId = allAuthors.Count > 0
                ? allAuthors.Max(a => a.Id!.Value) + 1 : 0;

            authorList.AppendChild(new CommentAuthor
            {
                Id = authorId,
                Name = "OfficeTalk",
                Initials = "OT",
                LastIndex = 0
            });
        }

        // Ensure slide has a comments part
        var slideCommentsPart = slidePart.SlideCommentsPart;
        if (slideCommentsPart == null)
        {
            slideCommentsPart = slidePart.AddNewPart<SlideCommentsPart>();
            slideCommentsPart.CommentList = new CommentList();
        }

        var commentList = slideCommentsPart.CommentList;

        // Generate unique comment index
        uint commentIdx = lastIdx + 1;

        // Update author's last index
        var author = authorList.Elements<CommentAuthor>()
            .First(a => a.Name?.Value == "OfficeTalk");
        author.LastIndex = commentIdx;

        // Create the comment — positioned at 0,0 (top-left of slide)
        var comment = new Comment
        {
            AuthorId = authorId,
            DateTime = DateTime.UtcNow,
            Index = commentIdx
        };

        comment.Position = new DocumentFormat.OpenXml.Presentation.Position
        {
            X = 0,
            Y = 0
        };

        comment.AppendChild(new Text(operation.Content.Text));
        commentList.AppendChild(comment);
    }
}
