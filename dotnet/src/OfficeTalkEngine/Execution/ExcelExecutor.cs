using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeTalk.Ast;
using OfficeTalkEngine.Addressing;

namespace OfficeTalkEngine.Execution;

/// <summary>
/// Executes OfficeTalk operations on Excel documents using snapshot semantics.
/// </summary>
public class ExcelExecutor : IOfficeTalkExecutor
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

        using (var spreadsheetDoc = SpreadsheetDocument.Open(memoryStream, true))
        {
            Execute(document, spreadsheetDoc);
            spreadsheetDoc.Save();
        }

        var data = memoryStream.ToArray();
        using (var fs = new FileStream(workingPath, FileMode.Open, FileAccess.Write, FileShare.None))
        {
            fs.Write(data, 0, data.Length);
            fs.SetLength(data.Length);
        }
    }

    /// <summary>
    /// Executes operations against an already-open SpreadsheetDocument (for testing).
    /// </summary>
    public void Execute(OfficeTalkDocument document, SpreadsheetDocument spreadsheetDoc)
    {
        var resolver = new ExcelAddressResolver(spreadsheetDoc);

        // Snapshot mode: resolve all addresses before mutation
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
                    ExecuteOperation(spreadsheetDoc, element, operation);
                }
            }
        }
    }

    private static void ExecuteOperation(
        SpreadsheetDocument doc, OpenXmlElement element, Operation operation)
    {
        switch (operation)
        {
            case SetOperation set:
                ExecuteSet(doc, element, set);
                break;
            case DeleteOperation:
                ExecuteDelete(element);
                break;
            case CommentOperation comment:
                ExecuteComment(doc, element, comment);
                break;
            default:
                throw new NotSupportedException(
                    $"Operation type '{operation.GetType().Name}' is not yet supported for Excel documents.");
        }
    }

    private static void ExecuteSet(SpreadsheetDocument doc, OpenXmlElement element, SetOperation operation)
    {
        if (element is Cell cell)
        {
            cell.CellValue = new CellValue(operation.Content.Text);
            // If the value looks numeric, set type accordingly; otherwise use string
            if (double.TryParse(operation.Content.Text, out _))
            {
                cell.DataType = null; // Numeric — no explicit type needed
            }
            else
            {
                cell.DataType = CellValues.InlineString;
                cell.CellValue = null;
                cell.InlineString = new InlineString(
                    new Text(operation.Content.Text));
            }
        }
        else if (element is Row row)
        {
            // SET on a row — set all cells
            var cells = row.Elements<Cell>().ToList();
            foreach (var c in cells)
            {
                c.CellValue = new CellValue(operation.Content.Text);
            }
        }
        else
        {
            throw new NotSupportedException(
                $"SET is not supported on {element.GetType().Name} in Excel.");
        }
    }

    private static void ExecuteDelete(OpenXmlElement element)
    {
        element.Remove();
    }

    private static void ExecuteComment(
        SpreadsheetDocument doc, OpenXmlElement element, CommentOperation operation)
    {
        if (element is not Cell cell)
            throw new NotSupportedException(
                "COMMENT in Excel can only be applied to cells.");

        var cellRef = cell.CellReference?.Value
            ?? throw new InvalidOperationException("Cell has no reference.");

        // Find which worksheet part this cell belongs to
        var worksheetPart = cell.Ancestors<Worksheet>().FirstOrDefault()?.WorksheetPart
            ?? FindWorksheetPart(doc, cell);

        if (worksheetPart == null)
            throw new InvalidOperationException("Could not find worksheet for cell.");

        // Ensure WorksheetCommentsPart exists
        var commentsPart = worksheetPart.WorksheetCommentsPart;
        if (commentsPart == null)
        {
            commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();
            commentsPart.Comments = new Comments(
                new Authors(new Author("OfficeTalk")),
                new CommentList());
        }

        var comments = commentsPart.Comments;
        var authors = comments.GetFirstChild<Authors>()!;
        var commentList = comments.GetFirstChild<CommentList>()!;

        // Ensure "OfficeTalk" author exists
        uint authorId = 0;
        var existingAuthors = authors.Elements<Author>().ToList();
        bool found = false;
        for (int i = 0; i < existingAuthors.Count; i++)
        {
            if (existingAuthors[i].Text == "OfficeTalk")
            {
                authorId = (uint)i;
                found = true;
                break;
            }
        }
        if (!found)
        {
            authorId = (uint)existingAuthors.Count;
            authors.AppendChild(new Author("OfficeTalk"));
        }

        // Create the comment
        var comment = new Comment
        {
            Reference = cellRef,
            AuthorId = authorId
        };

        var commentText = new CommentText();
        var lines = operation.Content.Text.Split('\n');
        foreach (var line in lines)
        {
            var run = new Run();
            run.AppendChild(new Text(line.TrimEnd('\r')));
            commentText.AppendChild(run);
        }
        comment.AppendChild(commentText);
        commentList.AppendChild(comment);
    }

    private static WorksheetPart? FindWorksheetPart(SpreadsheetDocument doc, OpenXmlElement element)
    {
        // Walk up the tree to find the Worksheet, then its part
        var worksheet = element.Ancestors<Worksheet>().FirstOrDefault();
        if (worksheet == null) return null;

        var workbookPart = doc.WorkbookPart;
        if (workbookPart == null) return null;

        foreach (var wsPart in workbookPart.WorksheetParts)
        {
            if (wsPart.Worksheet == worksheet) return wsPart;
        }

        return null;
    }
}
