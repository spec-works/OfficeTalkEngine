using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeTalk.Ast;
using OfficeTalkEngine.Addressing;

namespace OfficeTalkEngine.Execution;

/// <summary>
/// Executes OfficeTalk operations on Word documents using snapshot semantics.
/// All addresses are resolved before any mutations occur.
/// </summary>
public class WordExecutor : IOfficeTalkExecutor
{
    public void Execute(OfficeTalkDocument document, string targetPath, string? outputPath = null)
    {
        string workingPath = outputPath ?? targetPath;
        if (outputPath != null && outputPath != targetPath)
            File.Copy(targetPath, outputPath, overwrite: true);

        // Read file with FileShare.ReadWrite so we can read even if Word has it open
        using var memoryStream = new MemoryStream();
        using (var fs = new FileStream(workingPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            fs.CopyTo(memoryStream);
        }
        memoryStream.Position = 0;

        using (var wordDoc = WordprocessingDocument.Open(memoryStream, true))
        {
            // Phase 1: Snapshot — resolve all addresses before mutation
            var resolvedBlocks = new List<(OperationBlock Block, IReadOnlyList<OpenXmlElement> Elements)>();
            var resolver = new WordAddressResolver(wordDoc);

            foreach (var block in document.OperationBlocks)
            {
                var elements = resolver.Resolve(block.Address);
                resolvedBlocks.Add((block, elements));
            }

            // Phase 2: Execute operations against resolved elements
            foreach (var (block, elements) in resolvedBlocks)
            {
                foreach (var element in elements)
                {
                    foreach (var operation in block.Operations)
                    {
                        ExecuteOperation(wordDoc, element, operation);
                    }
                }
            }

            // Apply document-level property settings
            foreach (var prop in document.PropertySettings)
            {
                ApplyProperty(wordDoc, prop.Name, prop.Value);
            }

            wordDoc.Save();
        }

        // Write modified content back to disk
        File.WriteAllBytes(workingPath, memoryStream.ToArray());
    }

    /// <summary>
    /// Executes operations against an already-open WordprocessingDocument (for testing).
    /// </summary>
    public void Execute(OfficeTalkDocument document, WordprocessingDocument wordDoc)
    {
        var resolver = new WordAddressResolver(wordDoc);
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
                    ExecuteOperation(wordDoc, element, operation);
                }
            }
        }

        foreach (var prop in document.PropertySettings)
        {
            ApplyProperty(wordDoc, prop.Name, prop.Value);
        }
    }

    private static void ExecuteOperation(
        WordprocessingDocument wordDoc, OpenXmlElement element, Operation operation)
    {
        switch (operation)
        {
            case SetOperation set:
                ExecuteSet(element, set);
                break;
            case ReplaceOperation replace:
                ExecuteReplace(element, replace);
                break;
            case DeleteOperation:
                ExecuteDelete(element);
                break;
            case AppendOperation append:
                ExecuteAppend(element, append);
                break;
            case PrependOperation prepend:
                ExecutePrepend(element, prepend);
                break;
            case StyleOperation style:
                ExecuteStyle(element, style);
                break;
            case FormatOperation:
                throw new NotImplementedException("FORMAT operations are not yet supported.");
            case InsertBeforeOperation:
                throw new NotImplementedException("INSERT BEFORE operations are not yet supported.");
            case InsertAfterOperation:
                throw new NotImplementedException("INSERT AFTER operations are not yet supported.");
            case InsertRowOperation:
                throw new NotImplementedException("INSERT ROW operations are not yet supported.");
            case InsertColumnOperation:
                throw new NotImplementedException("INSERT COLUMN operations are not yet supported.");
            case MergeCellsOperation:
                throw new NotImplementedException("MERGE CELLS operations are not yet supported.");
            case SetCellsOperation:
                throw new NotImplementedException("SET CELLS operations are not yet supported.");
            case InsertSlideOperation:
                throw new NotImplementedException("INSERT SLIDE operations are not yet supported.");
            case DuplicateSlideOperation:
                throw new NotImplementedException("DUPLICATE SLIDE operations are not yet supported.");
            case RenameSheetOperation:
                throw new NotImplementedException("RENAME SHEET operations are not yet supported.");
            case AddSheetOperation:
                throw new NotImplementedException("ADD SHEET operations are not yet supported.");
            default:
                throw new NotSupportedException($"Operation type '{operation.GetType().Name}' is not supported.");
        }
    }

    private static void ExecuteSet(OpenXmlElement element, SetOperation operation)
    {
        if (element is Paragraph paragraph)
        {
            // Remove all existing runs
            paragraph.RemoveAllChildren<Run>();

            // Add new run with content
            var run = new Run(new Text(operation.Content.Text) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(run);
        }
        else if (element is TableCell cell)
        {
            var cellParagraph = cell.GetFirstChild<Paragraph>() ?? cell.AppendChild(new Paragraph());
            cellParagraph.RemoveAllChildren<Run>();
            var run = new Run(new Text(operation.Content.Text) { Space = SpaceProcessingModeValues.Preserve });
            cellParagraph.AppendChild(run);
        }
    }

    private static void ExecuteReplace(OpenXmlElement element, ReplaceOperation operation)
    {
        var runs = element.Descendants<Run>().ToList();

        foreach (var run in runs)
        {
            var textElement = run.GetFirstChild<Text>();
            if (textElement == null) continue;

            string text = textElement.Text;
            if (!text.Contains(operation.Search)) continue;

            if (operation.IsAll)
            {
                textElement.Text = text.Replace(operation.Search, operation.Replacement);
            }
            else
            {
                int index = text.IndexOf(operation.Search, StringComparison.Ordinal);
                if (index >= 0)
                {
                    textElement.Text = string.Concat(
                        text.AsSpan(0, index),
                        operation.Replacement,
                        text.AsSpan(index + operation.Search.Length));
                }
            }
        }
    }

    private static void ExecuteDelete(OpenXmlElement element)
    {
        element.Remove();
    }

    private static void ExecuteAppend(OpenXmlElement element, AppendOperation operation)
    {
        if (element is Paragraph paragraph)
        {
            var run = new Run(new Text(operation.Content.Text) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(run);
        }
    }

    private static void ExecutePrepend(OpenXmlElement element, PrependOperation operation)
    {
        if (element is Paragraph paragraph)
        {
            var run = new Run(new Text(operation.Content.Text) { Space = SpaceProcessingModeValues.Preserve });
            var firstRun = paragraph.GetFirstChild<Run>();
            if (firstRun != null)
                paragraph.InsertBefore(run, firstRun);
            else
                paragraph.AppendChild(run);
        }
    }

    private static void ExecuteStyle(OpenXmlElement element, StyleOperation operation)
    {
        if (element is Paragraph paragraph)
        {
            var props = paragraph.ParagraphProperties ?? paragraph.PrependChild(new ParagraphProperties());
            // Normalize display name (e.g. "Heading 2") to style ID (e.g. "Heading2")
            var styleId = operation.StyleName.Replace(" ", "");
            props.ParagraphStyleId = new ParagraphStyleId { Val = styleId };
        }
    }

    private static void ApplyProperty(WordprocessingDocument wordDoc, string name, string value)
    {
        var coreProps = wordDoc.PackageProperties;

        switch (name.ToLowerInvariant())
        {
            case "title":
                coreProps.Title = value;
                break;
            case "author":
                coreProps.Creator = value;
                break;
            case "subject":
                coreProps.Subject = value;
                break;
            case "description":
                coreProps.Description = value;
                break;
            case "keywords":
                coreProps.Keywords = value;
                break;
            case "category":
                coreProps.Category = value;
                break;
            default:
                throw new NotSupportedException($"Document property '{name}' is not supported.");
        }
    }
}
