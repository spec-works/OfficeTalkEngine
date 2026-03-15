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

        // Write modified content back using FileMode.Open to preserve NTFS
        // reparse points (used by OneDrive cloud files for sync tracking).
        // File.WriteAllBytes uses FileMode.Create which destroys them.
        var data = memoryStream.ToArray();
        using (var fs = new FileStream(workingPath, FileMode.Open, FileAccess.Write, FileShare.None))
        {
            fs.Write(data, 0, data.Length);
            fs.SetLength(data.Length);
        }
    }

    /// <summary>
    /// Executes operations against an already-open WordprocessingDocument (for testing).
    /// </summary>
    public void Execute(OfficeTalkDocument document, WordprocessingDocument wordDoc)
    {
        var resolver = new WordAddressResolver(wordDoc);

        bool hasStructuralOps = document.OperationBlocks.Any(b =>
            b.Operations.Any(op => op is InsertBeforeOperation or InsertAfterOperation or DeleteOperation));

        if (hasStructuralOps)
        {
            // Sequential mode: resolve and execute each block in order.
            // Required when operations change document structure so that
            // subsequent addresses reflect the updated document.
            foreach (var block in document.OperationBlocks)
            {
                var elements = resolver.Resolve(block.Address);
                foreach (var element in elements)
                {
                    foreach (var operation in block.Operations)
                    {
                        ExecuteOperation(wordDoc, element, operation);
                    }
                }
            }
        }
        else
        {
            // Snapshot mode: resolve all addresses upfront then execute.
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
            case DeleteOperation delete:
                ExecuteDelete(element, delete);
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
            case FormatOperation format:
                ExecuteFormat(element, format);
                break;
            case InsertBeforeOperation insertBefore:
                ExecuteInsertBefore(element, insertBefore);
                break;
            case InsertAfterOperation insertAfter:
                ExecuteInsertAfter(element, insertAfter);
                break;
            case InsertRowOperation insertRow:
                ExecuteInsertRow(element, insertRow);
                break;
            case InsertColumnOperation insertColumn:
                ExecuteInsertColumn(element, insertColumn);
                break;
            case MergeCellsOperation merge:
                ExecuteMergeCells(element, merge);
                break;
            case SetCellsOperation setCells:
                ExecuteSetCells(element, setCells);
                break;
            case InsertSlideOperation:
                throw new NotImplementedException("INSERT SLIDE operations are not supported for Word documents.");
            case DuplicateSlideOperation:
                throw new NotImplementedException("DUPLICATE SLIDE operations are not supported for Word documents.");
            case RenameSheetOperation:
                throw new NotImplementedException("RENAME SHEET operations are not supported for Word documents.");
            case AddSheetOperation:
                throw new NotImplementedException("ADD SHEET operations are not supported for Word documents.");
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
        else if (element is BookmarkStart bookmarkStart)
        {
            var para = bookmarkStart.Parent as Paragraph;
            if (para != null)
            {
                para.RemoveAllChildren<Run>();
                var run = new Run(new Text(operation.Content.Text) { Space = SpaceProcessingModeValues.Preserve });
                para.AppendChild(run);
            }
        }
        else if (element is SdtBlock sdtBlock)
        {
            var contentBlock = sdtBlock.SdtContentBlock;
            if (contentBlock != null)
            {
                var para = contentBlock.GetFirstChild<Paragraph>();
                if (para != null)
                {
                    para.RemoveAllChildren<Run>();
                    var run = new Run(new Text(operation.Content.Text) { Space = SpaceProcessingModeValues.Preserve });
                    para.AppendChild(run);
                }
            }
        }
        else if (element is SdtRun sdtRun)
        {
            var contentRun = sdtRun.SdtContentRun;
            if (contentRun != null)
            {
                contentRun.RemoveAllChildren<Run>();
                var run = new Run(new Text(operation.Content.Text) { Space = SpaceProcessingModeValues.Preserve });
                contentRun.AppendChild(run);
            }
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

    private static void ExecuteDelete(OpenXmlElement element, DeleteOperation operation)
    {
        switch (operation.Target)
        {
            case DeleteTarget.Element:
                // For inline elements (Drawing, BookmarkStart), remove the containing paragraph
                // if it would be left empty
                if (element is Drawing or BookmarkStart)
                {
                    var containingPara = element.Ancestors<Paragraph>().FirstOrDefault();
                    if (containingPara != null)
                    {
                        containingPara.Remove();
                        break;
                    }
                }
                element.Remove();
                break;
            case DeleteTarget.Row:
                if (element is TableRow row)
                    row.Remove();
                else if (element.Parent is TableRow parentRow)
                    parentRow.Remove();
                else
                    element.Remove();
                break;
            case DeleteTarget.Column:
                // Delete a column by removing the cell at the same position from every row
                if (element is TableCell cell && cell.Parent is TableRow cellRow && cellRow.Parent is Table table)
                {
                    var cellIndex = cellRow.Elements<TableCell>().ToList().IndexOf(cell);
                    if (cellIndex >= 0)
                    {
                        foreach (var r in table.Elements<TableRow>().ToList())
                        {
                            var cells = r.Elements<TableCell>().ToList();
                            if (cellIndex < cells.Count)
                                cells[cellIndex].Remove();
                        }
                    }
                }
                break;
        }
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

    private static void ExecuteFormat(OpenXmlElement element, FormatOperation operation)
    {
        if (element is Paragraph paragraph)
        {
            var pProps = paragraph.ParagraphProperties ?? paragraph.PrependChild(new ParagraphProperties());

            // Build run properties for text formatting
            RunProperties? runProps = null;

            foreach (var (key, value) in operation.Properties)
            {
                var strValue = value?.ToString() ?? "";
                switch (key.ToLowerInvariant())
                {
                    // Text (run) properties — apply to all runs
                    case "bold":
                        runProps ??= new RunProperties();
                        if (strValue.Equals("true", StringComparison.OrdinalIgnoreCase))
                            runProps.Bold = new Bold();
                        else
                            runProps.Bold = new Bold { Val = false };
                        break;
                    case "italic":
                        runProps ??= new RunProperties();
                        if (strValue.Equals("true", StringComparison.OrdinalIgnoreCase))
                            runProps.Italic = new Italic();
                        else
                            runProps.Italic = new Italic { Val = false };
                        break;
                    case "underline":
                        runProps ??= new RunProperties();
                        runProps.Underline = new Underline
                        {
                            Val = strValue.Equals("true", StringComparison.OrdinalIgnoreCase)
                                ? UnderlineValues.Single
                                : UnderlineValues.None
                        };
                        break;
                    case "strikethrough":
                        runProps ??= new RunProperties();
                        runProps.Strike = new Strike
                        {
                            Val = strValue.Equals("true", StringComparison.OrdinalIgnoreCase)
                        };
                        break;
                    case "font-name":
                        runProps ??= new RunProperties();
                        runProps.RunFonts = new RunFonts { Ascii = strValue, HighAnsi = strValue };
                        break;
                    case "font-size":
                        runProps ??= new RunProperties();
                        // OpenXML uses half-points
                        if (TryParsePoints(strValue, out double pts))
                            runProps.FontSize = new FontSize { Val = ((int)(pts * 2)).ToString() };
                        break;
                    case "color":
                        runProps ??= new RunProperties();
                        runProps.Color = new Color { Val = ResolveColorHex(strValue) };
                        break;

                    // Paragraph properties
                    case "alignment":
                    case "align":
                        pProps.Justification = new Justification
                        {
                            Val = strValue.ToLowerInvariant() switch
                            {
                                "left" => JustificationValues.Left,
                                "center" => JustificationValues.Center,
                                "right" => JustificationValues.Right,
                                "justify" => JustificationValues.Both,
                                _ => JustificationValues.Left
                            }
                        };
                        break;
                    case "spacing-before":
                        var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                        if (TryParsePoints(strValue, out double beforePts))
                            spacing.Before = ((int)(beforePts * 20)).ToString(); // twips
                        break;
                    case "spacing-after":
                        var spacingAfter = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                        if (TryParsePoints(strValue, out double afterPts))
                            spacingAfter.After = ((int)(afterPts * 20)).ToString(); // twips
                        break;
                    case "line-spacing":
                        var lineSpacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                        if (TryParsePoints(strValue, out double linePts))
                            lineSpacing.Line = ((int)(linePts * 20)).ToString();
                        break;
                    case "indent-left":
                        var indent = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                        if (TryParsePoints(strValue, out double leftPts))
                            indent.Left = ((int)(leftPts * 20)).ToString();
                        break;
                    case "indent-right":
                        var indentR = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                        if (TryParsePoints(strValue, out double rightPts))
                            indentR.Right = ((int)(rightPts * 20)).ToString();
                        break;
                    case "style":
                        var sId = strValue.Replace(" ", "");
                        pProps.ParagraphStyleId = new ParagraphStyleId { Val = sId };
                        break;
                }
            }

            // Apply run properties to all existing runs
            if (runProps != null)
            {
                foreach (var run in paragraph.Elements<Run>())
                {
                    var existing = run.RunProperties ?? run.PrependChild(new RunProperties());
                    MergeRunProperties(existing, runProps);
                }
            }
        }
        else if (element is Run run)
        {
            // Apply formatting directly to this run
            RunProperties? runProps = null;
            foreach (var (key, value) in operation.Properties)
            {
                var strValue = value?.ToString() ?? "";
                switch (key.ToLowerInvariant())
                {
                    case "bold":
                        runProps ??= new RunProperties();
                        runProps.Bold = strValue.Equals("true", StringComparison.OrdinalIgnoreCase)
                            ? new Bold() : new Bold { Val = false };
                        break;
                    case "italic":
                        runProps ??= new RunProperties();
                        runProps.Italic = strValue.Equals("true", StringComparison.OrdinalIgnoreCase)
                            ? new Italic() : new Italic { Val = false };
                        break;
                    case "underline":
                        runProps ??= new RunProperties();
                        runProps.Underline = new Underline
                        {
                            Val = strValue.Equals("true", StringComparison.OrdinalIgnoreCase)
                                ? UnderlineValues.Single : UnderlineValues.None
                        };
                        break;
                    case "strikethrough":
                        runProps ??= new RunProperties();
                        runProps.Strike = new Strike
                        {
                            Val = strValue.Equals("true", StringComparison.OrdinalIgnoreCase)
                        };
                        break;
                    case "font-name":
                        runProps ??= new RunProperties();
                        runProps.RunFonts = new RunFonts { Ascii = strValue, HighAnsi = strValue };
                        break;
                    case "font-size":
                        runProps ??= new RunProperties();
                        if (TryParsePoints(strValue, out double runPts))
                            runProps.FontSize = new FontSize { Val = ((int)(runPts * 2)).ToString() };
                        break;
                    case "color":
                        runProps ??= new RunProperties();
                        runProps.Color = new Color { Val = ResolveColorHex(strValue) };
                        break;
                }
            }
            if (runProps != null)
            {
                var existing = run.RunProperties ?? run.PrependChild(new RunProperties());
                MergeRunProperties(existing, runProps);
            }
        }
        else if (element is TableCell tableCell)
        {
            foreach (var (key, value) in operation.Properties)
            {
                var strValue = value?.ToString() ?? "";
                switch (key.ToLowerInvariant())
                {
                    case "fill-color":
                    case "background-color":
                        var tcPr = tableCell.TableCellProperties ??
                                   tableCell.PrependChild(new TableCellProperties());
                        tcPr.Shading = new Shading
                        {
                            Fill = strValue.TrimStart('#'),
                            Val = ShadingPatternValues.Clear
                        };
                        break;
                }
            }
        }
    }

    private static void MergeRunProperties(RunProperties target, RunProperties source)
    {
        if (source.Bold != null) target.Bold = (Bold)source.Bold.CloneNode(true);
        if (source.Italic != null) target.Italic = (Italic)source.Italic.CloneNode(true);
        if (source.Underline != null) target.Underline = (Underline)source.Underline.CloneNode(true);
        if (source.Strike != null) target.Strike = (Strike)source.Strike.CloneNode(true);
        if (source.RunFonts != null) target.RunFonts = (RunFonts)source.RunFonts.CloneNode(true);
        if (source.FontSize != null) target.FontSize = (FontSize)source.FontSize.CloneNode(true);
        if (source.Color != null) target.Color = (Color)source.Color.CloneNode(true);
    }

    private static bool TryParsePoints(string value, out double points)
    {
        // Strip unit suffix (pt, in, cm, mm)
        var trimmed = value.TrimEnd();
        if (trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed[..^2];
        else if (trimmed.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            if (double.TryParse(trimmed[..^2], out double inches))
            {
                points = inches * 72;
                return true;
            }
        }
        else if (trimmed.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            if (double.TryParse(trimmed[..^2], out double cm))
            {
                points = cm * 28.3465;
                return true;
            }
        }

        return double.TryParse(trimmed, out points);
    }

    /// <summary>
    /// Resolves a color value (named or hex) to a 6-digit hex string for OpenXML.
    /// </summary>
    private static string ResolveColorHex(string color)
    {
        var hex = color.Trim().ToLowerInvariant() switch
        {
            "black" => "000000",
            "red" => "FF0000",
            "green" => "008000",
            "blue" => "0000FF",
            "white" => "FFFFFF",
            "yellow" => "FFFF00",
            "orange" => "FFA500",
            "purple" => "800080",
            "gray" or "grey" => "808080",
            _ => color.TrimStart('#')
        };
        return hex;
    }

    private static void ExecuteInsertBefore(OpenXmlElement element, InsertBeforeOperation operation)
    {
        var newParagraph = CreateParagraphFromContent(operation.Content);
        element.InsertBeforeSelf(newParagraph);
    }

    private static void ExecuteInsertAfter(OpenXmlElement element, InsertAfterOperation operation)
    {
        var newParagraph = CreateParagraphFromContent(operation.Content);
        element.InsertAfterSelf(newParagraph);
    }

    private static Paragraph CreateParagraphFromContent(ContentValue content)
    {
        if (content.IsContentBlock)
        {
            // Content blocks may contain multiple lines — create paragraph with the full text
            // Future: could split into multiple paragraphs on blank lines
            var paragraph = new Paragraph();
            var lines = content.Text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i].TrimEnd('\r');
                if (i > 0)
                    paragraph.AppendChild(new Run(new Break()));
                paragraph.AppendChild(new Run(
                    new Text(line) { Space = SpaceProcessingModeValues.Preserve }));
            }
            return paragraph;
        }
        else
        {
            return new Paragraph(
                new Run(new Text(content.Text) { Space = SpaceProcessingModeValues.Preserve }));
        }
    }

    private static void ExecuteInsertRow(OpenXmlElement element, InsertRowOperation operation)
    {
        if (element is not TableRow targetRow || targetRow.Parent is not Table)
            return;

        // Clone the row structure (preserving cell count and properties) but clear content
        var newRow = (TableRow)targetRow.CloneNode(true);
        foreach (var cell in newRow.Elements<TableCell>())
        {
            foreach (var p in cell.Elements<Paragraph>())
                p.RemoveAllChildren<Run>();
        }

        if (operation.Position == InsertPosition.Before)
            targetRow.InsertBeforeSelf(newRow);
        else
            targetRow.InsertAfterSelf(newRow);
    }

    private static void ExecuteInsertColumn(OpenXmlElement element, InsertColumnOperation operation)
    {
        // Element should be a cell; insert a new cell at the same column index in every row
        if (element is not TableCell targetCell || targetCell.Parent is not TableRow targetRow ||
            targetRow.Parent is not Table table)
            return;

        var cellIndex = targetRow.Elements<TableCell>().ToList().IndexOf(targetCell);

        foreach (var row in table.Elements<TableRow>())
        {
            var cells = row.Elements<TableCell>().ToList();
            var referenceCell = cellIndex < cells.Count ? cells[cellIndex] : cells.LastOrDefault();
            if (referenceCell == null) continue;

            // Create new cell with same properties but empty content
            var newCell = new TableCell(new Paragraph());
            if (referenceCell.TableCellProperties != null)
                newCell.TableCellProperties = (TableCellProperties)referenceCell.TableCellProperties.CloneNode(true);

            if (operation.Position == InsertPosition.Before)
                referenceCell.InsertBeforeSelf(newCell);
            else
                referenceCell.InsertAfterSelf(newCell);
        }
    }

    private static void ExecuteMergeCells(OpenXmlElement element, MergeCellsOperation operation)
    {
        // Horizontal merge: merge cells from the addressed cell to the target cell
        if (element is not TableCell startCell || startCell.Parent is not TableRow row)
            return;

        var cells = row.Elements<TableCell>().ToList();
        var startIndex = cells.IndexOf(startCell);

        // Parse target address to find end cell index
        // The target address is relative (e.g., "row[2]/cell[3]")
        // For simplicity, check if it has a cell segment with positional predicate
        var cellSegment = operation.TargetAddress.Segments
            .FirstOrDefault(s => s.Identifier.Equals("cell", StringComparison.OrdinalIgnoreCase));

        if (cellSegment == null) return;

        var posPred = cellSegment.Predicates.OfType<PositionalPredicate>().FirstOrDefault();
        if (posPred == null) return;

        int endIndex = posPred.Position - 1;
        if (endIndex <= startIndex || endIndex >= cells.Count) return;

        // Apply horizontal merge
        var startProps = startCell.TableCellProperties ?? startCell.PrependChild(new TableCellProperties());
        startProps.HorizontalMerge = new HorizontalMerge { Val = MergedCellValues.Restart };

        for (int i = startIndex + 1; i <= endIndex; i++)
        {
            var cellProps = cells[i].TableCellProperties ?? cells[i].PrependChild(new TableCellProperties());
            cellProps.HorizontalMerge = new HorizontalMerge { Val = MergedCellValues.Continue };
        }
    }

    private static void ExecuteSetCells(OpenXmlElement element, SetCellsOperation operation)
    {
        // SET CELLS populates cells in a row by position
        if (element is not TableRow row) return;

        var cells = row.Elements<TableCell>().ToList();
        for (int i = 0; i < operation.Values.Count && i < cells.Count; i++)
        {
            var cellParagraph = cells[i].GetFirstChild<Paragraph>() ?? cells[i].AppendChild(new Paragraph());
            cellParagraph.RemoveAllChildren<Run>();
            cellParagraph.AppendChild(new Run(
                new Text(operation.Values[i]) { Space = SpaceProcessingModeValues.Preserve }));
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
