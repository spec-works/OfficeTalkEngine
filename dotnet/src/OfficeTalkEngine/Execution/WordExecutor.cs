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
            case CommentOperation comment:
                ExecuteComment(wordDoc, element, comment);
                break;
            case InsertImageOperation insertImage:
                ExecuteInsertImage(wordDoc, element, insertImage);
                break;
            case InsertTableOperation insertTable:
                ExecuteInsertTable(element, insertTable);
                break;
            case LinkOperation link:
                ExecuteLink(wordDoc, element, link);
                break;
            case InsertListOperation insertList:
                ExecuteInsertList(wordDoc, element, insertList);
                break;
            case SetRunsOperation setRuns:
                ExecuteSetRuns(wordDoc, element, setRuns);
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

                    // Run-level background/shading (applied to runs below)
                    case "background-color":
                        runProps ??= new RunProperties();
                        runProps.Shading = new Shading
                        {
                            Fill = ResolveColorHex(strValue),
                            Val = ShadingPatternValues.Clear
                        };
                        break;
                    case "highlight":
                        runProps ??= new RunProperties();
                        runProps.Highlight = new Highlight { Val = ResolveHighlightColor(strValue) };
                        break;

                    // Paragraph border properties
                    case "border-top":
                    case "border-bottom":
                    case "border-left":
                    case "border-right":
                        ApplyParagraphBorder(pProps, key.ToLowerInvariant(), strValue, operation.Properties);
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
                    case "background-color":
                        runProps ??= new RunProperties();
                        runProps.Shading = new Shading
                        {
                            Fill = ResolveColorHex(strValue),
                            Val = ShadingPatternValues.Clear
                        };
                        break;
                    case "highlight":
                        runProps ??= new RunProperties();
                        runProps.Highlight = new Highlight { Val = ResolveHighlightColor(strValue) };
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
        if (source.Shading != null) target.Shading = (Shading)source.Shading.CloneNode(true);
        if (source.Highlight != null) target.Highlight = (Highlight)source.Highlight.CloneNode(true);
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

    private static void ExecuteComment(
        WordprocessingDocument wordDoc, OpenXmlElement element, CommentOperation operation)
    {
        var mainPart = wordDoc.MainDocumentPart
            ?? throw new InvalidOperationException("Document has no main part.");

        // Ensure comments part exists
        var commentsPart = mainPart.WordprocessingCommentsPart;
        if (commentsPart == null)
        {
            commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments();
        }

        var comments = commentsPart.Comments;

        // Generate unique comment ID (max existing + 1, or 0)
        int commentId = 0;
        var existingIds = comments.Elements<Comment>()
            .Select(c => c.Id?.Value)
            .Where(id => id != null)
            .Select(id => int.TryParse(id, out var n) ? n : -1)
            .Where(n => n >= 0);
        if (existingIds.Any())
            commentId = existingIds.Max() + 1;

        var commentIdStr = commentId.ToString();

        // Create the comment element
        var comment = new Comment
        {
            Id = commentIdStr,
            Author = "OfficeTalk",
            Date = DateTime.UtcNow
        };

        // Add comment text — split on newlines into paragraphs
        var lines = operation.Content.Text.Split('\n');
        foreach (var line in lines)
        {
            comment.AppendChild(new Paragraph(
                new Run(new Text(line.TrimEnd('\r')) { Space = SpaceProcessingModeValues.Preserve })));
        }
        comments.AppendChild(comment);

        // Wrap the addressed element with CommentRangeStart/End and reference
        var rangeStart = new CommentRangeStart { Id = commentIdStr };
        var rangeEnd = new CommentRangeEnd { Id = commentIdStr };

        if (element is Run run)
        {
            // For runs, wrap the run itself
            run.InsertBeforeSelf(rangeStart);
            run.InsertAfterSelf(rangeEnd);
            rangeEnd.InsertAfterSelf(new Run(
                new CommentReference { Id = commentIdStr }));
        }
        else if (element is Paragraph para)
        {
            // For paragraphs, put start before first run, end before paragraph mark
            var firstRun = para.GetFirstChild<Run>();
            if (firstRun != null)
                firstRun.InsertBeforeSelf(rangeStart);
            else
                para.PrependChild(rangeStart);

            // Insert end and reference before ParagraphProperties or at end
            var paraProps = para.GetFirstChild<ParagraphProperties>();
            if (paraProps != null)
            {
                // Range markers go after properties but at end of content
                para.AppendChild(rangeEnd);
                para.AppendChild(new Run(new CommentReference { Id = commentIdStr }));
            }
            else
            {
                para.AppendChild(rangeEnd);
                para.AppendChild(new Run(new CommentReference { Id = commentIdStr }));
            }
        }
        else
        {
            // For other elements (table, row, etc.), insert before/after the element
            element.InsertBeforeSelf(rangeStart);
            element.InsertAfterSelf(rangeEnd);
            rangeEnd.InsertAfterSelf(new Paragraph(
                new Run(new CommentReference { Id = commentIdStr })));
        }
    }

    private static void ExecuteInsertImage(
        WordprocessingDocument wordDoc, OpenXmlElement element, InsertImageOperation operation)
    {
        var mainPart = wordDoc.MainDocumentPart
            ?? throw new InvalidOperationException("Document has no main part.");

        // Read image bytes from file or URL
        byte[] imageBytes;
        string contentType;
        if (operation.Source.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            using var client = new System.Net.Http.HttpClient();
            imageBytes = client.GetByteArrayAsync(operation.Source).GetAwaiter().GetResult();
            contentType = "image/png"; // default
        }
        else
        {
            imageBytes = File.ReadAllBytes(operation.Source);
            contentType = GetImageContentType(operation.Source);
        }

        // Add image part
        var imagePart = mainPart.AddImagePart(
            contentType switch
            {
                "image/png" => ImagePartType.Png,
                "image/jpeg" => ImagePartType.Jpeg,
                "image/gif" => ImagePartType.Gif,
                "image/bmp" => ImagePartType.Bmp,
                "image/svg+xml" => ImagePartType.Svg,
                _ => ImagePartType.Png
            });

        using (var stream = new MemoryStream(imageBytes))
        {
            imagePart.FeedData(stream);
        }

        string relId = mainPart.GetIdOfPart(imagePart);

        // Parse dimensions (default 4in x 3in = 3657600 x 2743200 EMU)
        long widthEmu = 3657600L;
        long heightEmu = 2743200L;

        if (operation.Properties.TryGetValue("width", out var wVal) &&
            TryParseEmu(wVal?.ToString() ?? "", out var w))
            widthEmu = w;
        if (operation.Properties.TryGetValue("height", out var hVal) &&
            TryParseEmu(hVal?.ToString() ?? "", out var h))
            heightEmu = h;

        string altText = operation.Properties.TryGetValue("alt", out var alt)
            ? alt?.ToString() ?? "" : "";

        // Build the Drawing element
        var drawing = CreateInlineDrawing(relId, widthEmu, heightEmu, altText);
        var paragraph = new Paragraph(new Run(drawing));

        if (operation.Position == InsertPosition.Before)
            element.InsertBeforeSelf(paragraph);
        else
            element.InsertAfterSelf(paragraph);
    }

    private static DocumentFormat.OpenXml.Wordprocessing.Drawing CreateInlineDrawing(
        string relId, long widthEmu, long heightEmu, string altText)
    {
        var element = new DocumentFormat.OpenXml.Wordprocessing.Drawing(
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent
                {
                    Cx = widthEmu,
                    Cy = heightEmu
                },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties
                {
                    Id = 1U,
                    Name = "Image",
                    Description = altText
                },
                new DocumentFormat.OpenXml.Drawing.Graphic(
                    new DocumentFormat.OpenXml.Drawing.GraphicData(
                        new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
                                {
                                    Id = 0U,
                                    Name = "Image"
                                },
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                            new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                new DocumentFormat.OpenXml.Drawing.Blip
                                {
                                    Embed = relId
                                },
                                new DocumentFormat.OpenXml.Drawing.Stretch(
                                    new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                            new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                new DocumentFormat.OpenXml.Drawing.Transform2D(
                                    new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 },
                                    new DocumentFormat.OpenXml.Drawing.Extents
                                    {
                                        Cx = widthEmu,
                                        Cy = heightEmu
                                    }),
                                new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                    new DocumentFormat.OpenXml.Drawing.AdjustValueList())
                                {
                                    Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
                                }))
                    )
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                    }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            });

        return element;
    }

    private static string GetImageContentType(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext switch
        {
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".svg" => "image/svg+xml",
            _ => "image/png"
        };
    }

    private static bool TryParseEmu(string value, out long emu)
    {
        emu = 0;
        if (string.IsNullOrEmpty(value)) return false;

        if (value.EndsWith("emu"))
        {
            return long.TryParse(value[..^3], out emu);
        }
        if (value.EndsWith("in"))
        {
            if (double.TryParse(value[..^2], out var inches))
            {
                emu = (long)(inches * 914400);
                return true;
            }
        }
        if (value.EndsWith("cm"))
        {
            if (double.TryParse(value[..^2], out var cm))
            {
                emu = (long)(cm * 360000);
                return true;
            }
        }
        if (value.EndsWith("pt"))
        {
            if (double.TryParse(value[..^2], out var pt))
            {
                emu = (long)(pt * 12700);
                return true;
            }
        }
        if (value.EndsWith("px"))
        {
            if (double.TryParse(value[..^2], out var px))
            {
                emu = (long)(px * 9525);
                return true;
            }
        }
        return false;
    }

    private static void ExecuteInsertTable(OpenXmlElement element, InsertTableOperation operation)
    {
        var table = new Table();

        // Add table properties
        var tblProps = new TableProperties(
            new TableBorders(
                new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
            ));
        table.AppendChild(tblProps);

        // Create rows and cells
        for (int r = 0; r < operation.Rows; r++)
        {
            var row = new TableRow();
            for (int c = 0; c < operation.Columns; c++)
            {
                row.AppendChild(new TableCell(new Paragraph()));
            }
            table.AppendChild(row);
        }

        if (operation.Position == InsertPosition.Before)
            element.InsertBeforeSelf(table);
        else
            element.InsertAfterSelf(table);
    }

    private static void ExecuteLink(
        WordprocessingDocument wordDoc, OpenXmlElement element, LinkOperation operation)
    {
        var mainPart = wordDoc.MainDocumentPart
            ?? throw new InvalidOperationException("Document has no main part.");

        // Create hyperlink relationship
        var relId = mainPart.AddHyperlinkRelationship(new Uri(operation.Url), true).Id;

        if (element is Run run)
        {
            // Wrap the run in a hyperlink
            var parent = run.Parent;
            if (parent == null) return;

            var hyperlink = new Hyperlink { Id = relId };

            // Apply blue underline styling typical of links
            var linkRun = (Run)run.CloneNode(true);
            var rProps = linkRun.RunProperties ?? linkRun.PrependChild(new RunProperties());
            rProps.Color = new Color { Val = "0563C1" };
            rProps.Underline = new Underline { Val = UnderlineValues.Single };

            hyperlink.AppendChild(linkRun);
            run.InsertAfterSelf(hyperlink);
            run.Remove();
        }
        else if (element is Paragraph para)
        {
            // Wrap all runs in the paragraph into a hyperlink
            var hyperlink = new Hyperlink { Id = relId };
            var runs = para.Elements<Run>().ToList();

            foreach (var r in runs)
            {
                var linkRun = (Run)r.CloneNode(true);
                var rProps = linkRun.RunProperties ?? linkRun.PrependChild(new RunProperties());
                rProps.Color = new Color { Val = "0563C1" };
                rProps.Underline = new Underline { Val = UnderlineValues.Single };
                hyperlink.AppendChild(linkRun);
                r.Remove();
            }

            para.AppendChild(hyperlink);
        }
    }

    private static void ExecuteInsertList(
        WordprocessingDocument wordDoc, OpenXmlElement element, InsertListOperation operation)
    {
        var mainPart = wordDoc.MainDocumentPart
            ?? throw new InvalidOperationException("Document has no main part.");

        // Ensure numbering part exists
        var numberingPart = mainPart.NumberingDefinitionsPart;
        if (numberingPart == null)
        {
            numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering();
        }

        var numbering = numberingPart.Numbering;

        // Create abstract numbering definition
        int abstractNumId = numbering.Elements<AbstractNum>().Count() + 1;
        int numId = numbering.Elements<NumberingInstance>().Count() + 1;

        var abstractNum = new AbstractNum { AbstractNumberId = abstractNumId };

        if (operation.ListType == ListType.Ordered)
        {
            abstractNum.AppendChild(new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "%1." })
            { LevelIndex = 0 });
            abstractNum.AppendChild(new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.LowerLetter },
                new LevelText { Val = "%2." })
            { LevelIndex = 1 });
        }
        else
        {
            abstractNum.AppendChild(new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.Bullet },
                new LevelText { Val = "\u2022" })
            { LevelIndex = 0 });
            abstractNum.AppendChild(new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.Bullet },
                new LevelText { Val = "\u25E6" })
            { LevelIndex = 1 });
        }

        numbering.InsertAt(abstractNum, 0);
        numbering.AppendChild(new NumberingInstance(
            new AbstractNumId { Val = abstractNumId })
        { NumberID = numId });
        numbering.Save();

        // Create list item paragraphs and insert them
        var insertionPoint = element;
        foreach (var item in operation.Items)
        {
            int level = item.IsNested ? 1 : 0;
            var listParagraph = new Paragraph(
                new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = level },
                        new NumberingId { Val = numId })),
                new Run(new Text(item.Content.Text) { Space = SpaceProcessingModeValues.Preserve }));

            if (operation.Position == InsertPosition.Before && insertionPoint == element)
            {
                element.InsertBeforeSelf(listParagraph);
            }
            else
            {
                insertionPoint.InsertAfterSelf(listParagraph);
            }
            insertionPoint = listParagraph;
        }
    }

    private static void ExecuteSetRuns(
        WordprocessingDocument wordDoc, OpenXmlElement element, SetRunsOperation operation)
    {
        var mainPart = wordDoc.MainDocumentPart;

        Paragraph? targetParagraph = null;
        if (element is Paragraph para)
            targetParagraph = para;
        else if (element is TableCell cell)
            targetParagraph = cell.GetFirstChild<Paragraph>() ?? cell.AppendChild(new Paragraph());

        if (targetParagraph == null) return;

        // Remove all existing runs (preserve ParagraphProperties)
        targetParagraph.RemoveAllChildren<Run>();
        targetParagraph.RemoveAllChildren<Hyperlink>();

        foreach (var runDef in operation.Runs)
        {
            var text = new Text(runDef.Content.Text) { Space = SpaceProcessingModeValues.Preserve };
            var run = new Run(text);

            // Apply run properties
            if (runDef.Properties.Count > 0)
            {
                var rProps = new RunProperties();
                string? href = null;

                foreach (var (key, value) in runDef.Properties)
                {
                    var strValue = value?.ToString() ?? "";
                    switch (key.ToLowerInvariant())
                    {
                        case "bold":
                            if (strValue.Equals("true", StringComparison.OrdinalIgnoreCase))
                                rProps.Bold = new Bold();
                            break;
                        case "italic":
                            if (strValue.Equals("true", StringComparison.OrdinalIgnoreCase))
                                rProps.Italic = new Italic();
                            break;
                        case "underline":
                            rProps.Underline = new Underline
                            {
                                Val = strValue.ToLowerInvariant() switch
                                {
                                    "single" => UnderlineValues.Single,
                                    "double" => UnderlineValues.Double,
                                    "dotted" => UnderlineValues.Dotted,
                                    "dashed" => UnderlineValues.Dash,
                                    "wavy" => UnderlineValues.Wave,
                                    "none" => UnderlineValues.None,
                                    "true" => UnderlineValues.Single,
                                    _ => UnderlineValues.Single
                                }
                            };
                            break;
                        case "strikethrough":
                            rProps.Strike = new Strike
                            {
                                Val = strValue.Equals("true", StringComparison.OrdinalIgnoreCase)
                            };
                            break;
                        case "font-name":
                            rProps.RunFonts = new RunFonts { Ascii = strValue, HighAnsi = strValue };
                            break;
                        case "font-size":
                            if (TryParsePoints(strValue, out double pts))
                                rProps.FontSize = new FontSize { Val = ((int)(pts * 2)).ToString() };
                            break;
                        case "color":
                            rProps.Color = new Color { Val = ResolveColorHex(strValue) };
                            break;
                        case "highlight":
                            rProps.Highlight = new Highlight
                            {
                                Val = ResolveHighlightColor(strValue)
                            };
                            break;
                        case "background-color":
                            rProps.Shading = new Shading
                            {
                                Fill = ResolveColorHex(strValue),
                                Val = ShadingPatternValues.Clear
                            };
                            break;
                        case "href":
                            href = strValue;
                            // Style as hyperlink
                            rProps.Color ??= new Color { Val = "0563C1" };
                            rProps.Underline ??= new Underline { Val = UnderlineValues.Single };
                            break;
                    }
                }

                run.PrependChild(rProps);

                // If this run has an href, wrap it in a hyperlink
                if (href != null && mainPart != null)
                {
                    var relId = mainPart.AddHyperlinkRelationship(new Uri(href), true).Id;
                    var hyperlink = new Hyperlink(run) { Id = relId };
                    targetParagraph.AppendChild(hyperlink);
                    continue; // skip normal append
                }
            }

            targetParagraph.AppendChild(run);
        }
    }

    private static void ApplyParagraphBorder(
        ParagraphProperties pProps, string side, string style,
        Dictionary<string, object> allProperties)
    {
        var borders = pProps.ParagraphBorders ?? (pProps.ParagraphBorders = new ParagraphBorders());

        var borderVal = style.ToLowerInvariant() switch
        {
            "single" => BorderValues.Single,
            "double" => BorderValues.Double,
            "thick" => BorderValues.Thick,
            "dashed" => BorderValues.Dashed,
            "dotted" => BorderValues.Dotted,
            "none" => BorderValues.None,
            _ => BorderValues.Single
        };

        // Resolve shared border-color and border-width
        string? colorHex = null;
        uint sizeEighths = 4; // default 0.5pt (4 eighths of a point)
        if (allProperties.TryGetValue("border-color", out var colorVal))
            colorHex = ResolveColorHex(colorVal?.ToString() ?? "");
        if (allProperties.TryGetValue("border-width", out var widthVal) &&
            TryParsePoints(widthVal?.ToString() ?? "", out double widthPts))
            sizeEighths = (uint)(widthPts * 8);

        switch (side)
        {
            case "border-top":
                borders.TopBorder = new TopBorder { Val = borderVal, Size = sizeEighths, Space = 1 };
                if (colorHex != null) borders.TopBorder.Color = colorHex;
                break;
            case "border-bottom":
                borders.BottomBorder = new BottomBorder { Val = borderVal, Size = sizeEighths, Space = 1 };
                if (colorHex != null) borders.BottomBorder.Color = colorHex;
                break;
            case "border-left":
                borders.LeftBorder = new LeftBorder { Val = borderVal, Size = sizeEighths, Space = 1 };
                if (colorHex != null) borders.LeftBorder.Color = colorHex;
                break;
            case "border-right":
                borders.RightBorder = new RightBorder { Val = borderVal, Size = sizeEighths, Space = 1 };
                if (colorHex != null) borders.RightBorder.Color = colorHex;
                break;
        }
    }

    private static HighlightColorValues ResolveHighlightColor(string color)
    {
        return color.ToLowerInvariant() switch
        {
            "yellow" => HighlightColorValues.Yellow,
            "green" => HighlightColorValues.Green,
            "cyan" => HighlightColorValues.Cyan,
            "magenta" => HighlightColorValues.Magenta,
            "blue" => HighlightColorValues.Blue,
            "red" => HighlightColorValues.Red,
            "darkblue" => HighlightColorValues.DarkBlue,
            "darkcyan" => HighlightColorValues.DarkCyan,
            "darkgreen" => HighlightColorValues.DarkGreen,
            "darkmagenta" => HighlightColorValues.DarkMagenta,
            "darkred" => HighlightColorValues.DarkRed,
            "darkyellow" => HighlightColorValues.DarkYellow,
            "lightgray" or "lightgrey" => HighlightColorValues.LightGray,
            "darkgray" or "darkgrey" => HighlightColorValues.DarkGray,
            "black" => HighlightColorValues.Black,
            "white" => HighlightColorValues.White,
            _ => HighlightColorValues.Yellow
        };
    }
}
