using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeTalk.Ast;
using OfficeTalkEngine.Addressing;
using OfficeTalkEngine.Responses;

namespace OfficeTalkEngine.Inspection;

/// <summary>
/// Executes INSPECT operations against Word documents, producing
/// JSONL responses conforming to §14.3 of the OfficeTalk specification.
/// </summary>
public class WordInspector
{
    private readonly WordprocessingDocument _wordDoc;
    private readonly WordAddressResolver _resolver;

    public WordInspector(WordprocessingDocument wordDoc)
    {
        _wordDoc = wordDoc ?? throw new ArgumentNullException(nameof(wordDoc));
        _resolver = new WordAddressResolver(wordDoc);
    }

    /// <summary>
    /// Execute all INSPECT blocks in the document and return responses.
    /// </summary>
    public IReadOnlyList<InspectResponse> Inspect(OfficeTalkDocument document)
    {
        var responses = new List<InspectResponse>();

        foreach (var block in document.InspectBlocks)
        {
            responses.Add(InspectBlock(block));
        }

        return responses;
    }

    /// <summary>
    /// Execute a single INSPECT block.
    /// </summary>
    public InspectResponse InspectBlock(InspectBlock block)
    {
        var addressStr = block.Address.ToString();

        try
        {
            var elements = _resolver.Resolve(block.Address);
            var includeContent = block.Include.Contains(IncludeLayer.Content);
            var includeProperties = block.Include.Contains(IncludeLayer.Properties);

            var elementInfos = new List<ElementInfo>();
            var bodyChildren = GetBodyChildren();

            foreach (var element in elements)
            {
                var info = BuildElementInfo(
                    element, bodyChildren,
                    includeContent, includeProperties,
                    block.Depth, block.Context);
                elementInfos.Add(info);
            }

            return new InspectResponse
            {
                Address = addressStr,
                Matched = elements.Count,
                Elements = elementInfos
            };
        }
        catch (Exception ex)
        {
            return new InspectResponse
            {
                Address = addressStr,
                Matched = 0,
                Error = ex.Message
            };
        }
    }

    private ElementInfo BuildElementInfo(
        OpenXmlElement element,
        IReadOnlyList<OpenXmlElement> bodyChildren,
        bool includeContent,
        bool includeProperties,
        int depth,
        int contextCount)
    {
        var info = BuildAddressingLayer(element, bodyChildren);

        if (includeContent)
            info.Content = BuildContentLayer(element);

        if (includeProperties)
            info.Properties = BuildPropertiesLayer(element);

        var comments = BuildCommentsLayer(element);
        if (comments.Count > 0)
            info.Comments = comments;

        if (depth > 0)
            info.Children = BuildChildren(element, includeContent, includeProperties, depth - 1);

        if (contextCount > 0)
            info.Context = BuildContextLayer(element, bodyChildren, contextCount, includeContent);

        return info;
    }

    #region Addressing Layer

    private ElementInfo BuildAddressingLayer(OpenXmlElement element, IReadOnlyList<OpenXmlElement> bodyChildren)
    {
        return element switch
        {
            Paragraph para => BuildParagraphAddressing(para, bodyChildren),
            Table table => BuildTableAddressing(table, bodyChildren),
            TableRow row => BuildRowAddressing(row),
            TableCell cell => BuildCellAddressing(cell),
            BookmarkStart bm => BuildBookmarkAddressing(bm, bodyChildren),
            SdtBlock sdt => BuildContentControlAddressing(sdt, bodyChildren),
            SdtRun sdtRun => BuildContentControlRunAddressing(sdtRun),
            _ => new ElementInfo { Type = MapElementType(element) }
        };
    }

    private ElementInfo BuildParagraphAddressing(Paragraph para, IReadOnlyList<OpenXmlElement> bodyChildren)
    {
        var headingLevel = GetHeadingLevel(para);
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var (index, total) = GetPositionInParent(para, bodyChildren);

        if (headingLevel > 0)
        {
            return new ElementInfo
            {
                Type = "heading",
                Level = headingLevel,
                Style = styleId,
                Index = index,
                Of = total
            };
        }

        var info = new ElementInfo
        {
            Type = "paragraph",
            Index = index,
            Of = total
        };

        if (!string.IsNullOrEmpty(styleId))
            info.Style = styleId;

        return info;
    }

    private ElementInfo BuildTableAddressing(Table table, IReadOnlyList<OpenXmlElement> bodyChildren)
    {
        var (index, total) = GetPositionAmong<Table>(table, bodyChildren);
        return new ElementInfo
        {
            Type = "table",
            Index = index,
            Of = total
        };
    }

    private ElementInfo BuildRowAddressing(TableRow row)
    {
        var table = row.Parent as Table;
        if (table == null) return new ElementInfo { Type = "row" };

        var rows = table.Elements<TableRow>().ToList();
        var idx = rows.IndexOf(row);
        return new ElementInfo
        {
            Type = "row",
            Index = idx >= 0 ? idx + 1 : null,
            Of = rows.Count
        };
    }

    private ElementInfo BuildCellAddressing(TableCell cell)
    {
        var row = cell.Parent as TableRow;
        if (row == null) return new ElementInfo { Type = "cell" };

        var cells = row.Elements<TableCell>().ToList();
        var idx = cells.IndexOf(cell);
        return new ElementInfo
        {
            Type = "cell",
            Index = idx >= 0 ? idx + 1 : null,
            Of = cells.Count
        };
    }

    private ElementInfo BuildBookmarkAddressing(BookmarkStart bm, IReadOnlyList<OpenXmlElement> bodyChildren)
    {
        return new ElementInfo
        {
            Type = "bookmark",
            Name = bm.Name?.Value
        };
    }

    private ElementInfo BuildContentControlAddressing(SdtBlock sdt, IReadOnlyList<OpenXmlElement> bodyChildren)
    {
        var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
        var (index, total) = GetPositionAmong<SdtBlock>(sdt, bodyChildren);
        var info = new ElementInfo
        {
            Type = "content-control",
            Index = index,
            Of = total
        };
        if (!string.IsNullOrEmpty(tag))
            info.Tag = tag;
        return info;
    }

    private ElementInfo BuildContentControlRunAddressing(SdtRun sdtRun)
    {
        var tag = sdtRun.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
        var info = new ElementInfo { Type = "content-control" };
        if (!string.IsNullOrEmpty(tag))
            info.Tag = tag;
        return info;
    }

    #endregion

    #region Content Layer

    private ContentInfo BuildContentLayer(OpenXmlElement element)
    {
        return element switch
        {
            Paragraph para => new ContentInfo { Text = para.InnerText },
            TableRow row => new ContentInfo
            {
                Cells = row.Elements<TableCell>().Select(c => c.InnerText).ToList()
            },
            TableCell cell => new ContentInfo { Text = cell.InnerText },
            Table table => new ContentInfo
            {
                Text = $"[table with {table.Elements<TableRow>().Count()} rows]"
            },
            BookmarkStart bm => BuildBookmarkContentLayer(bm),
            SdtBlock sdt => new ContentInfo { Text = sdt.InnerText },
            SdtRun sdtRun => new ContentInfo { Text = sdtRun.InnerText },
            _ => new ContentInfo { Text = element.InnerText }
        };
    }

    private ContentInfo BuildBookmarkContentLayer(BookmarkStart bm)
    {
        // Collect text between BookmarkStart and matching BookmarkEnd
        var text = new System.Text.StringBuilder();
        var next = bm.NextSibling();
        while (next != null)
        {
            if (next is BookmarkEnd end && end.Id?.Value == bm.Id?.Value)
                break;
            text.Append(next.InnerText);
            next = next.NextSibling();
        }
        return new ContentInfo { Text = text.ToString() };
    }

    #endregion

    #region Properties Layer

    private Dictionary<string, object?> BuildPropertiesLayer(OpenXmlElement element)
    {
        var props = new Dictionary<string, object?>();

        if (element is Paragraph para)
        {
            BuildParagraphProperties(para, props);
        }
        else if (element is TableCell cell)
        {
            BuildCellProperties(cell, props);
        }
        else if (element is Table table)
        {
            BuildTableProperties(table, props);
        }

        // Get run-level properties from the first run
        var firstRun = element.Descendants<Run>().FirstOrDefault();
        if (firstRun?.RunProperties != null)
        {
            BuildRunProperties(firstRun.RunProperties, props);
        }

        return props.Count > 0 ? props : new Dictionary<string, object?>();
    }

    private void BuildParagraphProperties(Paragraph para, Dictionary<string, object?> props)
    {
        var pPr = para.ParagraphProperties;
        if (pPr == null) return;

        if (pPr.ParagraphStyleId?.Val?.Value is string style)
            props["style"] = style;

        if (pPr.Justification?.Val?.Value is JustificationValues alignment)
        {
            props["alignment"] = alignment.ToString().ToLowerInvariant() switch
            {
                "left" => "left",
                "center" => "center",
                "right" => "right",
                "both" => "justify",
                var other => other
            };
        }

        var spacing = pPr.SpacingBetweenLines;
        if (spacing != null)
        {
            if (spacing.Before?.Value is string before)
                props["spacing-before"] = $"{int.Parse(before) / 20}pt";
            if (spacing.After?.Value is string after)
                props["spacing-after"] = $"{int.Parse(after) / 20}pt";
            if (spacing.Line?.Value is string line)
                props["line-spacing"] = $"{int.Parse(line) / 240.0:F1}";
        }

        var indent = pPr.Indentation;
        if (indent != null)
        {
            if (indent.Left?.Value is string left)
                props["indent-left"] = $"{int.Parse(left) / 20}pt";
            if (indent.Right?.Value is string right)
                props["indent-right"] = $"{int.Parse(right) / 20}pt";
        }
    }

    private void BuildRunProperties(RunProperties rPr, Dictionary<string, object?> props)
    {
        if (rPr.Bold != null)
            props["bold"] = rPr.Bold.Val == null || rPr.Bold.Val.Value;
        if (rPr.Italic != null)
            props["italic"] = rPr.Italic.Val == null || rPr.Italic.Val.Value;
        if (rPr.Underline?.Val?.Value is UnderlineValues ul && ul != UnderlineValues.None)
            props["underline"] = true;
        if (rPr.Strike != null)
            props["strikethrough"] = rPr.Strike.Val == null || rPr.Strike.Val.Value;

        if (rPr.RunFonts?.Ascii?.Value is string fontName)
            props["font-name"] = fontName;
        if (rPr.FontSize?.Val?.Value is string fontSize)
            props["font-size"] = $"{int.Parse(fontSize) / 2}pt";
        if (rPr.Color?.Val?.Value is string color)
            props["color"] = $"#{color}";
    }

    private void BuildCellProperties(TableCell cell, Dictionary<string, object?> props)
    {
        var tcPr = cell.TableCellProperties;
        if (tcPr == null) return;

        if (tcPr.Shading?.Fill?.Value is string fill && fill != "auto")
            props["fill-color"] = $"#{fill}";

        if (tcPr.TableCellWidth?.Width?.Value is string width)
            props["width"] = $"{int.Parse(width) / 20}pt";
    }

    private void BuildTableProperties(Table table, Dictionary<string, object?> props)
    {
        var tblPr = table.GetFirstChild<TableProperties>();
        if (tblPr == null) return;

        if (tblPr.TableStyle?.Val?.Value is string tableStyle)
            props["table-style"] = tableStyle;

        var rows = table.Elements<TableRow>().Count();
        var cols = table.Elements<TableRow>().FirstOrDefault()?.Elements<TableCell>().Count() ?? 0;
        props["rows"] = rows;
        props["columns"] = cols;
    }

    #endregion

    #region Comments Layer

    private List<CommentInfo> BuildCommentsLayer(OpenXmlElement element)
    {
        var comments = new List<CommentInfo>();
        var commentIds = new HashSet<string>();

        foreach (var rangeStart in element.Descendants<CommentRangeStart>())
        {
            if (rangeStart.Id?.Value is string id)
                commentIds.Add(id);
        }

        if (commentIds.Count == 0) return comments;

        var commentsPart = _wordDoc.MainDocumentPart?.WordprocessingCommentsPart;
        if (commentsPart?.Comments == null) return comments;

        var commentMap = commentsPart.Comments.Elements<Comment>()
            .Where(c => c.Id?.Value != null)
            .ToDictionary(c => c.Id!.Value!, c => c);

        foreach (var id in commentIds)
        {
            if (commentMap.TryGetValue(id, out var comment))
            {
                var info = new CommentInfo
                {
                    Author = comment.Author?.Value ?? "Unknown",
                    Text = comment.InnerText
                };
                if (comment.Date?.Value is DateTime date)
                    info.Date = date.ToString("o");
                comments.Add(info);
            }
        }

        return comments;
    }

    #endregion

    #region Children (DEPTH)

    private List<ElementInfo> BuildChildren(
        OpenXmlElement element,
        bool includeContent, bool includeProperties,
        int remainingDepth)
    {
        var children = new List<ElementInfo>();
        var childElements = GetLogicalChildren(element);

        foreach (var child in childElements)
        {
            var info = BuildAddressingLayer(child, Array.Empty<OpenXmlElement>());

            if (includeContent)
                info.Content = BuildContentLayer(child);
            if (includeProperties)
                info.Properties = BuildPropertiesLayer(child);

            var comments = BuildCommentsLayer(child);
            if (comments.Count > 0)
                info.Comments = comments;

            if (remainingDepth > 0)
                info.Children = BuildChildren(child, includeContent, includeProperties, remainingDepth - 1);

            children.Add(info);
        }

        return children.Count > 0 ? children : null!;
    }

    private IEnumerable<OpenXmlElement> GetLogicalChildren(OpenXmlElement element)
    {
        return element switch
        {
            Table table => table.Elements<TableRow>(),
            TableRow row => row.Elements<TableCell>(),
            Paragraph para => para.Elements<Run>(),
            SdtBlock sdt => sdt.Descendants<Paragraph>(),
            _ => element.ChildElements.Where(c =>
                c is Paragraph or Table or TableRow or TableCell or Run or SdtBlock)
        };
    }

    #endregion

    #region Context Layer

    private ContextInfo BuildContextLayer(
        OpenXmlElement element,
        IReadOnlyList<OpenXmlElement> bodyChildren,
        int contextCount,
        bool includeContent)
    {
        var idx = IndexOf(bodyChildren, element);
        if (idx < 0)
        {
            // Element not a direct body child; try to find parent that is
            var parent = element.Parent;
            while (parent != null && IndexOf(bodyChildren, parent) < 0)
                parent = parent.Parent;
            if (parent != null)
                idx = IndexOf(bodyChildren, parent);
        }

        var context = new ContextInfo();
        if (idx < 0) return context;

        // Before
        for (int i = Math.Max(0, idx - contextCount); i < idx; i++)
        {
            var info = BuildAddressingLayer(bodyChildren[i], bodyChildren);
            if (includeContent)
                info.Content = BuildContentLayer(bodyChildren[i]);
            context.Before.Add(info);
        }

        // After
        for (int i = idx + 1; i <= Math.Min(bodyChildren.Count - 1, idx + contextCount); i++)
        {
            var info = BuildAddressingLayer(bodyChildren[i], bodyChildren);
            if (includeContent)
                info.Content = BuildContentLayer(bodyChildren[i]);
            context.After.Add(info);
        }

        return context;
    }

    #endregion

    #region Helpers

    private IReadOnlyList<OpenXmlElement> GetBodyChildren()
    {
        return _wordDoc.MainDocumentPart?.Document?.Body?.ChildElements
            .OfType<OpenXmlElement>()
            .ToList() ?? new List<OpenXmlElement>();
    }

    private static int GetHeadingLevel(Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrEmpty(styleId)) return 0;

        if (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(styleId.AsSpan(7), out int level))
            return level;

        var outlineLevel = para.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (outlineLevel.HasValue)
            return outlineLevel.Value + 1;

        return 0;
    }

    private static string MapElementType(OpenXmlElement element)
    {
        return element switch
        {
            Paragraph _ => "paragraph",
            Table _ => "table",
            TableRow _ => "row",
            TableCell _ => "cell",
            Run _ => "run",
            BookmarkStart _ => "bookmark",
            SdtBlock _ => "content-control",
            SdtRun _ => "content-control",
            Drawing _ => "image",
            _ => element.LocalName
        };
    }

    private static int IndexOf(IReadOnlyList<OpenXmlElement> list, OpenXmlElement item)
    {
        for (int i = 0; i < list.Count; i++)
        {
            if (ReferenceEquals(list[i], item)) return i;
        }
        return -1;
    }

    private static (int? index, int? total) GetPositionInParent(
        OpenXmlElement element, IReadOnlyList<OpenXmlElement> siblings)
    {
        var idx = IndexOf(siblings, element);
        if (idx < 0) return (null, null);
        return (idx + 1, siblings.Count);
    }

    private static (int? index, int? total) GetPositionAmong<T>(
        OpenXmlElement element, IReadOnlyList<OpenXmlElement> siblings) where T : OpenXmlElement
    {
        var typedSiblings = siblings.OfType<T>().ToList();
        var idx = typedSiblings.IndexOf((T)element);
        if (idx < 0) return (null, null);
        return (idx + 1, typedSiblings.Count);
    }

    #endregion
}
