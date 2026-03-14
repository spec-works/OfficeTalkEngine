using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using OfficeTalk.Parsing;
using OfficeTalkEngine.Addressing;

namespace OfficeTalk.Cli.Commands;

/// <summary>
/// Handles the inspect command — resolves an address and shows matched elements.
/// </summary>
public static class InspectCommand
{
    public static int Execute(FileInfo target, string address, int context)
    {
        try
        {
            if (!target.Exists)
            {
                Console.Error.WriteLine($"Error: Target file not found: {target.FullName}");
                return 1;
            }

            // Parse the address by wrapping it in a minimal OfficeTalk document
            var otkSource = $"OFFICETALK/1.0\nDOCTYPE word\nAT {address}\nDELETE";
            var lexer = new OfficeTalkLexer(otkSource);
            var tokens = lexer.Tokenize();
            var parser = new OfficeTalkParser(tokens);
            var document = parser.Parse();

            if (document.Errors.Count > 0)
            {
                Console.Error.WriteLine($"Error: Invalid address '{address}'");
                foreach (var error in document.Errors)
                {
                    Console.Error.WriteLine($"  {error.Message}");
                }
                return 2;
            }

            if (document.OperationBlocks.Count == 0)
            {
                Console.Error.WriteLine($"Error: Could not parse address '{address}'");
                return 2;
            }

            var parsedAddress = document.OperationBlocks[0].Address;

            // Resolve against target document
            // Open via memory stream to avoid holding a file lock that blocks OneDrive sync
            byte[] fileBytes = File.ReadAllBytes(target.FullName);
            using var memoryStream = new MemoryStream(fileBytes, writable: false);
            using var wordDoc = WordprocessingDocument.Open(memoryStream, false);
            var resolver = new WordAddressResolver(wordDoc);
            var elements = resolver.Resolve(parsedAddress);

            Console.WriteLine($"Address: {parsedAddress}");
            Console.WriteLine($"Matched: {elements.Count} element(s)");

            if (elements.Count == 0)
            {
                Console.WriteLine();
                Console.WriteLine("  No elements matched this address.");
                return 0;
            }

            Console.WriteLine();

            // Get all body children for position tracking
            var bodyChildren = wordDoc.MainDocumentPart?.Document?.Body?.ChildElements
                .OfType<OpenXmlElement>()
                .ToList() ?? new List<OpenXmlElement>();

            for (int i = 0; i < elements.Count; i++)
            {
                var element = elements[i];
                PrintElement(i + 1, element, bodyChildren, context);
            }

            return 0;
        }
        catch (Exception ex) when (ex.Message.Contains("address", StringComparison.OrdinalIgnoreCase)
                                   || ex.Message.Contains("resolve", StringComparison.OrdinalIgnoreCase))
        {
            Console.Error.WriteLine($"Error: Address resolution failed — {ex.Message}");
            return 3;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
    }

    private static void PrintElement(int index, OpenXmlElement element, List<OpenXmlElement> bodyChildren, int context)
    {
        var text = element.InnerText;
        var truncatedText = text.Length > 80 ? text[..80] + "..." : text;

        if (element is Paragraph paragraph)
        {
            var level = GetHeadingLevel(paragraph);
            var styleName = GetStyleName(paragraph);
            var position = bodyChildren.IndexOf(paragraph) + 1;
            var totalParagraphs = bodyChildren.Count;

            if (level > 0)
            {
                Console.WriteLine($"  [{index}] Heading (level {level})");
            }
            else
            {
                Console.WriteLine($"  [{index}] Paragraph");
            }

            Console.WriteLine($"      Text: \"{truncatedText}\"");
            if (!string.IsNullOrEmpty(styleName))
                Console.WriteLine($"      Style: \"{styleName}\"");
            if (position > 0)
                Console.WriteLine($"      Position: paragraph {position} of {totalParagraphs}");
        }
        else if (element is Table)
        {
            var position = bodyChildren.IndexOf(element) + 1;
            var rowCount = element.Elements<TableRow>().Count();
            Console.WriteLine($"  [{index}] Table");
            Console.WriteLine($"      Rows: {rowCount}");
            if (position > 0)
                Console.WriteLine($"      Position: element {position} of {bodyChildren.Count}");
        }
        else if (element is TableRow row)
        {
            var cellCount = row.Elements<TableCell>().Count();
            Console.WriteLine($"  [{index}] Table Row");
            Console.WriteLine($"      Cells: {cellCount}");
            Console.WriteLine($"      Text: \"{truncatedText}\"");
        }
        else if (element is TableCell)
        {
            Console.WriteLine($"  [{index}] Table Cell");
            Console.WriteLine($"      Text: \"{truncatedText}\"");
        }
        else
        {
            Console.WriteLine($"  [{index}] {element.GetType().Name}");
            Console.WriteLine($"      Text: \"{truncatedText}\"");
        }

        // Show context elements
        if (context > 0)
        {
            var idx = bodyChildren.IndexOf(element);
            if (idx >= 0)
            {
                var start = Math.Max(0, idx - context);
                var end = Math.Min(bodyChildren.Count - 1, idx + context);

                if (start < idx || end > idx)
                {
                    Console.WriteLine($"      Context:");
                    for (int c = start; c <= end; c++)
                    {
                        var ctxElement = bodyChildren[c];
                        var marker = c == idx ? ">>>" : "   ";
                        var ctxText = ctxElement.InnerText;
                        var ctxTruncated = ctxText.Length > 60 ? ctxText[..60] + "..." : ctxText;
                        Console.WriteLine($"        {marker} [{c + 1}] \"{ctxTruncated}\"");
                    }
                }
            }
        }

        Console.WriteLine();
    }

    private static int GetHeadingLevel(Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrEmpty(styleId))
            return 0;

        // Match patterns like "Heading1", "Heading2", etc.
        if (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(styleId.AsSpan(7), out int level))
        {
            return level;
        }

        // Check OutlineLevel
        var outlineLevel = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (outlineLevel.HasValue)
        {
            return outlineLevel.Value + 1;
        }

        return 0;
    }

    private static string? GetStyleName(Paragraph paragraph)
    {
        return paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
    }
}
