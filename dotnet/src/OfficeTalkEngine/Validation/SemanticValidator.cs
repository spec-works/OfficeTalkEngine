using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeTalk.Ast;
using OfficeTalk.Validation;
using OfficeTalkEngine.Addressing;

namespace OfficeTalkEngine.Validation;

/// <summary>
/// Validates an OfficeTalk document against a target Office document,
/// checking that addresses resolve, styles exist, and search strings are present.
/// </summary>
public class SemanticValidator
{
    public ValidationResult Validate(OfficeTalkDocument document, string targetPath)
    {
        return document.DocType switch
        {
            DocType.Word => ValidateWord(document, targetPath),
            _ => throw new NotImplementedException(
                $"Semantic validation for {document.DocType} documents is not yet supported.")
        };
    }

    private static ValidationResult ValidateWord(OfficeTalkDocument document, string targetPath)
    {
        var result = new ValidationResult();

        // Open with FileShare.ReadWrite so we can validate even if Word has the file open
        using var memoryStream = new MemoryStream();
        using (var fs = new FileStream(targetPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            fs.CopyTo(memoryStream);
        }
        memoryStream.Position = 0;
        using var wordDoc = WordprocessingDocument.Open(memoryStream, false);
        var resolver = new WordAddressResolver(wordDoc);

        foreach (var block in document.OperationBlocks)
        {
            ValidateBlock(block, resolver, wordDoc, result);
        }

        return result;
    }

    private static void ValidateBlock(
        OperationBlock block,
        WordAddressResolver resolver,
        WordprocessingDocument wordDoc,
        ValidationResult result)
    {
        var resolved = resolver.Resolve(block.Address);

        if (resolved.Count == 0)
        {
            result.Errors.Add(new ValidationDiagnostic
            {
                Category = ValidationCategory.AddressNotFound,
                Message = $"Address '{block.Address}' did not match any elements in the document.",
                Line = block.Line
            });
            return;
        }

        if (!block.IsEach && resolved.Count > 1)
        {
            result.Errors.Add(new ValidationDiagnostic
            {
                Category = ValidationCategory.AddressAmbiguous,
                Message = $"Address '{block.Address}' matched {resolved.Count} elements but 'AT EACH' was not specified.",
                Line = block.Line
            });
        }

        foreach (var operation in block.Operations)
        {
            ValidateOperation(operation, resolved, wordDoc, result);
        }
    }

    private static void ValidateOperation(
        Operation operation,
        IReadOnlyList<DocumentFormat.OpenXml.OpenXmlElement> elements,
        WordprocessingDocument wordDoc,
        ValidationResult result)
    {
        switch (operation)
        {
            case ReplaceOperation replace:
                ValidateReplace(replace, elements, result);
                break;
            case StyleOperation style:
                ValidateStyle(style, wordDoc, result);
                break;
        }
    }

    private static void ValidateReplace(
        ReplaceOperation operation,
        IReadOnlyList<DocumentFormat.OpenXml.OpenXmlElement> elements,
        ValidationResult result)
    {
        foreach (var element in elements)
        {
            string text = element.InnerText;
            if (!text.Contains(operation.Search, StringComparison.Ordinal))
            {
                result.Errors.Add(new ValidationDiagnostic
                {
                    Category = ValidationCategory.SearchNotFound,
                    Message = $"Search text '{operation.Search}' was not found in target element content.",
                    Line = operation.Line
                });
            }
        }
    }

    private static void ValidateStyle(
        StyleOperation operation,
        WordprocessingDocument wordDoc,
        ValidationResult result)
    {
        var stylesPart = wordDoc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart == null)
        {
            result.Errors.Add(new ValidationDiagnostic
            {
                Category = ValidationCategory.MissingStyle,
                Message = $"Style '{operation.StyleName}' not found — document has no styles defined.",
                Line = operation.Line
            });
            return;
        }

        var styles = stylesPart.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>();
        bool exists = styles.Any(s =>
            s.StyleId?.Value == operation.StyleName ||
            s.StyleName?.Val?.Value?.Equals(operation.StyleName, StringComparison.OrdinalIgnoreCase) == true);

        if (!exists)
        {
            result.Errors.Add(new ValidationDiagnostic
            {
                Category = ValidationCategory.MissingStyle,
                Message = $"Style '{operation.StyleName}' was not found in the document.",
                Line = operation.Line
            });
        }
    }
}
