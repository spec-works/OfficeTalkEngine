using System.Text.Json;
using OfficeTalk.Parsing;
using OfficeTalk.Validation;
using OfficeTalkEngine.Validation;

namespace OfficeTalk.Cli.Commands;

/// <summary>
/// Handles the validate command — checks an .otk file without applying.
/// </summary>
public static class ValidateCommand
{
    public static int Execute(FileInfo input, FileInfo? target, bool syntaxOnly, string format)
    {
        try
        {
            if (!input.Exists)
            {
                Console.Error.WriteLine($"Error: Input file not found: {input.FullName}");
                return 1;
            }

            var source = File.ReadAllText(input.FullName);

            // Lex and parse
            var lexer = new OfficeTalkLexer(source);
            var tokens = lexer.Tokenize();
            var parser = new OfficeTalkParser(tokens);
            var document = parser.Parse();

            var allErrors = new List<DiagnosticEntry>();
            var allWarnings = new List<DiagnosticEntry>();

            // Collect parse errors
            foreach (var error in document.Errors)
            {
                allErrors.Add(new DiagnosticEntry(error.Line, error.Column, "error", "Syntax", error.Message));
            }

            // Syntactic validation
            var syntacticValidator = new SyntacticValidator();
            var syntacticResult = syntacticValidator.Validate(document);

            foreach (var error in syntacticResult.Errors)
            {
                allErrors.Add(new DiagnosticEntry(error.Line ?? 0, error.Column ?? 0, "error", error.Category.ToString(), error.Message));
            }
            foreach (var warning in syntacticResult.Warnings)
            {
                allWarnings.Add(new DiagnosticEntry(warning.Line ?? 0, warning.Column ?? 0, "warning", warning.Category.ToString(), warning.Message));
            }

            // Semantic validation
            if (target != null && !syntaxOnly && allErrors.Count == 0)
            {
                if (!target.Exists)
                {
                    Console.Error.WriteLine($"Error: Target file not found: {target.FullName}");
                    return 1;
                }

                var semanticValidator = new SemanticValidator();
                var semanticResult = semanticValidator.Validate(document, target.FullName);

                foreach (var error in semanticResult.Errors)
                {
                    allErrors.Add(new DiagnosticEntry(error.Line ?? 0, error.Column ?? 0, "error", error.Category.ToString(), error.Message));
                }
                foreach (var warning in semanticResult.Warnings)
                {
                    allWarnings.Add(new DiagnosticEntry(warning.Line ?? 0, warning.Column ?? 0, "warning", warning.Category.ToString(), warning.Message));
                }
            }

            bool isValid = allErrors.Count == 0;
            int operationBlocks = document.OperationBlocks.Count;

            if (format == "json")
            {
                var jsonOutput = new
                {
                    valid = isValid,
                    errors = allErrors.Select(e => new { line = e.Line, column = e.Column, severity = e.Severity, category = e.Category, message = e.Message }),
                    warnings = allWarnings.Select(w => new { line = w.Line, column = w.Column, severity = w.Severity, category = w.Category, message = w.Message }),
                    operationBlocks
                };

                var jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                };

                Console.WriteLine(JsonSerializer.Serialize(jsonOutput, jsonOptions));
            }
            else
            {
                // Text format
                if (isValid)
                {
                    Console.WriteLine($"\u2713 Valid. {operationBlocks} operation block(s).");
                }
                else
                {
                    Console.WriteLine($"\u2717 Invalid. {allErrors.Count} error(s), {allWarnings.Count} warning(s).");
                    Console.WriteLine();
                }

                foreach (var error in allErrors)
                {
                    Console.Error.WriteLine($"{input.Name}:{error.Line}:{error.Column}: error: {error.Message}");
                }
                foreach (var warning in allWarnings)
                {
                    Console.Error.WriteLine($"{input.Name}:{warning.Line}:{warning.Column}: warning: {warning.Message}");
                }
            }

            foreach (var warning in allWarnings)
            {
                if (format == "text")
                {
                    // Already printed above
                }
            }

            return isValid ? 0 : 2;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
    }

    private record DiagnosticEntry(int Line, int Column, string Severity, string Category, string Message);
}
