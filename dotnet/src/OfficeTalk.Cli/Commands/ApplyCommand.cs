using System.Runtime.InteropServices;
using OfficeTalk.Ast;
using OfficeTalk.Parsing;
using OfficeTalk.Validation;
using OfficeTalkEngine.Execution;
using OfficeTalkEngine.Validation;

namespace OfficeTalk.Cli.Commands;

/// <summary>
/// Handles the apply command — executes an .otk file against a target document.
/// </summary>
public static class ApplyCommand
{
    public static int Execute(
        FileInfo? input,
        FileInfo target,
        FileInfo? output,
        bool force,
        bool dryRun,
        bool verbose)
    {
        try
        {
            // Read .otk source
            string source;
            string sourceName;
            if (input != null)
            {
                if (!input.Exists)
                {
                    Console.Error.WriteLine($"Error: Input file not found: {input.FullName}");
                    return 1;
                }
                source = File.ReadAllText(input.FullName);
                sourceName = input.Name;
            }
            else if (Console.IsInputRedirected)
            {
                source = Console.In.ReadToEnd();
                sourceName = "<stdin>";
            }
            else
            {
                Console.Error.WriteLine("Error: No input specified. Provide --input or pipe .otk content via stdin.");
                return 1;
            }

            if (verbose)
            {
                Console.WriteLine($"OfficeTalk - Apply");
                Console.WriteLine($"Source:  {sourceName}");
                Console.WriteLine($"Target:  {target.FullName}");
                if (output != null)
                    Console.WriteLine($"Output:  {output.FullName}");
                Console.WriteLine();
            }

            // Check output path
            string? outputPath = output?.FullName;
            if (outputPath != null && File.Exists(outputPath) && !force)
            {
                Console.Error.WriteLine($"Error: Output file already exists: {outputPath}");
                Console.Error.WriteLine("Use --force to overwrite.");
                return 1;
            }

            // Lex
            var lexer = new OfficeTalkLexer(source);
            var tokens = lexer.Tokenize();

            // Parse
            var parser = new OfficeTalkParser(tokens);
            var document = parser.Parse();

            // Check parse errors
            if (document.Errors.Count > 0)
            {
                foreach (var error in document.Errors)
                {
                    Console.Error.WriteLine($"{sourceName}:{error.Line}:{error.Column}: error: {error.Message}");
                }
                return 2;
            }

            // Syntactic validation
            var syntacticValidator = new SyntacticValidator();
            var syntacticResult = syntacticValidator.Validate(document);

            if (!syntacticResult.IsValid)
            {
                foreach (var error in syntacticResult.Errors)
                {
                    var line = error.Line ?? 0;
                    var col = error.Column ?? 0;
                    Console.Error.WriteLine($"{sourceName}:{line}:{col}: error: {error.Message}");
                }
                foreach (var warning in syntacticResult.Warnings)
                {
                    var line = warning.Line ?? 0;
                    var col = warning.Column ?? 0;
                    Console.Error.WriteLine($"{sourceName}:{line}:{col}: warning: {warning.Message}");
                }
                return 2;
            }

            // Print warnings even if valid
            foreach (var warning in syntacticResult.Warnings)
            {
                var line = warning.Line ?? 0;
                var col = warning.Column ?? 0;
                Console.Error.WriteLine($"{sourceName}:{line}:{col}: warning: {warning.Message}");
            }

            if (verbose)
            {
                Console.WriteLine($"Parsed {document.OperationBlocks.Count} operation block(s)");
                foreach (var block in document.OperationBlocks)
                {
                    var eachStr = block.IsEach ? " EACH" : "";
                    Console.WriteLine($"  AT{eachStr} {block.Address}");
                    foreach (var op in block.Operations)
                    {
                        Console.WriteLine($"    {DescribeOperation(op)}");
                    }
                }
                Console.WriteLine();
            }

            // Dry-run: validate but don't execute
            if (dryRun)
            {
                // Run semantic validation if target is provided
                var semanticValidator = new SemanticValidator();
                var semanticResult = semanticValidator.Validate(document, target.FullName);

                if (!semanticResult.IsValid)
                {
                    foreach (var error in semanticResult.Errors)
                    {
                        var line = error.Line ?? 0;
                        var col = error.Column ?? 0;
                        Console.Error.WriteLine($"{sourceName}:{line}:{col}: error: {error.Message}");
                    }
                    return 2;
                }

                foreach (var warning in semanticResult.Warnings)
                {
                    var line = warning.Line ?? 0;
                    var col = warning.Column ?? 0;
                    Console.Error.WriteLine($"{sourceName}:{line}:{col}: warning: {warning.Message}");
                }

                var opCount = document.OperationBlocks.Sum(b => b.Operations.Count);
                Console.WriteLine($"Dry run complete. {opCount} operation(s) would be applied. No changes written.");
                return 0;
            }

            // Choose executor: if on Windows and Word has the target open, use COM
            IOfficeTalkExecutor executor;
            bool useCom = false;

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) && outputPath == null)
            {
                useCom = WordComExecutor.IsAvailable(target.FullName);
            }

            if (useCom)
            {
                if (verbose)
                    Console.WriteLine("Word is open with target document \u2014 using COM executor.");
#pragma warning disable CA1416 // Platform compatibility (guarded by IsOSPlatform check above)
                executor = new WordComExecutor();
#pragma warning restore CA1416
            }
            else
            {
                executor = new WordExecutor();
            }

            executor.Execute(document, target.FullName, outputPath);

            var totalOps = document.OperationBlocks.Sum(b => b.Operations.Count);
            var targetFile = outputPath ?? target.FullName;
            Console.WriteLine($"\u2713 Applied {totalOps} operation(s) to {targetFile}");

            return 0;
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
        catch (IOException ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
        catch (NotImplementedException ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
        catch (NotSupportedException ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
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

    private static string DescribeOperation(Operation op)
    {
        return op switch
        {
            SetOperation set => $"SET \"{Truncate(set.Content.Text, 40)}\"",
            ReplaceOperation rep => rep.IsAll
                ? $"REPLACE ALL \"{Truncate(rep.Search, 20)}\" WITH \"{Truncate(rep.Replacement, 20)}\""
                : $"REPLACE \"{Truncate(rep.Search, 20)}\" WITH \"{Truncate(rep.Replacement, 20)}\"",
            DeleteOperation del => del.Target == DeleteTarget.Element ? "DELETE" : $"DELETE {del.Target}",
            StyleOperation style => $"STYLE \"{style.StyleName}\"",
            FormatOperation => "FORMAT ...",
            InsertBeforeOperation ins => $"INSERT BEFORE \"{Truncate(ins.Content.Text, 30)}\"",
            InsertAfterOperation ins => $"INSERT AFTER \"{Truncate(ins.Content.Text, 30)}\"",
            AppendOperation app => $"APPEND \"{Truncate(app.Content.Text, 30)}\"",
            PrependOperation prep => $"PREPEND \"{Truncate(prep.Content.Text, 30)}\"",
            _ => op.GetType().Name
        };
    }

    private static string Truncate(string text, int maxLength)
    {
        if (text.Length <= maxLength) return text;
        return text[..maxLength] + "...";
    }
}
