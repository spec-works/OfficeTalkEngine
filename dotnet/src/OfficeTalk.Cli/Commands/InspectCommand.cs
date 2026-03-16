using OfficeTalk.Ast;
using OfficeTalk.Parsing;
using OfficeTalk.Validation;
using OfficeTalkEngine.Inspection;
using OfficeTalkEngine.Responses;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeTalk.Cli.Commands;

/// <summary>
/// Handles the inspect command — resolves addresses and shows matched elements.
/// Supports both OTK documents (with INSPECT blocks) and bare --address mode.
/// </summary>
public static class InspectCommand
{
    /// <summary>
    /// Execute inspect using an OTK document containing INSPECT blocks.
    /// Produces JSONL output conforming to §14.3.
    /// </summary>
    public static int ExecuteOtk(FileInfo? input, FileInfo target, bool verbose)
    {
        try
        {
            // Read OTK source
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

            // Parse
            var lexer = new OfficeTalkLexer(source);
            var tokens = lexer.Tokenize();
            var parser = new OfficeTalkParser(tokens);
            var document = parser.Parse();

            if (document.Errors.Count > 0)
            {
                foreach (var error in document.Errors)
                {
                    Console.Error.WriteLine($"{sourceName}:{error.Line}:{error.Column}: error: {error.Message}");
                }
                return 2;
            }

            // Validate: must contain only INSPECT blocks
            var validator = new SyntacticValidator();
            var result = validator.Validate(document);
            if (!result.IsValid)
            {
                foreach (var error in result.Errors)
                {
                    Console.Error.WriteLine($"{sourceName}:{error.Line ?? 0}:{error.Column ?? 0}: error: {error.Message}");
                }
                return 2;
            }

            if (document.InspectBlocks.Count == 0)
            {
                Console.Error.WriteLine($"{sourceName}: error: Document contains no INSPECT blocks.");
                return 2;
            }

            if (document.OperationBlocks.Count > 0 || document.PropertySettings.Count > 0)
            {
                Console.Error.WriteLine($"{sourceName}: error: INSPECT document must not contain write operations (AT/PROPERTY). Use 'apply' for write operations.");
                return 2;
            }

            if (verbose)
            {
                Console.Error.WriteLine($"Inspecting {document.InspectBlocks.Count} block(s) against {target.FullName}");
            }

            // Sync COM state before reading
            ComSync.SyncIfNeeded(target.FullName, verbose);

            var extension = Path.GetExtension(target.FullName).ToLowerInvariant();

            using var memoryStream = new MemoryStream();
            using (var fs = new FileStream(target.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                fs.CopyTo(memoryStream);
            }
            memoryStream.Position = 0;

            var writer = new JsonlResponseWriter(Console.Out);

            switch (extension)
            {
                case ".docx" or ".docm":
                {
                    using var wordDoc = WordprocessingDocument.Open(memoryStream, false);
                    var inspector = new WordInspector(wordDoc);
                    var responses = inspector.Inspect(document);
                    foreach (var response in responses)
                    {
                        writer.WriteInspectResponse(response);
                    }
                    break;
                }
                default:
                    Console.Error.WriteLine($"Error: INSPECT not yet supported for '{extension}'. Currently supports: .docx");
                    return 1;
            }

            writer.Flush();
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

    /// <summary>
    /// Legacy mode: bare --address. Synthesizes a proper INSPECT document
    /// and delegates to the OTK path.
    /// </summary>
    public static int Execute(FileInfo target, string address, int context, bool verbose)
    {
        // Determine DOCTYPE from file extension
        var extension = Path.GetExtension(target.FullName).ToLowerInvariant();
        var docType = extension switch
        {
            ".docx" or ".docm" => "word",
            ".xlsx" or ".xlsm" => "excel",
            ".pptx" or ".pptm" => "powerpoint",
            _ => "word"
        };

        // Build a proper INSPECT document
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("OFFICETALK/1.0");
        sb.AppendLine($"DOCTYPE {docType}");
        sb.AppendLine();
        sb.AppendLine($"INSPECT {address}");
        sb.AppendLine("  INCLUDE content");
        if (context > 0)
            sb.AppendLine($"  CONTEXT {context}");

        // Write to a temp file and delegate to the OTK path
        var tempFile = Path.GetTempFileName();
        try
        {
            File.WriteAllText(tempFile, sb.ToString());
            return ExecuteOtk(new FileInfo(tempFile), target, verbose);
        }
        finally
        {
            File.Delete(tempFile);
        }
    }
}
