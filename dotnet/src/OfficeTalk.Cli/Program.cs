using System.CommandLine;
using OfficeTalk.Cli.Commands;

var rootCommand = new RootCommand("OfficeTalk - Apply deterministic operations to Office documents")
{
    Name = "officetalk"
};

// apply command
var applyCommand = new Command("apply", "Execute an .otk file against a target Office document");

var applyInputOption = new Option<FileInfo?>(
    aliases: new[] { "--input", "-i" },
    description: "Path to .otk file (reads from stdin if omitted and stdin is piped)"
);

var applyTargetOption = new Option<FileInfo>(
    aliases: new[] { "--target", "-t" },
    description: "Target Office document"
)
{
    IsRequired = true
};
applyTargetOption.AddValidator(result =>
{
    var fileInfo = result.GetValueForOption(applyTargetOption);
    if (fileInfo != null && !fileInfo.Exists)
    {
        result.ErrorMessage = $"Target file not found: {fileInfo.FullName}";
    }
});

var applyOutputOption = new Option<FileInfo?>(
    aliases: new[] { "--output", "-o" },
    description: "Output path (default: modify target in-place)"
);

var forceOption = new Option<bool>(
    aliases: new[] { "--force" },
    description: "Overwrite output if it already exists",
    getDefaultValue: () => false
);

var dryRunOption = new Option<bool>(
    aliases: new[] { "--dry-run" },
    description: "Validate and resolve but don't apply changes",
    getDefaultValue: () => false
);

var verboseOption = new Option<bool>(
    aliases: new[] { "--verbose", "-v" },
    description: "Show detailed operation log",
    getDefaultValue: () => false
);

applyCommand.AddOption(applyInputOption);
applyCommand.AddOption(applyTargetOption);
applyCommand.AddOption(applyOutputOption);
applyCommand.AddOption(forceOption);
applyCommand.AddOption(dryRunOption);
applyCommand.AddOption(verboseOption);

applyCommand.SetHandler((context) =>
{
    var input = context.ParseResult.GetValueForOption(applyInputOption);
    var target = context.ParseResult.GetValueForOption(applyTargetOption)!;
    var output = context.ParseResult.GetValueForOption(applyOutputOption);
    var force = context.ParseResult.GetValueForOption(forceOption);
    var dryRun = context.ParseResult.GetValueForOption(dryRunOption);
    var verbose = context.ParseResult.GetValueForOption(verboseOption);

    var exitCode = ApplyCommand.Execute(input, target, output, force, dryRun, verbose);
    context.ExitCode = exitCode;
});

rootCommand.AddCommand(applyCommand);

// validate command
var validateCommand = new Command("validate", "Check an .otk file without applying");

var validateInputOption = new Option<FileInfo>(
    aliases: new[] { "--input", "-i" },
    description: "Path to .otk file"
)
{
    IsRequired = true
};
validateInputOption.AddValidator(result =>
{
    var fileInfo = result.GetValueForOption(validateInputOption);
    if (fileInfo != null && !fileInfo.Exists)
    {
        result.ErrorMessage = $"Input file not found: {fileInfo.FullName}";
    }
});

var validateTargetOption = new Option<FileInfo?>(
    aliases: new[] { "--target", "-t" },
    description: "Target document for semantic validation"
);

var syntaxOnlyOption = new Option<bool>(
    aliases: new[] { "--syntax-only" },
    description: "Skip semantic validation",
    getDefaultValue: () => false
);

var formatOption = new Option<string>(
    aliases: new[] { "--format", "-f" },
    description: "Output format: text or json",
    getDefaultValue: () => "text"
);
formatOption.AddValidator(result =>
{
    var value = result.GetValueForOption(formatOption);
    if (value != "text" && value != "json")
    {
        result.ErrorMessage = "Format must be 'text' or 'json'";
    }
});

validateCommand.AddOption(validateInputOption);
validateCommand.AddOption(validateTargetOption);
validateCommand.AddOption(syntaxOnlyOption);
validateCommand.AddOption(formatOption);

validateCommand.SetHandler((context) =>
{
    var input = context.ParseResult.GetValueForOption(validateInputOption)!;
    var target = context.ParseResult.GetValueForOption(validateTargetOption);
    var syntaxOnly = context.ParseResult.GetValueForOption(syntaxOnlyOption);
    var format = context.ParseResult.GetValueForOption(formatOption)!;

    var exitCode = ValidateCommand.Execute(input, target, syntaxOnly, format);
    context.ExitCode = exitCode;
});

rootCommand.AddCommand(validateCommand);

// parse command
var parseCommand = new Command("parse", "Dump the AST of an .otk file as JSON");

var parseInputOption = new Option<FileInfo>(
    aliases: new[] { "--input", "-i" },
    description: "Path to .otk file"
)
{
    IsRequired = true
};
parseInputOption.AddValidator(result =>
{
    var fileInfo = result.GetValueForOption(parseInputOption);
    if (fileInfo != null && !fileInfo.Exists)
    {
        result.ErrorMessage = $"Input file not found: {fileInfo.FullName}";
    }
});

var prettyOption = new Option<bool>(
    aliases: new[] { "--pretty" },
    description: "Pretty-print JSON output",
    getDefaultValue: () => false
);

parseCommand.AddOption(parseInputOption);
parseCommand.AddOption(prettyOption);

parseCommand.SetHandler((context) =>
{
    var input = context.ParseResult.GetValueForOption(parseInputOption)!;
    var pretty = context.ParseResult.GetValueForOption(prettyOption);

    var exitCode = ParseCommand.Execute(input, pretty);
    context.ExitCode = exitCode;
});

rootCommand.AddCommand(parseCommand);

// inspect command
var inspectCommand = new Command("inspect", "Inspect a document using an .otk file with INSPECT blocks, or a bare address");

var inspectInputOption = new Option<FileInfo?>(
    aliases: new[] { "--input", "-i" },
    description: "Path to .otk file containing INSPECT blocks (reads from stdin if omitted and stdin is piped)"
);

var inspectTargetOption = new Option<FileInfo>(
    aliases: new[] { "--target", "-t" },
    description: "Target Office document"
)
{
    IsRequired = true
};
inspectTargetOption.AddValidator(result =>
{
    var fileInfo = result.GetValueForOption(inspectTargetOption);
    if (fileInfo != null && !fileInfo.Exists)
    {
        result.ErrorMessage = $"Target file not found: {fileInfo.FullName}";
    }
});

var addressOption = new Option<string?>(
    aliases: new[] { "--address", "-a" },
    description: "OfficeTalk address to resolve (legacy mode; prefer --input with .otk file)"
);

var contextOption = new Option<int>(
    aliases: new[] { "--context", "-c" },
    description: "Lines of surrounding context to show (legacy mode only)",
    getDefaultValue: () => 0
);

var inspectVerboseOption = new Option<bool>(
    aliases: new[] { "--verbose", "-v" },
    description: "Show detailed output",
    getDefaultValue: () => false
);

inspectCommand.AddOption(inspectInputOption);
inspectCommand.AddOption(inspectTargetOption);
inspectCommand.AddOption(addressOption);
inspectCommand.AddOption(contextOption);
inspectCommand.AddOption(inspectVerboseOption);

inspectCommand.SetHandler((context) =>
{
    var input = context.ParseResult.GetValueForOption(inspectInputOption);
    var target = context.ParseResult.GetValueForOption(inspectTargetOption)!;
    var address = context.ParseResult.GetValueForOption(addressOption);
    var ctxLines = context.ParseResult.GetValueForOption(contextOption);
    var verbose = context.ParseResult.GetValueForOption(inspectVerboseOption);

    int exitCode;
    if (input != null || (address == null && Console.IsInputRedirected))
    {
        // OTK mode: parse .otk file with INSPECT blocks, produce JSONL
        exitCode = InspectCommand.ExecuteOtk(input, target, verbose);
    }
    else if (address != null)
    {
        // Legacy mode: bare address, human-readable output
        exitCode = InspectCommand.Execute(target, address, ctxLines, verbose);
    }
    else
    {
        Console.Error.WriteLine("Error: Provide --input (.otk file) or --address (bare address).");
        exitCode = 1;
    }

    context.ExitCode = exitCode;
});

rootCommand.AddCommand(inspectCommand);

// version command
var versionCommand = new Command("version", "Display version information");
versionCommand.SetHandler(() =>
{
    Console.WriteLine("OfficeTalk CLI v0.1.0");
    Console.WriteLine(".NET 9 OfficeTalk Execution Engine");
});

rootCommand.AddCommand(versionCommand);

return await rootCommand.InvokeAsync(args);
