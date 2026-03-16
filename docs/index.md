[!INCLUDE [](../dotnet/README.md)]

## CLI Tool

OfficeTalkEngine includes a command-line tool for applying OfficeTalk documents:

```bash
dotnet tool install -g SpecWorks.OfficeTalk.CLI
```

### Commands

| Command | Description |
|---------|-------------|
| `officetalk apply` | Apply an OfficeTalk document to an Office file |
| `officetalk validate` | Validate an OfficeTalk document (syntactic + semantic) |
| `officetalk parse` | Parse and display the AST of an OfficeTalk document |
| `officetalk inspect` | Inspect the structure of an Office document |
| `officetalk version` | Display version information |

### Example

```bash
# Apply changes to a Word document
officetalk apply -i changes.otk document.docx

# Inspect a spreadsheet's structure
officetalk inspect budget.xlsx --address "sheet[1]" --depth 1

# Validate an OfficeTalk file
officetalk validate -i review.otk
```

### Live Editing with COM

When running on Windows with an Office application open, the CLI automatically
uses COM interop to apply changes live — edits appear instantly in the open
document. This works with Word, Excel, and PowerPoint.

## Supported Formats

| Format | Operations | COM Live Edit |
|--------|-----------|---------------|
| Word (.docx) | SET, DELETE, REPLACE, FORMAT, STYLE, INSERT, COMMENT | ✅ |
| Excel (.xlsx) | SET, DELETE, COMMENT, FORMAT | ✅ |
| PowerPoint (.pptx) | SET, DELETE, COMMENT, FORMAT | ✅ |

## Test Cases

Parity test cases in [`testcases/`](https://github.com/spec-works/OfficeTalkEngine/tree/main/testcases)
verify that the execution engine produces correct results for each operation type.

## API Reference

- [OfficeTalkEngine API Documentation](api/OfficeTalkEngine.html) - Engine, executors, and address resolvers
- [OfficeTalk.CLI API Documentation](api/OfficeTalk.Cli.html) - CLI commands and infrastructure
