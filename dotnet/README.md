# OfficeTalkEngine — .NET Library

.NET 9 execution engine for [OfficeTalk](https://github.com/spec-works/OfficeTalk) documents. Resolves addresses, validates semantics, and applies deterministic operations to Microsoft Office documents via the OpenXML SDK.

## Quick Start

### Parse and Execute

```csharp
using OfficeTalk.Parsing;
using OfficeTalkEngine.Execution;

var source = """
    OFFICETALK/1.0 Word
    AT paragraph[1]
        SET "Hello, World!"
    """;

// Parse the OfficeTalk source
var parser = new OfficeTalkParser();
var document = parser.Parse(source);

// Execute against a Word document
var executor = new WordExecutor();
executor.Execute(document, "template.docx", "output.docx");
```

### Semantic Validation

```csharp
using OfficeTalk.Parsing;
using OfficeTalkEngine.Validation;

var parser = new OfficeTalkParser();
var document = parser.Parse(source);

var validator = new SemanticValidator();
var result = validator.Validate(document, "template.docx");

if (!result.IsValid)
{
    foreach (var error in result.Errors)
        Console.WriteLine($"Error: {error}");
}
```

### Address Resolution

```csharp
using DocumentFormat.OpenXml.Packaging;
using OfficeTalk.Ast;
using OfficeTalkEngine.Addressing;

using var doc = WordprocessingDocument.Open("document.docx", false);
var resolver = new WordAddressResolver(doc);

var address = new Address
{
    Segments = { new AddressSegment { Identifier = "heading", Predicates = { new KeyValuePredicate { Key = "level", Operator = PredicateOperator.Equals, Value = "1" } } } }
};

var elements = resolver.Resolve(address);
```

## Architecture

OfficeTalkEngine implements a three-phase execution pipeline:

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│   1. Resolve     │────▶│   2. Validate    │────▶│   3. Execute    │
│                  │     │                  │     │                 │
│ Address segments │     │ Addresses exist? │     │ Snapshot all    │
│ matched against  │     │ Styles valid?    │     │ resolutions,    │
│ document tree    │     │ Search text      │     │ then apply ops  │
│                  │     │ present?         │     │ deterministically│
└─────────────────┘     └──────────────────┘     └─────────────────┘
```

### Phase 1: Address Resolution
`IAddressResolver` implementations navigate document structure segment by segment, applying predicates (positional, text match, level) to locate target elements.

### Phase 2: Semantic Validation
`SemanticValidator` checks that all addresses resolve, REPLACE search strings exist in target content, and STYLE references map to actual document styles.

### Phase 3: Execution
`IOfficeTalkExecutor` implementations apply operations with snapshot semantics — all addresses are resolved before any mutations occur, ensuring deterministic results.

## Project Structure

```
dotnet/
├── OfficeTalkEngine.slnx
├── src/
│   └── OfficeTalkEngine/
│       ├── Addressing/          # Address resolution against documents
│       │   ├── IAddressResolver.cs
│       │   └── WordAddressResolver.cs
│       ├── Execution/           # Operation execution
│       │   ├── IOfficeTalkExecutor.cs
│       │   ├── WordExecutor.cs        # OpenXML SDK executor
│       │   └── WordComExecutor.cs     # Word COM Interop executor (live editing)
│       └── Validation/          # Semantic validation
│           └── SemanticValidator.cs
└── tests/
    ├── OfficeTalkEngine.Tests/
    │   ├── Addressing/
    │   │   └── WordAddressResolverTests.cs
    │   └── Execution/
    │       └── WordExecutorTests.cs
    └── OfficeTalkEngine.ParityTests/   # COM vs OpenXML parity tests
        └── ParityTests.cs
```

## Supported Operations

| Operation | Word (OpenXML) | Word (COM) | Excel | PowerPoint |
|-----------|:-:|:-:|:-:|:-:|
| SET | ✅ | ✅ | ✅ | ✅ |
| REPLACE / REPLACE ALL | ✅ | ✅ | — | — |
| INSERT BEFORE/AFTER | ✅ | ✅ | — | — |
| DELETE | ✅ | ✅ | ✅ | ✅ |
| APPEND / PREPEND | ✅ | ✅ | — | — |
| FORMAT | ✅ | ✅ | ✅ | ✅ |
| STYLE | ✅ | ✅ | — | — |
| COMMENT | ✅ | ✅ | ✅ | — |
| INSERT ROW/COLUMN | ✅ | ✅ | — | — |
| SET CELLS / MERGE CELLS | ✅ | ✅ | — | — |
| INSERT IMAGE | ✅ | ✅ | — | — |
| INSERT TABLE | ✅ | ✅ | — | — |
| LINK | ✅ | ✅ | — | — |
| INSERT LIST | ✅ | ✅ | — | — |
| SET RUNS | ✅ | ✅ | — | — |
| INSERT SLIDE / DUPLICATE | — | — | — | ✅ |
| ADD/RENAME/DELETE SHEET | — | — | ✅ | — |
| INSPECT | ✅ | ✅ | ✅ | ✅ |
| PROPERTY | ✅ | ✅ | — | — |

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| [SpecWorks.OfficeTalk](https://github.com/spec-works/OfficeTalk) | (project ref) | Core parsing and AST |
| [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml) | 3.1.0 | OpenXML SDK for Office documents |

## License

MIT License
