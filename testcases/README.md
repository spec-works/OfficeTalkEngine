# OfficeTalkEngine Test Cases

Shared, language-independent test cases for the OfficeTalkEngine component.

## Format

Each test case consists of:
- An `.otk` file containing the OfficeTalk operations to apply
- A `.json` file describing the input document structure, expected output, and validation expectations

The JSON file includes:
- `name`: Test case name
- `description`: What the test verifies
- `input`: Description of the target document structure (paragraphs, headings, tables)
- `operations`: The `.otk` filename
- `expected`: Expected state after execution
- `isValid`: Whether the operations should execute successfully

### Positive Tests (Execution)

Files in the root of this directory test successful operation execution.

| File | Description |
|------|-------------|
| `set-paragraph.otk` / `.json` | SET replaces paragraph text content |
| `replace-in-paragraph.otk` / `.json` | REPLACE finds and replaces text within a paragraph |
| `replace-all-in-body.otk` / `.json` | REPLACE ALL replaces all occurrences across the body |
| `delete-paragraph.otk` / `.json` | DELETE removes a paragraph element |
| `append-text.otk` / `.json` | APPEND adds text to the end of a paragraph |
| `prepend-text.otk` / `.json` | PREPEND adds text to the beginning of a paragraph |
| `style-paragraph.otk` / `.json` | STYLE applies a named paragraph style |
| `set-properties.otk` / `.json` | PROPERTY sets document metadata |
| `resolve-heading.otk` / `.json` | Address resolution for headings with level and text predicates |
| `resolve-table-cell.otk` / `.json` | Address resolution for table/row/cell paths |
| `snapshot-semantics.otk` / `.json` | Verifies operations use snapshot (pre-resolution) semantics |
| `multiple-operations.otk` / `.json` | Multiple operations in a single block applied sequentially |

### Negative Tests (Validation Failures)

Files in the `negative/` subdirectory test semantic validation failures.

| File | Description |
|------|-------------|
| `address-not-found.otk` / `.json` | Address that doesn't match any element |
| `address-ambiguous.otk` / `.json` | Address matching multiple elements without EACH |
| `replace-search-not-found.otk` / `.json` | REPLACE where search text doesn't exist |
| `style-not-found.otk` / `.json` | STYLE referencing a non-existent style |
