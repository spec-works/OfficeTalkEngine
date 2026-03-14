using System.Text.Json;
using OfficeTalk.Ast;
using OfficeTalk.Parsing;

namespace OfficeTalk.Cli.Commands;

/// <summary>
/// Handles the parse command — dumps the AST as JSON.
/// </summary>
public static class ParseCommand
{
    public static int Execute(FileInfo input, bool pretty)
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

            // Build JSON representation
            var json = new
            {
                version = document.Version,
                docType = document.DocType.ToString().ToLowerInvariant(),
                operationBlocks = document.OperationBlocks.Select(b => new
                {
                    address = b.Address.ToString(),
                    isEach = b.IsEach,
                    operations = b.Operations.Select(SerializeOperation)
                }),
                propertySettings = document.PropertySettings.Select(p => new
                {
                    name = p.Name,
                    value = p.Value
                }),
                errors = document.Errors.Select(e => new
                {
                    line = e.Line,
                    column = e.Column,
                    message = e.Message
                })
            };

            var options = new JsonSerializerOptions
            {
                WriteIndented = pretty,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            Console.WriteLine(JsonSerializer.Serialize(json, options));
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
    }

    private static object SerializeOperation(Operation op)
    {
        return op switch
        {
            SetOperation set => new
            {
                type = "SET",
                content = set.Content.Text,
                isContentBlock = set.Content.IsContentBlock
            },
            ReplaceOperation rep => new
            {
                type = "REPLACE",
                search = rep.Search,
                replacement = rep.Replacement,
                isAll = rep.IsAll
            },
            DeleteOperation del => (object)new
            {
                type = "DELETE",
                target = del.Target.ToString()
            },
            StyleOperation style => new
            {
                type = "STYLE",
                styleName = style.StyleName
            },
            FormatOperation fmt => new
            {
                type = "FORMAT",
                properties = fmt.Properties
            },
            InsertBeforeOperation ins => new
            {
                type = "INSERT_BEFORE",
                content = ins.Content.Text,
                isContentBlock = ins.Content.IsContentBlock
            },
            InsertAfterOperation ins => new
            {
                type = "INSERT_AFTER",
                content = ins.Content.Text,
                isContentBlock = ins.Content.IsContentBlock
            },
            AppendOperation app => new
            {
                type = "APPEND",
                content = app.Content.Text,
                isContentBlock = app.Content.IsContentBlock
            },
            PrependOperation prep => new
            {
                type = "PREPEND",
                content = prep.Content.Text,
                isContentBlock = prep.Content.IsContentBlock
            },
            InsertRowOperation row => new
            {
                type = "INSERT_ROW",
                position = row.Position.ToString()
            },
            InsertColumnOperation col => new
            {
                type = "INSERT_COLUMN",
                position = col.Position.ToString()
            },
            MergeCellsOperation merge => new
            {
                type = "MERGE_CELLS",
                targetAddress = merge.TargetAddress.ToString()
            },
            SetCellsOperation cells => new
            {
                type = "SET_CELLS",
                values = cells.Values
            },
            InsertSlideOperation slide => new
            {
                type = "INSERT_SLIDE",
                position = slide.Position.ToString()
            },
            DuplicateSlideOperation => (object)new
            {
                type = "DUPLICATE_SLIDE"
            },
            RenameSheetOperation rename => new
            {
                type = "RENAME_SHEET",
                newName = rename.NewName
            },
            AddSheetOperation add => new
            {
                type = "ADD_SHEET",
                name = add.Name
            },
            _ => new
            {
                type = op.GetType().Name
            }
        };
    }
}
