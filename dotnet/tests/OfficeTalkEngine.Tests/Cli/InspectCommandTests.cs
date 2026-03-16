using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeTalk.Cli.Commands;
using Xunit;

namespace OfficeTalkEngine.Tests.Cli;

public class InspectCommandTests : IDisposable
{
    private readonly string _tempDir;

    public InspectCommandTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"officetalk-test-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try { Directory.Delete(_tempDir, true); } catch { }
    }

    #region Helpers

    private string CreateTestDocument(Action<Body> configure)
    {
        var path = Path.Combine(_tempDir, $"test-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();
        mainPart.Document.Body = body;
        configure(body);
        mainPart.Document.Save();
        return path;
    }

    private string CreateOtkFile(string content)
    {
        var path = Path.Combine(_tempDir, $"test-{Guid.NewGuid():N}.otk");
        File.WriteAllText(path, content);
        return path;
    }

    private static Paragraph MakeParagraph(string text)
    {
        return new Paragraph(
            new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Paragraph MakeHeading(string text, int level)
    {
        return new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId { Val = $"Heading{level}" }),
            new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Table MakeTable(params string[][] rows)
    {
        var table = new Table();
        foreach (var row in rows)
        {
            var tr = new TableRow();
            foreach (var cellText in row)
            {
                tr.AppendChild(new TableCell(MakeParagraph(cellText)));
            }
            table.AppendChild(tr);
        }
        return table;
    }

    /// <summary>
    /// Runs InspectCommand.ExecuteOtk, capturing stdout and returning JSONL lines.
    /// </summary>
    private (int exitCode, string[] jsonlLines, string stderr) RunInspectOtk(
        string otkContent, string docPath, bool verbose = false)
    {
        var otkPath = CreateOtkFile(otkContent);

        var originalOut = Console.Out;
        var originalErr = Console.Error;
        var outWriter = new StringWriter();
        var errWriter = new StringWriter();

        try
        {
            Console.SetOut(outWriter);
            Console.SetError(errWriter);

            var exitCode = InspectCommand.ExecuteOtk(
                new FileInfo(otkPath), new FileInfo(docPath), verbose);

            var output = outWriter.ToString();
            var lines = output.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            return (exitCode, lines, errWriter.ToString());
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalErr);
        }
    }

    /// <summary>
    /// Runs InspectCommand.Execute (legacy --address mode), capturing stdout.
    /// </summary>
    private (int exitCode, string[] jsonlLines, string stderr) RunInspectAddress(
        string docPath, string address, int context = 0, bool verbose = false)
    {
        var originalOut = Console.Out;
        var originalErr = Console.Error;
        var outWriter = new StringWriter();
        var errWriter = new StringWriter();

        try
        {
            Console.SetOut(outWriter);
            Console.SetError(errWriter);

            var exitCode = InspectCommand.Execute(
                new FileInfo(docPath), address, context, verbose);

            var output = outWriter.ToString();
            var lines = output.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            return (exitCode, lines, errWriter.ToString());
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalErr);
        }
    }

    #endregion

    // ─── OTK mode: basic functionality ──────────────────────────

    [Fact]
    public void ExecuteOtk_SimpleInspect_ReturnsJsonl()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeHeading("Title", 1));
            body.AppendChild(MakeParagraph("Text"));
        });

        var (exitCode, lines, _) = RunInspectOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading
  INCLUDE content
", docPath);

        exitCode.Should().Be(0);
        lines.Should().HaveCount(1);

        var json = JsonDocument.Parse(lines[0]);
        json.RootElement.GetProperty("matched").GetInt32().Should().Be(1);
        json.RootElement.GetProperty("elements")[0]
            .GetProperty("content").GetProperty("text").GetString()
            .Should().Be("Title");
    }

    [Fact]
    public void ExecuteOtk_MultipleBlocks_ReturnsMultipleLines()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeHeading("H1", 1));
            body.AppendChild(MakeParagraph("P1"));
            body.AppendChild(MakeTable(new[] { "A", "B" }));
        });

        var (exitCode, lines, _) = RunInspectOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading
  INCLUDE content

INSPECT body/paragraph
  INCLUDE content

INSPECT body/table
", docPath);

        exitCode.Should().Be(0);
        lines.Should().HaveCount(3);

        // Each line should be valid JSON
        foreach (var line in lines)
        {
            var action = () => JsonDocument.Parse(line);
            action.Should().NotThrow();
        }
    }

    [Fact]
    public void ExecuteOtk_NoMatches_ReturnsZeroMatched()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeParagraph("No headings here"));
        });

        var (exitCode, lines, _) = RunInspectOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading
", docPath);

        exitCode.Should().Be(0);
        lines.Should().HaveCount(1);

        var json = JsonDocument.Parse(lines[0]);
        json.RootElement.GetProperty("matched").GetInt32().Should().Be(0);
        json.RootElement.GetProperty("elements").GetArrayLength().Should().Be(0);
    }

    [Fact]
    public void ExecuteOtk_WithDepthAndContext_PopulatesFields()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeParagraph("Before"));
            body.AppendChild(MakeTable(
                new[] { "H1", "H2" },
                new[] { "V1", "V2" }
            ));
            body.AppendChild(MakeParagraph("After"));
        });

        var (exitCode, lines, _) = RunInspectOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/table[1]
  DEPTH 2
  INCLUDE content
  CONTEXT 1
", docPath);

        exitCode.Should().Be(0);
        var json = JsonDocument.Parse(lines[0]);
        var element = json.RootElement.GetProperty("elements")[0];

        // Has children (DEPTH)
        element.GetProperty("children").GetArrayLength().Should().BeGreaterThan(0);
        // Has context
        element.GetProperty("context").GetProperty("before").GetArrayLength().Should().Be(1);
        element.GetProperty("context").GetProperty("after").GetArrayLength().Should().Be(1);
    }

    // ─── OTK mode: JSONL format compliance ──────────────────────

    [Fact]
    public void ExecuteOtk_JsonlOmitsNullFields()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeHeading("Title", 1));
        });

        var (exitCode, lines, _) = RunInspectOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading[1]
", docPath);

        exitCode.Should().Be(0);
        var json = JsonDocument.Parse(lines[0]);
        var element = json.RootElement.GetProperty("elements")[0];

        // No INCLUDE → content/properties should be absent
        element.TryGetProperty("content", out _).Should().BeFalse();
        element.TryGetProperty("properties", out _).Should().BeFalse();
        // No CONTEXT → context should be absent
        element.TryGetProperty("context", out _).Should().BeFalse();
        // No DEPTH → children should be absent
        element.TryGetProperty("children", out _).Should().BeFalse();
        // No comments → comments should be absent
        element.TryGetProperty("comments", out _).Should().BeFalse();

        // Required fields should be present
        element.GetProperty("type").GetString().Should().Be("heading");
        element.GetProperty("index").GetInt32().Should().Be(1);
        element.GetProperty("of").GetInt32().Should().Be(1);
        element.GetProperty("level").GetInt32().Should().Be(1);
    }

    [Fact]
    public void ExecuteOtk_JsonlHasCorrectOpField()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeParagraph("Test"));
        });

        var (exitCode, lines, _) = RunInspectOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/paragraph
", docPath);

        exitCode.Should().Be(0);
        var json = JsonDocument.Parse(lines[0]);
        json.RootElement.GetProperty("op").GetString().Should().Be("inspect");
    }

    // ─── OTK mode: validation and error handling ────────────────

    [Fact]
    public void ExecuteOtk_MixedOperations_ReturnsError()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeParagraph("Test"));
        });

        var (exitCode, _, stderr) = RunInspectOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading

AT body/paragraph[1]
SET ""Hello""
", docPath);

        exitCode.Should().NotBe(0);
        stderr.Should().Contain("error");
    }

    [Fact]
    public void ExecuteOtk_InvalidSyntax_ReturnsError()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeParagraph("Test"));
        });

        var (exitCode, _, stderr) = RunInspectOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT
", docPath);

        exitCode.Should().NotBe(0);
        stderr.Should().NotBeEmpty();
    }

    [Fact]
    public void ExecuteOtk_NoInspectBlocks_ReturnsError()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeParagraph("Test"));
        });

        var (exitCode, _, stderr) = RunInspectOtk(@"OFFICETALK/1.0
DOCTYPE word
", docPath);

        exitCode.Should().NotBe(0);
        stderr.Should().Contain("INSPECT");
    }

    [Fact]
    public void ExecuteOtk_NonexistentInputFile_ReturnsError()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeParagraph("Test"));
        });

        var originalOut = Console.Out;
        var originalErr = Console.Error;
        var errWriter = new StringWriter();

        try
        {
            Console.SetOut(new StringWriter());
            Console.SetError(errWriter);

            var exitCode = InspectCommand.ExecuteOtk(
                new FileInfo(Path.Combine(_tempDir, "nonexistent.otk")),
                new FileInfo(docPath), false);

            exitCode.Should().Be(1);
            errWriter.ToString().Should().Contain("not found");
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalErr);
        }
    }

    [Fact]
    public void ExecuteOtk_Verbose_WritesToStderr()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeHeading("Title", 1));
        });

        var (exitCode, _, stderr) = RunInspectOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading
  INCLUDE content
", docPath, verbose: true);

        exitCode.Should().Be(0);
        stderr.Should().Contain("Inspecting");
    }

    // ─── Legacy --address mode ──────────────────────────────────

    [Fact]
    public void Execute_BareAddress_ProducesJsonl()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeHeading("Title", 1));
            body.AppendChild(MakeParagraph("Body text"));
        });

        var (exitCode, lines, _) = RunInspectAddress(docPath, "body/heading");

        exitCode.Should().Be(0);
        lines.Should().HaveCount(1);

        var json = JsonDocument.Parse(lines[0]);
        json.RootElement.GetProperty("matched").GetInt32().Should().Be(1);
    }

    [Fact]
    public void Execute_BareAddressWithContext_IncludesContext()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeParagraph("Before"));
            body.AppendChild(MakeHeading("Title", 1));
            body.AppendChild(MakeParagraph("After"));
        });

        var (exitCode, lines, _) = RunInspectAddress(docPath, "body/heading[1]", context: 2);

        exitCode.Should().Be(0);
        var json = JsonDocument.Parse(lines[0]);
        var element = json.RootElement.GetProperty("elements")[0];
        element.GetProperty("context").GetProperty("before").GetArrayLength().Should().BeGreaterThan(0);
    }

    [Fact]
    public void Execute_BareAddressAlwaysIncludesContent()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeHeading("Title", 1));
        });

        var (exitCode, lines, _) = RunInspectAddress(docPath, "body/heading[1]");

        exitCode.Should().Be(0);
        var json = JsonDocument.Parse(lines[0]);
        var element = json.RootElement.GetProperty("elements")[0];
        // Legacy mode always includes INCLUDE content
        element.GetProperty("content").GetProperty("text").GetString().Should().Be("Title");
    }

    [Fact]
    public void Execute_BareAddressMergesWithOtkPath()
    {
        // Verify that --address and --input produce structurally identical output
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeHeading("Title", 1));
            body.AppendChild(MakeParagraph("Body"));
        });

        var (exitCode1, lines1, _) = RunInspectAddress(docPath, "body/heading");
        var (exitCode2, lines2, _) = RunInspectOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading
  INCLUDE content
", docPath);

        exitCode1.Should().Be(0);
        exitCode2.Should().Be(0);
        lines1.Should().HaveCount(lines2.Length);

        // Both should report the same match count
        var json1 = JsonDocument.Parse(lines1[0]);
        var json2 = JsonDocument.Parse(lines2[0]);
        json1.RootElement.GetProperty("matched").GetInt32()
            .Should().Be(json2.RootElement.GetProperty("matched").GetInt32());
    }

    // ─── Exit codes ─────────────────────────────────────────────

    [Fact]
    public void ExecuteOtk_Success_ReturnsZero()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeParagraph("Test"));
        });

        var (exitCode, _, _) = RunInspectOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/paragraph
", docPath);

        exitCode.Should().Be(0);
    }

    [Fact]
    public void ExecuteOtk_ParseError_ReturnsTwo()
    {
        var docPath = CreateTestDocument(body =>
        {
            body.AppendChild(MakeParagraph("Test"));
        });

        var (exitCode, _, _) = RunInspectOtk(@"OFFICETALK/1.0
DOCTYPE word

INSPECT
", docPath);

        exitCode.Should().Be(2);
    }
}
