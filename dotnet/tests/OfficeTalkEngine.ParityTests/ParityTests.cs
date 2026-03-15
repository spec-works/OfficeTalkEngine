using System.Runtime.InteropServices;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeTalk.Parsing;
using OfficeTalkEngine.Execution;
using OfficeTalkEngine.Tests.TestCases;

namespace OfficeTalkEngine.ParityTests;

/// <summary>
/// Verifies that the COM executor and OpenXML executor produce identical results
/// for every test case. Requires Word to be running — skipped otherwise.
/// Separate project so CI can exclude it easily.
/// </summary>
public class ParityTests : IDisposable
{
    private static readonly string TestCasesDir = Path.Combine(
        AppContext.BaseDirectory, "testcases");

    private readonly List<string> _tempFiles = new();

    public static IEnumerable<object[]> TestCases()
    {
        if (!Directory.Exists(TestCasesDir))
            yield break;

        foreach (var jsonFile in Directory.GetFiles(TestCasesDir, "*.json"))
        {
            var json = File.ReadAllText(jsonFile);
            var tc = JsonSerializer.Deserialize<TestCaseDefinition>(json)!;
            if (!tc.IsValid) continue; // skip negative cases

            var name = Path.GetFileNameWithoutExtension(jsonFile);
            yield return new object[] { name, jsonFile };
        }
    }

    [SkippableTheory]
    [MemberData(nameof(TestCases))]
    public void COM_and_OpenXml_produce_same_result(string name, string jsonPath)
    {
        Skip.IfNot(RuntimeInformation.IsOSPlatform(OSPlatform.Windows),
            "COM parity tests require Windows");
        Skip.IfNot(TryGetWordApp(out var wordApp),
            "COM parity tests require Word to be running");

        var json = File.ReadAllText(jsonPath);
        var testCase = JsonSerializer.Deserialize<TestCaseDefinition>(json)!;

        var otkPath = Path.Combine(Path.GetDirectoryName(jsonPath)!, testCase.Operations);
        var otkText = File.ReadAllText(otkPath);

        // ── OpenXML path ──
        var otDocXml = ParseOtk(otkText);
        var (xmlStream, xmlBuildDoc) = TestDocumentBuilder.Build(testCase.Input);
        using (xmlStream)
        using (xmlBuildDoc)
        {
            new WordExecutor().Execute(otDocXml, xmlBuildDoc);
            xmlBuildDoc.Save();
        }
        var openXmlBytes = xmlStream.ToArray();

        // ── COM path ──
        var otDocCom = ParseOtk(otkText);
        var comInputPath = CreateTempDocx($"parity-in-{name}");
        var comOutputPath = CreateTempDocx($"parity-out-{name}");

        // Build fresh input doc and save to disk
        var (comStream, comBuildDoc) = TestDocumentBuilder.Build(testCase.Input);
        using (comStream)
        using (comBuildDoc)
        {
            comBuildDoc.Save();
        }
        File.WriteAllBytes(comInputPath, comStream.ToArray());

        // Open in Word (hidden), apply via COM, save to output, close
        dynamic doc = wordApp!.Documents.Open(comInputPath, ReadOnly: false, Visible: false);
        try
        {
            var comExecutor = new WordComExecutor();
            comExecutor.Execute(otDocCom, comInputPath);
            doc.SaveAs2(comOutputPath);
        }
        finally
        {
            doc.Close(0); // wdDoNotSaveChanges
        }

        // ── Compare ──
        CompareDocuments(openXmlBytes, comOutputPath, name);
    }

    #region Helpers

    private static OfficeTalk.Ast.OfficeTalkDocument ParseOtk(string text)
    {
        var lexer = new OfficeTalkLexer(text);
        var tokens = lexer.Tokenize();
        var parser = new OfficeTalkParser(tokens);
        return parser.Parse();
    }

    private static void CompareDocuments(byte[] openXmlBytes, string comOutputPath, string testName)
    {
        using var xmlStream = new MemoryStream(openXmlBytes);
        using var xmlDoc = WordprocessingDocument.Open(xmlStream, false);
        using var comDoc = WordprocessingDocument.Open(comOutputPath, false);

        var xmlElements = ReadBody(xmlDoc.MainDocumentPart!.Document.Body!);
        var comElements = ReadBody(comDoc.MainDocumentPart!.Document.Body!);

        comElements.Should().HaveCount(xmlElements.Count,
            "[{0}] body element count mismatch", testName);

        for (int i = 0; i < xmlElements.Count; i++)
        {
            var expected = xmlElements[i];
            var actual = comElements[i];

            actual.Type.Should().Be(expected.Type,
                "[{0}] element {1} type", testName, i);
            actual.Text.Should().Be(expected.Text,
                "[{0}] element {1} text", testName, i);

            if (expected.StyleId != null)
                actual.StyleId.Should().Be(expected.StyleId,
                    "[{0}] element {1} style", testName, i);

            if (expected.IsBold.HasValue)
                actual.IsBold.Should().Be(expected.IsBold,
                    "[{0}] element {1} bold", testName, i);

            if (expected.IsItalic.HasValue)
                actual.IsItalic.Should().Be(expected.IsItalic,
                    "[{0}] element {1} italic", testName, i);

            if (expected.FontName != null)
                actual.FontName.Should().Be(expected.FontName,
                    "[{0}] element {1} font-name", testName, i);

            if (expected.FontSize != null)
                actual.FontSize.Should().Be(expected.FontSize,
                    "[{0}] element {1} font-size", testName, i);

            if (expected.Rows != null)
            {
                actual.Rows.Should().NotBeNull("[{0}] element {1} should be a table", testName, i);
                actual.Rows!.Count.Should().Be(expected.Rows.Count,
                    "[{0}] element {1} row count", testName, i);
                for (int r = 0; r < expected.Rows.Count; r++)
                {
                    actual.Rows[r].Should().Equal(expected.Rows[r],
                        "[{0}] element {1} row {2}", testName, i, r);
                }
            }
        }

        // Compare document properties
        var xmlProps = xmlDoc.PackageProperties;
        var comProps = comDoc.PackageProperties;
        if (xmlProps.Title != null)
            comProps.Title.Should().Be(xmlProps.Title, "[{0}] title", testName);
        if (xmlProps.Subject != null)
            comProps.Subject.Should().Be(xmlProps.Subject, "[{0}] subject", testName);
        if (xmlProps.Creator != null)
            comProps.Creator.Should().Be(xmlProps.Creator, "[{0}] author", testName);
    }

    private static List<DocSnapshot> ReadBody(Body body)
    {
        var result = new List<DocSnapshot>();
        foreach (var child in body.ChildElements)
        {
            switch (child)
            {
                case Paragraph para:
                    var text = para.InnerText;
                    var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

                    // skip empty non-styled paragraphs (trailing empties)
                    if (string.IsNullOrEmpty(text) && styleId == null)
                        continue;

                    var firstRun = para.Descendants<Run>().FirstOrDefault();
                    var rp = firstRun?.RunProperties;

                    result.Add(new DocSnapshot
                    {
                        Type = styleId?.StartsWith("Heading") == true ? "heading" : "paragraph",
                        Text = text,
                        StyleId = styleId,
                        IsBold = rp?.Bold != null ? (rp.Bold.Val?.Value ?? true) : null,
                        IsItalic = rp?.Italic != null ? (rp.Italic.Val?.Value ?? true) : null,
                        FontName = rp?.RunFonts?.Ascii?.Value,
                        FontSize = rp?.FontSize?.Val?.Value,
                    });
                    break;

                case Table table:
                    result.Add(new DocSnapshot
                    {
                        Type = "table",
                        Rows = table.Elements<TableRow>()
                            .Select(row => row.Elements<TableCell>()
                                .Select(c => c.InnerText).ToList())
                            .ToList(),
                    });
                    break;
            }
        }
        return result;
    }

    private string CreateTempDocx(string prefix)
    {
        var path = Path.Combine(Path.GetTempPath(), $"{prefix}-{Guid.NewGuid():N}.docx");
        _tempFiles.Add(path);
        return path;
    }

    private static bool TryGetWordApp(out dynamic? wordApp)
    {
        wordApp = null;
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return false;
        try
        {
            wordApp = ComInteropHelper.GetActiveObject("Word.Application");
            return true;
        }
        catch
        {
            return false;
        }
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
        {
            try { File.Delete(f); } catch { }
        }
    }

    #endregion

    private class DocSnapshot
    {
        public string Type { get; set; } = "";
        public string? Text { get; set; }
        public string? StyleId { get; set; }
        public bool? IsBold { get; set; }
        public bool? IsItalic { get; set; }
        public string? FontName { get; set; }
        public string? FontSize { get; set; }
        public List<List<string>>? Rows { get; set; }
    }
}
