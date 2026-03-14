using System.Text.Json;
using FluentAssertions;
using OfficeTalk.Parsing;
using OfficeTalkEngine.Execution;
using Xunit;

namespace OfficeTalkEngine.Tests.TestCases;

public class PositiveTestCaseTests
{
    private static readonly string TestCasesDir = Path.Combine(
        AppContext.BaseDirectory, "testcases");

    public static IEnumerable<object[]> PositiveTestCases()
    {
        foreach (var jsonFile in Directory.GetFiles(TestCasesDir, "*.json"))
        {
            var name = Path.GetFileNameWithoutExtension(jsonFile);
            yield return new object[] { name, jsonFile };
        }
    }

    [Theory]
    [MemberData(nameof(PositiveTestCases))]
    public void Positive_fixture(string name, string jsonPath)
    {
        var json = File.ReadAllText(jsonPath);
        var testCase = JsonSerializer.Deserialize<TestCaseDefinition>(json)!;
        testCase.IsValid.Should().BeTrue("test case '{0}' should be marked as valid", name);

        // Build input document
        var (stream, wordDoc) = TestDocumentBuilder.Build(testCase.Input);
        using var disposableStream = stream;
        using var doc = wordDoc;

        // Parse OTK
        var otkPath = Path.Combine(Path.GetDirectoryName(jsonPath)!, testCase.Operations);
        var otkText = File.ReadAllText(otkPath);
        var lexer = new OfficeTalkLexer(otkText);
        var tokens = lexer.Tokenize();
        var parser = new OfficeTalkParser(tokens);
        var otDoc = parser.Parse();

        otDoc.Errors.Should().BeEmpty("OTK file '{0}' should parse without errors", testCase.Operations);

        // Execute
        var executor = new WordExecutor();
        executor.Execute(otDoc, doc);

        // Verify
        TestDocumentReader.AssertExpected(doc, testCase.Expected!);
    }
}
