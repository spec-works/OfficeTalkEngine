using System.Text.Json;
using FluentAssertions;
using OfficeTalk.Parsing;
using OfficeTalk.Validation;
using OfficeTalkEngine.Validation;
using Xunit;

namespace OfficeTalkEngine.Tests.TestCases;

public class NegativeTestCaseTests
{
    private static readonly string TestCasesDir = Path.Combine(
        AppContext.BaseDirectory, "testcases", "negative");

    public static IEnumerable<object[]> NegativeTestCases()
    {
        foreach (var jsonFile in Directory.GetFiles(TestCasesDir, "*.json"))
        {
            var name = Path.GetFileNameWithoutExtension(jsonFile);
            yield return new object[] { name, jsonFile };
        }
    }

    [Theory]
    [MemberData(nameof(NegativeTestCases))]
    public void Negative_fixture(string name, string jsonPath)
    {
        var json = File.ReadAllText(jsonPath);
        var testCase = JsonSerializer.Deserialize<TestCaseDefinition>(json)!;
        testCase.IsValid.Should().BeFalse("test case '{0}' should be marked as invalid", name);

        // Build input document and save to temp file for semantic validation
        var (stream, wordDoc) = TestDocumentBuilder.Build(testCase.Input);
        var tempPath = Path.Combine(Path.GetTempPath(), $"otk-test-{Guid.NewGuid()}.docx");
        try
        {
            wordDoc.Save();
            wordDoc.Dispose();
            stream.Position = 0;
            File.WriteAllBytes(tempPath, stream.ToArray());
            stream.Dispose();

            // Parse OTK
            var otkPath = Path.Combine(Path.GetDirectoryName(jsonPath)!, testCase.Operations);
            var otkText = File.ReadAllText(otkPath);
            var lexer = new OfficeTalkLexer(otkText);
            var tokens = lexer.Tokenize();
            var parser = new OfficeTalkParser(tokens);
            var otDoc = parser.Parse();

            otDoc.Errors.Should().BeEmpty(
                "OTK file '{0}' should parse without errors", testCase.Operations);

            // Validate semantically against the target document
            var validator = new SemanticValidator();
            var result = validator.Validate(otDoc, tempPath);

            result.IsValid.Should().BeFalse(
                "semantic validation should fail for '{0}'", name);

            var expectedCategory = MapErrorCategory(testCase.ExpectedError!);
            result.Errors.Should().Contain(
                e => e.Category == expectedCategory,
                "validation should report {0} for '{1}'", expectedCategory, name);
        }
        finally
        {
            if (File.Exists(tempPath))
                File.Delete(tempPath);
        }
    }

    private static ValidationCategory MapErrorCategory(string errorCode)
    {
        return errorCode switch
        {
            "ADDRESS_NOT_FOUND" => ValidationCategory.AddressNotFound,
            "ADDRESS_AMBIGUOUS" => ValidationCategory.AddressAmbiguous,
            "SEARCH_NOT_FOUND" => ValidationCategory.SearchNotFound,
            "MISSING_STYLE" => ValidationCategory.MissingStyle,
            _ => throw new ArgumentException($"Unknown error code: {errorCode}")
        };
    }
}
