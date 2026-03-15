using System.Text.Json.Serialization;

namespace OfficeTalkEngine.Tests.TestCases;

public class TestCaseDefinition
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("description")]
    public string Description { get; set; } = string.Empty;

    [JsonPropertyName("input")]
    public TestInputSpec Input { get; set; } = new();

    [JsonPropertyName("operations")]
    public string Operations { get; set; } = string.Empty;

    [JsonPropertyName("expected")]
    public TestExpectedOutput? Expected { get; set; }

    [JsonPropertyName("expectedError")]
    public string? ExpectedError { get; set; }

    [JsonPropertyName("isValid")]
    public bool IsValid { get; set; }
}

public class TestInputSpec
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = string.Empty;

    [JsonPropertyName("body")]
    public List<BodyElement> Body { get; set; } = new();
}

public class BodyElement
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = string.Empty;

    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("level")]
    public int? Level { get; set; }

    [JsonPropertyName("style")]
    public string? Style { get; set; }

    [JsonPropertyName("rows")]
    public List<TableRowSpec>? Rows { get; set; }

    [JsonPropertyName("bookmarkName")]
    public string? BookmarkName { get; set; }

    [JsonPropertyName("tag")]
    public string? Tag { get; set; }

    [JsonPropertyName("alt")]
    public string? Alt { get; set; }
}

public class TableRowSpec
{
    [JsonPropertyName("cells")]
    public List<string> Cells { get; set; } = new();
}

public class TestExpectedOutput
{
    [JsonPropertyName("body")]
    public List<BodyElement>? Body { get; set; }

    [JsonPropertyName("properties")]
    public Dictionary<string, string>? Properties { get; set; }
}
