using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeTalkEngine.Responses;

/// <summary>
/// Response for an INSPECT block — contains matched elements with detail layers.
/// Conforms to the inspect-response shape in §14.3.
/// </summary>
public class InspectResponse
{
    [JsonPropertyName("op")]
    public string Op => "inspect";

    [JsonPropertyName("address")]
    public required string Address { get; set; }

    [JsonPropertyName("matched")]
    public int Matched { get; set; }

    [JsonPropertyName("elements")]
    public List<ElementInfo> Elements { get; set; } = new();

    [JsonPropertyName("error")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Error { get; set; }
}

/// <summary>
/// Describes a matched element in an INSPECT response.
/// Conforms to the element shape in §14.3.
/// </summary>
public class ElementInfo
{
    [JsonPropertyName("type")]
    public required string Type { get; set; }

    [JsonPropertyName("index")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Index { get; set; }

    [JsonPropertyName("of")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Of { get; set; }

    // Identity fields (type-specific, always present in addressing layer)
    [JsonPropertyName("level")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Level { get; set; }

    [JsonPropertyName("style")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Style { get; set; }

    [JsonPropertyName("name")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; set; }

    [JsonPropertyName("reference")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Reference { get; set; }

    [JsonPropertyName("sheet")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Sheet { get; set; }

    [JsonPropertyName("placeholder")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Placeholder { get; set; }

    [JsonPropertyName("tag")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Tag { get; set; }

    // Detail layers
    [JsonPropertyName("content")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public ContentInfo? Content { get; set; }

    [JsonPropertyName("properties")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public Dictionary<string, object?>? Properties { get; set; }

    [JsonPropertyName("children")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<ElementInfo>? Children { get; set; }

    [JsonPropertyName("comments")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<CommentInfo>? Comments { get; set; }

    [JsonPropertyName("context")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public ContextInfo? Context { get; set; }
}

/// <summary>
/// Content layer — present when INCLUDE content is specified.
/// </summary>
public class ContentInfo
{
    [JsonPropertyName("text")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Text { get; set; }

    [JsonPropertyName("value")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; set; }

    [JsonPropertyName("dataType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DataType { get; set; }

    [JsonPropertyName("cells")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<string>? Cells { get; set; }
}

/// <summary>
/// Comment information — always included when comments exist on the element.
/// </summary>
public class CommentInfo
{
    [JsonPropertyName("author")]
    public required string Author { get; set; }

    [JsonPropertyName("text")]
    public required string Text { get; set; }

    [JsonPropertyName("date")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Date { get; set; }
}

/// <summary>
/// Context — sibling elements before and after the match.
/// </summary>
public class ContextInfo
{
    [JsonPropertyName("before")]
    public List<ElementInfo> Before { get; set; } = new();

    [JsonPropertyName("after")]
    public List<ElementInfo> After { get; set; } = new();
}

/// <summary>
/// Response for a write operation.
/// Conforms to the operation-response shape in §14.3.
/// </summary>
public class OperationResponse
{
    [JsonPropertyName("op")]
    public required string Op { get; set; }

    [JsonPropertyName("address")]
    public required string Address { get; set; }

    [JsonPropertyName("status")]
    public required string Status { get; set; }

    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; set; }

    [JsonPropertyName("detail")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public object? Detail { get; set; }
}
