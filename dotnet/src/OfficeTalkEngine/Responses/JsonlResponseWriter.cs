using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeTalkEngine.Responses;

/// <summary>
/// Writes OfficeTalk responses as JSONL (one JSON object per line).
/// Media type: application/officetalk-response
/// </summary>
public class JsonlResponseWriter
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly TextWriter _writer;

    public JsonlResponseWriter(TextWriter writer)
    {
        _writer = writer ?? throw new ArgumentNullException(nameof(writer));
    }

    public void WriteInspectResponse(InspectResponse response)
    {
        var json = JsonSerializer.Serialize(response, SerializerOptions);
        _writer.WriteLine(json);
    }

    public void WriteOperationResponse(OperationResponse response)
    {
        var json = JsonSerializer.Serialize(response, SerializerOptions);
        _writer.WriteLine(json);
    }

    public void Flush() => _writer.Flush();
}
