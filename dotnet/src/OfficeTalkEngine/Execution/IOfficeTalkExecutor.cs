using OfficeTalk.Ast;

namespace OfficeTalkEngine.Execution;

/// <summary>
/// Executes an OfficeTalk document against a target Office file,
/// applying all operations with snapshot semantics.
/// </summary>
public interface IOfficeTalkExecutor
{
    void Execute(OfficeTalkDocument document, string targetPath, string? outputPath = null);
}
