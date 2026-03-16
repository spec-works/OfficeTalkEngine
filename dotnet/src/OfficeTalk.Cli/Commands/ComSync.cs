using System.Runtime.InteropServices;
using OfficeTalkEngine.Execution;

namespace OfficeTalk.Cli.Commands;

/// <summary>
/// Ensures the on-disk copy of a document matches the in-memory state
/// when an Office application has it open via COM. Call before any
/// operation that reads from disk (INSPECT, validation, etc.).
/// </summary>
internal static class ComSync
{
    /// <summary>
    /// If the target document is open in an Office app via COM, save it
    /// to disk so that subsequent file reads see the current state.
    /// </summary>
    public static void SyncIfNeeded(string targetPath, bool verbose = false)
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return;

        var extension = Path.GetExtension(targetPath).ToLowerInvariant();

        try
        {
            switch (extension)
            {
                case ".docx" or ".docm":
                    SyncWordDocument(targetPath, verbose);
                    break;
                case ".xlsx" or ".xlsm":
                    SyncExcelDocument(targetPath, verbose);
                    break;
                case ".pptx" or ".pptm":
                    SyncPowerPointDocument(targetPath, verbose);
                    break;
            }
        }
        catch
        {
            // Best-effort: if COM sync fails, caller still reads from disk
        }
    }

#pragma warning disable CA1416 // Platform compatibility (guarded by RuntimeInformation check)

    private static void SyncWordDocument(string targetPath, bool verbose)
    {
        if (!WordComExecutor.IsAvailable(targetPath)) return;

        if (verbose)
            Console.Error.WriteLine("Word has document open — syncing to disk.");

        dynamic wordApp = ComInteropHelper.GetActiveObject("Word.Application");
        SaveMatchingDocument(wordApp.Documents, targetPath);
    }

    private static void SyncExcelDocument(string targetPath, bool verbose)
    {
        if (!ExcelComExecutor.IsAvailable(targetPath)) return;

        if (verbose)
            Console.Error.WriteLine("Excel has workbook open — syncing to disk.");

        dynamic excelApp = ComInteropHelper.GetActiveObject("Excel.Application");
        SaveMatchingDocument(excelApp.Workbooks, targetPath);
    }

    private static void SyncPowerPointDocument(string targetPath, bool verbose)
    {
        if (!PowerPointComExecutor.IsAvailable(targetPath)) return;

        if (verbose)
            Console.Error.WriteLine("PowerPoint has presentation open — syncing to disk.");

        dynamic pptApp = ComInteropHelper.GetActiveObject("PowerPoint.Application");
        SaveMatchingDocument(pptApp.Presentations, targetPath);
    }

    private static void SaveMatchingDocument(dynamic documents, string targetPath)
    {
        var normalizedTarget = Path.GetFullPath(targetPath);
        foreach (dynamic doc in documents)
        {
            if (string.Equals(Path.GetFullPath(doc.FullName), normalizedTarget, StringComparison.OrdinalIgnoreCase))
            {
                doc.Save();
                break;
            }
        }
    }

#pragma warning restore CA1416
}
