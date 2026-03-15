using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace OfficeTalkEngine.Execution;

/// <summary>
/// Provides GetActiveObject functionality removed in .NET Core/.NET 5+.
/// Uses direct P/Invoke to oleaut32.dll.
/// </summary>
[SupportedOSPlatform("windows")]
internal static class ComInteropHelper
{
    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    /// <summary>
    /// Gets a running COM object registered with the specified ProgID.
    /// </summary>
    public static object GetActiveObject(string progId)
    {
        var type = Type.GetTypeFromProgID(progId, throwOnError: true)!;
        GetActiveObject(type.GUID, IntPtr.Zero, out var obj);
        return obj;
    }
}
