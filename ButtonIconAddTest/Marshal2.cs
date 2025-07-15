using System.Runtime.InteropServices;

public class Marshal2
{

    public static object? GetActiveObject(string progId, bool throwOnError = false)
    {
        if (progId == null)
            throw new ArgumentNullException(nameof(progId));

        var hr = CLSIDFromProgIDEx(progId, out var clsid);
        if (hr < 0)
        {
            if (throwOnError)
                Marshal.ThrowExceptionForHR(hr);

            return null;
        }

        hr = GetActiveObject(clsid, nint.Zero, out var obj);
        if (hr < 0)
        {
            if (throwOnError)
                Marshal.ThrowExceptionForHR(hr);

            return null;
        }
        return obj;
    }
    [DllImport("ole32")]
    private static extern int CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid lpclsid);

    [DllImport("oleaut32")]
    private static extern int GetActiveObject([MarshalAs(UnmanagedType.LPStruct)] Guid rclsid, nint pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
}