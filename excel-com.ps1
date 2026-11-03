if (-not ('CsvExcelRunningObject' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CsvExcelRunningObject {
    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid clsid, IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object value);
    public static object Get() {
        Guid clsid = new Guid("00024500-0000-0000-C000-000000000046");
        object value;
        GetActiveObject(ref clsid, IntPtr.Zero, out value);
        return value;
    }
}
'@
}
