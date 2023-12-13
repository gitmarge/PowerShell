# Code example: Access Win32 API function with PowerShell and C#
#
# Microsoft Resources
# https://learn.microsoft.com/en-us/windows/win32/api/powersetting/nf-powersetting-powergetactivescheme
# https://devblogs.microsoft.com/scripting/use-powershell-to-interact-with-the-windows-api-part-1/

Add-Type @'
using System;
using System.Runtime;
using System.Runtime.InteropServices;

public class ActivePowerScheme
{
    [DllImport("PowrProf.dll")]
    public static extern int PowerGetActiveScheme(ref string UserRootPowerKey, out IntPtr ActivePolicyGuid);

    public struct GuidInfo
    {
        public Guid GUID;
    }

    public static GuidInfo GetGuid()
    {
        string urpk = null;
        IntPtr apg = IntPtr.Zero;
        int a = PowerGetActiveScheme(ref urpk, out apg);
        GuidInfo GuidInfoStruct = new GuidInfo();
        GuidInfoStruct.GUID = Marshal.PtrToStructure<Guid>(apg);
        return GuidInfoStruct;
    }
}
'@

$ActivePowerSchemeGuid = [ActivePowerScheme]::GetGuid()
$ActivePowerSchemeGuid
