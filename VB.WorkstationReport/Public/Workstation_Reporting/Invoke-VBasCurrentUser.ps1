# ============================================================
# FUNCTION : Invoke-VBasCurrentUser
# MODULE   : VB.WorkstationReport
# VERSION  : 1.0.0
# CHANGED  : 23-04-2026 -- Initial release
# AUTHOR   : Vibhu Bhatnagar
# PURPOSE  : Run a scriptblock as the currently logged-on user from SYSTEM context
# ENCODING : UTF-8 with BOM
# ============================================================
# Source: Adapted from RunAsUser by KelvinTegelaar (MIT Licence)
# Ref   : https://github.com/KelvinTegelaar/RunAsUser
# ============================================================

$script:source = @"
using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;

namespace RunAsUser
{
    internal class NativeHelpers
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct LUID
        {
            public int LowPart;
            public int HighPart;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct LUID_AND_ATTRIBUTES
        {
            public LUID Luid;
            public PrivilegeAttributes Attributes;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct PROCESS_INFORMATION
        {
            public IntPtr hProcess;
            public IntPtr hThread;
            public int dwProcessId;
            public int dwThreadId;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct STARTUPINFO
        {
            public int cb;
            public String lpReserved;
            public String lpDesktop;
            public String lpTitle;
            public uint dwX;
            public uint dwY;
            public uint dwXSize;
            public uint dwYSize;
            public uint dwXCountChars;
            public uint dwYCountChars;
            public uint dwFillAttribute;
            public uint dwFlags;
            public short wShowWindow;
            public short cbReserved2;
            public IntPtr lpReserved2;
            public IntPtr hStdInput;
            public IntPtr hStdOutput;
            public IntPtr hStdError;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct TOKEN_PRIVILEGES
        {
            public int PrivilegeCount;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
            public LUID_AND_ATTRIBUTES[] Privileges;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct WTS_SESSION_INFO
        {
            public readonly UInt32 SessionID;
            [MarshalAs(UnmanagedType.LPStr)]
            public readonly String pWinStationName;
            public readonly WTS_CONNECTSTATE_CLASS State;
        }

        public struct SECURITY_ATTRIBUTES
        {
            public Int32 nLength;
            public IntPtr lpSecurityDescriptor;
            public int bInheritHandle;
        }
    }

    internal class NativeMethods
    {
        [DllImport("kernel32", SetLastError = true)]
        public static extern int WaitForSingleObject(IntPtr hHandle, int dwMilliseconds);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool CloseHandle(IntPtr hSnapshot);

        [DllImport("userenv.dll", SetLastError = true)]
        public static extern bool CreateEnvironmentBlock(ref IntPtr lpEnvironment, SafeHandle hToken, bool bInherit);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool CreateProcessAsUserW(
            SafeHandle hToken,
            String lpApplicationName,
            StringBuilder lpCommandLine,
            IntPtr lpProcessAttributes,
            IntPtr lpThreadAttributes,
            bool bInheritHandle,
            uint dwCreationFlags,
            IntPtr lpEnvironment,
            String lpCurrentDirectory,
            ref NativeHelpers.STARTUPINFO lpStartupInfo,
            out NativeHelpers.PROCESS_INFORMATION lpProcessInformation);

        [DllImport("userenv.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DestroyEnvironmentBlock(IntPtr lpEnvironment);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool DuplicateTokenEx(
            SafeHandle ExistingTokenHandle,
            uint dwDesiredAccess,
            IntPtr lpThreadAttributes,
            SECURITY_IMPERSONATION_LEVEL ImpersonationLevel,
            TOKEN_TYPE TokenType,
            out SafeNativeHandle DuplicateTokenHandle);

        [DllImport("kernel32")]
        public static extern IntPtr GetCurrentProcess();

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool GetTokenInformation(
            SafeHandle TokenHandle,
            uint TokenInformationClass,
            SafeMemoryBuffer TokenInformation,
            int TokenInformationLength,
            out int ReturnLength);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool LookupPrivilegeName(
            string lpSystemName,
            ref NativeHelpers.LUID lpLuid,
            StringBuilder lpName,
            ref Int32 cchName);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool OpenProcessToken(
            IntPtr ProcessHandle,
            TokenAccessLevels DesiredAccess,
            out SafeNativeHandle TokenHandle);

        [DllImport("wtsapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool WTSEnumerateSessions(
            IntPtr hServer,
            int Reserved,
            int Version,
            ref IntPtr ppSessionInfo,
            ref int pCount);

        [DllImport("wtsapi32.dll")]
        public static extern void WTSFreeMemory(IntPtr pMemory);

        [DllImport("kernel32.dll")]
        public static extern uint WTSGetActiveConsoleSessionId();

        [DllImport("Wtsapi32.dll", SetLastError = true)]
        public static extern bool WTSQueryUserToken(uint SessionId, out SafeNativeHandle phToken);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr CreatePipe(
            ref IntPtr hReadPipe,
            ref IntPtr hWritePipe,
            ref NativeHelpers.SECURITY_ATTRIBUTES lpPipeAttributes,
            Int32 nSize);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool SetHandleInformation(IntPtr hObject, int dwMask, int dwFlags);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool ReadFile(
            IntPtr hFile,
            byte[] lpBuffer,
            int nNumberOfBytesToRead,
            ref int lpNumberOfBytesRead,
            IntPtr lpOverlapped);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool PeekNamedPipe(
            IntPtr handle,
            byte[] buffer,
            uint nBufferSize,
            ref uint bytesRead,
            ref uint bytesAvail,
            ref uint BytesLeftThisMessage);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DuplicateHandle(
            IntPtr hSourceProcessHandle,
            ushort hSourceHandle,
            IntPtr hTargetProcessHandle,
            out IntPtr lpTargetHandle,
            uint dwDesiredAccess,
            [MarshalAs(UnmanagedType.Bool)] bool bInheritHandle,
            uint dwOptions);
    }

    internal class SafeMemoryBuffer : SafeHandleZeroOrMinusOneIsInvalid
    {
        public SafeMemoryBuffer(int cb) : base(true) { base.SetHandle(Marshal.AllocHGlobal(cb)); }
        public SafeMemoryBuffer(IntPtr handle) : base(true) { base.SetHandle(handle); }
        protected override bool ReleaseHandle() { Marshal.FreeHGlobal(handle); return true; }
    }

    internal class SafeNativeHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        public SafeNativeHandle() : base(true) { }
        public SafeNativeHandle(IntPtr handle) : base(true) { this.handle = handle; }
        protected override bool ReleaseHandle() { return NativeMethods.CloseHandle(handle); }
    }

    internal enum SECURITY_IMPERSONATION_LEVEL
    {
        SecurityAnonymous      = 0,
        SecurityIdentification = 1,
        SecurityImpersonation  = 2,
        SecurityDelegation     = 3,
    }

    internal enum SW
    {
        SW_HIDE            = 0,
        SW_SHOWNORMAL      = 1,
        SW_NORMAL          = 1,
        SW_SHOWMINIMIZED   = 2,
        SW_SHOWMAXIMIZED   = 3,
        SW_MAXIMIZE        = 3,
        SW_SHOWNOACTIVATE  = 4,
        SW_SHOW            = 5,
        SW_MINIMIZE        = 6,
        SW_SHOWMINNOACTIVE = 7,
        SW_SHOWNA          = 8,
        SW_RESTORE         = 9,
        SW_SHOWDEFAULT     = 10,
        SW_MAX             = 10
    }

    internal enum TokenElevationType
    {
        TokenElevationTypeDefault = 1,
        TokenElevationTypeFull,
        TokenElevationTypeLimited,
    }

    internal enum TOKEN_TYPE
    {
        TokenPrimary       = 1,
        TokenImpersonation = 2
    }

    internal enum WTS_CONNECTSTATE_CLASS
    {
        WTSActive,
        WTSConnected,
        WTSConnectQuery,
        WTSShadow,
        WTSDisconnected,
        WTSIdle,
        WTSListen,
        WTSReset,
        WTSDown,
        WTSInit
    }

    [Flags]
    public enum PrivilegeAttributes : uint
    {
        Disabled         = 0x00000000,
        EnabledByDefault = 0x00000001,
        Enabled          = 0x00000002,
        Removed          = 0x00000004,
        UsedForAccess    = 0x80000000,
    }

    public class Win32Exception : System.ComponentModel.Win32Exception
    {
        private string _msg;
        public Win32Exception(string message) : this(Marshal.GetLastWin32Error(), message) { }
        public Win32Exception(int errorCode, string message) : base(errorCode)
        {
            _msg = String.Format("{0} ({1}, Win32ErrorCode {2} - 0x{2:X8})", message, base.Message, errorCode);
        }
        public override string Message { get { return _msg; } }
        public static explicit operator Win32Exception(string message) { return new Win32Exception(message); }
    }

    public static class ProcessExtensions
    {
        private const int CREATE_UNICODE_ENVIRONMENT  = 0x00000400;
        private const int CREATE_NO_WINDOW            = 0x08000000;
        private const int CREATE_NEW_CONSOLE          = 0x00000010;
        private const uint INVALID_SESSION_ID         = 0xFFFFFFFF;
        private static readonly IntPtr WTS_CURRENT_SERVER_HANDLE = IntPtr.Zero;
        private const int HANDLE_FLAG_INHERIT         = 0x00000001;
        private const int STARTF_USESTDHANDLES        = 0x00000100;
        private const int CREATE_BREAKAWAY_FROM_JOB   = 0x01000000;

        private static IntPtr out_read;
        private static IntPtr out_write;
        private static IntPtr err_read;
        private static IntPtr err_write;
        private static int BUFSIZE = 4096;

        private static SafeNativeHandle GetSessionUserToken(bool elevated)
        {
            var activeSessionId = INVALID_SESSION_ID;
            var pSessionInfo    = IntPtr.Zero;
            var sessionCount    = 0;

            if (NativeMethods.WTSEnumerateSessions(WTS_CURRENT_SERVER_HANDLE, 0, 1, ref pSessionInfo, ref sessionCount))
            {
                try
                {
                    var arrayElementSize = Marshal.SizeOf(typeof(NativeHelpers.WTS_SESSION_INFO));
                    var current          = pSessionInfo;

                    for (var i = 0; i < sessionCount; i++)
                    {
                        var si = (NativeHelpers.WTS_SESSION_INFO)Marshal.PtrToStructure(
                            current, typeof(NativeHelpers.WTS_SESSION_INFO));
                        current = IntPtr.Add(current, arrayElementSize);

                        if (si.State == WTS_CONNECTSTATE_CLASS.WTSActive)
                        {
                            activeSessionId = si.SessionID;
                            break;
                        }
                    }
                }
                finally
                {
                    NativeMethods.WTSFreeMemory(pSessionInfo);
                }
            }

            if (activeSessionId == INVALID_SESSION_ID)
                activeSessionId = NativeMethods.WTSGetActiveConsoleSessionId();

            SafeNativeHandle hImpersonationToken;
            if (!NativeMethods.WTSQueryUserToken(activeSessionId, out hImpersonationToken))
                throw new Win32Exception("WTSQueryUserToken failed to get access token.");

            using (hImpersonationToken)
            {
                TokenElevationType elevationType = GetTokenElevationType(hImpersonationToken);

                if (elevationType == TokenElevationType.TokenElevationTypeLimited && elevated == true)
                {
                    using (var linkedToken = GetTokenLinkedToken(hImpersonationToken))
                        return DuplicateTokenAsPrimary(linkedToken);
                }
                else
                {
                    return DuplicateTokenAsPrimary(hImpersonationToken);
                }
            }
        }

        public static string StartProcessAsCurrentUser(string appPath, string cmdLine = null, string workDir = null, bool visible = true, int wait = -1, bool elevated = true, bool redirectOutput = true, bool breakaway = false)
        {
            NativeHelpers.SECURITY_ATTRIBUTES saAttr = new NativeHelpers.SECURITY_ATTRIBUTES();
            saAttr.nLength              = Marshal.SizeOf(typeof(NativeHelpers.SECURITY_ATTRIBUTES));
            saAttr.bInheritHandle       = 0x1;
            saAttr.lpSecurityDescriptor = IntPtr.Zero;

            if (redirectOutput)
            {
                NativeMethods.CreatePipe(ref out_read, ref out_write, ref saAttr, 0);
                NativeMethods.CreatePipe(ref err_read, ref err_write, ref saAttr, 0);
                NativeMethods.SetHandleInformation(out_read, HANDLE_FLAG_INHERIT, 0);
                NativeMethods.SetHandleInformation(err_read, HANDLE_FLAG_INHERIT, 0);
            }

            var startInfo  = new NativeHelpers.STARTUPINFO();
            startInfo.cb   = Marshal.SizeOf(startInfo);

            uint dwCreationFlags = CREATE_UNICODE_ENVIRONMENT
                | (uint)(breakaway      ? CREATE_BREAKAWAY_FROM_JOB : 0)
                | (uint)(visible        ? CREATE_NEW_CONSOLE         : CREATE_NO_WINDOW);

            startInfo.wShowWindow = (short)(visible ? SW.SW_SHOW : SW.SW_HIDE);
            startInfo.hStdOutput  = out_write;
            startInfo.hStdError   = err_write;
            startInfo.dwFlags    |= (uint)STARTF_USESTDHANDLES;

            StringBuilder commandLine = new StringBuilder(cmdLine);
            var procInfo = new NativeHelpers.PROCESS_INFORMATION();

            using (var hUserToken = GetSessionUserToken(elevated))
            {
                IntPtr pEnv = IntPtr.Zero;

                if (!NativeMethods.CreateEnvironmentBlock(ref pEnv, hUserToken, false))
                    throw new Win32Exception("CreateEnvironmentBlock failed.");

                try
                {
                    if (!NativeMethods.CreateProcessAsUserW(hUserToken, appPath, commandLine,
                        IntPtr.Zero, IntPtr.Zero, redirectOutput, dwCreationFlags, pEnv,
                        workDir, ref startInfo, out procInfo))
                        throw new Win32Exception("CreateProcessAsUser failed.");

                    try   { NativeMethods.WaitForSingleObject(procInfo.hProcess, wait); }
                    finally
                    {
                        NativeMethods.CloseHandle(procInfo.hThread);
                        NativeMethods.CloseHandle(procInfo.hProcess);
                    }
                }
                finally { NativeMethods.DestroyEnvironmentBlock(pEnv); }
            }

            if (redirectOutput)
            {
                var sb      = new StringBuilder();
                byte[] buf  = new byte[BUFSIZE];
                int dwRead  = 0;

                while (true)
                {
                    if (Readable(out_read))
                    {
                        bool bSuccess = NativeMethods.ReadFile(out_read, buf, BUFSIZE, ref dwRead, IntPtr.Zero);
                        if (!bSuccess || dwRead == 0) break;
                        sb.AppendLine(Encoding.Default.GetString(buf).TrimEnd(new char[] { (char)0 }));
                    }
                    else { break; }
                }

                NativeMethods.CloseHandle(out_read);
                NativeMethods.CloseHandle(err_read);
                NativeMethods.CloseHandle(out_write);
                NativeMethods.CloseHandle(err_write);

                return sb.ToString();
            }
            else
            {
                return procInfo.dwProcessId.ToString();
            }
        }

        private static bool Readable(IntPtr streamHandle)
        {
            byte[] aPeekBuffer  = new byte[1];
            uint aPeekedBytes   = 0;
            uint aAvailBytes    = 0;
            uint aLeftBytes     = 0;

            bool aPeekedSuccess = NativeMethods.PeekNamedPipe(
                streamHandle, aPeekBuffer, 1,
                ref aPeekedBytes, ref aAvailBytes, ref aLeftBytes);

            return (aPeekedSuccess && aPeekBuffer[0] != 0);
        }

        private static SafeNativeHandle DuplicateTokenAsPrimary(SafeHandle hToken)
        {
            SafeNativeHandle pDupToken;
            if (!NativeMethods.DuplicateTokenEx(hToken, 0, IntPtr.Zero,
                SECURITY_IMPERSONATION_LEVEL.SecurityImpersonation,
                TOKEN_TYPE.TokenPrimary, out pDupToken))
                throw new Win32Exception("DuplicateTokenEx failed.");

            return pDupToken;
        }

        public static Dictionary<String, PrivilegeAttributes> GetTokenPrivileges()
        {
            Dictionary<string, PrivilegeAttributes> privileges = new Dictionary<string, PrivilegeAttributes>();

            using (SafeNativeHandle hToken   = OpenProcessToken(NativeMethods.GetCurrentProcess(), TokenAccessLevels.Query))
            using (SafeMemoryBuffer tokenInfo = GetTokenInformation(hToken, 3))
            {
                NativeHelpers.TOKEN_PRIVILEGES privilegeInfo =
                    (NativeHelpers.TOKEN_PRIVILEGES)Marshal.PtrToStructure(
                        tokenInfo.DangerousGetHandle(), typeof(NativeHelpers.TOKEN_PRIVILEGES));

                IntPtr ptrOffset = IntPtr.Add(tokenInfo.DangerousGetHandle(),
                    Marshal.SizeOf(privilegeInfo.PrivilegeCount));

                for (int i = 0; i < privilegeInfo.PrivilegeCount; i++)
                {
                    NativeHelpers.LUID_AND_ATTRIBUTES info =
                        (NativeHelpers.LUID_AND_ATTRIBUTES)Marshal.PtrToStructure(
                            ptrOffset, typeof(NativeHelpers.LUID_AND_ATTRIBUTES));

                    int nameLen = 0;
                    NativeHelpers.LUID privLuid = info.Luid;
                    NativeMethods.LookupPrivilegeName(null, ref privLuid, null, ref nameLen);

                    StringBuilder name = new StringBuilder(nameLen + 1);
                    if (!NativeMethods.LookupPrivilegeName(null, ref privLuid, name, ref nameLen))
                        throw new Win32Exception("LookupPrivilegeName() failed");

                    privileges[name.ToString()] = info.Attributes;
                    ptrOffset = IntPtr.Add(ptrOffset,
                        Marshal.SizeOf(typeof(NativeHelpers.LUID_AND_ATTRIBUTES)));
                }
            }

            return privileges;
        }

        private static TokenElevationType GetTokenElevationType(SafeHandle hToken)
        {
            using (SafeMemoryBuffer tokenInfo = GetTokenInformation(hToken, 18))
                return (TokenElevationType)Marshal.ReadInt32(tokenInfo.DangerousGetHandle());
        }

        private static SafeNativeHandle GetTokenLinkedToken(SafeHandle hToken)
        {
            using (SafeMemoryBuffer tokenInfo = GetTokenInformation(hToken, 19))
                return new SafeNativeHandle(Marshal.ReadIntPtr(tokenInfo.DangerousGetHandle()));
        }

        private static SafeMemoryBuffer GetTokenInformation(SafeHandle hToken, uint infoClass)
        {
            int returnLength;
            bool res = NativeMethods.GetTokenInformation(hToken, infoClass,
                new SafeMemoryBuffer(IntPtr.Zero), 0, out returnLength);
            int errCode = Marshal.GetLastWin32Error();

            if (!res && errCode != 24 && errCode != 122)
                throw new Win32Exception(errCode,
                    String.Format("GetTokenInformation({0}) failed to get buffer length", infoClass));

            SafeMemoryBuffer tokenInfo = new SafeMemoryBuffer(returnLength);
            if (!NativeMethods.GetTokenInformation(hToken, infoClass, tokenInfo, returnLength, out returnLength))
                throw new Win32Exception(String.Format("GetTokenInformation({0}) failed", infoClass));

            return tokenInfo;
        }

        private static SafeNativeHandle OpenProcessToken(IntPtr process, TokenAccessLevels access)
        {
            SafeNativeHandle hToken = null;
            if (!NativeMethods.OpenProcessToken(process, access, out hToken))
                throw new Win32Exception("OpenProcessToken() failed");

            return hToken;
        }
    }
}
"@

function Invoke-VBasCurrentUser {
    <#
    .SYNOPSIS
        Runs a scriptblock as the currently logged-on interactive user from SYSTEM context.

    .DESCRIPTION
        Invoke-VBasCurrentUser allows scripts running as SYSTEM (e.g. via Intune, RMM tools,
        or Task Scheduler) to execute code in the context of the currently logged-on user
        session. No external module dependencies -- the required C# type is compiled inline
        on first use and cached for the session.

        Useful for collecting user-specific data that is invisible to SYSTEM: applied GPOs
        (gpresult), mapped drives, printer mappings, HKCU registry values, and profile
        settings that only exist inside the user's session.

        Requires SeDelegateSessionUserImpersonatePrivilege -- only available when the
        calling process is running as SYSTEM.

    .PARAMETER ScriptBlock
        The scriptblock to execute in the context of the currently logged-on user.

    .PARAMETER NoWait
        Return immediately after launching the process. Do not wait for it to complete.

    .PARAMETER UseWindowsPowerShell
        Force use of Windows PowerShell (powershell.exe) regardless of the calling shell.

    .PARAMETER UseMicrosoftPowerShell
        Force use of PowerShell 7 (pwsh.exe). Must be installed at the standard path
        under $env:ProgramFiles\PowerShell\7\pwsh.exe.

    .PARAMETER NonElevatedSession
        Run the spawned process without elevation. Default is elevated (RunAs = true).

    .PARAMETER Visible
        Show the spawned PowerShell window. Default is hidden.

    .PARAMETER CacheToDisk
        Write the scriptblock to a temp .ps1 file before execution instead of base64-encoding
        it on the command line. Required when the scriptblock exceeds the OS command line
        length limit (32767 chars on Windows 6.2+).

    .PARAMETER CaptureOutput
        Capture and return the output from the spawned process.

    .PARAMETER Breakaway
        Launch the spawned process outside the current job object. Use when the calling
        process is inside a job that blocks child process creation.

    .PARAMETER ExpandStringVariables
        Expand variables in the scriptblock before execution. Use when you need to pass
        variable values from the SYSTEM calling scope into the user-context process.

    .EXAMPLE
        Invoke-VBasCurrentUser -ScriptBlock {
            gpresult /r /scope user | Out-File 'C:\Temp\gpo_report.txt' -Encoding UTF8
        }

        Runs gpresult as the logged-on user and saves applied GPO output to a file.
        Typical deployment: Intune or RMM script running as SYSTEM.

    .EXAMPLE
        $Username = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName -replace '.*\\'
        Invoke-VBasCurrentUser -ScriptBlock {
            gpresult /r /scope user | Out-File "C:\Temp\gpo_$Username.txt" -Encoding UTF8
        } -ExpandStringVariables

        Expands $Username from the SYSTEM calling scope into the scriptblock before launch.

    .EXAMPLE
        Invoke-VBasCurrentUser -ScriptBlock {
            Get-Printer | Select-Object Name, PortName |
                Export-Csv 'C:\Temp\printers.csv' -NoTypeInformation -Encoding UTF8
        } -CaptureOutput

        Enumerates printers visible in the user session and exports to CSV.

    .EXAMPLE
        Invoke-VBasCurrentUser -ScriptBlock {
            reg export 'HKCU\Software\Microsoft\Windows\CurrentVersion\Group Policy' 'C:\Temp\gpo_hkcu.reg' /y
        }

        Exports the user's HKCU Group Policy registry branch.

    .OUTPUTS
        None by default. When -CaptureOutput is used the output of the spawned process is
        returned as a string. On error, writes to the error stream.

    .NOTES
        Version  : 1.0.0
        Author   : Vibhu Bhatnagar
        Modified : 23-04-2026
        Category : Windows Workstation Administration

        Prerequisites:
          - Must run as SYSTEM (Intune, RMM, Task Scheduler as SYSTEM)
          - PowerShell 5.1
          - No external modules required -- C# type is compiled and cached inline

        Source: Adapted from the RunAsUser module by KelvinTegelaar (MIT Licence).
        Ref   : https://github.com/KelvinTegelaar/RunAsUser
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory = $false)]
        [switch]$NoWait,

        [Parameter(Mandatory = $false)]
        [switch]$UseWindowsPowerShell,

        [Parameter(Mandatory = $false)]
        [switch]$UseMicrosoftPowerShell,

        [Parameter(Mandatory = $false)]
        [switch]$NonElevatedSession,

        [Parameter(Mandatory = $false)]
        [switch]$Visible,

        [Parameter(Mandatory = $false)]
        [switch]$CacheToDisk,

        [Parameter(Mandatory = $false)]
        [switch]$CaptureOutput,

        [Parameter(Mandatory = $false)]
        [switch]$Breakaway,

        [Parameter(Mandatory = $false)]
        [switch]$ExpandStringVariables
    )

    # Step 1 -- Compile the C# type on first use, cached for the session
    if (-not ("RunAsUser.ProcessExtensions" -as [type])) {
        Add-Type -TypeDefinition $script:source -Language CSharp
    }

    # Step 2 -- Expand variables into scriptblock if requested
    if ($ExpandStringVariables) {
        $InnerScriptBlock = $ExecutionContext.InvokeCommand.ExpandString($ScriptBlock)
    }
    else {
        $InnerScriptBlock = $ScriptBlock
    }

    # Step 3 -- Build the PowerShell command line (disk cache or base64 encoded)
    if ($CacheToDisk) {
        $ScriptGuid  = New-Guid
        $TempScript  = Join-Path $env:TEMP "$ScriptGuid.ps1"
        $null        = New-Item -Path $TempScript -Value $InnerScriptBlock -Force
        $PwshCommand = "-ExecutionPolicy Bypass -Window Normal -file `"$TempScript`""
    }
    else {
        $EncodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($InnerScriptBlock))
        $PwshCommand    = "-ExecutionPolicy Bypass -Window Normal -EncodedCommand $EncodedCommand"
    }

    # Step 4 -- Enforce OS command line length limit
    $OsLevel   = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CurrentVersion
    $MaxLength = if ($OsLevel -lt 6.2) { 8190 } else { 32767 }

    if ($EncodedCommand.Length -gt $MaxLength -and -not $CacheToDisk) {
        Write-Error -Message "The encoded script exceeds the command line length limit. Re-run with -CacheToDisk."
        return
    }

    # Step 5 -- Validate pwsh.exe path when -UseMicrosoftPowerShell is requested
    $Pwsh7Path = Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
    if ($UseMicrosoftPowerShell -and -not (Test-Path -Path $Pwsh7Path)) {
        Write-Error -Message "PowerShell 7 (pwsh.exe) not found at $Pwsh7Path. Ensure it is installed."
        return
    }

    # Step 6 -- Verify the calling process has the required impersonation privilege
    $Privs = [RunAsUser.ProcessExtensions]::GetTokenPrivileges()['SeDelegateSessionUserImpersonatePrivilege']
    if (-not $Privs -or ($Privs -band [RunAsUser.PrivilegeAttributes]::Disabled)) {
        Write-Error -Message "Missing SeDelegateSessionUserImpersonatePrivilege. This function must run as SYSTEM."
        return
    }

    # Step 7 -- Launch the process as the current interactive user
    try {
        $WinPsPath = Join-Path $env:SystemRoot 'system32\WindowsPowerShell\v1.0\powershell.exe'

        $PwshPath = if ($UseWindowsPowerShell)     { $WinPsPath }
                    elseif ($UseMicrosoftPowerShell) { $Pwsh7Path }
                    else                             { (Get-Process -Id $PID).Path }

        $ProcWaitTime = if ($NoWait) { 1 } else { -1 }
        $RunAsAdmin   = -not $NonElevatedSession

        [RunAsUser.ProcessExtensions]::StartProcessAsCurrentUser(
            $PwshPath,
            "`"$PwshPath`" $PwshCommand",
            (Split-Path $PwshPath -Parent),
            $Visible,
            $ProcWaitTime,
            $RunAsAdmin,
            $CaptureOutput,
            $Breakaway
        )

        if ($CacheToDisk) { $null = Remove-Item -Path $TempScript -Force }
    }
    catch {
        Write-Error -Message "Could not execute as currently logged-on user: $($_.Exception.Message)" -Exception $_.Exception
        return
    }
}
