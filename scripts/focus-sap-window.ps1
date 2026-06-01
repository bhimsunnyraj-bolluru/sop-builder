# Brings the SAP GUI main window to the foreground before capture.
$ErrorActionPreference = 'SilentlyContinue'

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;

public class SapFocusWin32 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern IntPtr GetParent(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    public static string ReadClassName(IntPtr hWnd) {
        if (hWnd == IntPtr.Zero) return "";
        var sb = new StringBuilder(256);
        GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    public static IntPtr FindSapRoot(IntPtr hWnd) {
        var current = hWnd;
        while (current != IntPtr.Zero) {
            if (ReadClassName(current) == "SAP_FRONTEND_SESSION") return current;
            current = GetParent(current);
        }
        return IntPtr.Zero;
    }

    public static List<IntPtr> EnumSapWindows() {
        var list = new List<IntPtr>();
        EnumWindows((hWnd, lParam) => {
            if (!IsWindowVisible(hWnd)) return true;
            if (ReadClassName(hWnd) == "SAP_FRONTEND_SESSION") list.Add(hWnd);
            return true;
        }, IntPtr.Zero);
        return list;
    }
}
"@ | Out-Null

function Get-SapRootWindow {
    $fg = [SapFocusWin32]::GetForegroundWindow()
    $sapRoot = [SapFocusWin32]::FindSapRoot($fg)
    if ($sapRoot -ne [IntPtr]::Zero) { return $sapRoot }
    $windows = [SapFocusWin32]::EnumSapWindows()
    if ($windows.Count -gt 0) { return $windows[0] }
    return [IntPtr]::Zero
}

$sapRoot = Get-SapRootWindow
if ($sapRoot -eq [IntPtr]::Zero) {
    Write-Output '{"ok":false,"reason":"no-sap-window"}'
    exit 0
}

[void][SapFocusWin32]::SetForegroundWindow($sapRoot)
Start-Sleep -Milliseconds 150
Write-Output '{"ok":true}'
