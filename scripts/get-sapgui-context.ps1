# Returns JSON with SAP GUI window / session context for SOP step descriptions.
# Strategies (in order): SAP GUI Scripting COM, Win32 foreground + parent walk, SAP window enum.
$ErrorActionPreference = 'Stop'

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;

public class SapWin32 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern IntPtr GetParent(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    public static string ReadWindowText(IntPtr hWnd) {
        if (hWnd == IntPtr.Zero) return "";
        var sb = new StringBuilder(2048);
        GetWindowText(hWnd, sb, sb.Capacity);
        var text = sb.ToString();
        if (!string.IsNullOrWhiteSpace(text)) return text.Trim();
        return "";
    }

    public static string ReadClassName(IntPtr hWnd) {
        if (hWnd == IntPtr.Zero) return "";
        var sb = new StringBuilder(256);
        GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    public static IntPtr FindSapRoot(IntPtr hWnd) {
        var current = hWnd;
        while (current != IntPtr.Zero) {
            var cls = ReadClassName(current);
            if (cls == "SAP_FRONTEND_SESSION") { return current; }
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
"@ -ErrorAction SilentlyContinue

function Invoke-ComMethod {
    param($Object, [string]$Method, [object[]]$Args = @())
    if ($null -eq $Object) { return $null }
    return $Object.GetType().InvokeMember($Method, [System.Reflection.BindingFlags]::InvokeMethod, $null, $Object, $Args)
}

function Get-ComProperty {
    param($Object, [string]$Property, [object[]]$Args = @())
    if ($null -eq $Object) { return $null }
    return $Object.GetType().InvokeMember($Property, [System.Reflection.BindingFlags]::GetProperty, $null, $Object, $Args)
}

function New-Result {
    param(
        [string]$Title = "",
        [string]$Transaction = "",
        [string]$Program = "",
        [string]$ScreenNumber = "",
        [string]$SystemName = "",
        [string]$StatusBar = "",
        [string]$WindowClass = "",
        [string]$Source = "",
        [string]$SapWindowTitle = ""
    )
    [ordered]@{
        title         = $Title
        transaction   = $Transaction
        program       = $Program
        screenNumber  = $ScreenNumber
        systemName    = $SystemName
        statusBar     = $StatusBar
        windowClass   = $WindowClass
        sapWindowTitle = $SapWindowTitle
        source        = $Source
    }
}

function Build-Title {
    param($Result)
    if ($Result.title) { return $Result.title }
    $parts = @()
    if ($Result.sapWindowTitle) { $parts += $Result.sapWindowTitle }
    if ($Result.transaction) { $parts += $Result.transaction }
    if ($Result.program) { $parts += "($($Result.program))" }
    if ($Result.statusBar) { $parts += "- $($Result.statusBar)" }
    return ($parts -join " ")
}

# --- Win32: foreground SAP window ---
$fg = [SapWin32]::GetForegroundWindow()
$fgClass = [SapWin32]::ReadClassName($fg)
$fgTitle = [SapWin32]::ReadWindowText($fg)
$sapRoot = [SapWin32]::FindSapRoot($fg)
$sapRootTitle = ""
$sapRootClass = ""
if ($sapRoot -ne [IntPtr]::Zero) {
    $sapRootTitle = [SapWin32]::ReadWindowText($sapRoot)
    $sapRootClass = [SapWin32]::ReadClassName($sapRoot)
}

$win32Title = if ($sapRootTitle) { $sapRootTitle } elseif ($fgTitle) { $fgTitle } else { "" }
$win32Class = if ($sapRootClass) { $sapRootClass } else { $fgClass }

# If foreground is not SAP, prefer a visible SAP window that still looks active (often only one)
if (-not $sapRootTitle) {
    $sapWindows = [SapWin32]::EnumSapWindows()
    foreach ($h in $sapWindows) {
        $t = [SapWin32]::ReadWindowText($h)
        if ($t -and $t -ne "SAP") {
            $win32Title = $t
            $win32Class = "SAP_FRONTEND_SESSION"
            break
        }
    }
    if (-not $win32Title -and $sapWindows.Count -gt 0) {
        $win32Title = [SapWin32]::ReadWindowText($sapWindows[0])
        $win32Class = "SAP_FRONTEND_SESSION"
    }
}

$result = New-Result -Title $win32Title -WindowClass $win32Class -SapWindowTitle $win32Title -Source "win32"

# --- SAP GUI Scripting COM (best metadata when enabled) ---
try {
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
    $sapGuiAuto = $null
    try {
        $sapGuiAuto = [System.Runtime.InteropServices.Marshal]::GetActiveObject("SAPGUI")
    } catch {
        $sapGuiAuto = [Microsoft.VisualBasic.Interaction]::GetObject("SAPGUI")
    }

    if ($null -ne $sapGuiAuto) {
        $application = Invoke-ComMethod $sapGuiAuto "GetScriptingEngine"
        if ($null -ne $application) {
            $connections = Get-ComProperty $application "Children"
            $connCount = [int](Get-ComProperty $connections "Count")
            $bestSession = $null
            $bestScore = -1

            for ($ci = 0; $ci -lt $connCount; $ci++) {
                $connection = Get-ComProperty $connections "Item" @($ci)
                if ($null -eq $connection) { continue }
                $sessions = Get-ComProperty $connection "Children"
                $sessCount = [int](Get-ComProperty $sessions "Count")

                for ($si = 0; $si -lt $sessCount; $si++) {
                    $session = Get-ComProperty $sessions "Item" @($si)
                    if ($null -eq $session) { continue }

                    $busy = $false
                    try { $busy = [bool](Invoke-ComMethod $session "Busy") } catch { $busy = $false }
                    if ($busy) { continue }

                    $info = Get-ComProperty $session "Info"
                    $txn = ""
                    $prog = ""
                    $screen = ""
                    $sysName = ""
                    try { $txn = [string](Get-ComProperty $info "Transaction") } catch {}
                    try { $prog = [string](Get-ComProperty $info "Program") } catch {}
                    try { $screen = [string](Get-ComProperty $info "ScreenNumber") } catch {}
                    try { $sysName = [string](Get-ComProperty $info "SystemName") } catch {}

                    $wndText = ""
                    $sbar = ""
                    try {
                        $wnd = Invoke-ComMethod $session "findById" @("wnd[0]")
                        if ($wnd) { $wndText = [string](Get-ComProperty $wnd "Text") }
                    } catch {}
                    try {
                        $sb = Invoke-ComMethod $session "findById" @("wnd[0]/sbar")
                        if ($sb) { $sbar = [string](Get-ComProperty $sb "Text") }
                    } catch {}

                    $score = 0
                    if ($wndText) { $score += 2 }
                    if ($txn) { $score += 2 }
                    if ($win32Title -and $wndText -and ($wndText -eq $win32Title -or $win32Title.Contains($wndText) -or $wndText.Contains($win32Title))) {
                        $score += 10
                    }
                    if ($win32Title -and $txn -and $win32Title.ToUpper().Contains($txn.ToUpper())) {
                        $score += 5
                    }

                    if ($score -gt $bestScore) {
                        $bestScore = $score
                        $bestSession = [ordered]@{
                            transaction  = $txn
                            program      = $prog
                            screenNumber = $screen
                            systemName   = $sysName
                            sapWindowTitle = $wndText
                            statusBar    = $sbar
                        }
                    }
                }
            }

            if ($null -ne $bestSession) {
                $scriptTitle = $bestSession.sapWindowTitle
                if (-not $scriptTitle) {
                    $scriptTitle = if ($bestSession.transaction) {
                        "$($bestSession.transaction)"
                    } else { $win32Title }
                }

                $composed = $scriptTitle
                if ($bestSession.transaction -and $scriptTitle -notmatch [regex]::Escape($bestSession.transaction)) {
                    $composed = "$scriptTitle - $($bestSession.transaction)"
                }
                if ($bestSession.systemName) {
                    $composed = "$composed [$($bestSession.systemName)]"
                }

                $result = New-Result `
                    -Title $composed `
                    -Transaction $bestSession.transaction `
                    -Program $bestSession.program `
                    -ScreenNumber $bestSession.screenNumber `
                    -SystemName $bestSession.systemName `
                    -StatusBar $bestSession.statusBar `
                    -WindowClass $win32Class `
                    -SapWindowTitle $bestSession.sapWindowTitle `
                    -Source "sap-scripting"
            }
        }
    }
} catch {
    # Scripting disabled or SAP not running — keep win32 result
}

if (-not $result.title) {
    $result.title = Build-Title $result
}

$result | ConvertTo-Json -Compress
