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

    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public static bool TryGetWindowRect(IntPtr hWnd, out RECT rect) {
        rect = new RECT();
        if (hWnd == IntPtr.Zero) return false;
        return GetWindowRect(hWnd, out rect);
    }

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

    public class TextCandidate {
        public string Text;
        public int Top;
        public int Bottom;
        public string ClassName;
    }

    public static void CollectVisibleTexts(IntPtr hWnd, System.Collections.Generic.List<TextCandidate> items) {
        if (hWnd == IntPtr.Zero) return;
        EnumChildWindows(hWnd, (child, lParam) => {
            if (!IsWindowVisible(child)) return true;
            var text = ReadWindowText(child);
            var cls = ReadClassName(child);
            RECT r;
            if (!string.IsNullOrWhiteSpace(text) && TryGetWindowRect(child, out r)) {
                items.Add(new TextCandidate {
                    Text = text.Trim(),
                    Top = r.Top,
                    Bottom = r.Bottom,
                    ClassName = cls
                });
            }
            CollectVisibleTexts(child, items);
            return true;
        }, IntPtr.Zero);
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

function Get-SessionContext {
    param($session)
    if ($null -eq $session) { return $null }

    $busy = $false
    try { $busy = [bool](Invoke-ComMethod $session "Busy") } catch { $busy = $false }
    # Do not skip Busy sessions — after Save (e.g. VA01) the order number appears
    # on the status bar while the session may still report Busy.

    $info = $null
    try { $info = Get-ComProperty $session "Info" } catch {}
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
        if ($sb) {
            $sbar = [string](Get-ComProperty $sb "Text")
            if (-not $sbar) {
                try { $sbar = [string](Get-ComProperty $sb "MessageText") } catch {}
            }
        }
    } catch {}

    if ($busy -and -not $sbar -and -not $wndText -and -not $txn) { return $null }

    return [ordered]@{
        transaction    = $txn
        program        = $prog
        screenNumber   = $screen
        systemName     = $sysName
        sapWindowTitle = $wndText
        statusBar      = $sbar
    }
}

function Apply-SessionContext {
    param($Result, $SessionContext, [string]$Source, [string]$WindowClass)
    if ($null -eq $SessionContext) { return $Result }
    return New-Result `
        -Transaction $SessionContext.transaction `
        -Program $SessionContext.program `
        -ScreenNumber $SessionContext.screenNumber `
        -SystemName $SessionContext.systemName `
        -StatusBar $SessionContext.statusBar `
        -WindowClass $WindowClass `
        -SapWindowTitle $(if ($SessionContext.sapWindowTitle) { $SessionContext.sapWindowTitle } else { $Result.sapWindowTitle }) `
        -Source $Source
}

function Score-SessionContext {
    param($SessionContext, [string]$Win32Title)
    if ($null -eq $SessionContext) { return -1 }
    $score = 0
    if ($SessionContext.sapWindowTitle) { $score += 2 }
    if ($SessionContext.transaction) { $score += 2 }
    if ($SessionContext.statusBar) { $score += 1 }
    if ($Win32Title -and $SessionContext.sapWindowTitle -and (
        $SessionContext.sapWindowTitle -eq $Win32Title -or
        $Win32Title.Contains($SessionContext.sapWindowTitle) -or
        $SessionContext.sapWindowTitle.Contains($Win32Title)
    )) { $score += 10 }
    if ($Win32Title -and $SessionContext.transaction -and $Win32Title.ToUpper().Contains($SessionContext.transaction.ToUpper())) {
        $score += 5
    }
    return $score
}

function Get-SapStatusBarWin32 {
    param([IntPtr]$SapRoot)
    if ($SapRoot -eq [IntPtr]::Zero) { return "" }

    $rect = New-Object SapWin32+RECT
    if (-not [SapWin32]::TryGetWindowRect($SapRoot, [ref]$rect)) { return "" }

    $items = New-Object 'System.Collections.Generic.List[SapWin32+TextCandidate]'
    [SapWin32]::CollectVisibleTexts($SapRoot, $items)

    $windowHeight = [Math]::Max(1, $rect.Bottom - $rect.Top)
    $thresholdTop = $rect.Top + [int]($windowHeight * 0.84)
    $best = ""
    $bestScore = -1

    foreach ($item in $items) {
        if ($item.Top -lt $thresholdTop) { continue }
        $text = $item.Text.Trim()
        if ($text.Length -lt 4) { continue }
        if ($text -eq $rect.ToString()) { continue }

        $score = $text.Length
        if ($item.ClassName -match 'status|sbar|pane') { $score += 40 }
        if ($text -match '(?i)(saved|created|changed|order|document|error|warning|success|standard|invoice|delivery|material|message)') {
            $score += 30
        }
        if ($text -match '\d{3,}') { $score += 15 }

        if ($score -gt $bestScore) {
            $bestScore = $score
            $best = $text
        }
    }

    return $best
}

function Get-SapStatusBarUia {
    param([IntPtr]$SapRoot)
    if ($SapRoot -eq [IntPtr]::Zero) { return "" }

    try {
        Add-Type -AssemblyName UIAutomationClient
        Add-Type -AssemblyName UIAutomationTypes
    } catch {
        return ""
    }

    try {
        $root = [System.Windows.Automation.AutomationElement]::FromHandle($SapRoot)
        if ($null -eq $root) { return "" }

        $statusType = [System.Windows.Automation.ControlType]::StatusBar
        $condStatus = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            $statusType
        )
        $statusBars = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condStatus)
        $best = ""
        foreach ($bar in $statusBars) {
            $name = [string]$bar.Current.Name
            if ($name.Length -gt $best.Length) { $best = $name.Trim() }
        }
        if ($best) { return $best }

        $paneType = [System.Windows.Automation.ControlType]::Pane
        $condPane = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            $paneType
        )
        $panes = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condPane)
        $rect = $root.Current.BoundingRectangle
        $threshold = $rect.Bottom - ($rect.Height * 0.12)
        foreach ($pane in $panes) {
            $r = $pane.Current.BoundingRectangle
            if ($r.Bottom -lt $threshold) { continue }
            $name = [string]$pane.Current.Name
            if ($name.Length -gt 12 -and $name.Length -gt $best.Length) { $best = $name.Trim() }
        }
        return $best
    } catch {
        return ""
    }
}

function Get-SapWindowRectJson {
    param([IntPtr]$SapRoot)
    if ($SapRoot -eq [IntPtr]::Zero) { return $null }
    $rect = New-Object SapWin32+RECT
    if (-not [SapWin32]::TryGetWindowRect($SapRoot, [ref]$rect)) { return $null }
    return @{
        left   = $rect.Left
        top    = $rect.Top
        right  = $rect.Right
        bottom = $rect.Bottom
    }
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

$result = New-Result -WindowClass $win32Class -SapWindowTitle $win32Title -Source "win32"

$sapRootForRect = $sapRoot
if ($sapRootForRect -eq [IntPtr]::Zero -and $win32Class -eq "SAP_FRONTEND_SESSION") {
    $sapWindows = [SapWin32]::EnumSapWindows()
    if ($sapWindows.Count -gt 0) { $sapRootForRect = $sapWindows[0] }
}

$win32Status = Get-SapStatusBarWin32 $sapRootForRect
if (-not $win32Status) {
    $win32Status = Get-SapStatusBarUia $sapRootForRect
    if ($win32Status) { $result.source = "win32+uia" }
} else {
    $result.source = "win32+child"
}
if ($win32Status) {
    $result.statusBar = $win32Status
}

$sapWindowRect = Get-SapWindowRectJson $sapRootForRect

# --- SAP GUI Scripting COM (transaction, status bar, system name) ---
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
            $scriptContext = $null

            # Prefer SAP's active session when available
            try {
                $activeSession = Get-ComProperty $application "ActiveSession"
                $scriptContext = Get-SessionContext $activeSession
                if ($null -ne $scriptContext) {
                    $result = Apply-SessionContext $result $scriptContext "sap-scripting-active" $win32Class
                }
            } catch {}

            # Otherwise pick the best matching open session
            if ($result.source -eq "win32") {
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
                        $ctx = Get-SessionContext $session
                        $score = Score-SessionContext $ctx $win32Title
                        if ($score -gt $bestScore) {
                            $bestScore = $score
                            $bestSession = $ctx
                        }
                    }
                }

                if ($null -ne $bestSession) {
                    $result = Apply-SessionContext $result $bestSession "sap-scripting" $win32Class
                }
            }

            # Scripting connected but status bar empty — keep Win32/UIA status if any
            if (-not $result.statusBar -and $win32Status) {
                $result.statusBar = $win32Status
            }
        }
    }
} catch {
    # Scripting disabled or SAP not running — keep win32 result
}

if (-not $result.sapWindowTitle -and $win32Title) {
    $result.sapWindowTitle = $win32Title
}

$output = [ordered]@{}
foreach ($key in $result.Keys) { $output[$key] = $result[$key] }
if ($null -ne $sapWindowRect) { $output.sapWindowRect = $sapWindowRect }

$output | ConvertTo-Json -Compress
