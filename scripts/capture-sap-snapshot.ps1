# Captures full SAP GUI screen state via SAP GUI Scripting (wnd[0]/usr traversal).
# Falls back to Win32/UI Automation when scripting is disabled.
param(
    [switch]$UiaOnly
)
$ErrorActionPreference = 'Continue'

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;

public class SapSnapWin32 {
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
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public class UiNode {
        public string Text;
        public string ClassName;
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
        public string Source;
        public string ControlType;
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
        var t = sb.ToString();
        return string.IsNullOrWhiteSpace(t) ? "" : t.Trim();
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

    public static void CollectUiNodes(IntPtr hWnd, List<UiNode> items) {
        if (hWnd == IntPtr.Zero) return;
        EnumChildWindows(hWnd, (child, lParam) => {
            if (!IsWindowVisible(child)) return true;
            var text = ReadWindowText(child);
            var cls = ReadClassName(child);
            RECT r;
            if (!string.IsNullOrWhiteSpace(text) && TryGetWindowRect(child, out r)) {
                items.Add(new UiNode {
                    Text = text.Trim(),
                    ClassName = cls,
                    Left = r.Left,
                    Top = r.Top,
                    Right = r.Right,
                    Bottom = r.Bottom,
                    Source = "win32",
                    ControlType = cls
                });
            }
            CollectUiNodes(child, items);
            return true;
        }, IntPtr.Zero);
    }
}
"@ -ErrorAction SilentlyContinue

function Invoke-ComMethod {
    param($Object, [string]$Method, [object[]]$ComArgs = @())
    if ($null -eq $Object) { return $null }
    try {
        return $Object.GetType().InvokeMember($Method, [System.Reflection.BindingFlags]::InvokeMethod, $null, $Object, $ComArgs)
    } catch {
        return $null
    }
}

function Get-ComProperty {
    param($Object, [string]$Property, [object[]]$ComArgs = @())
    if ($null -eq $Object) { return $null }
    try {
        return $Object.GetType().InvokeMember($Property, [System.Reflection.BindingFlags]::GetProperty, $null, $Object, $ComArgs)
    } catch {
        return $null
    }
}

function Get-ComPropertySafe {
    param($Object, [string]$Property)
    try {
        $v = Get-ComProperty $Object $Property
        if ($null -eq $v) { return "" }
        return [string]$v
    } catch {
        return ""
    }
}

function Get-ComIntProperty {
    param($Object, [string]$Property)
    foreach ($name in @($Property)) {
        try {
            return [int](Get-ComProperty $Object $name)
        } catch {}
    }
    return 0
}

function Get-SapScriptingRegistryDiag {
    $diag = @{}
    foreach ($regPath in @(
        "HKCU:\Software\SAP\SAPGUI Front\SAP Frontend Server\Security",
        "HKCU:\Software\SAP\SAPGUI Front\SAP Frontend Server\Scripting"
    )) {
        try {
            if (-not (Test-Path $regPath)) { continue }
            $vals = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
            if ($null -eq $vals) { continue }
            if ($null -ne $vals.UserScripting) { $diag.UserScripting = [int]$vals.UserScripting }
            if ($null -ne $vals.ScriptingSecurityLevel) { $diag.ScriptingSecurityLevel = [int]$vals.ScriptingSecurityLevel }
            if ($null -ne $vals.WarnOnAttach) { $diag.WarnOnAttach = [int]$vals.WarnOnAttach }
        } catch {}
    }
    return $diag
}

function Get-SapGuiAutoObject {
    param([ref]$MethodUsed)
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue

    # Preferred: SAP ROT Wrapper (works when direct GetObject fails)
    foreach ($progId in @("SapROTWr.SapROTWrapper", "SAPROTWr.SapROTWrapper")) {
        try {
            $wrapper = New-Object -ComObject $progId
            if ($null -eq $wrapper) { continue }
            $entry = Invoke-ComMethod $wrapper "GetROTEntry" @("SAPGUI")
            if ($null -ne $entry) {
                if ($null -ne $MethodUsed) { $MethodUsed.Value = "SapROTWrapper" }
                return $entry
            }
        } catch {}
    }

    try {
        $obj = [System.Runtime.InteropServices.Marshal]::GetActiveObject("SAPGUI")
        if ($null -ne $obj) {
            if ($null -ne $MethodUsed) { $MethodUsed.Value = "GetActiveObject" }
            return $obj
        }
    } catch {}

    try {
        $obj = [Microsoft.VisualBasic.Interaction]::GetObject("SAPGUI")
        if ($null -ne $obj) {
            if ($null -ne $MethodUsed) { $MethodUsed.Value = "GetObject" }
            return $obj
        }
    } catch {}

    try {
        $obj = New-Object -ComObject "SAPGUI"
        if ($null -ne $obj) {
            if ($null -ne $MethodUsed) { $MethodUsed.Value = "CreateObject" }
            return $obj
        }
    } catch {}

    return $null
}

function Get-SapGuiApplication {
    param([ref]$Diagnostic)
    $diag = @{
        sapGuiAuto = $false
        scriptingEngine = $false
        error = ""
        connectMethod = ""
        registry = (Get-SapScriptingRegistryDiag)
        powershellBitness = [IntPtr]::Size * 8
    }
    try {
        $method = ""
        $sapGuiAuto = Get-SapGuiAutoObject ([ref]$method)
        if ($null -eq $sapGuiAuto) {
            $diag.error = "SAPGUI COM object not found (open SAP Logon and a session, then capture again)"
            if ($null -ne $Diagnostic) { $Diagnostic.Value = $diag }
            return $null
        }
        $diag.sapGuiAuto = $true
        $diag.connectMethod = $method

        $application = $null
        try {
            $application = $sapGuiAuto.GetScriptingEngine()
        } catch {}
        if ($null -eq $application) {
            $application = Invoke-ComMethod $sapGuiAuto "GetScriptingEngine"
        }
        if ($null -eq $application) {
            $sec = $diag.registry.ScriptingSecurityLevel
            $hint = 'GetScriptingEngine returned null.'
            if ($diag.registry.UserScripting -eq 1) {
                $hint += ' Client scripting is enabled but the SAP server may still block scripting. Ask Basis to set RZ11 sapgui/user_scripting = TRUE and log off and back in.'
            } elseif ($sec -eq 2) {
                $hint += ' Scripting security is Disabled in registry.'
            } elseif ($sec -eq 1) {
                $hint += ' Scripting security is Prompt - click Allow when SAP asks.'
            } else {
                $hint += ' Check SAP GUI Scripting security Allow, server RZ11 sapgui/user_scripting=TRUE, and run capture while SAP session is open.'
            }
            if ([IntPtr]::Size -eq 8) {
                $hint += ' If this persists, use 32-bit SAP GUI with 32-bit PowerShell.'
            }
            $diag.error = $hint
            if ($null -ne $Diagnostic) { $Diagnostic.Value = $diag }
            return $null
        }
        $diag.scriptingEngine = $true
        if ($null -ne $Diagnostic) { $Diagnostic.Value = $diag }
        return $application
    } catch {
        $diag.error = $_.Exception.Message
        if ($null -ne $Diagnostic) { $Diagnostic.Value = $diag }
        return $null
    }
}

function Get-ForegroundSapTitle {
    $fg = [SapSnapWin32]::GetForegroundWindow()
    $sapRoot = [SapSnapWin32]::FindSapRoot($fg)
    if ($sapRoot -ne [IntPtr]::Zero) {
        $title = [SapSnapWin32]::ReadWindowText($sapRoot)
        if ($title) { return $title }
    }
    foreach ($h in [SapSnapWin32]::EnumSapWindows()) {
        $title = [SapSnapWin32]::ReadWindowText($h)
        if ($title -and $title -ne "SAP") { return $title }
    }
    return ""
}

function Focus-SapWindow {
    $sapRoot = Get-SapRootWindow
    if ($sapRoot -eq [IntPtr]::Zero) { return $false }
    try {
        # Foreground only - do NOT call ShowWindow (it restores/shrinks maximized SAP windows).
        [void][SapSnapWin32]::SetForegroundWindow($sapRoot)
        Start-Sleep -Milliseconds 200
        return $true
    } catch {
        return $false
    }
}

function Get-SapWindowRectJson {
    $sapRoot = Get-SapRootWindow
    if ($sapRoot -eq [IntPtr]::Zero) { return $null }
    $rect = New-Object SapSnapWin32+RECT
    if (-not [SapSnapWin32]::TryGetWindowRect($sapRoot, [ref]$rect)) { return $null }
    return @{
        left   = [int]$rect.Left
        top    = [int]$rect.Top
        right  = [int]$rect.Right
        bottom = [int]$rect.Bottom
    }
}

function Get-SessionContext {
    param($session)
    if ($null -eq $session) { return $null }

    $info = $null
    try { $info = Get-ComProperty $session "Info" } catch {}

    $txn = ""
    $prog = ""
    $screen = ""
    $sysName = ""
    $wndText = ""
    $sbar = ""
    if ($null -ne $info) {
        $txn = Get-ComPropertySafe $info "Transaction"
        $prog = Get-ComPropertySafe $info "Program"
        $screen = Get-ComPropertySafe $info "ScreenNumber"
        $sysName = Get-ComPropertySafe $info "SystemName"
    }
    try {
        $wnd = Invoke-ComMethod $session "findById" @("wnd[0]")
        if ($wnd) { $wndText = Get-ComPropertySafe $wnd "Text" }
    } catch {}
    try {
        $sb = Invoke-ComMethod $session "findById" @("wnd[0]/sbar")
        if ($sb) {
            $sbar = Get-ComPropertySafe $sb "Text"
            if (-not $sbar) { $sbar = Get-ComPropertySafe $sb "MessageText" }
        }
    } catch {}

    return [ordered]@{
        transaction    = $txn
        program        = $prog
        screenNumber   = $screen
        systemName     = $sysName
        sapWindowTitle = $wndText
        statusBar      = $sbar
    }
}

function Score-SessionContext {
    param($SessionContext, [string]$Win32Title)
    if ($null -eq $SessionContext) { return -1 }
    $score = 0
    if ($SessionContext.sapWindowTitle) { $score += 2 }
    if ($SessionContext.transaction) { $score += 2 }
    if ($SessionContext.statusBar) { $score += 1 }
    if ($Win32Title -and $SessionContext.sapWindowTitle) {
        $a = $Win32Title.ToUpper()
        $b = [string]$SessionContext.sapWindowTitle
        if ($a -eq $b.ToUpper() -or $a.Contains($b.ToUpper()) -or $b.ToUpper().Contains($a)) {
            $score += 10
        }
    }
    if ($Win32Title -and $SessionContext.transaction) {
        if ($Win32Title.ToUpper().Contains([string]$SessionContext.transaction.ToUpper())) {
            $score += 5
        }
    }
    return $score
}

function Get-BestSapSession {
    param([string]$Win32Title = "")
    $scriptDiag = @{}
    $application = Get-SapGuiApplication ([ref]$scriptDiag)
    if ($null -eq $application) { return @{ session = $null; scriptDiag = $scriptDiag } }

    $bestSession = $null
    $bestScore = -1

    try {
        $active = Get-ComProperty $application "ActiveSession"
        if ($null -ne $active) {
            $ctx = Get-SessionContext $active
            $score = Score-SessionContext $ctx $Win32Title
            if ($score -gt $bestScore) {
                $bestScore = $score
                $bestSession = $active
            }
        }
    } catch {}

    try {
        $connections = Get-ComProperty $application "Children"
        if ($null -ne $connections) {
            $connCount = [int](Get-ComProperty $connections "Count")
            for ($ci = 0; $ci -lt $connCount; $ci++) {
                $connection = Get-ComProperty $connections "Item" @($ci)
                if ($null -eq $connection) { continue }
                $sessions = Get-ComProperty $connection "Children"
                if ($null -eq $sessions) { continue }
                $sessCount = [int](Get-ComProperty $sessions "Count")
                for ($si = 0; $si -lt $sessCount; $si++) {
                    $session = Get-ComProperty $sessions "Item" @($si)
                    if ($null -eq $session) { continue }
                    $ctx = Get-SessionContext $session
                    $score = Score-SessionContext $ctx $Win32Title
                    if ($score -gt $bestScore) {
                        $bestScore = $score
                        $bestSession = $session
                    }
                }
            }
        }
    } catch {}

    if ($bestScore -lt 0) { return @{ session = $null; scriptDiag = $scriptDiag } }
    return @{ session = $bestSession; scriptDiag = $scriptDiag }
}

function Get-ActiveSapSession {
    param([string]$Win32Title = "")
    $result = Get-BestSapSession $Win32Title
    return $result.session
}

function Read-FieldValue {
    param($control, [string]$Type)
    $candidates = @("Text", "Key", "Value")
    if ($Type -match '(?i)ComboBox|ListBox|DropDown') {
        $candidates = @("Key", "Text", "Value")
    }
    if ($Type -match '(?i)CTextField|TextField') {
        $candidates = @("Text", "Key", "Value", "SelectedText")
    }
    foreach ($prop in $candidates) {
        $v = Get-ComPropertySafe $control $prop
        if (-not [string]::IsNullOrWhiteSpace($v)) { return $v.Trim() }
    }
    return ""
}

function Resolve-ControlName {
    param([string]$Name, [string]$Id)
    if (-not [string]::IsNullOrWhiteSpace($Name)) { return $Name.Trim() }
    if ($Id -match '(?i)(?:ctxt|txt|cmb|chk|btn|rdo|pwd|sub)([^/]+)$') {
        return $Matches[1]
    }
    return $Id.Trim()
}

function Read-ControlNode {
    param($control)
    if ($null -eq $control) { return $null }

    $type = Get-ComPropertySafe $control "Type"
    $id = Get-ComPropertySafe $control "Id"
    $name = Resolve-ControlName (Get-ComPropertySafe $control "Name") $id
    $text = Read-FieldValue $control $type
    if (-not $text -and $type -match '(?i)Label') {
        $text = Get-ComPropertySafe $control "Text"
    }

    $left = Get-ComIntProperty $control "ScreenLeft"
    $top = Get-ComIntProperty $control "ScreenTop"
    $width = Get-ComIntProperty $control "ScreenWidth"
    $height = Get-ComIntProperty $control "ScreenHeight"
    if ($width -le 0) { $width = Get-ComIntProperty $control "Width" }
    if ($height -le 0) { $height = Get-ComIntProperty $control "Height" }
    if ($left -eq 0 -and $top -eq 0) {
        $left = Get-ComIntProperty $control "Left"
        $top = Get-ComIntProperty $control "Top"
    }

    return [ordered]@{
        id      = $id
        type    = $type
        name    = $name
        text    = $text
        value   = $text
        tooltip = Get-ComPropertySafe $control "Tooltip"
        left    = $left
        top     = $top
        width   = $width
        height  = $height
    }
}

function Test-HasChildren {
    param($control)
    try {
        $children = Get-ComProperty $control "Children"
        $count = [int](Get-ComProperty $children "Count")
        return ($count -gt 0)
    } catch {
        return $false
    }
}

function Test-IsEmptyControl {
    param($node)
    if ($null -eq $node) { return $true }
    $hasIdentity = -not [string]::IsNullOrWhiteSpace($node.name) -or -not [string]::IsNullOrWhiteSpace($node.id)
    $hasContent = -not [string]::IsNullOrWhiteSpace($node.text) -or -not [string]::IsNullOrWhiteSpace($node.value)
    if ($hasContent) { return $false }
    if ($node.type -match '(?i)Label|TextField|CTextField|ComboBox|CheckBox|RadioButton|Button|OkCodeField') {
        return -not $hasIdentity
    }
    return $true
}

function Invoke-SapControlWalker {
    param($control, [ref]$Accum)
    if ($null -eq $control) { return }

    $node = Read-ControlNode $control
    if ($null -ne $node -and -not (Test-IsEmptyControl $node)) {
        $Accum.Value += ,$node
    }

    if (Test-HasChildren $control) {
        try {
            $children = Get-ComProperty $control "Children"
            $count = [int](Get-ComProperty $children "Count")
            for ($i = 0; $i -lt $count; $i++) {
                $child = Get-ComProperty $children "Item" @($i)
                Invoke-SapControlWalker $child ([ref]$Accum)
            }
        } catch {}
    }
}

function New-EmptySnapshot {
    param([string]$ErrorMessage = "")
    $result = [ordered]@{
        timestamp   = (Get-Date).ToUniversalTime().ToString("o")
        transaction = ""
        program     = ""
        screen      = ""
        controls    = @()
    }
    if ($ErrorMessage) { $result.error = $ErrorMessage }
    return $result
}

function Get-SapRootWindow {
    $fg = [SapSnapWin32]::GetForegroundWindow()
    $sapRoot = [SapSnapWin32]::FindSapRoot($fg)
    if ($sapRoot -ne [IntPtr]::Zero) { return $sapRoot }
    $sapWindows = [SapSnapWin32]::EnumSapWindows()
    if ($sapWindows.Count -gt 0) { return $sapWindows[0] }
    return [IntPtr]::Zero
}

function Test-IsFieldClass {
    param([string]$ClassName, [string]$ControlType = "")
    $c = "$ClassName $ControlType"
    return ($c -match '(?i)\bEdit\b|ComboBox|ComboLBox|TextBox|Spinner|MSComCtl|SapEdit|SAP.*Edit|GuiCTextField|TextField')
}

function Test-IsLabelClass {
    param([string]$ClassName, [string]$ControlType = "", [string]$Text = "")
    if ($Text -match '[:：]\s*$') { return $true }
    $c = "$ClassName $ControlType"
    return ($c -match '(?i)Static|Label|TextBlock|LText')
}

function Find-NearestLabel {
    param($Field, $Labels, [int]$RowTolerance = 14)
    $best = $null
    $bestScore = [double]::PositiveInfinity
    foreach ($label in $Labels) {
        $rowDelta = [Math]::Abs($label.Top - $Field.Top)
        if ($rowDelta -gt $RowTolerance) { continue }
        if ($label.Left -ge $Field.Left) { continue }
        $gap = $Field.Left - $label.Left
        if ($gap -gt 600) { continue }
        $score = ($rowDelta * 10) + $gap
        if ($score -lt $bestScore) {
            $bestScore = $score
            $best = $label
        }
    }
    if ($null -eq $best) { return "" }
    return ($best.Text -replace '[:：]\s*$', '').Trim()
}

function Convert-UiNodeToControl {
    param($Node, [string]$SapType, [string]$Name, [string]$Value)
    return [ordered]@{
        id      = $Name
        type    = $SapType
        name    = $Name
        text    = $Value
        value   = $Value
        tooltip = ""
        left    = [int]$Node.Left
        top     = [int]$Node.Top
        width   = [int]([Math]::Max(0, $Node.Right - $Node.Left))
        height  = [int]([Math]::Max(0, $Node.Bottom - $Node.Top))
    }
}

function Get-FieldsViaUia {
    param([IntPtr]$SapRoot)
    $controls = New-Object System.Collections.ArrayList
    if ($SapRoot -eq [IntPtr]::Zero) { return @($controls.ToArray()) }

    try {
        Add-Type -AssemblyName UIAutomationClient -ErrorAction Stop
        Add-Type -AssemblyName UIAutomationTypes -ErrorAction Stop
    } catch {
        return @($controls.ToArray())
    }

    try {
        $root = [System.Windows.Automation.AutomationElement]::FromHandle($SapRoot)
        if ($null -eq $root) { return @($controls.ToArray()) }

        $rect = $root.Current.BoundingRectangle
        $topCutoff = $rect.Top + ($rect.Height * 0.10)
        $bottomCutoff = $rect.Bottom - ($rect.Height * 0.12)

        $trueCondition = [System.Windows.Automation.Condition]::TrueCondition
        $all = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $trueCondition)
        $labels = New-Object System.Collections.ArrayList
        $fields = New-Object System.Collections.ArrayList

        foreach ($el in $all) {
            try {
                $name = [string]$el.Current.Name
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                $r = $el.Current.BoundingRectangle
                if ($r.Bottom -lt $topCutoff -or $r.Top -gt $bottomCutoff) { continue }

                $node = [SapSnapWin32+UiNode]@{
                    Text = $name.Trim()
                    ClassName = [string]$el.Current.ClassName
                    Left = [int]$r.Left
                    Top = [int]$r.Top
                    Right = [int]$r.Right
                    Bottom = [int]$r.Bottom
                    Source = "uia"
                    ControlType = [string]$el.Current.ControlType.ProgrammaticName
                }

                $ct = [string]$el.Current.ControlType.ProgrammaticName
                $value = ""
                try {
                    $vp = $el.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
                    if ($null -ne $vp) { $value = [string]$vp.Current.Value }
                } catch {}

                if ($ct -match 'Edit|ComboBox|Document' -and -not [string]::IsNullOrWhiteSpace($value)) {
                    $node.Text = $value.Trim()
                    [void]$fields.Add($node)
                } elseif ($ct -match 'Text' -or (Test-IsLabelClass $node.ClassName $ct $node.Text)) {
                    [void]$labels.Add($node)
                }
            } catch {}
        }

        $seen = @{}
        foreach ($field in $fields) {
            $value = $field.Text
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            $labelText = Find-NearestLabel $field $labels
            $name = if ($labelText) { ($labelText -replace '\s+', '_') } else { "field_$($field.Left)_$($field.Top)" }
            if ($seen.ContainsKey($name)) { continue }
            $seen[$name] = $true
            $sapType = if ($field.ControlType -match 'ComboBox') { "GuiComboBox" } else { "GuiTextField" }
            [void]$controls.Add((Convert-UiNodeToControl $field $sapType $name $value))
            if ($labelText) {
                [void]$controls.Add((Convert-UiNodeToControl $field "GuiLabel" "${name}_label" $labelText))
            }
        }
    } catch {}

    return @($controls.ToArray())
}

function Get-FieldsViaWin32 {
    param([IntPtr]$SapRoot)
    $controls = New-Object System.Collections.ArrayList
    if ($SapRoot -eq [IntPtr]::Zero) { return @($controls.ToArray()) }

    $rect = New-Object SapSnapWin32+RECT
    if (-not [SapSnapWin32]::TryGetWindowRect($SapRoot, [ref]$rect)) { return @($controls.ToArray()) }

    $items = New-Object 'System.Collections.Generic.List[SapSnapWin32+UiNode]'
    [SapSnapWin32]::CollectUiNodes($SapRoot, $items)

    $windowHeight = [Math]::Max(1, $rect.Bottom - $rect.Top)
    $topCutoff = $rect.Top + [int]($windowHeight * 0.10)
    $bottomCutoff = $rect.Bottom - [int]($windowHeight * 0.12)

    $labels = @()
    $fields = @()
    foreach ($item in $items) {
        if ($item.Top -lt $topCutoff -or $item.Bottom -gt $bottomCutoff) { continue }
        if ($item.Text.Length -gt 120) { continue }
        if (Test-IsFieldClass $item.ClassName $item.ControlType) {
            $fields += $item
        } elseif (Test-IsLabelClass $item.ClassName $item.ControlType $item.Text) {
            $labels += $item
        } elseif ($item.Text -match '^\s*\S.{0,40}[:：]\s*$') {
            $labels += $item
        }
    }

    $seen = @{}
    foreach ($field in $fields) {
        $value = $field.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        $labelText = Find-NearestLabel $field $labels
        $name = if ($labelText) { ($labelText -replace '\s+', '_') } else { "field_$($field.Left)_$($field.Top)" }
        if ($seen.ContainsKey($name)) { continue }
        $seen[$name] = $true
        $sapType = if ($field.ClassName -match '(?i)Combo') { "GuiComboBox" } else { "GuiTextField" }
        [void]$controls.Add((Convert-UiNodeToControl $field $sapType $name $value))
        if ($labelText) {
            [void]$controls.Add((Convert-UiNodeToControl $field "GuiLabel" "${name}_label" $labelText))
        }
    }

    return @($controls.ToArray())
}

function Get-FallbackFieldControls {
    param([IntPtr]$SapRoot)
    $uia = Get-FieldsViaUia $SapRoot
    if ($uia.Count -gt 0) { return $uia }
    return Get-FieldsViaWin32 $SapRoot
}

function Test-ShouldSkipChromeId {
    param([string]$Id)
    if ([string]::IsNullOrWhiteSpace($Id)) { return $false }
    return ($Id -match '(?i)/(sbar|tbar|mbar|toolbar|titl|okcd)(/|$)')
}

function Invoke-ScriptingSnapshot {
    param($session)
    $transaction = ""
    $program = ""
    $screen = ""
    $info = Get-ComProperty $session "Info"
    if ($null -ne $info) {
        $transaction = Get-ComPropertySafe $info "Transaction"
        $program = Get-ComPropertySafe $info "Program"
        $screen = Get-ComPropertySafe $info "ScreenNumber"
    }

    $controls = New-Object System.Collections.ArrayList
    $seenIds = @{}

    function Add-WalkedControl {
        param($control)
        if ($null -eq $control) { return }
        $ref = [ref]$controls
        Invoke-SapControlWalker $control $ref
    }

    $usr = Invoke-ComMethod $session "findById" @("wnd[0]/usr")
    if ($null -ne $usr) {
        Add-WalkedControl $usr
    }

    $wnd0 = Invoke-ComMethod $session "findById" @("wnd[0]")
    if ($null -ne $wnd0) {
        try {
            $children = Get-ComProperty $wnd0 "Children"
            $count = [int](Get-ComProperty $children "Count")
            for ($i = 0; $i -lt $count; $i++) {
                $child = Get-ComProperty $children "Item" @($i)
                if ($null -eq $child) { continue }
                $childId = Get-ComPropertySafe $child "Id"
                if (Test-ShouldSkipChromeId $childId) { continue }
                if ($childId -and $childId -ne "wnd[0]/usr") {
                    Add-WalkedControl $child
                }
            }
        } catch {}
    }

    # De-duplicate by control id
    $unique = New-Object System.Collections.ArrayList
    $dedupe = @{}
    foreach ($item in $controls) {
        $key = if ($item.id) { $item.id } else { "$($item.name)|$($item.top)|$($item.left)" }
        if ($dedupe.ContainsKey($key)) { continue }
        $dedupe[$key] = $true
        [void]$unique.Add($item)
    }

    return @{
        transaction = $transaction
        program     = $program
        screen      = $screen
        controls    = @($unique.ToArray())
        captureSource = "scripting"
    }
}

try {
    [void](Focus-SapWindow)
    $win32Title = Get-ForegroundSapTitle
    $sessionResult = Get-BestSapSession $win32Title
    $session = $sessionResult.session
    $scriptDiag = $sessionResult.scriptDiag
    $transaction = ""
    $program = ""
    $screen = ""
    $controls = @()
    $captureSource = ""
    $errorMessage = ""

    if (-not $UiaOnly -and $null -ne $session) {
        $scriptSnap = Invoke-ScriptingSnapshot $session
        $transaction = $scriptSnap.transaction
        $program = $scriptSnap.program
        $screen = $scriptSnap.screen
        $controls = $scriptSnap.controls
        $captureSource = $scriptSnap.captureSource
    }

    if ($controls.Count -eq 0) {
        $sapRoot = Get-SapRootWindow
        if ($sapRoot -ne [IntPtr]::Zero) {
            $fallbackControls = Get-FallbackFieldControls $sapRoot
            if ($fallbackControls.Count -gt 0) {
                $controls = $fallbackControls
                $captureSource = "uia-fallback"
                $errorMessage = ""
            }
        }
    }

    if ($controls.Count -eq 0) {
        if ($UiaOnly) {
            $errorMessage = "No SAP fields detected via UI Automation on this screen"
        } elseif ($null -eq $session) {
            $errorMessage = 'SAP GUI Scripting unavailable - enable Options, Accessibility and Scripting, Enable scripting for field capture (Order Type, etc.)'
        } else {
            $errorMessage = "No SAP fields detected on this screen"
        }
    }

    $sapRoot = Get-SapRootWindow
    $sapRect = Get-SapWindowRectJson
    $scriptingEngine = Get-SapGuiApplication

    $result = [ordered]@{
        timestamp     = (Get-Date).ToUniversalTime().ToString("o")
        transaction   = $transaction
        program       = $program
        screen        = $screen
        controls      = @($controls)
        captureSource = $captureSource
        sapWindowRect = $sapRect
        debug         = [ordered]@{
            win32Title               = $win32Title
            scriptingEngineAvailable = ($null -ne $scriptingEngine)
            sessionConnected         = ($null -ne $session)
            sapRootFound             = ($sapRoot -ne [IntPtr]::Zero)
            controlCount             = $controls.Count
            scriptingDiag            = $scriptDiag
        }
    }
    if ($errorMessage) { $result.error = $errorMessage }
    $result | ConvertTo-Json -Compress -Depth 8
} catch {
    New-EmptySnapshot -ErrorMessage $_.Exception.Message | ConvertTo-Json -Compress -Depth 8
}
