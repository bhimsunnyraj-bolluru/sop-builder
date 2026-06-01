# Quick SAP GUI Scripting diagnostic — run while SAP session (e.g. VA01) is open.
$ErrorActionPreference = 'Continue'

Write-Host "=== SAP GUI Scripting Diagnostic ===" -ForegroundColor Cyan
Write-Host "PowerShell bitness: $([IntPtr]::Size * 8)-bit"

function Try-ComMethod($Object, [string]$Method, [object[]]$Args = @()) {
    if ($null -eq $Object) { return $null }
    try {
        return $Object.GetType().InvokeMember($Method, [System.Reflection.BindingFlags]::InvokeMethod, $null, $Object, $Args)
    } catch { return $null }
}

Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue

$sapGuiAuto = $null
$method = ""
foreach ($progId in @("SapROTWr.SapROTWrapper", "SAPROTWr.SapROTWrapper")) {
    try {
        $wrapper = New-Object -ComObject $progId
        $sapGuiAuto = Try-ComMethod $wrapper "GetROTEntry" @("SAPGUI")
        if ($null -ne $sapGuiAuto) { $method = $progId; break }
    } catch {}
}
if (-not $sapGuiAuto) {
    try { $sapGuiAuto = [Runtime.InteropServices.Marshal]::GetActiveObject("SAPGUI"); $method = "GetActiveObject" } catch {}
}
if (-not $sapGuiAuto) {
    try { $sapGuiAuto = [Microsoft.VisualBasic.Interaction]::GetObject("SAPGUI"); $method = "GetObject" } catch {}
}

Write-Host "SAPGUI COM: $(if ($sapGuiAuto) { "OK via $method" } else { "NOT FOUND" })"

$engine = $null
if ($sapGuiAuto) {
    try {
        $engine = $sapGuiAuto.GetScriptingEngine()
        if ($null -ne $engine) { Write-Host "GetScriptingEngine (direct call): OK" -ForegroundColor Green }
    } catch {
        Write-Host "GetScriptingEngine (direct call): $($_.Exception.Message)"
    }
    if ($null -eq $engine) {
        $engine = Try-ComMethod $sapGuiAuto "GetScriptingEngine"
        if ($null -ne $engine) { Write-Host "GetScriptingEngine (InvokeMember): OK" -ForegroundColor Green }
    }
}
Write-Host "GetScriptingEngine: $(if ($engine) { "OK" } else { "NULL - scripting blocked or no active session" })" -ForegroundColor $(if ($engine) { "Green" } else { "Red" })

if ($engine) {
    try {
        $session = $engine.ActiveSession
        $info = $session.Info
        Write-Host "Active session t-code: $($info.Transaction)"
        $fld = Try-ComMethod $session "findById" @("wnd[0]/usr/ctxtVBAK-AUART")
        if ($fld) {
            Write-Host "Order Type field: $($fld.Text)" -ForegroundColor Green
        } else {
            Write-Host "Order Type field id not found on this screen"
        }
    } catch {
        Write-Host "Session read error: $($_.Exception.Message)"
    }
}

Write-Host "`nRegistry (scripting):"
Get-ChildItem "HKCU:\Software\SAP" -Recurse -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match 'Script' } |
    ForEach-Object {
        Write-Host "  $($_.Name)"
        try {
            Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction Stop |
                Format-List UserScripting, ScriptingSecurityLevel, *Script* |
                Out-String | Write-Host
        } catch {}
    }

Write-Host "All values under Scripting key:"
try {
    Get-ItemProperty "HKCU:\Software\SAP\SAPGUI Front\SAP Frontend Server\Scripting" |
        Format-List * -Exclude PS* | Out-String | Write-Host
} catch { Write-Host "  (not found)" }

Write-Host "Search UserScripting anywhere under HKCU\Software\SAP:"
Get-ChildItem "HKCU:\Software\SAP" -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
    if ($null -eq $props) { return }
    if ($null -ne $props.UserScripting) {
        Write-Host "  $($_.PSPath) UserScripting=$($props.UserScripting)"
    }
    if ($null -ne $props.ScriptingSecurityLevel) {
        Write-Host "  $($_.PSPath) ScriptingSecurityLevel=$($props.ScriptingSecurityLevel)"
    }
}

Write-Host "`nIf GetScriptingEngine is NULL with Enable scripting checked:"
Write-Host "  1. SAP GUI Options > Scripting > open Security / Configuration > set Allow"
Write-Host "  2. Ask Basis: RZ11 parameter sapgui/user_scripting = TRUE on system S18"
Write-Host "  3. Close Options dialog, keep VA01 open, run this script again"
