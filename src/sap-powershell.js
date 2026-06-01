/**
 * SAP GUI for Windows is usually 32-bit. Run SAP COM scripts via 32-bit PowerShell when available.
 */

const fs = require("fs");
const path = require("path");

function getPowerShellExe(prefer32Bit = false) {
	if (process.platform !== "win32") return "powershell";

	const systemRoot = process.env.SystemRoot || process.env.WINDIR || "C:\\Windows";
	const ps64 = path.join(systemRoot, "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
	const ps32 = path.join(systemRoot, "SysWOW64", "WindowsPowerShell", "v1.0", "powershell.exe");

	if (prefer32Bit && fs.existsSync(ps32)) return ps32;
	if (fs.existsSync(ps64)) return ps64;
	if (fs.existsSync(ps32)) return ps32;
	return "powershell";
}

/** Prefer 32-bit PowerShell for SAP GUI Scripting COM (GetScriptingEngine). */
function getSapPowerShellExe() {
	return getPowerShellExe(true);
}

module.exports = { getPowerShellExe, getSapPowerShellExe };
