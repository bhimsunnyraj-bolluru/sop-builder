const { execFile } = require("child_process");
const { promisify } = require("util");
const { getFocusSapScriptPath } = require("../paths");
const { getSapPowerShellExe } = require("./sap-powershell");

const execFileAsync = promisify(execFile);

async function focusSapWindow() {
	if (process.platform !== "win32") return false;
	try {
		const { stdout } = await execFileAsync(
			getSapPowerShellExe(),
			["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", getFocusSapScriptPath()],
			{ encoding: "utf8", windowsHide: true, timeout: 5000 }
		);
		const raw = (stdout || "").trim();
		if (!raw) return false;
		const parsed = JSON.parse(raw);
		return parsed.ok === true;
	} catch {
		return false;
	}
}

module.exports = { focusSapWindow };
