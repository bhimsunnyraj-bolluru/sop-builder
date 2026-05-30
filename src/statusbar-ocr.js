const { execFile } = require("child_process");
const { promisify } = require("util");
const path = require("path");
const { getProjectRoot } = require("../paths");

const execFileAsync = promisify(execFile);

function getOcrScriptPath() {
	const root = getProjectRoot();
	const { app } = require("electron");
	if (app.isPackaged) {
		return path.join(
			process.resourcesPath,
			"app.asar.unpacked",
			"scripts",
			"read-statusbar-ocr.ps1"
		);
	}
	return path.join(root, "scripts", "read-statusbar-ocr.ps1");
}

function normalizeStatusBarText(text) {
	if (!text) return "";
	return String(text)
		.replace(/\s+/g, " ")
		.replace(/[|]/g, " ")
		.trim();
}

/** Read SAP status bar text from the bottom band of a screenshot (Windows OCR). */
async function readStatusBarFromScreenshot(imagePath, sapWindowRect) {
	if (process.platform !== "win32" || !imagePath) return "";

	const args = ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", getOcrScriptPath(), "-ImagePath", imagePath];
	if (sapWindowRect && sapWindowRect.right > sapWindowRect.left && sapWindowRect.bottom > sapWindowRect.top) {
		args.push(
			"-Left", String(Math.round(sapWindowRect.left)),
			"-Top", String(Math.round(sapWindowRect.top)),
			"-Right", String(Math.round(sapWindowRect.right)),
			"-Bottom", String(Math.round(sapWindowRect.bottom))
		);
	}

	try {
		const { stdout } = await execFileAsync("powershell", args, {
			encoding: "utf8",
			windowsHide: true,
			timeout: 20000,
			maxBuffer: 1024 * 1024,
		});
		return normalizeStatusBarText(stdout);
	} catch {
		return "";
	}
}

module.exports = { readStatusBarFromScreenshot, normalizeStatusBarText };
