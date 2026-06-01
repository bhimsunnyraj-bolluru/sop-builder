const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");
const { promisify } = require("util");
const screenshot = require("screenshot-desktop");
const { getProjectRoot, isPackaged } = require("../paths");

const execFileAsync = promisify(execFile);

function getCropScriptPath() {
	if (isPackaged()) {
		return path.join(process.resourcesPath, "app.asar.unpacked", "scripts", "crop-screenshot.ps1");
	}
	return path.join(getProjectRoot(), "scripts", "crop-screenshot.ps1");
}

function normalizeCropRect(rect) {
	if (!rect || typeof rect !== "object") return null;
	const left = Number(rect.left);
	const top = Number(rect.top);
	const right = Number(rect.right);
	const bottom = Number(rect.bottom);
	if (!Number.isFinite(left) || !Number.isFinite(top) || !Number.isFinite(right) || !Number.isFinite(bottom)) {
		return null;
	}
	if (right - left < 80 || bottom - top < 80) return null;
	return { left, top, right, bottom };
}

async function cropScreenshot(imagePath, outPath, rect) {
	const script = getCropScriptPath();
	await execFileAsync(
		"powershell",
		[
			"-NoProfile",
			"-ExecutionPolicy",
			"Bypass",
			"-File",
			script,
			"-ImagePath",
			imagePath,
			"-OutPath",
			outPath,
			"-Left",
			String(rect.left),
			"-Top",
			String(rect.top),
			"-Right",
			String(rect.right),
			"-Bottom",
			String(rect.bottom),
		],
		{ windowsHide: true, timeout: 30000 }
	);
}

/**
 * Capture screen PNG. When cropRect is provided (SAP window bounds), crop to that region.
 * @param {string} filename
 * @param {{ left, top, right, bottom } | null} cropRect
 */
async function takeScreenshot(filename, cropRect = null) {
	const rect = normalizeCropRect(cropRect);
	if (!rect) {
		await screenshot({ filename });
		return filename;
	}

	const tempFile = filename.replace(/\.png$/i, "._full.png");
	await screenshot({ filename: tempFile });
	try {
		await cropScreenshot(tempFile, filename, rect);
	} finally {
		try {
			if (fs.existsSync(tempFile)) fs.unlinkSync(tempFile);
		} catch {
			/* ignore */
		}
	}
	return filename;
}

module.exports = { takeScreenshot, normalizeCropRect };
