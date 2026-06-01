const fs = require("fs");
const path = require("path");

const DEFAULT_SETTINGS = {
	captureHotkey: "Alt+C",
	sessionHotkey: "CommandOrControl+Shift+S",
	lastProjectPath: "",
	maxCaptureFields: 10,
	/** "sap" = focus SAP GUI and capture fields; "non-sap" = active window only (any app) */
	captureTarget: "sap",
};

function normalizeCaptureTarget(value) {
	return value === "non-sap" ? "non-sap" : "sap";
}

function getSettingsPath(baseDir) {
	return path.join(baseDir, "data", "settings.json");
}

function loadSettings(baseDir) {
	const filePath = getSettingsPath(baseDir);
	try {
		if (fs.existsSync(filePath)) {
			return { ...DEFAULT_SETTINGS, ...JSON.parse(fs.readFileSync(filePath, "utf8")) };
		}
	} catch {
		/* use defaults */
	}
	return { ...DEFAULT_SETTINGS };
}

function saveSettings(baseDir, settings) {
	const dir = path.join(baseDir, "data");
	if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
	const merged = { ...DEFAULT_SETTINGS, ...settings };
	fs.writeFileSync(getSettingsPath(baseDir), JSON.stringify(merged, null, 2));
	return merged;
}

module.exports = { DEFAULT_SETTINGS, normalizeCaptureTarget, loadSettings, saveSettings };
