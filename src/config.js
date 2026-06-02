const fs = require("fs");
const path = require("path");

/** Company-wide branding reused on every Word export. */
const DEFAULT_BRANDING = {
	companyName: "Business Core Solutions",
	logoPath: "",
	accentColor: "#F97316",
	footerText: "Confidential — for internal use only",
	/** "standard" = cover + headers/footers; "minimal" = simple header only; "detailed" = cover + revision history */
	template: "standard",
};

const DEFAULT_SETTINGS = {
	captureHotkey: "Alt+C",
	sessionHotkey: "CommandOrControl+Shift+S",
	lastProjectPath: "",
	maxCaptureFields: 10,
	/** "sap" = focus SAP GUI and capture fields; "non-sap" = active window only (any app) */
	captureTarget: "sap",
	branding: { ...DEFAULT_BRANDING },
};

function normalizeCaptureTarget(value) {
	return value === "non-sap" ? "non-sap" : "sap";
}

/** Merge a partial branding object over the defaults so exports never see missing keys. */
function normalizeBranding(value) {
	const b = value && typeof value === "object" ? value : {};
	return { ...DEFAULT_BRANDING, ...b };
}

function getSettingsPath(baseDir) {
	return path.join(baseDir, "data", "settings.json");
}

function loadSettings(baseDir) {
	const filePath = getSettingsPath(baseDir);
	try {
		if (fs.existsSync(filePath)) {
			const parsed = JSON.parse(fs.readFileSync(filePath, "utf8"));
			const merged = { ...DEFAULT_SETTINGS, ...parsed };
			merged.branding = normalizeBranding(parsed.branding);
			return merged;
		}
	} catch {
		/* use defaults */
	}
	return { ...DEFAULT_SETTINGS, branding: { ...DEFAULT_BRANDING } };
}

function saveSettings(baseDir, settings) {
	const dir = path.join(baseDir, "data");
	if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
	const merged = { ...DEFAULT_SETTINGS, ...settings };
	merged.branding = normalizeBranding(settings ? settings.branding : null);
	fs.writeFileSync(getSettingsPath(baseDir), JSON.stringify(merged, null, 2));
	return merged;
}

module.exports = {
	DEFAULT_SETTINGS,
	DEFAULT_BRANDING,
	normalizeCaptureTarget,
	normalizeBranding,
	loadSettings,
	saveSettings,
};
