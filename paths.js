const fs = require("fs");
const path = require("path");
const { app } = require("electron");

function isPackaged() {
	// `electron .` / npm start
	if (process.defaultApp === true) return false;
	try {
		if (app.isPackaged) return true;
	} catch {
		/* app may be unavailable in some renderer contexts */
	}
	// Packaged app runs as SOP Builder.exe, not electron.exe
	return path.basename(process.execPath).toLowerCase() !== "electron.exe";
}

/** Writable app folder: project root in dev, folder containing the .exe when installed. */
function getProjectRoot() {
	if (isPackaged()) return path.dirname(process.execPath);
	return __dirname;
}

function getScreenshotsDir() {
	return path.join(getProjectRoot(), "screenshots");
}

function getDataDir() {
	return path.join(getProjectRoot(), "data");
}

function getSopsDir() {
	return path.join(getDataDir(), "sops");
}

function getExportsDir() {
	return path.join(getProjectRoot(), "exports");
}

function getPsScriptPath() {
	if (isPackaged()) {
		return path.join(
			process.resourcesPath,
			"app.asar.unpacked",
			"scripts",
			"get-sapgui-context.ps1"
		);
	}
	return path.join(getProjectRoot(), "scripts", "get-sapgui-context.ps1");
}

function ensureDir(dir) {
	if (fs.existsSync(dir)) {
		if (!fs.statSync(dir).isDirectory()) {
			throw new Error(`Expected a folder but found a file: ${dir}`);
		}
		return dir;
	}
	fs.mkdirSync(dir, { recursive: true });
	return dir;
}

module.exports = {
	isPackaged,
	getProjectRoot,
	getScreenshotsDir,
	getDataDir,
	getSopsDir,
	getExportsDir,
	getPsScriptPath,
	ensureDir,
};
