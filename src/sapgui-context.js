const path = require("path");
const { execFile } = require("child_process");
const { promisify } = require("util");

const execFileAsync = promisify(execFile);

const PS_SCRIPT = path.join(__dirname, "..", "scripts", "get-sapgui-context.ps1");

function emptyContext() {
	return {
		title: "",
		transaction: "",
		program: "",
		screenNumber: "",
		systemName: "",
		statusBar: "",
		windowClass: "",
		sapWindowTitle: "",
		source: "none",
	};
}

function parseContextOutput(stdout) {
	const raw = (stdout || "").trim();
	if (!raw) return emptyContext();
	try {
		const parsed = JSON.parse(raw);
		return { ...emptyContext(), ...parsed };
	} catch {
		return { ...emptyContext(), title: raw, source: "raw" };
	}
}

async function getSapGuiContext() {
	if (process.platform !== "win32") return emptyContext();
	try {
		const { stdout } = await execFileAsync(
			"powershell",
			["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", PS_SCRIPT],
			{ encoding: "utf8", windowsHide: true, timeout: 15000, maxBuffer: 1024 * 1024 }
		);
		return parseContextOutput(stdout);
	} catch {
		return emptyContext();
	}
}

function formatStepDescription(context) {
	if (!context) return "";
	if (context.title) return context.title;

	const parts = [];
	if (context.sapWindowTitle) parts.push(context.sapWindowTitle);
	if (context.transaction) parts.push(context.transaction);
	if (context.program) parts.push(`(${context.program})`);
	if (context.statusBar) parts.push(`- ${context.statusBar}`);
	return parts.join(" ");
}

module.exports = { getSapGuiContext, formatStepDescription };
