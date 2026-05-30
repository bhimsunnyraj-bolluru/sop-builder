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
		sapWindowRect: null,
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

/** Build a readable step title from SAP window / scripting metadata. */
function formatStepDescription(context) {
	if (!context) return "";

	const windowTitle = (context.sapWindowTitle || context.title || "").trim();
	const transaction = (context.transaction || "").trim();
	const statusBar = (context.statusBar || "").trim();
	const systemName = (context.systemName || "").trim();

	let desc = windowTitle || transaction;

	if (transaction && desc && !desc.toUpperCase().includes(transaction.toUpperCase())) {
		desc = `${desc} (${transaction})`;
	}

	if (systemName && desc && !desc.includes(`[${systemName}]`) && !desc.includes(systemName)) {
		desc = `${desc} [${systemName}]`;
	}

	if (statusBar && desc && !desc.includes(statusBar)) {
		desc = `${desc} — ${statusBar}`;
	} else if (statusBar && !desc) {
		desc = statusBar;
	}

	return desc.trim();
}

function describeCaptureContext(context) {
	if (!context || context.source === "none") return "";
	const parts = [];
	if (context.sapWindowTitle || context.title) parts.push("window title");
	if (context.transaction) parts.push(`t-code ${context.transaction}`);
	if (context.statusBar) parts.push("status bar");
	if (context.source && String(context.source).includes("ocr")) parts.push("OCR");
	if (context.systemName) parts.push(context.systemName);
	if (!parts.length) return "";
	const via = (() => {
		const src = context.source || "";
		if (src.includes("sap-scripting")) return "SAP Scripting";
		if (src.includes("ocr")) return "screenshot OCR";
		if (src.includes("uia")) return "UI Automation";
		if (src.includes("child")) return "Win32 controls";
		if (src === "win32") return "Win32";
		return src;
	})();
	return ` (${via}: ${parts.join(", ")})`;
}

async function getSapGuiContext() {
	if (process.platform !== "win32") return emptyContext();
	const { execFile } = require("child_process");
	const { promisify } = require("util");
	const { getPsScriptPath } = require("../paths");
	const execFileAsync = promisify(execFile);
	try {
		const { stdout } = await execFileAsync(
			"powershell",
			["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", getPsScriptPath()],
			{ encoding: "utf8", windowsHide: true, timeout: 15000, maxBuffer: 1024 * 1024 }
		);
		return parseContextOutput(stdout);
	} catch {
		return emptyContext();
	}
}

module.exports = {
	getSapGuiContext,
	formatStepDescription,
	describeCaptureContext,
	emptyContext,
};
