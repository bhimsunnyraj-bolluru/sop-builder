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
		foregroundIsSap: false,
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

/** True when the captured foreground window is SAP GUI (not Chrome, Office, etc.). */
function isSapForegroundContext(context) {
	if (!context) return false;
	if (context.foregroundIsSap === true) return true;
	if (context.foregroundIsSap === false) return false;
	const cls = String(context.windowClass || "");
	if (cls === "SAP_FRONTEND_SESSION") return true;
	if (/^Chrome_/i.test(cls) || /^MozillaWindowClass$/i.test(cls)) return false;
	if (String(context.source || "").includes("sap-scripting") && context.transaction) return true;
	return false;
}

/** True when foreground is a normal web browser (Chrome/Edge/Firefox), not SAP. */
function isWebBrowserForeground(context) {
	if (!context || isSapForegroundContext(context)) return false;
	const cls = String(context.windowClass || "");
	if (cls === "SAP_FRONTEND_SESSION") return false;
	const title = resolveForegroundTitle(context);
	if (!title) return false;
	// SAP GUI for Windows uses Chrome_WidgetWin_1 for some embedded views — require browser suffix in title.
	if (/^Chrome_/i.test(cls)) {
		return /\s-\s*Google Chrome\s*$/i.test(title) || /\s-\s*Microsoft Edge\s*$/i.test(title) || /\s-\s*Mozilla Firefox\s*$/i.test(title);
	}
	if (cls === "MozillaWindowClass") return true;
	if (cls === "ApplicationFrameWindow" && /\s-\s*Microsoft Edge\s*$/i.test(title)) return true;
	return false;
}

function isOwnAppWindowTitle(title) {
	return /\bSOP Builder\b/i.test(String(title || ""));
}

function isDevToolOrOwnWindowTitle(title) {
	const t = String(title || "").trim();
	if (!t) return true;
	if (isOwnAppWindowTitle(t)) return true;
	if (/\s-\s*Cursor\s*$/i.test(t)) return true;
	if (/Visual Studio Code/i.test(t)) return true;
	return false;
}

/** Prefer a real foreground window title (never our own Electron shell or IDE). */
function resolveForegroundTitle(context) {
	const candidates = [
		context.sapWindowTitle,
		context.title,
	].map((t) => String(t || "").trim()).filter(Boolean);
	for (const t of candidates) {
		if (!isDevToolOrOwnWindowTitle(t)) return t;
	}
	return candidates[0] || "";
}

/** SAP status bar text that indicates a completed save/post/create action. */
function isSapSuccessStatusBar(text) {
	const s = String(text || "").trim();
	if (!s) return false;
	const lower = s.toLowerCase();
	if (lower.includes("success")) return true;
	if (/\b(has been|was)\s+saved\b/i.test(s)) return true;
	if (/\b(has been|was)\s+(created|posted|changed|updated|released)\b/i.test(s)) return true;
	if (/\b(saved|created|posted)\b/i.test(s) && /\b\d{4,}\b/.test(s)) return true;
	return false;
}

/** Build a readable step title from SAP window / scripting metadata. */
function formatStepDescription(context) {
	if (!context) return "";

	const statusBar = (context.statusBar || "").trim();
	if (isSapSuccessStatusBar(statusBar)) return statusBar;

	const windowTitle = resolveForegroundTitle(context);
	const transaction = isSapForegroundContext(context) ? (context.transaction || "").trim() : "";
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

/** Fast context from VBS snapshot — avoids PowerShell on the capture hot path. */
function contextFromSnapshot(snapshot) {
	if (!snapshot) return emptyContext();
	const attach = snapshot.debug?.sessionAttach || {};
	const txn = String(snapshot.transaction || attach.attachedTxn || "").trim();
	const title = String(attach.attachedWindowTitle || txn).trim();
	return {
		title,
		transaction: txn,
		program: String(snapshot.program || attach.attachedProgram || "").trim(),
		screenNumber: String(snapshot.screen || attach.attachedScreen || "").trim(),
		systemName: "",
		statusBar: "",
		windowClass: "",
		sapWindowTitle: title,
		sapWindowRect: snapshot.sapWindowRect || null,
		source: snapshot.debug?.sessionConnected ? "sap-scripting-vbs" : "none",
	};
}

/** Merge VBS snapshot metadata with PowerShell context (status bar, window rect). */
function mergeContextFromSnapshotAndPs(snapshot, psContext, captureIntent = "sap") {
	if (!psContext || psContext.source === "none") {
		return contextFromSnapshot(snapshot);
	}

	const psTitle = resolveForegroundTitle(psContext);
	const rect = psContext.sapWindowRect || null;
	const nonSapCapture = captureIntent === "non-sap";

	if (nonSapCapture) {
		return {
			...emptyContext(),
			title: psTitle,
			sapWindowTitle: psTitle,
			sapWindowRect: null,
			windowClass: String(psContext.windowClass || "").trim(),
			foregroundIsSap: false,
			captureTarget: "non-sap",
			source: psContext.source || "win32-foreground",
		};
	}

	const base = contextFromSnapshot(snapshot);
	const statusBar = String(psContext.statusBar || "").trim();
	let source = base.source;
	if (statusBar) {
		source = `${base.source}+ps-status`;
	} else if (psContext.source && psContext.source !== "none") {
		source = `${base.source}+${psContext.source}`;
	}

	const sapTitle = isDevToolOrOwnWindowTitle(psTitle) ? base.sapWindowTitle || psTitle : psTitle || base.sapWindowTitle;

	return {
		...base,
		statusBar: statusBar || base.statusBar,
		systemName: String(psContext.systemName || base.systemName || "").trim(),
		sapWindowTitle: sapTitle,
		sapWindowRect: rect || base.sapWindowRect,
		title: sapTitle,
		transaction: base.transaction || String(psContext.transaction || "").trim(),
		program: base.program || String(psContext.program || "").trim(),
		screenNumber: base.screenNumber || String(psContext.screenNumber || "").trim(),
		windowClass: psContext.windowClass || base.windowClass,
		foregroundIsSap: true,
		captureTarget: "sap",
		source,
	};
}

async function getSapGuiContext(options = {}) {
	if (process.platform !== "win32") return emptyContext();
	const captureTarget = options.captureTarget === "non-sap" ? "non-sap" : "sap";
	const psMode = captureTarget === "non-sap" ? "NonSap" : "Sap";
	const { execFile } = require("child_process");
	const { promisify } = require("util");
	const { getPsScriptPath } = require("../paths");
	const { getSapPowerShellExe } = require("./sap-powershell");
	const execFileAsync = promisify(execFile);
	try {
		const { stdout } = await execFileAsync(
			getSapPowerShellExe(),
			["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", getPsScriptPath(), "-CaptureMode", psMode],
			{ encoding: "utf8", windowsHide: true, timeout: 15000, maxBuffer: 1024 * 1024 }
		);
		return parseContextOutput(stdout);
	} catch {
		return emptyContext();
	}
}

module.exports = {
	getSapGuiContext,
	contextFromSnapshot,
	mergeContextFromSnapshotAndPs,
	isSapForegroundContext,
	isWebBrowserForeground,
	isSapSuccessStatusBar,
	resolveForegroundTitle,
	formatStepDescription,
	describeCaptureContext,
	emptyContext,
};
