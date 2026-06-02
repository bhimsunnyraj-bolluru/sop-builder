/**
 * Captures SAP GUI screen state via VBScript. No PS fallback.
 */

const { execFile } = require("child_process");
const { promisify } = require("util");
const { getSapSnapshotVbsPath } = require("../../../paths");
const { deferSnapshotDebug } = require("../../captureLog");

const execFileAsync = promisify(execFile);
// The VBScript enforces its own ~7s wall-clock budget and returns partial
// results; this host timeout sits above it so cscript is only hard-killed if
// the bridge itself hangs (COM stall), not during normal deep discovery.
const VBS_TIMEOUT_MS = 10000;

function emptySnapshot(error = "") {
	return {
		timestamp: new Date().toISOString(),
		transaction: "",
		program: "",
		screen: "",
		controls: [],
		...(error ? { error } : {}),
	};
}

function normalizeSnapshot(parsed) {
	if (!parsed || typeof parsed !== "object") return emptySnapshot("Invalid snapshot response");
	return {
		timestamp: parsed.timestamp || new Date().toISOString(),
		transaction: parsed.transaction || "",
		program: parsed.program || "",
		screen: String(parsed.screen || ""),
		controls: Array.isArray(parsed.controls) ? parsed.controls : [],
		captureSource: parsed.captureSource || "",
		sapWindowRect: parsed.sapWindowRect || null,
		debug: parsed.debug || null,
		...(parsed.error ? { error: parsed.error } : {}),
	};
}

function isVbsTimeoutError(message) {
	return /timed out|ETIMEDOUT|SIGTERM|kill|ENOENT.*cscript/i.test(String(message || ""));
}

async function captureSnapshot(options = {}) {
	if (process.platform !== "win32") {
		return emptySnapshot("SAP snapshot capture requires Windows");
	}

	let maxFields = Number(options.maxFields);
	if (!Number.isFinite(maxFields) || maxFields < 1) maxFields = 10;
	if (maxFields > 50) maxFields = 50;

	const vbsPath = getSapSnapshotVbsPath();
	try {
		const { stdout } = await execFileAsync(
			"cscript",
			["//Nologo", vbsPath, String(maxFields)],
			{ encoding: "utf8", windowsHide: true, timeout: VBS_TIMEOUT_MS, maxBuffer: 4 * 1024 * 1024 }
		);
		const raw = (stdout || "").trim();
		if (!raw) {
			const snapshot = emptySnapshot("VBScript snapshot returned empty output");
			deferSnapshotDebug(snapshot);
			return snapshot;
		}
		const snapshot = normalizeSnapshot(JSON.parse(raw));
		if (snapshot.debug?.sessionConnected !== true && !snapshot.controls.length) {
			snapshot.error = snapshot.error || "SAP scripting not connected";
		}
		deferSnapshotDebug(snapshot);
		return snapshot;
	} catch (err) {
		const msg = err && err.message ? err.message : String(err);
		if (isVbsTimeoutError(msg)) {
			const snapshot = emptySnapshot("SAP capture timed out — keep SAP focused and try again");
			snapshot.debug = { bridge: "vbscript", sessionConnected: false, timedOut: true };
			deferSnapshotDebug(snapshot, { vbsError: msg });
			return snapshot;
		}
		const snapshot = emptySnapshot(msg);
		deferSnapshotDebug(snapshot, { vbsError: msg });
		return snapshot;
	}
}

module.exports = { captureSnapshot, emptySnapshot, normalizeSnapshot };
