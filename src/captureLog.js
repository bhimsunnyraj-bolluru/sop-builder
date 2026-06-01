/**
 * Append-only capture / snapshot diagnostics log (data/capture.log).
 */

const fs = require("fs");
const path = require("path");
const { getDataDir, ensureDir } = require("../paths");

function getLogPath() {
	return path.join(getDataDir(), "capture.log");
}

function getSnapshotDebugPath() {
	return path.join(getDataDir(), "snapshot-last.json");
}

function appendCaptureLog(entry) {
	try {
		ensureDir(getDataDir());
		const line = JSON.stringify({
			timestamp: new Date().toISOString(),
			...entry,
		});
		fs.appendFileSync(getLogPath(), line + "\n", "utf8");
	} catch {
		/* ignore log write failures */
	}
}

function writeSnapshotDebug(snapshot, extra = {}) {
	try {
		ensureDir(getDataDir());
		const controls = Array.isArray(snapshot && snapshot.controls) ? snapshot.controls : [];
		const sampleFields = controls
			.filter((c) => String(c.value ?? c.text ?? "").trim() || c.label)
			.slice(0, 20)
			.map((c) => ({
				name: c.name,
				label: c.label || "",
				value: String(c.value ?? c.text ?? "").trim(),
			}));

		const payload = {
			timestamp: new Date().toISOString(),
			transaction: snapshot.transaction || "",
			program: snapshot.program || "",
			screen: snapshot.screen || "",
			controlCount: controls.length,
			filledFieldCount: controls.filter((c) => String(c.value ?? c.text ?? "").trim()).length,
			captureSource: snapshot.captureSource || "",
			error: snapshot.error || "",
			debug: snapshot.debug || null,
			sampleFields,
			sessionAttach: snapshot.debug?.sessionAttach || null,
			vbsDiscoverMs: snapshot.debug?.discoverMs ?? undefined,
			...extra,
		};
		fs.writeFileSync(getSnapshotDebugPath(), JSON.stringify(payload, null, 2), "utf8");
		appendCaptureLog({
			event: "snapshot",
			controlCount: payload.controlCount,
			filledFieldCount: payload.filledFieldCount,
			captureSource: payload.captureSource,
			error: payload.error || undefined,
			transaction: payload.transaction || undefined,
		});
		return payload;
	} catch {
		return null;
	}
}

function deferSnapshotDebug(snapshot, extra = {}) {
	setImmediate(() => writeSnapshotDebug(snapshot, extra));
}

module.exports = {
	appendCaptureLog,
	writeSnapshotDebug,
	deferSnapshotDebug,
	getLogPath,
	getSnapshotDebugPath,
};
