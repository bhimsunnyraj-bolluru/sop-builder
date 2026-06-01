/**
 * SOP Session Recorder — snapshot diff pipeline integrated with existing project steps.
 */

const { compareSnapshots, collectFilledFields } = require("../sap/deltaEngine");
const { resolveChanges } = require("../sap/labelResolver");
const { formatStepDescription, isSapSuccessStatusBar, isSapForegroundContext, resolveForegroundTitle } = require("../../sapgui-context");

let _previousSnapshot = null;
let _stepNumber = 0;

function resetSessionRecorder() {
	_previousSnapshot = null;
	_stepNumber = 0;
}

function collectFilledFromControls(controls) {
	return (controls || [])
		.filter((c) => String(c.value ?? c.text ?? "").trim())
		.map((c) => ({
			fieldId: c.name || c.id,
			old: "",
			new: String(c.value ?? c.text ?? "").trim(),
		}));
}

function formatChangePhrase(change) {
	const label = String(change.label || change.technicalName || "Field").trim();
	if (change.oldValue && change.newValue) {
		return `Set ${label}: ${change.oldValue} → ${change.newValue}`;
	}
	return `Set ${label}: ${change.newValue}`;
}

/** True when step.description already includes every formatted change (avoid duplicate UI/export). */
function descriptionListsChanges(description, changes) {
	const desc = String(description || "").trim();
	if (!desc || !Array.isArray(changes) || !changes.length) return false;
	return changes.every((c) => desc.includes(formatChangePhrase(c)));
}

function buildStepDescription(changes, snapshot, context) {
	const statusBar = context && context.statusBar ? String(context.statusBar).trim() : "";
	if (isSapSuccessStatusBar(statusBar)) return statusBar;

	if (changes && changes.length) {
		return changes.map(formatChangePhrase).join("; ");
	}

	const ctxDesc = context ? formatStepDescription(context) : "";
	if (ctxDesc) return ctxDesc;

	const title = snapshot && snapshot.transaction ? `${snapshot.transaction} screen ${snapshot.screen || ""}`.trim() : "";
	return title || (context && resolveForegroundTitle(context)) || "Captured step";
}

function shouldResetBaseline(previous, snapshot, context) {
	if (!previous) return false;
	const prevTxn = String(previous.transaction || "");
	const prevScreen = String(previous.screen || "");
	const nextTxn = String(snapshot.transaction || "");
	const nextScreen = String(snapshot.screen || "");
	if (prevTxn && nextTxn && prevTxn !== nextTxn) return true;
	if (prevScreen && nextScreen && prevScreen !== nextScreen) return true;
	// Do not diff SAP fields when the new step is non-SAP (browser, calculator, etc.)
	if (context && !isSapForegroundContext(context) && (prevTxn || (previous.controls && previous.controls.length))) {
		return true;
	}
	return false;
}

function recordStep(input) {
	const screenshot = input && input.screenshot ? input.screenshot : "";
	const snapshot = input && input.snapshot ? input.snapshot : { controls: [] };
	const context = input && input.context ? input.context : null;
	const statusBar = context && context.statusBar ? String(context.statusBar).trim() : "";
	const statusBarOnly = isSapSuccessStatusBar(statusBar);
	const sapCapture = isSapForegroundContext(context);

	if (shouldResetBaseline(_previousSnapshot, snapshot, context)) {
		_previousSnapshot = null;
	}

	let rawChanges = [];
	if (sapCapture && !statusBarOnly) {
		if (_previousSnapshot) {
			rawChanges = compareSnapshots(_previousSnapshot, snapshot);
		} else {
			rawChanges = collectFilledFields(snapshot);
		}
		if (!rawChanges.length && Array.isArray(snapshot.controls) && snapshot.controls.length) {
			rawChanges = collectFilledFromControls(snapshot.controls);
		}
	}

	const baselineSnapshot =
		statusBarOnly || !sapCapture ? { ...snapshot, controls: [] } : snapshot;
	_previousSnapshot = sapCapture ? baselineSnapshot : null;
	_stepNumber += 1;

	const changes = sapCapture && !statusBarOnly ? resolveChanges(snapshot, rawChanges) : [];
	const controlCount =
		sapCapture && !statusBarOnly && Array.isArray(snapshot.controls) ? snapshot.controls.length : 0;

	const step = {
		stepNumber: _stepNumber,
		timestamp: snapshot.timestamp || new Date().toISOString(),
		screenshot,
		image: screenshot,
		transaction: sapCapture ? snapshot.transaction || (context && context.transaction) || "" : "",
		screen: sapCapture ? snapshot.screen || (context && context.screenNumber) || "" : "",
		program: sapCapture ? snapshot.program || (context && context.program) || "" : "",
		changes,
		description: buildStepDescription(changes, snapshot, context),
		sapContext: context || undefined,
		snapshotMeta: {
			transaction: sapCapture ? snapshot.transaction || "" : "",
			program: sapCapture ? snapshot.program || "" : "",
			screen: sapCapture ? snapshot.screen || "" : "",
			controlCount,
			captureSource: sapCapture ? snapshot.captureSource || "" : "win32-foreground",
			statusBarOnly,
			foregroundIsSap: !!sapCapture,
			...(snapshot.error && sapCapture ? { error: snapshot.error } : {}),
		},
		sessionRecord: true,
	};

	return step;
}

function getPreviousSnapshot() {
	return _previousSnapshot;
}

function getStepNumber() {
	return _stepNumber;
}

module.exports = {
	recordStep,
	resetSessionRecorder,
	buildStepDescription,
	formatChangePhrase,
	descriptionListsChanges,
	getPreviousSnapshot,
	getStepNumber,
};
