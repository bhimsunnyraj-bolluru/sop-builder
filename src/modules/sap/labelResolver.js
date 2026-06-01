/**
 * Labels come from VBS capture (DefaultTooltip → control.label). No pairing or cache.
 */

const { extractTechnicalName } = require("./deltaEngine");

function fieldLabel(control) {
	const label = String(control?.label || control?.defaultTooltip || "").trim();
	if (label) return label.replace(/[:：]\s*$/, "");
	const name = extractTechnicalName(control?.id || control?.name || "") || control?.name || "";
	return name.replace(/_/g, " ").trim() || name;
}

function buildControlIndex(controls) {
	const byKey = new Map();
	for (const c of controls || []) {
		const keys = [
			c.id,
			c.name,
			extractTechnicalName(c.id || ""),
			extractTechnicalName(c.name || ""),
		].filter(Boolean);
		for (const k of keys) {
			if (!byKey.has(k)) byKey.set(k, c);
		}
	}
	return byKey;
}

function resolveChanges(snapshot, rawChanges) {
	const byKey = buildControlIndex(snapshot?.controls);
	return (rawChanges || []).map((change) => {
		const fieldId = String(change.fieldId || "");
		const tech = extractTechnicalName(fieldId) || fieldId;
		const ctrl = byKey.get(fieldId) || byKey.get(tech);
		return {
			label: ctrl ? fieldLabel(ctrl) : tech,
			technicalName: tech,
			oldValue: change.old ?? "",
			newValue: change.new ?? "",
		};
	});
}

function resolveLabels(controls) {
	return (controls || [])
		.filter((c) => String(c.value ?? c.text ?? "").trim())
		.map((c) => ({
			technicalName: extractTechnicalName(c.id || c.name || "") || c.name || "",
			label: fieldLabel(c),
			value: String(c.value ?? c.text ?? "").trim(),
		}));
}

function clearLabelCache() {}

module.exports = {
	resolveChanges,
	resolveLabels,
	fieldLabel,
	clearLabelCache,
};
