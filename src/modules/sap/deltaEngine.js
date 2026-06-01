/**
 * Compares two SAP GUI snapshots and returns only meaningful field changes.
 */

const TECHNICAL_TYPES = new Set([
	"GuiMenubar",
	"GuiToolbar",
	"GuiTitlebar",
	"GuiStatusPane",
	"GuiOkCodeField",
	"GuiNavigationPanel",
	"GuiSplitter",
	"GuiScrollContainer",
	"GuiContainerShell",
]);

const TECHNICAL_ID_RE = /(?:\/sbar|\/tbar|\/mbar|\/toolbar|\/okcd|\/titl)/i;
const TECHNICAL_NAME_RE = /^(OK|%)?CODE$/i;
const LABEL_TYPE_RE = /Label/i;
const EDITABLE_TYPE_RE = /TextField|CTextField|ComboBox|CheckBox|RadioButton|PasswordField/i;

function extractTechnicalName(id) {
	const m = String(id || "").match(/(?:ctxt|txt|cmb|chk|btn|rdo|pwd)([^/[\]]+)/i);
	return m ? m[1] : "";
}

function isLabelControl(control) {
	return LABEL_TYPE_RE.test(String(control.type || ""));
}

function isEditableControl(control) {
	const type = String(control.type || "");
	const id = String(control.id || "");
	if (EDITABLE_TYPE_RE.test(type)) return true;
	if (/(?:ctxt|\/txt|cmb|chk|pwd)/i.test(id)) return true;
	if (/\/cell\[/i.test(id)) return true;
	if (/GridViewCell/i.test(type)) return true;
	return false;
}

function isTechnicalControl(control) {
	if (!control) return true;
	const type = String(control.type || "");
	if (TECHNICAL_TYPES.has(type)) return true;
	if (/^GuiShell$/i.test(type)) return true;
	const id = String(control.id || "");
	const name = String(control.name || "");
	if (TECHNICAL_ID_RE.test(id)) return true;
	if (TECHNICAL_NAME_RE.test(name)) return true;
	return false;
}

function fieldKey(control) {
	const id = String(control.id || "").trim();
	const bracket = id.match(/\[(\d+),(\d+)\]/);
	const fromId = extractTechnicalName(id);
	if (fromId && bracket) return `${fromId}[${bracket[1]},${bracket[2]}]`;
	if (/\/cell\[/i.test(id)) {
		const name = String(control.name || "").trim();
		if (name && bracket) return `${name}[${bracket[1]},${bracket[2]}]`;
		if (name) return name;
	}
	if (fromId) return fromId;
	const name = String(control.name || "").trim();
	if (name) return name;
	return id;
}

function fieldValue(control) {
	return String(control.value ?? control.text ?? "").trim();
}

function compareSnapshots(before, after) {
	const beforeMap = new Map();
	const beforeControls = before && Array.isArray(before.controls) ? before.controls : [];

	for (const control of beforeControls) {
		if (isTechnicalControl(control) || isLabelControl(control)) continue;
		if (!isEditableControl(control)) continue;
		const key = fieldKey(control);
		if (!key) continue;
		beforeMap.set(key, fieldValue(control));
	}

	const changes = [];
	const afterControls = after && Array.isArray(after.controls) ? after.controls : [];
	const seen = new Set();

	for (const control of afterControls) {
		if (isTechnicalControl(control) || isLabelControl(control)) continue;
		if (!isEditableControl(control)) continue;
		const key = fieldKey(control);
		if (!key || seen.has(key)) continue;
		seen.add(key);

		const newVal = fieldValue(control);
		const oldVal = beforeMap.has(key) ? beforeMap.get(key) : "";

		if (newVal === oldVal) continue;
		if (!newVal && !oldVal) continue;

		changes.push({
			fieldId: key,
			old: oldVal,
			new: newVal,
		});
	}

	return changes;
}

/** First capture (no baseline): treat all filled fields as recorded values. */
function collectFilledFields(snapshot) {
	const changes = [];
	const controls = snapshot && Array.isArray(snapshot.controls) ? snapshot.controls : [];
	const seen = new Set();

	for (const control of controls) {
		if (isTechnicalControl(control) || isLabelControl(control)) continue;
		if (!isEditableControl(control)) continue;
		const key = fieldKey(control);
		const newVal = fieldValue(control);
		if (!key || !newVal || seen.has(key)) continue;
		seen.add(key);
		changes.push({ fieldId: key, old: "", new: newVal });
	}

	return changes;
}

module.exports = {
	compareSnapshots,
	collectFilledFields,
	isTechnicalControl,
	isLabelControl,
	isEditableControl,
	fieldKey,
	fieldValue,
	extractTechnicalName,
};
