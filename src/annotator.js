/**
 * User-friendly screenshot annotator for SAP SOP steps.
 */

let _canvas = null;
let _ctx = null;
let _baseImage = null;
let _shapes = [];
let _tool = "rect";
let _color = "#f97316";
let _lineWidth = 4;
let _drawing = false;
let _start = null;
let _preview = null;
let _imagePath = "";
let _stepIndex = -1;
let _onSave = null;
let _nextNumber = 1;
let _textPendingPoint = null;

const TOOL_HINTS = {
	rect: "Click and drag to draw a box around a button or field",
	arrow: "Click and drag from the label toward the item you mean",
	highlight: "Click and drag to highlight an area (yellow overlay)",
	number: "Click once to place a step number badge",
	text: "Click where you want a text label, then type below",
};

const COLOR_PRESETS = [
	{ id: "red", value: "#e53935", label: "Red" },
	{ id: "blue", value: "#1e88e5", label: "Blue" },
	{ id: "orange", value: "#fb8c00", label: "Orange" },
	{ id: "green", value: "#43a047", label: "Green" },
];

function $(id) {
	return document.getElementById(id);
}

function canvasPoint(evt) {
	const rect = _canvas.getBoundingClientRect();
	const scaleX = _canvas.width / rect.width;
	const scaleY = _canvas.height / rect.height;
	return {
		x: (evt.clientX - rect.left) * scaleX,
		y: (evt.clientY - rect.top) * scaleY,
		screenX: evt.clientX,
		screenY: evt.clientY,
	};
}

function scaledLineWidth() {
	if (!_canvas) return _lineWidth;
	return Math.max(3, Math.round(_canvas.width / 400));
}

function scaledNumberSize() {
	if (!_canvas) return 14;
	return Math.max(11, Math.round(_canvas.width / 130));
}

function redraw() {
	if (!_ctx || !_baseImage) return;
	_ctx.clearRect(0, 0, _canvas.width, _canvas.height);
	_ctx.drawImage(_baseImage, 0, 0);
	for (const shape of _shapes) drawShape(shape);
	if (_preview) drawShape(_preview);
}

function drawShape(shape) {
	const lw = scaledLineWidth();
	const fs = scaledFontSize();
	const numSize = scaledNumberSize();
	const { type, x1, y1, x2, y2, color, text, number } = shape;
	_ctx.save();
	if (type === "highlight") {
		_ctx.fillStyle = color || "rgba(255, 235, 59, 0.42)";
		_ctx.fillRect(x1, y1, x2 - x1, y2 - y1);
	} else if (type === "rect") {
		_ctx.strokeStyle = color || "#e53935";
		_ctx.lineWidth = lw;
		_ctx.strokeRect(x1, y1, x2 - x1, y2 - y1);
	} else if (type === "arrow") {
		drawArrow(x1, y1, x2, y2, color || "#e53935", lw);
	} else if (type === "number") {
		drawNumberBadge(x1, y1, number || 1, color || "#e53935", numSize);
	} else if (type === "text" && text) {
		drawLabel(x1, y1, text, color || "#e53935", fs);
	}
	_ctx.restore();
}

function drawArrow(x1, y1, x2, y2, color, lw) {
	const head = Math.max(12, lw * 4);
	const angle = Math.atan2(y2 - y1, x2 - x1);
	_ctx.strokeStyle = color;
	_ctx.fillStyle = color;
	_ctx.lineWidth = lw;
	_ctx.lineCap = "round";
	_ctx.beginPath();
	_ctx.moveTo(x1, y1);
	_ctx.lineTo(x2, y2);
	_ctx.stroke();
	_ctx.beginPath();
	_ctx.moveTo(x2, y2);
	_ctx.lineTo(x2 - head * Math.cos(angle - Math.PI / 6), y2 - head * Math.sin(angle - Math.PI / 6));
	_ctx.lineTo(x2 - head * Math.cos(angle + Math.PI / 6), y2 - head * Math.sin(angle + Math.PI / 6));
	_ctx.closePath();
	_ctx.fill();
}

function drawNumberBadge(x, y, num, color, fs) {
	const r = fs * 0.72;
	_ctx.fillStyle = color;
	_ctx.beginPath();
	_ctx.arc(x, y, r, 0, Math.PI * 2);
	_ctx.fill();
	_ctx.fillStyle = "#fff";
	_ctx.font = `bold ${Math.max(10, Math.round(fs * 0.82))}px Arial`;
	_ctx.textAlign = "center";
	_ctx.textBaseline = "middle";
	_ctx.fillText(String(num), x, y + 1);
}

function scaledFontSize() {
	if (!_canvas) return 20;
	return Math.max(16, Math.round(_canvas.width / 70));
}

function drawLabel(x, y, text, color, fs) {
	_ctx.font = `bold ${fs}px Arial`;
	const pad = 8;
	const metrics = _ctx.measureText(text);
	const w = metrics.width + pad * 2;
	const h = fs + pad;
	_ctx.fillStyle = "rgba(255,255,255,0.92)";
	_ctx.strokeStyle = color;
	_ctx.lineWidth = 2;
	_ctx.fillRect(x, y - h + 4, w, h);
	_ctx.strokeRect(x, y - h + 4, w, h);
	_ctx.fillStyle = color;
	_ctx.textAlign = "left";
	_ctx.textBaseline = "bottom";
	_ctx.fillText(text, x + pad, y + 4);
}

function setTool(tool) {
	_tool = tool;
	hideTextEditor();
	document.querySelectorAll(".annot-tool").forEach((btn) => {
		btn.classList.toggle("active", btn.dataset.tool === tool);
	});
	const hint = $("annotToolHint");
	if (hint) hint.textContent = TOOL_HINTS[tool] || "";
	if (_canvas) {
		_canvas.style.cursor = tool === "text" || tool === "number" ? "pointer" : "crosshair";
	}
}

function setColor(color) {
	_color = color;
	const colorInput = $("annotColor");
	if (colorInput) colorInput.value = color;
	document.querySelectorAll(".annot-color").forEach((btn) => {
		btn.classList.toggle("active", btn.dataset.color === color);
	});
}

function updateMeta(options = {}) {
	const titleEl = $("annotStepTitle");
	const subEl = $("annotStepSub");
	if (titleEl) {
		const n = options.stepNumber ? `Step ${options.stepNumber}` : "Annotate screenshot";
		titleEl.textContent = options.title ? `${n}: ${options.title}` : n;
	}
	if (subEl) {
		subEl.textContent = options.subtitle || "Mark up the screenshot, then click Done — or Skip if no markup is needed.";
	}
}

function hideTextEditor() {
	const editor = $("annotTextEditor");
	if (editor) editor.style.display = "none";
	_textPendingPoint = null;
}

function showTextEditor(screenX, screenY, canvasPoint) {
	const editor = $("annotTextEditor");
	const input = $("annotTextInput");
	if (!editor || !input) return;
	_textPendingPoint = canvasPoint;
	editor.style.display = "flex";
	editor.style.left = Math.min(screenX, window.innerWidth - 320) + "px";
	editor.style.top = Math.min(screenY + 8, window.innerHeight - 80) + "px";
	input.value = "";
	setTimeout(() => { input.focus(); }, 30);
}

function confirmTextLabel() {
	const input = $("annotTextInput");
	if (!input || !_textPendingPoint) return;
	const text = input.value.trim();
	if (text) {
		_shapes.push({
			type: "text",
			x1: _textPendingPoint.x,
			y1: _textPendingPoint.y,
			color: _color,
			text,
		});
		redraw();
	}
	hideTextEditor();
}

function bindCanvasEvents() {
	if (!_canvas || _canvas.dataset.bound === "1") return;
	_canvas.dataset.bound = "1";

	const onPointerDown = (evt) => {
		evt.preventDefault();
		_canvas.setPointerCapture(evt.pointerId);
		const p = canvasPoint(evt);

		if (_tool === "text") {
			showTextEditor(p.screenX, p.screenY, p);
			return;
		}
		if (_tool === "number") {
			_shapes.push({
				type: "number",
				x1: p.x,
				y1: p.y,
				number: _nextNumber++,
				color: _color,
			});
			redraw();
			return;
		}

		_drawing = true;
		_start = p;
		_preview = null;
	};

	const onPointerMove = (evt) => {
		if (!_drawing || !_start) return;
		const p = canvasPoint(evt);
		if (_tool === "arrow") {
			_preview = { type: "arrow", x1: _start.x, y1: _start.y, x2: p.x, y2: p.y, color: _color };
		} else {
			const x1 = Math.min(_start.x, p.x);
			const y1 = Math.min(_start.y, p.y);
			const x2 = Math.max(_start.x, p.x);
			const y2 = Math.max(_start.y, p.y);
			_preview = {
				type: _tool,
				x1, y1, x2, y2,
				color: _tool === "highlight" ? "rgba(255, 235, 59, 0.42)" : _color,
			};
		}
		redraw();
	};

	const onPointerUp = (evt) => {
		if (!_drawing || !_start) return;
		_drawing = false;
		const p = canvasPoint(evt);

		if (_tool === "arrow") {
			if (Math.hypot(p.x - _start.x, p.y - _start.y) > 8) {
				_shapes.push({ type: "arrow", x1: _start.x, y1: _start.y, x2: p.x, y2: p.y, color: _color });
			}
		} else {
			const x1 = Math.min(_start.x, p.x);
			const y1 = Math.min(_start.y, p.y);
			const x2 = Math.max(_start.x, p.x);
			const y2 = Math.max(_start.y, p.y);
			if (Math.abs(x2 - x1) > 6 || Math.abs(y2 - y1) > 6) {
				_shapes.push({
					type: _tool,
					x1, y1, x2, y2,
					color: _tool === "highlight" ? "rgba(255, 235, 59, 0.42)" : _color,
				});
			}
		}
		_preview = null;
		_start = null;
		redraw();
		try { _canvas.releasePointerCapture(evt.pointerId); } catch { /* ignore */ }
	};

	_canvas.addEventListener("pointerdown", onPointerDown);
	_canvas.addEventListener("pointermove", onPointerMove);
	_canvas.addEventListener("pointerup", onPointerUp);
	_canvas.addEventListener("pointercancel", onPointerUp);
}

function fitCanvasDisplay() {
	if (!_canvas || !_baseImage) return;
	const wrap = _canvas.parentElement;
	if (!wrap) return;
	const maxW = wrap.clientWidth - 16;
	const maxH = wrap.clientHeight - 16;
	if (maxW <= 0 || maxH <= 0) return;
	const scale = Math.min(maxW / _canvas.width, maxH / _canvas.height, 1);
	_canvas.style.width = Math.round(_canvas.width * scale) + "px";
	_canvas.style.height = Math.round(_canvas.height * scale) + "px";
}

function openAnnotator(imagePath, stepIndex, onSave, meta = {}) {
	_imagePath = imagePath;
	_stepIndex = stepIndex;
	_onSave = onSave;
	_shapes = [];
	_preview = null;
	_nextNumber = meta.stepNumber || 1;
	hideTextEditor();
	updateMeta(meta);

	const modal = $("annotatorModal");
	const img = new Image();
	img.onload = () => {
		_baseImage = img;
		_canvas = $("annotCanvas");
		_ctx = _canvas.getContext("2d");
		_canvas.width = img.naturalWidth;
		_canvas.height = img.naturalHeight;
		_canvas.style.width = "";
		_canvas.style.height = "";
		bindCanvasEvents();
		redraw();
		modal.classList.add("open");
		setTool("rect");
		setTimeout(fitCanvasDisplay, 80);
		window.addEventListener("resize", fitCanvasDisplay);
	};
	img.onerror = () => alert("Could not load screenshot for annotation.");
	try {
		const { pathToFileURL } = require("url");
		img.src = pathToFileURL(imagePath).href;
	} catch {
		img.src = imagePath;
	}
}

function closeAnnotator() {
	hideTextEditor();
	window.removeEventListener("resize", fitCanvasDisplay);
	const modal = $("annotatorModal");
	if (modal) modal.classList.remove("open");
	_shapes = [];
	_preview = null;
	_baseImage = null;
}

function saveAnnotation() {
	if (!_canvas || !_imagePath) return;
	hideTextEditor();
	redraw();
	const fs = require("fs");
	const dataUrl = _canvas.toDataURL("image/png");
	const base64 = dataUrl.replace(/^data:image\/png;base64,/, "");
	fs.writeFileSync(_imagePath, Buffer.from(base64, "base64"));
	if (typeof _onSave === "function") _onSave(_stepIndex, _imagePath);
	closeAnnotator();
}

function undoAnnotation() {
	if (_shapes.length) {
		const last = _shapes[_shapes.length - 1];
		if (last.type === "number" && last.number === _nextNumber - 1) _nextNumber--;
		_shapes.pop();
		redraw();
	}
}

function clearAnnotations() {
	_shapes = [];
	_nextNumber = 1;
	redraw();
}

function setupKeyboardShortcuts() {
	document.addEventListener("keydown", (e) => {
		if (!document.body.classList.contains("annotating")) return;
		if (e.target.id === "annotTextInput") {
			if (e.key === "Enter") { e.preventDefault(); confirmTextLabel(); }
			if (e.key === "Escape") { e.preventDefault(); hideTextEditor(); }
			return;
		}
		if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") {
			e.preventDefault();
			undoAnnotation();
		}
		if (e.key === "1") setTool("rect");
		if (e.key === "2") setTool("arrow");
		if (e.key === "3") setTool("highlight");
		if (e.key === "4") setTool("number");
		if (e.key === "5") setTool("text");
	});
}

function initAnnotator() {
	document.querySelectorAll(".annot-tool").forEach((btn) => {
		btn.addEventListener("click", () => setTool(btn.dataset.tool));
	});
	document.querySelectorAll(".annot-color").forEach((btn) => {
		btn.addEventListener("click", () => setColor(btn.dataset.color));
	});
	const colorInput = $("annotColor");
	if (colorInput) {
		colorInput.addEventListener("input", (e) => setColor(e.target.value));
	}
	$("annotTextOk")?.addEventListener("click", confirmTextLabel);
	$("annotTextCancel")?.addEventListener("click", hideTextEditor);
	_canvas = $("annotCanvas");
	setColor(_color);
	setTool("rect");
	setupKeyboardShortcuts();
}

module.exports = {
	initAnnotator,
	openAnnotator,
	closeAnnotator,
	saveAnnotation,
	undoAnnotation,
	clearAnnotations,
};
