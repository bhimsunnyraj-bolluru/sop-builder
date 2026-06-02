const {
	Document,
	Packer,
	Paragraph,
	ImageRun,
	LevelFormat,
	AlignmentType,
	HeadingLevel,
	TextRun,
	LevelSuffix,
	Header,
	Footer,
	PageNumber,
	PageBreak,
	Table,
	TableRow,
	TableCell,
	WidthType,
	BorderStyle,
	VerticalAlign,
	TabStopType,
} = require("docx");
const fsExporter = require("fs");
const path = require("path");
const sizeOf = require("image-size");
const { getExportsDir, ensureDir } = require("../paths");
const { DEFAULT_BRANDING } = require("./config");
const { formatChangePhrase, descriptionListsChanges } = require("./modules/recording/sessionRecorder");

/** Width of the printable area for Letter @ 1" margins, in twips. Used for tab stops/table widths. */
const CONTENT_WIDTH_TWIPS = 9360;
const LIGHT_SHADE = "F2F2F2";
const TEXT_MUTED = "666666";

function sanitizeFileName(name) {
	return name.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").slice(0, 120);
}

/** Strip a leading '#' and upper-case a hex color; fall back to the default accent on bad input. */
function normalizeHex(value, fallback = "F97316") {
	const hex = String(value || "").replace(/^#/, "").trim().toUpperCase();
	return /^[0-9A-F]{6}$/.test(hex) ? hex : fallback.replace(/^#/, "").toUpperCase();
}

function resolveBranding(branding) {
	const b = { ...DEFAULT_BRANDING, ...(branding || {}) };
	b.accentColor = normalizeHex(b.accentColor, DEFAULT_BRANDING.accentColor);
	return b;
}

function imageType(imagePath) {
	const ext = path.extname(imagePath).toLowerCase();
	if (ext === ".jpg" || ext === ".jpeg") return "jpg";
	if (ext === ".gif") return "gif";
	if (ext === ".bmp") return "bmp";
	return "png";
}

/** Build an ImageRun for a logo/screenshot scaled to maxWidth (keeping aspect ratio), or null. */
function buildImageRun(imagePath, maxWidth, maxHeight) {
	if (!imagePath || !fsExporter.existsSync(imagePath)) return null;
	let stat;
	try {
		stat = fsExporter.statSync(imagePath);
	} catch {
		return null;
	}
	if (!stat.isFile()) return null;

	let width = maxWidth;
	let height = Math.round(maxWidth * 0.6);
	try {
		const dims = sizeOf(imagePath);
		width = dims.width || maxWidth;
		height = dims.height || height;
	} catch {
		/* use defaults */
	}
	if (width > maxWidth) {
		const ratio = maxWidth / width;
		width = Math.round(width * ratio);
		height = Math.round(height * ratio);
	}
	if (maxHeight && height > maxHeight) {
		const ratio = maxHeight / height;
		width = Math.round(width * ratio);
		height = Math.round(height * ratio);
	}

	let data;
	try {
		data = fsExporter.readFileSync(imagePath);
	} catch {
		return null;
	}
	return new ImageRun({ type: imageType(imagePath), data, transformation: { width, height } });
}

/** Split "Set A: x; Set B: y" into Word lines (semicolon on all but last). Returns null if not splittable. */
function splitDescriptionLinesForWord(description) {
	const text = String(description || "").trim();
	if (!text.includes(";")) return null;
	const parts = text
		.split(";")
		.map((p) => p.trim())
		.filter(Boolean);
	if (parts.length <= 1) return null;
	return parts.map((p, i) => (i < parts.length - 1 ? `${p};` : p));
}

function stepDescriptionParagraphs(description) {
	const lines = splitDescriptionLinesForWord(description);
	if (!lines) {
		return [
			new Paragraph({
				children: [new TextRun({ text: description || " " })],
				numbering: { reference: "stepNumbering", level: 0 },
				spacing: { after: 80 },
			}),
		];
	}

	const paragraphs = [
		new Paragraph({
			children: [new TextRun({ text: "" })],
			numbering: { reference: "stepNumbering", level: 0 },
			spacing: { after: 40 },
		}),
	];
	for (let i = 0; i < lines.length; i++) {
		paragraphs.push(
			new Paragraph({
				children: [new TextRun({ text: lines[i] })],
				indent: { left: 720 },
				spacing: { after: i === lines.length - 1 ? 80 : 40 },
			})
		);
	}
	return paragraphs;
}

function formatDisplayDate(value) {
	if (!value) return "";
	try {
		const d = new Date(value);
		if (Number.isNaN(d.getTime())) return value;
		return d.toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
	} catch {
		return value;
	}
}

function imageParagraph(imagePath) {
	const run = buildImageRun(imagePath, 600, 760);
	if (!run) return null;
	return new Paragraph({ children: [run], spacing: { after: 240 } });
}

/* ------------------------------------------------------------------ *
 *  Branding building blocks: cover page, tables, headers, footers
 * ------------------------------------------------------------------ */

function noBorders() {
	const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
	return { top: none, bottom: none, left: none, right: none };
}

/** Two-column label/value table used for both Document Control and Revision History. */
function infoTable(rows, accent) {
	const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD" };
	const borders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };

	const tableRows = rows.map(
		(r) =>
			new TableRow({
				children: [
					new TableCell({
						width: { size: 32, type: WidthType.PERCENTAGE },
						shading: { fill: LIGHT_SHADE },
						margins: { top: 60, bottom: 60, left: 120, right: 120 },
						verticalAlign: VerticalAlign.CENTER,
						children: [
							new Paragraph({
								children: [new TextRun({ text: r.label, bold: true, color: "333333" })],
							}),
						],
					}),
					new TableCell({
						width: { size: 68, type: WidthType.PERCENTAGE },
						margins: { top: 60, bottom: 60, left: 120, right: 120 },
						verticalAlign: VerticalAlign.CENTER,
						children: [
							new Paragraph({
								children: [new TextRun({ text: r.value == null ? "" : String(r.value) })],
							}),
						],
					}),
				],
			})
	);

	return new Table({
		width: { size: 100, type: WidthType.PERCENTAGE },
		borders,
		rows: tableRows,
	});
}

function sectionHeading(text, accent) {
	return new Paragraph({
		spacing: { before: 240, after: 120 },
		border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: accent } },
		children: [new TextRun({ text, bold: true, size: 26, color: accent })],
	});
}

function buildCoverChildren(project, branding, accent) {
	const children = [];

	// Push content down a little so the cover breathes.
	children.push(new Paragraph({ text: "", spacing: { before: 600 } }));

	const logoRun = buildImageRun(branding.logoPath, 200, 120);
	if (logoRun) {
		children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [logoRun], spacing: { after: 160 } }));
	}

	if (branding.companyName) {
		children.push(
			new Paragraph({
				alignment: AlignmentType.CENTER,
				spacing: { after: 80 },
				children: [new TextRun({ text: branding.companyName, bold: true, size: 30, color: accent })],
			})
		);
	}

	children.push(
		new Paragraph({
			alignment: AlignmentType.CENTER,
			spacing: { after: 480 },
			children: [new TextRun({ text: "Standard Operating Procedure", size: 22, color: TEXT_MUTED, allCaps: true })],
		})
	);

	children.push(
		new Paragraph({
			alignment: AlignmentType.CENTER,
			spacing: { before: 240, after: 120 },
			border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: accent } },
			children: [new TextRun({ text: project.title || "Untitled SOP", bold: true, size: 48 })],
		})
	);

	children.push(new Paragraph({ text: "", spacing: { after: 360 } }));

	// Document Control table.
	const controlRows = [
		{ label: "Author", value: project.author },
		{ label: "Version", value: project.version },
		{ label: "Review Date", value: formatDisplayDate(project.reviewDate) },
		{ label: "Document ID", value: project.documentId },
		{ label: "Department", value: project.department },
		{ label: "Classification", value: project.classification },
	].filter((r) => r.value);

	if (controlRows.length) {
		children.push(sectionHeading("Document Control", accent));
		children.push(infoTable(controlRows, accent));
	}

	// Detailed template adds a revision history table.
	if (branding.template === "detailed") {
		children.push(sectionHeading("Revision History", accent));
		children.push(
			revisionTable(
				[
					{
						version: project.version || "1.0",
						date: formatDisplayDate(project.reviewDate),
						author: project.author || "",
						notes: "Initial documented version.",
					},
				]
			)
		);
	}

	children.push(new Paragraph({ children: [new PageBreak()] }));
	return children;
}

function revisionTable(entries) {
	const headerBorder = { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD" };
	const borders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };
	const cols = ["Version", "Date", "Author", "Notes"];

	const headerRow = new TableRow({
		tableHeader: true,
		children: cols.map(
			(c) =>
				new TableCell({
					shading: { fill: LIGHT_SHADE },
					margins: { top: 60, bottom: 60, left: 120, right: 120 },
					children: [new Paragraph({ children: [new TextRun({ text: c, bold: true, color: "333333" })] })],
				})
		),
	});

	const bodyRows = entries.map(
		(e) =>
			new TableRow({
				children: [e.version, e.date, e.author, e.notes].map(
					(v) =>
						new TableCell({
							margins: { top: 60, bottom: 60, left: 120, right: 120 },
							children: [new Paragraph({ children: [new TextRun({ text: v == null ? "" : String(v) })] })],
						})
				),
			})
	);

	return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, borders, rows: [headerRow, ...bodyRows] });
}

function buildHeader(project, branding, accent) {
	return new Header({
		children: [
			new Paragraph({
				tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_WIDTH_TWIPS }],
				border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: accent } },
				spacing: { after: 120 },
				children: [
					new TextRun({ text: branding.companyName || "", bold: true, color: accent, size: 18 }),
					new TextRun({ text: "\t" + (project.title || "Standard Operating Procedure"), color: TEXT_MUTED, size: 18 }),
				],
			}),
		],
	});
}

function buildFooter(branding, accent) {
	return new Footer({
		children: [
			new Paragraph({
				tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_WIDTH_TWIPS }],
				border: { top: { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD" } },
				spacing: { before: 80 },
				children: [
					new TextRun({ text: branding.footerText || "", color: TEXT_MUTED, size: 16 }),
					new TextRun({ text: "\tPage ", color: TEXT_MUTED, size: 16 }),
					new TextRun({ children: [PageNumber.CURRENT], color: TEXT_MUTED, size: 16 }),
					new TextRun({ text: " of ", color: TEXT_MUTED, size: 16 }),
					new TextRun({ children: [PageNumber.TOTAL_PAGES], color: TEXT_MUTED, size: 16 }),
				],
			}),
		],
	});
}

function buildStepChildren(project, accent, withProcedureHeading) {
	const children = [];
	if (withProcedureHeading && (project.steps || []).length) {
		children.push(sectionHeading("Procedure", accent));
	}

	for (const s of project.steps || []) {
		children.push(...stepDescriptionParagraphs(s.description));

		if (
			Array.isArray(s.changes) &&
			s.changes.length &&
			!descriptionListsChanges(s.description, s.changes)
		) {
			for (const c of s.changes) {
				children.push(
					new Paragraph({
						children: [new TextRun({ text: formatChangePhrase(c) })],
						indent: { left: 720 },
						spacing: { after: 60 },
					})
				);
			}
		}

		const imgPara = imageParagraph(s.image || s.screenshot);
		if (imgPara) children.push(imgPara);
	}
	return children;
}

/* ------------------------------------------------------------------ *
 *  Main entry point
 * ------------------------------------------------------------------ */

async function exportWord(project, branding) {
	const b = resolveBranding(branding);
	const accent = b.accentColor;
	const template = b.template === "minimal" || b.template === "detailed" ? b.template : "standard";

	const numberingConfig = [
		{
			reference: "stepNumbering",
			levels: [
				{
					level: 0,
					format: LevelFormat.DECIMAL,
					text: "Step %1.",
					suffix: LevelSuffix.TAB,
					start: 1,
					alignment: AlignmentType.START,
				},
			],
		},
	];

	const children = [];
	const hasCover = template === "standard" || template === "detailed";

	if (hasCover) {
		children.push(...buildCoverChildren(project, b, accent));
		children.push(...buildStepChildren(project, accent, true));
	} else {
		// Minimal: simple title + meta lines + steps, no cover page.
		if (project.title) {
			children.push(new Paragraph({ text: project.title, heading: HeadingLevel.HEADING_1 }));
		}
		const metaRows = [
			["Author", project.author],
			["Version", project.version],
			["Review Date", formatDisplayDate(project.reviewDate)],
			["Document ID", project.documentId],
			["Department", project.department],
			["Classification", project.classification],
		].filter(([, v]) => v);
		for (const [label, value] of metaRows) {
			children.push(
				new Paragraph({
					children: [
						new TextRun({ text: `${label}: `, bold: true }),
						new TextRun({ text: String(value) }),
					],
				})
			);
		}
		if (metaRows.length) children.push(new Paragraph({ text: "" }));
		children.push(...buildStepChildren(project, accent, false));
	}

	const sectionProps = {
		children,
	};

	if (template === "minimal") {
		// Header (company + title) and page numbers on every page; no title page.
		sectionProps.headers = { default: buildHeader(project, b, accent) };
		sectionProps.footers = { default: buildFooter(b, accent) };
	} else {
		// Title page (cover) stays clean; headers/footers appear from page 2 on.
		sectionProps.properties = { titlePage: true };
		sectionProps.headers = { default: buildHeader(project, b, accent), first: new Header({ children: [] }) };
		sectionProps.footers = { default: buildFooter(b, accent), first: new Footer({ children: [] }) };
	}

	const exportsDir = ensureDir(getExportsDir());
	const fileName = sanitizeFileName(project.title || "SOP") + ".docx";
	const outPath = path.join(exportsDir, fileName);

	const doc = new Document({
		numbering: { config: numberingConfig },
		sections: [sectionProps],
	});

	const buf = await Packer.toBuffer(doc);
	fsExporter.writeFileSync(outPath, buf);
	return outPath;
}

module.exports = { exportWord };
