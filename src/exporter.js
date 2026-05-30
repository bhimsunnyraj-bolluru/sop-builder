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
} = require("docx");
const fsExporter = require("fs");
const path = require("path");
const sizeOf = require("image-size");
const { getExportsDir, ensureDir } = require("../paths");

function sanitizeFileName(name) {
	return name.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").slice(0, 120);
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

function metaLine(label, value) {
	if (!value) return null;
	return new Paragraph({
		children: [
			new TextRun({ text: `${label}: `, bold: true }),
			new TextRun({ text: String(value) }),
		],
	});
}

function imageParagraph(imagePath) {
	if (!imagePath || !fsExporter.existsSync(imagePath)) return null;
	let stat;
	try {
		stat = fsExporter.statSync(imagePath);
	} catch {
		return null;
	}
	if (!stat.isFile()) return null;
	const img = fsExporter.readFileSync(imagePath);
	const maxWidth = 600;
	let width = maxWidth;
	let height = Math.round(maxWidth * 0.6);
	try {
		const dims = sizeOf(imagePath);
		width = dims.width || maxWidth;
		height = dims.height || height;
		if (width > maxWidth) {
			const ratio = maxWidth / width;
			width = Math.round(width * ratio);
			height = Math.round(height * ratio);
		}
	} catch {
		/* use defaults */
	}
	return new Paragraph({
		children: [
			new ImageRun({
				type: "png",
				data: img,
				transformation: { width, height },
			}),
		],
		spacing: { after: 240 },
	});
}

async function exportWord(project) {
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

	if (project.title) {
		children.push(new Paragraph({ text: project.title, heading: HeadingLevel.HEADING_1 }));
	}

	const metaRows = [
		metaLine("Author", project.author),
		metaLine("Version", project.version),
		metaLine("Review Date", formatDisplayDate(project.reviewDate)),
	].filter(Boolean);

	children.push(...metaRows);
	if (metaRows.length) {
		children.push(new Paragraph({ text: "" }));
	}

	for (const s of project.steps || []) {
		children.push(
			new Paragraph({
				children: [new TextRun({ text: s.description || " " })],
				numbering: { reference: "stepNumbering", level: 0 },
				spacing: { after: 120 },
			})
		);
		const imgPara = imageParagraph(s.image);
		if (imgPara) children.push(imgPara);
	}

	const exportsDir = ensureDir(getExportsDir());

	const fileName = sanitizeFileName(project.title || "SOP") + ".docx";
	const outPath = path.join(exportsDir, fileName);

	const doc = new Document({
		numbering: { config: numberingConfig },
		sections: [{ children }],
	});

	const buf = await Packer.toBuffer(doc);
	fsExporter.writeFileSync(outPath, buf);
	return outPath;
}

module.exports = { exportWord };
