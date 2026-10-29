// review.md(간단 마크다운)를 docx로 변환: 제목(#), 표(|), 번호 목록(1.), 글머리(-), 굵게(**)
// 사용: NODE_PATH=/opt/node22/lib/node_modules node build_review.js review.md out.docx
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  AlignmentType, LevelFormat, ShadingType, HeadingLevel } = require("docx");

const [, , SRC, OUT] = process.argv;
const FONT = { ascii: "Arial", hAnsi: "Arial", eastAsia: "맑은 고딕", cs: "Arial" };
const lines = fs.readFileSync(SRC, "utf8").split("\n");

let meta = {};
let i = 0;
if (lines[0] === "---") {
  i = 1;
  while (lines[i] !== "---") {
    const m = lines[i].match(/^(\w+):\s*"(.*)"$/);
    if (m) meta[m[1]] = m[2];
    i++;
  }
  i++;
}

function runs(text, o = {}) {
  return text.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((s) =>
    s.startsWith("**") ? new TextRun({ text: s.slice(2, -2), bold: true, font: FONT, size: o.size || 20 })
      : new TextRun({ text: s, bold: o.bold, font: FONT, size: o.size || 20, color: o.color }));
}
const border = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
const W = 9638;   // A4, 여백 20mm

const children = [];
if (meta.title) children.push(new Paragraph({ children: runs(meta.title, { size: 30, bold: true }), alignment: AlignmentType.CENTER, spacing: { after: 120 } }));
if (meta.subtitle) children.push(new Paragraph({ children: runs(meta.subtitle, { size: 20, color: "444444" }), alignment: AlignmentType.CENTER, spacing: { after: 80 } }));
if (meta.date) children.push(new Paragraph({ children: runs(meta.date, { size: 18, color: "666666" }), alignment: AlignmentType.CENTER, spacing: { after: 240 } }));

let tableRows = [];
const flushTable = () => {
  if (!tableRows.length) return;
  const rows = tableRows.filter((r) => !/^\|\s*-/.test(r)).map((r) => r.replace(/^\||\|$/g, "").split("|").map((c) => c.trim()));
  // 열 너비: 열마다 가장 긴 칸의 글자 수에 비례(긴 문장 열은 60자로 상한)
  const lens = rows[0].map((_, ci) => Math.min(60, Math.max(...rows.map((r) => (r[ci] || "").replace(/\*\*/g, "").length))) + 4);
  const sum = lens.reduce((a, b) => a + b, 0);
  const widths = lens.map((l) => Math.floor((W * l) / sum));
  children.push(new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((cells, ri) => new TableRow({
      tableHeader: ri === 0,
      children: cells.map((c, ci) => new TableCell({
        width: { size: widths[ci], type: WidthType.DXA }, borders,
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        shading: ri === 0 ? { fill: "1F4E78", type: ShadingType.CLEAR, color: "auto" } : undefined,
        children: [new Paragraph({ children: ri === 0 ? [new TextRun({ text: c, bold: true, color: "FFFFFF", font: FONT, size: 18 })] : runs(c, { size: 18 }) })],
      })),
    })),
  }));
  children.push(new Paragraph({ children: [], spacing: { after: 120 } }));
  tableRows = [];
};

for (; i < lines.length; i++) {
  const l = lines[i];
  if (l.startsWith("|")) { tableRows.push(l); continue; }
  flushTable();
  if (!l.trim()) continue;
  if (l.startsWith("# ")) {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: runs(l.slice(2), { size: 26, bold: true }), spacing: { before: 280, after: 120 } }));
  } else if (/^\d+\.\s/.test(l)) {
    children.push(new Paragraph({ numbering: { reference: "num", level: 0 }, children: runs(l.replace(/^\d+\.\s/, "")), spacing: { after: 60 } }));
  } else if (l.startsWith("- ")) {
    children.push(new Paragraph({ numbering: { reference: "bul", level: 0 }, children: runs(l.slice(2)), spacing: { after: 60 } }));
  } else {
    children.push(new Paragraph({ children: runs(l), spacing: { after: 100 } }));
  }
}
flushTable();

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 20 } } } },
  numbering: {
    config: [
      { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 300 } } } }] },
      { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 300 } } } }] },
    ],
  },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } }, children }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(OUT, b); console.log("saved", OUT); });
