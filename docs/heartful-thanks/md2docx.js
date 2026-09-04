// Markdown (this report's subset) -> monochrome formal DOCX
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, HeadingLevel, BorderStyle, ShadingType, LevelFormat, PageNumber,
  Footer, Header, VerticalAlign, PageBreak,
} = require("docx");

const MINCHO = "游明朝";
const GOTHIC = "游ゴシック";
const BLACK = "000000";
const GRAY_FILL = "E6E6E6";
const GRAY_FILL2 = "F2F2F2";
const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 1134; // A4, 20mm
const USABLE = PAGE_W - 2 * MARGIN; // 9638

const [,, inPath, outPath] = process.argv;
const md = fs.readFileSync(inPath, "utf8").split("\n");

// ---------- inline runs ----------
function runs(text, base = {}) {
  const out = [];
  const re = /\*\*(.+?)\*\*/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(mk(text.slice(last, m.index), base));
    out.push(mk(m[1], { ...base, bold: true }));
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push(mk(text.slice(last), base));
  return out;
}
function mk(t, o) {
  const f = o.gothic ? GOTHIC : MINCHO;
  return new TextRun({ text: t, font: { ascii: f, eastAsia: f, hAnsi: f }, size: o.size || 21, bold: o.bold, color: BLACK });
}
function para(text, o = {}) {
  return new Paragraph({
    alignment: o.align || AlignmentType.BOTH,
    spacing: { before: o.before ?? 0, after: o.after ?? 120, line: o.line || 320 },
    indent: o.indent,
    children: runs(text, o),
  });
}

// ---------- numbering (one reference per list block) ----------
const numberingConfigs = [];
let listCounter = 0;
function newNumbered() {
  const ref = `num${listCounter++}`;
  numberingConfigs.push({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1．", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 600, hanging: 600 } } } }] });
  return ref;
}
numberingConfigs.push({ reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "・", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 300 } } } }] });

// ---------- tables ----------
const border = { style: BorderStyle.SINGLE, size: 4, color: BLACK };
const borders = { top: border, bottom: border, left: border, right: border };
function parseRow(line) {
  return line.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map((c) => c.trim());
}
function buildTable(lines) {
  const header = parseRow(lines[0]);
  const aligns = parseRow(lines[1]).map((a) => (/^-+:$/.test(a) ? "R" : /^:-+:$/.test(a) ? "C" : "L"));
  const rows = lines.slice(2).map(parseRow);
  const n = header.length;
  const fontSize = n >= 12 ? 15 : n >= 8 ? 16 : 18; // half-points: 7.5 / 8 / 9pt
  // column widths: first column wider for label tables
  let widths;
  if (n >= 12) { const first = 1300; const w = Math.floor((USABLE - first) / (n - 1)); widths = [first, ...Array(n - 1).fill(w)]; }
  else if (n === 8) { const first = 1500; const w = Math.floor((USABLE - first) / 7); widths = [first, ...Array(7).fill(w)]; }
  else if (n === 2) widths = [1500, USABLE - 1500];
  else if (n === 3) { widths = aligns[2] === "L" ? [3600, 1500, USABLE - 5100] : [4400, 2600, USABLE - 7000]; }
  else if (n === 4) { widths = aligns[1] === "L" ? [2300, 4300, 1400, USABLE - 8000] : [3600, 2000, 2000, USABLE - 7600]; }
  else if (n === 5) { widths = aligns[1] === "L" ? [1900, 3300, 1500, 1500, USABLE - 8200] : [2600, 2400, 1500, 1500, USABLE - 8000]; }
  else if (n === 6) { widths = [2400, 1200, 1500, 1200, 1600, USABLE - 7900]; }
  else { const w = Math.floor(USABLE / n); widths = Array(n).fill(w); }
  const sum = widths.reduce((a, b) => a + b, 0); widths[widths.length - 1] += USABLE - sum;

  const cell = (text, i, opts) => {
    const isBold = /^\*\*.*\*\*$/.test(text);
    const t = text.replace(/^\*\*|\*\*$/g, "");
    const al = opts.head ? AlignmentType.CENTER : aligns[i] === "R" ? AlignmentType.RIGHT : aligns[i] === "C" ? AlignmentType.CENTER : AlignmentType.LEFT;
    return new TableCell({
      width: { size: widths[i], type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
      shading: opts.fill ? { type: ShadingType.CLEAR, color: opts.fill, fill: opts.fill } : undefined,
      margins: { top: 40, bottom: 40, left: 70, right: 70 },
      children: [new Paragraph({ alignment: al, spacing: { after: 0, line: 240 }, children: runs(t, { size: fontSize, bold: opts.head || isBold, gothic: opts.head }) })],
    });
  };
  const trs = [new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, { head: true, fill: GRAY_FILL })) })];
  rows.forEach((r) => {
    const total = r.some((c) => /^\*\*.*\*\*$/.test(c));
    trs.push(new TableRow({ children: header.map((_, i) => cell(r[i] ?? "", i, { fill: total ? GRAY_FILL2 : undefined })) }));
  });
  return new Table({ columnWidths: widths, width: { size: USABLE, type: WidthType.DXA }, rows: trs });
}

// ---------- walk ----------
const children = [];
let i = 0, seenTitle = false, listRef = null;
const push = (x) => children.push(x);
while (i < md.length) {
  const line = md[i];
  const t = line.trim();
  if (t === "") { i++; listRef = null; continue; }
  if (t === "---") { i++; continue; }
  if (t.startsWith("| ")) {
    const block = [];
    while (i < md.length && md[i].trim().startsWith("|")) block.push(md[i++]);
    push(buildTable(block));
    push(new Paragraph({ spacing: { after: 80 }, children: [] }));
    continue;
  }
  if (t.startsWith("# ")) {
    seenTitle = true;
    push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240, after: 120 }, children: [mk(t.slice(2), { gothic: true, size: 30, bold: true })] }));
    i++; continue;
  }
  if (t.startsWith("## ")) {
    const isAppendix = t.includes("別紙");
    if (isAppendix) push(new Paragraph({ children: [new PageBreak()] }));
    push(new Paragraph({
      heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLACK, space: 2 } },
      children: [mk(t.slice(3), { gothic: true, size: 25, bold: true })],
    }));
    i++; continue;
  }
  if (t.startsWith("### ")) {
    push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 100 }, children: [mk(t.slice(4), { gothic: true, size: 22, bold: true })] }));
    i++; continue;
  }
  if (/^（\d）/.test(t) && t.length < 60) {
    push(new Paragraph({ spacing: { before: 160, after: 80 }, children: [mk(t, { gothic: true, size: 20, bold: true })] }));
    i++; continue;
  }
  if (t.startsWith("- ")) {
    push(new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 80, line: 300 }, children: runs(t.slice(2)) }));
    i++; continue;
  }
  if (/^\d+\.\s/.test(t)) {
    if (!listRef) listRef = newNumbered();
    push(new Paragraph({ numbering: { reference: listRef, level: 0 }, spacing: { after: 100, line: 300 }, children: runs(t.replace(/^\d+\.\s/, "")) }));
    i++; continue;
  }
  if (t.startsWith("※")) { push(para(t, { size: 17, after: 60, line: 260 })); i++; continue; }
  if (!seenTitle) {
    // header block
    if (/^\d{4}年\d{1,2}月\d{1,2}日$/.test(t)) push(para(t, { align: AlignmentType.RIGHT, after: 240 }));
    else if (t.endsWith("御中")) push(para(t, { align: AlignmentType.LEFT, size: 22, after: 360 }));
    else push(para(t, { align: AlignmentType.RIGHT, after: 40 }));
    i++; continue;
  }
  if (t.startsWith("―") && t.endsWith("―")) { push(para(t, { align: AlignmentType.CENTER, size: 20, after: 360 })); i++; continue; }
  if (t === "敬具" || t === "以上") { push(para(t, { align: AlignmentType.RIGHT, before: 120, after: 240 })); i++; continue; }
  push(para(t));
  i++;
}

const doc = new Document({
  creator: "株式会社ハートフルサンク",
  title: "業績改善の進捗および今後の見通しについて（ご報告）",
  styles: {
    default: { document: { run: { font: { ascii: MINCHO, eastAsia: MINCHO, hAnsi: MINCHO }, size: 21, color: BLACK } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: GOTHIC, size: 25, bold: true, color: BLACK } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: GOTHIC, size: 22, bold: true, color: BLACK } },
    ],
  },
  numbering: { config: numberingConfigs },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: 1250, bottom: 1150, left: MARGIN, right: MARGIN } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [mk("株式会社ハートフルサンク　業績改善の進捗および今後の見通しについて（ご報告）　　貴行限り", { size: 16 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT, " / ", PageNumber.TOTAL_PAGES], font: MINCHO, size: 18, color: BLACK })] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(outPath, buf); console.log("written", outPath, buf.length); });
