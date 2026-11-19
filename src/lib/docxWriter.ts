/**
 * TipTap 문서(JSON) → 진짜 DOCX(OOXML) 변환기.
 *
 * 예전에는 html-docx-js 로 저장했는데, 그 라이브러리는 HTML 을 "altChunk" 로 통째로 집어넣는
 * 방식이라 (1) 이 앱(mammoth)에서 다시 열면 내용이 비어 있고 (2) 한글/기타 뷰어에서도 빈 문서로
 * 보이는 문제가 있었다. 여기서는 문단/제목/목록/표/이미지를 실제 w:p / w:tbl 로 만든다.
 */
import JSZip from "jszip";

export type DocxPageOptions = {
  /** mm */
  top: number; bottom: number; left: number; right: number; header: number; footer: number;
  number: "none" | "footer" | "header";
};

type J = { type?: string; attrs?: Record<string, any>; content?: J[]; marks?: { type: string; attrs?: Record<string, any> }[]; text?: string };

const esc = (s: string) =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");

const mmToTwip = (mm: number) => Math.round((mm * 1440) / 25.4);

type Ctx = {
  rels: string[];                 // document.xml.rels 항목
  media: { name: string; data: Uint8Array }[];
  nextRel: number;
  nextNum: number;                // 번호 목록 인스턴스(번호 다시 시작용)
  nums: number[];                 // 사용된 번호 목록 numId
  contentWidthTwip: number;
  docPrId: number;
};

function firstFont(v: string | undefined | null): string | null {
  if (!v) return null;
  const f = v.split(",")[0].trim().replace(/^["']|["']$/g, "");
  return f || null;
}

function runProps(marks: J["marks"], extra?: { bold?: boolean }): string {
  let p = "";
  let font: string | null = null, size: string | null = null, color: string | null = null;
  let bold = !!extra?.bold, italic = false, underline = false, strike = false, code = false, link = false;
  for (const m of marks ?? []) {
    switch (m.type) {
      case "bold": bold = true; break;
      case "italic": italic = true; break;
      case "underline": underline = true; break;
      case "strike": strike = true; break;
      case "code": code = true; break;
      case "link": link = true; break;
      case "textStyle": {
        font = firstFont(m.attrs?.fontFamily) ?? font;
        if (m.attrs?.fontSize) { const n = parseFloat(String(m.attrs.fontSize)); if (n > 0) size = String(Math.round(n * 2)); }
        if (m.attrs?.color) { const c = String(m.attrs.color).trim(); const hex = cssColorToHex(c); if (hex) color = hex; }
        break;
      }
    }
  }
  if (code && !font) font = "Consolas";
  if (font) p += `<w:rFonts w:ascii="${esc(font)}" w:hAnsi="${esc(font)}" w:eastAsia="${esc(font)}" w:cs="${esc(font)}"/>`;
  if (bold) p += "<w:b/><w:bCs/>";
  if (italic) p += "<w:i/><w:iCs/>";
  if (strike) p += "<w:strike/>";
  if (color) p += `<w:color w:val="${color}"/>`;
  else if (link) p += `<w:color w:val="0563C1"/>`;
  if (size) p += `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>`;
  if (underline || link) p += `<w:u w:val="single"/>`;
  return p ? `<w:rPr>${p}</w:rPr>` : "";
}

function cssColorToHex(c: string): string | null {
  let m = /^#([0-9a-f]{6})$/i.exec(c); if (m) return m[1].toUpperCase();
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(c); if (m) return (m[1] + m[1] + m[2] + m[2] + m[3] + m[3]).toUpperCase();
  m = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/i.exec(c);
  if (m) return [m[1], m[2], m[3]].map((x) => Math.min(255, +x).toString(16).padStart(2, "0")).join("").toUpperCase();
  return null;
}

function textRuns(text: string, rPr: string): string {
  // 탭과 줄바꿈을 DOCX 요소로 바꾼다.
  const parts = text.split(/(\t|\n)/);
  let out = "";
  for (const part of parts) {
    if (part === "") continue;
    if (part === "\t") out += `<w:r>${rPr}<w:tab/></w:r>`;
    else if (part === "\n") out += `<w:r>${rPr}<w:br/></w:r>`;
    else out += `<w:r>${rPr}<w:t xml:space="preserve">${esc(part)}</w:t></w:r>`;
  }
  return out;
}

function dataUriToBytes(src: string): { bytes: Uint8Array; ext: string } | null {
  const m = /^data:image\/(png|jpe?g|gif|bmp);base64,(.*)$/is.exec(src);
  if (!m) return null;
  const bin = atob(m[2].replace(/\s/g, ""));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return { bytes, ext: m[1].toLowerCase().replace("jpeg", "jpg") };
}

function imageSize(b: Uint8Array, ext: string): [number, number] | null {
  try {
    if (ext === "png") return [(b[16] << 24) | (b[17] << 16) | (b[18] << 8) | b[19], (b[20] << 24) | (b[21] << 16) | (b[22] << 8) | b[23]];
    if (ext === "gif") return [b[6] | (b[7] << 8), b[8] | (b[9] << 8)];
    if (ext === "jpg") {
      let i = 2;
      while (i < b.length) {
        if (b[i] !== 0xff) { i++; continue; }
        const mk = b[i + 1];
        if (mk >= 0xc0 && mk <= 0xcf && mk !== 0xc4 && mk !== 0xc8 && mk !== 0xcc) return [(b[i + 7] << 8) | b[i + 8], (b[i + 5] << 8) | b[i + 6]];
        i += 2 + ((b[i + 2] << 8) | b[i + 3]);
      }
    }
  } catch { /* fallthrough */ }
  return null;
}

function inlineXml(nodes: J[] | undefined, ctx: Ctx, extra?: { bold?: boolean }): string {
  let out = "";
  for (const n of nodes ?? []) {
    if (n.type === "text") {
      const rPr = runProps(n.marks, extra);
      const link = n.marks?.find((m) => m.type === "link");
      const runs = textRuns(n.text ?? "", rPr);
      if (link?.attrs?.href) {
        const id = `rId${ctx.nextRel++}`;
        ctx.rels.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${esc(String(link.attrs.href))}" TargetMode="External"/>`);
        out += `<w:hyperlink r:id="${id}">${runs}</w:hyperlink>`;
      } else out += runs;
    } else if (n.type === "hardBreak") {
      out += `<w:r><w:br/></w:r>`;
    } else if (n.type === "image") {
      out += imageXml(n, ctx);
    }
  }
  return out;
}

function imageXml(n: J, ctx: Ctx): string {
  const src = String(n.attrs?.src ?? "");
  const img = dataUriToBytes(src);
  if (!img) return "";
  const name = `image${ctx.media.length + 1}.${img.ext}`;
  ctx.media.push({ name, data: img.bytes });
  const id = `rId${ctx.nextRel++}`;
  ctx.rels.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${name}"/>`);
  const nat = imageSize(img.bytes, img.ext) ?? [300, 200];
  let w = Number(n.attrs?.width) || nat[0] || 300;
  let h = Number(n.attrs?.height) || (nat[1] && nat[0] ? (w * nat[1]) / nat[0] : 200);
  const maxW = (ctx.contentWidthTwip / 1440) * 96;
  if (w > maxW) { h = (h * maxW) / w; w = maxW; }
  const cx = Math.round(w * 9525), cy = Math.round(h * 9525);
  const pid = ctx.docPrId++;
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${pid}" name="그림 ${pid}" descr="${esc(String(n.attrs?.alt ?? ""))}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="${pid}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

const ALIGN: Record<string, string> = { left: "left", center: "center", right: "right", justify: "both" };

function paragraphXml(n: J, ctx: Ctx, opts: { style?: string; numId?: number; ilvl?: number; bold?: boolean; indentLeft?: number } = {}): string {
  let pPr = "";
  if (opts.style) pPr += `<w:pStyle w:val="${opts.style}"/>`;
  if (opts.numId !== undefined) pPr += `<w:numPr><w:ilvl w:val="${opts.ilvl ?? 0}"/><w:numId w:val="${opts.numId}"/></w:numPr>`;
  if (opts.indentLeft) pPr += `<w:ind w:left="${opts.indentLeft}"/>`;
  const al = n.attrs?.textAlign && ALIGN[n.attrs.textAlign];
  if (al && al !== "left") pPr += `<w:jc w:val="${al}"/>`;
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${inlineXml(n.content, ctx, { bold: opts.bold })}</w:p>`;
}

function blocksXml(nodes: J[] | undefined, ctx: Ctx, list?: { numId: number; ilvl: number }, inCell = false): string {
  let out = "";
  for (const n of nodes ?? []) out += blockXml(n, ctx, list, inCell);
  return out;
}

function blockXml(n: J, ctx: Ctx, list?: { numId: number; ilvl: number }, inCell = false): string {
  switch (n.type) {
    case "paragraph":
      return list
        ? paragraphXml(n, ctx, { style: "ListParagraph", numId: list.numId, ilvl: list.ilvl })
        : paragraphXml(n, ctx);
    case "heading": {
      const lv = Math.min(6, Math.max(1, Number(n.attrs?.level) || 1));
      return paragraphXml(n, ctx, { style: `Heading${lv}` });
    }
    case "bulletList":
    case "orderedList": {
      const ordered = n.type === "orderedList";
      const ilvl = list ? Math.min(list.ilvl + 1, 8) : 0;
      let numId: number;
      if (ordered) { numId = ctx.nextNum++; ctx.nums.push(numId); } // 목록마다 새 번호 인스턴스
      else numId = 1; // 글머리 기호 공용
      let out = "";
      for (const li of n.content ?? []) {
        let first = true;
        for (const child of li.content ?? []) {
          if (child.type === "paragraph" && first) { out += paragraphXml(child, ctx, { style: "ListParagraph", numId, ilvl }); first = false; }
          else if (child.type === "paragraph") out += paragraphXml(child, ctx, { style: "ListParagraph", indentLeft: 720 * (ilvl + 1) });
          else out += blockXml(child, ctx, { numId, ilvl }, inCell);
        }
        if (first) out += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr></w:p>`;
      }
      return out;
    }
    case "blockquote":
      return (n.content ?? []).map((c) => c.type === "paragraph" ? paragraphXml(c, ctx, { style: "Quote" }) : blockXml(c, ctx, undefined, inCell)).join("");
    case "codeBlock": {
      const text = (n.content ?? []).map((c) => c.text ?? "").join("");
      return text.split("\n").map((line) => `<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr>${textRuns(line, "")}</w:p>`).join("");
    }
    case "horizontalRule":
      return `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="808080"/></w:pBdr></w:pPr></w:p>`;
    case "table":
      return tableXml(n, ctx);
    default:
      if (n.content) return blocksXml(n.content, ctx, list, inCell);
      return "";
  }
}

function tableXml(n: J, ctx: Ctx): string {
  const rows = n.content ?? [];
  // 열 개수와 열 너비
  let cols = 0;
  for (const r of rows) {
    let c = 0;
    for (const cell of r.content ?? []) c += Number(cell.attrs?.colspan) || 1;
    cols = Math.max(cols, c);
  }
  cols = Math.max(cols, 1);
  const widths: number[] = new Array(cols).fill(0);
  for (const r of rows) {
    let c = 0;
    for (const cell of r.content ?? []) {
      const span = Number(cell.attrs?.colspan) || 1;
      const cw = cell.attrs?.colwidth as number[] | null | undefined;
      if (cw && cw.length === span) for (let k = 0; k < span; k++) if (cw[k] && !widths[c + k]) widths[c + k] = Math.round(cw[k] * 15);
      c += span;
    }
  }
  const known = widths.filter(Boolean);
  const total = ctx.contentWidthTwip;
  const knownSum = known.reduce((a, b) => a + b, 0);
  const missing = cols - known.length;
  if (missing > 0) {
    const rest = Math.max(missing * 700, total - knownSum);
    for (let i = 0; i < cols; i++) if (!widths[i]) widths[i] = Math.round(rest / missing);
  }
  const sum = widths.reduce((a, b) => a + b, 0);
  if (sum > total) for (let i = 0; i < cols; i++) widths[i] = Math.max(300, Math.floor((widths[i] * total) / sum));

  const grid = `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>`;
  const occ: { rows: number; span: number }[] = Array.from({ length: cols }, () => ({ rows: 0, span: 1 })); // 위 행에서 이어지는 rowspan
  let out = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${widths.reduce((a, b) => a + b, 0)}" w:type="dxa"/><w:tblLayout w:type="fixed"/>` +
    `<w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="808080"/>`).join("")}</w:tblBorders>` +
    `<w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>${grid}`;
  for (const r of rows) {
    let tcs = "";
    let c = 0;
    const cells = [...(r.content ?? [])];
    let ci = 0;
    while (c < cols) {
      if (occ[c].rows > 0) {
        // rowspan 으로 덮인 칸: 세로 병합 이어짐
        const span = occ[c].span;
        let w = 0; for (let k = 0; k < span; k++) w += widths[c + k] ?? 0;
        tcs += `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ""}<w:vMerge/></w:tcPr><w:p/></w:tc>`;
        for (let k = 0; k < span; k++) occ[c + k].rows--;
        c += span; continue;
      }
      const cell = cells[ci++];
      if (!cell) break;
      const colspan = Math.min(Number(cell.attrs?.colspan) || 1, cols - c);
      const rowspan = Number(cell.attrs?.rowspan) || 1;
      let w = 0; for (let k = 0; k < colspan; k++) w += widths[c + k] ?? 0;
      const isHeader = cell.type === "tableHeader";
      let tcPr = `<w:tcW w:w="${w}" w:type="dxa"/>`;
      if (colspan > 1) tcPr += `<w:gridSpan w:val="${colspan}"/>`;
      if (rowspan > 1) tcPr += `<w:vMerge w:val="restart"/>`;
      const bg = cell.attrs?.backgroundColor ? cssColorToHex(String(cell.attrs.backgroundColor)) : null;
      if (bg) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
      else if (isHeader) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="F0F2F5"/>`;
      let inner = "";
      for (const child of cell.content ?? []) {
        if (child.type === "paragraph") inner += paragraphXml(child, ctx, { bold: isHeader });
        else inner += blockXml(child, ctx, undefined, true);
      }
      if (!inner.endsWith("</w:p>")) inner += "<w:p/>"; // 셀은 반드시 문단으로 끝나야 한다
      tcs += `<w:tc><w:tcPr>${tcPr}</w:tcPr>${inner}</w:tc>`;
      if (rowspan > 1) for (let k = 0; k < colspan; k++) occ[c + k] = { rows: rowspan - 1, span: colspan };
      c += colspan;
    }
    out += `<w:tr>${tcs}</w:tr>`;
  }
  out += `</w:tbl><w:p><w:pPr><w:spacing w:after="0" w:line="120" w:lineRule="exact"/></w:pPr></w:p>`;
  return out;
}

const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
  `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
  `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;

function stylesXml(): string {
  const heading = (n: number, sz: number) =>
    `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
    `<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr></w:style>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="맑은 고딕" w:cs="Malgun Gothic"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="en-US" w:eastAsia="ko-KR"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
    heading(1, 36) + heading(2, 30) + heading(3, 26) + heading(4, 24) + heading(5, 22) + heading(6, 21) +
    `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="40"/><w:ind w:left="720"/></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:i/><w:color w:val="555555"/></w:rPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Consolas"/></w:rPr></w:style>` +
    `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>` +
    `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>` +
    `</w:styles>`;
}

function numberingXml(extraNums: number[]): string {
  const bullets = ["•", "○", "■"];
  const bulletLvls = Array.from({ length: 9 }, (_, i) =>
    `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${bullets[i % 3]}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="맑은 고딕" w:hint="default"/></w:rPr></w:lvl>`).join("");
  const fmts = ["decimal", "lowerLetter", "lowerRoman"];
  const decLvls = Array.from({ length: 9 }, (_, i) =>
    `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmts[i % 3]}"/><w:lvlText w:val="%${i + 1}."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NS}>` +
    `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${bulletLvls}</w:abstractNum>` +
    `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${decLvls}</w:abstractNum>` +
    `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
    extraNums.map((id) => `<w:num w:numId="${id}"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`).join("") +
    `</w:numbering>`;
}

function pageNumberPart(kind: "hdr" | "ftr"): string {
  const tag = kind === "hdr" ? "w:hdr" : "w:ftr";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${tag} ${NS}><w:p><w:pPr><w:jc w:val="center"/></w:pPr>` +
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>` +
    `</w:p></${tag}>`;
}

export async function tiptapJsonToDocx(json: J, page: DocxPageOptions): Promise<Uint8Array> {
  const pageW = 11906, pageH = 16838;
  const ml = mmToTwip(page.left), mr = mmToTwip(page.right);
  const ctx: Ctx = { rels: [], media: [], nextRel: 10, nextNum: 2, nums: [], contentWidthTwip: pageW - ml - mr, docPrId: 1 };

  let body = blocksXml(json.content, ctx);
  if (!body) body = "<w:p/>";

  let sectRefs = "";
  const extraRels: string[] = [];
  const zip = new JSZip();
  if (page.number === "footer") {
    sectRefs = `<w:footerReference w:type="default" r:id="rId4"/>`;
    extraRels.push(`<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>`);
    zip.file("word/footer1.xml", pageNumberPart("ftr"));
  } else if (page.number === "header") {
    sectRefs = `<w:headerReference w:type="default" r:id="rId4"/>`;
    extraRels.push(`<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`);
    zip.file("word/header1.xml", pageNumberPart("hdr"));
  }
  const sect = `<w:sectPr>${sectRefs}<w:pgSz w:w="${pageW}" w:h="${pageH}"/>` +
    `<w:pgMar w:top="${mmToTwip(page.top)}" w:right="${mr}" w:bottom="${mmToTwip(page.bottom)}" w:left="${ml}" w:header="${mmToTwip(page.header)}" w:footer="${mmToTwip(page.footer)}" w:gutter="0"/></w:sectPr>`;

  zip.file("[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="bmp" ContentType="image/bmp"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>` +
    (page.number === "footer" ? `<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>` : "") +
    (page.number === "header" ? `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>` : "") +
    `</Types>`);
  zip.file("_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>` +
    extraRels.join("") + ctx.rels.join("") + `</Relationships>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}${sect}</w:body></w:document>`);
  zip.file("word/styles.xml", stylesXml());
  zip.file("word/numbering.xml", numberingXml(ctx.nums));
  for (const m of ctx.media) zip.file(`word/media/${m.name}`, m.data);
  return await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

/**
 * 예전 버전 앱(html-docx-js)이 저장한 DOCX 는 내용이 word/afchunk.mht(HTML)에 들어 있어서
 * 일반 DOCX 리더(mammoth 포함)로는 비어 보인다. 그 HTML 을 꺼내 되살린다.
 */
export async function recoverAltChunkHtml(bytes: Uint8Array): Promise<string | null> {
  try {
    const zip = await JSZip.loadAsync(bytes);
    const f = zip.file(/^word\/afchunk[^/]*\.(mht|html?)$/i)[0];
    if (!f) return null;
    const raw = await f.async("uint8array");
    const text = new TextDecoder("utf-8").decode(raw);
    // text/html 파트의 본문: 헤더 뒤 첫 빈 줄 이후 ~ 다음 boundary 전
    const m = /Content-Type:\s*text\/html[\s\S]*?\r?\n\r?\n([\s\S]*?)\r?\n--/i.exec(text);
    let body = m ? m[1] : text;
    const qp = /Content-Transfer-Encoding:\s*quoted-printable/i.test(text);
    if (qp) {
      const src = new TextEncoder().encode(body);
      const out: number[] = [];
      for (let i = 0; i < src.length; i++) {
        const c = src[i];
        if (c === 0x3d) { // '='
          if (src[i + 1] === 0x0d && src[i + 2] === 0x0a) { i += 2; continue; }
          if (src[i + 1] === 0x0a) { i += 1; continue; }
          const hex = String.fromCharCode(src[i + 1] ?? 0, src[i + 2] ?? 0);
          if (/^[0-9a-fA-F]{2}$/.test(hex)) { out.push(parseInt(hex, 16)); i += 2; continue; }
        }
        out.push(c);
      }
      body = new TextDecoder("utf-8").decode(Uint8Array.from(out));
    }
    const d = new DOMParser().parseFromString(body, "text/html");
    return d.body.innerHTML || null;
  } catch { return null; }
}
