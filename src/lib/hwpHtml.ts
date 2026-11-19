/**
 * TipTap(HTML) → rhwp.pasteHtml 이 제대로 해석하는 HTML 로 바꿔 준다.
 *
 * rhwp 의 HTML 파서는 <style> 시트나 CSS 클래스를 보지 않고 "인라인 style" 만 읽는다.
 * 그런데 TipTap 의 getHTML() 은 표에 테두리/너비/여백 스타일이 하나도 없고 제목(h1~h6)도 글자 크기가
 * 없어서, DOCX/TXT → HWP 로 변환하면 표가 테두리 없이 흩어져 보이는(=표가 깨진) 문제가 있었다.
 */

const HEADING_PT: Record<string, number> = { H1: 18, H2: 15, H3: 13, H4: 12, H5: 11, H6: 10.5 };
const PX_PER_PT = 96 / 72;

function textWidthPx(text: string, fontPx: number): number {
  let w = 0;
  for (const ch of text) w += /[\u1100-\u11ff\u3130-\u318f\uac00-\ud7af\u3000-\u9fff\uff00-\uffef]/.test(ch) ? fontPx : fontPx * 0.55;
  return w;
}

type Grid = { cols: number; widths: number[] };

function readGrid(table: HTMLTableElement, totalPx: number): Grid {
  const rows = Array.from(table.rows);
  let cols = 0;
  for (const r of rows) { let c = 0; for (const cell of Array.from(r.cells)) c += cell.colSpan || 1; cols = Math.max(cols, c); }
  cols = Math.max(cols, 1);
  const widths: number[] = new Array(cols).fill(0);
  // 1) TipTap 이 열 너비를 조절한 경우: <col style="width:..px"> 또는 셀의 data-colwidth
  const colEls = Array.from(table.querySelectorAll("colgroup > col"));
  colEls.forEach((c, i) => {
    const m = /(?:^|;)\s*width:\s*([\d.]+)px/.exec(c.getAttribute("style") ?? "");
    if (m && i < cols) widths[i] = parseFloat(m[1]);
  });
  for (const r of rows) {
    let c = 0;
    for (const cell of Array.from(r.cells)) {
      const span = cell.colSpan || 1;
      const cw = (cell.getAttribute("colwidth") ?? cell.getAttribute("data-colwidth") ?? "").split(",").map(Number);
      if (cw.length === span) for (let k = 0; k < span; k++) if (cw[k] > 0 && !widths[c + k]) widths[c + k] = cw[k];
      c += span;
    }
  }
  const known = widths.filter(Boolean);
  const missing = cols - known.length;
  if (missing > 0) {
    const rest = Math.max(missing * 40, totalPx - known.reduce((a, b) => a + b, 0));
    for (let i = 0; i < cols; i++) if (!widths[i]) widths[i] = rest / missing;
  }
  const sum = widths.reduce((a, b) => a + b, 0);
  if (sum > totalPx) for (let i = 0; i < cols; i++) widths[i] = (widths[i] * totalPx) / sum;
  return { cols, widths: widths.map((w) => Math.round(w)) };
}

function cellFontPx(cell: HTMLElement): number {
  const m = /font-size:\s*([\d.]+)pt/.exec(cell.innerHTML);
  return (m ? parseFloat(m[1]) : 10) * PX_PER_PT;
}

function decorateTable(table: HTMLTableElement, totalPx: number) {
  const { cols, widths } = readGrid(table, totalPx);
  const doc = table.ownerDocument;
  table.setAttribute("border", "1");
  table.setAttribute("cellspacing", "0");
  table.setAttribute("cellpadding", "0");
  table.setAttribute("style", `border-collapse:collapse;width:${widths.reduce((a, b) => a + b, 0)}px`);
  table.querySelectorAll("colgroup").forEach((g) => g.remove());
  // 열 너비 (rhwp 는 셀의 width 를 사용한다)
  const occ: number[] = new Array(cols).fill(0); // rowspan 으로 점유된 남은 행 수
  for (const row of Array.from(table.rows)) {
    let c = 0;
    let rowH = 0;
    for (const cell of Array.from(row.cells)) {
      while (c < cols && occ[c] > 0) { occ[c]--; c++; }
      const span = Math.min(cell.colSpan || 1, cols - c);
      const rs = cell.rowSpan || 1;
      let w = 0; for (let k = 0; k < span; k++) w += widths[c + k] ?? 0;
      const isTh = cell.tagName === "TH";
      const bg = isTh ? "background-color:#f0f2f5;" : "";
      const bold = isTh ? "font-weight:bold;" : "";
      cell.setAttribute("style", `border:1px solid #808080;width:${w}px;${bg}${bold}`);
      // th 는 같은 모양의 td 로 바꾼다 (제목 셀의 굵기는 위 style 로 유지)
      if (isTh) {
        const td = doc.createElement("td");
        for (const a of Array.from(cell.attributes)) td.setAttribute(a.name, a.value);
        td.innerHTML = cell.innerHTML;
        td.querySelectorAll("p").forEach((p) => { if (!p.querySelector("strong,b")) p.innerHTML = `<strong>${p.innerHTML}</strong>`; });
        cell.replaceWith(td);
      }
      // 행 높이 추정 (rhwp 는 HTML 의 height 를 최소 높이로 쓴다)
      if (rs === 1) {
        const fontPx = cellFontPx(cell);
        const inner = Math.max(20, w - 16);
        let lines = 0;
        const blocks = Array.from(cell.querySelectorAll("p,li"));
        if (blocks.length === 0) lines = Math.max(1, Math.ceil(textWidthPx(cell.textContent ?? "", fontPx) / inner));
        for (const p of blocks) lines += Math.max(1, Math.ceil(textWidthPx(p.textContent ?? "", fontPx) / inner));
        rowH = Math.max(rowH, Math.ceil(lines * fontPx * 1.6 + 12));
      }
      for (let k = 0; k < span; k++) if (rs > 1) occ[c + k] = rs - 1;
      c += span;
    }
    if (rowH > 0) for (const cell of Array.from(row.cells)) {
      const st = cell.getAttribute("style") ?? "";
      cell.setAttribute("style", `${st}height:${rowH}px;`);
    }
  }
}

export function htmlForHwp(html: string, bodyWidthPx: number): string {
  const d = new DOMParser().parseFromString(`<body>${html}</body>`, "text/html");
  const total = Math.max(120, Math.floor(bodyWidthPx));

  // 제목: 글자 크기 + 굵게를 span 으로 명시한 문단으로
  d.querySelectorAll<HTMLElement>("h1,h2,h3,h4,h5,h6").forEach((h) => {
    const p = d.createElement("p");
    const al = h.style.textAlign; if (al) p.setAttribute("style", `text-align:${al}`);
    const span = d.createElement("span");
    span.setAttribute("style", `font-size:${HEADING_PT[h.tagName]}pt;font-weight:bold`);
    span.innerHTML = h.innerHTML;
    p.appendChild(span);
    h.replaceWith(p);
  });
  // 번호 목록: rhwp 는 <ol> 을 글머리 기호로 바꿔 버리므로 번호를 글자로 직접 붙인다.
  Array.from(d.querySelectorAll("ol")).reverse().forEach((ol) => {
    const depth = (() => { let n = 0, e: Element | null = ol.parentElement; while (e) { if (e.tagName === "OL" || e.tagName === "UL") n++; e = e.parentElement; } return n; })();
    let i = Number(ol.getAttribute("start") ?? "1") || 1;
    const frag = d.createDocumentFragment();
    Array.from(ol.children).forEach((li) => {
      if (li.tagName !== "LI") return;
      const kids = Array.from(li.children);
      const first = kids.find((k) => k.tagName === "P");
      const p = d.createElement("p");
      if (depth > 0) p.setAttribute("style", `margin-left:${depth * 20}px`);
      p.innerHTML = `${i}. ${first ? first.innerHTML : li.textContent ?? ""}`;
      frag.appendChild(p);
      first?.remove();
      kids.filter((k) => k !== first).forEach((k) => frag.appendChild(k));
      i++;
    });
    ol.replaceWith(frag);
  });
  // 표 (중첩 표는 바깥 것부터 처리)
  Array.from(d.querySelectorAll<HTMLTableElement>("table")).forEach((t) => {
    const nested = t.parentElement?.closest("table");
    decorateTable(t, nested ? Math.round(total / 2) : total);
  });
  // 표 안의 문단 간격이 과하지 않도록 빈 문단은 유지하되, 표 바로 뒤에 문단이 없으면 하나 둔다
  d.querySelectorAll("table").forEach((t) => {
    if (!t.nextElementSibling && t.parentElement === d.body) d.body.appendChild(d.createElement("p"));
  });
  return d.body.innerHTML;
}
