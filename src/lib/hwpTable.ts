import type { HwpDocument } from "../wasm/rhwp.js";

/** 본문에 놓인 표 하나를 가리킨다. (sec: 구역, parent: 표가 들어 있는 본문 문단, ctrl: 그 문단 속 컨트롤 번호) */
export type TableRef = { sec: number; parent: number; ctrl: number };
export type CellInfo = { idx: number; row: number; col: number; rs: number; cs: number };
export type Dir = "left" | "right" | "up" | "down";
export type GridRect = { r1: number; c1: number; r2: number; c2: number };

/** 표의 모든 셀(병합 정보 포함)을 줄·칸 순서로 돌려준다. */
export function listCells(d: HwpDocument, t: TableRef): CellInfo[] {
  const dim = JSON.parse(d.getTableDimensions(t.sec, t.parent, t.ctrl)) as { cellCount: number };
  const cells: CellInfo[] = [];
  for (let i = 0; i < dim.cellCount; i++) {
    try {
      const c = JSON.parse(d.getCellInfo(t.sec, t.parent, t.ctrl, i)) as { row: number; col: number; rowSpan: number; colSpan: number };
      cells.push({ idx: i, row: c.row, col: c.col, rs: Math.max(1, c.rowSpan), cs: Math.max(1, c.colSpan) });
    } catch { /* 읽을 수 없는 셀은 건너뛴다 */ }
  }
  return cells.sort((a, b) => a.row - b.row || a.col - b.col);
}

/** (row, col) 칸을 덮고 있는 셀 (병합된 셀 포함). */
export function cellAt(cells: CellInfo[], row: number, col: number): CellInfo | undefined {
  return cells.find((c) => row >= c.row && row < c.row + c.rs && col >= c.col && col < c.col + c.cs);
}

export function cellByIdx(cells: CellInfo[], idx: number): CellInfo | undefined {
  return cells.find((c) => c.idx === idx);
}

/** 기준 셀에서 방향키 방향으로 이웃한 셀. 표 밖이면 undefined. */
export function neighborCell(cells: CellInfo[], cur: CellInfo, dir: Dir): CellInfo | undefined {
  switch (dir) {
    case "left": return cellAt(cells, cur.row, cur.col - 1);
    case "right": return cellAt(cells, cur.row, cur.col + cur.cs);
    case "up": return cellAt(cells, cur.row - 1, cur.col);
    case "down": return cellAt(cells, cur.row + cur.rs, cur.col);
  }
}

/** 두 셀(시작/끝)을 모서리로 하는 직사각형. 걸쳐 있는 병합 셀은 통째로 포함하도록 넓힌다. */
export function blockRect(cells: CellInfo[], a: CellInfo, f: CellInfo): GridRect {
  const rect: GridRect = {
    r1: Math.min(a.row, f.row), c1: Math.min(a.col, f.col),
    r2: Math.max(a.row + a.rs - 1, f.row + f.rs - 1), c2: Math.max(a.col + a.cs - 1, f.col + f.cs - 1),
  };
  for (let again = true; again;) {
    again = false;
    for (const c of cells) {
      const hit = c.row <= rect.r2 && c.row + c.rs - 1 >= rect.r1 && c.col <= rect.c2 && c.col + c.cs - 1 >= rect.c1;
      if (!hit) continue;
      const n = {
        r1: Math.min(rect.r1, c.row), c1: Math.min(rect.c1, c.col),
        r2: Math.max(rect.r2, c.row + c.rs - 1), c2: Math.max(rect.c2, c.col + c.cs - 1),
      };
      if (n.r1 !== rect.r1 || n.c1 !== rect.c1 || n.r2 !== rect.r2 || n.c2 !== rect.c2) { Object.assign(rect, n); again = true; }
    }
  }
  return rect;
}

export function cellsInRect(cells: CellInfo[], r: GridRect): CellInfo[] {
  return cells.filter((c) => c.row <= r.r2 && c.row + c.rs - 1 >= r.r1 && c.col <= r.c2 && c.col + c.cs - 1 >= r.c1);
}

/** 한 셀 안의 문단 수 */
export function cellParaCount(d: HwpDocument, t: TableRef, cellIdx: number): number {
  return d.getCellParagraphCount(t.sec, t.parent, t.ctrl, cellIdx);
}

export const cellPathJson = (t: TableRef, cellIdx: number, para = 0) =>
  JSON.stringify([{ controlIndex: t.ctrl, cellIndex: cellIdx, cellParaIndex: para }]);

/** 셀의 글(문단 사이는 줄바꿈) */
export function cellText(d: HwpDocument, t: TableRef, cellIdx: number): string {
  const n = cellParaCount(d, t, cellIdx);
  const parts: string[] = [];
  for (let p = 0; p < n; p++) {
    try { parts.push(d.getTextInCellByPath(t.sec, t.parent, cellPathJson(t, cellIdx, p), 0, 100000)); }
    catch { parts.push(""); }
  }
  return parts.join("\n");
}

/** 블록으로 잡은 셀들의 글을 줄×칸 격자로 만든다. (병합으로 가려진 칸은 빈 글) */
export function blockGrid(d: HwpDocument, t: TableRef, cells: CellInfo[], r: GridRect): string[][] {
  const grid: string[][] = [];
  for (let row = r.r1; row <= r.r2; row++) {
    const line: string[] = [];
    for (let col = r.c1; col <= r.c2; col++) {
      const c = cellAt(cells, row, col);
      line.push(c && c.row === row && c.col === col ? cellText(d, t, c.idx) : "");
    }
    grid.push(line);
  }
  return grid;
}

export const gridToTsv = (g: string[][]) => g.map((l) => l.map((s) => s.replace(/[\t\r\n]+/g, " ")).join("\t")).join("\n");

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
export const gridToHtml = (g: string[][]) =>
  `<table border="1">${g.map((l) => `<tr>${l.map((s) => `<td>${esc(s).replace(/\n/g, "<br>")}</td>`).join("")}</tr>`).join("")}</table>`;

export function tableInfoAt(d: HwpDocument, t: TableRef): { rows: number; cols: number } | null {
  try {
    const dim = JSON.parse(d.getTableDimensions(t.sec, t.parent, t.ctrl)) as { rowCount: number; colCount: number };
    return { rows: dim.rowCount, cols: dim.colCount };
  } catch { return null; }
}
