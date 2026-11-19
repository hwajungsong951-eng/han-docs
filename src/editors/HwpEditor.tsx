import { forwardRef, useEffect, useImperativeHandle, useRef, useState, useCallback } from "react";
import { HwpDocument } from "../wasm/rhwp.js";
import { ensureWasm } from "../wasm/init";
import { readText, writeHtml, writeText } from "@tauri-apps/plugin-clipboard-manager";
import { open } from "@tauri-apps/plugin-dialog";
import { readBytes } from "../lib/fileio";
import { reflowControlParagraphs, hasOverlappingTables } from "../lib/hwpReflow";
import PlacementFields from "./PlacementFields";
import { floatingDefaults, placementDiff, placementFromProps, type Placement } from "../lib/hwpObject";
import { nextWordOffset, prevWordOffset } from "../lib/hwpText";
import { HANDLES, HU_PER_MM as PIC_HU_PER_MM, HU_PER_PX, handleCursor, handlePos, huToMm as picHuToMm, mmToHu as picMmToHu, resizeBox, type Box, type Handle, type PicRef } from "../lib/hwpPicture";
import {
  blockGrid, blockRect, cellAt, cellByIdx, cellsInRect, cellParaCount, cellPathJson, gridToHtml, gridToTsv,
  listCells, neighborCell, tableInfoAt,
  type CellInfo, type Dir, type GridRect, type TableRef,
} from "../lib/hwpTable";
import { PDFDocument } from "pdf-lib";
import {
  Bold, Italic, Underline, Strikethrough, AlignLeft, AlignCenter, AlignRight, AlignJustify,
  Table as TableIcon, Square, Circle, Type, Minus, Undo2, Redo2, ZoomIn, ZoomOut, Image as ImageIcon, Smile, Settings2, Search as SearchIcon,
} from "lucide-react";
import { TBtn, Sep, TSelect, SIZES } from "../components/ui";
import type { ExportFormat } from "../lib/fileio";
import type { DocEditorHandle } from "./RichEditor";

type CellEntry = { controlIndex: number; cellIndex: number; cellParaIndex: number };
type CellPos = { parentPara: number; path: CellEntry[] };
type Pos = { sec: number; para: number; off: number; cell?: CellPos };
type Caret = { page: number; x: number; y: number; h: number };
type SelRect = { page: number; x: number; y: number; w: number; h: number };
type SelRange = { start: Pos; end: Pos };

/** 셀 블록 선택 상태. mode: cell=셀 하나(방향키로 블록 이동), extend=확장(방향키로 넓힘), table=표 전체(개체 선택) */
type CellSel = { t: TableRef; a: number; f: number; mode: "cell" | "extend" | "table"; back: Pos };
type CellOverlay = { cells: SelRect[]; frames: SelRect[]; focus: SelRect | null };
type TableOp = "rowAbove" | "rowBelow" | "colLeft" | "colRight" | "delRow" | "delCol" | "merge" | "split" | "clear" | "delTable";
const EMPTY_OVERLAY: CellOverlay = { cells: [], frames: [], focus: null };

/** 표 바로 안(깊이 1)의 셀에 있는 위치면 그 표와 셀 번호를 돌려준다. (표 안의 표는 대상 아님) */
const tableOfPos = (p: Pos | null | undefined): { t: TableRef; cell: number } | null => {
  if (!p?.cell || p.cell.path.length !== 1) return null;
  const e = p.cell.path[0];
  return { t: { sec: p.sec, parent: p.cell.parentPara, ctrl: e.controlIndex }, cell: e.cellIndex };
};
const sameTable = (a: TableRef, b: TableRef) => a.sec === b.sec && a.parent === b.parent && a.ctrl === b.ctrl;
const cellPosOf = (t: TableRef, cellIdx: number, para = 0, off = 0): Pos => ({
  sec: t.sec, para, off,
  cell: { parentPara: t.parent, path: [{ controlIndex: t.ctrl, cellIndex: cellIdx, cellParaIndex: para }] },
});
const wholeRect = (cells: CellInfo[]): GridRect => ({
  r1: 0, c1: 0,
  r2: Math.max(0, ...cells.map((c) => c.row + c.rs - 1)), c2: Math.max(0, ...cells.map((c) => c.col + c.cs - 1)),
});
type ImageBox = { x: number; y: number; w: number; h: number; ref: PicRef; z: number };
type PicDlg = { ref: PicRef; wMm: string; hMm: string; keep: boolean; pl: Placement; plPrev: Placement; ratio: number };
type TblDlg = { t: TableRef; pl: Placement; plPrev: Placement; wMm: number; hMm: number };
type MoveDrag = { sx: number; sy: number; page: number; box: { x: number; y: number; w: number; h: number }; dx: number; dy: number };
type TableBox = { x: number; y: number; w: number; h: number; t: TableRef; cells: { idx: number; row: number; col: number; x: number; w: number }[] };

const errorText = (e: unknown) => (e instanceof Error ? e.message : String(e));

const HWP_FONTS: [string, string][] = [
  ["함초롬바탕", "함초롬바탕"],
  ["함초롬돋움", "함초롬돋움"],
  ["맑은 고딕", "맑은 고딕"],
  ["바탕", "바탕"],
  ["돋움", "돋움"],
  ["굴림", "굴림"],
  ["궁서", "궁서"],
  ["Arial", "Arial"],
  ["Times New Roman", "Times New Roman"],
];

/** Windows 기호/이모지 패널에서 자주 쓰는 문자들 */const SYMBOL_GROUPS: { name: string; items: string[] }[] = [
  { name: "일반", items: ["★", "☆", "○", "●", "◎", "◇", "◆", "□", "■", "△", "▲", "▽", "▼", "※", "§", "¶", "†", "‡", "…", "·", "•", "‥", "―", "～"] },
  { name: "화살표", items: ["→", "←", "↑", "↓", "↔", "↕", "⇒", "⇐", "⇔", "↗", "↘", "↙", "↖", "⇧", "⇩", "⇨", "⇦", "➡", "⬅", "⬆", "⬇"] },
  { name: "수학", items: ["±", "×", "÷", "≠", "≤", "≥", "∞", "∑", "∏", "√", "∫", "∂", "∆", "∇", "≈", "≡", "∝", "∠", "⊥", "∥", "∈", "∉", "∪", "∩"] },
  { name: "단위/통화", items: ["℃", "℉", "㎜", "㎝", "㎞", "㎡", "㎥", "㎏", "㏄", "％", "‰", "$", "€", "£", "¥", "¢", "₩", "₽", "₹", "°", "′", "″", "Ω", "μ"] },
  { name: "원문자", items: ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫", "⑬", "⑭", "⑮", "⑯", "⑰", "⑱", "⑲", "⑳", "ⓐ", "ⓑ", "ⓒ", "ⓓ"] },
  { name: "체크/기타", items: ["✓", "✔", "☑", "☐", "☞", "☜", "☝", "☎", "☕", "♠", "♣", "♦", "♥", "♡", "♪", "♫", "☀", "☁", "☂", "☃", "⚡", "✈", "✉", "✏"] },
  { name: "이모지", items: ["😀", "😄", "😊", "😍", "😘", "😎", "🤔", "😅", "😂", "😭", "😡", "👍", "👎", "👏", "🙏", "💪", "🎉", "🔥", "⭐", "❤️", "💡", "📌", "✅", "❌"] },
];

const TOGGLE_KEYS = ["bold", "italic", "underline", "strikethrough"] as const;

const posFromHit = (h: Record<string, unknown>): Pos => {
  const cellPath = h.cellPath as CellEntry[] | undefined;
  const cell: CellPos | undefined = Array.isArray(cellPath) && cellPath.length
    ? {
        parentPara: Number(h.parentParaIndex ?? 0),
        path: cellPath.map((c) => ({
          controlIndex: c.controlIndex, cellIndex: c.cellIndex, cellParaIndex: c.cellParaIndex,
        })),
      }
    : undefined;
  return { sec: Number(h.sectionIndex), para: Number(h.paragraphIndex), off: Number(h.charOffset), cell };
};

const sameCell = (a: Pos, b: Pos) => {
  if (!a.cell || !b.cell) return !a.cell && !b.cell;
  return a.cell.parentPara === b.cell.parentPara
    && JSON.stringify(a.cell.path.slice(0, -1)) === JSON.stringify(b.cell.path.slice(0, -1));
};

const cmpPos = (a: Pos, b: Pos) => (a.para - b.para) || (a.off - b.off);

/** 셀 문단 위치를 rhwp cellPath JSON으로 변환한다. */
const cellPathJsonFor = (p: Pos) => {
  if (!p.cell) return "[]";
  const path = p.cell.path.map((entry, i) =>
    i === p.cell!.path.length - 1 ? { ...entry, cellParaIndex: p.para } : entry);
  return JSON.stringify(path);
};

type PageMargins = { top: string; bottom: string; left: string; right: string; header: string; footer: string };
const A4_DEFAULT_MARGINS: PageMargins = { top: "15", bottom: "15", left: "20", right: "20", header: "10", footer: "10" };
const HU_PER_MM = 7200 / 25.4;
const huToMm = (hu: number) => String(Math.round((hu / HU_PER_MM) * 10) / 10).replace(/\.0$/, "");
const mmToHu = (mm: string) => Math.round((Number(mm) || 0) * HU_PER_MM);

/** HWP/HWPX 페이지를 고해상도로 렌더링해 PDF로 묶는다. */
async function hwpToPdf(doc: HwpDocument): Promise<Uint8Array> {
  const count = doc.pageCount();
  if (count <= 0) throw new Error("PDF로 저장할 페이지가 없습니다");
  const pdf = await PDFDocument.create();
  const scale = 2;
  for (let i = 0; i < count; i++) {
    const src = document.createElement("canvas");
    doc.renderPageToCanvas(i, src, scale);
    const out = document.createElement("canvas");
    out.width = src.width;
    out.height = src.height;
    const ctx = out.getContext("2d");
    if (!ctx) throw new Error("PDF 페이지를 만들 수 없습니다");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, out.width, out.height);
    ctx.drawImage(src, 0, 0);
    const dataUrl = out.toDataURL("image/jpeg", 0.92);
    const bytes = Uint8Array.from(atob(dataUrl.split(",")[1]), (c) => c.charCodeAt(0));
    const image = await pdf.embedJpg(bytes);
    // 96dpi 렌더링 → PDF point(72dpi)
    const wPt = (src.width / scale) * 0.75;
    const hPt = (src.height / scale) * 0.75;
    const page = pdf.addPage([wPt, hPt]);
    page.drawImage(image, { x: 0, y: 0, width: wPt, height: hPt });
  }
  return await pdf.save();
}

/** HWP / HWPX 편집기 (rhwp 엔진, Canvas 렌더링) */
export const HwpEditor = forwardRef<DocEditorHandle, { bytes: Uint8Array | null; onDirty: () => void; onError: (m: string) => void }>(
  ({ bytes, onDirty, onError }, ref) => {
    const docRef = useRef<HwpDocument | null>(null);
    const [pageCount, setPageCount] = useState(0);
    const [rev, setRev] = useState(0);              // 다시 그리기 트리거
    const [zoom, setZoom] = useState(1);
    const [pos, setPos] = useState<Pos | null>(null);
    const [anchor, setAnchor] = useState<Pos | null>(null); // 선택 시작
    const [caret, setCaret] = useState<Caret | null>(null);
    const [selRects, setSelRects] = useState<SelRect[]>([]);
    const [size, setSize] = useState("10");
    const [color, setColor] = useState("#000000");
    const [font, setFont] = useState("함초롬바탕");
    const [bold, setBold] = useState(false);
    const [italic, setItalic] = useState(false);
    const [underline, setUnderline] = useState(false);
    const [strike, setStrike] = useState(false);
    const [loadError, setLoadError] = useState<string | null>(null);
    const [reloadKey, setReloadKey] = useState(0);
    const [loading, setLoading] = useState(true);
    const [visiblePages, setVisiblePages] = useState<Set<number>>(() => new Set([0]));
    const [symbolsOpen, setSymbolsOpen] = useState(false);
    const [pendingHint, setPendingHint] = useState(false);
    const [pageOpen, setPageOpen] = useState(false);
    const [pageMargins, setPageMargins] = useState<PageMargins>(A4_DEFAULT_MARGINS);
    const [pageNum, setPageNum] = useState<"none" | "footer" | "header">("none");
    const canvases = useRef<(HTMLCanvasElement | null)[]>([]);
    const pageRefs = useRef<(HTMLDivElement | null)[]>([]);
    const scrollRef = useRef<HTMLDivElement>(null);
    const symbolRef = useRef<HTMLDivElement>(null);
    const pagePanelRef = useRef<HTMLDivElement>(null);
    const inputRef = useRef<HTMLTextAreaElement>(null);
    const undoStack = useRef<number[]>([]);
    const redoStack = useRef<number[]>([]);
    const pendingRef = useRef<Record<string, unknown> | null>(null);
    const pendingPosRef = useRef<Pos | null>(null);
    const preferredXRef = useRef(-1);
    const lastEditRef = useRef<{ kind: string; at: number } | null>(null);
    const dragRef = useRef<{ active: boolean; start: Pos; moved: boolean } | null>(null);

    // 셀 블록/표 선택
    const [cellSel, setCellSelState] = useState<CellSel | null>(null);
    const cellSelRef = useRef<CellSel | null>(null);
    const setCellSel = useCallback((v: CellSel | null) => { cellSelRef.current = v; setCellSelState(v); }, []);
    const [cellOverlay, setCellOverlay] = useState<CellOverlay>(EMPTY_OVERLAY);
    const cellClipRef = useRef<{ tsv: string; grid: string[][] } | null>(null);
    const scrollNextRef = useRef(false);
    const layoutCache = useRef<{ rev: number; map: Map<number, TableBox[]>; img: Map<number, ImageBox[]> }>({ rev: -1, map: new Map(), img: new Map() });

    // 그림 선택/크기 조절
    const [picSel, setPicSelState] = useState<PicRef | null>(null);
    const picSelRef = useRef<PicRef | null>(null);
    const setPicSel = useCallback((v: PicRef | null) => { picSelRef.current = v; setPicSelState(v); }, []);
    const [picBox, setPicBox] = useState<(Box & { page: number }) | null>(null);
    const [picGhost, setPicGhost] = useState<(Box & { page: number }) | null>(null);
    const [picDlg, setPicDlg] = useState<PicDlg | null>(null);
    const [tblDlg, setTblDlg] = useState<TblDlg | null>(null);
    const [tblGhost, setTblGhost] = useState<{ page: number; x: number; y: number; w: number; h: number } | null>(null);
    const tblDragRef = useRef<(MoveDrag & { t: TableRef }) | null>(null);
    const picMoveRef = useRef<(MoveDrag & { ref: PicRef }) | null>(null);
    const moveTableRef = useRef<(t: TableRef, dh: number, dv: number) => void>(() => {});
    const movePicRef = useRef<(ref: PicRef, dh: number, dv: number) => void>(() => {});
    const picDragRef = useRef<{ h: Handle; sx: number; sy: number; box: Box & { page: number }; ref: PicRef; wHU: number; hHU: number; last?: { w: number; h: number } } | null>(null);
    const zoomRef = useRef(1);
    zoomRef.current = zoom;
    const applyPicSizeRef = useRef<(ref: PicRef, w: number, h: number) => void>(() => {});

    // 편집 보조: 수정(덮어쓰기) 모드, 겹글쇠(Ctrl+N,…), 찾기/바꾸기
    const [overwrite, setOverwrite] = useState(false);
    const chordRef = useRef(0);
    const [findOpen, setFindOpen] = useState<null | "find" | "replace">(null);
    const [findQ, setFindQ] = useState("");
    const [replQ, setReplQ] = useState("");
    const [findCase, setFindCase] = useState(false);
    const [findMsg, setFindMsg] = useState("");
    const findInputRef = useRef<HTMLInputElement>(null);

    // 문서 로드
    useEffect(() => {
      let cancelled = false;
      (async () => {
        setLoadError(null);
        setLoading(true);
        try {
          await ensureWasm();
          if (cancelled) return;
          // 로딩 표시를 먼저 화면에 그린 뒤 무거운 파싱을 시작한다.
          await new Promise((r) => setTimeout(r, 0));
          if (cancelled) return;

          docRef.current?.free();
          docRef.current = null;
          setPageCount(0); setCaret(null); setSelRects([]); setVisiblePages(new Set([0])); setCellSel(null); setPicSel(null); setPicDlg(null); setTblDlg(null);

          let doc: HwpDocument;
          if (bytes) doc = new HwpDocument(bytes);
          else { doc = HwpDocument.createEmpty(); doc.createBlankDocument(); }
          if (cancelled) { doc.free(); return; }
          docRef.current = doc;
          // 이전 버전에서 표를 붙여넣고 저장해 표끼리 겹쳐 저장된 파일이면 위치를 다시 계산해 바로잡는다.
          if (bytes) {
            try {
              if (hasOverlappingTables(doc)) {
                for (let sec = 0; sec < doc.getSectionCount(); sec++) reflowControlParagraphs(doc, sec);
                onDirty();
                onError("겹쳐 저장돼 있던 표 위치를 바로잡았습니다. 저장하면 반영됩니다.");
              }
            } catch { /* 보정 실패 시 원본 그대로 표시 */ }
          }
          undoStack.current = []; redoStack.current = []; lastEditRef.current = null;
          setPageCount(doc.pageCount());
          setPos({ sec: 0, para: 0, off: 0 }); setAnchor(null);
          setRev((r) => r + 1);
          setLoading(false);
        } catch (e) {
          if (cancelled) return;
          const msg = `HWP 문서를 열 수 없습니다: ${errorText(e)}`;
          setLoadError(msg);
          setPageCount(0); setCaret(null); setLoading(false);
          onError(msg);
        }
      })();
      return () => { cancelled = true; };
    }, [bytes, reloadKey, onError, setCellSel, setPicSel]);

    // 화면에 보이는 페이지만 렌더링한다. (큰 문서에서 키 입력마다 전체 페이지를
    // 다시 그리던 지연을 없앤다.)
    useEffect(() => {
      const doc = docRef.current;
      if (!doc) return;
      // 고해상도(Windows 배율 125%/150% 등) 화면에서는 실제 픽셀 수만큼 그려야 표 테두리/글자가 흐려지거나
      // 어긋나 보이지 않는다. 캔버스 CSS 크기는 항상 (쪽 크기 × 확대율)로 고정해 좌표 계산(zoom)과 맞춘다.
      const dpr = Math.min(Math.max(window.devicePixelRatio || 1, 1), 3);
      for (let i = 0; i < pageCount; i++) {
        const c = canvases.current[i];
        if (!c) continue;
        if (!visiblePages.has(i)) {
          // 아직 그리지 않은 쪽도 자리를 미리 잡아 두어 스크롤/레이아웃이 튀지 않게 한다.
          try {
            const info = JSON.parse(doc.getPageInfo(i)) as { width: number; height: number };
            c.style.width = `${info.width * zoom}px`; c.style.height = `${info.height * zoom}px`;
          } catch { /* 쪽 정보를 못 읽으면 그리기 때 크기가 정해진다 */ }
          continue;
        }
        try {
          doc.renderPageToCanvas(i, c, zoom * dpr);
          c.style.width = `${c.width / dpr}px`;
          c.style.height = `${c.height / dpr}px`;
        } catch (e) { console.error(e); }
      }
    }, [rev, pageCount, zoom, visiblePages]);

    // 페이지 가시성 추적 (스크롤 시 새로 보이는 페이지만 렌더)
    useEffect(() => {
      const root = scrollRef.current;
      if (!root || pageCount === 0) return;
      const obs = new IntersectionObserver((entries) => {
        setVisiblePages((prev) => {
          const next = new Set(prev);
          let changed = false;
          for (const entry of entries) {
            const idx = Number((entry.target as HTMLElement).dataset.page);
            if (Number.isNaN(idx)) continue;
            if (entry.isIntersecting) {
              if (!next.has(idx)) { next.add(idx); changed = true; }
            } else if (next.delete(idx)) {
              changed = true;
            }
          }
          return changed ? next : prev;
        });
      }, { root, rootMargin: "200px 0px" });
      for (let i = 0; i < pageCount; i++) {
        const el = pageRefs.current[i];
        if (el) obs.observe(el);
      }
      return () => obs.disconnect();
    }, [pageCount]);

    // 기호 패널 바깥 클릭 시 닫기
    useEffect(() => {
      if (!symbolsOpen) return;
      const onDown = (e: MouseEvent) => {
        if (symbolRef.current && !symbolRef.current.contains(e.target as Node)) setSymbolsOpen(false);
      };
      document.addEventListener("mousedown", onDown);
      return () => document.removeEventListener("mousedown", onDown);
    }, [symbolsOpen]);

    // 페이지 설정 패널 바깥 클릭 시 닫기
    useEffect(() => {
      if (!pageOpen) return;
      const onDown = (e: MouseEvent) => {
        if (pagePanelRef.current && !pagePanelRef.current.contains(e.target as Node)) setPageOpen(false);
      };
      document.addEventListener("mousedown", onDown);
      return () => document.removeEventListener("mousedown", onDown);
    }, [pageOpen]);

    const selectionRange = useCallback((): SelRange | null => {
      if (!pos || !anchor) return null;
      if (anchor.sec !== pos.sec || !sameCell(anchor, pos)) return null;
      if (anchor.para === pos.para && anchor.off === pos.off) return null;
      return cmpPos(anchor, pos) <= 0 ? { start: anchor, end: pos } : { start: pos, end: anchor };
    }, [anchor, pos]);

    const charPropsAt = (d: HwpDocument, p: Pos) => {
      const raw = p.cell
        ? d.getCellCharPropertiesAtByPath(p.sec, p.cell.parentPara, cellPathJsonFor(p), p.off)
        : d.getCharPropertiesAt(p.sec, p.para, p.off);
      return JSON.parse(raw) as Record<string, unknown>;
    };

    const paraLength = (d: HwpDocument, p: Pos) => p.cell
      ? d.getCellParagraphLengthByPath(p.sec, p.cell.parentPara, cellPathJsonFor(p))
      : d.getParagraphLength(p.sec, p.para);

    const applyCharRange = (d: HwpDocument, p: Pos, from: number, to: number, propsJson: string) => {
      if (to <= from) return;
      if (p.cell) d.applyCharFormatInCellByPath(p.sec, p.cell.parentPara, cellPathJsonFor(p), from, to, propsJson);
      else d.applyCharFormat(p.sec, p.para, from, to, propsJson);
    };

    const applyParaToPos = (d: HwpDocument, p: Pos, propsJson: string) => {
      if (p.cell) {
        const last = p.cell.path[p.cell.path.length - 1];
        d.applyParaFormatInCell(p.sec, p.cell.parentPara, last.controlIndex, last.cellIndex, p.para, propsJson);
      } else {
        d.applyParaFormat(p.sec, p.para, propsJson);
      }
    };

    const deleteAt = (d: HwpDocument, p: Pos, off: number, count: number) => {
      if (count <= 0) return;
      if (p.cell) d.deleteTextInCellByPath(p.sec, p.cell.parentPara, cellPathJsonFor(p), off, count);
      else d.deleteText(p.sec, p.para, off, count);
    };

    const mergePara = (d: HwpDocument, p: Pos) => {
      if (p.cell) d.mergeParagraphInCellByPath(p.sec, p.cell.parentPara, cellPathJsonFor(p));
      else d.mergeParagraph(p.sec, p.para);
    };

    /** 선택 범위를 지운다. 여러 문단이면 문단을 합쳐 한 문단으로 만든다. */
    const deleteRange = (d: HwpDocument, r: SelRange) => {
      if (r.start.para === r.end.para) {
        deleteAt(d, r.start, r.start.off, r.end.off - r.start.off);
        return;
      }
      deleteAt(d, r.start, r.start.off, paraLength(d, r.start) - r.start.off);
      for (let par = r.start.para + 1; par < r.end.para; par++) {
        const p = { ...r.start, para: par };
        deleteAt(d, p, 0, paraLength(d, p));
      }
      deleteAt(d, r.end, 0, r.end.off);
      for (let par = r.end.para; par > r.start.para; par--) {
        mergePara(d, { ...r.start, para: par });
      }
    };

    /** 선택 범위의 각 문단에 글자 서식을 적용한다. */
    const applyCharToRange = (d: HwpDocument, r: SelRange, props: Record<string, unknown>) => {
      const propsJson = JSON.stringify(props);
      for (let par = r.start.para; par <= r.end.para; par++) {
        const p = { ...r.start, para: par };
        const from = par === r.start.para ? r.start.off : 0;
        const to = par === r.end.para ? r.end.off : paraLength(d, p);
        applyCharRange(d, p, from, to, propsJson);
      }
    };

    const pendingProps = () => pendingRef.current;
    const clearPending = () => { pendingRef.current = null; pendingPosRef.current = null; setPendingHint(false); };

    /** 커서/선택 위치의 서식 상태를 툴바에 반영한다. */
    useEffect(() => {
      const d = docRef.current;
      if (!d || !pos) return;
      const range = selectionRange();
      if (!range && pendingRef.current) return;
      try {
        const p = range ? range.start : pos;
        const c = charPropsAt(d, p);
        if (typeof c.fontSize === "number") setSize(String(Math.round(c.fontSize / 100)));
        if (typeof c.textColor === "string") setColor(c.textColor);
        if (typeof c.fontFamily === "string" && c.fontFamily) setFont(c.fontFamily);
        setBold(!!c.bold); setItalic(!!c.italic);
        setUnderline(!!c.underline); setStrike(!!c.strikethrough);
      } catch { /* 커서가 문서 밖이면 무시 */ }
    }, [pos, anchor, rev, selectionRange]);

    // 캐럿 위치 갱신
    useEffect(() => {
      const doc = docRef.current;
      if (!doc || !pos) { setCaret(null); return; }
      try {
        const raw = pos.cell
          ? doc.getCursorRectByPath(pos.sec, pos.cell.parentPara, cellPathJsonFor(pos), pos.off)
          : doc.getCursorRect(pos.sec, pos.para, pos.off);
        const r = JSON.parse(raw);
        setCaret({ page: r.pageIndex, x: r.x, y: r.y, h: r.height });
        if (scrollNextRef.current && !cellSelRef.current) { scrollNextRef.current = false; scrollToRect(r.pageIndex, r.x, r.y, r.height); }
      } catch { setCaret(null); }
    }, [pos, rev]);

    // 선택 영역 하이라이트
    useEffect(() => {
      const doc = docRef.current;
      const range = selectionRange();
      if (!doc || !range) { setSelRects((prev) => (prev.length ? [] : prev)); return; }
      try {
        const raw = range.start.cell
          ? doc.getSelectionRectsInCellByPath(
              range.start.sec, range.start.cell.parentPara, cellPathJsonFor(range.start),
              range.start.para, range.start.off, range.end.para, range.end.off)
          : doc.getSelectionRects(range.start.sec, range.start.para, range.start.off, range.end.para, range.end.off);
        const rects = JSON.parse(raw) as { pageIndex: number; x: number; y: number; width: number; height: number }[];
        setSelRects(rects.map((r) => ({ page: r.pageIndex, x: r.x, y: r.y, w: r.width, h: r.height })));
      } catch { setSelRects((prev) => (prev.length ? [] : prev)); }
    }, [anchor, pos, rev, selectionRange]);


    // 셀 블록 / 표 선택 표시 (파란 칸, 표 테두리, 포커스 점)
    useEffect(() => {
      const d = docRef.current;
      const cs = cellSel;
      if (!d || !cs) { setCellOverlay((o) => (o.cells.length || o.frames.length || o.focus ? EMPTY_OVERLAY : o)); return; }
      try {
        const info = selInfo(d, cs);
        if (!info) { setCellOverlay(EMPTY_OVERLAY); return; }
        const picked = new Set(info.sel.map((c) => c.idx));
        const cells: SelRect[] = [], frames: SelRect[] = [];
        let focus: SelRect | null = null;
        for (const pg of [...visiblePages].sort((a, b) => a - b)) {
          if (pg >= pageCount) continue;
          const lay = JSON.parse(d.getPageControlLayout(pg)) as { controls?: Record<string, unknown>[] };
          for (const c of lay.controls ?? []) {
            if (c.type !== "table" || Number(c.secIdx ?? 0) !== cs.t.sec || c.paraIdx !== cs.t.parent || Number(c.controlIdx ?? 0) !== cs.t.ctrl) continue;
            frames.push({ page: pg, x: Number(c.x), y: Number(c.y), w: Number(c.w), h: Number(c.h) });
            for (const cell of (c.cells as Record<string, number>[] | undefined) ?? []) {
              const r = { page: pg, x: cell.x, y: cell.y, w: cell.w, h: cell.h };
              if (picked.has(cell.cellIdx)) cells.push(r);
              if (cell.cellIdx === cs.f) focus = r;
            }
          }
        }
        setCellOverlay({ cells, frames, focus });
        if (scrollNextRef.current && cs.mode === "table") scrollNextRef.current = false;
        else if (scrollNextRef.current && focus) {
          scrollNextRef.current = false;
          const fr = focus as SelRect;
          scrollToRect(fr.page, fr.x, fr.y, fr.h);
        }
      } catch { setCellOverlay(EMPTY_OVERLAY); }
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [cellSel, rev, visiblePages, pageCount]);

    /** 편집 동작 래퍼: 스냅샷 → 실행 → 재렌더
     *  연속 타이핑은 하나의 undo 단위로 묶어 매 글자마다 스냅샷을 뜨지 않는다. */
    const edit = useCallback((fn: (d: HwpDocument) => void, kind: "typing" | "format" | "edit" = "edit") => {
      const d = docRef.current;
      if (!d) return;
      try {
        const now = performance.now();
        const last = lastEditRef.current;
        const coalesce = kind === "typing" && last?.kind === "typing" && now - last.at < 700;
        if (!coalesce) {
          undoStack.current.push(d.saveSnapshot());
          redoStack.current.forEach((id) => d.discardSnapshot(id)); redoStack.current = [];
        }
        lastEditRef.current = { kind, at: now };
        fn(d);
        setPageCount(d.pageCount());
        setRev((r) => r + 1);
        onDirty();
      } catch (e) { onError(String(e)); }
    }, [onDirty, onError]);

    const undo = () => {
      const d = docRef.current; const id = undoStack.current.pop();
      if (!d || id === undefined) return;
      lastEditRef.current = null;
      redoStack.current.push(d.saveSnapshot());
      d.restoreSnapshot(id); d.discardSnapshot(id);
      setCellSel(null); setPicSel(null); fixPos(d);
      setPageCount(d.pageCount()); setRev((r) => r + 1);
    };
    const redo = () => {
      const d = docRef.current; const id = redoStack.current.pop();
      if (!d || id === undefined) return;
      lastEditRef.current = null;
      undoStack.current.push(d.saveSnapshot());
      d.restoreSnapshot(id); d.discardSnapshot(id);
      setCellSel(null); setPicSel(null); fixPos(d);
      setPageCount(d.pageCount()); setRev((r) => r + 1);
    };

    /** 되돌리기/다시 실행 뒤 커서가 없어진 문단을 가리키면 문서 안쪽으로 되돌린다. */
    const fixPos = (d: HwpDocument) => {
      setPos((p) => {
        if (!p) return p;
        try {
          if (p.cell) {
            d.getCursorRectByPath(p.sec, p.cell.parentPara, cellPathJsonFor(p), p.off);
            return p;
          }
          const n = d.getParagraphCount(p.sec);
          if (p.para >= n) return { sec: p.sec, para: n - 1, off: d.getParagraphLength(p.sec, n - 1) };
          const len = d.getParagraphLength(p.sec, p.para);
          return p.off > len ? { ...p, off: len } : p;
        } catch { return { sec: 0, para: 0, off: 0 }; }
      });
      setAnchor(null);
    };

    const hitTestClient = useCallback((clientX: number, clientY: number): Pos | null => {
      const d = docRef.current;
      if (!d) return null;
      for (let i = 0; i < pageCount; i++) {
        const c = canvases.current[i];
        if (!c) continue;
        const r = c.getBoundingClientRect();
        if (clientX < r.left || clientX > r.right || clientY < r.top || clientY > r.bottom) continue;
        try {
          const h = JSON.parse(d.hitTest(i, (clientX - r.left) / zoom, (clientY - r.top) / zoom));
          return posFromHit(h);
        } catch { return null; }
      }
      return null;
    }, [pageCount, zoom]);

    // 드래그 선택 (캔버스 밖으로 나가도 계속 추적)
    useEffect(() => {
      const onMove = (e: MouseEvent) => {
        const drag = dragRef.current;
        if (!drag?.active) return;
        const hit = hitTestClient(e.clientX, e.clientY);
        if (!hit) return;
        // 같은 표 안에서 다른 셀로 끌면 글자 선택이 아니라 셀 블록(직사각형)으로 선택한다.
        const from = tableOfPos(drag.start), to = tableOfPos(hit);
        if (from && to && sameTable(from.t, to.t)) {
          if (from.cell !== to.cell) {
            drag.moved = true;
            setCellSel({ t: from.t, a: from.cell, f: to.cell, mode: "extend", back: drag.start });
            setAnchor(null); setPos(hit);
          } else if (cellSelRef.current) {
            setCellSel(null); setAnchor(drag.start); setPos(hit);
          } else if (hit.para !== drag.start.para || hit.off !== drag.start.off) { drag.moved = true; setPos(hit); }
          return;
        }
        if (cellSelRef.current) return; // 셀 블록을 끄는 중에는 표 밖 위치를 무시
        if (hit.para !== drag.start.para || hit.off !== drag.start.off) drag.moved = true;
        setPos(hit);
      };
      const onUp = () => {
        const drag = dragRef.current;
        dragRef.current = null;
        if (drag?.active && !drag.moved) { setAnchor(null); setSelRects([]); }
      };
      window.addEventListener("mousemove", onMove);
      window.addEventListener("mouseup", onUp);
      return () => {
        window.removeEventListener("mousemove", onMove);
        window.removeEventListener("mouseup", onUp);
      };
    }, [hitTestClient, setCellSel]);

    // 마우스 클릭/드래그 → 커서 이동 및 선택
    const onPageMouseDown = (page: number, e: React.MouseEvent<HTMLCanvasElement>) => {
      // mousedown 기본 동작이 textarea 포커스를 BODY로 되돌리므로 막는다.
      e.preventDefault();
      const d = docRef.current; if (!d) return;
      const rect = e.currentTarget.getBoundingClientRect();
      const x = (e.clientX - rect.left) / zoom, y = (e.clientY - rect.top) / zoom;
      // 그림을 누르면 그림을 개체로 선택한다. (손잡이로 크기 조절)
      const img = imageAt(page, x, y);
      if (img && !e.shiftKey) {
        clearPending(); dragRef.current = null;
        let floating = false;
        try { floating = !picProps(d, img.ref).treatAsChar; } catch { /* 무시 */ }
        if (floating) picMoveRef.current = { ref: img.ref, page, sx: e.clientX, sy: e.clientY, box: { x: img.x, y: img.y, w: img.w, h: img.h }, dx: 0, dy: 0 };
        selectPicture(img.ref);
        inputRef.current?.focus();
        return;
      }
      if (picSelRef.current) { setPicSel(null); setPicDlg(null); }
      // 표 바깥 테두리를 누르면 표 전체를 개체로 선택한다.
      const border = tableBorderAt(page, x, y);
      if (border && !e.shiftKey) {
        clearPending(); dragRef.current = null;
        selectTableObject(border.t);
        // 떠 있는 표(글자처럼 취급 아님)는 테두리를 끌어 옮긴다.
        if (isFloatingTable(border.t)) {
          tblDragRef.current = { t: border.t, page, sx: e.clientX, sy: e.clientY, box: { x: border.x, y: border.y, w: border.w, h: border.h }, dx: 0, dy: 0 };
        }
        inputRef.current?.focus();
        return;
      }
      let p: Pos;
      try {
        p = posFromHit(JSON.parse(d.hitTest(page, x, y)));
      } catch (err) { console.warn(err); return; }
      // 세 번 클릭: 문단 전체 선택
      if (e.detail >= 3) {
        try {
          const len = paraLength(d, p);
          setCellSel(null); dragRef.current = null; clearPending();
          setAnchor({ ...p, off: 0 }); setPos({ ...p, off: len });
        } catch { /* 무시 */ }
        inputRef.current?.focus();
        return;
      }
      // Shift+클릭: 셀 블록이 있으면 클릭한 셀까지 넓힌다.
      const csNow = cellSelRef.current;
      const clicked = tableOfPos(p);
      if (csNow && e.shiftKey && clicked && sameTable(clicked.t, csNow.t)) {
        setCellSel({ ...csNow, f: clicked.cell, mode: "extend" });
        setPos(p); dragRef.current = null; inputRef.current?.focus();
        return;
      }
      if (csNow) setCellSel(null);

      const pendingPos = pendingPosRef.current;
      const keepPending = !e.shiftKey && !!pendingRef.current && !!pendingPos
        && p.sec === pendingPos.sec && p.para === pendingPos.para && p.off === pendingPos.off
        && sameCell(p, pendingPos);
      if (!keepPending) clearPending();
      preferredXRef.current = -1;
      lastEditRef.current = null;
      if (e.shiftKey && pos) {
        // Shift+클릭: 기존 커서에서 새 위치까지 선택
        if (!anchor) setAnchor(pos);
        setPos(p);
        dragRef.current = null;
      } else {
        setAnchor(p);
        setPos(p);
        dragRef.current = { active: true, start: p, moved: false };
      }
      inputRef.current?.focus();
    };

    // 더블클릭 → 단어 선택
    const onPageDoubleClick = (page: number, e: React.MouseEvent<HTMLCanvasElement>) => {
      e.preventDefault();
      {
        const r0 = e.currentTarget.getBoundingClientRect();
        const im = imageAt(page, (e.clientX - r0.left) / zoom, (e.clientY - r0.top) / zoom);
        if (im) { selectPicture(im.ref); openPicDialog(im.ref); return; }
      }
      if (cellSelRef.current) setCellSel(null);
      const d = docRef.current; if (!d) return;
      const rect = e.currentTarget.getBoundingClientRect();
      let p: Pos;
      try {
        p = posFromHit(JSON.parse(d.hitTest(page, (e.clientX - rect.left) / zoom, (e.clientY - rect.top) / zoom)));
      } catch { return; }
      try {
        const text = p.cell
          ? d.getTextInCellByPath(p.sec, p.cell.parentPara, cellPathJsonFor(p), 0, 100000)
          : d.getTextRange(p.sec, p.para, 0, 100000);
        const re = /[A-Za-z0-9가-힣_]+/g;
        let m: RegExpExecArray | null;
        while ((m = re.exec(text))) {
          const start = m.index, end = start + m[0].length;
          if (p.off >= start && p.off <= end) {
            setAnchor({ ...p, off: start });
            setPos({ ...p, off: end });
            break;
          }
        }
      } catch { /* 단어 선택 실패 시 커서만 유지 */ }
      inputRef.current?.focus();
    };

    const insert = (text: string) => {
      if (!pos || !text) return;
      const range = selectionRange();
      scrollNextRef.current = true;
      edit((d) => {
        let p = pos;
        if (range) { deleteRange(d, range); p = range.start; }
        const charCount = [...text].length;
        if (overwrite && !range) { // 수정(덮어쓰기) 모드: 입력한 글자 수만큼 뒤 글자를 지운다
          const room = paraLength(d, p) - p.off;
          if (room > 0) deleteAt(d, p, p.off, Math.min(charCount, room));
        }
        const plain = !/[\r\n\t]/.test(text) && charCount <= 8;
        if (/[\r\n]/.test(text)) {
          // 한글처럼 줄바꿈은 문단 나누기로 넣는다. (rhwp insertText는 문단 안 줄바꿈 문자로 넣어 버린다)
          const lines = text.replace(/\r\n?/g, "\n").split("\n");
          let cur: Pos = p;
          lines.forEach((ln, k) => {
            if (k > 0) {
              if (cur.cell) d.splitParagraphInCellByPath(cur.sec, cur.cell.parentPara, cellPathJsonFor(cur), cur.off);
              else d.splitParagraph(cur.sec, cur.para, cur.off, undefined);
              cur = { ...cur, para: cur.para + 1, off: 0 };
            }
            if (ln) {
              if (cur.cell) d.insertTextInCellByPath(cur.sec, cur.cell.parentPara, cellPathJsonFor(cur), cur.off, ln);
              else d.insertText(cur.sec, cur.para, cur.off, ln);
              cur = { ...cur, off: cur.off + [...ln].length };
            }
          });
          reflowControlParagraphs(d, p.sec);
          setPos(cur); setAnchor(null); setSelRects([]);
          return;
        }
        if (!range && !p.cell && plain) {
          // 한 글자/짧은 IME 커밋은 페이지 전체 재배치 없이 지역 편집한다.
          try { d.replaceBodyTextLocal(p.sec, p.para, p.off, 0, text); }
          catch { d.insertText(p.sec, p.para, p.off, text); }
        } else if (!range && p.cell && p.cell.path.length === 1 && plain) {
          const c = p.cell.path[0];
          try {
            d.insertTextInCellDeferredPagination(
              p.sec, p.cell.parentPara, c.controlIndex, c.cellIndex, p.para, p.off, text);
          } catch {
            d.insertTextInCellByPath(p.sec, p.cell.parentPara, cellPathJsonFor(p), p.off, text);
          }
        } else if (p.cell) d.insertTextInCellByPath(p.sec, p.cell.parentPara, cellPathJsonFor(p), p.off, text);
        else d.insertText(p.sec, p.para, p.off, text);
        const pending = pendingProps();
        if (pending) applyCharRange(d, p, p.off, p.off + charCount, JSON.stringify(pending));
        const newPos = { ...p, off: p.off + charCount };
        if (pending) pendingPosRef.current = newPos;
        setPos(newPos); setAnchor(null); setSelRects([]);
      }, "typing");
    };

    /** textarea에 커밋된 텍스트를 정확히 한 번만 반영한다. (IME 중복 입력 방지) */
    const commitTextarea = (el: HTMLTextAreaElement) => {
      const text = el.value;
      el.value = "";
      if (text) insert(text);
    };

    // 복사 / 잘라내기 / 붙여넣기 (Tauri 네이티브 클립보드)
    const writeClipboard = async (text: string, html: string) => {
      try {
        if (html) await writeHtml(html, text);
        else await writeText(text);
      } catch {
        try { await navigator.clipboard?.writeText(text); } catch { /* 무시 */ }
      }
    };

    const copyInternal = (): { text: string; html: string } | null => {
      const d = docRef.current;
      if (!d) return null;
      const range = selectionRange();
      let text = "";
      let html = "";
      try {
        if (range) {
          if (range.start.cell) {
            d.copySelectionInCellByPath(
              range.start.sec, range.start.cell.parentPara, cellPathJsonFor(range.start),
              range.start.para, range.start.off, range.end.para, range.end.off);
            try {
              html = d.exportSelectionInCellHtmlByPath(
                range.start.sec, range.start.cell.parentPara, cellPathJsonFor(range.start),
                range.start.para, range.start.off, range.end.para, range.end.off);
            } catch { /* HTML 실패 시 텍스트만 */ }
          } else {
            d.copySelection(range.start.sec, range.start.para, range.start.off, range.end.para, range.end.off);
            try {
              html = d.exportSelectionHtml(
                range.start.sec, range.start.para, range.start.off, range.end.para, range.end.off);
            } catch { /* HTML 실패 시 텍스트만 */ }
          }
          text = d.getClipboardText();
        } else if (pos?.cell) {
          // 선택 없이 표 안에 있으면 표 전체를 복사한다.
          const path = pos.cell.path;
          const last = path[path.length - 1];
          const parentPath = path.slice(0, -1);
          const parentJson = parentPath.length ? JSON.stringify(parentPath) : "[]";
          d.copyControl(pos.sec, pos.cell.parentPara, parentJson, last.controlIndex);
          text = d.getClipboardText() || "[표]";
          try { html = d.exportControlHtml(pos.sec, pos.cell.parentPara, parentJson, last.controlIndex); } catch { /* HTML 실패 시 텍스트만 */ }
        } else {
          return null;
        }
        return { text, html };
      } catch (err) {
        onError(`복사 실패: ${errorText(err)}`);
        return null;
      }
    };

    const copySelectionToClipboard = async () => {
      const data = copyInternal();
      if (data) await writeClipboard(data.text, data.html);
    };

    const cutSelectionToClipboard = async () => {
      const range = selectionRange();
      if (!range) { await copySelectionToClipboard(); return; }
      const data = copyInternal();
      if (data) await writeClipboard(data.text, data.html);
      edit((d) => {
        deleteRange(d, range);
        setPos(range.start); setAnchor(null);
      });
    };

    const pasteFromClipboard = async () => {
      const d = docRef.current;
      if (!d || !pos) return;
      let osText = "";
      try { osText = await readText(); }
      catch { try { osText = await navigator.clipboard?.readText() ?? ""; } catch { /* 무시 */ } }
      if (!osText) return;
      // 셀 블록으로 복사한 내용이면 셀 칸에 맞춰 채워 넣는다.
      const grid = clipGrid(osText);
      const csNow = cellSelRef.current;
      if (grid) {
        const single = grid.length === 1 && grid[0].length === 1;
        if (csNow || !single) { pasteCellGrid(grid); return; }
      }
      if (csNow) {
        // 셀 블록에 일반 글을 붙이면 맨 앞 셀의 내용을 바꾼다.
        if (d.hasInternalClipboard() && d.clipboardHasControl()) { onError("셀 안에는 표를 붙여넣을 수 없습니다"); return; }
        const info = selInfo(d, csNow);
        if (info) edit((doc) => { setCellText(doc, csNow.t, info.sel[0].idx, osText.replace(/\r\n/g, "\n")); });
        return;
      }
      const internalText = d.getClipboardText();
      const norm = (s: string) => s.replace(/\r\n/g, "\n");
      const useInternal = d.hasInternalClipboard() && norm(osText) === norm(internalText);
      try {
        if (useInternal) {
          const range = selectionRange();
          edit((doc) => {
            let p = pos;
            if (range) { deleteRange(doc, range); p = range.start; }
            if (doc.clipboardHasControl()) {
              if (p.cell) { onError("표는 본문에만 붙여넣을 수 있습니다"); return; }
              const r = JSON.parse(doc.pasteControl(p.sec, p.para, p.off));
              reflowControlParagraphs(doc, p.sec); // 뒤따르는 표가 옛 위치에 남아 겹치는 것을 막는다
              setPos({ sec: p.sec, para: (r.paraIdx ?? p.para) + 1, off: 0 });
              setAnchor(null);
            } else if (p.cell) {
              const r = JSON.parse(doc.pasteInternalInCellByPath(
                p.sec, p.cell.parentPara, cellPathJsonFor(p), p.off));
              setPos({ ...p, off: r.charOffset ?? p.off });
              setAnchor(null);
            } else {
              const r = JSON.parse(doc.pasteInternal(p.sec, p.para, p.off));
              setPos({ sec: p.sec, para: r.paraIdx ?? p.para, off: r.charOffset ?? p.off });
              setAnchor(null);
            }
          });
        } else {
          insert(osText);
        }
      } catch (err) { onError(`붙여넣기 실패: ${errorText(err)}`); }
    };

    const selectAll = () => {
      const d = docRef.current;
      if (!d || !pos) return;
      clearPending();
      if (pos.cell) {
        const path = cellPathJsonFor(pos);
        const count = d.getCellParagraphCountByPath(pos.sec, pos.cell.parentPara, path);
        if (count <= 0) return;
        const last = { ...pos, para: count - 1 };
        const len = d.getCellParagraphLengthByPath(pos.sec, pos.cell.parentPara, cellPathJsonFor(last));
        // 셀 글이 이미 전부 선택돼 있거나 비어 있으면 Ctrl+A 한 번 더 → 표 전체 선택
        const tp = tableOfPos(pos);
        const sr = selectionRange();
        const coversCell = sr ? (sr.start.para === 0 && sr.start.off === 0 && sr.end.para === count - 1 && sr.end.off >= len) : (count === 1 && len === 0);
        if (tp && coversCell) { selectTableObject(tp.t, pos); return; }
        setAnchor({ ...pos, para: 0, off: 0 });
        setPos({ ...last, off: len });
      } else {
        const count = d.getParagraphCount(pos.sec);
        if (count <= 0) return;
        const len = d.getParagraphLength(pos.sec, count - 1);
        setAnchor({ sec: pos.sec, para: 0, off: 0 });
        setPos({ sec: pos.sec, para: count - 1, off: len });
      }
      inputRef.current?.focus();
    };

    const paraCount = (d: HwpDocument, p: Pos) => p.cell
      ? d.getCellParagraphCountByPath(p.sec, p.cell.parentPara, cellPathJsonFor(p))
      : d.getParagraphCount(p.sec);

    const applyCaret = (next: Pos, extend: boolean) => {
      if (!pos) return;
      if (extend) { if (!anchor) setAnchor(pos); }
      else setAnchor(null);
      lastEditRef.current = null;
      setPos(next);
    };

    /** 좌우 이동. 문단/셀 경계를 넘어가면 앞/뒤 문단으로 이동한다. */
    const moveHorizontal = (delta: number, extend: boolean) => {
      const d = docRef.current;
      if (!d || !pos) return;
      let next = pos;
      if (delta < 0) {
        if (pos.off > 0) next = { ...pos, off: pos.off - 1 };
        else if (pos.para > 0) {
          const prev = { ...pos, para: pos.para - 1 };
          next = { ...prev, off: paraLength(d, prev) };
        }
      } else {
        if (pos.off < paraLength(d, pos)) next = { ...pos, off: pos.off + 1 };
        else if (pos.para + 1 < paraCount(d, pos)) next = { ...pos, para: pos.para + 1, off: 0 };
      }
      if (next.para === pos.para && next.off === pos.off) return;
      preferredXRef.current = -1;
      applyCaret(next, extend);
    };

    /** 상하 이동. rhwp의 moveVertical을 사용해 줄/문단/페이지 경계를 자연스럽게 넘는다. */
    const moveCaretVertical = (delta: number, extend: boolean) => {
      const d = docRef.current;
      if (!d || !pos) return;
      try {
        const res = pos.cell
          ? JSON.parse(d.moveVerticalByPath(
              pos.sec, pos.cell.parentPara, cellPathJsonFor(pos), pos.off, delta, preferredXRef.current))
          : JSON.parse(d.moveVertical(
              pos.sec, pos.para, pos.off, delta, preferredXRef.current,
              0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF));
        if (typeof res.preferredX === "number") preferredXRef.current = res.preferredX;
        let next: Pos;
        if (pos.cell) {
          const last = pos.cell.path.length - 1;
          next = {
            sec: res.sectionIndex ?? pos.sec,
            para: res.paragraphIndex ?? res.cellParaIndex ?? pos.para,
            off: res.charOffset ?? pos.off,
            cell: {
              parentPara: res.parentParaIndex ?? pos.cell.parentPara,
              path: pos.cell.path.map((entry, i) => i === last ? {
                ...entry,
                controlIndex: res.controlIndex ?? entry.controlIndex,
                cellIndex: res.cellIndex ?? entry.cellIndex,
                cellParaIndex: res.cellParaIndex ?? entry.cellParaIndex,
              } : entry),
            },
          };
        } else {
          next = {
            sec: res.sectionIndex ?? pos.sec,
            para: res.paragraphIndex ?? pos.para,
            off: res.charOffset ?? pos.off,
          };
        }
        applyCaret(next, extend);
      } catch (err) { console.warn(err); }
    };


    /** Ctrl+←/→ : 단어 단위 이동 */
    const moveWord = (dir: -1 | 1, extend: boolean) => {
      const d = docRef.current;
      if (!d || !pos) return;
      let text = "";
      try {
        text = pos.cell
          ? d.getTextInCellByPath(pos.sec, pos.cell.parentPara, cellPathJsonFor(pos), 0, 100000)
          : d.getTextRange(pos.sec, pos.para, 0, 100000);
      } catch { moveHorizontal(dir, extend); return; }
      // 그림 같은 개체가 끼어 있으면 글 길이와 문단 길이가 달라 위치를 믿을 수 없으므로 한 글자씩 이동한다.
      if (text.length !== paraLength(d, pos)) { moveHorizontal(dir, extend); return; }
      const o = dir > 0 ? nextWordOffset(text, pos.off) : prevWordOffset(text, pos.off);
      if (o < 0) { moveHorizontal(dir, extend); return; } // 문단 경계는 이웃 문단으로
      preferredXRef.current = -1;
      applyCaret({ ...pos, off: o }, extend);
    };

    /** PageUp/PageDown : 화면 한 쪽(보이는 높이)만큼 이동하고 화면도 같이 넘긴다. */
    const movePage = (dir: -1 | 1, extend: boolean) => {
      const d = docRef.current, root = scrollRef.current;
      if (!d || !root || !pos) return;
      const cur = caret;
      if (!cur) { moveCaretVertical(dir * 10, extend); return; }
      const el = pageRefs.current[cur.page];
      if (!el) return;
      const dy = dir * Math.max(120, root.clientHeight - 80);
      let ty = el.offsetTop + cur.y * zoom + dy;
      let target = -1;
      for (let i = 0; i < pageCount; i++) {
        const pe = pageRefs.current[i];
        if (!pe) continue;
        if (ty < pe.offsetTop) { target = i; ty = pe.offsetTop + 2; break; }
        if (ty < pe.offsetTop + pe.offsetHeight) { target = i; break; }
      }
      if (target < 0) { moveDocEdge(dir > 0, extend); root.scrollTop += dy; return; }
      const te = pageRefs.current[target]!;
      try {
        const hit = posFromHit(JSON.parse(d.hitTest(target, cur.x, Math.max(0, (ty - te.offsetTop) / zoom))));
        preferredXRef.current = -1;
        applyCaret(hit, extend);
        root.scrollTop = Math.max(0, root.scrollTop + dy);
      } catch (err) { console.warn(err); }
    };

    /** 현재 줄의 시작/끝 오프셋. Home/End 및 Shift+Home/End에 사용한다. */
    const lineBounds = (d: HwpDocument, p: Pos) => {
      try {
        const last = p.cell?.path[p.cell.path.length - 1];
        const li = p.cell
          ? JSON.parse(d.getLineInfoInCell(p.sec, p.cell.parentPara, last!.controlIndex, last!.cellIndex, p.para, p.off))
          : JSON.parse(d.getLineInfo(p.sec, p.para, p.off));
        return { start: Number(li.char_start ?? li.charStart ?? p.off), end: Number(li.char_end ?? li.charEnd ?? p.off) };
      } catch { return null; }
    };

    const moveLineEdge = (toEnd: boolean, extend: boolean) => {
      const d = docRef.current;
      if (!d || !pos) return;
      const b = lineBounds(d, pos);
      if (!b) return;
      const off = toEnd ? b.end : b.start;
      if (off === pos.off) return;
      preferredXRef.current = -1;
      applyCaret({ ...pos, off }, extend);
    };

    const moveDocEdge = (toEnd: boolean, extend: boolean) => {
      const d = docRef.current;
      if (!d || !pos) return;
      preferredXRef.current = -1;
      if (pos.cell) {
        const count = d.getCellParagraphCountByPath(pos.sec, pos.cell.parentPara, cellPathJsonFor(pos));
        const target = { ...pos, para: toEnd ? Math.max(0, count - 1) : 0 };
        const len = toEnd ? d.getCellParagraphLengthByPath(pos.sec, pos.cell.parentPara, cellPathJsonFor(target)) : 0;
        applyCaret({ ...target, off: len }, extend);
      } else {
        const count = d.getParagraphCount(pos.sec);
        const target = { sec: pos.sec, para: toEnd ? Math.max(0, count - 1) : 0, off: 0 };
        if (toEnd) target.off = d.getParagraphLength(pos.sec, target.para);
        applyCaret(target, extend);
      }
    };



    // ───────────────────────── 그림: 선택 / 크기 조절 / 속성 ─────────────────────────

    /** 그림 컨트롤이 놓인 문단 안의 글자 위치 */
    const ctrlOffsetOf = (d: HwpDocument, ref: PicRef) => {
      try { return Number((JSON.parse(d.getControlTextPositions(ref.sec, ref.para)) as number[])[ref.ctrl] ?? 0); } catch { return 0; }
    };
    const picProps = (d: HwpDocument, ref: PicRef) => JSON.parse(d.getPictureProperties(ref.sec, ref.para, ref.ctrl)) as Record<string, unknown>;

    const selectPicture = (ref: PicRef) => {
      const d = docRef.current;
      if (!d) return;
      clearPending(); setCellSel(null); setAnchor(null); setSelRects([]);
      scrollNextRef.current = true;
      setPicSel(ref);
      setPos({ sec: ref.sec, para: ref.para, off: ctrlOffsetOf(d, ref) });
    };

    const applyPicSize = (ref: PicRef, wHU: number, hHU: number) => {
      edit((d) => {
        d.setPictureProperties(ref.sec, ref.para, ref.ctrl, JSON.stringify({ width: Math.max(300, Math.round(wHU)), height: Math.max(300, Math.round(hHU)) }));
        reflowControlParagraphs(d, ref.sec);
      });
    };
    applyPicSizeRef.current = applyPicSize;

    const deletePicture = (ref: PicRef) => {
      edit((d) => {
        const off = ctrlOffsetOf(d, ref);
        d.deletePictureControl(ref.sec, ref.para, ref.ctrl);
        reflowControlParagraphs(d, ref.sec);
        setPicSel(null); setPicDlg(null);
        setPos({ sec: ref.sec, para: ref.para, off: Math.min(off, d.getParagraphLength(ref.sec, ref.para)) });
      });
    };

    const copyPicture = async (ref: PicRef, cut: boolean) => {
      const d = docRef.current;
      if (!d) return;
      try {
        d.copyControl(ref.sec, ref.para, "[]", ref.ctrl);
        const text = d.getClipboardText() || "[그림]";
        let html = "";
        try { html = d.exportControlHtml(ref.sec, ref.para, "[]", ref.ctrl); } catch { /* 텍스트만 */ }
        await writeClipboard(text, html);
        if (cut) deletePicture(ref);
      } catch (err) { onError(`그림 복사 실패: ${errorText(err)}`); }
    };

    const openPicDialog = (ref: PicRef) => {
      const d = docRef.current;
      if (!d) return;
      try {
        const pr = picProps(d, ref);
        const w = Number(pr.width), h = Number(pr.height);
        const pl = placementFromProps(pr);
        setTblDlg(null);
        setPicDlg({ ref, wMm: String(picHuToMm(w)), hMm: String(picHuToMm(h)), keep: true, pl, plPrev: pl, ratio: h / w });
      } catch (err) { onError(`그림 속성을 읽을 수 없습니다: ${errorText(err)}`); }
    };

    const applyPicDlg = () => {
      if (!picDlg) return;
      const w = Number(picDlg.wMm), h = Number(picDlg.hMm);
      if (!(w > 0) || !(h > 0)) { onError("너비와 높이는 0보다 커야 합니다"); return; }
      const ref = picDlg.ref;
      edit((d) => {
        const cur = picProps(d, ref);
        const props: Record<string, unknown> = { width: picMmToHu(w), height: picMmToHu(h) };
        void cur;
        Object.assign(props, placementDiff(picDlg.pl, picDlg.plPrev));
        d.setPictureProperties(ref.sec, ref.para, ref.ctrl, JSON.stringify(props));
        reflowControlParagraphs(d, ref.sec);
      });
      setPicDlg(null);
    };

    /** 그림 손잡이를 잡고 끌기 시작 */
    const startPicResize = (e: React.MouseEvent, h: Handle) => {
      e.preventDefault(); e.stopPropagation();
      const d = docRef.current, ref = picSelRef.current;
      if (!d || !ref || !picBox) return;
      try {
        const pr = picProps(d, ref);
        picDragRef.current = { h, sx: e.clientX, sy: e.clientY, box: picBox, ref, wHU: Number(pr.width), hHU: Number(pr.height) };
        setPicGhost({ ...picBox });
      } catch (err) { console.warn(err); }
      inputRef.current?.focus();
    };

    useEffect(() => {
      const mv = (e: MouseEvent) => {
        const dr = picDragRef.current;
        if (!dr) return;
        const z = zoomRef.current;
        const r = resizeBox(dr.box, dr.h, (e.clientX - dr.sx) / z, (e.clientY - dr.sy) / z, e.shiftKey, e.altKey);
        dr.last = r;
        setPicGhost({ page: dr.box.page, x: dr.box.x, y: dr.box.y, w: r.w, h: r.h });
      };
      const up = () => {
        const dr = picDragRef.current;
        picDragRef.current = null;
        setPicGhost(null);
        if (dr?.last && (Math.abs(dr.last.w - dr.box.w) > 1 || Math.abs(dr.last.h - dr.box.h) > 1)) {
          applyPicSizeRef.current(dr.ref, dr.wHU * (dr.last.w / dr.box.w), dr.hHU * (dr.last.h / dr.box.h));
        }
      };
      window.addEventListener("mousemove", mv);
      window.addEventListener("mouseup", up);
      return () => { window.removeEventListener("mousemove", mv); window.removeEventListener("mouseup", up); };
    }, []);

    /** 떠 있는 개체를 옮길 때 쪽 밖/문단 위로 나가지 않게 이동량(HU)을 다듬는다 */
    const clampMove = (pr: Record<string, unknown>, dh: number, dv: number): [number, number] => {
      const vo = Number(pr.vertOffset ?? 0), ho = Number(pr.horzOffset ?? 0);
      if (vo + dv < 0) dv = -vo;
      const hr = String(pr.horzRelTo ?? "Para");
      if ((hr === "Paper" || hr === "Page") && ho + dh < 0) dh = -ho;
      return [Math.round(dh), Math.round(dv)];
    };
    movePicRef.current = (ref, dh, dv) => {
      edit((d) => {
        const pr = picProps(d, ref);
        if (pr.treatAsChar) return;
        const [h, v] = clampMove(pr, dh, dv);
        d.setPictureProperties(ref.sec, ref.para, ref.ctrl, JSON.stringify({ horzOffset: Number(pr.horzOffset ?? 0) + h, vertOffset: Number(pr.vertOffset ?? 0) + v }));
        reflowControlParagraphs(d, ref.sec);
      });
    };
    moveTableRef.current = (t, dh, dv) => {
      let np: { ppi?: number; ci?: number } = {};
      edit((d) => {
        const pr = JSON.parse(d.getTableProperties(t.sec, t.parent, t.ctrl)) as Record<string, unknown>;
        if (pr.treatAsChar) return;
        const [h, v] = clampMove(pr, dh, dv);
        if (!h && !v) return;
        np = JSON.parse(d.moveTableOffset(t.sec, t.parent, t.ctrl, h, v));
        reflowControlParagraphs(d, t.sec);
      });
      if (typeof np.ppi === "number" && typeof np.ci === "number" && (np.ppi !== t.parent || np.ci !== t.ctrl)) {
        const ppi = np.ppi, ci = np.ci;
        const cs = cellSelRef.current;
        if (cs && sameTable(cs.t, t)) setCellSel({ ...cs, t: { ...cs.t, parent: ppi, ctrl: ci } });
      }
    };

    /** 떠 있는 표/그림을 마우스로 끌어 옮기기 */
    useEffect(() => {
      const bound = (page: number, box: { x: number; y: number; w: number; h: number }, dx: number, dy: number) => {
        const cv = canvases.current[page];
        const z = zoomRef.current;
        const pw = cv ? cv.clientWidth / z : 1e6, ph = cv ? cv.clientHeight / z : 1e6;
        const x = Math.min(Math.max(0, box.x + dx), Math.max(0, pw - box.w));
        const y = Math.min(Math.max(0, box.y + dy), Math.max(0, ph - box.h));
        return { x, y, dx: x - box.x, dy: y - box.y };
      };
      const mv = (e: MouseEvent) => {
        const dr = tblDragRef.current ?? picMoveRef.current;
        if (!dr) return;
        const z = zoomRef.current;
        const b = bound(dr.page, dr.box, (e.clientX - dr.sx) / z, (e.clientY - dr.sy) / z);
        dr.dx = b.dx; dr.dy = b.dy;
        if (Math.abs(b.dx) > 2 || Math.abs(b.dy) > 2) setTblGhost({ page: dr.page, x: b.x, y: b.y, w: dr.box.w, h: dr.box.h });
      };
      const up = () => {
        const td = tblDragRef.current, pd = picMoveRef.current;
        tblDragRef.current = null; picMoveRef.current = null;
        setTblGhost(null);
        if (td && (Math.abs(td.dx) > 2 || Math.abs(td.dy) > 2)) moveTableRef.current(td.t, td.dx * HU_PER_PX, td.dy * HU_PER_PX);
        if (pd && (Math.abs(pd.dx) > 2 || Math.abs(pd.dy) > 2)) movePicRef.current(pd.ref, pd.dx * HU_PER_PX, pd.dy * HU_PER_PX);
      };
      window.addEventListener("mousemove", mv);
      window.addEventListener("mouseup", up);
      return () => { window.removeEventListener("mousemove", mv); window.removeEventListener("mouseup", up); };
    }, []);

    // 선택한 그림의 위치(프레임)
    useEffect(() => {
      const d = docRef.current;
      const ps = picSel;
      if (!d || !ps) { setPicBox(null); return; }
      try {
        picProps(d, ps);
        let found: (Box & { page: number }) | null = null;
        for (const pg of [...visiblePages].sort((a, b) => a - b)) {
          if (pg >= pageCount) continue;
          const im = imagesOnPage(pg).find((b) => b.ref.sec === ps.sec && b.ref.para === ps.para && b.ref.ctrl === ps.ctrl);
          if (im) { found = { page: pg, x: im.x, y: im.y, w: im.w, h: im.h }; break; }
        }
        setPicBox(found);
        if (found && scrollNextRef.current) { scrollNextRef.current = false; scrollToRect(found.page, found.x, found.y, found.h); }
      } catch { setPicSel(null); setPicBox(null); }
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [picSel, rev, visiblePages, pageCount]);

    /** 그림이 선택된 상태의 키 처리. 처리했으면 true */
    const handlePicKey = (e: React.KeyboardEvent<HTMLTextAreaElement>): boolean => {
      const d = docRef.current, ps = picSelRef.current;
      if (!d || !ps) return false;
      const key = e.key, mod = e.ctrlKey || e.metaKey;
      const stop = () => { e.preventDefault(); return true; };
      const off = ctrlOffsetOf(d, ps);
      const caretAt = (o: number) => { setPicSel(null); setPicDlg(null); gotoPos({ sec: ps.sec, para: ps.para, off: o }); };
      const STEP = Math.round(PIC_HU_PER_MM); // 1mm
      if (key === "Escape") { caretAt(off); return stop(); }
      if (key === "Delete" || key === "Backspace") { deletePicture(ps); return stop(); }
      if (key.startsWith("Arrow") && e.shiftKey && !mod) { // Shift+방향키: 크기
        try {
          const pr = picProps(d, ps);
          let w = Number(pr.width), h = Number(pr.height);
          if (key === "ArrowRight") w += STEP; else if (key === "ArrowLeft") w -= STEP;
          else if (key === "ArrowDown") h += STEP; else h -= STEP;
          applyPicSize(ps, Math.max(STEP, w), Math.max(STEP, h));
        } catch (err) { console.warn(err); }
        return stop();
      }
      if (key.startsWith("Arrow") && e.altKey && !mod) { // Alt+방향키: 위치 (글자처럼 취급이 아닐 때)
        try {
          const pr = picProps(d, ps);
          if (!pr.treatAsChar) {
            const props: Record<string, number> = {};
            if (key === "ArrowLeft") props.horzOffset = Number(pr.horzOffset) - STEP;
            else if (key === "ArrowRight") props.horzOffset = Number(pr.horzOffset) + STEP;
            else if (key === "ArrowUp") props.vertOffset = Number(pr.vertOffset) - STEP;
            else props.vertOffset = Number(pr.vertOffset) + STEP;
            edit((doc) => { doc.setPictureProperties(ps.sec, ps.para, ps.ctrl, JSON.stringify(props)); reflowControlParagraphs(doc, ps.sec); });
          }
        } catch (err) { console.warn(err); }
        return stop();
      }
      if (!mod && !e.altKey && e.shiftKey && (key === "Home" || key === "End")) { // 글 앞으로 / 글 뒤로
        try {
          const pr = picProps(d, ps);
          const w = key === "Home" ? "InFrontOfText" : "BehindText";
          const pl = placementFromProps(pr);
          edit((doc) => {
            doc.setPictureProperties(ps.sec, ps.para, ps.ctrl, JSON.stringify(placementDiff(pl.inline ? { ...floatingDefaults(pl), wrap: w } : { ...pl, wrap: w }, pl)));
            reflowControlParagraphs(doc, ps.sec);
          });
        } catch (err) { console.warn(err); }
        return stop();
      }
      if (!mod && !e.shiftKey && !e.altKey && key.startsWith("Arrow")) {
        try { // 떠 있는 그림은 방향키로 1mm씩 움직인다 (한글과 같음)
          if (!picProps(d, ps).treatAsChar) {
            movePicRef.current(ps, key === "ArrowLeft" ? -STEP : key === "ArrowRight" ? STEP : 0, key === "ArrowUp" ? -STEP : key === "ArrowDown" ? STEP : 0);
            return stop();
          }
        } catch { /* 아래에서 캐럿 이동 */ }
      }
      if (!mod && !e.shiftKey && !e.altKey) {
        if (key === "ArrowLeft" || key === "ArrowUp") { caretAt(off); return stop(); }
        if (key === "ArrowRight" || key === "ArrowDown" || key === "Enter") { caretAt(off + 1); return stop(); }
        if (e.code === "KeyP") { openPicDialog(ps); return stop(); }
      }
      if (mod && key.toLowerCase() === "c") { void copyPicture(ps, false); return stop(); }
      if (mod && key.toLowerCase() === "x") { void copyPicture(ps, true); return stop(); }
      if (mod && key.toLowerCase() === "v") { setPicSel(null); return false; }
      if (mod) return false;
      if (key === "Home" || key === "End" || key === "PageUp" || key === "PageDown") { caretAt(key === "Home" ? off : off + 1); return false; }
      if (key.length === 1 || key === "Process") { caretAt(off + 1); return false; } // 글자를 치면 그림 뒤에서 입력
      return false;
    };

    // ───────────────────────── 표: 셀 이동 / 셀 블록 / 표 선택 (한컴 한글 방식) ─────────────────────────

    /** 캐럿/선택이 화면 밖이면 보이도록 스크롤한다. (쪽 좌표 기준) */
    const scrollToRect = (page: number, x: number, y: number, h: number) => {
      const root = scrollRef.current, el = pageRefs.current[page];
      if (!root || !el) return;
      const top = el.offsetTop + y * zoom, bottom = top + h * zoom;
      const left = el.offsetLeft + x * zoom;
      if (top < root.scrollTop + 24) root.scrollTop = Math.max(0, top - 80);
      else if (bottom > root.scrollTop + root.clientHeight - 24) root.scrollTop = bottom - root.clientHeight + 80;
      if (left < root.scrollLeft + 8) root.scrollLeft = Math.max(0, left - 40);
    };

    /** 쪽에 놓인 표의 위치(쪽 좌표). 같은 화면(rev)에서는 캐시한다. */
    const tablesOnPage = (page: number): TableBox[] => {
      const d = docRef.current;
      if (!d) return [];
      const cache = layoutCache.current;
      if (cache.rev !== rev) { cache.rev = rev; cache.map.clear(); cache.img.clear(); }
      const hit = cache.map.get(page);
      if (hit) return hit;
      const list: TableBox[] = [];
      const imgs: ImageBox[] = [];
      try {
        const lay = JSON.parse(d.getPageControlLayout(page)) as { controls?: Record<string, unknown>[] };
        for (const c of lay.controls ?? []) {
          if (c.type === "image" && typeof c.paraIdx === "number") {
            const ref: PicRef = { sec: Number(c.secIdx ?? 0), para: Number(c.paraIdx), ctrl: Number(c.controlIdx ?? 0) };
            try { d.getPictureProperties(ref.sec, ref.para, ref.ctrl); } catch { continue; } // 표 안의 그림 등은 건너뜀
            imgs.push({ x: Number(c.x), y: Number(c.y), w: Number(c.w), h: Number(c.h), ref, z: Number(c.zOrder ?? 0) });
            continue;
          }
          if (c.type !== "table" || typeof c.paraIdx !== "number") continue;
          list.push({
            x: Number(c.x), y: Number(c.y), w: Number(c.w), h: Number(c.h),
            t: { sec: Number(c.secIdx ?? 0), parent: Number(c.paraIdx), ctrl: Number(c.controlIdx ?? 0) },
            cells: ((c.cells as Record<string, number>[] | undefined) ?? []).map((k) => ({ idx: k.cellIdx, row: k.row, col: k.col, x: k.x, w: k.w })),
          });
        }
      } catch { /* 레이아웃을 못 읽으면 표 테두리 선택은 건너뛴다 */ }
      cache.map.set(page, list);
      cache.img.set(page, imgs);
      return list;
    };

    /** 쪽 위의 그림들 (같은 화면에서는 캐시) */
    const imagesOnPage = (page: number): ImageBox[] => {
      tablesOnPage(page);
      return layoutCache.current.img.get(page) ?? [];
    };
    const imageAt = (page: number, x: number, y: number): ImageBox | null => {
      let best: ImageBox | null = null;
      for (const b of imagesOnPage(page)) {
        if (x >= b.x && x <= b.x + b.w && y >= b.y && y <= b.y + b.h && (!best || b.z >= best.z)) best = b;
      }
      return best;
    };

    /** (x, y)가 표의 바깥 테두리 근처(안쪽 4px ~ 바깥 4px)면 그 표를 돌려준다. */
    const tableBorderAt = (page: number, x: number, y: number): TableBox | null => {
      const tol = 4 / zoom;
      for (const b of tablesOnPage(page)) {
        const inside = x >= b.x - tol && x <= b.x + b.w + tol && y >= b.y - tol && y <= b.y + b.h + tol;
        const inner = x > b.x + tol && x < b.x + b.w - tol && y > b.y + tol && y < b.y + b.h - tol;
        if (inside && !inner) return b;
      }
      return null;
    };

    const onPageHover = (page: number, e: React.MouseEvent<HTMLCanvasElement>) => {
      if (e.buttons) return;
      const r = e.currentTarget.getBoundingClientRect();
      const hx = (e.clientX - r.left) / zoom, hy = (e.clientY - r.top) / zoom;
      const want = imageAt(page, hx, hy) ? "pointer" : tableBorderAt(page, hx, hy) ? "move" : "";
      if (e.currentTarget.style.cursor !== want) e.currentTarget.style.cursor = want;
    };


    // ───────────────────────── 표 개체 속성: 글자처럼 취급 / 본문과의 배치 / 위치 ─────────────────────────
    const tblProps = (d: HwpDocument, t: TableRef) => JSON.parse(d.getTableProperties(t.sec, t.parent, t.ctrl)) as Record<string, unknown>;
    const isFloatingTable = (t: TableRef) => {
      const d = docRef.current;
      if (!d) return false;
      try { return !tblProps(d, t).treatAsChar; } catch { return false; }
    };
    const openTableDialog = (t: TableRef) => {
      const d = docRef.current;
      if (!d) return;
      try {
        const pr = tblProps(d, t);
        const pl = placementFromProps(pr);
        setPicDlg(null);
        setTblDlg({ t, pl, plPrev: pl, wMm: picHuToMm(Number(pr.tableWidth ?? 0)), hMm: picHuToMm(Number(pr.tableHeight ?? 0)) });
      } catch (err) { onError(`표 속성을 읽을 수 없습니다: ${errorText(err)}`); }
    };
    const applyTableDlg = () => {
      if (!tblDlg) return;
      const { t, pl, plPrev } = tblDlg;
      const diff = placementDiff(pl, plPrev);
      if (Object.keys(diff).length) {
        edit((d) => {
          d.setTableProperties(t.sec, t.parent, t.ctrl, JSON.stringify(diff));
          reflowControlParagraphs(d, t.sec); // 표 위치가 바뀌면 뒤따르는 문단/표의 줄 위치도 다시 계산
        });
      }
      setTblDlg(null);
    };
    /** Shift+Home / Shift+End : 글 앞으로 / 글 뒤로 */
    const setTableWrap = (t: TableRef, wrap: string) => {
      const d = docRef.current;
      if (!d) return;
      try {
        const pl = placementFromProps(tblProps(d, t));
        const next = pl.inline ? { ...floatingDefaults(pl), wrap } : { ...pl, wrap };
        edit((doc) => {
          doc.setTableProperties(t.sec, t.parent, t.ctrl, JSON.stringify(placementDiff(next, pl)));
          reflowControlParagraphs(doc, t.sec);
        });
      } catch (err) { onError(`표 배치를 바꿀 수 없습니다: ${errorText(err)}`); }
    };

    /** 선택에 쓰는 셀 정보. 표가 사라졌거나 셀이 없으면 null */
    const selInfo = (d: HwpDocument, cs: CellSel) => {
      try {
        const cells = listCells(d, cs.t);
        const a = cellByIdx(cells, cs.a), f = cellByIdx(cells, cs.f);
        if (!a || !f) return null;
        const rect = cs.mode === "table" ? wholeRect(cells) : blockRect(cells, a, f);
        return { cells, a, f, rect, sel: cellsInRect(cells, rect) };
      } catch { return null; }
    };

    /** 캐럿이 놓인 셀 하나를 블록처럼 다루기 위한 정보 (셀 블록이 없을 때) */
    const targetBlock = (d: HwpDocument) => {
      const cs = cellSelRef.current;
      if (cs) {
        const info = selInfo(d, cs);
        return info ? { t: cs.t, ...info, cs } : null;
      }
      const tp = tableOfPos(pos);
      if (!tp) return null;
      try {
        const cells = listCells(d, tp.t);
        const cur = cellByIdx(cells, tp.cell);
        if (!cur) return null;
        const rect = blockRect(cells, cur, cur);
        return { t: tp.t, cells, a: cur, f: cur, rect, sel: cellsInRect(cells, rect), cs: null as CellSel | null };
      } catch { return null; }
    };

    /** 캐럿을 옮긴다. (셀 블록 해제, 선택 해제) */
    const gotoPos = (p: Pos) => {
      clearPending(); setCellSel(null); setAnchor(null); setSelRects([]);
      preferredXRef.current = -1; lastEditRef.current = null;
      scrollNextRef.current = true;
      setPos(p);
    };

    const cellStartPos = (t: TableRef, idx: number) => cellPosOf(t, idx, 0, 0);
    const cellEndPos = (d: HwpDocument, t: TableRef, idx: number) => {
      const last = Math.max(0, cellParaCount(d, t, idx) - 1);
      let len = 0;
      try { len = d.getCellParagraphLengthByPath(t.sec, t.parent, cellPathJson(t, idx, last)); } catch { /* 0 */ }
      return cellPosOf(t, idx, last, len);
    };

    /** 표 앞(위)/뒤(아래)의 본문 위치로 빠져나간다. */
    const exitTable = (t: TableRef, after: boolean) => {
      const d = docRef.current;
      if (!d) return;
      if (after) {
        if (t.parent + 1 < d.getParagraphCount(t.sec)) gotoPos({ sec: t.sec, para: t.parent + 1, off: 0 });
      } else if (t.parent > 0) {
        gotoPos({ sec: t.sec, para: t.parent - 1, off: d.getParagraphLength(t.sec, t.parent - 1) });
      } else gotoPos({ sec: t.sec, para: t.parent, off: 0 });
    };

    /** 표 전체를 개체로 선택한다. (Esc / 표 테두리 클릭 / F5 세 번) */
    const selectTableObject = (t: TableRef, back?: Pos) => {
      const d = docRef.current;
      if (!d) return;
      try {
        const cells = listCells(d, t);
        if (!cells.length) return;
        const first = cells[0], last = cells[cells.length - 1];
        const keep = back ?? (tableOfPos(pos) && sameTable(tableOfPos(pos)!.t, t) ? pos! : cellStartPos(t, first.idx));
        clearPending(); setAnchor(null); setSelRects([]);
        scrollNextRef.current = true;
        setCellSel({ t, a: first.idx, f: last.idx, mode: "table", back: keep });
        setPos(cellStartPos(t, first.idx));
      } catch (err) { console.warn(err); }
    };

    /** 셀 블록 선택을 끝내고 글자 입력 상태로 돌아간다. */
    const exitSel = (cs: CellSel, atFocus = false) => {
      const d = docRef.current;
      if (atFocus && d) gotoPos(cellEndPos(d, cs.t, cs.f));
      else gotoPos(cs.back);
    };

    /** 셀 하나의 글을 모두 지운다. (문단 서식은 첫 문단 것을 유지) */
    const clearCell = (d: HwpDocument, t: TableRef, idx: number) => {
      const n = cellParaCount(d, t, idx);
      for (let k = n - 1; k >= 0; k--) {
        const len = d.getCellParagraphLengthByPath(t.sec, t.parent, cellPathJson(t, idx, k));
        if (len > 0) d.deleteTextInCellByPath(t.sec, t.parent, cellPathJson(t, idx, k), 0, len);
      }
      for (let k = n - 1; k >= 1; k--) d.mergeParagraphInCellByPath(t.sec, t.parent, cellPathJson(t, idx, k));
    };

    const setCellText = (d: HwpDocument, t: TableRef, idx: number, text: string) => {
      clearCell(d, t, idx);
      const lines = text.split("\n");
      lines.forEach((ln, k) => {
        if (k > 0) {
          const prev = d.getCellParagraphLengthByPath(t.sec, t.parent, cellPathJson(t, idx, k - 1));
          d.splitParagraphInCellByPath(t.sec, t.parent, cellPathJson(t, idx, k - 1), prev);
        }
        if (ln) d.insertTextInCellByPath(t.sec, t.parent, cellPathJson(t, idx, k), 0, ln);
      });
    };

    const deleteTable = (t: TableRef) => {
      edit((doc) => {
        doc.deleteTableControl(t.sec, t.parent, t.ctrl);
        reflowControlParagraphs(doc, t.sec); // 뒤따르는 표/그림이 옛 위치에 남지 않게 한다
        const para = Math.max(0, Math.min(t.parent, doc.getParagraphCount(t.sec) - 1));
        setCellSel(null); setAnchor(null); setSelRects([]);
        setPos({ sec: t.sec, para, off: 0 });
      });
    };

    /** 표 편집 동작 (줄/칸 추가·삭제, 셀 합치기/나누기, 내용 지우기, 표 지우기) */
    const tableOp = (op: TableOp) => {
      const d = docRef.current;
      if (!d) return;
      const tb = targetBlock(d);
      if (!tb) return;
      const { t, cells, rect, sel, cs } = tb;
      const dims = tableInfoAt(d, t);
      if (!dims) return;
      const caretCell = tableOfPos(pos);
      /** (줄, 칸) 좌표가 바뀐 뒤에도 같은 셀을 가리키도록 옮긴다 */
      const remap = (doc: HwpDocument, shift: (row: number, col: number) => [number, number]) => {
        const nc = listCells(doc, t);
        const find = (c: CellInfo | undefined) => {
          if (!c) return undefined;
          const [r, k] = shift(c.row, c.col);
          return cellAt(nc, r, k);
        };
        if (cs) {
          const a = find(cellByIdx(cells, cs.a)), f = find(cellByIdx(cells, cs.f));
          if (a && f) {
            setCellSel({ ...cs, a: a.idx, f: f.idx, back: cellStartPos(t, a.idx) });
            setPos(cellStartPos(t, f.idx));
          } else setCellSel(null);
        } else if (caretCell && pos) {
          const c = find(cellByIdx(cells, caretCell.cell));
          if (c) setPos({ ...cellPosOf(t, c.idx, pos.para, pos.off) });
        }
      };
      const afterDelete = (doc: HwpDocument, row: number, col: number) => {
        const nc = listCells(doc, t);
        const c = cellAt(nc, Math.max(0, row), Math.max(0, col)) ?? nc[nc.length - 1];
        setCellSel(null); setAnchor(null); setSelRects([]);
        if (c) setPos(cellStartPos(t, c.idx));
      };
      switch (op) {
        case "delTable": deleteTable(t); return;
        case "clear":
          edit((doc) => { for (const c of sel) if (c.row >= rect.r1 && c.col >= rect.c1) clearCell(doc, t, c.idx); });
          return;
        case "merge":
          if (sel.length < 2) { onError("합칠 셀을 두 개 이상 블록으로 잡아 주세요 (F5 → 방향키, 또는 마우스로 끌기)"); return; }
          edit((doc) => {
            doc.mergeTableCells(t.sec, t.parent, t.ctrl, rect.r1, rect.c1, rect.r2, rect.c2);
            reflowControlParagraphs(doc, t.sec);
            const m = cellAt(listCells(doc, t), rect.r1, rect.c1);
            if (m) { setCellSel({ t, a: m.idx, f: m.idx, mode: "cell", back: cellStartPos(t, m.idx) }); setPos(cellStartPos(t, m.idx)); setAnchor(null); }
          });
          return;
        case "split": {
          const merged = sel.filter((c) => c.rs > 1 || c.cs > 1);
          if (merged.length) {
            edit((doc) => {
              for (const c of merged) doc.splitTableCell(t.sec, t.parent, t.ctrl, c.row, c.col);
              reflowControlParagraphs(doc, t.sec);
              const m = cellAt(listCells(doc, t), rect.r1, rect.c1);
              if (m) { setCellSel({ t, a: m.idx, f: m.idx, mode: "cell", back: cellStartPos(t, m.idx) }); setPos(cellStartPos(t, m.idx)); }
            });
            return;
          }
          const n = Number(prompt("나눌 줄 수", "2")), m = Number(prompt("나눌 칸 수", "1"));
          if (!(n >= 1) || !(m >= 1) || (n === 1 && m === 1)) return;
          edit((doc) => {
            doc.splitTableCellsInRange(t.sec, t.parent, t.ctrl, rect.r1, rect.c1, rect.r2, rect.c2, n, m, true);
            reflowControlParagraphs(doc, t.sec);
            const c = cellAt(listCells(doc, t), rect.r1, rect.c1);
            setCellSel(null); setAnchor(null);
            if (c) setPos(cellStartPos(t, c.idx));
          });
          return;
        }
        case "rowAbove":
          edit((doc) => { doc.insertTableRow(t.sec, t.parent, t.ctrl, rect.r1, false); reflowControlParagraphs(doc, t.sec); remap(doc, (r, k) => [r >= rect.r1 ? r + 1 : r, k]); });
          return;
        case "rowBelow":
          edit((doc) => { doc.insertTableRow(t.sec, t.parent, t.ctrl, rect.r2, true); reflowControlParagraphs(doc, t.sec); remap(doc, (r, k) => [r, k]); });
          return;
        case "colLeft":
          edit((doc) => { doc.insertTableColumn(t.sec, t.parent, t.ctrl, rect.c1, false); reflowControlParagraphs(doc, t.sec); remap(doc, (r, k) => [r, k >= rect.c1 ? k + 1 : k]); });
          return;
        case "colRight":
          edit((doc) => { doc.insertTableColumn(t.sec, t.parent, t.ctrl, rect.c2, true); reflowControlParagraphs(doc, t.sec); remap(doc, (r, k) => [r, k]); });
          return;
        case "delRow": {
          const n = rect.r2 - rect.r1 + 1;
          if (n >= dims.rows) { onError("표의 모든 줄을 지울 수 없습니다. 표를 지우려면 표를 선택(Esc)하고 Delete를 누르세요"); return; }
          edit((doc) => {
            for (let r = rect.r2; r >= rect.r1; r--) doc.deleteTableRow(t.sec, t.parent, t.ctrl, r);
            reflowControlParagraphs(doc, t.sec);
            afterDelete(doc, Math.min(rect.r1, dims.rows - n - 1), rect.c1);
          });
          return;
        }
        case "delCol": {
          const n = rect.c2 - rect.c1 + 1;
          if (n >= dims.cols) { onError("표의 모든 칸을 지울 수 없습니다. 표를 지우려면 표를 선택(Esc)하고 Delete를 누르세요"); return; }
          edit((doc) => {
            for (let c = rect.c2; c >= rect.c1; c--) doc.deleteTableColumn(t.sec, t.parent, t.ctrl, c);
            reflowControlParagraphs(doc, t.sec);
            afterDelete(doc, rect.r1, Math.min(rect.c1, dims.cols - n - 1));
          });
          return;
        }
      }
    };

    /** 셀 블록의 글자 서식(굵게 등)을 블록 안 모든 셀에 적용한다. */
    const applyFmtToBlock = (props: Record<string, unknown>) => {
      const d = docRef.current, cs = cellSelRef.current;
      if (!d || !cs) return;
      const info = selInfo(d, cs);
      if (!info) return;
      const first = info.sel[0];
      const next: Record<string, unknown> = { ...props };
      try {
        const cur = JSON.parse(d.getCellCharPropertiesAtByPath(cs.t.sec, cs.t.parent, cellPathJson(cs.t, first.idx, 0), 0)) as Record<string, unknown>;
        for (const k of TOGGLE_KEYS) if (k in next) next[k] = !cur[k];
      } catch { /* 현재 값을 못 읽으면 그대로 적용 */ }
      const fontName = typeof next.fontFamily === "string" ? next.fontFamily : null;
      if (fontName) { const id = d.findOrCreateFontId(fontName); delete next.fontFamily; if (id >= 0) next.fontId = id; }
      const json = JSON.stringify(next);
      edit((doc) => {
        for (const c of info.sel) {
          const n = cellParaCount(doc, cs.t, c.idx);
          for (let k = 0; k < n; k++) {
            const len = doc.getCellParagraphLengthByPath(cs.t.sec, cs.t.parent, cellPathJson(cs.t, c.idx, k));
            if (len > 0) doc.applyCharFormatInCellByPath(cs.t.sec, cs.t.parent, cellPathJson(cs.t, c.idx, k), 0, len, json);
          }
        }
      }, "format");
    };

    /** 셀 블록 복사/오려두기. 표 전체 선택이면 표 개체로 복사한다. */
    const copyCells = async (cs: CellSel, cut: boolean) => {
      const d = docRef.current;
      if (!d) return;
      if (cs.mode === "table") {
        cellClipRef.current = null;
        const data = copyInternal();
        if (data) await writeClipboard(data.text, data.html);
        if (cut) deleteTable(cs.t);
        return;
      }
      const info = selInfo(d, cs);
      if (!info) return;
      const grid = blockGrid(d, cs.t, info.cells, info.rect);
      const tsv = gridToTsv(grid);
      cellClipRef.current = { tsv, grid };
      try { d.clearClipboard(); } catch { /* 무시 */ }
      await writeClipboard(tsv, gridToHtml(grid));
      if (cut) tableOp("clear");
    };

    /** 복사해 둔 셀 격자를 셀 블록(또는 캐럿 셀)부터 채워 넣는다. 표 밖이면 새 표로 만든다. */
    /** 새 표: 한컴 한글처럼 "글자처럼 취급" 으로 만든다. (rhwp createTable 기본값은 떠 있는 표) */
    const createInlineTable = (doc: HwpDocument, sec: number, para: number, off: number, rows: number, cols: number) => {
      const r = JSON.parse(doc.createTable(sec, para, off, rows, cols)) as { paraIdx: number; controlIdx: number };
      try { doc.setTableProperties(sec, r.paraIdx, r.controlIdx, JSON.stringify({ treatAsChar: true })); }
      catch (err) { console.warn("글자처럼 취급 설정 실패", err); }
      return r;
    };

    const pasteCellGrid = (grid: string[][]) => {
      const rows = grid.length, cols = Math.max(...grid.map((r) => r.length));
      edit((doc) => {
        const cs = cellSelRef.current;
        let t: TableRef, r0 = 0, c0 = 0;
        const tp = tableOfPos(pos);
        if (cs) {
          const info = selInfo(doc, cs);
          if (!info) return;
          t = cs.t; r0 = info.rect.r1; c0 = info.rect.c1;
        } else if (tp) {
          const cur = cellByIdx(listCells(doc, tp.t), tp.cell);
          if (!cur) return;
          t = tp.t; r0 = cur.row; c0 = cur.col;
        } else if (pos && !pos.cell) {
          const r = createInlineTable(doc, pos.sec, pos.para, pos.off, rows, cols);
          t = { sec: pos.sec, parent: r.paraIdx, ctrl: r.controlIdx };
        } else { onError("여기에는 셀을 붙여넣을 수 없습니다"); return; }
        const nc = listCells(doc, t);
        for (let i = 0; i < rows; i++) {
          for (let j = 0; j < cols; j++) {
            const target = cellAt(nc, r0 + i, c0 + j);
            if (target && target.row === r0 + i && target.col === c0 + j) setCellText(doc, t, target.idx, grid[i][j] ?? "");
          }
        }
        if (!tp && !cs) { reflowControlParagraphs(doc, t.sec); setCellSel(null); setPos({ sec: t.sec, para: t.parent + 1, off: 0 }); }
      });
    };

    const clipGrid = (osText: string): string[][] | null => {
      const clip = cellClipRef.current;
      if (!clip) return null;
      return osText.replace(/\r\n/g, "\n") === clip.tsv ? clip.grid : null;
    };

    /** 셀 블록 선택 상태의 키 처리. 처리했으면 true */
    const handleCellSelKey = (e: React.KeyboardEvent<HTMLTextAreaElement>, cs: CellSel, d: HwpDocument): boolean => {
      const key = e.key, mod = e.ctrlKey || e.metaKey;
      const stop = () => { e.preventDefault(); return true; };
      const info = selInfo(d, cs);
      if (!info) { setCellSel(null); return false; }
      const { cells, f } = info;
      const setSel = (next: Partial<CellSel>) => {
        const n = { ...cs, ...next };
        scrollNextRef.current = true;
        setCellSel(n); setAnchor(null); setSelRects([]);
        setPos(cellStartPos(n.t, n.f));
      };
      const first = cells[0], last = cells[cells.length - 1];
      const lower = key.toLowerCase();

      if (key === "Escape") { exitSel(cs); return stop(); }
      if (key === "Enter" && !mod && !e.altKey) { exitSel(cs, true); return stop(); }
      if (key === "Enter" && mod) { tableOp("rowBelow"); return stop(); }
      if (key === "Backspace" && mod) { tableOp("delRow"); return stop(); }
      if (e.altKey && key === "Insert") { tableOp(e.shiftKey ? "colRight" : "rowBelow"); return stop(); }
      if (e.altKey && key === "Delete") { tableOp(e.shiftKey ? "delCol" : "delRow"); return stop(); }
      if (key === "F5" || key === "F3") {
        if (e.shiftKey || cs.mode === "cell") setSel({ mode: "extend" });
        else if (cs.mode === "extend") setSel({ mode: "table", a: first.idx, f: last.idx });
        else setSel({ mode: "cell", a: first.idx, f: first.idx });
        return stop();
      }
      if (key === "F7" || key === "F8") { // 칸 전체 / 줄 전체 블록
        const top = key === "F7" ? cellAt(cells, 0, f.col) : cellAt(cells, f.row, 0);
        const bottom = key === "F7" ? cellAt(cells, wholeRect(cells).r2, f.col) : cellAt(cells, f.row, wholeRect(cells).c2);
        if (top && bottom) setSel({ mode: "extend", a: top.idx, f: bottom.idx });
        return stop();
      }
      if (!mod && !e.altKey && !e.shiftKey && e.code === "KeyP") { openTableDialog(cs.t); return stop(); } // 개체 속성 / 표·셀 속성
      if (cs.mode === "table" && !mod && !e.altKey && e.shiftKey && (key === "Home" || key === "End")) {
        setTableWrap(cs.t, key === "Home" ? "InFrontOfText" : "BehindText");
        return stop();
      }
      if (cs.mode === "table" && !mod && !e.altKey && !e.shiftKey && key.startsWith("Arrow") && isFloatingTable(cs.t)) {
        const st = Math.round(PIC_HU_PER_MM); // 떠 있는 표는 방향키로 1mm씩 이동
        moveTableRef.current(cs.t, key === "ArrowLeft" ? -st : key === "ArrowRight" ? st : 0, key === "ArrowUp" ? -st : key === "ArrowDown" ? st : 0);
        return stop();
      }
      const dir: Dir | null = key === "ArrowLeft" ? "left" : key === "ArrowRight" ? "right" : key === "ArrowUp" ? "up" : key === "ArrowDown" ? "down" : null;
      if (dir) {
        if (mod || e.altKey) return stop();
        if (cs.mode === "table") return stop();
        const nb = neighborCell(cells, f, dir);
        if (!nb) return stop();
        if (e.shiftKey || cs.mode === "extend") setSel({ f: nb.idx, mode: "extend" });
        else setSel({ a: nb.idx, f: nb.idx });
        return stop();
      }
      if (key === "Tab" && !mod) {
        const i = cells.findIndex((c) => c.idx === f.idx);
        const n = cells[i + (e.shiftKey ? -1 : 1)];
        if (n) setSel({ a: n.idx, f: n.idx, mode: "cell" });
        return stop();
      }
      if (key === "Delete" || key === "Backspace") {
        if (cs.mode === "table") deleteTable(cs.t); else tableOp("clear");
        return stop();
      }
      if (mod && e.code === "KeyE") { tableOp("clear"); return stop(); }
      if (mod && lower === "c") { void copyCells(cs, false); return stop(); }
      if (mod && lower === "x") { void copyCells(cs, true); return stop(); }
      if (mod && lower === "v") { void pasteFromClipboard(); return stop(); }
      if (mod && lower === "a") { setSel({ mode: "table", a: first.idx, f: last.idx }); return stop(); }
      if (mod) return false; // 되돌리기, 굵게 등은 기존 처리
      if (!e.altKey && e.code === "KeyM") { tableOp("merge"); return stop(); }
      if (!e.altKey && e.code === "KeyS") { tableOp("split"); return stop(); }
      if (key === "Home" || key === "End" || key === "PageUp" || key === "PageDown") return stop();
      if (key.length === 1 || key === "Process") { exitSel(cs, true); return false; } // 글자를 치면 편집 상태로 돌아가 입력
      return false;
    };

    /** 본문 문단 para 에 놓인 표 (캐럿 가까운 쪽에서 찾는다) */
    const tableInPara = (sec: number, para: number): TableBox | null => {
      const cp = caret?.page ?? 0;
      for (const pg of [cp, cp - 1, cp + 1]) {
        if (pg < 0 || pg >= pageCount) continue;
        const b = tablesOnPage(pg).find((x) => x.t.sec === sec && x.t.parent === para);
        if (b) return b;
      }
      return null;
    };

    /** 본문에서 방향키로 표 속으로 들어간다. (표 위/아래 글에서 ↓ ↑ → ←) */
    const handleBodyKeyIntoTable = (e: React.KeyboardEvent<HTMLTextAreaElement>, d: HwpDocument): boolean => {
      if (!pos || pos.cell || e.ctrlKey || e.metaKey || e.altKey || e.shiftKey) return false;
      const key = e.key;
      if (key !== "ArrowDown" && key !== "ArrowUp" && key !== "ArrowRight" && key !== "ArrowLeft") return false;
      const len = d.getParagraphLength(pos.sec, pos.para);
      const lb = lineBounds(d, pos);
      const pick = (b: TableBox, last: boolean, horizontal = false) => {
        if (horizontal) { // ←/→ 로 들어갈 때는 맨 앞(첫 셀) / 맨 뒤(마지막 셀)
          const ord = [...b.cells].sort((p, q) => p.row - q.row || p.col - q.col);
          return last ? ord[ord.length - 1] : ord[0];
        }
        const rows = b.cells.map((c) => c.row);
        const row = last ? Math.max(...rows) : Math.min(...rows);
        const inRow = b.cells.filter((c) => c.row === row);
        const cx = caret?.x ?? 0;
        return inRow.find((c) => cx >= c.x && cx < c.x + c.w) ?? inRow[0];
      };
      const enter = (b: TableBox, last: boolean, end: boolean) => {
        const c = pick(b, last, key === "ArrowLeft" || key === "ArrowRight");
        if (!c) return false;
        gotoPos(end ? cellEndPos(d, b.t, c.idx) : cellStartPos(b.t, c.idx));
        return true;
      };
      // 표가 든 문단 맨 앞에서 ↓ → 로 표 안으로
      const here = tableInPara(pos.sec, pos.para);
      if (here && pos.off === 0 && (key === "ArrowDown" || key === "ArrowRight")) { if (enter(here, false, false)) { e.preventDefault(); return true; } }
      if (key === "ArrowDown" || key === "ArrowRight") {
        const atEnd = key === "ArrowRight" ? pos.off >= len : (!lb || lb.end >= len);
        const next = pos.para + 1 < d.getParagraphCount(pos.sec) ? tableInPara(pos.sec, pos.para + 1) : null;
        if (atEnd && next && !here) { if (enter(next, false, false)) { e.preventDefault(); return true; } }
      } else {
        const atStart = key === "ArrowLeft" ? pos.off === 0 : (!lb || lb.start <= 0);
        const prev = pos.para > 0 ? tableInPara(pos.sec, pos.para - 1) : null;
        if (atStart && prev) { if (enter(prev, true, key === "ArrowLeft")) { e.preventDefault(); return true; } }
      }
      return false;
    };

    /** 표 안(캐럿 상태)의 키 처리. 처리했으면 true */
    const handleTableKey = (e: React.KeyboardEvent<HTMLTextAreaElement>): boolean => {
      const d = docRef.current;
      if (!d || !pos) return false;
      const cs = cellSelRef.current;
      if (cs) return handleCellSelKey(e, cs, d);
      const tp = tableOfPos(pos);
      if (!tp) return handleBodyKeyIntoTable(e, d);
      const { t } = tp;
      const key = e.key, mod = e.ctrlKey || e.metaKey;
      const stop = () => { e.preventDefault(); return true; };

      if (key === "F5") {
        clearPending(); setAnchor(null); setSelRects([]);
        scrollNextRef.current = true;
        setCellSel({ t, a: tp.cell, f: tp.cell, mode: e.shiftKey ? "extend" : "cell", back: pos });
        return stop();
      }
      if (key === "Tab" && !mod && !e.altKey) {
        const cells = listCells(d, t);
        const i = cells.findIndex((c) => c.idx === tp.cell);
        if (i < 0) return false;
        if (e.shiftKey) { if (i > 0) gotoPos(cellStartPos(t, cells[i - 1].idx)); return stop(); }
        if (i + 1 < cells.length) { gotoPos(cellStartPos(t, cells[i + 1].idx)); return stop(); }
        // 마지막 셀에서 Tab: 한글처럼 새 줄을 하나 추가하고 첫 칸으로 이동
        const lastRow = Math.max(...cells.map((c) => c.row + c.rs - 1));
        edit((doc) => {
          doc.insertTableRow(t.sec, t.parent, t.ctrl, lastRow, true);
          reflowControlParagraphs(doc, t.sec);
          const nc = cellAt(listCells(doc, t), lastRow + 1, 0);
          if (nc) { setCellSel(null); setAnchor(null); setSelRects([]); scrollNextRef.current = true; setPos(cellStartPos(t, nc.idx)); }
        });
        return stop();
      }
      if (e.altKey && !mod && key.startsWith("Arrow")) { // Alt+방향키: 셀 단위 이동
        const dir: Dir = key === "ArrowLeft" ? "left" : key === "ArrowRight" ? "right" : key === "ArrowUp" ? "up" : "down";
        const cells = listCells(d, t);
        const cur = cellByIdx(cells, tp.cell);
        const nb = cur && neighborCell(cells, cur, dir);
        if (nb) gotoPos(cellStartPos(t, nb.idx));
        return stop();
      }
      if (mod && !e.altKey && key === "Enter") { tableOp("rowBelow"); return stop(); }
      if (mod && !e.altKey && key === "Backspace") { tableOp("delRow"); return stop(); }
      if (e.altKey && key === "Insert") { tableOp(e.shiftKey ? "colRight" : "rowBelow"); return stop(); }
      if (e.altKey && key === "Delete") { tableOp(e.shiftKey ? "delCol" : "delRow"); return stop(); }
      if (key === "Escape" && !mod) {
        if (selectionRange()) return false; // 글자 선택이 있으면 먼저 그것만 해제
        clearPending();
        selectTableObject(t, pos);
        return stop();
      }
      if (mod || e.altKey || e.shiftKey) return false;

      // 셀 경계에서 좌우/상하 방향키: 이웃 셀 → 표 밖으로
      const atCellEnd = pos.para >= paraCount(d, pos) - 1 && pos.off >= paraLength(d, pos);
      const atCellStart = pos.para === 0 && pos.off === 0;
      if (key === "ArrowRight" && atCellEnd) {
        const cells = listCells(d, t);
        const i = cells.findIndex((c) => c.idx === tp.cell);
        if (i >= 0 && i + 1 < cells.length) gotoPos(cellStartPos(t, cells[i + 1].idx)); else exitTable(t, true);
        return stop();
      }
      if (key === "ArrowLeft" && atCellStart) {
        const cells = listCells(d, t);
        const i = cells.findIndex((c) => c.idx === tp.cell);
        if (i > 0) gotoPos(cellEndPos(d, t, cells[i - 1].idx)); else exitTable(t, false);
        return stop();
      }
      if (key === "ArrowUp" || key === "ArrowDown") {
        const delta = key === "ArrowUp" ? -1 : 1;
        try {
          const res = JSON.parse(d.moveVerticalByPath(pos.sec, pos.cell!.parentPara, cellPathJsonFor(pos), pos.off, delta, preferredXRef.current));
          const samePlace = res.cellIndex === tp.cell && (res.cellParaIndex ?? res.paragraphIndex) === pos.para && res.charOffset === pos.off;
          if (samePlace) { exitTable(t, delta > 0); return stop(); } // 첫 줄에서 위로 / 마지막 줄에서 아래로
        } catch { /* 기본 처리 */ }
      }
      return false;
    };


    // ───────────────────────── 찾기 / 바꾸기 ─────────────────────────
    const openFind = (mode: "find" | "replace") => {
      const r = selectionRange();
      const d = docRef.current;
      if (r && d && r.start.para === r.end.para) {
        try {
          const t = r.start.cell
            ? d.getTextInCellByPath(r.start.sec, r.start.cell.parentPara, cellPathJsonFor(r.start), r.start.off, r.end.off - r.start.off)
            : d.getTextRange(r.start.sec, r.start.para, r.start.off, r.end.off - r.start.off);
          if (t && !t.includes("\n")) setFindQ(t);
        } catch { /* 선택 글을 못 읽으면 기존 검색어 유지 */ }
      }
      setFindMsg(""); setFindOpen(mode);
      setTimeout(() => { findInputRef.current?.focus(); findInputRef.current?.select(); }, 0);
    };
    const closeFind = () => { setFindOpen(null); inputRef.current?.focus(); };

    /** from 위치부터 앞/뒤로 찾아 찾은 글을 선택한다. */
    const findFrom = (from: Pos, forward: boolean): boolean => {
      const d = docRef.current;
      if (!d || !findQ) return false;
      try {
        const para = from.cell ? from.cell.parentPara : from.para;
        const off = from.cell ? (forward ? 0 : 0) : from.off;
        const r = JSON.parse(d.searchText(findQ, from.sec, para, off, forward, findCase, true)) as {
          found: boolean; wrapped?: boolean; sec: number; para: number; charOffset: number; length: number;
          cellContext?: { parentPara: number; ctrlIdx: number; cellIdx: number; cellPara: number };
        };
        if (!r.found) { setFindMsg("찾을 수 없습니다"); return false; }
        const cc = r.cellContext;
        const start: Pos = cc
          ? { sec: r.sec, para: cc.cellPara, off: r.charOffset, cell: { parentPara: cc.parentPara, path: [{ controlIndex: cc.ctrlIdx, cellIndex: cc.cellIdx, cellParaIndex: cc.cellPara }] } }
          : { sec: r.sec, para: r.para, off: r.charOffset };
        clearPending(); setCellSel(null); setPicSel(null);
        scrollNextRef.current = true;
        setAnchor(start); setPos({ ...start, off: start.off + r.length });
        setFindMsg(r.wrapped ? "처음부터 다시 찾았습니다" : "");
        return true;
      } catch (err) { setFindMsg(`찾기 실패: ${errorText(err)}`); return false; }
    };
    const runFind = (forward: boolean) => {
      if (!pos || !findQ) return;
      const r = selectionRange();
      findFrom(r ? (forward ? r.end : r.start) : pos, forward);
    };
    /** 선택된 글이 찾는 글과 같으면 바꾸고 다음으로 */
    const replaceCurrent = () => {
      const d = docRef.current;
      const r = selectionRange();
      if (!d || !r || !findQ) { runFind(true); return; }
      try {
        const got = r.start.cell
          ? d.getTextInCellByPath(r.start.sec, r.start.cell.parentPara, cellPathJsonFor(r.start), r.start.off, r.end.off - r.start.off)
          : d.getTextRange(r.start.sec, r.start.para, r.start.off, r.end.off - r.start.off);
        const same = findCase ? got === findQ : got.toLowerCase() === findQ.toLowerCase();
        if (!same || r.start.para !== r.end.para) { runFind(true); return; }
      } catch { runFind(true); return; }
      let after: Pos = r.start;
      edit((doc) => {
        deleteAt(doc, r.start, r.start.off, r.end.off - r.start.off);
        if (replQ) {
          if (r.start.cell) doc.insertTextInCellByPath(r.start.sec, r.start.cell.parentPara, cellPathJsonFor(r.start), r.start.off, replQ);
          else doc.insertText(r.start.sec, r.start.para, r.start.off, replQ);
        }
        after = { ...r.start, off: r.start.off + [...replQ].length };
        setAnchor(null); setPos(after);
      });
      findFrom(after, true);
    };
    const replaceEverything = () => {
      if (!findQ) return;
      let count = 0;
      edit((doc) => {
        const r = JSON.parse(doc.replaceAll(findQ, replQ, findCase)) as { count?: number };
        count = r.count ?? 0;
        setAnchor(null); setSelRects([]);
        for (let sec = 0; sec < doc.getSectionCount(); sec++) reflowControlParagraphs(doc, sec);
      });
      setFindMsg(`${count}건을 바꿨습니다`);
    };

    const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if (e.nativeEvent.isComposing || !pos) return;
      if (!/^(Shift|Control|Alt|Meta|CapsLock)$/.test(e.key)) scrollNextRef.current = true;
      // 겹글쇠 Ctrl+N,I (그림) / Ctrl+N,T (표)
      if (!/^(Shift|Control|Alt|Meta)$/.test(e.key)) {
        const armed = chordRef.current && Date.now() - chordRef.current < 1500;
        chordRef.current = 0;
        if (armed && e.code === "KeyI") { e.preventDefault(); void insertImage(); return; }
        if (armed && e.code === "KeyT") { e.preventDefault(); insertTable(); return; }
        if ((e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && e.code === "KeyN") { e.preventDefault(); chordRef.current = Date.now(); return; }
      }
      if (picSelRef.current && handlePicKey(e)) return;
      if (handleTableKey(e)) return;
      const mod = e.ctrlKey || e.metaKey;
      if (mod && !e.altKey && e.code === "KeyF" && !e.shiftKey) { e.preventDefault(); openFind("find"); return; }
      if (mod && !e.altKey && e.code === "KeyH" && !e.shiftKey) { e.preventDefault(); openFind("replace"); return; }
      if (e.key === "F3") { e.preventDefault(); if (findQ) runFind(!e.shiftKey); else openFind("find"); return; }
      if (mod && e.shiftKey && !e.altKey) { // 한글 맞춤 단축키: Ctrl+Shift+L/C/R/M
        const al = ({ KeyL: "left", KeyC: "center", KeyR: "right", KeyM: "justify" } as Record<string, string>)[e.code];
        if (al) { e.preventDefault(); paraFmt({ alignment: al }); return; }
      }
      if (mod && e.key.toLowerCase() === "c") { e.preventDefault(); void copySelectionToClipboard(); return; }
      if (mod && e.key.toLowerCase() === "x") { e.preventDefault(); void cutSelectionToClipboard(); return; }
      if (mod && e.key.toLowerCase() === "v") { e.preventDefault(); void pasteFromClipboard(); return; }
      if (mod && e.key.toLowerCase() === "a") { e.preventDefault(); selectAll(); return; }
      if (mod && e.key.toLowerCase() === "z") { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
      if (mod && e.key.toLowerCase() === "y") { e.preventDefault(); redo(); return; }
      if (mod && e.key.toLowerCase() === "b") { e.preventDefault(); charFmt({ bold: true }); return; }
      if (mod && e.key.toLowerCase() === "i") { e.preventDefault(); charFmt({ italic: true }); return; }
      if (mod && e.key.toLowerCase() === "u") { e.preventDefault(); charFmt({ underline: true }); return; }
      const range = selectionRange();
      switch (e.key) {
        case "Enter": {
          e.preventDefault();
          preferredXRef.current = -1;
          if (mod && !e.shiftKey && !pos.cell) { // Ctrl+Enter: 쪽 나누기 (표 안에서는 줄 추가가 먼저 처리됨)
            edit((d) => {
              let p = pos;
              if (range) { deleteRange(d, range); p = range.start; }
              const r = JSON.parse(d.insertPageBreak(p.sec, p.para, p.off)) as { paraIdx?: number; charOffset?: number };
              reflowControlParagraphs(d, p.sec);
              setPos({ sec: p.sec, para: r.paraIdx ?? p.para + 1, off: r.charOffset ?? 0 }); setAnchor(null);
            });
            break;
          }
          edit((d) => {
            let p = pos;
            if (range) { deleteRange(d, range); p = range.start; }
            if (p.cell) d.splitParagraphInCellByPath(p.sec, p.cell.parentPara, cellPathJsonFor(p), p.off);
            else d.splitParagraph(p.sec, p.para, p.off, undefined);
            setPos({ ...p, para: p.para + 1, off: 0 }); setAnchor(null);
          });
          break;
        }
        case "Backspace": {
          e.preventDefault();
          preferredXRef.current = -1;
          if (range) {
            edit((d) => { deleteRange(d, range); setPos(range.start); setAnchor(null); });
          } else if (pos.off > 0) {
            edit((d) => {
              if (!pos.cell) {
                try { d.replaceBodyTextLocal(pos.sec, pos.para, pos.off - 1, 1, ""); }
                catch { deleteAt(d, pos, pos.off - 1, 1); }
              } else if (pos.cell.path.length === 1) {
                const c = pos.cell.path[0];
                try {
                  d.deleteTextInCellDeferredPagination(
                    pos.sec, pos.cell.parentPara, c.controlIndex, c.cellIndex, pos.para, pos.off - 1, 1);
                } catch { deleteAt(d, pos, pos.off - 1, 1); }
              } else deleteAt(d, pos, pos.off - 1, 1);
              setPos({ ...pos, off: pos.off - 1 });
            }, "typing");
          } else if (pos.para > 0) {
            edit((d) => {
              const raw = pos.cell
                ? d.mergeParagraphInCellByPath(pos.sec, pos.cell.parentPara, cellPathJsonFor(pos))
                : d.mergeParagraph(pos.sec, pos.para);
              const r = JSON.parse(raw);
              setPos({ ...pos, para: pos.para - 1, off: r.charOffset ?? r.char_offset ?? 0 });
            });
          }
          break;
        }
        case "Delete": {
          e.preventDefault();
          preferredXRef.current = -1;
          if (range) edit((d) => { deleteRange(d, range); setPos(range.start); setAnchor(null); });
          else if (pos.off < paraLength(docRef.current!, pos)) edit((d) => {
            if (!pos.cell) {
              try { d.replaceBodyTextLocal(pos.sec, pos.para, pos.off, 1, ""); }
              catch { deleteAt(d, pos, pos.off, 1); }
            } else if (pos.cell.path.length === 1) {
              const c = pos.cell.path[0];
              try {
                d.deleteTextInCellDeferredPagination(
                  pos.sec, pos.cell.parentPara, c.controlIndex, c.cellIndex, pos.para, pos.off, 1);
              } catch { deleteAt(d, pos, pos.off, 1); }
            } else deleteAt(d, pos, pos.off, 1);
          }, "typing");
          break;
        }
        case "Home":
          e.preventDefault(); clearPending();
          if (mod) moveDocEdge(false, e.shiftKey); else moveLineEdge(false, e.shiftKey);
          break;
        case "End":
          e.preventDefault(); clearPending();
          if (mod) moveDocEdge(true, e.shiftKey); else moveLineEdge(true, e.shiftKey);
          break;
        case "PageUp":
          e.preventDefault(); clearPending();
          movePage(-1, e.shiftKey);
          break;
        case "PageDown":
          e.preventDefault(); clearPending();
          movePage(1, e.shiftKey);
          break;
        case "ArrowLeft":
          e.preventDefault(); clearPending();
          if (mod) moveWord(-1, e.shiftKey); else moveHorizontal(-1, e.shiftKey);
          break;
        case "ArrowRight":
          e.preventDefault(); clearPending();
          if (mod) moveWord(1, e.shiftKey); else moveHorizontal(1, e.shiftKey);
          break;
        case "Insert":
          e.preventDefault();
          if (e.shiftKey) void pasteFromClipboard();
          else if (mod) void copySelectionToClipboard();
          else setOverwrite((v) => !v);
          break;
        case "ArrowUp":
          e.preventDefault(); clearPending();
          moveCaretVertical(-1, e.shiftKey);
          break;
        case "ArrowDown":
          e.preventDefault(); clearPending();
          moveCaretVertical(1, e.shiftKey);
          break;
        case "Escape":
          e.preventDefault(); clearPending(); setAnchor(null); setSelRects([]); setCellSel(null);
          break;
      }
    };

    // 글자 서식
    const charFmt = (props: Record<string, unknown>) => {
      const d = docRef.current;
      if (!d) return;
      if (cellSelRef.current) { applyFmtToBlock(props); return; }
      const range = selectionRange();
      const base = range ? range.start : pos;
      if (!base) return;
      const next: Record<string, unknown> = { ...props };
      const current = { ...charPropsAt(d, base), ...(range ? {} : (pendingProps() ?? {})) };
      for (const k of TOGGLE_KEYS) {
        if (k in next) next[k] = !current[k];
      }
      const fontName = typeof next.fontFamily === "string" ? next.fontFamily : null;
      if (fontName) {
        const id = d.findOrCreateFontId(fontName);
        delete next.fontFamily;
        if (id >= 0) next.fontId = id;
      }
      if (range) {
        // 블록 선택이 있을 때만 기존 글자에 적용한다.
        edit((doc) => applyCharToRange(doc, range, next));
      }
      setPendingHint(!range);
      // 선택이 없을 때는 다음 입력에 적용될 서식만 예약한다.
      if (typeof next.fontSize === "number") setSize(String(Math.round(next.fontSize / 100)));
      if (typeof next.textColor === "string") setColor(next.textColor);
      if (fontName) setFont(fontName);
      if (typeof next.bold === "boolean") setBold(next.bold);
      if (typeof next.italic === "boolean") setItalic(next.italic);
      if (typeof next.underline === "boolean") setUnderline(next.underline);
      if (typeof next.strikethrough === "boolean") setStrike(next.strikethrough);
      pendingRef.current = { ...(pendingProps() ?? {}), ...next };
      pendingPosRef.current = range ? null : pos;
    };

    // 문단 서식
    const paraFmt = (props: Record<string, unknown>) => {
      if (!pos) return;
      const range = selectionRange();
      edit((d) => {
        if (!range) { applyParaToPos(d, pos, JSON.stringify(props)); return; }
        for (let par = range.start.para; par <= range.end.para; par++) {
          applyParaToPos(d, { ...range.start, para: par }, JSON.stringify(props));
        }
      });
    };

    // 표 / 그림 / 도형 삽입
    const insertTable = () => {
      if (!pos) return;
      if (pos.cell) { onError("표 안에는 표를 넣을 수 없습니다. 표 밖 본문에 커서를 두고 삽입하세요"); return; }
      const rows = Number(prompt("행 개수", "3")); const cols = Number(prompt("열 개수", "3"));
      if (!rows || !cols) return;
      edit((d) => {
        const r = createInlineTable(d, pos.sec, pos.para, pos.off, rows, cols);
        reflowControlParagraphs(d, pos.sec);
        // 한글처럼 만든 표의 첫 칸에 커서를 둔다.
        clearPending(); setCellSel(null); setAnchor(null); setSelRects([]);
        scrollNextRef.current = true;
        setPos(cellStartPos({ sec: pos.sec, parent: r.paraIdx, ctrl: r.controlIdx }, 0));
      });
    };
    const insertImage = async () => {
      if (!pos) return;
      if (pos.cell) { onError("표 안에는 아직 그림을 넣을 수 없습니다. 표 밖 본문에 커서를 두고 삽입하세요"); return; }
      try {
        const path = await open({
          multiple: false,
          filters: [{ name: "그림 (png, jpg, gif, bmp)", extensions: ["png", "jpg", "jpeg", "gif", "bmp", "webp"] }],
        });
        if (typeof path !== "string") return;
        const bytes = await readBytes(path);
        const ext = (path.split(".").pop() || "png").toLowerCase();
        const mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "bmp" ? "image/bmp" : `image/${ext}`;
        const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
        const bmp = await createImageBitmap(new Blob([buf], { type: mime }));
        // 한글처럼 원래 크기로 넣되 본문 너비(약 148mm)를 넘으면 줄인다.
        const scale = Math.min(1, 560 / bmp.width, 760 / bmp.height);
        const wPx = Math.max(8, Math.round(bmp.width * scale));
        const hPx = Math.max(8, Math.round(bmp.height * scale));
        const p = pos;
        edit((d) => {
          const r = JSON.parse(d.insertPicture(
            p.sec, p.cell ? p.cell.parentPara : p.para, p.off,
            p.cell ? cellPathJsonFor(p) : "[]",
            bytes, wPx * HU_PER_PX, hPx * HU_PER_PX,
            bmp.width, bmp.height, ext === "jpeg" ? "jpg" : ext, "그림",
          )) as { paraIdx?: number; controlIdx?: number };
          if (!p.cell && typeof r.paraIdx === "number" && typeof r.controlIdx === "number") {
            // 한글 기본값: 글자처럼 취급 → 커서 위치에 놓이고 뒤따르는 글/표가 아래로 밀린다.
            d.setPictureProperties(p.sec, r.paraIdx, r.controlIdx, JSON.stringify({ treatAsChar: true }));
            reflowControlParagraphs(d, p.sec);
            const ref: PicRef = { sec: p.sec, para: r.paraIdx, ctrl: r.controlIdx };
            clearPending(); setCellSel(null); setAnchor(null); setSelRects([]);
            scrollNextRef.current = true;
            setPicSel(ref);
            setPos({ sec: p.sec, para: r.paraIdx, off: ctrlOffsetOf(d, ref) });
          }
        });
        bmp.close();
        inputRef.current?.focus();
      } catch (e) { onError(`그림 삽입 실패: ${errorText(e)}`); }
    };
    const insertShape = (shapeType: "rectangle" | "ellipse" | "textbox" | "line") => {
      if (!pos) return;
      edit((d) => {
        d.createShapeControl(JSON.stringify({
          sectionIdx: pos.sec, paraIdx: pos.para, charOffset: pos.off,
          shapeType, width: 14173, height: shapeType === "line" ? 0 : 7087, // HWPUNIT (5cm × 2.5cm)
        }));
        if (!pos.cell) reflowControlParagraphs(d, pos.sec);
      });
    };

    // 페이지 여백 / 쪽 번호
    const hasPageNumber = (info: Record<string, unknown>) => {
      if (!info.exists) return false;
      const text = String(info.text ?? "");
      return text.includes("\u0015") || (text.length === 1 && text.trim() === "");
    };

    const readPageSettings = (d: HwpDocument) => {
      try {
        const pd = JSON.parse(d.getPageDef(0)) as Record<string, number>;
        setPageMargins({
          top: huToMm(pd.marginTop ?? 0),
          bottom: huToMm(pd.marginBottom ?? 0),
          left: huToMm(pd.marginLeft ?? 0),
          right: huToMm(pd.marginRight ?? 0),
          header: huToMm(pd.marginHeader ?? 0),
          footer: huToMm(pd.marginFooter ?? 0),
        });
      } catch { /* 페이지 정의를 읽지 못하면 기본값 유지 */ }
      try {
        const footer = JSON.parse(d.getHeaderFooter(0, false, 0)) as Record<string, unknown>;
        const header = JSON.parse(d.getHeaderFooter(0, true, 0)) as Record<string, unknown>;
        setPageNum(hasPageNumber(footer) ? "footer" : hasPageNumber(header) ? "header" : "none");
      } catch { setPageNum("none"); }
    };

    const removePageNumber = (d: HwpDocument, sec: number, isHeader: boolean) => {
      try {
        const info = JSON.parse(d.getHeaderFooter(sec, isHeader, 0)) as Record<string, unknown>;
        if (!info.exists) return;
        const text = String(info.text ?? "");
        const idx = text.indexOf("\u0015");
        if (idx >= 0) d.deleteTextInHeaderFooter(sec, isHeader, 0, 0, idx, 1);
        else if (text.length === 1 && text.trim() === "") d.deleteTextInHeaderFooter(sec, isHeader, 0, 0, 0, 1);
      } catch { /* 없으면 무시 */ }
    };

    const ensurePageNumber = (d: HwpDocument, sec: number, isHeader: boolean) => {
      const info = JSON.parse(d.getHeaderFooter(sec, isHeader, 0)) as Record<string, unknown>;
      if (!info.exists) d.createHeaderFooter(sec, isHeader, 0);
      if (hasPageNumber(info)) {
        d.applyParaFormatInHf(sec, isHeader, 0, 0, JSON.stringify({ alignment: "center" }));
        return;
      }
      const pinfo = JSON.parse(d.getHeaderFooterParaInfo(sec, isHeader, 0, 0)) as Record<string, unknown>;
      const text = String(pinfo.text ?? "");
      const offset = Number(pinfo.charCount ?? text.length ?? 0);
      d.insertFieldInHf(sec, isHeader, 0, 0, offset, 1); // 1 = 쪽 번호 필드
      d.applyParaFormatInHf(sec, isHeader, 0, 0, JSON.stringify({ alignment: "center" }));
    };

    const applyPageSettings = () => {
      const d = docRef.current;
      if (!d) return;
      try {
        const pageDef = JSON.stringify({
          width: mmToHu("210"), height: mmToHu("297"),
          marginTop: mmToHu(pageMargins.top), marginBottom: mmToHu(pageMargins.bottom),
          marginLeft: mmToHu(pageMargins.left), marginRight: mmToHu(pageMargins.right),
          marginHeader: mmToHu(pageMargins.header), marginFooter: mmToHu(pageMargins.footer),
        });
        edit((doc) => {
          const sections = doc.getSectionCount();
          for (let sec = 0; sec < sections; sec++) {
            doc.setPageDef(sec, pageDef);
            if (pageNum === "none") {
              removePageNumber(doc, sec, false);
              removePageNumber(doc, sec, true);
            } else {
              const isHeader = pageNum === "header";
              ensurePageNumber(doc, sec, isHeader);
              removePageNumber(doc, sec, !isHeader);
            }
          }
        });
        setPageOpen(false);
        onError("페이지 여백과 쪽 번호가 적용되었습니다");
      } catch (e) {
        onError(`페이지 설정 실패: ${errorText(e)}`);
      }
    };

    // 문서가 바뀌면 현재 여백/쪽 번호 상태를 툴바에 반영한다.
    useEffect(() => {
      const d = docRef.current;
      if (d && pageCount > 0) readPageSettings(d);
    }, [pageCount, bytes, reloadKey]);

    useImperativeHandle(ref, () => ({
      async exportAs(fmt: ExportFormat) {
        const d = docRef.current; if (!d) throw new Error("문서 없음");
        if (fmt === "hwp") return d.exportHwp();
        if (fmt === "hwpx") return d.exportHwpx();
        if (fmt === "pdf") return await hwpToPdf(d);
        throw new Error(`HWP 문서는 ${fmt}로 저장할 수 없습니다`);
      },
    }));

    // 문서에 이미 지정된 글꼴이 목록에 없어도 현재 값이 비어 보이지 않게 한다.
    const fontOptions: [string, string][] = font && !HWP_FONTS.some(([v]) => v === font)
      ? [[font, font], ...HWP_FONTS]
      : HWP_FONTS;

    return (
      <div className="flex flex-col h-full">
        <div className="flex flex-wrap items-center gap-0.5 px-3 py-1.5 bg-white border-b border-gray-200">
          <TBtn title="실행 취소 (Ctrl+Z)" onClick={undo}><Undo2 size={16} /></TBtn>
          <TBtn title="다시 실행 (Ctrl+Y)" onClick={redo}><Redo2 size={16} /></TBtn>
          <Sep />
          <TSelect title="글꼴" width="w-28" value={font} options={fontOptions}
            onChange={(v) => { setFont(v); charFmt({ fontFamily: v }); }} />
          <TSelect title="크기" width="w-16" value={size} options={SIZES.map((x) => [x, x] as [string, string])}
            onChange={(v) => { setSize(v); charFmt({ fontSize: Number(v) * 100 }); }} />
          <input type="color" title="글자 색" value={color}
            onChange={(e) => { setColor(e.target.value); charFmt({ textColor: e.target.value }); }}
            className="h-8 w-8 p-1 rounded-md border border-gray-200 bg-white cursor-pointer" />
          {pendingHint && (
            <span className="ml-1 whitespace-nowrap text-[11px] text-amber-600">다음 입력에 적용</span>
          )}
          <TBtn title="굵게 (Ctrl+B)" active={bold} onClick={() => charFmt({ bold: true })}><Bold size={16} /></TBtn>
          <TBtn title="기울임 (Ctrl+I)" active={italic} onClick={() => charFmt({ italic: true })}><Italic size={16} /></TBtn>
          <TBtn title="밑줄 (Ctrl+U)" active={underline} onClick={() => charFmt({ underline: true })}><Underline size={16} /></TBtn>
          <TBtn title="취소선" active={strike} onClick={() => charFmt({ strikethrough: true })}><Strikethrough size={16} /></TBtn>
          <Sep />
          <TBtn title="왼쪽 정렬" onClick={() => paraFmt({ alignment: "left" })}><AlignLeft size={16} /></TBtn>
          <TBtn title="가운데 정렬" onClick={() => paraFmt({ alignment: "center" })}><AlignCenter size={16} /></TBtn>
          <TBtn title="오른쪽 정렬" onClick={() => paraFmt({ alignment: "right" })}><AlignRight size={16} /></TBtn>
          <TBtn title="양쪽 정렬" onClick={() => paraFmt({ alignment: "justify" })}><AlignJustify size={16} /></TBtn>
          <Sep />
          <button title="삽입 / 수정(덮어쓰기) 전환 (Insert)" onMouseDown={(e) => e.preventDefault()} onClick={() => { setOverwrite((v) => !v); inputRef.current?.focus(); }}
            className={`h-6 rounded px-1.5 text-[11px] ${overwrite ? "bg-amber-100 font-medium text-amber-700" : "text-gray-400 hover:bg-gray-100"}`}>{overwrite ? "수정" : "삽입"}</button>
          <div className="relative" ref={symbolRef}>
            <TBtn title="기호 삽입" active={symbolsOpen} onClick={() => setSymbolsOpen((v) => !v)}><Smile size={16} /></TBtn>
            {symbolsOpen && (
              <div className="absolute left-0 top-9 z-50 w-[360px] rounded-lg border border-gray-200 bg-white p-2 shadow-xl">
                <div className="mb-1 flex items-center justify-between px-1">
                  <span className="text-xs font-medium text-gray-600">기호 · 이모지</span>
                  <button onMouseDown={(e) => e.preventDefault()} onClick={() => setSymbolsOpen(false)}
                    className="text-xs text-gray-400 hover:text-gray-600">닫기</button>
                </div>
                <div className="max-h-64 overflow-y-auto pr-1">
                  {SYMBOL_GROUPS.map((group) => (
                    <div key={group.name} className="mb-2">
                      <div className="mb-1 text-[11px] text-gray-400">{group.name}</div>
                      <div className="grid grid-cols-10 gap-1">
                        {group.items.map((s, idx) => (
                          <button key={`${group.name}-${idx}`} title={s}
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => { insert(s); inputRef.current?.focus(); }}
                            className="flex h-7 w-7 items-center justify-center rounded text-base leading-none hover:bg-blue-50">
                            {s}
                          </button>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
          <TBtn title="표 삽입 (Ctrl+N, T)" onClick={insertTable}><TableIcon size={16} /></TBtn>
          <TBtn title="그림 삽입 (Ctrl+N, I)" onClick={insertImage}><ImageIcon size={16} /></TBtn>
          <TBtn title="사각형 삽입" onClick={() => insertShape("rectangle")}><Square size={16} /></TBtn>
          <TBtn title="타원 삽입" onClick={() => insertShape("ellipse")}><Circle size={16} /></TBtn>
          <TBtn title="글상자 삽입" onClick={() => insertShape("textbox")}><Type size={16} /></TBtn>
          <TBtn title="선 삽입" onClick={() => insertShape("line")}><Minus size={16} /></TBtn>
          <Sep />
          <TBtn title="찾기 / 바꾸기 (Ctrl+F, Ctrl+H)" onClick={() => openFind("find")}><SearchIcon size={16} /></TBtn>
          <TBtn title="축소" onClick={() => setZoom((z) => Math.max(0.5, +(z - 0.1).toFixed(1)))}><ZoomOut size={16} /></TBtn>
          <span className="text-xs text-gray-500 w-10 text-center">{Math.round(zoom * 100)}%</span>
          <TBtn title="확대" onClick={() => setZoom((z) => Math.min(3, +(z + 0.1).toFixed(1)))}><ZoomIn size={16} /></TBtn>
          <Sep />
          <div className="relative" ref={pagePanelRef}>
            <TBtn title="페이지 여백 / 쪽 번호" active={pageOpen} onClick={() => setPageOpen((v) => !v)}><Settings2 size={16} /></TBtn>
            {pageOpen && (
              <div className="absolute right-0 top-9 z-50 w-[300px] rounded-lg border border-gray-200 bg-white p-3 shadow-xl">
                <div className="mb-2 flex items-center justify-between">
                  <span className="text-xs font-medium text-gray-700">페이지 설정 (A4)</span>
                  <button onMouseDown={(e) => e.preventDefault()} onClick={() => setPageOpen(false)}
                    className="text-xs text-gray-400 hover:text-gray-600">닫기</button>
                </div>
                <div className="grid grid-cols-2 gap-2">
                  {([
                    ["top", "위 여백(mm)"], ["bottom", "아래 여백(mm)"],
                    ["left", "왼쪽 여백(mm)"], ["right", "오른쪽 여백(mm)"],
                    ["header", "머리말(mm)"], ["footer", "꼬리말(mm)"],
                  ] as [keyof PageMargins, string][]).map(([key, label]) => (
                    <label key={key} className="flex items-center justify-between gap-2 text-[11px] text-gray-600">
                      <span>{label}</span>
                      <input type="number" min="0" step="1" title={label} value={pageMargins[key]}
                        onChange={(e) => setPageMargins((m) => ({ ...m, [key]: e.target.value }))}
                        className="h-7 w-16 rounded-md border border-gray-200 px-2 text-right text-xs" />
                    </label>
                  ))}
                </div>
                <label className="mt-2 flex items-center justify-between gap-2 text-[11px] text-gray-600">
                  <span>쪽 번호</span>
                  <select title="쪽 번호" value={pageNum}
                    onChange={(e) => setPageNum(e.target.value as "none" | "footer" | "header")}
                    className="h-7 rounded-md border border-gray-200 px-2 text-xs">
                    <option value="none">없음</option>
                    <option value="footer">아래 가운데</option>
                    <option value="header">위 가운데</option>
                  </select>
                </label>
                <div className="mt-3 flex items-center justify-between">
                  <button onMouseDown={(e) => e.preventDefault()} onClick={() => setPageMargins(A4_DEFAULT_MARGINS)}
                    className="h-7 rounded-md border border-gray-200 px-2 text-xs text-gray-600 hover:bg-gray-50">A4 기본값</button>
                  <button onMouseDown={(e) => e.preventDefault()} onClick={applyPageSettings}
                    className="h-7 rounded-md bg-blue-600 px-3 text-xs text-white hover:bg-blue-700">적용</button>
                </div>
              </div>
            )}
          </div>
        </div>

        {!loading && picSel && !picDlg && (
          <div className="fixed right-3 top-[132px] z-40 w-[132px] rounded-lg border border-gray-200 bg-white/95 p-1.5 text-[11px] shadow-lg"
            onMouseDown={(e) => e.preventDefault()}>
            <div className="mb-1 px-1 font-medium text-gray-700">그림 선택됨</div>
            {([
              ["그림 속성 (P)", () => openPicDialog(picSel)],
              ["복사 (Ctrl+C)", () => void copyPicture(picSel, false)],
              ["오려 두기 (Ctrl+X)", () => void copyPicture(picSel, true)],
              ["지우기 (Delete)", () => deletePicture(picSel)],
            ] as [string, () => void][]).map(([label, fn]) => (
              <button key={label} onMouseDown={(e) => e.preventDefault()} onClick={() => { fn(); inputRef.current?.focus(); }}
                className="block w-full rounded px-1.5 py-[3px] text-left text-gray-700 hover:bg-blue-50">{label}</button>
            ))}
            <div className="mt-1 border-t border-gray-100 px-1 pt-1 leading-tight text-gray-400">손잡이를 끌어 크기 조절<br />Shift+방향키: 크기 1mm</div>
          </div>
        )}

        {picDlg && (
          <div className="fixed right-3 top-[132px] z-50 w-[300px] rounded-lg border border-gray-200 bg-white p-3 text-xs shadow-xl">
            <div className="mb-2 flex items-center justify-between">
              <span className="font-medium text-gray-700">그림 속성</span>
              <button className="text-gray-400 hover:text-gray-600" onClick={() => { setPicDlg(null); inputRef.current?.focus(); }}>닫기</button>
            </div>
            {([["너비 (mm)", "wMm"], ["높이 (mm)", "hMm"]] as [string, "wMm" | "hMm"][]).map(([label, key]) => (
              <label key={key} className="mb-1.5 flex items-center justify-between gap-2 text-gray-600">
                <span>{label}</span>
                <input type="number" min="1" step="0.1" title={label} value={picDlg[key]}
                  onChange={(e) => {
                    const v = e.target.value; const n = Number(v);
                    setPicDlg((dlg) => {
                      if (!dlg) return dlg;
                      if (!dlg.keep || !(n > 0)) return { ...dlg, [key]: v };
                      return key === "wMm"
                        ? { ...dlg, wMm: v, hMm: String(Math.round(n * dlg.ratio * 10) / 10) }
                        : { ...dlg, hMm: v, wMm: String(Math.round((n / dlg.ratio) * 10) / 10) };
                    });
                  }}
                  onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); applyPicDlg(); setTimeout(() => inputRef.current?.focus(), 0); } if (e.key === "Escape") { setPicDlg(null); inputRef.current?.focus(); } }}
                  className="h-7 w-24 rounded-md border border-gray-200 px-2 text-right" />
              </label>
            ))}
            <label className="mb-1 flex items-center gap-2 text-gray-600">
              <input type="checkbox" checked={picDlg.keep} onChange={(e) => setPicDlg({ ...picDlg, keep: e.target.checked })} /> 가로 세로 비율 유지
            </label>
            <PlacementFields kind="picture" original={picDlg.plPrev} value={picDlg.pl} onChange={(pl) => setPicDlg((dlg) => (dlg ? { ...dlg, pl } : dlg))} />
            <div className="flex justify-end gap-2">
              <button className="h-7 rounded-md border border-gray-200 px-3 text-gray-600 hover:bg-gray-50" onClick={() => { setPicDlg(null); inputRef.current?.focus(); }}>취소</button>
              <button className="h-7 rounded-md bg-blue-600 px-3 text-white hover:bg-blue-700" onClick={() => { applyPicDlg(); inputRef.current?.focus(); }}>적용</button>
            </div>
          </div>
        )}

        {tblDlg && (
          <div className="fixed right-3 top-[132px] z-50 w-[300px] rounded-lg border border-gray-200 bg-white p-3 text-xs shadow-xl"
            onKeyDown={(e) => { if (e.key === "Escape") { e.preventDefault(); setTblDlg(null); inputRef.current?.focus(); } }}>
            <div className="mb-2 flex items-center justify-between">
              <span className="font-medium text-gray-700">표 속성</span>
              <button className="text-gray-400 hover:text-gray-600" onClick={() => { setTblDlg(null); inputRef.current?.focus(); }}>닫기</button>
            </div>
            <div className="mb-2 text-[11px] text-gray-500">크기 {tblDlg.wMm} × {tblDlg.hMm} mm</div>
            <PlacementFields kind="table" original={tblDlg.plPrev} value={tblDlg.pl} onChange={(pl) => setTblDlg((dlg) => (dlg ? { ...dlg, pl } : dlg))} />
            <div className="mb-2 text-[10px] leading-tight text-gray-400">떠 있는 표는 테두리를 끌거나, 표 선택 상태에서 방향키(1mm)로 옮깁니다. Shift+Home 글 앞으로, Shift+End 글 뒤로.</div>
            <div className="flex justify-end gap-2">
              <button className="h-7 rounded-md border border-gray-200 px-3 text-gray-600 hover:bg-gray-50" onClick={() => { setTblDlg(null); inputRef.current?.focus(); }}>취소</button>
              <button className="h-7 rounded-md bg-blue-600 px-3 text-white hover:bg-blue-700" onClick={() => { applyTableDlg(); inputRef.current?.focus(); }}>적용</button>
            </div>
          </div>
        )}

        {findOpen && (
          <div onKeyDown={(e) => { if (e.key === "Escape") { e.preventDefault(); closeFind(); } }} className="fixed left-1/2 top-[84px] z-50 w-[420px] -translate-x-1/2 rounded-lg border border-gray-200 bg-white p-3 text-xs shadow-xl">
            <div className="mb-2 flex items-center justify-between">
              <span className="font-medium text-gray-700">{findOpen === "replace" ? "찾아 바꾸기" : "찾기"}</span>
              <button className="text-gray-400 hover:text-gray-600" onClick={closeFind}>닫기 (Esc)</button>
            </div>
            <div className="mb-1.5 flex items-center gap-2">
              <span className="w-12 text-gray-500">찾을 말</span>
              <input ref={findInputRef} value={findQ} onChange={(e) => { setFindQ(e.target.value); setFindMsg(""); }}
                onKeyDown={(e) => {
                  if (e.nativeEvent.isComposing) return;
                  if (e.key === "Enter") { e.preventDefault(); runFind(!e.shiftKey); }
                  if (e.key === "Escape") { e.preventDefault(); closeFind(); }
                }}
                title="찾을 말" className="h-7 flex-1 rounded-md border border-gray-200 px-2" />
            </div>
            {findOpen === "replace" && (
              <div className="mb-1.5 flex items-center gap-2">
                <span className="w-12 text-gray-500">바꿀 말</span>
                <input value={replQ} onChange={(e) => setReplQ(e.target.value)} title="바꿀 말"
                  onKeyDown={(e) => { if (e.nativeEvent.isComposing) return; if (e.key === "Enter") { e.preventDefault(); replaceCurrent(); } if (e.key === "Escape") { e.preventDefault(); closeFind(); } }}
                  className="h-7 flex-1 rounded-md border border-gray-200 px-2" />
              </div>
            )}
            <div className="flex items-center gap-1.5">
              <label className="mr-1 flex items-center gap-1 text-gray-600"><input type="checkbox" checked={findCase} onChange={(e) => setFindCase(e.target.checked)} /> 대소문자 구분</label>
              <button className="h-7 rounded-md border border-gray-200 px-2 hover:bg-gray-50" onClick={() => runFind(false)}>이전</button>
              <button className="h-7 rounded-md bg-blue-600 px-2 text-white hover:bg-blue-700" onClick={() => runFind(true)}>다음 찾기</button>
              {findOpen === "replace" && (<>
                <button className="h-7 rounded-md border border-gray-200 px-2 hover:bg-gray-50" onClick={replaceCurrent}>바꾸기</button>
                <button className="h-7 rounded-md border border-gray-200 px-2 hover:bg-gray-50" onClick={replaceEverything}>모두 바꾸기</button>
              </>)}
            </div>
            {findMsg && <div className="mt-1.5 text-[11px] text-amber-600">{findMsg}</div>}
          </div>
        )}

        {!loading && !picSel && (cellSel || tableOfPos(pos)) && (
          <div className="fixed right-3 top-[132px] z-40 w-[132px] rounded-lg border border-gray-200 bg-white/95 p-1.5 text-[11px] shadow-lg"
            onMouseDown={(e) => e.preventDefault()}>
            <div className="mb-1 px-1 font-medium text-gray-700">
              표 {cellSel ? (cellSel.mode === "table" ? "· 표 선택됨" : cellSel.mode === "extend" ? "· 셀 블록(확장)" : "· 셀 선택") : "편집"}
            </div>
            {([
              ["표 속성 (P)", () => { const tp = tableOfPos(pos); const t = cellSel?.t ?? tp?.t; if (t) openTableDialog(t); }, "글자처럼 취급 / 본문과의 배치 / 위치"],
              ["표 선택 (Esc)", () => { const tp = tableOfPos(pos); if (tp) selectTableObject(tp.t, pos ?? undefined); }, "표 전체를 개체로 선택"],
              ["셀 블록 (F5)", () => { const tp = tableOfPos(pos); if (tp && pos && !cellSel) { setAnchor(null); setCellSel({ t: tp.t, a: tp.cell, f: tp.cell, mode: "cell", back: pos }); } }, "현재 셀을 블록으로 선택"],
              ["셀 합치기 (M)", () => tableOp("merge"), "블록으로 잡은 셀을 하나로"],
              ["셀 나누기 (S)", () => tableOp("split"), "합쳐진 셀 되돌리기 / 줄·칸으로 나누기"],
              ["위에 줄 추가", () => tableOp("rowAbove"), "현재 줄 위에 줄 추가"],
              ["아래 줄 추가 (Ctrl+Enter)", () => tableOp("rowBelow"), "현재 줄 아래에 줄 추가"],
              ["왼쪽 칸 추가", () => tableOp("colLeft"), "현재 칸 왼쪽에 칸 추가"],
              ["오른쪽 칸 추가", () => tableOp("colRight"), "현재 칸 오른쪽에 칸 추가"],
              ["줄 지우기 (Alt+Delete)", () => tableOp("delRow"), "현재 줄(블록의 줄들) 삭제"],
              ["칸 지우기", () => tableOp("delCol"), "현재 칸(블록의 칸들) 삭제"],
              ["셀 내용 지우기", () => tableOp("clear"), "블록 셀의 글 지우기"],
              ["표 지우기", () => tableOp("delTable"), "표 전체 삭제"],
            ] as [string, () => void, string][]).map(([label, fn, tip]) => (
              <button key={label} title={tip} onMouseDown={(e) => e.preventDefault()} onClick={() => { fn(); inputRef.current?.focus(); }}
                className="block w-full rounded px-1.5 py-[3px] text-left text-gray-700 hover:bg-blue-50">{label}</button>
            ))}
          </div>
        )}

        {loading && (
          <div className="px-4 py-2 text-sm text-gray-500 bg-gray-50 border-b border-gray-200">
            문서를 불러오는 중…
          </div>
        )}

        {loadError && (
          <div className="mx-4 mt-4 flex items-center gap-3 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700">
            <span className="flex-1">{loadError}</span>
            <button onClick={() => setReloadKey((k) => k + 1)}
              className="h-7 px-2.5 rounded-md bg-white border border-red-200 hover:bg-red-100">다시 시도</button>
          </div>
        )}

        <div ref={scrollRef} className="flex-1 overflow-auto py-8 relative">
          <textarea
            ref={inputRef}
            className="fixed opacity-0 w-px h-px -left-10"
            onKeyDown={onKeyDown}
            onCompositionEnd={(e) => commitTextarea(e.currentTarget)}
            onInput={(e) => {
              if ((e.nativeEvent as InputEvent).isComposing) return;
              commitTextarea(e.currentTarget);
            }}
          />
          {Array.from({ length: pageCount }, (_, i) => (
            <div key={i} ref={(el) => { pageRefs.current[i] = el; }} data-page={i}
              className="relative mx-auto mb-6 w-fit shadow-sm border border-gray-200 bg-white">
              <canvas ref={(el) => { canvases.current[i] = el; }} onMouseDown={(e) => onPageMouseDown(i, e)} onMouseMove={(e) => onPageHover(i, e)} onDoubleClick={(e) => onPageDoubleClick(i, e)} className="block cursor-text" />
              {selRects.filter((r) => r.page === i).map((r, j) => (
                <div key={j} className="absolute bg-blue-500/25 pointer-events-none"
                  style={{ left: r.x * zoom, top: r.y * zoom, width: r.w * zoom, height: r.h * zoom }} />
              ))}
              {cellOverlay.cells.filter((r) => r.page === i).map((r, j) => (
                <div key={`cs${j}`} className="absolute bg-blue-500/35 pointer-events-none"
                  style={{ left: r.x * zoom, top: r.y * zoom, width: r.w * zoom, height: r.h * zoom }} />
              ))}
              {cellSel?.mode === "table" && cellOverlay.frames.filter((r) => r.page === i).map((r, j) => (
                <div key={`fr${j}`} className="absolute pointer-events-none border-2 border-blue-600"
                  style={{ left: r.x * zoom - 2, top: r.y * zoom - 2, width: r.w * zoom + 4, height: r.h * zoom + 4 }}>
                  {[[-4, -4], ["50%", -4], ["100%", -4], [-4, "50%"], ["100%", "50%"], [-4, "100%"], ["50%", "100%"], ["100%", "100%"]].map(([lx, ty], k) => (
                    <span key={k} className="absolute h-[7px] w-[7px] border border-blue-600 bg-white"
                      style={{ left: lx as number | string, top: ty as number | string, marginLeft: typeof lx === "string" ? -4 : 0, marginTop: typeof ty === "string" ? -4 : 0 }} />
                  ))}
                </div>
              ))}
              {cellSel && cellSel.mode !== "table" && cellOverlay.focus && cellOverlay.focus.page === i && (
                <div className={`absolute pointer-events-none h-2.5 w-2.5 rounded-full border border-white ${cellSel.mode === "extend" ? "bg-red-500" : "bg-gray-500"}`}
                  style={{
                    left: (cellOverlay.focus.x + cellOverlay.focus.w / 2) * zoom - 5,
                    top: (cellOverlay.focus.y + cellOverlay.focus.h / 2) * zoom - 5,
                  }} />
              )}
              {picBox && picBox.page === i && (
                <div className="absolute pointer-events-none border-2 border-blue-600"
                  style={{ left: picBox.x * zoom - 2, top: picBox.y * zoom - 2, width: picBox.w * zoom + 4, height: picBox.h * zoom + 4 }}>
                  {HANDLES.map((h) => (
                    <span key={h} onMouseDown={(e) => startPicResize(e, h)}
                      className="pointer-events-auto absolute h-[9px] w-[9px] border border-blue-600 bg-white"
                      style={{ left: `${handlePos[h][0] * 100}%`, top: `${handlePos[h][1] * 100}%`, marginLeft: -5, marginTop: -5, cursor: handleCursor[h] }} />
                  ))}
                </div>
              )}
              {tblGhost && tblGhost.page === i && (
                <div className="absolute pointer-events-none border-2 border-dashed border-blue-500 bg-blue-200/20"
                  style={{ left: tblGhost.x * zoom, top: tblGhost.y * zoom, width: tblGhost.w * zoom, height: tblGhost.h * zoom }} />
              )}
              {picGhost && picGhost.page === i && (
                <div className="absolute pointer-events-none border-2 border-dashed border-red-500 bg-red-200/20"
                  style={{ left: picGhost.x * zoom, top: picGhost.y * zoom, width: picGhost.w * zoom, height: picGhost.h * zoom }}>
                  <span className="absolute -top-5 left-0 whitespace-nowrap rounded bg-red-500 px-1 text-[10px] text-white">
                    {picHuToMm(picGhost.w * HU_PER_PX)} × {picHuToMm(picGhost.h * HU_PER_PX)} mm
                  </span>
                </div>
              )}
              {caret && !cellSel && !picSel && caret.page === i && (
                <div className="absolute w-[2px] bg-black hwp-caret pointer-events-none"
                  style={{ left: caret.x * zoom, top: caret.y * zoom, height: caret.h * zoom }} />
              )}
            </div>
          ))}
        </div>
      </div>
    );
  },
);
