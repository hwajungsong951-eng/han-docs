import { forwardRef, useEffect, useImperativeHandle, useRef, useState, useCallback } from "react";
import init, { HwpDocument } from "../wasm/rhwp.js";
import { open } from "@tauri-apps/plugin-dialog";
import { readBytes } from "../lib/fileio";
import { PDFDocument } from "pdf-lib";
import {
  Bold, Italic, Underline, Strikethrough, AlignLeft, AlignCenter, AlignRight, AlignJustify,
  Table as TableIcon, Square, Circle, Type, Minus, Undo2, Redo2, ZoomIn, ZoomOut, Image as ImageIcon, Smile, Settings2,
} from "lucide-react";
import { TBtn, Sep, TSelect, SIZES } from "../components/ui";
import type { ExportFormat } from "../lib/fileio";
import type { DocEditorHandle } from "./RichEditor";

let wasmReady: Promise<unknown> | null = null;
const ensureWasm = () => {
  if (!wasmReady) {
    // 실패한 Promise를 캐시하지 않는다. 다음 문서를 열 때 초기화를 다시 시도한다.
    wasmReady = init().catch((err) => {
      wasmReady = null;
      throw err;
    });
  }
  return wasmReady;
};

type CellEntry = { controlIndex: number; cellIndex: number; cellParaIndex: number };
type CellPos = { parentPara: number; path: CellEntry[] };
type Pos = { sec: number; para: number; off: number; cell?: CellPos };
type Caret = { page: number; x: number; y: number; h: number };
type SelRect = { page: number; x: number; y: number; w: number; h: number };
type SelRange = { start: Pos; end: Pos };

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
          setPageCount(0); setCaret(null); setSelRects([]); setVisiblePages(new Set([0]));

          let doc: HwpDocument;
          if (bytes) doc = new HwpDocument(bytes);
          else { doc = HwpDocument.createEmpty(); doc.createBlankDocument(); }
          if (cancelled) { doc.free(); return; }
          docRef.current = doc;
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
    }, [bytes, reloadKey, onError]);

    // 화면에 보이는 페이지만 렌더링한다. (큰 문서에서 키 입력마다 전체 페이지를
    // 다시 그리던 지연을 없앤다.)
    useEffect(() => {
      const doc = docRef.current;
      if (!doc) return;
      for (let i = 0; i < pageCount; i++) {
        if (!visiblePages.has(i)) continue;
        const c = canvases.current[i];
        if (c) { try { doc.renderPageToCanvas(i, c, zoom); } catch (e) { console.error(e); } }
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
      setPageCount(d.pageCount()); setRev((r) => r + 1);
    };
    const redo = () => {
      const d = docRef.current; const id = redoStack.current.pop();
      if (!d || id === undefined) return;
      lastEditRef.current = null;
      undoStack.current.push(d.saveSnapshot());
      d.restoreSnapshot(id); d.discardSnapshot(id);
      setPageCount(d.pageCount()); setRev((r) => r + 1);
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
    }, [hitTestClient]);

    // 마우스 클릭/드래그 → 커서 이동 및 선택
    const onPageMouseDown = (page: number, e: React.MouseEvent<HTMLCanvasElement>) => {
      // mousedown 기본 동작이 textarea 포커스를 BODY로 되돌리므로 막는다.
      e.preventDefault();
      const d = docRef.current; if (!d) return;
      const rect = e.currentTarget.getBoundingClientRect();
      const x = (e.clientX - rect.left) / zoom, y = (e.clientY - rect.top) / zoom;
      let p: Pos;
      try {
        p = posFromHit(JSON.parse(d.hitTest(page, x, y)));
      } catch (err) { console.warn(err); return; }

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
      edit((d) => {
        let p = pos;
        if (range) { deleteRange(d, range); p = range.start; }
        const charCount = [...text].length;
        const plain = !/[\r\n\t]/.test(text) && charCount <= 8;
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

    const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if (e.nativeEvent.isComposing || !pos) return;
      const mod = e.ctrlKey || e.metaKey;
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
          moveCaretVertical(-10, e.shiftKey);
          break;
        case "PageDown":
          e.preventDefault(); clearPending();
          moveCaretVertical(10, e.shiftKey);
          break;
        case "ArrowLeft":
          e.preventDefault(); clearPending();
          moveHorizontal(-1, e.shiftKey);
          break;
        case "ArrowRight":
          e.preventDefault(); clearPending();
          moveHorizontal(1, e.shiftKey);
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
          e.preventDefault(); clearPending(); setAnchor(null); setSelRects([]);
          break;
      }
    };

    // 글자 서식
    const charFmt = (props: Record<string, unknown>) => {
      const d = docRef.current;
      if (!d) return;
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
      const rows = Number(prompt("행 개수", "3")); const cols = Number(prompt("열 개수", "3"));
      if (!rows || !cols) return;
      edit((d) => d.createTable(pos.sec, pos.para, pos.off, rows, cols));
    };
    const insertImage = async () => {
      if (!pos) return;
      try {
        const path = await open({
          multiple: false,
          filters: [{ name: "이미지", extensions: ["png", "jpg", "jpeg", "gif", "bmp", "webp"] }],
        });
        if (typeof path !== "string") return;
        const bytes = await readBytes(path);
        const ext = (path.split(".").pop() || "png").toLowerCase();
        const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
        const bmp = await createImageBitmap(new Blob([buf], { type: `image/${ext === "jpg" ? "jpeg" : ext}` }));
        const scale = Math.min(1, 400 / bmp.width);
        const wPx = Math.max(1, Math.round(bmp.width * scale));
        const hPx = Math.max(1, Math.round(bmp.height * scale));
        const p = pos;
        edit((d) => d.insertPicture(
          p.sec, p.cell ? p.cell.parentPara : p.para, p.off,
          p.cell ? cellPathJsonFor(p) : "[]",
          bytes, Math.round(wPx * 75), Math.round(hPx * 75),
          bmp.width, bmp.height, ext, "그림",
        ));
        bmp.close();
      } catch (e) { onError(`그림 삽입 실패: ${errorText(e)}`); }
    };
    const insertShape = (shapeType: "rectangle" | "ellipse" | "textbox" | "line") => {
      if (!pos) return;
      edit((d) => d.createShapeControl(JSON.stringify({
        sectionIdx: pos.sec, paraIdx: pos.para, charOffset: pos.off,
        shapeType, width: 14173, height: shapeType === "line" ? 0 : 7087, // HWPUNIT (5cm × 2.5cm)
      })));
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
          <span className="text-[11px] text-gray-400 px-1">삽입</span>
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
          <TBtn title="표 삽입" onClick={insertTable}><TableIcon size={16} /></TBtn>
          <TBtn title="그림 삽입" onClick={insertImage}><ImageIcon size={16} /></TBtn>
          <TBtn title="사각형 삽입" onClick={() => insertShape("rectangle")}><Square size={16} /></TBtn>
          <TBtn title="타원 삽입" onClick={() => insertShape("ellipse")}><Circle size={16} /></TBtn>
          <TBtn title="글상자 삽입" onClick={() => insertShape("textbox")}><Type size={16} /></TBtn>
          <TBtn title="선 삽입" onClick={() => insertShape("line")}><Minus size={16} /></TBtn>
          <Sep />
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
              <canvas ref={(el) => { canvases.current[i] = el; }} onMouseDown={(e) => onPageMouseDown(i, e)} onDoubleClick={(e) => onPageDoubleClick(i, e)} className="block cursor-text" />
              {selRects.filter((r) => r.page === i).map((r, j) => (
                <div key={j} className="absolute bg-blue-500/25 pointer-events-none"
                  style={{ left: r.x * zoom, top: r.y * zoom, width: r.w * zoom, height: r.h * zoom }} />
              ))}
              {caret && caret.page === i && (
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
