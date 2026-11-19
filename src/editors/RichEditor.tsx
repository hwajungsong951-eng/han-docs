import { forwardRef, useImperativeHandle, useEffect, useRef, useState } from "react";
import { useEditor, EditorContent, Node as TipTapNode, mergeAttributes } from "@tiptap/react";
import StarterKit from "@tiptap/starter-kit";
import { TableKit } from "@tiptap/extension-table";
import { TextStyle, Color, FontFamily, FontSize } from "@tiptap/extension-text-style";
import TextAlign from "@tiptap/extension-text-align";
import mammoth from "mammoth";
import html2canvas from "html2canvas-pro";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { HwpDocument } from "../wasm/rhwp.js";
import { ensureWasm } from "../wasm/init";
import {
  Bold, Italic, Underline, Strikethrough, AlignLeft, AlignCenter, AlignRight, AlignJustify,
  List, ListOrdered, Table as TableIcon, Undo2, Redo2, Rows3, Columns3, Trash2, Merge, Split, Settings2,
} from "lucide-react";
import { TBtn, Sep, TSelect, FONTS, SIZES } from "../components/ui";
import type { DocFormat, ExportFormat } from "../lib/fileio";
import { tiptapJsonToDocx, recoverAltChunkHtml } from "../lib/docxWriter";
import { htmlForHwp } from "../lib/hwpHtml";
import { decodeTxt, encodeTxt, type TxtEncoding } from "../lib/textCodec";

/** DOCX 안의 그림을 보존하기 위한 인라인 이미지 노드 */
const ImageNode = TipTapNode.create({
  name: "image",
  inline: true,
  group: "inline",
  atom: true,
  draggable: true,
  addAttributes() {
    return {
      src: { default: null },
      alt: { default: null },
      width: { default: null, parseHTML: (el: HTMLElement) => el.getAttribute("width") },
      height: { default: null, parseHTML: (el: HTMLElement) => el.getAttribute("height") },
    };
  },
  parseHTML() { return [{ tag: "img[src]" }]; },
  renderHTML({ HTMLAttributes }) { return ["img", mergeAttributes(HTMLAttributes, { style: "max-width:100%;height:auto" })]; },
});

const num = (v: string, d: number) => { const n = parseFloat(v); return Number.isFinite(n) ? n : d; };

/** mammoth 가 버리는 문단 정렬 / 글자 크기를 HTML 에 살려 둔다. */
const MAMMOTH_SIZES = [6, 7, 8, 9, 10, 10.5, 11, 12, 13, 14, 15, 16, 18, 20, 22, 24, 26, 28, 32, 36, 40, 48, 60, 72];
const MAMMOTH_STYLE_MAP = [
  "p[style-name='Title'] => h1:fresh",
  "p[style-name='Subtitle'] => h2:fresh",
  "p[style-name='__al:center'] => p.al-center:fresh",
  "p[style-name='__al:right'] => p.al-right:fresh",
  "p[style-name='__al:both'] => p.al-justify:fresh",
  ...MAMMOTH_SIZES.map((n) => `r[style-name='__fs:${n}'] => span.fs-${String(n).replace(".", "_")}:fresh`),
];
function mammothTransform(doc: any): any {
  const walk = (el: any): any => {
    if (el.children) el = { ...el, children: el.children.map(walk) };
    if (el.type === "paragraph" && !el.styleName && !el.numbering && (el.alignment === "center" || el.alignment === "right" || el.alignment === "both"))
      return { ...el, styleId: null, styleName: `__al:${el.alignment}` };
    if (el.type === "run" && !el.styleName && el.fontSize && MAMMOTH_SIZES.includes(el.fontSize) && el.fontSize !== 10.5)
      return { ...el, styleId: null, styleName: `__fs:${el.fontSize}` };
    return el;
  };
  return walk(doc);
}
function mammothHtmlToTipTap(html: string): string {
  const d = new DOMParser().parseFromString(`<body>${html}</body>`, "text/html");
  d.querySelectorAll<HTMLElement>("p[class^='al-']").forEach((p) => {
    const a = p.className.replace("al-", ""); p.removeAttribute("class"); p.style.textAlign = a;
  });
  d.querySelectorAll<HTMLElement>("span[class^='fs-']").forEach((sp) => {
    const n = sp.className.replace("fs-", "").replace("_", "."); sp.removeAttribute("class"); sp.style.fontSize = `${n}pt`;
  });
  return d.body.innerHTML;
}

export interface DocEditorHandle {
  exportAs(fmt: ExportFormat): Promise<Uint8Array>;
}

type RichPageSettings = {
  top: string; bottom: string; left: string; right: string;
  header: string; footer: string;
  number: "none" | "footer" | "header";
};
const RICH_A4_DEFAULT: RichPageSettings = {
  top: "15", bottom: "15", left: "20", right: "20", header: "10", footer: "10", number: "none",
};

/** 편집 영역을 고해상도로 캡처해 A4 PDF 페이지로 묶는다. */
async function elementToPdf(el: HTMLElement, settings: RichPageSettings): Promise<Uint8Array> {
  const canvas = await html2canvas(el, {
    scale: 2,
    backgroundColor: "#ffffff",
    useCORS: true,
    logging: false,
  });
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const A4_W = 595.28;
  const A4_H = 841.89;
  const pageW = Math.max(1, canvas.width);
  const pageH = Math.max(1, Math.round(pageW * (A4_H / A4_W)));
  const mmToPt = (mm: string) => (Number(mm) || 0) * 72 / 25.4;
  let pageIndex = 0;
  for (let offset = 0; offset < canvas.height || offset === 0; offset += pageH) {
    const sliceH = Math.min(pageH, Math.max(1, canvas.height - offset));
    const pageCanvas = document.createElement("canvas");
    pageCanvas.width = pageW;
    pageCanvas.height = pageH;
    const ctx = pageCanvas.getContext("2d");
    if (!ctx) throw new Error("PDF 페이지를 만들 수 없습니다");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, pageW, pageH);
    ctx.drawImage(canvas, 0, offset, pageW, sliceH, 0, 0, pageW, sliceH);
    const dataUrl = pageCanvas.toDataURL("image/jpeg", 0.92);
    const bytes = Uint8Array.from(atob(dataUrl.split(",")[1]), (c) => c.charCodeAt(0));
    const image = await pdf.embedJpg(bytes);
    const page = pdf.addPage([A4_W, A4_H]);
    page.drawImage(image, { x: 0, y: 0, width: A4_W, height: A4_H });
    pageIndex += 1;
    if (settings.number !== "none") {
      const label = String(pageIndex);
      const size = 9;
      const textWidth = font.widthOfTextAtSize(label, size);
      const y = settings.number === "footer"
        ? Math.max(8, mmToPt(settings.footer))
        : Math.min(A4_H - 18, A4_H - mmToPt(settings.header));
      page.drawText(label, { x: (A4_W - textWidth) / 2, y, size, font, color: rgb(0, 0, 0) });
    }
    if (canvas.height <= pageH) break;
  }
  return await pdf.save();
}

/** docx / txt 편집기 (TipTap 기반) */
export const RichEditor = forwardRef<DocEditorHandle, { bytes: Uint8Array | null; format: DocFormat; onDirty: () => void; onError?: (m: string) => void }>(
  ({ bytes, format, onDirty, onError }, ref) => {
    const isTxt = format === "txt";
    const pageRef = useRef<HTMLDivElement>(null);
    const pagePanelRef = useRef<HTMLDivElement>(null);
    const [pageSettings, setPageSettings] = useState<RichPageSettings>(RICH_A4_DEFAULT);
    const [pageOpen, setPageOpen] = useState(false);
    const [editorSlow, setEditorSlow] = useState(false);
    const [loadError, setLoadError] = useState<string | null>(null);
    // TXT 는 열었던 인코딩/줄바꿈 방식 그대로 저장한다.
    const txtMeta = useRef<{ encoding: TxtEncoding; eol: "\r\n" | "\n" }>({ encoding: "utf-8", eol: "\r\n" });

    useEffect(() => {
      if (!pageOpen) return;
      const onDown = (e: MouseEvent) => {
        if (pagePanelRef.current && !pagePanelRef.current.contains(e.target as Node)) setPageOpen(false);
      };
      document.addEventListener("mousedown", onDown);
      return () => document.removeEventListener("mousedown", onDown);
    }, [pageOpen]);
    const editor = useEditor({
      extensions: [
        StarterKit,
        ImageNode,
        TableKit.configure({ table: { resizable: true } }),
        TextStyle, Color, FontFamily, FontSize,
        TextAlign.configure({ types: ["heading", "paragraph"] }),
      ],
      content: "",
      autofocus: false,
      editable: true,
      shouldRerenderOnTransaction: true,
      onUpdate: () => onDirty(),
    });

    useEffect(() => {
      if (editor) { setEditorSlow(false); return; }
      const t = setTimeout(() => setEditorSlow(true), 5000);
      return () => clearTimeout(t);
    }, [editor]);

    // 파일 로드
    useEffect(() => {
      if (!editor) return;
      let cancelled = false;
      // 불러온 내용은 실행 취소(Ctrl+Z) 기록에 남기지 않는다. (남기면 Ctrl+Z 한 번에 문서가 비어 버린다)
      const load = (html: string) => {
        if (cancelled || editor.isDestroyed) return;
        editor.chain().setMeta("addToHistory", false).setContent(html || "<p></p>", { emitUpdate: false }).run();
        editor.commands.focus("start");
      };
      (async () => {
        setLoadError(null);
        if (!bytes) { load("<p></p>"); return; }
        try {
          if (format === "docx") {
            const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
            const { value, messages } = await mammoth.convertToHtml({ arrayBuffer: buf }, {
              styleMap: MAMMOTH_STYLE_MAP,
              transformDocument: mammothTransform,
            } as any);
            if (messages.some((m) => /altChunk/.test(m.message)) && !value.trim()) {
              // 예전 버전 앱이 저장한 DOCX (내용이 altChunk 로 들어 있음) → HTML 을 꺼내 복구한다.
              const recovered = await recoverAltChunkHtml(bytes);
              if (!recovered) throw new Error("이전 버전 앱에서 저장된 DOCX 라 내용을 읽을 수 없습니다");
              load(recovered);
              onError?.("이전 버전 앱으로 저장된 DOCX 를 복구해서 열었습니다. 다시 저장하면 정상 DOCX 로 저장됩니다.");
              return;
            }
            load(mammothHtmlToTipTap(value));
          } else {
            const { text, encoding, eol } = decodeTxt(bytes);
            txtMeta.current = { encoding, eol };
            const html = text.split(/\r\n|\n|\r/).map((l) => `<p>${escapeHtml(l)}</p>`).join("");
            load(html);
          }
        } catch (e) {
          const msg = `문서를 불러오지 못했습니다: ${e instanceof Error ? e.message : String(e)}`;
          setLoadError(msg); onError?.(msg);
          load("<p></p>");
        }
      })();
      return () => { cancelled = true; };
    }, [editor, bytes, format]);

    useImperativeHandle(ref, () => ({
      async exportAs(fmt) {
        if (fmt === "pdf") {
          if (!pageRef.current) throw new Error("PDF로 변환할 편집 영역이 없습니다");
          return await elementToPdf(pageRef.current, pageSettings);
        }
        if (fmt === "hwp" || fmt === "hwpx") {
          if (!editor) throw new Error("편집기 준비 안됨");
          await ensureWasm();
          const doc = HwpDocument.createEmpty();
          try {
            doc.createBlankDocument();
            // 편집기의 페이지 여백(mm)을 HWP 용지 설정에 먼저 반영하고, 그 본문 폭에 맞춰 표 너비를 계산한다.
            const hu = (mm: number) => Math.round((mm * 7200) / 25.4);
            doc.setPageDef(0, JSON.stringify({
              width: hu(210), height: hu(297),
              marginTop: hu(num(pageSettings.top, 15)), marginBottom: hu(num(pageSettings.bottom, 15)),
              marginLeft: hu(num(pageSettings.left, 20)), marginRight: hu(num(pageSettings.right, 20)),
              marginHeader: hu(num(pageSettings.header, 10)), marginFooter: hu(num(pageSettings.footer, 10)),
            }));
            const info = JSON.parse(doc.getPageInfo(0)) as { bodyLeft: number; bodyRight: number };
            doc.pasteHtml(0, 0, 0, htmlForHwp(editor.getHTML(), info.bodyRight - info.bodyLeft - 24) /* rhwp 는 붙여넣은 표를 본문보다 약 24px 안쪽에서 시작한다 */);
            return fmt === "hwp" ? doc.exportHwp() : doc.exportHwpx();
          } finally {
            doc.free();
          }
        }
        if (!editor) throw new Error("편집기 준비 안됨");
        if (fmt === "txt") {
          const text = editor.getText({ blockSeparator: "\n" }).replace(/\n/g, txtMeta.current.eol);
          return encodeTxt(text, txtMeta.current.encoding);
        }
        if (fmt === "docx") {
          return await tiptapJsonToDocx(editor.getJSON(), {
            top: num(pageSettings.top, 15), bottom: num(pageSettings.bottom, 15), left: num(pageSettings.left, 20), right: num(pageSettings.right, 20),
            header: num(pageSettings.header, 10), footer: num(pageSettings.footer, 10), number: pageSettings.number,
          });
        }
        throw new Error(`${fmt} 형식으로는 저장할 수 없습니다`);
      },
    }), [editor, pageSettings]);

    if (!editor) {
      return (
        <div className="grid h-full place-items-center text-sm text-gray-500">
          {editorSlow ? "편집기를 초기화하지 못했습니다. 앱을 다시 실행해 주세요." : "문서 편집기를 준비하는 중…"}
        </div>
      );
    }

    const s = {
      bold: editor.isActive("bold"),
      italic: editor.isActive("italic"),
      underline: editor.isActive("underline"),
      strike: editor.isActive("strike"),
      left: editor.isActive({ textAlign: "left" }),
      center: editor.isActive({ textAlign: "center" }),
      right: editor.isActive({ textAlign: "right" }),
      justify: editor.isActive({ textAlign: "justify" }),
      inTable: editor.isActive("table"),
      font: (editor.getAttributes("textStyle").fontFamily as string) ?? "",
      size: ((editor.getAttributes("textStyle").fontSize as string) ?? "").replace("pt", ""),
      color: (editor.getAttributes("textStyle").color as string) ?? "#000000",
    };
    const c = () => editor.chain().focus();

    return (
      <div className="flex flex-col h-full">
        <div className="flex flex-wrap items-center gap-0.5 px-3 py-1.5 bg-white border-b border-gray-200">
          <TBtn title="실행 취소" onClick={() => c().undo().run()}><Undo2 size={16} /></TBtn>
          <TBtn title="다시 실행" onClick={() => c().redo().run()}><Redo2 size={16} /></TBtn>
          {!isTxt && (<>
            <Sep />
            <TSelect title="글꼴" value={s.font} options={FONTS} onChange={(v) => v ? c().setFontFamily(v).run() : c().unsetFontFamily().run()} />
            <TSelect title="크기" width="w-16" value={s.size} options={[["", "크기"], ...SIZES.map((x) => [x, x] as [string, string])]}
              onChange={(v) => v ? c().setFontSize(`${v}pt`).run() : c().unsetFontSize().run()} />
            <input type="color" title="글자 색" value={s.color} onChange={(e) => c().setColor(e.target.value).run()}
              className="h-8 w-8 p-1 rounded-md border border-gray-200 bg-white cursor-pointer" />
            <Sep />
            <TBtn title="굵게 (Ctrl+B)" active={s.bold} onClick={() => c().toggleBold().run()}><Bold size={16} /></TBtn>
            <TBtn title="기울임 (Ctrl+I)" active={s.italic} onClick={() => c().toggleItalic().run()}><Italic size={16} /></TBtn>
            <TBtn title="밑줄 (Ctrl+U)" active={s.underline} onClick={() => c().toggleUnderline().run()}><Underline size={16} /></TBtn>
            <TBtn title="취소선" active={s.strike} onClick={() => c().toggleStrike().run()}><Strikethrough size={16} /></TBtn>
            <Sep />
            <TBtn title="왼쪽 정렬" active={s.left} onClick={() => c().setTextAlign("left").run()}><AlignLeft size={16} /></TBtn>
            <TBtn title="가운데 정렬" active={s.center} onClick={() => c().setTextAlign("center").run()}><AlignCenter size={16} /></TBtn>
            <TBtn title="오른쪽 정렬" active={s.right} onClick={() => c().setTextAlign("right").run()}><AlignRight size={16} /></TBtn>
            <TBtn title="양쪽 정렬" active={s.justify} onClick={() => c().setTextAlign("justify").run()}><AlignJustify size={16} /></TBtn>
            <Sep />
            <TBtn title="글머리 목록" onClick={() => c().toggleBulletList().run()}><List size={16} /></TBtn>
            <TBtn title="번호 목록" onClick={() => c().toggleOrderedList().run()}><ListOrdered size={16} /></TBtn>
            <Sep />
            <TBtn title="표 삽입 (3×3)" onClick={() => c().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()}><TableIcon size={16} /></TBtn>
            {s.inTable && (<>
              <TBtn title="행 추가" onClick={() => c().addRowAfter().run()}><Rows3 size={16} /></TBtn>
              <TBtn title="열 추가" onClick={() => c().addColumnAfter().run()}><Columns3 size={16} /></TBtn>
              <TBtn title="셀 병합" onClick={() => c().mergeCells().run()}><Merge size={16} /></TBtn>
              <TBtn title="셀 나누기" onClick={() => c().splitCell().run()}><Split size={16} /></TBtn>
              <TBtn title="표 삭제" onClick={() => c().deleteTable().run()}><Trash2 size={16} /></TBtn>
            </>)}
          </>)}
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
                  ] as [keyof RichPageSettings, string][]).map(([key, label]) => (
                    <label key={key} className="flex items-center justify-between gap-2 text-[11px] text-gray-600">
                      <span>{label}</span>
                      <input type="number" min="0" step="1" title={label} value={pageSettings[key] as string}
                        onChange={(e) => setPageSettings((s2) => ({ ...s2, [key]: e.target.value }))}
                        className="h-7 w-16 rounded-md border border-gray-200 px-2 text-right text-xs" />
                    </label>
                  ))}
                </div>
                <label className="mt-2 flex items-center justify-between gap-2 text-[11px] text-gray-600">
                  <span>쪽 번호</span>
                  <select title="쪽 번호" value={pageSettings.number}
                    onChange={(e) => setPageSettings((s2) => ({ ...s2, number: e.target.value as RichPageSettings["number"] }))}
                    className="h-7 rounded-md border border-gray-200 px-2 text-xs">
                    <option value="none">없음</option>
                    <option value="footer">아래 가운데</option>
                    <option value="header">위 가운데</option>
                  </select>
                </label>
                <div className="mt-3 flex items-center justify-between">
                  <button onMouseDown={(e) => e.preventDefault()} onClick={() => setPageSettings(RICH_A4_DEFAULT)}
                    className="h-7 rounded-md border border-gray-200 px-2 text-xs text-gray-600 hover:bg-gray-50">A4 기본값</button>
                  <button onMouseDown={(e) => e.preventDefault()} onClick={() => setPageOpen(false)}
                    className="h-7 rounded-md bg-blue-600 px-3 text-xs text-white hover:bg-blue-700">적용</button>
                </div>
              </div>
            )}
          </div>
        </div>
        {loadError && (
          <div className="mx-4 mt-3 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700">{loadError}</div>
        )}
        {/* 회색 작업 영역 어디를 눌러도 편집기에 커서가 들어가도록 한다 */}
        <div className="flex-1 overflow-auto py-8 px-4" onMouseDown={(e) => {
          if (!(e.target as HTMLElement).closest(".ProseMirror") && !(e.target as HTMLElement).closest("button,input,select")) {
            e.preventDefault(); editor.chain().focus("end").run();
          }
        }}>
          <div className="relative mx-auto" style={{ width: "210mm", maxWidth: "none" }}>
            <div ref={pageRef} className={`bg-white shadow-sm border border-gray-200 box-border w-full ${isTxt ? "font-mono text-sm txt-page" : ""}`}
              style={{
                minHeight: "297mm",
                padding: `${pageSettings.top}mm ${pageSettings.right}mm ${pageSettings.bottom}mm ${pageSettings.left}mm`,
                // 페이지 빈 곳을 눌러도 편집기 안으로 인식되도록 편집 영역을 페이지 높이만큼 채운다
                ["--pm-min" as string]: `calc(297mm - ${num(pageSettings.top, 15) + num(pageSettings.bottom, 15)}mm)`,
              }}>
              <EditorContent editor={editor} />
            </div>
            {pageSettings.number !== "none" && (
              <div className="pointer-events-none absolute left-0 right-0 text-center text-[11px] text-gray-500"
                style={pageSettings.number === "footer"
                  ? { bottom: `${pageSettings.footer}mm` }
                  : { top: `${pageSettings.header}mm` }}>1</div>
            )}
          </div>
        </div>
      </div>
    );
  },
);

function escapeHtml(s: string) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;") || "<br>";
}
