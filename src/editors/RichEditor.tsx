import { forwardRef, useImperativeHandle, useEffect, useRef, useState } from "react";
import { useEditor, EditorContent, useEditorState } from "@tiptap/react";
import StarterKit from "@tiptap/starter-kit";
import { TableKit } from "@tiptap/extension-table";
import { TextStyle, Color, FontFamily, FontSize } from "@tiptap/extension-text-style";
import TextAlign from "@tiptap/extension-text-align";
import mammoth from "mammoth";
import { asBlob } from "html-docx-js-typescript";
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
export const RichEditor = forwardRef<DocEditorHandle, { bytes: Uint8Array | null; format: DocFormat; onDirty: () => void }>(
  ({ bytes, format, onDirty }, ref) => {
    const isTxt = format === "txt";
    const pageRef = useRef<HTMLDivElement>(null);
    const pagePanelRef = useRef<HTMLDivElement>(null);
    const [pageSettings, setPageSettings] = useState<RichPageSettings>(RICH_A4_DEFAULT);
    const [pageOpen, setPageOpen] = useState(false);

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
        TableKit.configure({ table: { resizable: true } }),
        TextStyle, Color, FontFamily, FontSize,
        TextAlign.configure({ types: ["heading", "paragraph"] }),
      ],
      content: "",
      onUpdate: () => onDirty(),
    });

    // 파일 로드
    useEffect(() => {
      if (!editor) return;
      (async () => {
        if (!bytes) { editor.commands.setContent("<p></p>", { emitUpdate: false }); return; }
        if (format === "docx") {
          const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
          const { value } = await mammoth.convertToHtml({ arrayBuffer: buf });
          editor.commands.setContent(value || "<p></p>", { emitUpdate: false });
        } else {
          const text = decodeText(bytes);
          const html = text.split(/\r?\n/).map((l) => `<p>${escapeHtml(l)}</p>`).join("");
          editor.commands.setContent(html, { emitUpdate: false });
        }
      })();
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
            doc.pasteHtml(0, 0, 0, editor.getHTML());
            return fmt === "hwp" ? doc.exportHwp() : doc.exportHwpx();
          } finally {
            doc.free();
          }
        }
        if (!editor) throw new Error("편집기 준비 안됨");
        if (fmt === "txt") return new TextEncoder().encode(editor.getText({ blockSeparator: "\n" }));
        if (fmt === "docx") {
          const html = `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${editor.getHTML()}</body></html>`;
          const blob = (await asBlob(html)) as Blob;
          return new Uint8Array(await blob.arrayBuffer());
        }
        throw new Error(`${fmt} 형식으로는 저장할 수 없습니다`);
      },
    }), [editor, pageSettings]);

    const s = useEditorState({
      editor,
      selector: ({ editor: e }) => e ? {
        bold: e.isActive("bold"), italic: e.isActive("italic"), underline: e.isActive("underline"), strike: e.isActive("strike"),
        left: e.isActive({ textAlign: "left" }), center: e.isActive({ textAlign: "center" }),
        right: e.isActive({ textAlign: "right" }), justify: e.isActive({ textAlign: "justify" }),
        inTable: e.isActive("table"),
        font: (e.getAttributes("textStyle").fontFamily as string) ?? "",
        size: ((e.getAttributes("textStyle").fontSize as string) ?? "").replace("pt", ""),
        color: (e.getAttributes("textStyle").color as string) ?? "#000000",
      } : null,
    });

    if (!editor || !s) return null;
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
        <div className="flex-1 overflow-auto py-8">
          <div className="relative mx-auto w-fit">
            <div ref={pageRef} className={`bg-white shadow-sm border border-gray-200 ${isTxt ? "max-w-4xl font-mono text-sm" : "max-w-[210mm]"} min-h-[297mm]`}
              style={{ padding: `${pageSettings.top}mm ${pageSettings.right}mm ${pageSettings.bottom}mm ${pageSettings.left}mm` }}>
              <EditorContent editor={editor} className="h-full" />
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

function decodeText(b: Uint8Array): string {
  try { return new TextDecoder("utf-8", { fatal: true }).decode(b); }
  catch { return new TextDecoder("euc-kr").decode(b); } // 한글 윈도우 ANSI 텍스트 대응
}

function escapeHtml(s: string) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;") || "<br>";
}
