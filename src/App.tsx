import { useRef, useState, useCallback, useEffect } from "react";
import { FolderOpen, Save, FilePlus2, FileText, FileDown, X } from "lucide-react";
import { RichEditor, type DocEditorHandle } from "./editors/RichEditor";
import { HwpEditor } from "./editors/HwpEditor";
import { pickOpenPath, pickSavePath, readBytes, writeBytes, formatOf, saveFormatOf, baseName, type DocFormat, type ExportFormat } from "./lib/fileio";

type Doc = { path: string | null; format: DocFormat; bytes: Uint8Array | null; key: number };

/** 형식별 저장 가능 대상 */
const SAVE_TARGETS: Record<DocFormat, ExportFormat[]> = {
  hwp: ["hwp", "hwpx", "pdf"], hwpx: ["hwpx", "hwp", "pdf"],
  docx: ["docx", "txt", "pdf"], txt: ["txt", "docx", "pdf"],
};

export default function App() {
  const [doc, setDoc] = useState<Doc | null>(null);
  const [dirty, setDirty] = useState(false);
  const [toast, setToast] = useState<string | null>(null);
  const editorRef = useRef<DocEditorHandle>(null);

  const notify = useCallback((m: string) => { setToast(m); setTimeout(() => setToast(null), 3500); }, []);
  const markDirty = useCallback(() => setDirty(true), []);

  const openFile = async () => {
    const path = await pickOpenPath(); if (!path) return;
    const format = formatOf(path);
    if (!format) return notify("지원하지 않는 형식입니다 (hwp, hwpx, docx, txt)");
    try {
      const bytes = await readBytes(path);
      setDoc({ path, format, bytes, key: Date.now() }); setDirty(false);
    } catch (e) { notify(String(e)); }
  };

  const newFile = (format: DocFormat) => { setDoc({ path: null, format, bytes: null, key: Date.now() }); setDirty(false); };

  const saveFile = async (saveAs = false, force?: ExportFormat) => {
    if (!doc || !editorRef.current) return;
    const wantsPdf = force === "pdf";
    let path = wantsPdf ? null : doc.path;
    if (saveAs || !path || wantsPdf) {
      const def = wantsPdf ? "새 문서.pdf" : (path ?? `새 문서.${doc.format}`);
      const targets = wantsPdf ? (["pdf"] as ExportFormat[]) : SAVE_TARGETS[doc.format];
      path = await pickSavePath(def, targets);
      if (!path) return;
    }
    const fmt = wantsPdf ? "pdf" : (saveFormatOf(path) ?? doc.format);
    try {
      if (fmt === "pdf") notify("PDF로 변환 중...");
      const data = await editorRef.current.exportAs(fmt);
      await writeBytes(path, data);
      if (fmt === "pdf") {
        // PDF는 내보내기이므로 원본 문서 경로/저장 상태는 그대로 둔다.
        notify(`PDF로 저장됨: ${baseName(path)}`);
      } else {
        setDoc({ ...doc, path }); setDirty(false);
        notify(`저장됨: ${baseName(path)}`);
      }
    } catch (e) { notify(String(e)); }
  };

  // 단축키
  useEffect(() => {
    const h = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey;
      if (mod && e.key.toLowerCase() === "s") { e.preventDefault(); saveFile(e.shiftKey); }
      if (mod && e.key.toLowerCase() === "o") { e.preventDefault(); openFile(); }
    };
    window.addEventListener("keydown", h);
    return () => window.removeEventListener("keydown", h);
  });

  const isHwp = doc?.format === "hwp" || doc?.format === "hwpx";

  return (
    <div className="flex flex-col h-full">
      {/* 상단 바 */}
      <header className="flex items-center gap-2 h-12 px-3 bg-white border-b border-gray-200">
        <div className="flex items-center gap-2 font-semibold text-gray-800 mr-2">
          <div className="w-7 h-7 rounded-lg bg-blue-600 text-white grid place-items-center text-xs">한</div>
          Docs
          <span title="2026-11-07 빌드" style={{ fontSize: 10, opacity: 0.55, fontWeight: 400 }}>BUILD 11-07</span>
        </div>
        <HBtn onClick={openFile} icon={<FolderOpen size={16} />}>열기</HBtn>
        <HBtn onClick={() => saveFile(false)} icon={<Save size={16} />} disabled={!doc}>저장</HBtn>
        <HBtn onClick={() => saveFile(true)} disabled={!doc}>다른 이름으로</HBtn>
        <HBtn onClick={() => saveFile(true, "pdf")} icon={<FileDown size={16} />} disabled={!doc}>PDF</HBtn>
        <div className="flex-1 text-center text-sm text-gray-600 truncate">
          {doc && (<>{doc.path ? baseName(doc.path) : `새 문서.${doc.format}`}{dirty && <span className="text-blue-600"> •</span>}
            <span className="ml-2 text-[11px] uppercase px-1.5 py-0.5 rounded bg-gray-100 text-gray-500">{doc.format}</span></>)}
        </div>
        {doc && <HBtn onClick={() => setDoc(null)} icon={<X size={16} />}>닫기</HBtn>}
      </header>

      <main className="flex-1 min-h-0">
        {!doc && <Welcome onOpen={openFile} onNew={newFile} />}
        {doc && isHwp && <HwpEditor key={doc.key} ref={editorRef} bytes={doc.bytes} onDirty={markDirty} onError={notify} />}
        {doc && !isHwp && <RichEditor key={doc.key} ref={editorRef} bytes={doc.bytes} format={doc.format} onDirty={markDirty} />}
      </main>

      {toast && (
        <div className="fixed bottom-5 left-1/2 -translate-x-1/2 bg-gray-900 text-white text-sm px-4 py-2 rounded-lg shadow-lg">{toast}</div>
      )}
    </div>
  );
}

function HBtn(p: { onClick: () => void; icon?: React.ReactNode; children: React.ReactNode; disabled?: boolean }) {
  return (
    <button onClick={p.onClick} disabled={p.disabled}
      className="h-8 px-3 inline-flex items-center gap-1.5 rounded-md text-sm text-gray-700 hover:bg-gray-100 disabled:opacity-40">
      {p.icon}{p.children}
    </button>
  );
}

function Welcome({ onOpen, onNew }: { onOpen: () => void; onNew: (f: DocFormat) => void }) {
  const cards: [DocFormat, string, string][] = [
    ["hwp", "한글 문서", "bg-sky-50 text-sky-700"], ["hwpx", "한글 문서 (XML)", "bg-cyan-50 text-cyan-700"],
    ["docx", "Word 문서", "bg-indigo-50 text-indigo-700"], ["txt", "텍스트", "bg-gray-100 text-gray-700"],
  ];
  return (
    <div className="h-full grid place-items-center">
      <div className="w-[560px]">
        <h1 className="text-2xl font-bold text-gray-800 mb-1">무엇을 작성할까요?</h1>
        <p className="text-sm text-gray-500 mb-6">hwp · hwpx · docx · txt 문서를 열고 편집할 수 있습니다.</p>
        <button onClick={onOpen} className="w-full h-12 mb-6 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-medium inline-flex items-center justify-center gap-2">
          <FolderOpen size={18} /> 문서 열기 <span className="text-blue-200 text-xs">Ctrl+O</span>
        </button>
        <div className="text-xs font-medium text-gray-400 mb-2">새 문서</div>
        <div className="grid grid-cols-4 gap-3">
          {cards.map(([f, label, cls]) => (
            <button key={f} onClick={() => onNew(f)} className="p-4 rounded-xl bg-white border border-gray-200 hover:border-blue-300 hover:shadow-sm text-left transition">
              <div className={`w-9 h-9 rounded-lg grid place-items-center mb-3 ${cls}`}>{f === "txt" ? <FileText size={18} /> : <FilePlus2 size={18} />}</div>
              <div className="text-sm font-medium text-gray-800">{label}</div>
              <div className="text-xs text-gray-400 uppercase">.{f}</div>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}
