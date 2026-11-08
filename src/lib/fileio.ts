import { invoke } from "@tauri-apps/api/core";
import { open, save } from "@tauri-apps/plugin-dialog";

export type DocFormat = "hwp" | "hwpx" | "docx" | "txt";
export type ExportFormat = DocFormat | "pdf";

export const SUPPORTED: DocFormat[] = ["hwp", "hwpx", "docx", "txt"];
export const EXPORT_SUPPORTED: ExportFormat[] = ["hwp", "hwpx", "docx", "txt", "pdf"];

export function formatOf(path: string): DocFormat | null {
  const ext = path.split(".").pop()?.toLowerCase();
  return SUPPORTED.includes(ext as DocFormat) ? (ext as DocFormat) : null;
}

export function saveFormatOf(path: string): ExportFormat | null {
  const ext = path.split(".").pop()?.toLowerCase();
  return EXPORT_SUPPORTED.includes(ext as ExportFormat) ? (ext as ExportFormat) : null;
}

export async function pickOpenPath(): Promise<string | null> {
  const p = await open({
    multiple: false,
    filters: [
      { name: "지원 문서", extensions: SUPPORTED },
      { name: "한글", extensions: ["hwp", "hwpx"] },
      { name: "Word", extensions: ["docx"] },
      { name: "텍스트", extensions: ["txt"] },
    ],
  });
  return typeof p === "string" ? p : null;
}

export async function pickSavePath(defaultPath: string, exts: ExportFormat[]): Promise<string | null> {
  return await save({ defaultPath, filters: exts.map((e) => ({ name: e.toUpperCase(), extensions: [e] })) });
}

export async function readBytes(path: string): Promise<Uint8Array> {
  const arr = await invoke<number[]>("read_file_bytes", { path });
  return new Uint8Array(arr);
}

export async function writeBytes(path: string, data: Uint8Array): Promise<void> {
  await invoke("write_file_bytes", { path, data: Array.from(data) });
}

export function baseName(path: string) {
  return path.split(/[\\/]/).pop() ?? path;
}
