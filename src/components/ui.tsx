import { ReactNode } from "react";

export function TBtn(props: { title: string; onClick: () => void; active?: boolean; disabled?: boolean; children: ReactNode }) {
  return (
    <button
      title={props.title}
      disabled={props.disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={props.onClick}
      className={`h-8 min-w-8 px-1.5 inline-flex items-center justify-center rounded-md text-sm transition
        ${props.active ? "bg-blue-100 text-blue-700" : "text-gray-700 hover:bg-gray-100"}
        disabled:opacity-40 disabled:pointer-events-none`}
    >
      {props.children}
    </button>
  );
}

export function Sep() {
  return <div className="w-px h-5 bg-gray-200 mx-1" />;
}

export function TSelect(props: { title: string; value: string; options: [string, string][]; onChange: (v: string) => void; width?: string }) {
  return (
    <select
      title={props.title}
      value={props.value}
      onChange={(e) => props.onChange(e.target.value)}
      className={`h-8 ${props.width ?? "w-28"} rounded-md border border-gray-200 bg-white px-2 text-sm text-gray-700 hover:border-gray-300 focus:outline-none focus:ring-2 focus:ring-blue-200`}
    >
      {props.options.map(([v, label]) => (
        <option key={v} value={v}>{label}</option>
      ))}
    </select>
  );
}

export const FONTS: [string, string][] = [
  ["", "기본 글꼴"],
  ["Malgun Gothic, Apple SD Gothic Neo, sans-serif", "맑은 고딕"],
  ["Batang, AppleMyungjo, serif", "바탕"],
  ["Dotum, sans-serif", "돋움"],
  ["Gulim, sans-serif", "굴림"],
  ["Arial, sans-serif", "Arial"],
  ["Times New Roman, serif", "Times New Roman"],
];

export const SIZES = ["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32", "40"];
