import { HORZ_ALIGN, HORZ_REL, VERT_ALIGN, VERT_REL, WRAP_OPTIONS, floatingDefaults, type Placement } from "../lib/hwpObject";

const selCls = "h-6 rounded border border-gray-200 bg-white px-1 text-[11px] disabled:bg-gray-50 disabled:text-gray-300";
const numCls = "h-6 w-[58px] rounded border border-gray-200 px-1 text-right text-[11px] disabled:bg-gray-50 disabled:text-gray-300";

/** 개체 속성의 "위치" 칸: 글자처럼 취급 / 본문과의 배치 / 가로·세로 위치 */
export default function PlacementFields({
  value, onChange, kind, original,
}: { value: Placement; onChange: (p: Placement) => void; kind: "table" | "picture"; original?: Placement }) {
  const set = (p: Partial<Placement>) => onChange({ ...value, ...p });
  const off = value.inline;
  const wraps = WRAP_OPTIONS.some(([v]) => v === value.wrap) ? WRAP_OPTIONS : [...WRAP_OPTIONS, [value.wrap, value.wrap] as [string, string]];
  return (
    <div className="mb-2 rounded-md border border-gray-100 bg-gray-50/60 p-2">
      <label className="mb-1.5 flex items-center gap-2 text-gray-700">
        <input type="checkbox" checked={value.inline}
          onChange={(e) => onChange(e.target.checked ? { ...value, inline: true } : original && !original.inline ? { ...original } : floatingDefaults({ ...value, inline: false }))} />
        글자처럼 취급
      </label>
      <div className={`mb-1.5 ${off ? "opacity-40" : ""}`}>
        <div className="mb-0.5 text-[11px] text-gray-500">본문과의 배치</div>
        <div className="grid grid-cols-2 gap-x-2 gap-y-0.5">
          {wraps.map(([v, label]) => (
            <label key={v} className="flex items-center gap-1 text-[11px] text-gray-700">
              <input type="radio" name="obj-wrap" disabled={off} checked={value.wrap === v} onChange={() => set({ wrap: v })} /> {label}
            </label>
          ))}
        </div>
      </div>
      <div className={`grid grid-cols-[28px_1fr] items-center gap-x-1 gap-y-1 text-[11px] text-gray-600 ${off ? "opacity-40" : ""}`}>
        <span>가로</span>
        <div className="flex items-center gap-1">
          <select title="가로 기준" disabled={off} className={selCls} value={value.horzRelTo} onChange={(e) => set({ horzRelTo: e.target.value })}>
            {HORZ_REL.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select>
          <select title="가로 정렬" disabled={off} className={selCls} value={value.horzAlign} onChange={(e) => set({ horzAlign: e.target.value })}>
            {HORZ_ALIGN.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select>
          <input title="가로 위치 (mm)" type="number" step="0.5" disabled={off} className={numCls} value={value.horzMm} onChange={(e) => set({ horzMm: e.target.value })} />
          <span>mm</span>
        </div>
        <span>세로</span>
        <div className="flex items-center gap-1">
          <select title="세로 기준" disabled={off} className={selCls} value={value.vertRelTo} onChange={(e) => set({ vertRelTo: e.target.value })}>
            {VERT_REL.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select>
          <select title="세로 정렬" disabled={off} className={selCls} value={value.vertAlign} onChange={(e) => set({ vertAlign: e.target.value })}>
            {VERT_ALIGN.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select>
          <input title="세로 위치 (mm)" type="number" step="0.5" disabled={off} className={numCls} value={value.vertMm} onChange={(e) => set({ vertMm: e.target.value })} />
          <span>mm</span>
        </div>
      </div>
      {!off && value.wrap === "Square" && (
        <div className="mt-1.5 text-[10px] leading-tight text-amber-600">
          어울림: 이 편집기 화면에서는 글이 {kind === "table" ? "표를" : "그림을"} 피해 흐르지 않고 겹쳐 보일 수 있습니다. 겹치기 싫으면 "자리 차지"를 쓰세요.
        </div>
      )}
    </div>
  );
}
