/** 그림 선택/크기 조절에 쓰는 보조 함수 */

export type PicRef = { sec: number; para: number; ctrl: number };
export type Box = { x: number; y: number; w: number; h: number };
export type Handle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

export const HANDLES: Handle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
export const HU_PER_MM = 7200 / 25.4;
export const HU_PER_PX = 75; // 96dpi 에서 1px = 75 HWPUNIT
export const huToMm = (hu: number) => Math.round((hu / HU_PER_MM) * 10) / 10;
export const mmToHu = (mm: number) => Math.round(mm * HU_PER_MM);

export const handleCursor: Record<Handle, string> = {
  nw: "nwse-resize", se: "nwse-resize", ne: "nesw-resize", sw: "nesw-resize",
  n: "ns-resize", s: "ns-resize", e: "ew-resize", w: "ew-resize",
};

/** 손잡이 위치 (그림 상자 기준 0~1) */
export const handlePos: Record<Handle, [number, number]> = {
  nw: [0, 0], n: [0.5, 0], ne: [1, 0], e: [1, 0.5], se: [1, 1], s: [0.5, 1], sw: [0, 1], w: [0, 0.5],
};

const isCorner = (h: Handle) => h.length === 2;

/**
 * 손잡이를 (dx, dy)만큼 끌었을 때의 새 크기.
 * 모서리 손잡이는 가로세로 비율을 유지하고, 가장자리 손잡이는 한 방향만 바꾼다.
 * Shift 를 누르면 가장자리에서도 비율 유지, Alt 를 누르면 모서리에서도 자유 변형.
 */
export function resizeBox(box: Box, handle: Handle, dx: number, dy: number, shift: boolean, alt: boolean): { w: number; h: number } {
  const sx = handle.includes("e") ? 1 : handle.includes("w") ? -1 : 0;
  const sy = handle.includes("s") ? 1 : handle.includes("n") ? -1 : 0;
  let w = box.w + sx * dx, h = box.h + sy * dy;
  const keep = (isCorner(handle) && !alt) || (!isCorner(handle) && shift);
  if (keep && box.w > 0 && box.h > 0) {
    const kw = w / box.w, kh = h / box.h;
    let k: number;
    if (sx && sy) k = Math.max(kw, kh);
    else if (sx) k = kw;
    else k = kh;
    w = box.w * k; h = box.h * k;
  }
  return { w: Math.max(8, w), h: Math.max(8, h) };
}
