/** 표/그림 같은 개체의 "위치" 속성 (한컴 한글 개체 속성 > 기본 > 위치) */
import { huToMm, mmToHu } from "./hwpPicture";

export type Placement = {
  /** 글자처럼 취급 */
  inline: boolean;
  /** 본문과의 배치: TopAndBottom(자리 차지) / Square(어울림) / InFrontOfText(글 앞으로) / BehindText(글 뒤로) (+ Tight, Through 는 그대로 보존) */
  wrap: string;
  vertRelTo: string;
  vertAlign: string;
  vertMm: string;
  horzRelTo: string;
  horzAlign: string;
  horzMm: string;
};

export const WRAP_OPTIONS: [string, string][] = [
  ["TopAndBottom", "자리 차지"],
  ["Square", "어울림"],
  ["InFrontOfText", "글 앞으로"],
  ["BehindText", "글 뒤로"],
];
export const VERT_REL: [string, string][] = [["Paper", "종이"], ["Page", "쪽"], ["Para", "문단"]];
export const HORZ_REL: [string, string][] = [["Paper", "종이"], ["Page", "쪽"], ["Column", "단"], ["Para", "문단"]];
export const VERT_ALIGN: [string, string][] = [["Top", "위"], ["Center", "가운데"], ["Bottom", "아래"]];
export const HORZ_ALIGN: [string, string][] = [["Left", "왼쪽"], ["Center", "가운데"], ["Right", "오른쪽"]];

const str = (v: unknown, dflt: string) => (typeof v === "string" && v ? v : dflt);

/** getTableProperties / getPictureProperties 결과에서 위치 속성을 읽는다 */
export function placementFromProps(pr: Record<string, unknown>): Placement {
  return {
    inline: !!pr.treatAsChar,
    wrap: str(pr.textWrap, "TopAndBottom"),
    vertRelTo: str(pr.vertRelTo, "Para"),
    vertAlign: str(pr.vertAlign, "Top"),
    vertMm: String(huToMm(Number(pr.vertOffset ?? 0))),
    horzRelTo: str(pr.horzRelTo, "Para"),
    horzAlign: str(pr.horzAlign, "Left"),
    horzMm: String(huToMm(Number(pr.horzOffset ?? 0))),
  };
}

/** 바뀐 속성만 담아 setTableProperties / setPictureProperties 에 넘긴다 */
export function placementDiff(next: Placement, prev: Placement): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  if (next.inline !== prev.inline) out.treatAsChar = next.inline;
  if (!next.inline) {
    if (next.wrap !== prev.wrap || prev.inline) out.textWrap = next.wrap;
    if (next.vertRelTo !== prev.vertRelTo || prev.inline) out.vertRelTo = next.vertRelTo;
    if (next.horzRelTo !== prev.horzRelTo || prev.inline) out.horzRelTo = next.horzRelTo;
    if (next.vertAlign !== prev.vertAlign || prev.inline) out.vertAlign = next.vertAlign;
    if (next.horzAlign !== prev.horzAlign || prev.inline) out.horzAlign = next.horzAlign;
    if (next.vertMm !== prev.vertMm || prev.inline) out.vertOffset = mmToHu(Number(next.vertMm) || 0);
    if (next.horzMm !== prev.horzMm || prev.inline) out.horzOffset = mmToHu(Number(next.horzMm) || 0);
  }
  return out;
}

/** 글자처럼 취급을 껐을 때: 개체에 남아 있던 위치/배치 값을 그대로 쓰고, 알 수 없는 값만 기본값(자리 차지)으로 */
export function floatingDefaults(pl: Placement): Placement {
  const known = (v: string, list: [string, string][]) => list.some(([k]) => k === v);
  return {
    ...pl,
    inline: false,
    wrap: pl.wrap === "TopAndBottom" || pl.wrap === "Square" || pl.wrap === "InFrontOfText" || pl.wrap === "BehindText" ? pl.wrap : "TopAndBottom",
    vertRelTo: known(pl.vertRelTo, VERT_REL) ? pl.vertRelTo : "Para",
    horzRelTo: known(pl.horzRelTo, HORZ_REL) ? pl.horzRelTo : "Para",
    vertAlign: known(pl.vertAlign, VERT_ALIGN) ? pl.vertAlign : "Top",
    horzAlign: known(pl.horzAlign, HORZ_ALIGN) ? pl.horzAlign : "Left",
  };
}
