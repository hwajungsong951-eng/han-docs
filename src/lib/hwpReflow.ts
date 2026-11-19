import type { HwpDocument } from "../wasm/rhwp.js";

/**
 * 표/그림/도형이 들어 있는 문단의 줄 위치(line seg)를 다시 계산시킨다.
 *
 * rhwp 는 표를 "글자처럼 취급" 하는 문단의 세로 위치/높이를 문단의 line seg 에 저장해 두고 그대로 쓴다.
 * 그런데 붙여넣기(pasteControl) 나 표 삽입처럼 문단을 새로 끼워 넣는 작업은 "뒤따르는" 표 문단의
 * line seg 를 갱신하지 않아서, 이미 있던 표가 옛 위치에 그대로 남아 새로 붙인 표나 글자와 겹쳐 그려진다.
 * (저장해서 다시 열어도 같은 line seg 가 저장되므로 겹침이 그대로 남는다.)
 *
 * 컨트롤이 있는 문단마다 글자 한 개를 넣었다가 지우면 그 문단이 다시 조판되어 line seg 가 새로 계산된다.
 * 호출은 edit() 안에서(= 되돌리기 스냅샷 안에서) 해야 한다.
 */
export function reflowControlParagraphs(d: HwpDocument, sec: number): void {
  const paras = new Set<number>();
  try {
    const pages = d.pageCount();
    for (let p = 0; p < pages; p++) {
      const layout = JSON.parse(d.getPageControlLayout(p)) as { controls?: { secIdx?: number; paraIdx?: number }[] };
      for (const c of layout.controls ?? []) {
        if ((c.secIdx ?? 0) === sec && typeof c.paraIdx === "number") paras.add(c.paraIdx);
      }
    }
  } catch { return; }
  for (const i of [...paras].sort((a, b) => a - b)) {
    try {
      d.insertText(sec, i, 0, "x");
      d.deleteText(sec, i, 0, 1);
    } catch { /* 이 문단은 건너뛴다 */ }
  }
}

/**
 * 같은 쪽에서 본문 흐름 속 표끼리 세로로 겹쳐 있는지(= 예전 버전에서 붙여넣기 후 저장된 파일) 검사한다.
 */
export function hasOverlappingTables(d: HwpDocument): boolean {
  try {
    const pages = d.pageCount();
    for (let p = 0; p < pages; p++) {
      const layout = JSON.parse(d.getPageControlLayout(p)) as {
        controls?: { type?: string; x: number; y: number; w: number; h: number; plane?: number; secIdx?: number; paraIdx?: number; controlIdx?: number }[];
      };
      // 본문 흐름 속 "글자처럼 취급"·"위아래 배치" 표만 비교한다 (글 앞/뒤, 어울림 배치 표는 일부러 겹칠 수 있다).
      const tables = (layout.controls ?? []).filter((c) => {
        if (c.type !== "table" || c.plane !== 2 || !(c.h > 0)) return false;
        try {
          const tp = JSON.parse(d.getTableProperties(c.secIdx ?? 0, c.paraIdx ?? 0, c.controlIdx ?? 0)) as { treatAsChar?: boolean; textWrap?: string; vertRelTo?: string };
          return tp.treatAsChar === true || (tp.textWrap === "TopAndBottom" && tp.vertRelTo === "Para");
        } catch { return false; }
      });
      for (let i = 0; i < tables.length; i++) {
        for (let j = i + 1; j < tables.length; j++) {
          const a = tables[i], b = tables[j];
          const ox = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
          const oy = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
          if (ox > 2 && oy > 2) return true;
        }
      }
    }
  } catch { /* 검사 실패는 무시 */ }
  return false;
}
