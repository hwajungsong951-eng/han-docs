/** 글 한 문단 안에서 Ctrl+←/→ 단어 이동 위치를 구한다. (한글처럼 공백으로 나뉜 덩어리가 한 단어) */

const isSpace = (ch: string) => /\s/.test(ch);

/** Ctrl+→ : 현재 단어의 끝을 지나 다음 단어의 첫 글자로. 문단 끝이면 문단 끝(-1을 돌려주면 다음 문단으로) */
export function nextWordOffset(text: string, off: number): number {
  const n = text.length;
  if (off >= n) return -1;
  let i = off;
  while (i < n && !isSpace(text[i])) i++; // 지금 단어 건너뛰기
  while (i < n && isSpace(text[i])) i++;  // 공백 건너뛰기
  return i;
}

/** Ctrl+← : 지금 단어의 처음(이미 처음이면 앞 단어의 처음)으로. 문단 맨 앞이면 -1 */
export function prevWordOffset(text: string, off: number): number {
  if (off <= 0) return -1;
  let i = Math.min(off, text.length);
  while (i > 0 && isSpace(text[i - 1])) i--;
  while (i > 0 && !isSpace(text[i - 1])) i--;
  return i;
}
