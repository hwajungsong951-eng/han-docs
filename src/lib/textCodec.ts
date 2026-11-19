/** TXT 인코딩 감지/복원: UTF-8(BOM 유무), UTF-16, 한글 ANSI(CP949/EUC-KR)를 구분해 원래 형식으로 저장한다. */
export type TxtEncoding = "utf-8" | "utf-8-bom" | "utf-16le" | "utf-16be" | "cp949";
export type DecodedText = { text: string; encoding: TxtEncoding; eol: "\r\n" | "\n" };

export function decodeTxt(b: Uint8Array): DecodedText {
  let encoding: TxtEncoding = "utf-8";
  let text: string;
  if (b.length >= 3 && b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) {
    encoding = "utf-8-bom"; text = new TextDecoder("utf-8").decode(b.subarray(3));
  } else if (b.length >= 2 && b[0] === 0xff && b[1] === 0xfe) {
    encoding = "utf-16le"; text = new TextDecoder("utf-16le").decode(b.subarray(2));
  } else if (b.length >= 2 && b[0] === 0xfe && b[1] === 0xff) {
    encoding = "utf-16be"; text = new TextDecoder("utf-16be").decode(b.subarray(2));
  } else {
    try { text = new TextDecoder("utf-8", { fatal: true }).decode(b); }
    catch { encoding = "cp949"; text = new TextDecoder("euc-kr").decode(b); } // WHATWG euc-kr = CP949
  }
  const eol = text.includes("\r\n") ? "\r\n" : text.includes("\n") ? "\n" : "\r\n";
  return { text, encoding, eol };
}

let cp949Map: Map<string, number> | null = null;
function buildCp949Map(): Map<string, number> {
  const m = new Map<string, number>();
  const dec = new TextDecoder("euc-kr");
  const buf = new Uint8Array(2);
  for (let lead = 0x81; lead <= 0xfe; lead++) {
    for (let trail = 0x41; trail <= 0xfe; trail++) {
      buf[0] = lead; buf[1] = trail;
      const s = dec.decode(buf);
      if (s.length === 1 && s !== "\ufffd" && !m.has(s)) m.set(s, (lead << 8) | trail);
    }
  }
  return m;
}

export function encodeTxt(text: string, encoding: TxtEncoding): Uint8Array {
  switch (encoding) {
    case "utf-8": return new TextEncoder().encode(text);
    case "utf-8-bom": {
      const body = new TextEncoder().encode(text);
      const out = new Uint8Array(body.length + 3); out.set([0xef, 0xbb, 0xbf]); out.set(body, 3); return out;
    }
    case "utf-16le": case "utf-16be": {
      const le = encoding === "utf-16le";
      const out = new Uint8Array(2 + text.length * 2);
      out[0] = le ? 0xff : 0xfe; out[1] = le ? 0xfe : 0xff;
      for (let i = 0; i < text.length; i++) {
        const c = text.charCodeAt(i);
        out[2 + i * 2] = le ? c & 0xff : c >> 8; out[3 + i * 2] = le ? c >> 8 : c & 0xff;
      }
      return out;
    }
    case "cp949": {
      cp949Map ??= buildCp949Map();
      const bytes: number[] = [];
      for (const ch of text) {
        const cp = ch.codePointAt(0)!;
        if (cp < 0x80) { bytes.push(cp); continue; }
        const v = cp949Map.get(ch);
        if (v !== undefined) bytes.push(v >> 8, v & 0xff);
        else return encodeTxt(text, "utf-8-bom"); // ANSI 로 표현 못 하는 문자(이모지 등)가 있으면 글자 손실 대신 UTF-8 로 저장
      }
      return Uint8Array.from(bytes);
    }
  }
}
