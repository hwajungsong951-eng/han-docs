import init from "./rhwp.js";

let wasmReady: Promise<unknown> | null = null;

/** rhwp WASM을 한 번만 초기화한다. 실패하면 다음 호출에서 다시 시도한다. */
export const ensureWasm = () => {
  if (!wasmReady) {
    wasmReady = init().catch((err) => {
      wasmReady = null;
      throw err;
    });
  }
  return wasmReady;
};
