// 파일 입출력은 Rust 백엔드에서 처리 (바이너리 안전, 경로 제한 없음)

#[tauri::command]
fn read_file_bytes(path: String) -> Result<Vec<u8>, String> {
    std::fs::read(&path).map_err(|e| format!("파일 읽기 실패: {e}"))
}

#[tauri::command]
fn write_file_bytes(path: String, data: Vec<u8>) -> Result<(), String> {
    std::fs::write(&path, data).map_err(|e| format!("파일 저장 실패: {e}"))
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![read_file_bytes, write_file_bytes])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
