# 한 Docs

HWP/HWPX/DOCX/TXT 문서를 열고 편집하며 PDF로 저장할 수 있는 Windows/macOS 데스크톱 앱입니다.
Tauri + React + rhwp(WASM) 기반입니다.

## 주요 기능

- HWP / HWPX 문서 열기·편집·저장
- DOCX / TXT 문서 열기·편집·저장
- HWP/HWPX/DOCX/TXT → PDF 저장
- 글자 크기·색·글꼴·굵게/기울임/밑줄/취소선
- 마우스 드래그, Shift+방향키, Ctrl+A, 더블클릭 단어 선택
- **복사/잘라내기/붙여넣기 (Ctrl+C / Ctrl+X / Ctrl+V)**
  - HWP: 블록 선택 후 텍스트 복사, 표 안에서 선택 없이 Ctrl+C하면 표 전체 복사
  - DOCX/TXT: 일반 문서처럼 복사·붙여넣기
- A4 페이지 여백(mm) 조정 및 쪽 번호(머리말/꼬리말)
- 표·그림·도형·기호 삽입
- 한글 IME 입력

## Windows 실행

`outputs/HanDocs/HanDocs.exe`를 실행하거나 `HanDocs-setup.exe`로 설치합니다.
(릴리스 파일은 별도로 배포됩니다.)

## 개발

```bash
npm ci
npm run tauri dev
```

## 빌드

### Windows

```bash
npm run tauri build
```

### macOS

```bash
chmod +x build-macos.command
./build-macos.command
```

또는 GitHub Actions의 **Build macOS DMG** 워크플로를 실행한 뒤 Artifacts에서 DMG를 받습니다.

### GitHub Actions로 macOS DMG 만들기

현재 GitHub 토큰 권한 문제로 워크플로 파일은 저장소 루트의 `mac-build-workflow.yml`로 들어 있습니다.
GitHub 웹에서 아래 주소를 열고 새 파일을 만드세요.

```text
https://github.com/hwajungsong951-eng/han-docs/new/main/.github/workflows
```

파일 이름:

```text
build-macos.yml
```

내용은 저장소 루트의 `mac-build-workflow.yml`을 복사해서 붙여넣으면 됩니다.
그다음 **Actions → Build macOS DMG → Run workflow**를 실행하고, 완료 후 Artifacts에서 `HanDocs-macOS-dmg`를 내려받으세요.

## 라이선스

- 앱 코드: 저장소 정책에 따름
- rhwp 엔진: `LICENSE-rhwp` 참고 (https://github.com/edwardkim/rhwp)
