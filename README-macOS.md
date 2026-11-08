# 한 Docs — macOS(M1/M2/M3/M4) 설치 안내

이 ZIP은 **macOS에서 직접 빌드하기 위한 소스 키트**입니다.
Windows에서는 macOS용 DMG를 만들 수 없으므로, 미리 만들어진 DMG는 포함되어 있지 않습니다.

## 가장 빠른 방법 (M1)

1. ZIP을 풉니다.
2. 터미널을 열고 `han-docs` 폴더로 이동합니다.
3. 아래 한 줄을 실행합니다.

```bash
chmod +x build-macos.command
./build-macos.command
```

스크립트가 가능한 범위에서 자동으로 확인/설치합니다.

- Xcode Command Line Tools가 없으면 설치 창을 띄웁니다.
  설치가 끝나면 스크립트를 다시 실행하세요.
- Homebrew가 있으면 Node.js를 자동 설치합니다.
- Rust가 없으면 rustup으로 자동 설치합니다.
- 그 다음 `npm ci`와 DMG 빌드를 자동 실행합니다.

완료되면 DMG 폴더가 자동으로 열립니다.

```text
src-tauri/target/release/bundle/dmg/
```

## 수동 빌드

```bash
xcode-select --install   # 처음 한 번
cd han-docs
npm ci

# Rust가 없다면
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y
source "$HOME/.cargo/env"

npm run tauri build -- --bundles dmg
```

Intel Mac까지 함께 배포하려면:

```bash
rustup target add aarch64-apple-darwin x86_64-apple-darwin
npm run tauri build -- --target universal-apple-darwin --bundles dmg
```

## 설치

1. 생성된 `.dmg`를 엽니다.
2. `한 Docs` 앱을 `Applications` 폴더로 드래그합니다.
3. Applications에서 실행합니다.

서명·공증이 없어 macOS가 막으면:

- Finder에서 앱 **오른쪽 클릭 → 열기 → 열기**
- 또는 `시스템 설정 → 개인정보 보호 및 보안 → 그래도 열기`
- 필요 시:

```bash
xattr -dr com.apple.quarantine "/Applications/한 Docs.app"
```

## 로컬에 아무것도 설치하지 않고 DMG 만들기 (GitHub Actions)

이 폴더를 GitHub 저장소에 올린 뒤:

1. `mac-build-workflow.yml`을 `.github/workflows/build-macos.yml`로 복사합니다.
2. GitHub 저장소의 **Actions → Build macOS DMG → Run workflow**를 실행합니다.
3. 완료 후 Artifacts에서 `HanDocs-macOS-dmg`를 내려받습니다.
4. DMG를 열고 Applications로 드래그합니다.

이 방법은 GitHub의 macOS 러너에서 빌드하므로, 내 맥에는 Node/Rust/Xcode를 설치하지 않아도 됩니다.
(GitHub 저장소와 Actions 사용 권한은 필요합니다.)

## 참고

- `npm ci`는 `package-lock.json` 기준으로 의존성을 설치합니다.
- 첫 빌드는 Rust 의존성 컴파일 때문에 몇 분 걸릴 수 있습니다.
- 배포용으로 쓰려면 Apple Developer ID 서명과 공증(notarization)이 필요합니다.
- `tauri.conf.json`의 identifier가 `.app`으로 끝난다는 경고가 나올 수 있지만 빌드는 진행됩니다.
