#!/bin/bash
# 한 Docs — macOS(M1/M2/M3/M4) 원클릭 빌드 스크립트
set -e
cd "$(dirname "$0")"

echo "======================================"
echo " 한 Docs macOS 빌드"
echo "======================================"

if ! xcode-select -p >/dev/null 2>&1; then
  echo "Xcode Command Line Tools가 필요합니다."
  echo "지금 설치 창을 띄웁니다. 설치가 끝난 뒤 이 파일을 다시 실행하세요."
  xcode-select --install || true
  read -n 1 -s -r -p "Enter 키를 누르면 종료합니다..."
  echo
  exit 1
fi

if ! command -v node >/dev/null 2>&1; then
  if command -v brew >/dev/null 2>&1; then
    echo "Node.js를 Homebrew로 설치합니다..."
    brew install node
  else
    echo "Node.js 22.12 이상이 필요합니다."
    echo "https://nodejs.org 에서 설치한 뒤 다시 실행하세요."
    read -n 1 -s -r -p "Enter 키를 누르면 종료합니다..."
    echo
    exit 1
  fi
fi

if ! command -v rustc >/dev/null 2>&1; then
  echo "Rust를 설치합니다..."
  curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y
  # shellcheck disable=SC1091
  source "$HOME/.cargo/env"
fi

echo "의존성 설치 (npm ci)..."
npm ci

echo "DMG 빌드 중... (첫 빌드는 몇 분 걸릴 수 있습니다)"
npm run tauri build -- --bundles dmg

echo
echo "완료되었습니다."
echo "DMG 폴더: src-tauri/target/release/bundle/dmg/"
open "src-tauri/target/release/bundle/dmg" 2>/dev/null || true
read -n 1 -s -r -p "Enter 키를 누르면 종료합니다..."
echo
