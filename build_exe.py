#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EXE 파일 생성 스크립트 (PyInstaller 사용)

실행 전 설치:
    pip install pyinstaller

실행:
    python build_exe.py
"""

import subprocess
import sys
import os

print("=" * 70)
print("일일결산 자동화 시스템 - EXE 빌드")
print("=" * 70)

# PyInstaller 설치 확인
try:
    import PyInstaller
    print("✅ PyInstaller 설치됨")
except ImportError:
    print("❌ PyInstaller가 설치되지 않았습니다.")
    print("\n설치 중...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    print("✅ PyInstaller 설치 완료")

# 빌드 옵션
options = [
    "daily_report_fast.py",
    "--name=일일결산자동화",
    "--onefile",
    "--windowed",
    "--icon=NONE",
    "--hidden-import=openpyxl",
    "--hidden-import=pandas",
    "--hidden-import=xlrd",
    "--hidden-import=requests",
    "--hidden-import=win32com.client",
    "--hidden-import=pythoncom",
    "--hidden-import=file_cache_manager",
    "--clean",
]

print("\n빌드 시작...")
print(f"명령: pyinstaller {' '.join(options)}")
print()

try:
    subprocess.check_call(["pyinstaller"] + options)

    # 배포 폴더 구성: dist에 config.json 복사
    import shutil
    dist_dir = os.path.abspath("dist")
    config_src = os.path.join(os.path.dirname(__file__), "config.json")
    config_dst = os.path.join(dist_dir, "config.json")
    if os.path.exists(config_src):
        shutil.copy2(config_src, config_dst)
        print(f"✅ config.json → dist/ 복사 완료")

    print("\n" + "=" * 70)
    print("✅ EXE 파일 생성 완료!")
    print("=" * 70)
    print(f"\n위치: {os.path.join(dist_dir, '일일결산자동화.exe')}")
    print("\n배포 방법:")
    print("  dist 폴더 전체를 복사하면 됩니다.")
    print("  (일일결산자동화.exe + config.json)")
except subprocess.CalledProcessError as e:
    print("\n❌ 빌드 실패")
    print(f"오류: {e}")
    sys.exit(1)
