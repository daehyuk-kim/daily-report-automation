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
    "--add-data=config.json;.",
    "--hidden-import=openpyxl",
    "--hidden-import=pandas",
    "--hidden-import=xlrd",
    "--hidden-import=requests",
    "--hidden-import=win32com.client",
    "--hidden-import=pythoncom",
    "--hidden-import=pyodbc",
    "--hidden-import=file_cache_manager",
    "--hidden-import=certifi",
    "--collect-data=certifi",
    "--clean",
]

print("\n빌드 시작...")
print(f"명령: pyinstaller {' '.join(options)}")
print()

try:
    subprocess.check_call(["pyinstaller"] + options)

    dist_dir = os.path.abspath("dist")

    print("\n" + "=" * 70)
    print("✅ EXE 파일 생성 완료! (config.json 내장)")
    print("=" * 70)
    print(f"\n위치: {os.path.join(dist_dir, '일일결산자동화.exe')}")
    print("\n배포 방법:")
    print("  일일결산자동화.exe 하나만 복사하면 됩니다.")
    print("  (설정 변경 필요 시 exe 옆에 config.json을 두면 우선 적용)")
except subprocess.CalledProcessError as e:
    print("\n❌ 빌드 실패")
    print(f"오류: {e}")
    sys.exit(1)
