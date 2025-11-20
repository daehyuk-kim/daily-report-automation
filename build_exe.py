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
    "--onefile",  # 단일 EXE 파일
    "--windowed",  # GUI 모드 (콘솔 숨김)
    "--icon=NONE",  # 아이콘 없음 (필요시 .ico 파일 지정)
    "--add-data=config.json;.",  # config.json 포함
    "--hidden-import=openpyxl",
    "--hidden-import=pandas",
    "--hidden-import=xlrd",
    "--hidden-import=win32com.client",
    "--hidden-import=pythoncom",
    "--clean",  # 이전 빌드 정리
]

print("\n빌드 시작...")
print(f"명령: pyinstaller {' '.join(options)}")
print()

try:
    subprocess.check_call(["pyinstaller"] + options)
    print("\n" + "=" * 70)
    print("✅ EXE 파일 생성 완료!")
    print("=" * 70)
    print(f"\n위치: {os.path.abspath('dist/일일결산자동화.exe')}")
    print("\n배포 방법:")
    print("  1. dist 폴더의 일일결산자동화.exe 복사")
    print("  2. config.json을 같은 폴더에 복사")
    print("  3. 더블클릭으로 실행")
except subprocess.CalledProcessError as e:
    print("\n❌ 빌드 실패")
    print(f"오류: {e}")
    sys.exit(1)
