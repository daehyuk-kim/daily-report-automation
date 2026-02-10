@echo off
chcp 65001 >nul
title 일일결산 자동화 시스템

echo ============================================
echo   일일결산 자동화 시스템 시작
echo ============================================
echo.

REM Python 경로 자동 감지
where python >nul 2>&1
if %errorlevel% equ 0 (
    python daily_report_fast.py
) else (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo.
    echo Python 3.9 이상을 설치해주세요:
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)
