# 일일결산 자동화 실행 방법

## ✅ 최신 버전 실행 (v1.2.0 - Clean White UI)

```bash
python3 daily_report_upgraded.py
```

## 📁 파일 설명

### ✨ daily_report_upgraded.py (26K) - **최신 버전 (사용 권장)**
- ✅ Clean White UI (민트-그린 미니멀 디자인)
- ✅ 실시간 폴더 모니터링 (watchdog)
- ✅ Hospital Schedule API 연동 (검사실 직원 자동 로드)
- ✅ 2단 레이아웃 (입력 | 결과 그리드)
- ✅ 카드형 결과 (3x5 그리드, 큰 입력칸)

### 📦 daily_report_fast.py (77K) - **구버전 (사용 안함)**
- ❌ 구버전 UI
- ❌ 실시간 모니터링 없음
- ❌ API 연동 없음

### 📦 daily_report_clean_white.py (21K) - **개발 중간 버전**
- 개발 과정 파일

### 🔧 daily_report_mcp.py (6.4K) - **MCP Server (Python 3.10+ 필요)**
- Claude Code용 MCP 서버
- 현재 Python 3.9로 실행 불가

## 🚀 빠른 시작

### 1. 필수 라이브러리 설치

```bash
pip3 install watchdog requests openpyxl pandas
```

### 2. 최신 버전 실행

```bash
cd /Users/muffinmac/Desktop/Seraneye-Projects/daily-report
python3 daily_report_upgraded.py
```

### 3. Mac에서 실행 시

Mac에서 tkinter 창이 안 보이면:

```bash
pythonw daily_report_upgraded.py
```

또는 Python 3.11 설치:

```bash
brew install python@3.11
python3.11 daily_report_upgraded.py
```

## ⚠️ 주의사항

- **daily_report_fast.py는 구버전**입니다. 실행하지 마세요.
- **반드시 daily_report_upgraded.py를 실행**하세요.
- 실행 전 config.json 확인 (장비 경로, 템플릿 경로)

## 📋 새 기능 사용법

### 실시간 모니터링
- 체크박스로 ON/OFF
- 파일 생성 감지 시 자동 스캔 (2초 딜레이)

### Hospital Schedule API
- 날짜 변경 시 자동 업데이트
- Enter 키로 수동 새로고침
- 검사실 근무 직원 자동 로드

### Clean White UI
- 왼쪽: 입력 패널 (300px)
- 오른쪽: 결과 그리드 (3x5 카드)
- 민트-그린 색상 테마 (#11998e)

## 🔧 트러블슈팅

### watchdog 없음
```bash
pip3 install watchdog
```

### Hospital Schedule API 연결 안됨
```bash
# IP 확인
ping 192.168.0.210

# 포트 확인
nc -zv 192.168.0.210 3001
```

### 창이 안 보임 (Mac)
```bash
pythonw daily_report_upgraded.py
```

---

*최종 업데이트: 2025-01-26*
