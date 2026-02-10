# 일일결산 자동화 실행 방법

## ✅ 권장 버전 실행 (기존 UI + API 연동)

```bash
python3 daily_report_fast.py
```

## 📁 파일 설명

### ✨ daily_report_fast.py (77K) - **권장 버전** ⭐
- ✅ 기존 익숙한 UI (3칸 레이아웃)
- ✅ Hospital Schedule API 연동 (검사실 직원 자동 로드)
- ✅ 🔄 새로고침 버튼 (API 상태 표시)
- ✅ 날짜 변경 시 직원 자동 업데이트
- ✅ Enter 키로 수동 새로고침
- ✅ API 실패 시 자동 fallback (config 사용)

### 📦 daily_report_upgraded.py (26K) - **새 UI 실험 버전**
- Clean White UI (민트-그린 미니멀 디자인)
- 실시간 폴더 모니터링 (watchdog)
- 2단 레이아웃 (입력 | 결과 그리드)
- ⚠️ UI가 다르므로 익숙하지 않을 수 있음

### 📦 daily_report_clean_white.py (21K) - **개발 중간 버전**
- 개발 과정 파일

### 🔧 daily_report_mcp.py (6.4K) - **MCP Server (Python 3.10+ 필요)**
- Claude Code용 MCP 서버
- 현재 Python 3.9로 실행 불가

## 🚀 빠른 시작

### 1. 필수 라이브러리 설치

```bash
pip3 install requests openpyxl pandas
```

### 2. 권장 버전 실행

```bash
cd /Users/muffinmac/Desktop/Seraneye-Projects/daily-report
python3 daily_report_fast.py
```

### 3. Mac에서 실행 시

Mac에서 tkinter 창이 안 보이면:

```bash
pythonw daily_report_fast.py
```

또는 Python 3.11 설치:

```bash
brew install python@3.11
python3.11 daily_report_fast.py
```

## ⚠️ 주의사항

- **daily_report_fast.py가 권장 버전**입니다. (기존 UI + API 연동)
- 실행 전 config.json 확인 (장비 경로, 템플릿 경로)
- Hospital Schedule API 주소: http://192.168.0.210:3001

## 📋 새 기능 사용법

### Hospital Schedule API 연동 ⭐ NEW!
- **날짜 변경 시 자동 업데이트**: "오늘", "어제" 버튼 클릭 시
- **Enter 키로 수동 새로고침**: 날짜 입력란에서 Enter
- **🔄 버튼**: 직원 목록 수동 새로고침
- **API 상태 표시**:
  - 🟢 "API 연동 (N명)" - 성공
  - 🟠 "config 사용" - API 실패, fallback
  - 🔴 "API 오류" - 연결 불가

### API 동작 방식
1. 프로그램 시작 시 오늘 날짜로 API 호출
2. 성공하면 검사실 근무 직원 자동 로드
3. 실패하면 config.json의 staff_list 사용
4. 날짜 변경할 때마다 새로 조회

## 🔧 트러블슈팅

### requests 없음
```bash
pip3 install requests
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
pythonw daily_report_fast.py
```

### API 상태가 "config 사용"으로 나옴
- Hospital Schedule 시스템이 실행 중인지 확인
- API 주소가 맞는지 확인: http://192.168.0.210:3001
- config 사용이어도 정상 작동 (fallback 모드)

---

*최종 업데이트: 2025-01-26 (API 연동 추가)*
