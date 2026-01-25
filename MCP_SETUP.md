# Daily Report MCP Server 설치 가이드

## ⚠️ 필요 조건

- **Python 3.10 이상** (현재: Python 3.9)
- FastMCP 라이브러리

## 🚀 설치 방법

### 1. Python 업그레이드 (선택)

```bash
# Homebrew로 Python 3.10+ 설치
brew install python@3.11

# 가상환경 생성
python3.11 -m venv venv-mcp
source venv-mcp/bin/activate

# FastMCP 설치
pip install fastmcp
```

### 2. MCP 서버 실행

```bash
cd /Users/muffinmac/Desktop/Seraneye-Projects/daily-report
python daily_report_mcp.py
```

### 3. Claude Code에 MCP 서버 등록

```bash
# Claude Code 설정에 추가
claude mcp add daily-report /path/to/daily_report_mcp.py
```

## 📋 현재 상태

✅ **Skill 파일들은 바로 사용 가능**:
- `.claude/skills/seraneye-assistant.md` (통합)
- `.claude/skills/hospital-schedule.md`
- `.claude/skills/daily-report.md`
- `.claude/skills/lens-manager.md`

❌ **MCP Server는 Python 3.10+ 필요**:
- `daily_report_mcp.py` (코드는 완성, 실행만 안됨)

## 🎯 Skill 사용 방법

현재 바로 사용 가능:

```
"근무표 수정해줘"
→ hospital-schedule.md Skill 자동 활성화

"일일결산 오류 수정"
→ daily-report.md Skill 자동 활성화

"렌즈 추가해줘"
→ lens-manager.md Skill 자동 활성화
```

## 💡 추천

**당장 사용**: Skill 파일들 (Python 버전 무관)
**나중에**: Python 업그레이드 후 MCP 서버 활성화
