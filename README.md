# 일일결산 자동화 시스템

안과 검사실의 일일 통계를 자동으로 수집하고 PDF 보고서를 생성하는 Python GUI 프로그램입니다.

## 주요 기능

- **날짜 선택**: 오늘/어제 버튼으로 과거 날짜 결산 가능
- **자동 파일 스캔**: 6개 검사 장비의 네트워크 경로에서 파일/폴더 자동 스캔
- **실시간/정리 후 스캔**: 날짜 폴더 정리 전/후 모두 정상 인식
- **차트번호 추출**: 정규식 패턴 매칭으로 파일/폴더명에서 차트번호 자동 추출
- **중복 제거**: 같은 환자의 양안 검사(R/L) 및 다중 촬영을 1명으로 카운트
- **특수 계산**: 녹내장(HFA∩OCT), 안저(Fundus+Secondary), OCT+OCTS 합산
- **예약 파일 처리**: 엑셀 예약 파일에서 수술 키워드 검색 및 집계
- **엑셀/PDF 생성**: 템플릿 기반 엑셀 자동 작성 및 PDF 변환
- **사용자 친화적 GUI**: tkinter 기반의 직관적인 인터페이스
- **캐시 시스템**: 대용량 폴더 빠른 스캔 (선택사항)

## 시스템 요구사항

### 필수 요구사항
- Python 3.8 이상
- Windows OS (PDF 변환 기능 사용 시)
- Microsoft Excel 설치 (PDF 변환 기능 사용 시)
- 네트워크 경로 접근 권한

### Python 패키지
```
openpyxl>=3.0.9
pandas>=1.3.0
pywin32>=301
```

## 설치 방법

### 1. 저장소 클론
```bash
git clone https://github.com/daehyuk-kim/daily-report-automation.git
cd daily-report-automation
```

### 2. 필요 패키지 설치
```bash
pip install -r requirements.txt
```

Windows에서 pywin32 설치 후 추가 작업:
```bash
python Scripts/pywin32_postinstall.py -install
```

### 3. (선택) 캐시 시스템 설치
대용량 폴더 스캔 속도 향상을 원하면:
```bash
pip install diskcache
```

## 설정

### config.json 설정

프로그램 실행 전에 `config.json` 파일을 환경에 맞게 수정해야 합니다.

#### 1. 템플릿 및 출력 경로
```json
{
  "template_file": "F:\\결산\\일일결산_템플릿.xlsx",
  "output_pdf": "F:\\결산\\PDF\\일일결산_{date}.pdf",
  "target_sheet": "복구됨_Sheet2"
}
```

#### 2. 장비별 네트워크 경로 및 패턴

```json
{
  "equipment": {
    "SP": {
      "name": "내피",
      "path": "\\\\192.168.0.120\\sp",
      "pattern": "(\\d+)_[LR]_",
      "scan_type": "file",
      "folder_structure": "YYYY\\MM\\MM.DD"
    },
    "TOPO": {
      "name": "Tomey",
      "path": "\\\\topo\\Topo",
      "pattern": "\\s(\\d+)-\\d+",
      "scan_type": "file",
      "folder_structure": "YYYY\\MM\\TOPO MM.DD"
    },
    "ORB": {
      "name": "ORB",
      "path": "\\\\orb\\orb",
      "pattern": "\\s(\\d+)\\s+o[ds]-",
      "scan_type": "file",
      "folder_structure": "YYYY\\YYYY.MM\\ORB MM.DD"
    },
    "OCT": {
      "name": "OCT",
      "path": "\\\\192.168.0.221\\pdf",
      "pattern": "\\s(\\d+)-\\d+$",
      "scan_type": "both",
      "folder_structure": "YYYY\\MM\\oct MM.DD"
    },
    "HFA": {
      "name": "시야",
      "path": "\\\\geomsa-main2\\hfa",
      "pattern": "^(\\d+)_",
      "scan_type": "both",
      "folder_structure": "YYYY\\MM\\MM.DD"
    },
    "OQAS": {
      "name": "백내장",
      "path": "\\\\oqas\\oqass",
      "pattern": "\\s(\\d+)\\s+o[ds]-",
      "scan_type": "file",
      "folder_structure": "YYYY\\MM\\MM.DD"
    }
  }
}
```

**scan_type 설명:**
- `"file"`: 파일명에서 차트번호 추출
- `"both"`: 파일명 + 폴더명에서 차트번호 추출

**folder_structure 설명:**
- `YYYY`: 연도 (2025)
- `MM`: 월 (11)
- `DD`: 일 (18)
- `MM.DD`: 월.일 (11.18)

#### 3. 특수 항목 - 녹내장 (HFA ∩ OCT)

```json
{
  "special_items": {
    "녹내장": {
      "type": "intersection",
      "sources": ["HFA", "OCT"]
    }
  }
}
```

#### 4. 특수 항목 - 안저 (Fundus + Secondary)

```json
{
  "special_items": {
    "안저": {
      "folders": {
        "fundus": {
          "path": "\\\\AFC-210-PC\\Fundus2",
          "pattern": "_(\\d+)_\\d{3}\\.",
          "folder_structure": "YYYY\\MM\\MM.DD"
        },
        "secondary": {
          "path": "\\\\192.168.0.213\\images\\Secondary",
          "pattern": "^(\\d+)-\\d{8}@",
          "use_creation_time": false
        }
      }
    }
  }
}
```

**안저 스캔 방식:**
- **Fundus**: 날짜 폴더 또는 최상위 폴더에서 파일 스캔
- **Secondary**: 파일명에 날짜 포함 (`20251118` 형식 필터링)
- 중복 제거 후 합산

#### 5. OCT + OCTS 합산

```json
{
  "combined_items": {
    "OCT_TOTAL": {
      "auto_source": "OCT",
      "manual_source": "OCTS",
      "target_cell": {"row": 12, "col": 3}
    }
  }
}
```

**동작:**
- OCT: 자동 스캔
- OCTS: 수기 입력
- 최종: OCT + OCTS 합산하여 셀에 입력

#### 6. 수기 입력 항목

```json
{
  "manual_input": {
    "라식": {"row": 10, "col": 3},
    "FAG": {"row": 18, "col": 3},
    "안경검사": {"row": 17, "col": 3},
    "OCTS": {"row": null, "col": null, "add_to": "OCT"}
  }
}
```

#### 7. 예약 파일 키워드

```json
{
  "reservation": {
    "verion_keywords": ["toric"],
    "lensx_keywords": ["lens x", "lensx", "vivity"],
    "ex500_keywords": ["라식", "라섹", "lasik", "lasek", "올레이저"]
  }
}
```

#### 8. 직원 명단

```json
{
  "staff_list": [
    "김창범", "김수지", "김대혁", "김민지",
    "박현진", "고은서", "김현호", "임량현",
    "유예나", "김태호"
  ]
}
```

## 사용 방법

### 프로그램 실행

```bash
python daily_report_fast.py
```

또는 Windows에서 `daily_report_fast.py` 파일을 더블클릭하여 실행할 수 있습니다.

### 사용 절차

#### 1. 결산 날짜 선택
   - 기본값: 오늘 날짜
   - "오늘" 버튼: 오늘 날짜로 설정
   - "어제" 버튼: 어제 날짜로 설정
   - 직접 입력: `YYYY-MM-DD` 형식 (예: 2025-11-15)

#### 2. 근무 인원 선택
   - 오늘 근무한 직원을 체크박스로 선택
   - 기본값: 모든 직원 선택

#### 3. 수기 입력
   - **OCTS**: OCT 검사 중 수기로 기록된 건수
   - **라식**: 라식/라섹 수술 건수
   - **FAG**: 형광안저촬영 건수
   - **안경검사**: 안경 처방 건수
   - 해당 항목이 없으면 0으로 남겨둡니다

#### 4. 예약 파일 선택 (선택사항)
   - "파일 선택..." 버튼 클릭
   - 예약 엑셀 파일 선택 (.xlsx, .xls)
   - 여러 파일 동시 선택 가능
   - 예약 파일이 없으면 건너뛰기

#### 5. 결산 실행
   - "🚀 결산 실행" 버튼 클릭
   - 우측 로그 영역에서 실시간 진행 상황 확인
   - 완료되면 PDF 파일 자동 열림

### 실행 결과

프로그램이 성공적으로 실행되면:

1. 임시 엑셀 파일 생성 (`일일결산_YYYYMMDD_temp.xlsx`)
2. PDF 변환 (`F:\결산\PDF\일일결산_YYYYMMDD.pdf`)
3. PDF 파일 자동으로 열림
4. 임시 엑셀 파일 자동 삭제

## 데이터 수집 방식

### 장비별 폴더 구조

#### 그룹 A: 실시간 저장 → 저녁 정리
**SP (내피), HFA (시야), Fundus (안저)**

**낮 동안:**
```
\\192.168.0.120\sp\
  ├─ 권영숙 206338_L_Center.jpg
  ├─ 김철수 205123_R_Center.jpg
  └─ ...
```

**저녁 정리 후:**
```
\\192.168.0.120\sp\2025\11\11.18\
  ├─ 권영숙 206338_L_Center.jpg
  ├─ 김철수 205123_R_Center.jpg
  └─ ...
```

**프로그램 동작:**
- 날짜 폴더 있음 → 날짜 폴더 스캔
- 날짜 폴더 없음 → 최상위 폴더 스캔 (정리 전)

---

#### 그룹 B: 자동 날짜 폴더 저장
**TOPO, ORB, OCT, OQAS**

장비가 처음부터 날짜 폴더에 저장:
```
\\topo\Topo\2025\11\TOPO 11.18\
  ├─ 구민선 165774-16.bmp
  └─ ...
```

**프로그램 동작:**
- 항상 날짜 폴더 스캔

---

### 특수 계산 항목

#### 1. 녹내장 (HFA ∩ OCT)
```
HFA 검사 환자: {10643, 20356, 30123, ...}
OCT 검사 환자: {20356, 30123, 40567, ...}
녹내장 (교집합): {20356, 30123}
```

#### 2. 안저 (Fundus + Secondary)
```
Fundus:    {148022, 204775, ...}
Secondary: {148022, 109891, ...}
안저 (합집합): {148022, 204775, 109891, ...}
```
- 한 환자당 여러 장 촬영 가능
- 중복 제거하여 환자 수만 집계

#### 3. OCT + OCTS
```
OCT (자동):    35건
OCTS (수기):   5건
OCT 합계:     40건
```

### 예약 파일 처리

예약 엑셀에서 키워드 검색:

| 수술 유형 | 키워드 |
|----------|--------|
| Verion (Toric) | toric |
| Lensx | lens x, lensx, vivity |
| EX500/FS200 | 라식, 라섹, lasik, lasek, 올레이저 |

## 문제 해결

### 1. 네트워크 경로 접근 불가
```
⚠️ 경로 없음: \\192.168.0.120\sp
```

**해결:**
- Windows 탐색기에서 경로 접근 가능한지 확인
- 네트워크 드라이브 연결 확인
- VPN 연결 상태 확인

### 2. 날짜 폴더 없음 (정상)
```
📂 스캔 경로: \\geomsa-main2\hfa (날짜 폴더 미정리)
🔍 최상위 폴더 스캔 (정리 전)
```

**설명:**
- 낮 동안은 날짜 폴더가 없는 것이 정상
- 최상위 폴더에서 정상 스캔됨
- 저녁 정리 후에는 날짜 폴더 스캔

### 3. 안저 건수 확인
```
📂 Fundus 스캔: \\AFC-210-PC\Fundus2
  ✅ 최상위 파일 매칭: 25건
📂 Secondary 스캔: \\192.168.0.213\images\Secondary
  오늘 날짜 파일: 120개
  ✅ Secondary: 30명 (중복 제거)
📊 안저 최종 집계: 45명 (중복 제거 완료)
```

**설명:**
- 파일 개수 ≠ 환자 수
- 한 환자당 여러 장 촬영
- 중복 제거하여 환자 수만 집계

### 4. PDF 변환 실패
```
❌ PDF 변환 실패: Excel을 열 수 없습니다
```

**해결:**
1. Microsoft Excel 설치 확인
2. pywin32 재설치:
   ```bash
   pip install --upgrade pywin32
   python Scripts/pywin32_postinstall.py -install
   ```
3. Excel을 관리자 권한으로 한 번 실행
4. 실패해도 엑셀 파일은 생성되므로 수동 변환 가능

### 5. 특정 장비 0건
```
TOPO (Tomey)
  📂 스캔 경로: \\topo\Topo\2025\11\TOPO 11.18
  ⚠️ 날짜 폴더 없음 (휴무일일 수 있음)
  📊 매칭: 0건
```

**확인:**
- 휴무일이면 정상
- 평일인데 0건이면 경로/패턴 확인

## 파일 구조

```
daily-report-automation/
├── daily_report_fast.py    # 메인 프로그램
├── file_cache_manager.py   # 캐시 시스템 (선택)
├── config.json              # 설정 파일
├── requirements.txt         # Python 패키지
├── .gitignore              # Git 무시 규칙
└── README.md               # 사용 설명서
```

## 성능 최적화

### 캐시 시스템 (선택사항)
대용량 폴더 스캔 속도 향상:

```bash
pip install diskcache
```

**동작:**
- 이전 스캔 파일 캐싱
- 새 파일만 확인
- 최대 10배 속도 향상

**비활성화:**
- diskcache 미설치 시 자동 비활성화
- 정상 작동에는 영향 없음

## 주의사항

1. **네트워크 경로 표기**:
   - Windows 경로: `D:\\결산\\`
   - 네트워크 경로: `\\\\192.168.0.120\\sp`

2. **파일 정리 시점**:
   - SP, HFA, Fundus는 저녁 결산 후 날짜 폴더로 정리
   - 정리 전/후 모두 정상 인식됨

3. **차트번호 범위**:
   - 유효 범위: 1 ~ 210,000
   - 범위 외 숫자는 무시

4. **안저 집계 방식**:
   - 파일 개수가 아닌 환자 수 집계
   - 중복 자동 제거

5. **OCT + OCTS**:
   - OCT: 자동 스캔
   - OCTS: 수기 입력
   - 자동 합산하여 표시

## 업데이트 내역

### v1.1.0 (2025-11-18)
- 날짜 선택 기능 추가 (오늘/어제 버튼)
- OCT + OCTS 합산 기능 추가
- SP/HFA/Fundus 실시간 스캔 로직 개선
- 안저 중복 제거 로그 상세화
- 녹내장 계산 안정성 개선
- 저장소 정리 (디버그 파일 제거)

### v1.0.0 (2025-11-14)
- 최초 버전 릴리스
- 6개 장비 자동 스캔 기능
- 특수 항목 계산 (녹내장, 안저)
- 예약 파일 처리 기능
- 엑셀/PDF 자동 생성
- GUI 인터페이스

## 라이선스

이 프로그램은 안과 검사실 내부 사용을 위해 개발되었습니다.

## 기여

개선 사항이나 버그 발견 시 GitHub Issues로 제보해주세요.

## 문의

프로그램 사용 중 문제가 발생하거나 기능 개선이 필요한 경우:
- GitHub Issues: https://github.com/daehyuk-kim/daily-report-automation/issues
- 담당자에게 문의
