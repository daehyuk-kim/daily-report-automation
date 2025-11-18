# 일일결산 프로그램 문제 분석 및 해결방안

## 문제 1: HFA (시야) 인식 안됨 (0건으로 표시)

### 원인 분석:

HFA는 `scan_type: "both"`로 설정되어 있어 **폴더명**을 스캔해야 합니다.

**폴더 구조:**
```
\\geomsa-main2\hfa\2025\11\11.15\
  ├─ 10643_JOO_DONG MIN_19320821\
  ├─ 20356_KIM_YOUNG_HEE_19450103\
  └─ ...
```

**패턴:** `^(\d+)_`
**예상 동작:** 폴더명 `10643_JOO_DONG MIN_19320821`에서 `10643` 추출

### 가능한 문제:

1. **날짜 폴더가 없을 때의 fallback 로직 문제**
   - `get_today_folder_path()`가 None을 반환하면 base_path를 직접 스캔
   - 하지만 HFA의 경우 환자 폴더명에 날짜 정보가 없음
   - 따라서 날짜 필터링이 작동하지 않아 0건 반환

2. **scan_type: "both" 로직 확인 필요**
   - `os.walk()` 사용 시 폴더명 매칭이 정상적으로 작동하는지
   - `dirs` 리스트에서 패턴 매칭이 되는지

### 해결 방안:

#### 방안 1: 날짜 폴더 필수로 변경 (권장)
```python
def get_today_folder_path(self, base_path: str, equipment_id: str, today=None) -> str:
    if today is None:
        today = self.today

    equipment = self.config['equipment'][equipment_id]
    if 'folder_structure' not in equipment:
        return base_path

    folder_structure = equipment['folder_structure']

    # 날짜 변환
    folder = folder_structure
    folder = folder.replace('YYYY.MM', today.strftime('%Y.%m'))
    folder = folder.replace('YYYY', today.strftime('%Y'))
    folder = folder.replace('MM.DD', today.strftime('%m.%d'))
    folder = folder.replace('MM', today.strftime('%m'))
    folder = folder.replace('DD', today.strftime('%d'))

    full_path = os.path.join(base_path, folder)

    if os.path.exists(full_path):
        return full_path
    else:
        # 날짜 폴더 구조가 있는 경우 폴더가 없으면 None 반환
        # fallback 하지 않음!
        return None
```

그리고 `scan_directory_fast()`에서:
```python
# 오늘 날짜 폴더 경로 찾기
today_folder = self.get_today_folder_path(base_path, equipment_id)

if today_folder is None:
    if 'folder_structure' in equipment:
        # 날짜 폴더 구조가 있는데 폴더가 없으면 0건 반환
        log_callback(f"  ⚠️  날짜 폴더 없음 (휴무일일 수 있음)")
        return chart_numbers
    else:
        # 폴더 구조가 없는 경우만 base_path 스캔
        today_folder = base_path
        use_creation_time = equipment.get('use_creation_time', False)
        log_callback(f"     📂 스캔 경로: {today_folder}")
        # ... 기존 fallback 로직
```

#### 방안 2: HFA 전용 스캔 로직 추가
`calculate_hfa()` 메서드를 별도로 만들어 HFA 특화 로직 구현

---

## 문제 2: 녹내장 인식 안됨 (0건으로 표시)

### 원인:
```python
def calculate_glaucoma(self, log_callback) -> int:
    """녹내장 계산 (HFA ∩ OCT)"""
    hfa_charts = self.chart_numbers.get('HFA', set())
    oct_charts = self.chart_numbers.get('OCT', set())
    glaucoma_charts = hfa_charts & oct_charts
    return len(glaucoma_charts)
```

HFA가 0건이면 교집합도 0건입니다.

### 해결 방안:
**HFA 문제를 먼저 해결**하면 자동으로 해결됩니다.

---

## 문제 3: 안저 인식은 되는데 수가 틀림

### 현재 로직:

```python
def calculate_fundus(self, log_callback) -> int:
    fundus_charts = set()

    # 1. Fundus 폴더 스캔
    #    패턴: _(\d+)_\d{3}\.
    #    예: 구정순 _148022_000.jpg → 148022

    # 2. Secondary 폴더 스캔
    #    패턴: ^(\d+)-\d{8}@
    #    예: 204775-20251115@161455-l4-s.jpg → 204775

    # 3. 합집합 (중복 제거)
    fundus_charts.update(secondary_charts)

    return len(fundus_charts)  # 유니크 환자 수 반환
```

### 예시:
```
Secondary 폴더에서 2025-11-15:
- 109891-20251115@105433-L4-S.jpg
- 109891-20251115@105433-L5-S.jpg
- 109891-20251115@105433-L6-S.jpg
- 109891-20251115@105433-R1-S.jpg
- 109891-20251115@105433-R2-S.jpg
→ 5개 파일, 하지만 환자는 1명 (109891)

전체 311개 파일이 있어도, 유니크 환자는 50명일 수 있음
```

### 가능한 원인:

1. **사용자가 파일 개수를 기대**
   - 현재: 50건 (유니크 환자 수)
   - 기대: 311건 (전체 파일 수)

2. **중복 제거 로직 문제**
   - Fundus와 Secondary가 같은 환자를 중복으로 카운팅?

3. **패턴 매칭 문제**
   - 일부 파일이 패턴 매칭에 실패?

### 해결 방안:

#### 방안 1: 현재 로직이 맞다고 설명
```
"안저 검사는 한 환자당 여러 장의 사진을 촬영합니다.
311개의 파일이 있지만, 실제 검사받은 환자는 50명입니다.
중복을 제거한 환자 수를 집계하는 것이 정확합니다."
```

#### 방안 2: 파일 개수와 환자 수를 모두 표시
```python
fundus_file_count = len([...])  # 전체 파일 수
fundus_patient_count = len(fundus_charts)  # 유니크 환자 수

log_callback(f"  ✓ 안저: {fundus_patient_count}명 (파일 {fundus_file_count}개)")
```

---

## 디버그 방법

### 1. 상세 디버그 스크립트 실행 (Windows에서):
```bash
python debug_detailed.py 2025-11-15
```

이 스크립트는:
- HFA 폴더 구조 확인
- 파일/폴더 샘플 출력
- 패턴 매칭 상세 분석
- 안저 파일 중복 통계

### 2. 결과 분석:
- HFA에서 폴더가 매칭되는지 확인
- 안저에서 파일 수 vs 환자 수 확인

---

## 권장 수정사항

### 1. HFA 폴더 없을 때 처리 개선
`daily_report_fast.py` 의 `scan_directory_fast()` 메서드 수정

### 2. 로그 개선
```python
# HFA 로그 예시
log_callback(f"  ✓ HFA: {len(chart_set)}건 (폴더 {total_dirs_count}개)")

# 안저 로그 예시
log_callback(f"  ✓ 안저: {patient_count}명 (Fundus: {fundus_count}명, Secondary: {secondary_count}명, 중복: {overlap}명)")
```

### 3. 에러 처리 강화
날짜 폴더가 없는 경우를 에러가 아닌 정상 상황으로 처리 (휴무일)

---

## 테스트 계획

1. **2025-11-15 (금요일 - 정상 영업일)**
   - HFA 폴더 존재
   - 모든 장비 정상 스캔
   - 예상: 모든 항목에 정상 값

2. **2025-11-17 (일요일 - 휴무일)**
   - 대부분 폴더 없음
   - 예상: 안전하게 0건 처리

3. **로그 확인**
   - 각 장비별 스캔 경로
   - 매칭 건수
   - 패턴 매칭 성공/실패 사례
