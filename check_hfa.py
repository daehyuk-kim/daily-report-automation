#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HFA(시야) 인식 디버깅 스크립트
왜 0으로 인식되는지 단계별로 확인
"""

import os
import json
import re
from datetime import date

print("=" * 80)
print("HFA(시야) 디버깅 스크립트")
print("=" * 80)

# 1. Config 로드
print("\n[1단계] config.json 로드")
try:
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
    hfa_config = config['equipment']['HFA']
    print(f"✅ Config 로드 성공")
    print(f"   이름: {hfa_config['name']}")
    print(f"   경로: {hfa_config['path']}")
    print(f"   패턴: {hfa_config['pattern']}")
    print(f"   스캔타입: {hfa_config['scan_type']}")
    print(f"   폴더구조: {hfa_config['folder_structure']}")
except Exception as e:
    print(f"❌ Config 로드 실패: {e}")
    exit(1)

# 2. 경로 존재 확인
print("\n[2단계] 경로 존재 확인")
base_path = hfa_config['path']
print(f"   기본 경로: {base_path}")
if os.path.exists(base_path):
    print(f"✅ 경로 존재함")
else:
    print(f"❌ 경로 없음 - 프로그램이 여기서 멈춥니다!")
    print(f"   Windows에서는 네트워크 경로가 마운트되어 있어야 합니다.")
    print(f"   예: \\\\geomsa-main2\\hfa")
    exit(1)

# 3. 오늘 날짜 폴더 경로 생성
print("\n[3단계] 오늘 날짜 폴더 경로 생성")
today = date.today()
print(f"   오늘 날짜: {today.strftime('%Y-%m-%d')}")

folder_structure = hfa_config['folder_structure']
folder = folder_structure
folder = folder.replace('YYYY.MM', today.strftime('%Y.%m'))
folder = folder.replace('YYYY', today.strftime('%Y'))
folder = folder.replace('MM.DD', today.strftime('%m.%d'))
folder = folder.replace('MM', today.strftime('%m'))
folder = folder.replace('DD', today.strftime('%d'))

today_folder = os.path.join(base_path, folder)
print(f"   예상 경로: {today_folder}")

if os.path.exists(today_folder):
    print(f"✅ 오늘 날짜 폴더 존재")
    scan_path = today_folder
else:
    print(f"⚠️  오늘 날짜 폴더 없음 - 최상위 폴더 스캔")
    scan_path = base_path

# 4. 폴더 내용 확인
print(f"\n[4단계] 폴더 내용 스캔: {scan_path}")
try:
    items = os.listdir(scan_path)
    print(f"   전체 항목: {len(items)}개")

    files = [item for item in items if os.path.isfile(os.path.join(scan_path, item))]
    dirs = [item for item in items if os.path.isdir(os.path.join(scan_path, item))]

    print(f"   파일: {len(files)}개")
    print(f"   폴더: {len(dirs)}개")

    # 처음 5개만 출력
    if dirs:
        print(f"\n   📁 폴더 샘플 (최대 5개):")
        for d in dirs[:5]:
            print(f"      - {d}")

    if files:
        print(f"\n   📄 파일 샘플 (최대 5개):")
        for f in files[:5]:
            print(f"      - {f}")

except Exception as e:
    print(f"❌ 폴더 읽기 실패: {e}")
    exit(1)

# 5. 정규식 패턴 테스트
print(f"\n[5단계] 정규식 패턴 테스트")
pattern = re.compile(hfa_config['pattern'])
print(f"   패턴: {hfa_config['pattern']}")

# 파일 확장자 필터
valid_extensions = config['validation']['file_extensions']
print(f"   유효 확장자: {valid_extensions}")

chart_numbers = set()
folder_matches = []
file_matches = []

# 폴더 매칭
print(f"\n   📁 폴더 매칭 테스트:")
for dir_name in dirs[:10]:  # 최대 10개
    match = pattern.search(dir_name)
    if match:
        # 이중 그룹 처리
        chart_num = match.group(1) or (match.group(2) if match.lastindex > 1 else None)
        folder_matches.append((dir_name, chart_num, match.groups()))

if folder_matches:
    for dir_name, chart_num, groups in folder_matches[:5]:
        print(f"      ✅ {dir_name[:50]:50} → {chart_num} (groups: {groups})")
else:
    print(f"      ⚠️  매칭된 폴더 없음")

# 파일 매칭
print(f"\n   📄 파일 매칭 테스트:")
for file_name in files[:10]:  # 최대 10개
    # 확장자 체크
    if not any(file_name.lower().endswith(ext) for ext in valid_extensions):
        continue

    match = pattern.search(file_name)
    if match:
        # 이중 그룹 처리
        chart_num = match.group(1) or (match.group(2) if match.lastindex > 1 else None)
        file_matches.append((file_name, chart_num, match.groups()))

if file_matches:
    for file_name, chart_num, groups in file_matches[:5]:
        print(f"      ✅ {file_name[:50]:50} → {chart_num} (groups: {groups})")
else:
    print(f"      ⚠️  매칭된 파일 없음")

# 6. 차트번호 유효성 검증
print(f"\n[6단계] 차트번호 유효성 검증")
min_val = config['validation']['chart_number_min']
max_val = config['validation']['chart_number_max']
allow_leading_zero = config['validation']['allow_leading_zero']

print(f"   유효 범위: {min_val} ~ {max_val}")
print(f"   앞자리 0 허용: {allow_leading_zero}")

def is_valid_chart_number(chart_num_str):
    try:
        if chart_num_str.startswith('0') and len(chart_num_str) > 1:
            return False, "앞자리 0"
        chart_num = int(chart_num_str)
        if not (min_val <= chart_num <= max_val):
            return False, f"범위 벗어남 ({chart_num})"
        return True, "유효"
    except ValueError:
        return False, "숫자 변환 실패"

# 폴더 차트번호 검증
print(f"\n   📁 폴더 차트번호 검증:")
valid_folder_charts = []
invalid_folder_charts = []

for dir_name, chart_num, _ in folder_matches:
    is_valid, reason = is_valid_chart_number(chart_num)
    if is_valid:
        valid_folder_charts.append(chart_num)
        chart_numbers.add(chart_num)
    else:
        invalid_folder_charts.append((dir_name, chart_num, reason))

print(f"      ✅ 유효: {len(valid_folder_charts)}개")
if invalid_folder_charts:
    print(f"      ❌ 무효: {len(invalid_folder_charts)}개")
    for dir_name, chart_num, reason in invalid_folder_charts[:3]:
        print(f"         - {dir_name[:40]:40} → {chart_num} ({reason})")

# 파일 차트번호 검증
print(f"\n   📄 파일 차트번호 검증:")
valid_file_charts = []
invalid_file_charts = []

for file_name, chart_num, _ in file_matches:
    is_valid, reason = is_valid_chart_number(chart_num)
    if is_valid:
        valid_file_charts.append(chart_num)
        chart_numbers.add(chart_num)
    else:
        invalid_file_charts.append((file_name, chart_num, reason))

print(f"      ✅ 유효: {len(valid_file_charts)}개")
if invalid_file_charts:
    print(f"      ❌ 무효: {len(invalid_file_charts)}개")
    for file_name, chart_num, reason in invalid_file_charts[:3]:
        print(f"         - {file_name[:40]:40} → {chart_num} ({reason})")

# 7. 최종 결과
print(f"\n" + "=" * 80)
print(f"최종 결과")
print("=" * 80)
print(f"총 차트번호 개수: {len(chart_numbers)}개")

if len(chart_numbers) == 0:
    print(f"\n❌ 시야 0건 - 원인 분석:")

    if not dirs and not files:
        print(f"   1. 폴더가 비어있음")
    elif len(folder_matches) == 0 and len(file_matches) == 0:
        print(f"   2. 정규식 패턴이 매칭되지 않음")
        print(f"      - 실제 폴더명/파일명 형식을 확인하세요")
    elif len(valid_folder_charts) == 0 and len(valid_file_charts) == 0:
        print(f"   3. 모든 차트번호가 유효성 검증 실패")
        print(f"      - 범위 확인: {min_val} ~ {max_val}")
        if invalid_folder_charts or invalid_file_charts:
            print(f"      - 무효 사유: {invalid_folder_charts[0][2] if invalid_folder_charts else invalid_file_charts[0][2]}")
else:
    print(f"✅ 시야 {len(chart_numbers)}건 인식 성공")
    print(f"\n   차트번호 샘플 (최대 10개):")
    for chart_num in sorted(chart_numbers)[:10]:
        print(f"      - {chart_num}")

print("\n" + "=" * 80)
