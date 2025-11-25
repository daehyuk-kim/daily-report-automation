#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
실제 차트번호 범위 분석

최근 데이터를 보고 실제 차트번호가 몇 범위인지 확인
"""

import os
import json
import re
from collections import defaultdict

with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

print("=" * 80)
print("실제 차트번호 범위 분석")
print("=" * 80)

# 각 장비별로 분석
equipment_stats = {}

for eq_id, eq_info in config['equipment'].items():
    print(f"\n[{eq_info['name']} - {eq_id}]")
    print(f"경로: {eq_info['path']}")
    print(f"패턴: {eq_info['pattern']}")

    base_path = eq_info['path']
    pattern = re.compile(eq_info['pattern'])

    if not os.path.exists(base_path):
        print(f"⚠️  경로 없음 (Windows에서 실행 필요)")
        continue

    # 샘플 수집 (최대 100개)
    samples = []
    try:
        for root, dirs, files in os.walk(base_path):
            # 폴더명 확인
            for dir_name in dirs[:50]:
                match = pattern.search(dir_name)
                if match:
                    for group in match.groups():
                        if group and group.isdigit():
                            samples.append(int(group))

            # 파일명 확인
            for file_name in files[:50]:
                match = pattern.search(file_name)
                if match:
                    for group in match.groups():
                        if group and group.isdigit():
                            samples.append(int(group))

            if len(samples) >= 100:
                break
    except Exception as e:
        print(f"❌ 오류: {e}")
        continue

    if samples:
        samples = sorted(set(samples))
        print(f"✅ 샘플 수집: {len(samples)}개")
        print(f"   최소값: {min(samples)}")
        print(f"   최대값: {max(samples)}")
        print(f"   범위: {max(samples) - min(samples) + 1}")
        print(f"   샘플 (처음 10개): {samples[:10]}")
        print(f"   샘플 (마지막 10개): {samples[-10:]}")

        # 자릿수 분석
        digit_counts = defaultdict(int)
        for num in samples:
            digit_counts[len(str(num))] += 1

        print(f"   자릿수 분포:")
        for digits in sorted(digit_counts.keys()):
            print(f"      {digits}자리: {digit_counts[digits]}개")

        equipment_stats[eq_id] = {
            'name': eq_info['name'],
            'min': min(samples),
            'max': max(samples),
            'samples': len(samples),
            'digits': dict(digit_counts)
        }
    else:
        print(f"⚠️  샘플 없음")

# 최종 권장사항
print("\n" + "=" * 80)
print("권장 설정")
print("=" * 80)

if equipment_stats:
    all_maxes = [stats['max'] for stats in equipment_stats.values()]
    recommended_max = max(all_maxes) + 50000  # 여유 50000

    print(f"현재 최대 차트번호: {max(all_maxes)}")
    print(f"권장 chart_number_max: {recommended_max}")
    print(f"\n이유:")
    print(f"  - 현재 가장 큰 번호보다 50000 여유")
    print(f"  - 주민번호/생년월일 등 7-8자리 숫자는 제외됨")
else:
    print("Windows에서 실행하여 실제 데이터를 분석해주세요.")
    print("python analyze_chart_numbers.py")

print("=" * 80)
