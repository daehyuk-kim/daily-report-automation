#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
코드 버전 확인 스크립트
"""

import json
import inspect

print("=" * 70)
print("코드 버전 확인")
print("=" * 70)

# 1. config.json의 HFA 패턴 확인
print("\n[1] config.json HFA 패턴:")
with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)
print(f"   {config['equipment']['HFA']['pattern']}")

expected_pattern = r"_(\d{5,6})$|^(\d{5,6})_"
if config['equipment']['HFA']['pattern'] == expected_pattern:
    print(f"   ✅ 최신 패턴")
else:
    print(f"   ❌ 구 패턴! 업데이트 필요!")

# 2. daily_report_fast.py에 extract_chart_number 메소드 있는지 확인
print("\n[2] daily_report_fast.py 메소드 확인:")
try:
    from daily_report_fast import DailyReportSystem

    # extract_chart_number 메소드 존재 확인
    if hasattr(DailyReportSystem, 'extract_chart_number'):
        print(f"   ✅ extract_chart_number 메소드 존재")

        # 메소드 소스 코드 확인
        source = inspect.getsource(DailyReportSystem.extract_chart_number)
        if 'match.group(2)' in source:
            print(f"   ✅ 이중 그룹 처리 로직 포함")
        else:
            print(f"   ❌ 이중 그룹 처리 로직 없음!")
    else:
        print(f"   ❌ extract_chart_number 메소드 없음! 구버전!")

except Exception as e:
    print(f"   ❌ 오류: {e}")

print("\n" + "=" * 70)
print("결론:")
print("=" * 70)

if (config['equipment']['HFA']['pattern'] == expected_pattern and
    hasattr(DailyReportSystem, 'extract_chart_number')):
    print("✅ 최신 코드 적용됨")
    print("\n다른 문제가 있을 수 있습니다. 상세 로그가 필요합니다.")
else:
    print("❌ 구버전 코드 사용 중!")
    print("\n해결: git pull 후 프로그램 재시작")

print("=" * 70)
