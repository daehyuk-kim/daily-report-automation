#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
차트번호 최대값 검증 제거 패치

적용하면:
- 1 이상의 모든 차트번호 허용
- 앞자리 0은 여전히 차단
- 최대값 제한 없음
"""

import json

print("=" * 70)
print("차트번호 최대값 검증 제거")
print("=" * 70)

# 1. config.json 수정
with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

old_max = config['validation']['chart_number_max']
config['validation']['chart_number_max'] = None  # None = 무제한

with open('config.json', 'w', encoding='utf-8') as f:
    json.dump(config, f, indent=2, ensure_ascii=False)

print(f"✅ config.json 업데이트")
print(f"   기존 최대값: {old_max}")
print(f"   새 최대값: 무제한 (None)")

# 2. daily_report_fast.py 수정 (is_valid_chart_number)
print("\n⚠️  daily_report_fast.py도 수정이 필요합니다:")
print("   - is_valid_chart_number() 함수에서")
print("   - max_val == None 체크 추가")
print()
print("수동 수정 또는 자동 패치를 선택하세요.")

# 자동 패치 코드
auto_patch = input("자동 패치를 적용하시겠습니까? (y/n): ").lower()

if auto_patch == 'y':
    with open('daily_report_fast.py', 'r', encoding='utf-8') as f:
        code = f.read()

    # 기존 코드 찾기
    old_code = """    def is_valid_chart_number(self, chart_num_str: str) -> bool:
        \"\"\"차트번호 유효성 검증\"\"\"
        try:
            if chart_num_str.startswith('0') and len(chart_num_str) > 1:
                return False
            chart_num = int(chart_num_str)
            min_val = self.config['validation']['chart_number_min']
            max_val = self.config['validation']['chart_number_max']
            return min_val <= chart_num <= max_val
        except (ValueError, KeyError):
            return False"""

    # 새 코드
    new_code = """    def is_valid_chart_number(self, chart_num_str: str) -> bool:
        \"\"\"차트번호 유효성 검증\"\"\"
        try:
            if chart_num_str.startswith('0') and len(chart_num_str) > 1:
                return False
            chart_num = int(chart_num_str)
            min_val = self.config['validation']['chart_number_min']
            max_val = self.config['validation'].get('chart_number_max')

            # 최소값 체크
            if chart_num < min_val:
                return False

            # 최대값 체크 (None이면 무제한)
            if max_val is not None and chart_num > max_val:
                return False

            return True
        except (ValueError, KeyError):
            return False"""

    if old_code in code:
        code = code.replace(old_code, new_code)
        with open('daily_report_fast.py', 'w', encoding='utf-8') as f:
            f.write(code)
        print("✅ daily_report_fast.py 자동 패치 완료")
    else:
        print("❌ 코드 패턴을 찾을 수 없습니다. 수동 수정이 필요합니다.")
else:
    print("\n수동 수정이 필요합니다.")

print("\n" + "=" * 70)
print("완료!")
print("=" * 70)
