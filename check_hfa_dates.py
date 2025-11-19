#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""HFA 간단 테스트 - 날짜별"""

import os
import json
import re
from datetime import date, timedelta

with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

hfa = config['equipment']['HFA']
base_path = hfa['path']
pattern = re.compile(hfa['pattern'])

print("=" * 70)
print("HFA 날짜별 테스트")
print("=" * 70)

# 최근 7일 테스트
for days_ago in range(7):
    test_date = date.today() - timedelta(days=days_ago)

    # 날짜 폴더 경로 생성
    folder = hfa['folder_structure']
    folder = folder.replace('YYYY', test_date.strftime('%Y'))
    folder = folder.replace('MM.DD', test_date.strftime('%m.%d'))
    folder = folder.replace('MM', test_date.strftime('%m'))

    date_folder = os.path.join(base_path, folder)

    print(f"\n[{test_date.strftime('%Y-%m-%d')}]")
    print(f"경로: {date_folder}")

    if os.path.exists(date_folder):
        try:
            items = os.listdir(date_folder)
            dirs = [item for item in items if os.path.isdir(os.path.join(date_folder, item))]

            chart_numbers = set()
            for dir_name in dirs:
                match = pattern.search(dir_name)
                if match:
                    chart_num = match.group(1) or (match.group(2) if match.lastindex > 1 else None)
                    if chart_num:
                        chart_numbers.add(chart_num)

            print(f"✅ 폴더 존재 | 총 {len(dirs)}개 폴더 | 차트번호 {len(chart_numbers)}건")
            if len(chart_numbers) > 0:
                print(f"   샘플: {sorted(chart_numbers)[:3]}")
        except Exception as e:
            print(f"❌ 오류: {e}")
    else:
        print(f"❌ 폴더 없음")

print("\n" + "=" * 70)
