#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
scan_type 확인 스크립트
"""

import json

with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

print("=" * 70)
print("HFA scan_type 확인")
print("=" * 70)

hfa = config['equipment']['HFA']

print(f"\nHFA 전체 설정:")
for key, value in hfa.items():
    print(f"  {key:20} : {value}")

print(f"\n" + "=" * 70)
print(f"scan_type 값: '{hfa['scan_type']}'")
print(f"scan_type == 'both': {hfa['scan_type'] == 'both'}")
print(f"type(scan_type): {type(hfa['scan_type'])}")
print("=" * 70)

if hfa['scan_type'] == 'both':
    print("✅ scan_type이 'both'입니다 - 정상")
else:
    print(f"❌ scan_type이 '{hfa['scan_type']}'입니다 - 'both'여야 합니다!")
    print(f"\nconfig.json을 다시 확인하세요!")
