#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
백업 스크립트 방식 스캔 테스트
경로에 날짜가 포함된 파일만 빠르게 필터링

장점:
- getctime() 호출 없음 (매우 빠름)
- 경로 문자열 체크만으로 오늘 파일 판별

단점:
- 경로에 날짜가 없는 폴더 (SP, HFA)에서는 사용 불가
"""

import os
import re
from datetime import date, timedelta
import time


def scan_with_path_date_filter(source_dir, date_patterns):
    """백업 스크립트 방식: 경로에 날짜 포함 여부로 필터링"""
    matching_files = []
    total_scanned = 0

    try:
        for root, dirs, files in os.walk(source_dir):
            for filename in files:
                total_scanned += 1
                file_path = os.path.join(root, filename)

                # 경로에 오늘 날짜가 포함되어 있는지 확인
                if any(dp in file_path for dp in date_patterns):
                    matching_files.append(file_path)

            # 진행 상황 (매 1000개마다)
            if total_scanned % 1000 == 0:
                print(f"   ... {total_scanned}개 스캔, {len(matching_files)}개 매칭")

    except Exception as e:
        print(f"❌ 스캔 오류: {e}")

    return matching_files, total_scanned


def test_backup_style_scan():
    """백업 스크립트 방식 테스트"""
    print("=" * 80)
    print("백업 스크립트 방식 스캔 테스트")
    print("=" * 80)
    print()

    # 오늘/어제 날짜 패턴
    today = date.today()
    yesterday = today - timedelta(days=1)

    date_patterns = [
        today.strftime('%m.%d'),      # 11.16
        yesterday.strftime('%m.%d'),  # 11.15
        today.strftime('%Y%m%d'),     # 20251116
        today.strftime('%Y-%m-%d'),   # 2025-11-16
    ]

    print(f"📅 검색할 날짜 패턴: {date_patterns}")
    print()

    # 테스트할 경로들 (백업 스크립트에서 가져옴)
    test_paths = {
        "OCT": r"\\192.168.0.221\pdf",
        "TOPO": r"\\topo\topo",
        "ORB": r"\\192.168.0.209\orb",
        "OQAS": r"\\192.168.0.228\oqass",
        # SP와 HFA는 경로에 날짜가 없어서 이 방식 부적합
        # "SP": r"\\192.168.0.120\sp",
        # "HFA": r"\\GEOMSA-MAIN2\hfa",
    }

    results = {}

    for name, path in test_paths.items():
        print(f"\n📁 {name}: {path}")
        print("-" * 80)

        if not os.path.exists(path):
            print(f"   ❌ 경로 접근 불가")
            continue

        start_time = time.time()
        matching_files, total_scanned = scan_with_path_date_filter(path, date_patterns)
        elapsed = time.time() - start_time

        print(f"   ✅ 완료!")
        print(f"   📊 전체: {total_scanned:,}개 / 매칭: {len(matching_files)}개")
        print(f"   ⏱️  소요 시간: {elapsed:.2f}초")

        if matching_files:
            print(f"   📄 샘플 (최대 5개):")
            for f in matching_files[:5]:
                print(f"      - {f}")

        results[name] = {
            'total': total_scanned,
            'matched': len(matching_files),
            'time': elapsed
        }

    # 결과 요약
    print("\n" + "=" * 80)
    print("📋 결과 요약")
    print("=" * 80)

    for name, data in results.items():
        print(f"{name:10s}: {data['matched']:5,}개 매칭 / {data['total']:10,}개 중 / {data['time']:.2f}초")

    print()
    print("💡 결론:")
    print("- 경로에 날짜가 포함된 폴더 (OCT, TOPO, ORB, OQAS): 이 방식이 가장 빠름")
    print("- 경로에 날짜가 없는 폴더 (SP, HFA): getctime() 또는 캐시 필요")


if __name__ == "__main__":
    test_backup_style_scan()
