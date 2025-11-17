#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
디버그용 스캔 테스트 스크립트
각 장비의 오늘 날짜 폴더를 찾고 실제 매칭을 테스트합니다.
"""

import os
import re
import json
from datetime import date

def debug_scan():
    """디버그 스캔 실행"""

    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)

    today = date.today()

    results = []
    results.append("=" * 60)
    results.append(f"디버그 스캔 - {today.strftime('%Y-%m-%d')}")
    results.append("=" * 60)
    results.append("")

    # 각 장비별 테스트
    for eq_id, eq_info in config['equipment'].items():
        results.append(f"=== {eq_id} ({eq_info['name']}) ===")
        base_path = eq_info['path']
        pattern = re.compile(eq_info['pattern'])
        scan_type = eq_info['scan_type']

        results.append(f"기본 경로: {base_path}")
        results.append(f"패턴: {eq_info['pattern']}")
        results.append(f"스캔 타입: {scan_type}")

        # 1. 기본 경로 존재 확인
        if not os.path.exists(base_path):
            results.append(f"❌ 기본 경로 접근 불가!")
            results.append("")
            continue
        results.append(f"✅ 기본 경로 접근 가능")

        # 2. 오늘 날짜 폴더 경로 생성
        today_folder = None
        if 'folder_structure' in eq_info:
            folder_structure = eq_info['folder_structure']
            folder = folder_structure
            folder = folder.replace('YYYY.MM', today.strftime('%Y.%m'))
            folder = folder.replace('YYYY', today.strftime('%Y'))
            folder = folder.replace('MM.DD', today.strftime('%m.%d'))
            folder = folder.replace('MM', today.strftime('%m'))
            folder = folder.replace('DD', today.strftime('%d'))

            today_folder = os.path.join(base_path, folder)
            results.append(f"폴더 구조: {folder_structure}")
            results.append(f"오늘 폴더 경로: {today_folder}")

            if os.path.exists(today_folder):
                results.append(f"✅ 오늘 폴더 존재")
            else:
                results.append(f"❌ 오늘 폴더 없음!")
                # 상위 폴더 확인
                parent_folder = os.path.dirname(today_folder)
                if os.path.exists(parent_folder):
                    results.append(f"   상위 폴더 내용:")
                    try:
                        items = os.listdir(parent_folder)[:10]
                        for item in items:
                            results.append(f"     - {item}")
                    except Exception as e:
                        results.append(f"   상위 폴더 접근 오류: {e}")
                results.append("")
                continue
        else:
            today_folder = base_path
            results.append(f"폴더 구조 없음 - 기본 경로 사용")

        # 3. 실제 스캔 테스트
        chart_numbers = set()
        try:
            items = os.listdir(today_folder)
            results.append(f"전체 항목 수: {len(items)}개")

            # 샘플 5개 매칭 테스트
            results.append(f"샘플 매칭 테스트 (최대 10개):")
            test_count = 0

            for item in items[:30]:  # 처음 30개 중에서
                if test_count >= 10:
                    break

                # 파일/폴더 구분
                item_path = os.path.join(today_folder, item)
                is_file = os.path.isfile(item_path)
                is_dir = os.path.isdir(item_path)

                # 스캔 타입에 따라 처리
                should_scan = False
                if scan_type == 'file' and is_file:
                    # 확장자 체크
                    if any(item.lower().endswith(ext) for ext in config['validation']['file_extensions']):
                        should_scan = True
                elif scan_type == 'both':
                    if is_file:
                        if any(item.lower().endswith(ext) for ext in config['validation']['file_extensions']):
                            should_scan = True
                    elif is_dir:
                        should_scan = True

                if should_scan:
                    test_count += 1
                    match = pattern.search(item)
                    item_type = "파일" if is_file else "폴더"

                    if match:
                        chart_num = match.group(1)
                        # 유효성 검사
                        try:
                            num = int(chart_num)
                            min_val = config['validation']['chart_number_min']
                            max_val = config['validation']['chart_number_max']
                            if min_val <= num <= max_val:
                                chart_numbers.add(chart_num)
                                results.append(f"  ✅ [{item_type}] {item}")
                                results.append(f"     → 차트번호: {chart_num}")
                            else:
                                results.append(f"  ⚠️ [{item_type}] {item}")
                                results.append(f"     → 차트번호 {chart_num} 범위 초과 ({min_val}-{max_val})")
                        except:
                            results.append(f"  ⚠️ [{item_type}] {item}")
                            results.append(f"     → 차트번호 파싱 오류: {chart_num}")
                    else:
                        results.append(f"  ❌ [{item_type}] {item}")
                        results.append(f"     → 패턴 매칭 실패")

            if test_count == 0:
                results.append(f"  ⚠️ 스캔 대상 없음 (scan_type={scan_type})")
                results.append(f"     실제 항목 샘플:")
                for item in items[:5]:
                    item_path = os.path.join(today_folder, item)
                    item_type = "폴더" if os.path.isdir(item_path) else "파일"
                    results.append(f"       [{item_type}] {item}")

            results.append(f"\n총 매칭: {len(chart_numbers)}건")

        except Exception as e:
            results.append(f"❌ 스캔 오류: {e}")

        results.append("")

    # 안저 테스트
    results.append("=== 안저 (Fundus + Secondary) ===")
    fundus_config = config['special_items']['안저']['folders']

    # Fundus 테스트
    if 'fundus' in fundus_config:
        fundus_info = fundus_config['fundus']
        results.append(f"\n--- Fundus ---")
        results.append(f"경로: {fundus_info['path']}")

        if os.path.exists(fundus_info['path']):
            # 오늘 폴더 경로
            folder_structure = fundus_info.get('folder_structure', '')
            if folder_structure:
                folder = folder_structure
                folder = folder.replace('YYYY.MM', today.strftime('%Y.%m'))
                folder = folder.replace('YYYY', today.strftime('%Y'))
                folder = folder.replace('MM.DD', today.strftime('%m.%d'))
                folder = folder.replace('MM', today.strftime('%m'))
                today_folder = os.path.join(fundus_info['path'], folder)

                results.append(f"오늘 폴더: {today_folder}")
                if os.path.exists(today_folder):
                    items = os.listdir(today_folder)
                    results.append(f"항목 수: {len(items)}개")

                    pattern = re.compile(fundus_info['pattern'])
                    for item in items[:5]:
                        match = pattern.search(item)
                        if match:
                            results.append(f"  ✅ {item} → {match.group(1)}")
                        else:
                            results.append(f"  ❌ {item}")
                else:
                    results.append(f"❌ 오늘 폴더 없음")
        else:
            results.append(f"❌ 경로 접근 불가")

    # Secondary 테스트
    if 'secondary' in fundus_config:
        secondary_info = fundus_config['secondary']
        results.append(f"\n--- Secondary ---")
        results.append(f"경로: {secondary_info['path']}")

        if os.path.exists(secondary_info['path']):
            today_str = today.strftime('%Y%m%d')
            results.append(f"오늘 날짜 패턴: {today_str}")

            try:
                items = os.listdir(secondary_info['path'])
                results.append(f"전체 항목: {len(items)}개")

                # 오늘 날짜 파일만
                today_files = [f for f in items if today_str in f]
                results.append(f"오늘 날짜 파일: {len(today_files)}개")

                pattern = re.compile(secondary_info['pattern'])
                for item in today_files[:5]:
                    match = pattern.search(item)
                    if match:
                        results.append(f"  ✅ {item} → {match.group(1)}")
                    else:
                        results.append(f"  ❌ {item}")
            except Exception as e:
                results.append(f"❌ 스캔 오류: {e}")
        else:
            results.append(f"❌ 경로 접근 불가")

    results.append("")

    # 결과 저장
    output_file = f"debug_scan_{today.strftime('%Y%m%d')}.txt"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(results))

    print(f"디버그 결과가 {output_file}에 저장되었습니다.")
    print('\n'.join(results))

if __name__ == "__main__":
    debug_scan()
