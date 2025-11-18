#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
상세 디버그 스크립트 - HFA와 안저 문제 진단
사용법: python debug_detailed.py 2025-11-15
"""

import os
import re
import json
import sys
from datetime import datetime, date

def debug_detailed(target_date):
    """상세 디버그 실행"""

    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)

    results = []
    results.append("=" * 80)
    results.append(f"상세 디버그 - {target_date.strftime('%Y-%m-%d')}")
    results.append("=" * 80)
    results.append("")

    # ========================================
    # HFA (시야) 상세 분석
    # ========================================
    results.append("=" * 80)
    results.append("HFA (시야) 상세 분석")
    results.append("=" * 80)

    hfa_config = config['equipment']['HFA']
    base_path = hfa_config['path']
    pattern = re.compile(hfa_config['pattern'])
    scan_type = hfa_config['scan_type']

    results.append(f"기본 경로: {base_path}")
    results.append(f"패턴: {hfa_config['pattern']}")
    results.append(f"스캔 타입: {scan_type}")
    results.append("")

    if not os.path.exists(base_path):
        results.append("❌ 기본 경로 접근 불가!")
    else:
        results.append("✅ 기본 경로 접근 가능")

        # 날짜 폴더 경로
        folder_structure = hfa_config['folder_structure']
        folder = folder_structure
        folder = folder.replace('YYYY.MM', target_date.strftime('%Y.%m'))
        folder = folder.replace('YYYY', target_date.strftime('%Y'))
        folder = folder.replace('MM.DD', target_date.strftime('%m.%d'))
        folder = folder.replace('MM', target_date.strftime('%m'))
        folder = folder.replace('DD', target_date.strftime('%d'))

        today_folder = os.path.join(base_path, folder)
        results.append(f"대상 폴더: {today_folder}")

        if not os.path.exists(today_folder):
            results.append("❌ 대상 폴더 없음!")

            # 상위 폴더 확인
            parent = os.path.dirname(today_folder)
            if os.path.exists(parent):
                results.append(f"\n상위 폴더 내용 ({parent}):")
                items = sorted(os.listdir(parent))[-20:]
                for item in items:
                    results.append(f"  - {item}")
        else:
            results.append("✅ 대상 폴더 존재!")
            results.append("")

            # 폴더 내용 상세 분석
            try:
                all_items = os.listdir(today_folder)
                results.append(f"전체 항목 수: {len(all_items)}개")
                results.append("")

                # 파일과 폴더 구분
                files = []
                dirs = []
                for item in all_items:
                    item_path = os.path.join(today_folder, item)
                    if os.path.isfile(item_path):
                        files.append(item)
                    elif os.path.isdir(item_path):
                        dirs.append(item)

                results.append(f"파일: {len(files)}개")
                results.append(f"폴더: {len(dirs)}개")
                results.append("")

                # 파일 샘플
                if files:
                    results.append("파일 샘플 (최대 10개):")
                    for f in files[:10]:
                        file_path = os.path.join(today_folder, f)
                        ext = os.path.splitext(f)[1].lower()
                        match = pattern.search(f)

                        if match:
                            results.append(f"  ✅ [파일{ext}] {f}")
                            results.append(f"      → 차트번호: {match.group(1)}")
                        else:
                            results.append(f"  ❌ [파일{ext}] {f}")
                            results.append(f"      → 패턴 매칭 실패")
                    results.append("")

                # 폴더 샘플
                if dirs:
                    results.append("폴더 샘플 (최대 10개):")
                    chart_numbers = set()

                    for d in dirs[:10]:
                        match = pattern.search(d)

                        if match:
                            chart_num = match.group(1)
                            try:
                                num = int(chart_num)
                                min_val = config['validation']['chart_number_min']
                                max_val = config['validation']['chart_number_max']

                                if min_val <= num <= max_val:
                                    chart_numbers.add(chart_num)
                                    results.append(f"  ✅ [폴더] {d}")
                                    results.append(f"      → 차트번호: {chart_num} (유효)")
                                else:
                                    results.append(f"  ⚠️ [폴더] {d}")
                                    results.append(f"      → 차트번호: {chart_num} (범위 초과)")
                            except:
                                results.append(f"  ⚠️ [폴더] {d}")
                                results.append(f"      → 차트번호 파싱 오류")
                        else:
                            results.append(f"  ❌ [폴더] {d}")
                            results.append(f"      → 패턴 매칭 실패")

                    results.append("")
                    results.append(f"매칭된 차트번호 (중복제거): {len(chart_numbers)}건")

                    # 전체 폴더 스캔
                    if len(dirs) > 10:
                        results.append("")
                        results.append("전체 폴더 스캔 중...")
                        all_chart_numbers = set()
                        for d in dirs:
                            match = pattern.search(d)
                            if match:
                                chart_num = match.group(1)
                                try:
                                    num = int(chart_num)
                                    if min_val <= num <= max_val:
                                        all_chart_numbers.add(chart_num)
                                except:
                                    pass
                        results.append(f"전체 매칭 결과: {len(all_chart_numbers)}건")

            except Exception as e:
                results.append(f"❌ 폴더 스캔 오류: {e}")

    results.append("")
    results.append("")

    # ========================================
    # 안저 상세 분석
    # ========================================
    results.append("=" * 80)
    results.append("안저 상세 분석")
    results.append("=" * 80)

    fundus_config = config['special_items']['안저']['folders']

    # Fundus 폴더
    if 'fundus' in fundus_config:
        fundus_info = fundus_config['fundus']
        results.append("\n[1] Fundus 폴더")
        results.append(f"경로: {fundus_info['path']}")

        if os.path.exists(fundus_info['path']):
            # 날짜 폴더
            folder_structure = fundus_info.get('folder_structure', '')
            if folder_structure:
                folder = folder_structure
                folder = folder.replace('YYYY.MM', target_date.strftime('%Y.%m'))
                folder = folder.replace('YYYY', target_date.strftime('%Y'))
                folder = folder.replace('MM.DD', target_date.strftime('%m.%d'))
                folder = folder.replace('MM', target_date.strftime('%m'))
                today_folder = os.path.join(fundus_info['path'], folder)

                results.append(f"대상 폴더: {today_folder}")

                if os.path.exists(today_folder):
                    results.append("✅ 폴더 존재")

                    items = os.listdir(today_folder)
                    results.append(f"전체 항목: {len(items)}개")

                    pattern = re.compile(fundus_info['pattern'])
                    fundus_charts = set()

                    results.append("\n파일 샘플 (최대 20개):")
                    for item in items[:20]:
                        match = pattern.search(item)
                        if match:
                            chart_num = match.group(1)
                            fundus_charts.add(chart_num)
                            results.append(f"  ✅ {item}")
                            results.append(f"      → 차트번호: {chart_num}")
                        else:
                            results.append(f"  ❌ {item}")
                            results.append(f"      → 패턴 매칭 실패")

                    # 전체 스캔
                    all_fundus_charts = set()
                    for item in items:
                        match = pattern.search(item)
                        if match:
                            chart_num = match.group(1)
                            try:
                                num = int(chart_num)
                                if config['validation']['chart_number_min'] <= num <= config['validation']['chart_number_max']:
                                    all_fundus_charts.add(chart_num)
                            except:
                                pass

                    results.append(f"\n✅ Fundus 매칭 결과: {len(all_fundus_charts)}건 (중복제거)")
                else:
                    results.append("❌ 폴더 없음")
        else:
            results.append("❌ 경로 접근 불가")

    # Secondary 폴더
    if 'secondary' in fundus_config:
        secondary_info = fundus_config['secondary']
        results.append("\n\n[2] Secondary 폴더")
        results.append(f"경로: {secondary_info['path']}")

        if os.path.exists(secondary_info['path']):
            results.append("✅ 경로 접근 가능")

            today_str = target_date.strftime('%Y%m%d')
            results.append(f"날짜 필터: {today_str}")

            try:
                items = os.listdir(secondary_info['path'])
                results.append(f"전체 파일: {len(items)}개")

                # 날짜 필터링
                today_files = [f for f in items if today_str in f]
                results.append(f"해당 날짜 파일: {len(today_files)}개")

                pattern = re.compile(secondary_info['pattern'])
                secondary_charts = set()

                results.append("\n파일 샘플 (최대 20개):")
                for item in today_files[:20]:
                    match = pattern.search(item)
                    if match:
                        chart_num = match.group(1)
                        secondary_charts.add(chart_num)
                        results.append(f"  ✅ {item}")
                        results.append(f"      → 차트번호: {chart_num}")
                    else:
                        results.append(f"  ❌ {item}")
                        results.append(f"      → 패턴 매칭 실패")

                # 전체 스캔
                all_secondary_charts = set()
                for item in today_files:
                    match = pattern.search(item)
                    if match:
                        chart_num = match.group(1)
                        try:
                            num = int(chart_num)
                            if config['validation']['chart_number_min'] <= num <= config['validation']['chart_number_max']:
                                all_secondary_charts.add(chart_num)
                        except:
                            pass

                results.append(f"\n✅ Secondary 매칭 결과: {len(all_secondary_charts)}건 (중복제거)")

                # 통계 정보
                if all_secondary_charts:
                    # 중복 분석
                    from collections import Counter
                    chart_count = Counter()
                    for item in today_files:
                        match = pattern.search(item)
                        if match:
                            chart_num = match.group(1)
                            try:
                                num = int(chart_num)
                                if config['validation']['chart_number_min'] <= num <= config['validation']['chart_number_max']:
                                    chart_count[chart_num] += 1
                            except:
                                pass

                    results.append(f"\n파일 중복 통계:")
                    results.append(f"  유니크 환자 수: {len(chart_count)}명")
                    results.append(f"  총 파일 수: {sum(chart_count.values())}개")
                    results.append(f"  환자당 평균 파일: {sum(chart_count.values()) / len(chart_count):.1f}개")

                    # 가장 많은 파일을 가진 환자
                    top_patients = chart_count.most_common(5)
                    results.append(f"\n  파일이 가장 많은 환자:")
                    for chart_num, count in top_patients:
                        results.append(f"    - 차트번호 {chart_num}: {count}개")

            except Exception as e:
                results.append(f"❌ 스캔 오류: {e}")
        else:
            results.append("❌ 경로 접근 불가")

    results.append("")
    results.append("=" * 80)

    # 결과 저장
    output_file = f"debug_detailed_{target_date.strftime('%Y%m%d')}.txt"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(results))

    print(f"\n상세 디버그 결과가 {output_file}에 저장되었습니다.\n")
    print('\n'.join(results))

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # 날짜 인자가 있으면 파싱
        date_str = sys.argv[1]
        try:
            target_date = datetime.strptime(date_str, '%Y-%m-%d').date()
        except:
            print(f"날짜 형식 오류: {date_str}")
            print("사용법: python debug_detailed.py 2025-11-15")
            sys.exit(1)
    else:
        target_date = date.today()

    debug_detailed(target_date)
