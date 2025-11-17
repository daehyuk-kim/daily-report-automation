#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
장비별 샘플 데이터 추출 스크립트
2025년 10월 데이터를 기준으로 폴더 구조와 파일명 샘플을 추출합니다.
"""

import os
import json
from datetime import datetime

def extract_samples():
    """각 장비별 샘플 데이터 추출"""

    # config.json 로드
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)

    results = []
    results.append("=" * 60)
    results.append(f"샘플 데이터 추출 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    results.append("=" * 60)
    results.append("")

    # 1. 장비별 샘플 추출
    for eq_id, eq_info in config['equipment'].items():
        results.append(f"=== {eq_id} ({eq_info['name']}) ===")
        base_path = eq_info['path']
        results.append(f"기본 경로: {base_path}")

        if not os.path.exists(base_path):
            results.append(f"❌ 경로 접근 불가: {base_path}")
            results.append("")
            continue

        # 폴더 구조 탐색 (2025/10월 데이터)
        target_year = "2025"
        target_month = "10"

        # 가능한 폴더 구조들 시도
        possible_paths = [
            # TOPO 형식: 2025\2025.10\TOPO 10.XX
            os.path.join(base_path, target_year, f"{target_year}.{target_month}"),
            # ORB 형식: 2025\2025.10\ORB 10.XX
            os.path.join(base_path, target_year, f"{target_year}.{target_month}"),
            # OCT 형식: 2025\10\XX
            os.path.join(base_path, target_year, target_month),
            # OQAS 형식: 2025\10\XX.10
            os.path.join(base_path, target_year, target_month),
            # 직접 경로 (SP, HFA 등)
            base_path
        ]

        found_samples = False

        for check_path in possible_paths:
            if not os.path.exists(check_path):
                continue

            results.append(f"탐색 경로: {check_path}")

            # 하위 폴더 확인
            try:
                items = os.listdir(check_path)

                # 날짜별 폴더가 있는 경우 (TOPO 10.15, ORB 10.15, oct 10.15 등)
                date_folders = [f for f in items if '10.' in f or f.isdigit()]

                if date_folders:
                    # 날짜 폴더 중 하나 선택 (중간 날짜)
                    date_folders.sort()
                    sample_folder = date_folders[len(date_folders)//2] if len(date_folders) > 1 else date_folders[0]
                    sample_path = os.path.join(check_path, sample_folder)

                    results.append(f"샘플 폴더: {sample_folder}")
                    results.append(f"전체 경로: {sample_path}")

                    if os.path.exists(sample_path) and os.path.isdir(sample_path):
                        # 폴더 내 파일/폴더 샘플 추출
                        folder_items = os.listdir(sample_path)[:10]  # 최대 10개
                        results.append(f"항목 수: {len(os.listdir(sample_path))}개")
                        results.append("샘플 항목:")
                        for i, item in enumerate(folder_items, 1):
                            item_path = os.path.join(sample_path, item)
                            item_type = "폴더" if os.path.isdir(item_path) else "파일"

                            # 수정 시간 확인
                            try:
                                mtime = os.path.getmtime(item_path)
                                mtime_str = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M')
                            except:
                                mtime_str = "N/A"

                            results.append(f"  {i}. [{item_type}] {item}")
                            results.append(f"     수정일: {mtime_str}")

                        found_samples = True
                        break
                else:
                    # 날짜 폴더가 없는 경우 (SP, HFA 등) - 직접 파일 샘플
                    # 2025년 10월 생성 파일 찾기
                    results.append(f"전체 항목 수: {len(items)}개")

                    sample_files = []
                    for item in items[:100]:  # 처음 100개만 확인
                        item_path = os.path.join(check_path, item)
                        if os.path.isfile(item_path):
                            try:
                                ctime = os.path.getctime(item_path)
                                ctime_dt = datetime.fromtimestamp(ctime)
                                # 2025년 10월 파일
                                if ctime_dt.year == 2025 and ctime_dt.month == 10:
                                    sample_files.append((item, ctime_dt))
                            except:
                                pass

                    if sample_files:
                        results.append(f"2025년 10월 생성 파일 샘플:")
                        for i, (fname, cdt) in enumerate(sample_files[:10], 1):
                            results.append(f"  {i}. {fname}")
                            results.append(f"     생성일: {cdt.strftime('%Y-%m-%d %H:%M')}")
                        found_samples = True
                    else:
                        # 최근 파일이라도 샘플로
                        results.append("2025년 10월 파일 없음. 최근 파일 샘플:")
                        for i, item in enumerate(items[:10], 1):
                            item_path = os.path.join(check_path, item)
                            if os.path.isfile(item_path):
                                try:
                                    ctime = os.path.getctime(item_path)
                                    ctime_str = datetime.fromtimestamp(ctime).strftime('%Y-%m-%d %H:%M')
                                except:
                                    ctime_str = "N/A"
                                results.append(f"  {i}. {item}")
                                results.append(f"     생성일: {ctime_str}")
                        found_samples = True
                    break

            except Exception as e:
                results.append(f"❌ 오류: {e}")

        if not found_samples:
            results.append("⚠️ 샘플 데이터를 찾을 수 없습니다.")

        results.append("")

    # 2. 안저 폴더 샘플
    results.append("=== 안저 (Fundus) ===")
    for folder in config['special_items']['안저']['folders']:
        results.append(f"경로: {folder}")

        if not os.path.exists(folder):
            results.append(f"❌ 경로 접근 불가")
            continue

        try:
            items = os.listdir(folder)
            results.append(f"전체 항목 수: {len(items)}개")

            # 2025년 10월 생성 파일 찾기
            sample_files = []
            for item in items[:200]:
                item_path = os.path.join(folder, item)
                if os.path.isfile(item_path):
                    try:
                        ctime = os.path.getctime(item_path)
                        ctime_dt = datetime.fromtimestamp(ctime)
                        if ctime_dt.year == 2025 and ctime_dt.month == 10:
                            sample_files.append((item, ctime_dt))
                    except:
                        pass

            if sample_files:
                results.append(f"2025년 10월 생성 파일 샘플:")
                for i, (fname, cdt) in enumerate(sample_files[:5], 1):
                    results.append(f"  {i}. {fname}")
                    results.append(f"     생성일: {cdt.strftime('%Y-%m-%d %H:%M')}")
            else:
                results.append("최근 파일 샘플:")
                for i, item in enumerate(items[:5], 1):
                    results.append(f"  {i}. {item}")
        except Exception as e:
            results.append(f"❌ 오류: {e}")

        results.append("")

    # 결과 저장
    output_file = "samples_2025_10.txt"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(results))

    print(f"샘플 데이터가 {output_file}에 저장되었습니다.")
    print("\n".join(results))

if __name__ == "__main__":
    extract_samples()
