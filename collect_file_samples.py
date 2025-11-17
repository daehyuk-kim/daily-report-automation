#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
파일명 샘플 수집 스크립트

각 장비 디렉토리를 깊이 탐색하여 실제 파일명 1000개씩 수집
Claude가 파일명 패턴을 분석할 수 있도록 텍스트 파일로 저장
"""

import os
import json
from datetime import date, datetime


def load_config(config_path="config_real.json"):
    """설정 파일 로드"""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"설정 파일 없음: {config_path}")
        return None


def collect_files_from_directory(path, max_files=1000, min_year=2019):
    """디렉토리에서 파일명 수집 (깊이 탐색, 선택적 연도 필터)"""
    files = []
    skipped_old = 0

    if not os.path.exists(path):
        return files, "경로 접근 불가", skipped_old

    # 연도 필터 설정 (None이면 필터 없음)
    min_timestamp = None
    if min_year is not None:
        min_timestamp = datetime(min_year, 1, 1).timestamp()

    try:
        count = 0
        for root, dirs, filenames in os.walk(path):
            for filename in filenames:
                # 파일의 전체 경로와 파일명 둘 다 저장
                full_path = os.path.join(root, filename)

                # 연도 필터가 있을 때만 적용
                if min_timestamp is not None:
                    try:
                        file_stat = os.stat(full_path)
                        # 생성일과 수정일 중 더 최근 것 사용
                        file_time = max(file_stat.st_ctime, file_stat.st_mtime)

                        if file_time < min_timestamp:
                            skipped_old += 1
                            continue
                    except:
                        # stat 실패 시 일단 포함
                        pass

                # 상대 경로 계산
                rel_path = os.path.relpath(full_path, path)
                files.append({
                    'filename': filename,
                    'relative_path': rel_path,
                    'extension': os.path.splitext(filename)[1].lower()
                })
                count += 1
                if count >= max_files:
                    return files, "OK", skipped_old

            # 진행 상황 표시 (매 100개마다)
            if count > 0 and count % 100 == 0:
                if min_timestamp:
                    print(f"      ... {count}개 수집 중 (2019 이전 {skipped_old}개 스킵)")
                else:
                    print(f"      ... {count}개 수집 중")

        return files, "OK", skipped_old

    except Exception as e:
        return files, str(e), skipped_old


def main():
    """메인 실행"""
    print("=" * 80)
    print("파일명 샘플 수집 스크립트")
    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)
    print()

    # 설정 파일 로드
    config = load_config()
    if not config:
        return

    # 결과 저장할 파일
    output_file = f"file_samples_{date.today().strftime('%Y%m%d')}.txt"

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("=" * 80 + "\n")
        f.write("파일명 샘플 수집 결과\n")
        f.write(f"수집 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 80 + "\n\n")

        # 1. 장비 디렉토리 수집
        print("1. 장비 디렉토리 파일 수집")
        print("-" * 80)

        # SP, ORB, TOPO는 2019년 이전 파일도 수집
        no_filter_equipment = ['sp', 'orb', 'topo']

        for equipment_id, equipment_info in config['equipment'].items():
            name = equipment_info['name']
            path = equipment_info['path']

            print(f"\n{name} ({equipment_id})")
            print(f"   경로: {path}")

            # SP, ORB, TOPO는 필터 없이, 나머지는 2019년 이후만
            if equipment_id.lower() in no_filter_equipment:
                files, status, skipped = collect_files_from_directory(path, max_files=1000, min_year=None)
                filter_msg = "필터 없음"
            else:
                files, status, skipped = collect_files_from_directory(path, max_files=1000, min_year=2019)
                filter_msg = f"2019년 이전 {skipped}개 스킵"

            f.write(f"\n{'='*80}\n")
            f.write(f"{name} ({equipment_id})\n")
            f.write(f"경로: {path}\n")
            f.write(f"상태: {status}\n")
            f.write(f"수집된 파일 수: {len(files)} ({filter_msg})\n")
            f.write(f"{'='*80}\n\n")

            if files:
                print(f"   수집: {len(files)}개 파일 ({filter_msg})")

                # 확장자별 통계
                ext_count = {}
                for file_info in files:
                    ext = file_info['extension'] or '(없음)'
                    ext_count[ext] = ext_count.get(ext, 0) + 1

                f.write("확장자별 통계:\n")
                for ext, count in sorted(ext_count.items(), key=lambda x: -x[1]):
                    f.write(f"  {ext}: {count}개\n")
                f.write("\n")

                # 파일명 샘플 (상대 경로 포함)
                f.write("파일 목록 (상대경로):\n")
                f.write("-" * 80 + "\n")
                for file_info in files:
                    f.write(f"{file_info['relative_path']}\n")
                f.write("\n")

                # 파일명만 (폴더 구조 없이)
                f.write("파일명만:\n")
                f.write("-" * 80 + "\n")
                for file_info in files[:100]:  # 상위 100개만
                    f.write(f"{file_info['filename']}\n")
                if len(files) > 100:
                    f.write(f"... 외 {len(files) - 100}개\n")
                f.write("\n")
            else:
                print(f"   상태: {status}")
                f.write(f"파일 없음 또는 접근 불가\n\n")

        # 2. 특별 항목 (안저) 수집
        print("\n2. 특별 항목 디렉토리 파일 수집")
        print("-" * 80)

        if '안저' in config['special_items']:
            for idx, folder in enumerate(config['special_items']['안저']['folders'], 1):
                print(f"\n안저 경로 {idx}")
                print(f"   경로: {folder}")

                files, status, skipped = collect_files_from_directory(folder, max_files=1000)

                f.write(f"\n{'='*80}\n")
                f.write(f"안저 경로 {idx}\n")
                f.write(f"경로: {folder}\n")
                f.write(f"상태: {status}\n")
                f.write(f"수집된 파일 수: {len(files)} (2019년 이전 {skipped}개 스킵)\n")
                f.write(f"{'='*80}\n\n")

                if files:
                    print(f"   수집: {len(files)}개 파일 (2019 이전 {skipped}개 스킵)")

                    # 확장자별 통계
                    ext_count = {}
                    for file_info in files:
                        ext = file_info['extension'] or '(없음)'
                        ext_count[ext] = ext_count.get(ext, 0) + 1

                    f.write("확장자별 통계:\n")
                    for ext, count in sorted(ext_count.items(), key=lambda x: -x[1]):
                        f.write(f"  {ext}: {count}개\n")
                    f.write("\n")

                    # 파일명 샘플
                    f.write("파일 목록 (상대경로):\n")
                    f.write("-" * 80 + "\n")
                    for file_info in files:
                        f.write(f"{file_info['relative_path']}\n")
                    f.write("\n")

                    f.write("파일명만:\n")
                    f.write("-" * 80 + "\n")
                    for file_info in files[:100]:
                        f.write(f"{file_info['filename']}\n")
                    if len(files) > 100:
                        f.write(f"... 외 {len(files) - 100}개\n")
                    f.write("\n")
                else:
                    print(f"   상태: {status}")
                    f.write(f"파일 없음 또는 접근 불가\n\n")

    print("\n" + "=" * 80)
    print(f"결과 저장됨: {output_file}")
    print("=" * 80)
    print()
    print("다음 단계:")
    print(f"1. 생성된 '{output_file}' 파일을 Claude에게 전송")
    print("2. Claude가 파일명 패턴을 분석하여 차트번호 추출 로직 작성")
    print("3. 최적화된 스캔 방식 결정")


if __name__ == "__main__":
    main()
