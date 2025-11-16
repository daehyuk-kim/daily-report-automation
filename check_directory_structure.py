#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
디렉토리별 실제 파일 구조 확인 스크립트

각 장비 디렉토리의 실제 파일명과 폴더 구조를 확인하여
최적의 스캔 방식을 결정합니다.
"""

import os
import json
from datetime import date, datetime
import time


def load_config(config_path="config_real.json"):
    """설정 파일 로드"""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"❌ 설정 파일 없음: {config_path}")
        return None


def check_single_directory(name, path, max_files=20):
    """단일 디렉토리 구조 확인"""
    print(f"\n{'='*80}")
    print(f"📁 {name}")
    print(f"   경로: {path}")
    print(f"{'='*80}")

    if not os.path.exists(path):
        print(f"   ❌ 경로 접근 불가")
        return None

    result = {
        'name': name,
        'path': path,
        'accessible': True,
        'structure': None,
        'sample_files': [],
        'sample_folders': [],
        'has_date_in_path': False,
        'has_date_in_filename': False,
    }

    # 1. 최상위 내용 확인
    print(f"\n   📂 최상위 내용:")
    print(f"   {'-'*70}")

    try:
        items = os.listdir(path)
        files = []
        folders = []

        for item in items[:100]:  # 최대 100개만
            item_path = os.path.join(path, item)
            if os.path.isfile(item_path):
                files.append(item)
            elif os.path.isdir(item_path):
                folders.append(item)

        print(f"   파일: {len(files)}개 / 폴더: {len(folders)}개")

        # 폴더 샘플 (연도별 구조인지 확인)
        if folders:
            print(f"\n   📁 폴더 샘플:")
            for folder in sorted(folders)[:10]:
                print(f"      - {folder}")
            if len(folders) > 10:
                print(f"      ... 외 {len(folders) - 10}개")

            # 연도 폴더 확인
            year_folders = [f for f in folders if f.isdigit() and len(f) == 4]
            if year_folders:
                result['structure'] = 'yearly'
                print(f"\n   🗓️ 연도별 폴더 구조 감지: {year_folders}")

                # 현재 연도 폴더 탐색
                current_year = str(date.today().year)
                if current_year in folders:
                    year_path = os.path.join(path, current_year)
                    explore_year_folder(year_path, result)
            else:
                # 날짜 패턴이 폴더명에 있는지 확인
                date_folders = [f for f in folders if any(c.isdigit() for c in f)]
                if date_folders:
                    print(f"\n   📅 날짜 포함 폴더:")
                    for df in sorted(date_folders)[:5]:
                        print(f"      - {df}")

        # 파일 샘플
        if files:
            print(f"\n   📄 파일 샘플:")
            for f in sorted(files)[:max_files]:
                # 파일 생성일도 같이 표시
                try:
                    file_path = os.path.join(path, f)
                    ctime = os.path.getctime(file_path)
                    ctime_str = datetime.fromtimestamp(ctime).strftime('%Y-%m-%d %H:%M')
                    print(f"      - {f}")
                    print(f"        생성일: {ctime_str}")
                except:
                    print(f"      - {f}")
            if len(files) > max_files:
                print(f"      ... 외 {len(files) - max_files}개")

            result['sample_files'] = files[:max_files]

            # 파일명에 날짜 패턴이 있는지 확인
            today = date.today()
            date_patterns = [
                today.strftime('%Y%m%d'),     # 20251117
                today.strftime('%m.%d'),      # 11.17
                today.strftime('%Y-%m-%d'),   # 2025-11-17
                today.strftime('%Y.%m.%d'),   # 2025.11.17
            ]

            for f in files[:20]:
                if any(dp in f for dp in date_patterns):
                    result['has_date_in_filename'] = True
                    print(f"\n   ✅ 파일명에 날짜 패턴 발견!")
                    break

        result['sample_folders'] = folders[:10]

    except Exception as e:
        print(f"   ❌ 오류: {e}")
        result['accessible'] = False

    return result


def explore_year_folder(year_path, result):
    """연도 폴더 내부 탐색"""
    print(f"\n   📂 연도 폴더 내부: {year_path}")

    try:
        items = os.listdir(year_path)
        folders = [i for i in items if os.path.isdir(os.path.join(year_path, i))]
        files = [i for i in items if os.path.isfile(os.path.join(year_path, i))]

        print(f"      폴더: {len(folders)}개 / 파일: {len(files)}개")

        if folders:
            print(f"\n      📁 하위 폴더:")
            for f in sorted(folders)[:10]:
                print(f"         - {f}")

            # 월별 폴더 확인
            month_folders = [f for f in folders if f.isdigit() and len(f) <= 2]
            if month_folders:
                print(f"\n      🗓️ 월별 폴더 감지: {sorted(month_folders)}")

                # 현재 월 폴더 탐색
                current_month = str(date.today().month).zfill(2)
                if current_month in folders:
                    month_path = os.path.join(year_path, current_month)
                    explore_month_folder(month_path, result)
            else:
                # 다른 패턴 확인 (예: 2025.01, ORB 11.16 등)
                for f in folders[:5]:
                    if '.' in f:
                        print(f"\n      📅 특수 폴더 패턴: {f}")
                        result['has_date_in_path'] = True

    except Exception as e:
        print(f"      ❌ 오류: {e}")


def explore_month_folder(month_path, result):
    """월 폴더 내부 탐색"""
    print(f"\n      📂 월 폴더 내부: {month_path}")

    try:
        items = os.listdir(month_path)
        folders = [i for i in items if os.path.isdir(os.path.join(month_path, i))]
        files = [i for i in items if os.path.isfile(os.path.join(month_path, i))]

        print(f"         폴더: {len(folders)}개 / 파일: {len(files)}개")

        if folders:
            print(f"\n         📁 하위 폴더 (일별):")
            for f in sorted(folders)[:10]:
                print(f"            - {f}")
                if any(c.isdigit() for c in f):
                    result['has_date_in_path'] = True

        if files:
            print(f"\n         📄 파일 샘플:")
            for f in sorted(files)[:5]:
                print(f"            - {f}")

    except Exception as e:
        print(f"         ❌ 오류: {e}")


def main():
    """메인 실행"""
    print("=" * 80)
    print("📊 디렉토리 구조 확인 스크립트")
    print(f"   실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)

    # 설정 파일 로드
    config = load_config()
    if not config:
        return

    results = []

    # 1. 장비 디렉토리 확인
    print("\n" + "=" * 80)
    print("1️⃣ 장비 디렉토리")
    print("=" * 80)

    for equipment_id, equipment_info in config['equipment'].items():
        result = check_single_directory(
            f"{equipment_info['name']} ({equipment_id})",
            equipment_info['path']
        )
        if result:
            result['equipment_id'] = equipment_id
            results.append(result)

    # 2. 특별 항목 디렉토리 확인 (안저)
    print("\n" + "=" * 80)
    print("2️⃣ 특별 항목 디렉토리")
    print("=" * 80)

    if '안저' in config['special_items']:
        for idx, folder in enumerate(config['special_items']['안저']['folders'], 1):
            result = check_single_directory(f"안저 경로 {idx}", folder)
            if result:
                results.append(result)

    # 3. 결과 요약
    print("\n" + "=" * 80)
    print("📋 분석 결과 요약")
    print("=" * 80)

    print("\n경로 접근성:")
    for r in results:
        status = "✅ 접근 가능" if r['accessible'] else "❌ 접근 불가"
        print(f"   {r['name']:30s}: {status}")

    print("\n폴더 구조:")
    for r in results:
        if r['accessible']:
            structure = r.get('structure', '단일 폴더')
            print(f"   {r['name']:30s}: {structure}")

    print("\n날짜 패턴:")
    for r in results:
        if r['accessible']:
            path_date = "✅" if r.get('has_date_in_path') else "❌"
            file_date = "✅" if r.get('has_date_in_filename') else "❌"
            print(f"   {r['name']:30s}: 경로={path_date} / 파일명={file_date}")

    # 4. 최적화 권장사항
    print("\n" + "=" * 80)
    print("💡 최적화 권장사항")
    print("=" * 80)

    for r in results:
        if not r['accessible']:
            continue

        print(f"\n{r['name']}:")
        if r.get('has_date_in_path'):
            print(f"   → 백업 스크립트 방식 사용 가능 (경로에 날짜 포함)")
            print(f"   → getctime() 호출 불필요, 매우 빠름")
        elif r.get('has_date_in_filename'):
            print(f"   → 파일명 날짜 필터링 가능")
            print(f"   → getctime() 호출 최소화")
        else:
            print(f"   → 캐시 시스템 필요 (경로/파일명에 날짜 없음)")
            print(f"   → 첫 실행 후 캐시로 빠름")

    # 5. 결과 저장
    output_file = f"directory_structure_{date.today().strftime('%Y%m%d')}.json"
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2, default=str)
        print(f"\n✅ 상세 결과 저장됨: {output_file}")
    except Exception as e:
        print(f"\n⚠️  결과 저장 실패: {e}")

    print("\n" + "=" * 80)
    print("🎉 분석 완료")
    print("=" * 80)


if __name__ == "__main__":
    main()
