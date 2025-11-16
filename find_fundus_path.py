#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
안저 검사 경로 탐색 도구
\\192.168.0.213\ 하위의 실제 Fundus 폴더를 찾습니다.
"""

import os
import sys
from datetime import date

def explore_network_path(base_path, max_depth=3):
    """네트워크 경로 탐색"""
    print(f"\n{'='*80}")
    print(f"📁 경로 탐색: {base_path}")
    print(f"{'='*80}\n")

    if not os.path.exists(base_path):
        print(f"❌ 경로 접근 불가: {base_path}")
        print(f"   네트워크 연결을 확인하세요.")
        return []

    print(f"✅ 경로 접근 가능\n")

    found_paths = []

    try:
        # 1단계: 최상위 폴더 목록
        print("📂 최상위 폴더 목록:")
        print("-" * 80)

        top_level = []
        for item in os.listdir(base_path):
            item_path = os.path.join(base_path, item)
            if os.path.isdir(item_path):
                top_level.append(item)

        if not top_level:
            print("   (빈 폴더)\n")
            return []

        for i, folder in enumerate(sorted(top_level), 1):
            print(f"{i:2d}. {folder}")

        print()

        # 2단계: Fundus 관련 폴더 찾기
        print("🔍 'Fundus' 관련 폴더 검색:")
        print("-" * 80)

        fundus_keywords = ['fundus', 'afc', '안저', 'retina']

        for folder in top_level:
            folder_lower = folder.lower()
            if any(keyword in folder_lower for keyword in fundus_keywords):
                folder_path = os.path.join(base_path, folder)
                print(f"\n✓ 발견: {folder}")
                print(f"  전체 경로: {folder_path}")

                # 하위 폴더 확인
                try:
                    subfolders = [f for f in os.listdir(folder_path)
                                 if os.path.isdir(os.path.join(folder_path, f))]

                    if subfolders:
                        print(f"  하위 폴더 ({len(subfolders)}개):")
                        for subfolder in sorted(subfolders)[:10]:
                            print(f"    - {subfolder}")
                        if len(subfolders) > 10:
                            print(f"    ... 외 {len(subfolders) - 10}개")

                    # 파일 샘플
                    files = [f for f in os.listdir(folder_path)
                            if os.path.isfile(os.path.join(folder_path, f))]

                    if files:
                        print(f"  파일 샘플:")
                        for file in sorted(files)[:5]:
                            print(f"    - {file}")
                        if len(files) > 5:
                            print(f"    ... 외 {len(files) - 5}개")

                except Exception as e:
                    print(f"  ⚠️  하위 항목 읽기 오류: {e}")

                found_paths.append(folder_path)

        if not found_paths:
            print("   ❌ Fundus 관련 폴더를 찾지 못했습니다.")
            print()
            print("💡 수동으로 탐색하려면 위 폴더 목록을 참고하세요.")

        print()

    except PermissionError:
        print(f"❌ 권한 오류: {base_path}에 접근할 수 없습니다.")
    except Exception as e:
        print(f"❌ 오류: {e}")

    return found_paths


def main():
    """메인 함수"""
    print("\n" + "="*80)
    print("🔍 안저 검사 경로 탐색 도구")
    print("="*80)

    # 탐색할 경로들
    paths_to_check = [
        r"\\192.168.0.213",
        r"\\AFC-210-PC\Fundus2"  # 이미 알려진 경로도 확인
    ]

    all_found = []

    for path in paths_to_check:
        found = explore_network_path(path)
        all_found.extend(found)

    # 결과 요약
    print("\n" + "="*80)
    print("📋 탐색 결과 요약")
    print("="*80)
    print()

    if all_found:
        print("발견된 Fundus 경로:")
        for i, path in enumerate(all_found, 1):
            print(f"{i}. {path}")
    else:
        print("⚠️  자동으로 Fundus 경로를 찾지 못했습니다.")
        print()
        print("수동 확인이 필요합니다:")
        print("1. Windows 탐색기에서 \\\\192.168.0.213\\ 열기")
        print("2. Fundus 관련 폴더 찾기")
        print("3. 찾은 경로를 config_real.json에 업데이트")

    print()
    print("="*80)
    print()


if __name__ == "__main__":
    main()
