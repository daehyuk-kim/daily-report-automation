#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
빠른 폴더 구조 확인 도구 (샘플링 모드)
1000개 정도만 확인해서 폴더 구조 파악
"""

import os
import sys
from datetime import datetime, date

def quick_analyze(directory_path, max_sample=1000):
    """
    디렉토리를 빠르게 샘플링하여 구조 파악

    Args:
        directory_path: 분석할 경로
        max_sample: 최대 샘플링 개수
    """
    today = date.today()

    print("=" * 80)
    print(f"빠른 폴더 구조 확인 (최대 {max_sample}개 샘플)")
    print("=" * 80)
    print(f"경로: {directory_path}")
    print(f"오늘: {today.strftime('%Y-%m-%d')}")
    print()

    if not os.path.exists(directory_path):
        print(f"❌ 경로가 존재하지 않습니다!")
        return

    # 1단계: 최상위 폴더 확인
    print("📁 최상위 폴더 구조:")
    print("-" * 80)

    try:
        level1_items = os.listdir(directory_path)
        level1_dirs = [item for item in level1_items if os.path.isdir(os.path.join(directory_path, item))]
        level1_files = [item for item in level1_items if os.path.isfile(os.path.join(directory_path, item))]

        print(f"폴더: {len(level1_dirs)}개")
        print(f"파일: {len(level1_files)}개")
        print()

        # 폴더만 샘플링
        sample_dirs = level1_dirs[:20]  # 최대 20개만

        for dir_name in sample_dirs:
            dir_path = os.path.join(directory_path, dir_name)
            try:
                # 생성 날짜 확인
                ctime = os.path.getctime(dir_path)
                dir_date = date.fromtimestamp(ctime)
                is_today = dir_date == today

                # 하위 항목 개수
                sub_items = os.listdir(dir_path)
                sub_count = len(sub_items)

                marker = "🟢" if is_today else "  "
                print(f"{marker} {dir_name:30s} [{sub_count:4d}개] ({dir_date})")
            except:
                print(f"   {dir_name:30s} [접근 불가]")

        if len(level1_dirs) > 20:
            print(f"   ... 외 {len(level1_dirs) - 20}개")

    except Exception as e:
        print(f"❌ 오류: {e}")
        return

    print()
    print("=" * 80)
    print("📂 2단계 폴더 구조 (연도 폴더 탐색):")
    print("-" * 80)

    # 연도처럼 보이는 폴더 찾기
    year_folders = [d for d in level1_dirs if d.isdigit() and len(d) == 4]

    if year_folders:
        # 최신 연도 폴더 확인
        latest_year = max(year_folders)
        year_path = os.path.join(directory_path, latest_year)

        print(f"연도 폴더 발견: {', '.join(year_folders)}")
        print(f"최신 연도: {latest_year}")
        print()

        try:
            month_items = os.listdir(year_path)
            month_dirs = [item for item in month_items if os.path.isdir(os.path.join(year_path, item))]

            print(f"{latest_year} 폴더 내부:")

            # 샘플링
            for dir_name in month_dirs[:15]:
                dir_path = os.path.join(year_path, dir_name)
                try:
                    ctime = os.path.getctime(dir_path)
                    dir_date = date.fromtimestamp(ctime)
                    is_today = dir_date == today

                    sub_items = os.listdir(dir_path)
                    sub_count = len(sub_items)

                    marker = "🟢" if is_today else "  "
                    print(f"{marker} {latest_year}\\{dir_name:25s} [{sub_count:4d}개] ({dir_date})")
                except:
                    print(f"   {latest_year}\\{dir_name:25s} [접근 불가]")

            if len(month_dirs) > 15:
                print(f"   ... 외 {len(month_dirs) - 15}개")

        except Exception as e:
            print(f"❌ 오류: {e}")
    else:
        print("연도 폴더 없음 - 단일 폴더 구조인 것으로 보임")
        print()
        print("파일 샘플링 (처음 20개):")

        try:
            files_sample = level1_files[:20]
            for file_name in files_sample:
                file_path = os.path.join(directory_path, file_name)
                try:
                    ctime = os.path.getctime(file_path)
                    file_date = date.fromtimestamp(ctime)
                    is_today = file_date == today

                    marker = "🟢" if is_today else "  "
                    print(f"{marker} {file_name}")
                except:
                    print(f"   {file_name}")

            if len(level1_files) > 20:
                print(f"... 외 {len(level1_files) - 20}개")
        except Exception as e:
            print(f"❌ 오류: {e}")

    print()
    print("=" * 80)
    print("🔍 오늘 날짜 관련 폴더/파일 검색:")
    print("-" * 80)

    # 오늘 날짜 형식들
    today_patterns = [
        today.strftime('%Y%m%d'),      # 20250118
        today.strftime('%Y.%m.%d'),    # 2025.01.18
        today.strftime('%Y-%m-%d'),    # 2025-01-18
        today.strftime('%m.%d'),       # 01.18
        today.strftime('%m-%d'),       # 01-18
        today.strftime('%m\\%d'),      # 01\18
        today.strftime('%d'),          # 18
    ]

    print(f"검색 패턴: {', '.join(today_patterns[:4])}")
    print()

    found_today = []
    scanned = 0

    def search_today(path, depth=0, max_depth=3):
        nonlocal scanned, found_today

        if depth > max_depth or scanned >= max_sample:
            return

        try:
            items = os.listdir(path)

            for item in items:
                if scanned >= max_sample:
                    break

                scanned += 1
                item_path = os.path.join(path, item)

                # 이름에 오늘 날짜 패턴 포함 여부
                name_match = any(pattern in item for pattern in today_patterns)

                # 생성 날짜 확인
                try:
                    ctime = os.path.getctime(item_path)
                    item_date = date.fromtimestamp(ctime)
                    date_match = item_date == today
                except:
                    date_match = False

                if name_match or date_match:
                    rel_path = os.path.relpath(item_path, directory_path)
                    found_today.append((rel_path, name_match, date_match))

                # 폴더면 재귀
                if os.path.isdir(item_path) and depth < max_depth:
                    search_today(item_path, depth + 1, max_depth)

        except:
            pass

    search_today(directory_path)

    if found_today:
        print(f"발견: {len(found_today)}개 (샘플링: {scanned}개 중)")
        print()

        for rel_path, name_match, date_match in found_today[:10]:
            reason = []
            if name_match:
                reason.append("이름매칭")
            if date_match:
                reason.append("날짜매칭")
            reason_str = "+".join(reason)

            print(f"🟢 {rel_path}")
            print(f"   ({reason_str})")
            print()

        if len(found_today) > 10:
            print(f"... 외 {len(found_today) - 10}개")
    else:
        print("오늘 날짜 관련 폴더/파일을 찾지 못했습니다.")

    print()
    print("=" * 80)
    print("✅ 분석 완료!")
    print(f"총 스캔: {scanned}개")
    print("=" * 80)


def main():
    """메인 함수"""
    print()
    print("🚀 빠른 폴더 구조 확인 도구")
    print()

    # 경로 입력
    if len(sys.argv) > 1:
        directory_path = sys.argv[1]
    else:
        print("분석할 디렉토리 경로를 입력하세요:")
        print()
        print("예시:")
        print("  Windows: D:\\BACKUP\\SP")
        print("  네트워크: \\\\192.168.0.120\\sp")
        print()
        directory_path = input("경로: ").strip().strip('"').strip("'")

    if not directory_path:
        print("경로가 입력되지 않았습니다.")
        return

    # 샘플링 개수
    print()
    print("최대 샘플링 개수 (기본값: 1000, 빠른 분석: 100)")
    sample_input = input("개수 (엔터: 1000): ").strip()
    max_sample = int(sample_input) if sample_input.isdigit() else 1000

    print()

    # 분석 실행
    quick_analyze(directory_path, max_sample)

    # 결과 저장
    print()
    save = input("결과를 파일로 저장하시겠습니까? (y/n): ").strip().lower()
    if save == 'y':
        output_file = input("파일명 (기본값: quick_analysis.txt): ").strip()
        if not output_file:
            output_file = "quick_analysis.txt"

        try:
            # 다시 실행해서 파일로 저장
            import io
            from contextlib import redirect_stdout

            f = io.StringIO()
            with redirect_stdout(f):
                quick_analyze(directory_path, max_sample)

            with open(output_file, 'w', encoding='utf-8') as file:
                file.write(f.getvalue())

            print(f"✅ 저장 완료: {output_file}")
        except Exception as e:
            print(f"❌ 저장 실패: {e}")


if __name__ == "__main__":
    main()
