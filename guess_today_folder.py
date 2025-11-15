#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
오늘 날짜 폴더 경로 추측 도구
스캔 없이 가능한 경로 패턴을 생성해서 존재 여부만 확인 (초고속)
"""

import os
import sys
from datetime import date

def guess_today_folders(base_path):
    """
    오늘 날짜 폴더 경로를 추측하여 존재 여부 확인
    스캔 없이 경로만 확인하므로 1초 이내 완료
    """
    today = date.today()

    print("=" * 80)
    print("오늘 날짜 폴더 경로 추측 (스캔 없음)")
    print("=" * 80)
    print(f"기본 경로: {base_path}")
    print(f"오늘: {today.strftime('%Y-%m-%d (%A)')}")
    print()

    # 가능한 폴더 구조 패턴 생성
    patterns = []

    year = today.strftime('%Y')
    month = today.strftime('%m')
    day = today.strftime('%d')
    month_no_zero = str(int(month))  # 01 -> 1
    day_no_zero = str(int(day))      # 08 -> 8

    # 패턴 1: YYYY\MM\DD
    patterns.append(('YYYY\\MM\\DD', os.path.join(base_path, year, month, day)))

    # 패턴 2: YYYY\MM\DD.MM
    patterns.append(('YYYY\\MM\\DD.MM', os.path.join(base_path, year, month, f"{day}.{month}")))

    # 패턴 3: YYYY\MM (월 폴더)
    patterns.append(('YYYY\\MM', os.path.join(base_path, year, month)))

    # 패턴 4: YYYY\YYYY.MM (년.월 폴더)
    patterns.append(('YYYY\\YYYY.MM', os.path.join(base_path, year, f"{year}.{month}")))

    # 패턴 5: YYYY\MM\TOPO MM.DD (장비명 포함)
    patterns.append(('YYYY\\MM\\TOPO MM.DD', os.path.join(base_path, year, month, f"TOPO {month}.{day}")))
    patterns.append(('YYYY\\MM\\TOPO M.D', os.path.join(base_path, year, month, f"TOPO {month_no_zero}.{day_no_zero}")))

    # 패턴 6: YYYY\MM\ORB MM.DD
    patterns.append(('YYYY\\MM\\ORB MM.DD', os.path.join(base_path, year, month, f"ORB {month}.{day}")))
    patterns.append(('YYYY\\MM\\ORB M.D', os.path.join(base_path, year, month, f"ORB {month_no_zero}.{day_no_zero}")))

    # 패턴 7: YYYY\MM\OCT MM.DD
    patterns.append(('YYYY\\MM\\OCT MM.DD', os.path.join(base_path, year, month, f"OCT {month}.{day}")))

    # 패턴 8: YYYY\MM\SP MM.DD
    patterns.append(('YYYY\\MM\\SP MM.DD', os.path.join(base_path, year, month, f"SP {month}.{day}")))

    # 패턴 9: YYYY\MM\HFA MM.DD
    patterns.append(('YYYY\\MM\\HFA MM.DD', os.path.join(base_path, year, month, f"HFA {month}.{day}")))

    # 패턴 10: YYYY\YYYY.MM\장비명 MM.DD
    patterns.append(('YYYY\\YYYY.MM\\ORB MM.DD', os.path.join(base_path, year, f"{year}.{month}", f"ORB {month}.{day}")))
    patterns.append(('YYYY\\YYYY.MM\\TOPO MM.DD', os.path.join(base_path, year, f"{year}.{month}", f"TOPO {month}.{day}")))

    # 패턴 11: MM.DD (단일 폴더 구조)
    patterns.append(('MM.DD', os.path.join(base_path, f"{month}.{day}")))

    # 패턴 12: YYYYMMDD
    patterns.append(('YYYYMMDD', os.path.join(base_path, today.strftime('%Y%m%d'))))

    # 패턴 13: 기본 경로 (단일 폴더에 파일 저장)
    patterns.append(('기본 경로', base_path))

    # 존재하는 경로 확인
    print("존재하는 폴더 확인 중...")
    print("-" * 80)

    found = []

    for pattern_name, path in patterns:
        try:
            if os.path.exists(path):
                # 하위 항목 개수 확인
                try:
                    items = os.listdir(path)
                    count = len(items)
                    found.append((pattern_name, path, count))
                    print(f"✅ {pattern_name:30s} → 존재! ({count}개)")
                except:
                    found.append((pattern_name, path, '?'))
                    print(f"✅ {pattern_name:30s} → 존재! (접근 불가)")
            else:
                print(f"❌ {pattern_name:30s}")
        except Exception as e:
            print(f"⚠️  {pattern_name:30s} (오류: {e})")

    print()
    print("=" * 80)

    if found:
        print(f"발견된 오늘 날짜 폴더: {len(found)}개")
        print("=" * 80)
        print()

        for pattern_name, path, count in found:
            print(f"📁 {pattern_name}")
            print(f"   경로: {path}")
            if count != '?':
                print(f"   항목: {count}개")
            print()

            # 샘플 파일 몇 개만 확인
            if count != '?' and count > 0:
                try:
                    items = os.listdir(path)
                    print("   샘플 (처음 5개):")
                    for item in items[:5]:
                        print(f"      - {item}")
                    if count > 5:
                        print(f"      ... 외 {count - 5}개")
                    print()
                except:
                    pass
    else:
        print("오늘 날짜 폴더를 찾지 못했습니다.")
        print()
        print("추가 확인 사항:")
        print("1. 기본 경로가 올바른지 확인")
        print("2. 네트워크 연결 상태 확인")
        print("3. 폴더 명명 규칙이 다를 수 있음")
        print()
        print("기본 경로의 하위 폴더 확인 (처음 10개):")
        try:
            items = os.listdir(base_path)
            dirs = [item for item in items if os.path.isdir(os.path.join(base_path, item))]
            for d in dirs[:10]:
                print(f"   - {d}")
            if len(dirs) > 10:
                print(f"   ... 외 {len(dirs) - 10}개")
        except Exception as e:
            print(f"   ❌ 오류: {e}")

    print("=" * 80)
    print("✅ 완료!")
    print("=" * 80)

    return found


def main():
    """메인 함수"""
    print()
    print("🚀 오늘 날짜 폴더 추측 도구 (초고속)")
    print("   스캔 없이 경로 패턴만 확인 → 1초 이내 완료")
    print()

    # 경로 입력
    if len(sys.argv) > 1:
        base_path = sys.argv[1]
    else:
        print("장비 백업 폴더 경로를 입력하세요:")
        print()
        print("예시:")
        print("  SP:    D:\\BACKUP\\SP")
        print("  TOPO:  \\\\topo\\topo")
        print("  ORB:   \\\\192.168.0.120\\orb")
        print("  OCT:   D:\\BACKUP\\OCT")
        print()
        base_path = input("경로: ").strip().strip('"').strip("'")

    if not base_path:
        print("경로가 입력되지 않았습니다.")
        return

    print()

    # 경로 추측
    found = guess_today_folders(base_path)

    # 결과 저장
    if found:
        print()
        save = input("발견된 경로를 파일로 저장하시겠습니까? (y/n): ").strip().lower()
        if save == 'y':
            output_file = "today_folders.txt"
            try:
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(f"오늘 날짜 폴더 경로 ({date.today()})\n")
                    f.write("=" * 80 + "\n\n")
                    f.write(f"기본 경로: {base_path}\n\n")
                    for pattern_name, path, count in found:
                        f.write(f"{pattern_name}\n")
                        f.write(f"  {path}\n")
                        if count != '?':
                            f.write(f"  {count}개 항목\n")
                        f.write("\n")

                print(f"✅ 저장 완료: {output_file}")
            except Exception as e:
                print(f"❌ 저장 실패: {e}")


if __name__ == "__main__":
    main()
