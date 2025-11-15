# -*- coding: utf-8 -*-
import os
import shutil
import datetime
import ctypes
import re

# 공유 폴더 ↔ 백업 폴더 매핑
backup_pairs = {
    r"\\192.168.0.120\sp": r"D:\BACKUP\SP",
    r"\\GEOMSA-MAIN2\hfa": r"D:\BACKUP\hfa",
    r"\\topo\topo": r"D:\BACKUP\TOPO",
    r"\\192.168.0.120\iol700": r"D:\BACKUP\iol700",
    r"\\192.168.0.221\pdf": r"D:\BACKUP\OCT\2022.06~",
    r"\\192.168.0.228\oqass": r"D:\BACKUP\Oqas",
    r"\\192.168.0.209\orb": r"D:\BACKUP\ORB",
    r"\\192.168.0.231\slitlamp": r"D:\BACKUP\Slitlamp",
    r"\\antoct\antoct": r"D:\BACKUP\Antoct"
}

# 팝업 메시지 함수
def popup(title, message):
    try:
        ctypes.windll.user32.MessageBoxW(0, message, title, 0)
    except Exception as e:
        print(f"[팝업 실패] {title}: {message} → {e}")

# 오늘 날짜 기준 최근 2일 MM.DD 형식 생성
today = datetime.date.today()
recent_dates = {
    today.strftime("%m.%d"),
    (today - datetime.timedelta(days=1)).strftime("%m.%d")
}
print(f"📅 백업 대상 날짜: {', '.join(recent_dates)} (폴더 경로 기준)")

success_count = 0
fail_count = 0
targets = []

# 경로에서 날짜 문자열 포함 여부 확인
def path_matches_recent_date(path, date_set):
    for date_str in date_set:
        if date_str in path:
            return True
    return False

# 백업 대상 탐색
for source_dir, backup_root in backup_pairs.items():
    print(f"\n🔍 [{source_dir}] 파일 탐색 중...")

    try:
        for root, _, files in os.walk(source_dir):
            for filename in files:
                source_path = os.path.join(root, filename)
                if path_matches_recent_date(source_path, recent_dates):
                    rel_path = os.path.relpath(source_path, source_dir)
                    backup_path = os.path.join(backup_root, rel_path)

                    if not os.path.exists(backup_path):
                        targets.append((source_path, backup_path))
    except Exception as e:
        print(f"❌ 폴더 접근 실패: {source_dir} → {e}")
        continue

# 복사 시작
total = len(targets)
print(f"\n📁 백업 대상 파일: {total}개")

if total == 0:
    popup("백업 없음", "최근 2일 날짜가 포함된 경로의 파일이 없습니다.")
else:
    popup("백업 시작", f"{total}개 파일 백업을 시작합니다.")
    for idx, (src, dst) in enumerate(targets, start=1):
        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)
            success_count += 1
            print(f"[{idx}/{total}] ✅ 복사됨: {src}")
        except Exception as e:
            fail_count += 1
            print(f"[{idx}/{total}] ❌ 실패: {src} → {e}")

    # 결과 요약
    if fail_count == 0:
        popup("백업 완료", f"{success_count}개 파일 백업 완료!")
    else:
        popup("백업 일부 실패", f"{success_count}개 성공, {fail_count}개 실패.")
