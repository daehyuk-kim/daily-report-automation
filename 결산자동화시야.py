import os
import datetime

# 네트워크 경로
base_directory_path = r"\\GEOMSA-MAIN2\hfa"
# 결과 저장 파일 경로
output_file = r"C:\hfa_output_file.txt"

def is_created_today(path):
    """해당 경로가 오늘 생성되었는지 여부를 반환"""
    try:
        creation_time = datetime.datetime.fromtimestamp(os.path.getctime(path))
        today = datetime.datetime.today()
        return creation_time.date() == today.date()
    except Exception as e:
        print(f"[ERROR] 생성일 확인 실패: {e}")
        return False

def count_folders_created_today(base_path):
    """오늘 생성된 폴더 개수를 반환"""
    folder_count = 0

    try:
        items = os.listdir(base_path)
        print(f"[DEBUG] 전체 항목 수: {len(items)}")

        for item in items:
            item_path = os.path.join(base_path, item)
            if os.path.isdir(item_path) and is_created_today(item_path):
                folder_count += 1
                print(f"[DEBUG] 오늘 생성된 폴더: {item}")
    except Exception as e:
        error_message = f"디렉토리 접근 오류: {e}"
        print(error_message)
        with open(output_file, "a", encoding="utf-8") as file:
            file.write(f"{error_message}\n")

    return folder_count

# 오늘 생성된 폴더 수 계산
count = count_folders_created_today(base_directory_path)

# 결과 출력 및 기록
print(f"오늘 생성된 폴더 수: {count}개")
try:
    with open(output_file, "a", encoding="utf-8") as file:
        file.write(f"\n[{datetime.datetime.now()}] 디렉토리: {base_directory_path}\n")
        file.write(f"오늘 생성된 폴더 수: {count}개\n")
    print(f"결과가 '{output_file}'에 저장되었습니다.")
except Exception as e:
    print(f"결과 파일 저장 중 오류 발생: {e}")
