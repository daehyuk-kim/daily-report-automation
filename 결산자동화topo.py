import os
import re
import datetime

# 디렉토리 경로 설정
base_directory_path = r"\\topo\topo"  # 네트워크 경로 (topo로 변경)
output_file = "C:\\topo_output_file.txt"  # 결과 저장 경로 (topo로 변경)

def filter_files(file_list):
    """
    파일 리스트를 특정 패턴에 따라 필터링
    - 파일명에 숫자 1~6자리가 포함된 경우만 유지
    """
    filtered = []
    pattern = re.compile(r".* \d{1,6}-\d{2}")  # '허태용 107207-14' 형식의 파일명만 허용
    for file_name in file_list:
        if pattern.match(file_name):
            filtered.append(file_name)
    return filtered

def remove_duplicates(file_list):
    """파일명에서 숫자 부분과 다른 구분자까지 고려하여 중복 제거"""
    unique_files = []
    unique_identifiers = set()

    for file_name in file_list:
        # 파일명에서 '이름 숫자-번호' 형식으로 고유 식별자를 생성
        match = re.search(r"(.+ \d{1,6})-\d{2}", file_name)
        if match:
            identifier = match.group(1)  # 숫자 앞까지의 부분을 고유 식별자로 사용
            if identifier not in unique_identifiers:
                unique_identifiers.add(identifier)
                unique_files.append(file_name)
    
    return unique_files

def get_today_folder_path(base_directory_path):
    """오늘 날짜에 해당하는 폴더 경로 반환"""
    today = datetime.datetime.today()
    today_folder = today.strftime("%m\\TOPO %m.%d")  # '01\\TOPO 01.18' 형식으로 수정
    today_folder_path = os.path.join(base_directory_path, today.strftime("%Y") + "\\" + today_folder)
    
    if os.path.exists(today_folder_path):
        return today_folder_path
    else:
        print(f"오늘 날짜에 해당하는 폴더({today_folder_path})가 존재하지 않습니다.")
        return None

def get_all_files_from_directory(directory_path):
    """디렉토리와 하위 폴더의 모든 파일 리스트를 반환"""
    all_files = []
    try:
        for root, dirs, files in os.walk(directory_path):  # 하위 폴더를 포함하여 파일 목록을 가져옴
            for file in files:
                all_files.append(os.path.join(root, file))  # 파일 경로를 전체 경로로 추가
    except Exception as e:
        error_message = f"디렉토리 접근 오류: {e}"
        print(error_message)
        # 오류 메시지를 메모장에 기록
        with open(output_file, "a", encoding="utf-8") as file:
            file.write(f"{error_message}\n")
    
    return all_files

# 오늘 날짜 폴더 경로 찾기
today_folder_path = get_today_folder_path(base_directory_path)

if today_folder_path:
    # 파일 리스트 가져오기
    all_files = get_all_files_from_directory(today_folder_path)

    # 파일 필터링
    filtered_files = filter_files(all_files)
    print(f"필터링된 파일 수: {len(filtered_files)}개")
    print("필터링된 파일 목록:")
    for f in filtered_files:
        print(f"  - {f}")

    # 중복 제거
    unique_files = remove_duplicates(filtered_files)
    print(f"중복 제거 후 파일 수: {len(unique_files)}개")
    print("중복 제거된 파일 목록:")
    for f in unique_files:
        print(f"  - {f}")

    # 결과 저장
    try:
        with open(output_file, "a", encoding="utf-8") as file:
            file.write(f"\n디렉토리: {today_folder_path}\n조건에 맞는 파일 (중복 제거):\n")
            for item in unique_files:
                file.write(f"  {item}\n")
            file.write(f"총 {len(unique_files)}개\n\n")
        print(f"결과가 '{output_file}'에 저장되었습니다.")
    except Exception as e:
        error_message = f"결과 파일 저장 중 오류 발생: {e}"
        print(error_message)
        # 오류 메시지를 메모장에 기록
        with open(output_file, "a", encoding="utf-8") as file:
            file.write(f"{error_message}\n")
