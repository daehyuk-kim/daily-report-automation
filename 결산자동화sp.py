import os
import datetime
import re

# 디렉토리 경로 설정
directory_path = r"\\192.168.0.120\sp"  # 네트워크 경로
output_file = "C:\\sp_output_file.txt"  # 결과 저장 경로

def filter_files(file_list):
    """
    파일 리스트를 조건에 맞게 필터링
    - 오늘 날짜 포함
    - 특정 패턴에 일치
    """
    today = datetime.date.today()
    filtered_files = []
    pattern = re.compile(r"(.+ \d+)_")  # 이름과 숫자 조합 추출

    for file_name in file_list:
        # 파일 이름에서 날짜 추출 (예: '파일명_20250117.txt')
        match_date = re.search(r"_(\d{8})\.", file_name)
        if match_date:
            file_date = datetime.datetime.strptime(match_date.group(1), "%Y%m%d").date()
            if file_date != today:  # 오늘 날짜가 아니면 제외
                continue
        
        # 파일 이름이 패턴에 일치하지 않으면 제외
        match_identifier = pattern.match(file_name)
        if not match_identifier:
            continue
        
        filtered_files.append(file_name)
    
    return filtered_files

def remove_duplicates(file_list):
    """숫자를 기반으로 중복 제거"""
    unique_identifiers = set()
    unique_files = []

    for file_name in file_list:
        match = re.match(r"(.+ \d+)_", file_name)
        if match:
            identifier = match.group(1)
            if identifier not in unique_identifiers:
                unique_identifiers.add(identifier)
                unique_files.append(file_name)
    
    return unique_files

# 파일 리스트 가져오기
try:
    all_files = os.listdir(directory_path)
except Exception as e:
    print(f"디렉토리 접근 오류: {e}")
    all_files = []

# 파일 필터링
filtered_files = filter_files(all_files)

# 중복 제거
unique_files = remove_duplicates(filtered_files)

# 결과 저장
try:
    with open(output_file, "w", encoding="utf-8") as file:
        file.write(f"디렉토리: {directory_path}\n오늘 생성된 파일 (중복 제거):\n")
        for item in unique_files:
            file.write(f"  {item}\n")
        file.write(f"총 {len(unique_files)}개\n\n")
    print(f"결과가 '{output_file}'에 저장되었습니다.")
except Exception as e:
    print(f"결과 파일 저장 중 오류 발생: {e}")
