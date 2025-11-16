#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
파일 캐시 관리자
대용량 디렉토리의 파일 목록을 캐싱하여 스캔 속도를 크게 향상

원리:
1. 어제까지의 모든 파일 목록을 캐시 파일에 저장
2. 오늘 스캔 시: 전체 파일 목록 - 캐시된 파일 = 오늘 추가된 파일
3. 오늘 추가된 파일만 getctime() 확인

효과:
- 10만개 파일 중 오늘 추가된 50개만 확인
- 99.95% I/O 감소
"""

import os
import json
from datetime import date, datetime
from pathlib import Path


CACHE_DIR = ".file_cache"
CACHE_VERSION = "1.0"


def get_cache_path(directory_path: str) -> str:
    """캐시 파일 경로 생성"""
    # 디렉토리 경로를 해시하여 캐시 파일명 생성
    import hashlib
    path_hash = hashlib.md5(directory_path.encode('utf-8')).hexdigest()[:12]
    clean_name = directory_path.replace('\\', '_').replace('/', '_').replace(':', '')
    clean_name = clean_name[-50:] if len(clean_name) > 50 else clean_name

    return os.path.join(CACHE_DIR, f"cache_{clean_name}_{path_hash}.json")


def load_cache(directory_path: str) -> dict:
    """캐시 파일 로드"""
    cache_path = get_cache_path(directory_path)

    if not os.path.exists(cache_path):
        return {'files': set(), 'last_updated': None, 'version': CACHE_VERSION}

    try:
        with open(cache_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            data['files'] = set(data.get('files', []))
            return data
    except:
        return {'files': set(), 'last_updated': None, 'version': CACHE_VERSION}


def save_cache(directory_path: str, files: set):
    """캐시 파일 저장"""
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache_path = get_cache_path(directory_path)

    data = {
        'directory': directory_path,
        'files': list(files),
        'count': len(files),
        'last_updated': datetime.now().isoformat(),
        'version': CACHE_VERSION
    }

    try:
        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"캐시 저장 오류: {e}")
        return False


def get_new_files(directory_path: str, current_files: list) -> list:
    """캐시와 비교하여 새로 추가된 파일만 반환"""
    cache = load_cache(directory_path)
    cached_files = cache['files']

    # 차집합: 현재 파일 - 캐시된 파일 = 새 파일
    current_set = set(current_files)
    new_files = current_set - cached_files

    return list(new_files)


def update_cache_with_today_files(directory_path: str, all_files: list):
    """오늘 파일을 제외한 나머지를 캐시에 저장"""
    cache = load_cache(directory_path)
    cached_files = cache['files']

    # 기존 캐시 + 모든 파일 (오늘 파일 제외는 호출자가 처리)
    all_files_set = set(all_files)
    updated_files = cached_files | all_files_set

    save_cache(directory_path, updated_files)
    return len(updated_files)


def clear_cache(directory_path: str = None):
    """캐시 삭제"""
    if directory_path:
        cache_path = get_cache_path(directory_path)
        if os.path.exists(cache_path):
            os.remove(cache_path)
            return True
    else:
        # 모든 캐시 삭제
        if os.path.exists(CACHE_DIR):
            import shutil
            shutil.rmtree(CACHE_DIR)
            return True
    return False


def get_cache_info():
    """모든 캐시 파일 정보"""
    if not os.path.exists(CACHE_DIR):
        return []

    info = []
    for cache_file in os.listdir(CACHE_DIR):
        if cache_file.endswith('.json'):
            cache_path = os.path.join(CACHE_DIR, cache_file)
            try:
                with open(cache_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    info.append({
                        'file': cache_file,
                        'directory': data.get('directory', 'Unknown'),
                        'count': data.get('count', 0),
                        'last_updated': data.get('last_updated', 'Unknown')
                    })
            except:
                pass

    return info


# 테스트
if __name__ == "__main__":
    print("=" * 80)
    print("파일 캐시 관리자 테스트")
    print("=" * 80)
    print()

    # 캐시 정보 표시
    info = get_cache_info()
    if info:
        print("📁 현재 캐시 파일:")
        for item in info:
            print(f"  - {item['directory']}")
            print(f"    파일 수: {item['count']:,}개")
            print(f"    마지막 업데이트: {item['last_updated']}")
            print()
    else:
        print("캐시 파일 없음")
        print()

    # 예시: 캐시 사용법
    print("사용법:")
    print("1. 첫 실행: 모든 파일 목록을 캐시에 저장")
    print("2. 다음 실행: 캐시와 비교하여 새 파일만 확인")
    print("3. 효과: 10만개 파일 → 오늘 추가된 50개만 확인")
    print()

    # 캐시 삭제 옵션
    if info:
        clear = input("모든 캐시를 삭제하시겠습니까? (y/n): ").strip().lower()
        if clear == 'y':
            clear_cache()
            print("✅ 모든 캐시 삭제됨")
