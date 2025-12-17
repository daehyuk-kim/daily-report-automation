#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
일일결산 자동화 시스템 (최적화 버전)
안과 검사실의 일일 통계를 자동으로 수집하고 PDF 보고서를 생성하는 프로그램
"""

import os
import sys
import json
import re
import threading
from datetime import datetime, date, timedelta
from typing import Set, Dict, List, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

try:
    import openpyxl
    from openpyxl import load_workbook
except ImportError:
    print("오류: openpyxl이 설치되지 않았습니다.")
    print("설치: pip install openpyxl")
    sys.exit(1)

try:
    import pandas as pd
except ImportError:
    print("오류: pandas가 설치되지 않았습니다.")
    print("설치: pip install pandas")
    sys.exit(1)

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False

# 파일 캐시 시스템
try:
    from file_cache_manager import get_new_files, update_cache_with_today_files, load_cache
    HAS_CACHE = True
except ImportError:
    HAS_CACHE = False

# Windows에서만 pywin32 임포트
if sys.platform == 'win32':
    try:
        import win32com.client
        import pythoncom
        HAS_WIN32 = True
    except ImportError:
        HAS_WIN32 = False
        print("경고: pywin32가 설치되지 않았습니다. PDF 변환을 사용할 수 없습니다.")
else:
    HAS_WIN32 = False


class DailyReportSystem:
    """일일결산 시스템의 메인 클래스 (최적화 버전)"""

    def __init__(self, config_path: str = "config.json"):
        self.config = self.load_config(config_path)
        self.chart_numbers = {}
        self.results = {}
        self.today = date.today()

        # 정규식 패턴 미리 컴파일
        self.compiled_patterns = {}
        for eq_id, eq_info in self.config['equipment'].items():
            self.compiled_patterns[eq_id] = re.compile(eq_info['pattern'])

    def load_config(self, config_path: str) -> dict:
        """설정 파일 로드"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            messagebox.showerror("오류", f"설정 파일을 찾을 수 없습니다: {config_path}")
            sys.exit(1)
        except json.JSONDecodeError:
            messagebox.showerror("오류", "설정 파일 형식이 올바르지 않습니다.")
            sys.exit(1)

    def is_valid_chart_number(self, chart_num_str: str) -> bool:
        """차트번호 유효성 검증"""
        try:
            if chart_num_str.startswith('0') and len(chart_num_str) > 1:
                return False
            chart_num = int(chart_num_str)
            min_val = self.config['validation']['chart_number_min']
            max_val = self.config['validation']['chart_number_max']
            return min_val <= chart_num <= max_val
        except (ValueError, KeyError):
            return False

    def extract_chart_number(self, match) -> Optional[str]:
        """정규식 매칭에서 차트번호 추출 (단일/이중 그룹 패턴 지원)

        단일 그룹 패턴 (SP, TOPO 등): (\d+)_
        이중 그룹 패턴 (HFA): _(\d{5,6})$|^(\d{5,6})_
        """
        if not match:
            return None
        return match.group(1) or (match.group(2) if match.lastindex > 1 else None)

    def get_today_folder_path(self, base_path: str, equipment_id: str) -> Optional[str]:
        """오늘 날짜 폴더 경로 생성 (config의 folder_structure 사용)"""
        today = self.today

        # config에서 folder_structure 가져오기
        if equipment_id not in self.config['equipment']:
            return base_path

        equipment = self.config['equipment'][equipment_id]
        if 'folder_structure' not in equipment:
            return base_path

        # folder_structure 형식을 실제 경로로 변환
        # YYYY\MM\MM.DD -> 2025\11\11.17
        # YYYY\MM\TOPO MM.DD -> 2025\11\TOPO 11.17
        # YYYY\YYYY.MM\ORB MM.DD -> 2025\2025.11\ORB 11.17
        # YYYY\MM\oct MM.DD -> 2025\11\oct 11.17

        folder_structure = equipment['folder_structure']

        # 날짜 변환 (순서 중요: 긴 패턴부터 변환)
        folder = folder_structure
        folder = folder.replace('YYYY.MM', today.strftime('%Y.%m'))
        folder = folder.replace('YYYY', today.strftime('%Y'))
        folder = folder.replace('MM.DD', today.strftime('%m.%d'))
        folder = folder.replace('MM', today.strftime('%m'))
        folder = folder.replace('DD', today.strftime('%d'))

        full_path = os.path.join(base_path, folder)

        if os.path.exists(full_path):
            return full_path
        else:
            return None

    def scan_directory_fast(self, equipment_id: str, log_callback) -> Set[str]:
        """
        장비 디렉토리 스캔 (최적화 버전)
        - 오늘 날짜 폴더만 스캔
        - os.walk() 사용
        - 정규식 미리 컴파일
        - 경로에 날짜 포함 여부로 빠른 필터링
        """
        equipment = self.config['equipment'][equipment_id]
        base_path = equipment['path']
        pattern = self.compiled_patterns[equipment_id]
        scan_type = equipment['scan_type']

        chart_numbers = set()

        # 오늘 날짜 패턴들 (경로/파일명 매칭용) - 백업 스크립트 방식
        today_patterns = [
            self.today.strftime('%m.%d'),     # 11.16
            self.today.strftime('%Y%m%d'),    # 20251116
            self.today.strftime('%Y-%m-%d'),  # 2025-11-16
            self.today.strftime('%Y.%m.%d'),  # 2025.11.16
        ]

        if not os.path.exists(base_path):
            log_callback(f"  ⚠️  경로 없음: {base_path}")
            return chart_numbers

        try:
            # 오늘 날짜 폴더 경로 찾기
            today_folder = self.get_today_folder_path(base_path, equipment_id)
            is_realtime_scan = False  # 기본값

            if today_folder is None:
                # 날짜 폴더가 없는 경우: base_path를 직접 스캔
                # SP, HFA, Fundus 등은 낮에는 최상위 폴더에 직접 저장, 저녁에 날짜 폴더로 이동
                # 날짜 폴더가 없으면 최상위에 있는 것들이 오늘 것임
                today_folder = base_path
                is_realtime_scan = True  # 실시간 스캔 표시
                use_creation_time = equipment.get('use_creation_time', False)
                log_callback(f"     📂 스캔 경로: {today_folder} (날짜 폴더 미정리 - 최상위 전체 스캔)")
                if use_creation_time:
                    log_callback(f"     🔍 생성일 확인 모드")

                # 단일 폴더만 스캔 - os.scandir() 사용 (stat 캐싱으로 더 빠름)
                if scan_type == 'file':
                    log_callback(f"     ⚡ os.scandir() 사용 (stat 캐싱)")

                    valid_extensions = self.config['validation']['file_extensions']
                    total_files = 0
                    candidate_entries = []

                    # os.scandir()은 DirEntry 객체를 반환 (stat 정보 캐싱됨)
                    try:
                        with os.scandir(today_folder) as entries:
                            for entry in entries:
                                total_files += 1
                                if entry.is_file(follow_symlinks=False):
                                    if any(entry.name.lower().endswith(ext) for ext in valid_extensions):
                                        candidate_entries.append(entry)
                    except Exception as e:
                        log_callback(f"     ❌ 스캔 오류: {e}")
                        return chart_numbers

                    log_callback(f"     📊 전체: {total_files}개 / 유효 확장자: {len(candidate_entries)}개")

                    if not candidate_entries:
                        log_callback(f"     ⚠️  유효한 파일 없음")
                        return chart_numbers

                    # 최적화 1: 날짜 폴더 미정리 시 모든 파일을 오늘 것으로 간주
                    if is_realtime_scan:
                        log_callback(f"     🔍 실시간 스캔 모드: 모든 파일 매칭")
                        for entry in candidate_entries:
                            match = pattern.search(entry.name)
                            if match:
                                chart_num = self.extract_chart_number(match)
                                if self.is_valid_chart_number(chart_num):
                                    chart_numbers.add(chart_num)
                        log_callback(f"     ✅ 매칭 완료: {len(chart_numbers)}건")
                    else:
                        # 날짜 폴더가 있는 경우: 파일명/경로에 날짜 확인
                        filename_matched = 0
                        need_ctime_check = []

                        for entry in candidate_entries:
                            # 파일명 또는 전체 경로에 오늘 날짜가 있으면 바로 처리
                            if any(dp in entry.path for dp in today_patterns):
                                filename_matched += 1
                                match = pattern.search(entry.name)
                                if match:
                                    chart_num = self.extract_chart_number(match)
                                    if self.is_valid_chart_number(chart_num):
                                        chart_numbers.add(chart_num)
                            elif use_creation_time:
                                need_ctime_check.append(entry)

                        if filename_matched > 0:
                            log_callback(f"     ⚡ 파일명/경로 날짜 매칭: {filename_matched}개 → {len(chart_numbers)}건")

                        # 최적화 2: 생성일 확인이 필요한 경우 (파일명에 날짜 없음)
                        if need_ctime_check and use_creation_time:
                            log_callback(f"     🔍 생성일 확인 필요: {len(need_ctime_check)}개")

                            # 캐시 시스템 사용 (가장 빠름)
                            if HAS_CACHE:
                                cache = load_cache(today_folder)
                                if cache['last_updated']:
                                    log_callback(f"     ⚡ 캐시 사용: 마지막 업데이트 {cache['last_updated'][:10]}")
                                    entry_names = [e.name for e in need_ctime_check]
                                    new_file_names = get_new_files(today_folder, entry_names)
                                    new_file_set = set(new_file_names)
                                    need_ctime_check = [e for e in need_ctime_check if e.name in new_file_set]
                                    log_callback(f"     📊 캐시에 없는 새 파일: {len(need_ctime_check)}개 (기존 {len(entry_names) - len(need_ctime_check)}개 스킵)")

                                    if not need_ctime_check:
                                        log_callback(f"     ✅ 새 파일 없음 - 캐시에서 모두 확인됨")
                                        # 캐시 업데이트
                                        update_cache_with_today_files(today_folder, [e.name for e in candidate_entries])
                                        return chart_numbers
                                else:
                                    log_callback(f"     💾 캐시 없음 - 첫 실행 (다음부터 빨라짐)")

                            log_callback(f"     ⚡ os.scandir() stat 캐싱 사용 (getctime보다 10배 빠름)")

                            # DirEntry.stat()은 캐싱됨 - 네트워크 호출 최소화
                            def check_entry_date(entry):
                                try:
                                    # entry.stat()은 캐싱되어 있어 매우 빠름
                                    stat_info = entry.stat(follow_symlinks=False)
                                    ctime = stat_info.st_ctime
                                    file_date = date.fromtimestamp(ctime)
                                    if file_date == self.today:
                                        match = pattern.search(entry.name)
                                        if match:
                                            chart_num = self.extract_chart_number(match)
                                            if self.is_valid_chart_number(chart_num):
                                                return chart_num, file_date
                                    return None, file_date
                                except:
                                    pass
                                return None, None

                            # 배치 처리 (1000개씩) - entry.stat()은 캐싱되어 병렬 불필요
                            batch_size = 1000
                            total_checked = 0
                            consecutive_old_files = 0
                            ctime_matches = 0

                            for i in range(0, len(need_ctime_check), batch_size):
                                batch = need_ctime_check[i:i+batch_size]

                                # 순차 처리 (entry.stat()은 이미 캐싱됨, 병렬보다 오버헤드 적음)
                                batch_old_count = 0
                                for entry in batch:
                                    chart_num, file_date = check_entry_date(entry)
                                    if chart_num:
                                        chart_numbers.add(chart_num)
                                        ctime_matches += 1
                                        consecutive_old_files = 0
                                    elif file_date and file_date < self.today:
                                        batch_old_count += 1

                                # 이 배치에서 대부분 오래된 파일이면
                                if batch_old_count > len(batch) * 0.9:
                                    consecutive_old_files += 1

                                total_checked += len(batch)

                                # 진행 상황 로그
                                if total_checked % 2000 == 0 or i + batch_size >= len(need_ctime_check):
                                    log_callback(f"        ... {total_checked}/{len(need_ctime_check)} 확인 ({ctime_matches}건 발견)")

                                # 조기 종료: 연속 3배치가 모두 오래된 파일이면 중단
                                if consecutive_old_files >= 3:
                                    log_callback(f"     ⏹️  조기 종료: 최근 파일 없음 (총 {total_checked}개 확인)")
                                    break

                            log_callback(f"     ✅ 생성일 확인 완료: {ctime_matches}건 추가")

                            # 캐시 업데이트: 오늘 파일 제외한 모든 파일 저장
                            if HAS_CACHE:
                                # 오늘 생성된 파일을 제외한 나머지를 캐시에 추가
                                old_files = [e.name for e in candidate_entries if e.name not in chart_numbers]
                                update_cache_with_today_files(today_folder, old_files)
                                log_callback(f"     💾 캐시 업데이트 완료")

                        log_callback(f"     📊 최종 결과: {len(chart_numbers)}건 (중복 제외)")
                    return chart_numbers
                # scan_type == 'file'이 아닐 때는 아래 일반 스캔 로직으로 계속 진행

            # 오늘 폴더와 하위 폴더만 스캔 (os.walk 사용)
            log_callback(f"     📂 스캔 경로: {today_folder}")

            # 날짜 폴더가 없고 base_path를 스캔하는 경우 (실시간 파일/폴더)
            is_realtime_scan = (today_folder == base_path)

            total_files_count = 0
            total_dirs_count = 0

            # 디버그: scan_type과 is_realtime_scan 값 확인
            log_callback(f"     🔧 DEBUG: scan_type='{scan_type}', is_realtime_scan={is_realtime_scan}")

            # scan_type == 'both'이고 날짜 폴더 없을 때: 최상위 폴더 전체 스캔
            if scan_type == 'both' and is_realtime_scan:
                log_callback(f"     🔍 최상위 폴더 스캔 (정리 전)")

                try:
                    items = os.listdir(today_folder)
                    for item in items:
                        item_path = os.path.join(today_folder, item)

                        if os.path.isdir(item_path):
                            total_dirs_count += 1

                            # 패턴 매칭 (생성일 확인 없이)
                            match = pattern.search(item)
                            if match:
                                chart_num = self.extract_chart_number(match)
                                if self.is_valid_chart_number(chart_num):
                                    chart_numbers.add(chart_num)

                    log_callback(f"     📊 폴더: {total_dirs_count}개 / 매칭: {len(chart_numbers)}건")
                except Exception as e:
                    log_callback(f"     ❌ 스캔 오류: {e}")

            else:
                # 일반 스캔 (날짜 폴더가 있는 경우)
                for root, dirs, files in os.walk(today_folder):
                    total_files_count += len(files)
                    total_dirs_count += len(dirs)

                    # 파일 스캔
                    if scan_type in ['file', 'both']:
                        for file_name in files:
                            # 확장자 체크
                            if not any(file_name.lower().endswith(ext) for ext in self.config['validation']['file_extensions']):
                                continue

                            # 차트번호 추출
                            match = pattern.search(file_name)
                            if match:
                                chart_num = self.extract_chart_number(match)
                                if self.is_valid_chart_number(chart_num):
                                    chart_numbers.add(chart_num)

                    # 폴더 스캔 (OCT, HFA 등)
                    if scan_type == 'both':
                        for dir_name in dirs:
                            match = pattern.search(dir_name)
                            if match:
                                chart_num = self.extract_chart_number(match)
                                if self.is_valid_chart_number(chart_num):
                                    chart_numbers.add(chart_num)

            # 로그 출력 (실시간 스캔은 위에서 이미 출력)
            if not (scan_type == 'both' and is_realtime_scan):
                if scan_type == 'both':
                    log_callback(f"     📊 파일: {total_files_count}개 / 폴더: {total_dirs_count}개 / 매칭: {len(chart_numbers)}건")
                else:
                    log_callback(f"     📊 파일: {total_files_count}개 / 매칭: {len(chart_numbers)}건")

        except Exception as e:
            log_callback(f"  ❌ 오류: {equipment['name']} - {str(e)}")

        return chart_numbers

    def calculate_glaucoma(self, log_callback) -> int:
        """녹내장 계산 (HFA ∩ OCT)"""
        try:
            hfa_charts = self.chart_numbers.get('HFA', set())
            oct_charts = self.chart_numbers.get('OCT', set())
            glaucoma_charts = hfa_charts & oct_charts
            return len(glaucoma_charts)
        except Exception as e:
            log_callback(f"  ❌ 녹내장 계산 오류: {str(e)}")
            return 0

    def calculate_fundus(self, log_callback) -> int:
        """안저 계산 (Fundus + Secondary 폴더) - 최적화 버전"""
        fundus_charts = set()

        # 오늘 날짜 패턴
        today_str = self.today.strftime('%Y%m%d')
        today_str_dash = self.today.strftime('%Y-%m-%d')
        today_str_dot = self.today.strftime('%Y.%m.%d')
        date_patterns = [today_str, today_str_dash, today_str_dot]

        try:
            fundus_config = self.config['special_items']['안저']['folders']

            # 1. Fundus 폴더 처리 (날짜별 폴더 구조)
            if 'fundus' in fundus_config:
                fundus_info = fundus_config['fundus']
                base_path = fundus_info['path']
                pattern = re.compile(fundus_info['pattern'])

                log_callback(f"  📂 Fundus 스캔: {base_path}")

                if os.path.exists(base_path):
                    # 오늘 날짜 폴더 경로 생성
                    folder_structure = fundus_info.get('folder_structure', '')
                    today_folder = None

                    if folder_structure:
                        folder = folder_structure
                        folder = folder.replace('YYYY.MM', self.today.strftime('%Y.%m'))
                        folder = folder.replace('YYYY', self.today.strftime('%Y'))
                        folder = folder.replace('MM.DD', self.today.strftime('%m.%d'))
                        folder = folder.replace('MM', self.today.strftime('%m'))
                        folder = folder.replace('DD', self.today.strftime('%d'))
                        today_folder = os.path.join(base_path, folder)

                    # 1) 날짜 폴더가 있으면 우선 스캔 (저녁 정리 후)
                    if today_folder and os.path.exists(today_folder):
                        log_callback(f"     📂 날짜 폴더: {today_folder}")
                        items = os.listdir(today_folder)
                        log_callback(f"     전체: {len(items)}개")

                        for item in items:
                            match = pattern.search(item)
                            if match:
                                chart_num = self.extract_chart_number(match)
                                if self.is_valid_chart_number(chart_num):
                                    fundus_charts.add(chart_num)

                        log_callback(f"     ✅ 날짜 폴더 매칭: {len(fundus_charts)}건")

                    # 2) 날짜 폴더가 없으면 base_path 스캔 (정리 전 파일)
                    # 매일 저녁 100% 정리하므로 최상위에 있는 것 = 오늘 것
                    if not today_folder or not os.path.exists(today_folder):
                        log_callback(f"     📂 최상위 경로 스캔: {base_path} (정리 전)")

                        try:
                            items = os.listdir(base_path)
                            # 하위 폴더 제외, 파일만
                            files = [f for f in items if os.path.isfile(os.path.join(base_path, f))]
                            log_callback(f"     전체 파일: {len(files)}개")

                            base_fundus_charts = set()
                            valid_extensions = self.config['validation']['file_extensions']

                            for file_name in files:
                                # 확장자 체크
                                if not any(file_name.lower().endswith(ext) for ext in valid_extensions):
                                    continue

                                # 패턴 매칭 (생성일 확인 없이)
                                match = pattern.search(file_name)
                                if match:
                                    chart_num = self.extract_chart_number(match)
                                    if self.is_valid_chart_number(chart_num):
                                        base_fundus_charts.add(chart_num)

                            if base_fundus_charts:
                                log_callback(f"     ✅ 최상위 파일 매칭: {len(base_fundus_charts)}건")
                                fundus_charts.update(base_fundus_charts)
                            else:
                                log_callback(f"     ⚠️  매칭된 파일 없음")
                        except Exception as e:
                            log_callback(f"     ❌ 최상위 경로 스캔 오류: {e}")
                else:
                    log_callback(f"  ⚠️  경로 없음: {base_path}")

            # 2. Secondary 폴더 처리 (파일명에 날짜 포함)
            if 'secondary' in fundus_config:
                secondary_info = fundus_config['secondary']
                folder_path = secondary_info['path']
                pattern = re.compile(secondary_info['pattern'])

                log_callback(f"  📂 Secondary 스캔: {folder_path}")

                if os.path.exists(folder_path):
                    try:
                        items = os.listdir(folder_path)
                        total_items = len(items)
                        log_callback(f"     전체: {total_items}개")

                        # 파일명에 오늘 날짜가 포함된 것만 필터링
                        # 예: 204775-20250919@161455-l4-s.jpg
                        filename_matched = 0
                        secondary_charts = set()

                        for item in items:
                            if today_str in item:  # 20251117 형식
                                filename_matched += 1
                                match = pattern.search(item)
                                if match:
                                    chart_num = self.extract_chart_number(match)
                                    if self.is_valid_chart_number(chart_num):
                                        secondary_charts.add(chart_num)

                        log_callback(f"     오늘 날짜 파일: {filename_matched}개")
                        log_callback(f"     ✅ Secondary: {len(secondary_charts)}명 (중복 제거)")

                        # 합집합
                        before_merge = len(fundus_charts)
                        fundus_charts.update(secondary_charts)
                        after_merge = len(fundus_charts)

                        if before_merge > 0:
                            overlap = before_merge + len(secondary_charts) - after_merge
                            if overlap > 0:
                                log_callback(f"     💡 Fundus & Secondary 중복: {overlap}명")

                    except Exception as e:
                        log_callback(f"  ⚠️  Secondary 스캔 오류: {e}")
                else:
                    log_callback(f"  ⚠️  경로 없음: {folder_path}")

        except Exception as e:
            log_callback(f"  ❌ 안저 계산 오류: {str(e)}")

        log_callback(f"  📊 안저 최종 집계: {len(fundus_charts)}명 (중복 제거 완료)")
        return len(fundus_charts)

    def process_reservation_file(self, file_path: str, log_callback) -> Dict[str, int]:
        """예약 파일 처리 (.xlsx, .xls 모두 지원)"""
        counts = {'verion': 0, 'lensx': 0, 'ex500': 0}
        found_cells = set()
        search_keyword = self.config['reservation'].get('search_keyword', '예약비고:')

        try:
            # .xls 파일인 경우 xlrd로 읽기
            if file_path.lower().endswith('.xls') and not file_path.lower().endswith('.xlsx'):
                if not HAS_XLRD:
                    log_callback(f"  ⚠️  .xls 파일 읽기 실패: xlrd 라이브러리가 필요합니다")
                    log_callback(f"     설치: pip install xlrd")
                    return counts

                # xlrd로 .xls 파일 읽기
                import xlrd
                xls_book = xlrd.open_workbook(file_path)

                for sheet in xls_book.sheets():
                    for row_idx in range(sheet.nrows):
                        for col_idx in range(sheet.ncols):
                            cell = sheet.cell(row_idx, col_idx)
                            if cell.value is None or cell.value == '':
                                continue

                            cell_value = str(cell.value).lower()

                            if search_keyword.lower() not in cell_value:
                                continue

                            cell_key = f"{sheet.name}_{row_idx}_{col_idx}_{cell_value}"
                            if cell_key in found_cells:
                                continue
                            found_cells.add(cell_key)

                            # 각 셀마다 베리온/LensX/EX500 플래그 체크 (중복 방지, 대소문자 무시)
                            cell_value_lower = cell_value.lower()
                            has_verion = any(kw.lower() in cell_value_lower for kw in self.config['reservation']['verion_keywords'])
                            has_lensx = any(kw.lower() in cell_value_lower for kw in self.config['reservation']['lensx_keywords'])
                            has_ex500 = any(kw.lower() in cell_value_lower for kw in self.config['reservation']['ex500_keywords'])

                            if has_verion:
                                counts['verion'] += 1
                            if has_lensx:
                                counts['lensx'] += 1
                            if has_ex500:
                                counts['ex500'] += 1

                return counts

            # .xlsx 파일은 openpyxl로 읽기
            wb = load_workbook(file_path, data_only=True)

            for sheet in wb.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value is None:
                            continue

                        cell_value = str(cell.value).lower()

                        if search_keyword.lower() not in cell_value:
                            continue

                        cell_key = f"{sheet.title}_{cell.coordinate}_{cell_value}"
                        if cell_key in found_cells:
                            continue
                        found_cells.add(cell_key)

                        # 각 셀마다 베리온/LensX/EX500 플래그 체크 (중복 방지, 대소문자 무시)
                        cell_value_lower = cell_value.lower()
                        has_verion = any(kw.lower() in cell_value_lower for kw in self.config['reservation']['verion_keywords'])
                        has_lensx = any(kw.lower() in cell_value_lower for kw in self.config['reservation']['lensx_keywords'])
                        has_ex500 = any(kw.lower() in cell_value_lower for kw in self.config['reservation']['ex500_keywords'])

                        if has_verion:
                            counts['verion'] += 1
                        if has_lensx:
                            counts['lensx'] += 1
                        if has_ex500:
                            counts['ex500'] += 1

            wb.close()

        except Exception as e:
            log_callback(f"  ❌ 예약 파일 처리 오류: {str(e)}")

        return counts

    def write_excel(self, output_path: str, staff_selected: List[str],
                   manual_fag: int, manual_glasses: int, manual_lasik: int,
                   manual_octs: int, reservation_counts: Dict[str, int], log_callback,
                   glaucoma_count: int = None, fundus_count: int = None) -> bool:
        """엑셀 파일 작성"""
        try:
            template_file = self.config['template_file']
            if not os.path.exists(template_file):
                log_callback(f"  ❌ 템플릿 파일 없음: {template_file}")
                return False

            wb = load_workbook(template_file)
            ws = wb[self.config['target_sheet']]

            # 날짜 기입
            date_cell = self.config['date_cell']
            ws.cell(date_cell['row'], date_cell['col']).value = self.today.strftime('%Y-%m-%d')

            # 근무 인원 기입
            staff_cell = self.config['staff_cell']
            staff_count = len(staff_selected)
            staff_text = f"{staff_count}명( {', '.join(staff_selected)} )"
            ws.cell(staff_cell['row'], staff_cell['col']).value = staff_text

            # 각 장비별 결과 기입
            for equipment_id, chart_set in self.chart_numbers.items():
                if equipment_id in self.config['equipment']:
                    cell_info = self.config['equipment'][equipment_id]['cell']
                    # OCT는 OCTS 수기입력과 합산
                    if equipment_id == 'OCT':
                        oct_auto = len(chart_set)
                        oct_total = oct_auto + manual_octs
                        ws.cell(cell_info['row'], cell_info['col']).value = oct_total
                        log_callback(f"  ✓ OCT 합산: 자동({oct_auto}) + OCTS({manual_octs}) = {oct_total}")
                    else:
                        ws.cell(cell_info['row'], cell_info['col']).value = len(chart_set)

            # 특수 항목 기입 (전달된 값이 있으면 사용, 없으면 계산)
            if glaucoma_count is None:
                glaucoma_count = self.calculate_glaucoma(log_callback)
            glaucoma_cell = self.config['special_items']['녹내장']['cell']
            ws.cell(glaucoma_cell['row'], glaucoma_cell['col']).value = glaucoma_count

            if fundus_count is None:
                fundus_count = self.calculate_fundus(log_callback)
            fundus_cell = self.config['special_items']['안저']['cell']
            ws.cell(fundus_cell['row'], fundus_cell['col']).value = fundus_count

            # 수기 입력 항목
            lasik_cell = self.config['manual_input']['라식']
            ws.cell(lasik_cell['row'], lasik_cell['col']).value = manual_lasik

            fag_cell = self.config['manual_input']['FAG']
            ws.cell(fag_cell['row'], fag_cell['col']).value = manual_fag

            glasses_cell = self.config['manual_input']['안경검사']
            ws.cell(glasses_cell['row'], glasses_cell['col']).value = manual_glasses

            # 예약 파일 결과 (Verion은 예약파일에서만 추출)
            verion_count = reservation_counts.get('verion', 0)
            verion_cell = self.config['reservation']['cells']['verion']
            ws.cell(verion_cell['row'], verion_cell['col']).value = verion_count

            lensx_cell = self.config['reservation']['cells']['lensx']
            ws.cell(lensx_cell['row'], lensx_cell['col']).value = reservation_counts.get('lensx', 0)

            ex500_cell = self.config['reservation']['cells']['ex500']
            ws.cell(ex500_cell['row'], ex500_cell['col']).value = reservation_counts.get('ex500', 0)

            wb.save(output_path)
            wb.close()

            log_callback("  ✓ 엑셀 작성 완료")
            return True

        except Exception as e:
            log_callback(f"  ❌ 엑셀 작성 오류: {str(e)}")
            return False

    def convert_to_pdf(self, excel_path: str, pdf_path: str, log_callback) -> bool:
        """엑셀 파일을 PDF로 변환"""
        if not HAS_WIN32:
            log_callback("  ⚠️  pywin32가 없어 PDF 변환 불가")
            return False

        try:
            os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

            # COM 라이브러리 초기화
            pythoncom.CoInitialize()

            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False

                wb = excel.Workbooks.Open(os.path.abspath(excel_path))
                ws = wb.Worksheets(self.config['target_sheet'])

                ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path))

                wb.Close(SaveChanges=False)
                excel.Quit()

                log_callback(f"  ✓ PDF 생성 완료: {pdf_path}")
                return True

            finally:
                # COM 라이브러리 정리
                pythoncom.CoUninitialize()

        except Exception as e:
            log_callback(f"  ❌ PDF 변환 오류: {str(e)}")
            return False


class DailyReportGUI:
    """일일결산 시스템의 GUI 클래스"""

    def __init__(self, root: tk.Tk, system: DailyReportSystem):
        self.root = root
        self.system = system
        self.reservation_files = []
        self.log_file_handle = None  # 로그 파일 핸들
        self.scan_results = {}  # 스캔 결과 저장
        self.setup_gui()

    def setup_gui(self):
        """GUI 구성 요소 생성"""
        self.root.title("일일결산 자동화 시스템 (최적화)")
        self.root.geometry("900x850")
        self.root.resizable(True, True)

        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)

        # 좌측 입력 영역
        left_frame = ttk.Frame(main_frame, padding="5")
        left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 우측 로그 영역
        right_frame = ttk.Frame(main_frame, padding="5")
        right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(1, weight=1)

        # === 좌측 영역 구성 ===

        # 0. 결산 날짜 선택
        date_label = ttk.Label(left_frame, text="📅 결산 날짜", font=("", 12, "bold"))
        date_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))

        date_frame = ttk.Frame(left_frame)
        date_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))

        # 날짜 입력 (YYYY-MM-DD)
        self.date_entry = ttk.Entry(date_frame, width=12)
        self.date_entry.insert(0, date.today().strftime('%Y-%m-%d'))
        self.date_entry.grid(row=0, column=0, padx=(0, 5))

        today_btn = ttk.Button(date_frame, text="오늘", width=6,
                               command=lambda: self.set_date(date.today()))
        today_btn.grid(row=0, column=1, padx=2)

        yesterday_btn = ttk.Button(date_frame, text="어제", width=6,
                                   command=lambda: self.set_date(date.today() - timedelta(days=1)))
        yesterday_btn.grid(row=0, column=2, padx=2)

        ttk.Separator(left_frame, orient='horizontal').grid(row=2, column=0, columnspan=2,
                                                             sticky=(tk.W, tk.E), pady=5)

        # 1. 근무 인원 선택
        staff_label = ttk.Label(left_frame, text="📋 근무 인원", font=("", 12, "bold"))
        staff_label.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))

        staff_canvas = tk.Canvas(left_frame, height=200)
        staff_scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=staff_canvas.yview)
        staff_scrollable = ttk.Frame(staff_canvas)

        staff_scrollable.bind(
            "<Configure>",
            lambda e: staff_canvas.configure(scrollregion=staff_canvas.bbox("all"))
        )

        staff_canvas.create_window((0, 0), window=staff_scrollable, anchor="nw")
        staff_canvas.configure(yscrollcommand=staff_scrollbar.set)

        staff_canvas.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        staff_scrollbar.grid(row=4, column=1, sticky=(tk.N, tk.S), pady=(0, 10))

        # 직원 체크박스 생성
        self.staff_vars = {}
        for i, staff_name in enumerate(self.system.config['staff_list']):
            var = tk.BooleanVar(value=True)
            self.staff_vars[staff_name] = var
            cb = ttk.Checkbutton(staff_scrollable, text=staff_name, variable=var)
            cb.grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)

        # 2. 예약 파일 선택
        ttk.Separator(left_frame, orient='horizontal').grid(row=5, column=0, columnspan=2,
                                                             sticky=(tk.W, tk.E), pady=10)

        reservation_label = ttk.Label(left_frame, text="📁 예약 파일", font=("", 12, "bold"))
        reservation_label.grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))

        self.file_button = ttk.Button(left_frame, text="파일 선택...", command=self.select_files)
        self.file_button.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))

        self.file_label = ttk.Label(left_frame, text="선택된 파일: 없음", foreground="gray")
        self.file_label.grid(row=8, column=0, columnspan=2, sticky=tk.W)

        # 3. 수기 입력
        ttk.Separator(left_frame, orient='horizontal').grid(row=9, column=0, columnspan=2,
                                                             sticky=(tk.W, tk.E), pady=10)

        manual_label = ttk.Label(left_frame, text="✍ 수기 입력", font=("", 12, "bold"))
        manual_label.grid(row=10, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))

        lasik_label = ttk.Label(left_frame, text="라식:")
        lasik_label.grid(row=11, column=0, sticky=tk.W, padx=(0, 5))

        self.lasik_entry = ttk.Entry(left_frame, width=10)
        self.lasik_entry.insert(0, "0")
        self.lasik_entry.grid(row=11, column=1, sticky=tk.W, pady=3)

        fag_label = ttk.Label(left_frame, text="FAG:")
        fag_label.grid(row=12, column=0, sticky=tk.W, padx=(0, 5))

        self.fag_entry = ttk.Entry(left_frame, width=10)
        self.fag_entry.insert(0, "0")
        self.fag_entry.grid(row=12, column=1, sticky=tk.W, pady=3)

        glasses_label = ttk.Label(left_frame, text="안경검사:")
        glasses_label.grid(row=13, column=0, sticky=tk.W, padx=(0, 5))

        self.glasses_entry = ttk.Entry(left_frame, width=10)
        self.glasses_entry.insert(0, "0")
        self.glasses_entry.grid(row=13, column=1, sticky=tk.W, pady=3)

        octs_label = ttk.Label(left_frame, text="OCTS:")
        octs_label.grid(row=14, column=0, sticky=tk.W, padx=(0, 5))

        self.octs_entry = ttk.Entry(left_frame, width=10)
        self.octs_entry.insert(0, "0")
        self.octs_entry.grid(row=14, column=1, sticky=tk.W, pady=3)

        # 4. 스캔 버튼
        ttk.Separator(left_frame, orient='horizontal').grid(row=15, column=0, columnspan=2,
                                                             sticky=(tk.W, tk.E), pady=15)

        self.scan_button = ttk.Button(left_frame, text="🔍 스캔 시작", command=self.run_scan)
        self.scan_button.grid(row=16, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        # 5. 결과 확인 및 수정 (초기에는 숨김)
        ttk.Separator(left_frame, orient='horizontal').grid(row=17, column=0, columnspan=2,
                                                             sticky=(tk.W, tk.E), pady=10)

        result_label = ttk.Label(left_frame, text="📊 스캔 결과 (수정 가능)", font=("", 12, "bold"))
        result_label.grid(row=18, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))

        # 결과 프레임 (스크롤 가능)
        result_canvas = tk.Canvas(left_frame, height=300)
        result_scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=result_canvas.yview)
        self.result_frame = ttk.Frame(result_canvas)

        self.result_frame.bind(
            "<Configure>",
            lambda e: result_canvas.configure(scrollregion=result_canvas.bbox("all"))
        )

        result_canvas.create_window((0, 0), window=self.result_frame, anchor="nw")
        result_canvas.configure(yscrollcommand=result_scrollbar.set)

        result_canvas.grid(row=19, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        result_scrollbar.grid(row=19, column=1, sticky=(tk.N, tk.S))

        # 결과 항목들 (Entry 위젯) - 초기에는 비활성화
        self.result_entries = {}
        result_items = [
            ('OQAS', '백내장'),
            ('HFA', '시야'),
            ('OCT', 'OCT'),
            ('ORB', 'ORB'),
            ('SP', '내피'),
            ('TOPO', 'Tomey'),
            ('GLAUCOMA', '녹내장'),
            ('FUNDUS', '안저'),
            ('LASIK', '라식'),
            ('GLASSES', '안경검사'),
            ('FAG', 'FAG'),
            ('VERION', 'Verion'),
            ('LENSX', 'LensX'),
            ('EX500', 'EX500'),
        ]

        for idx, (key, label_text) in enumerate(result_items):
            label = ttk.Label(self.result_frame, text=f"{label_text}:")
            label.grid(row=idx, column=0, sticky=tk.W, padx=(0, 5), pady=2)

            entry = ttk.Entry(self.result_frame, width=10, state='disabled')
            entry.insert(0, "0")
            entry.grid(row=idx, column=1, sticky=tk.W, pady=2)
            self.result_entries[key] = entry

        # 6. PDF 출력 버튼 (초기에는 비활성화)
        self.output_button = ttk.Button(left_frame, text="✅ 확정 및 PDF 출력",
                                        command=self.run_output, state='disabled')
        self.output_button.grid(row=20, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        # === 우측 영역 구성 ===

        log_label = ttk.Label(right_frame, text="실행 로그", font=("", 12, "bold"))
        log_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        self.log_text = scrolledtext.ScrolledText(right_frame, width=50, height=30,
                                                   state='disabled', wrap=tk.WORD)
        self.log_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    def set_date(self, target_date: date):
        """날짜 설정"""
        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, target_date.strftime('%Y-%m-%d'))

    def log(self, message: str):
        """로그 메시지 출력 (화면 + 파일)"""
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')
        self.root.update()

        # 로그 파일에도 기록
        if self.log_file_handle:
            try:
                self.log_file_handle.write(message + '\n')
                self.log_file_handle.flush()  # 즉시 디스크에 쓰기
            except Exception as e:
                print(f"로그 파일 쓰기 오류: {e}")

    def select_files(self):
        """예약 파일 선택"""
        files = filedialog.askopenfilenames(
            title="예약 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if files:
            self.reservation_files = list(files)
            file_count = len(files)
            self.file_label.config(text=f"선택된 파일: {file_count}개", foreground="blue")
        else:
            self.reservation_files = []
            self.file_label.config(text="선택된 파일: 없음", foreground="gray")

    def get_selected_staff(self) -> List[str]:
        """선택된 직원 목록 반환"""
        return [name for name, var in self.staff_vars.items() if var.get()]

    def run_scan(self):
        """1단계: 스캔 실행"""
        self.scan_button.config(state='disabled')
        self.file_button.config(state='disabled')

        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')

        thread = threading.Thread(target=self.process_scan, daemon=True)
        thread.start()

    def run_output(self):
        """2단계: PDF 출력"""
        self.output_button.config(state='disabled')

        thread = threading.Thread(target=self.process_output, daemon=True)
        thread.start()

    def process_scan(self):
        """1단계: 스캔 처리 - 결과를 화면에 표시"""
        # 로그 파일 열기
        log_filename = f"결산로그_{date.today().strftime('%Y-%m-%d')}.txt"
        try:
            self.log_file_handle = open(log_filename, 'w', encoding='utf-8')
        except Exception as e:
            print(f"로그 파일 생성 오류: {e}")
            self.log_file_handle = None

        try:
            # 날짜 파싱
            date_str = self.date_entry.get()
            try:
                target_date = datetime.strptime(date_str, '%Y-%m-%d').date()
                self.system.today = target_date
            except ValueError:
                self.log("❌ 날짜 형식 오류! YYYY-MM-DD 형식으로 입력하세요.")
                self.scan_button.config(state='normal')
                self.file_button.config(state='normal')
                return

            self.log("=" * 54)
            self.log(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 스캔 시작")
            self.log(f"결산 날짜: {target_date.strftime('%Y-%m-%d')}")
            self.log(f"로그 파일: {log_filename}")
            self.log("=" * 54)
            self.log("")

            # 1. 디렉토리 자동 스캔
            self.log("[1/3] 디렉토리 자동 스캔 중...")
            for equipment_id in self.system.config['equipment'].keys():
                equipment_name = self.system.config['equipment'][equipment_id]['name']
                self.log(f"  🔍 {equipment_name} 스캔 중...")

                chart_set = self.system.scan_directory_fast(equipment_id, self.log)
                self.system.chart_numbers[equipment_id] = chart_set

                self.log(f"  ✓ {equipment_name}: {len(chart_set)}건")

            self.log("")

            # 2. 특수 항목 계산
            self.log("[2/3] 특수 항목 계산 중...")

            glaucoma_count = self.system.calculate_glaucoma(self.log)
            self.log(f"  ✓ 녹내장 (HFA ∩ OCT): {glaucoma_count}건")

            fundus_count = self.system.calculate_fundus(self.log)
            self.log(f"  ✓ 안저: {fundus_count}건")

            self.log("")

            # 3. 예약 파일 처리
            reservation_counts = {'verion': 0, 'lensx': 0, 'ex500': 0}

            if self.reservation_files:
                self.log(f"[3/3] 예약 파일 분석 중... ({len(self.reservation_files)}개 파일)")

                for file_path in self.reservation_files:
                    file_name = os.path.basename(file_path)
                    self.log(f"  📄 {file_name}")

                    file_counts = self.system.process_reservation_file(file_path, self.log)

                    for key in reservation_counts:
                        reservation_counts[key] += file_counts[key]

                self.log(f"  ✓ Verion (예약): {reservation_counts['verion']}건")
                self.log(f"  ✓ Lensx: {reservation_counts['lensx']}건")
                self.log(f"  ✓ EX500: {reservation_counts['ex500']}건")
            else:
                self.log("[3/3] 예약 파일 선택 안 함 (건너뜀)")

            self.log("")

            # 스캔 결과를 인스턴스 변수에 저장
            self.scan_results = {
                'glaucoma_count': glaucoma_count,
                'fundus_count': fundus_count,
                'reservation_counts': reservation_counts
            }

            # 결과 Entry 위젯 업데이트
            self.root.after(0, self.update_result_entries)

            self.log("=" * 54)
            self.log("✅ 스캔 완료! 결과를 확인하고 수정 후 PDF 출력 버튼을 클릭하세요.")
            self.log("=" * 54)
            self.log("")

        except Exception as e:
            self.log("")
            self.log("=" * 54)
            self.log(f"❌ 오류 발생: {str(e)}")
            self.log("=" * 54)
            self.scan_button.config(state='normal')
            self.file_button.config(state='normal')

        finally:
            # 로그 파일 닫기
            if self.log_file_handle:
                try:
                    self.log_file_handle.close()
                    self.log_file_handle = None
                except Exception as e:
                    print(f"로그 파일 닫기 오류: {e}")

    def update_result_entries(self):
        """스캔 결과를 Entry 위젯에 표시하고 편집 가능하게 설정"""
        # 각 항목의 값 설정
        entry_values = {
            'OQAS': len(self.system.chart_numbers.get('OQAS', set())),
            'HFA': len(self.system.chart_numbers.get('HFA', set())),
            'OCT': len(self.system.chart_numbers.get('OCT', set())) + int(self.octs_entry.get() or 0),
            'ORB': len(self.system.chart_numbers.get('ORB', set())),
            'SP': len(self.system.chart_numbers.get('SP', set())),
            'TOPO': len(self.system.chart_numbers.get('TOPO', set())),
            'GLAUCOMA': self.scan_results['glaucoma_count'],
            'FUNDUS': self.scan_results['fundus_count'],
            'LASIK': int(self.lasik_entry.get() or 0),
            'GLASSES': int(self.glasses_entry.get() or 0),
            'FAG': int(self.fag_entry.get() or 0),
            'VERION': self.scan_results['reservation_counts']['verion'],
            'LENSX': self.scan_results['reservation_counts']['lensx'],
            'EX500': self.scan_results['reservation_counts']['ex500'],
        }

        # Entry 위젯 업데이트 및 편집 가능하게 설정
        for key, value in entry_values.items():
            entry = self.result_entries[key]
            entry.config(state='normal')
            entry.delete(0, tk.END)
            entry.insert(0, str(value))

        # PDF 출력 버튼 활성화
        self.output_button.config(state='normal')
        self.scan_button.config(state='normal')
        self.file_button.config(state='normal')

    def process_output(self):
        """2단계: PDF 출력 - Entry 위젯의 값을 읽어서 엑셀/PDF 생성"""
        # 로그 파일 열기
        log_filename = f"결산로그_{date.today().strftime('%Y-%m-%d')}.txt"
        try:
            self.log_file_handle = open(log_filename, 'a', encoding='utf-8')  # append 모드
        except Exception as e:
            print(f"로그 파일 열기 오류: {e}")
            self.log_file_handle = None

        try:
            self.log("")
            self.log("=" * 54)
            self.log(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] PDF 출력 시작")
            self.log("=" * 54)
            self.log("")

            # Entry 위젯에서 값 읽기
            self.log("[1/2] 확정된 값:")
            try:
                result_values = {}
                for key, entry in self.result_entries.items():
                    value = int(entry.get() or 0)
                    result_values[key] = value
                    label_map = {
                        'OQAS': '백내장', 'HFA': '시야', 'OCT': 'OCT', 'ORB': 'ORB',
                        'SP': '내피', 'TOPO': 'Tomey', 'GLAUCOMA': '녹내장', 'FUNDUS': '안저',
                        'LASIK': '라식', 'GLASSES': '안경검사', 'FAG': 'FAG',
                        'VERION': 'Verion', 'LENSX': 'LensX', 'EX500': 'EX500'
                    }
                    self.log(f"  {label_map.get(key, key)}: {value}건")
            except ValueError as e:
                self.log(f"  ⚠️  값 읽기 오류: {e}")
                self.output_button.config(state='normal')
                return

            self.log("")

            # 엑셀 작성용 데이터 준비
            staff_selected = self.get_selected_staff()
            if not staff_selected:
                self.log("  ⚠️  경고: 직원이 선택되지 않았습니다.")

            # 예약 데이터
            reservation_counts = {
                'verion': result_values['VERION'],
                'lensx': result_values['LENSX'],
                'ex500': result_values['EX500']
            }

            # 수동 입력 데이터
            manual_lasik = result_values['LASIK']
            manual_fag = result_values['FAG']
            manual_glasses = result_values['GLASSES']
            manual_octs = 0  # OCTS는 OCT에 이미 포함됨

            # 자동 스캔 데이터를 직접 설정 (Entry 값으로 덮어쓰기)
            self.system.chart_numbers['OQAS'] = set(range(result_values['OQAS']))  # 더미 데이터
            self.system.chart_numbers['HFA'] = set(range(result_values['HFA']))
            self.system.chart_numbers['OCT'] = set(range(result_values['OCT']))
            self.system.chart_numbers['ORB'] = set(range(result_values['ORB']))
            self.system.chart_numbers['SP'] = set(range(result_values['SP']))
            self.system.chart_numbers['TOPO'] = set(range(result_values['TOPO']))

            # 특수 항목도 더미 데이터로 설정
            self.system.chart_numbers['녹내장'] = set(range(result_values['GLAUCOMA']))
            self.system.chart_numbers['안저'] = set(range(result_values['FUNDUS']))

            # 엑셀 작성
            self.log("[2/2] 엑셀 파일 작성 및 PDF 생성 중...")

            today_str = date.today().strftime('%Y%m%d')
            temp_excel = f"일일결산_{today_str}_temp.xlsx"

            success = self.system.write_excel(
                temp_excel, staff_selected, manual_fag, manual_glasses, manual_lasik,
                manual_octs, reservation_counts, self.log,
                glaucoma_count=result_values['GLAUCOMA'],
                fundus_count=result_values['FUNDUS']
            )

            if not success:
                self.log("")
                self.log("=" * 54)
                self.log("❌ 결산 실패: 엑셀 작성 오류")
                self.log("=" * 54)
                self.output_button.config(state='normal')
                return

            self.log("")

            # PDF 변환
            pdf_path = self.system.config['output_pdf'].replace('{date}', today_str)
            pdf_success = self.system.convert_to_pdf(temp_excel, pdf_path, self.log)

            self.log("")
            self.log("=" * 54)
            self.log("✅ 결산 완료!")
            self.log("=" * 54)
            self.log("")

            # PDF 열기
            if pdf_success and os.path.exists(pdf_path):
                self.log("📄 PDF 파일을 엽니다...")
                if sys.platform == 'win32':
                    os.startfile(pdf_path)
                else:
                    self.log(f"  PDF 경로: {pdf_path}")

                try:
                    os.remove(temp_excel)
                except:
                    pass
            else:
                self.log(f"📄 엑셀 파일이 저장되었습니다: {temp_excel}")

        except Exception as e:
            self.log("")
            self.log("=" * 54)
            self.log(f"❌ 오류 발생: {str(e)}")
            self.log("=" * 54)

        finally:
            # 로그 파일 닫기
            if self.log_file_handle:
                try:
                    self.log_file_handle.close()
                    self.log_file_handle = None
                except Exception as e:
                    print(f"로그 파일 닫기 오류: {e}")

            self.output_button.config(state='normal')

    def run_report(self):
        """결산 실행 (구버전 호환용 - 사용 안 함)"""
        self.run_button.config(state='disabled')
        self.file_button.config(state='disabled')

        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')

        thread = threading.Thread(target=self.process_report, daemon=True)
        thread.start()

    def process_report(self):
        """결산 처리 메인 로직"""
        # 로그 파일 열기
        log_filename = f"결산로그_{date.today().strftime('%Y-%m-%d')}.txt"
        try:
            self.log_file_handle = open(log_filename, 'w', encoding='utf-8')
        except Exception as e:
            print(f"로그 파일 생성 오류: {e}")
            self.log_file_handle = None

        try:
            # 날짜 파싱
            date_str = self.date_entry.get()
            try:
                target_date = datetime.strptime(date_str, '%Y-%m-%d').date()
                self.system.today = target_date
            except ValueError:
                self.log("❌ 날짜 형식 오류! YYYY-MM-DD 형식으로 입력하세요.")
                self.run_button.config(state='normal')
                self.file_button.config(state='normal')
                return

            self.log("=" * 54)
            self.log(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 결산 시작 (최적화 버전)")
            self.log(f"결산 날짜: {target_date.strftime('%Y-%m-%d')}")
            self.log(f"로그 파일: {log_filename}")
            self.log("=" * 54)
            self.log("")

            # 1. 디렉토리 자동 스캔
            self.log("[1/4] 디렉토리 자동 스캔 중...")
            for equipment_id in self.system.config['equipment'].keys():
                equipment_name = self.system.config['equipment'][equipment_id]['name']
                self.log(f"  🔍 {equipment_name} 스캔 중...")

                chart_set = self.system.scan_directory_fast(equipment_id, self.log)
                self.system.chart_numbers[equipment_id] = chart_set

                self.log(f"  ✓ {equipment_name}: {len(chart_set)}건")

            self.log("")

            # 2. 특수 항목 계산
            self.log("[2/4] 특수 항목 계산 중...")

            glaucoma_count = self.system.calculate_glaucoma(self.log)
            self.log(f"  ✓ 녹내장 (HFA ∩ OCT): {glaucoma_count}건")

            fundus_count = self.system.calculate_fundus(self.log)
            self.log(f"  ✓ 안저: {fundus_count}건")

            self.log("")

            # 3. 예약 파일 처리
            reservation_counts = {'verion': 0, 'lensx': 0, 'ex500': 0}

            if self.reservation_files:
                self.log(f"[3/4] 예약 파일 분석 중... ({len(self.reservation_files)}개 파일)")

                for file_path in self.reservation_files:
                    file_name = os.path.basename(file_path)
                    self.log(f"  📄 {file_name}")

                    file_counts = self.system.process_reservation_file(file_path, self.log)

                    for key in reservation_counts:
                        reservation_counts[key] += file_counts[key]

                self.log(f"  ✓ Verion (예약): {reservation_counts['verion']}건")
                self.log(f"  ✓ Lensx: {reservation_counts['lensx']}건")
                self.log(f"  ✓ EX500: {reservation_counts['ex500']}건")
            else:
                self.log("[3/4] 예약 파일 선택 안 함 (건너뜀)")

            self.log("")

            # 4. 엑셀 작성
            self.log("[4/4] 엑셀 파일 작성 중...")

            staff_selected = self.get_selected_staff()
            if not staff_selected:
                self.log("  ⚠️  경고: 직원이 선택되지 않았습니다.")

            try:
                manual_lasik = int(self.lasik_entry.get())
            except ValueError:
                manual_lasik = 0
                self.log("  ⚠️  라식 값이 올바르지 않아 0으로 설정합니다.")

            try:
                manual_fag = int(self.fag_entry.get())
            except ValueError:
                manual_fag = 0
                self.log("  ⚠️  FAG 값이 올바르지 않아 0으로 설정합니다.")

            try:
                manual_glasses = int(self.glasses_entry.get())
            except ValueError:
                manual_glasses = 0
                self.log("  ⚠️  안경검사 값이 올바르지 않아 0으로 설정합니다.")

            try:
                manual_octs = int(self.octs_entry.get())
            except ValueError:
                manual_octs = 0
                self.log("  ⚠️  OCTS 값이 올바르지 않아 0으로 설정합니다.")

            today_str = date.today().strftime('%Y%m%d')
            temp_excel = f"일일결산_{today_str}_temp.xlsx"

            success = self.system.write_excel(
                temp_excel, staff_selected, manual_fag, manual_glasses, manual_lasik,
                manual_octs, reservation_counts, self.log
            )

            if not success:
                self.log("")
                self.log("=" * 54)
                self.log("❌ 결산 실패: 엑셀 작성 오류")
                self.log("=" * 54)
                return

            self.log("")

            # 5. PDF 변환
            self.log("[5/5] PDF 생성 중...")

            pdf_path = self.system.config['output_pdf'].replace('{date}', today_str)
            pdf_success = self.system.convert_to_pdf(temp_excel, pdf_path, self.log)

            self.log("")
            self.log("=" * 54)
            self.log("✅ 결산 완료!")
            self.log("=" * 54)
            self.log("")

            # PDF 열기
            if pdf_success and os.path.exists(pdf_path):
                self.log("📄 PDF 파일을 엽니다...")
                if sys.platform == 'win32':
                    os.startfile(pdf_path)
                else:
                    self.log(f"  PDF 경로: {pdf_path}")

                try:
                    os.remove(temp_excel)
                except:
                    pass
            else:
                self.log(f"📄 엑셀 파일이 저장되었습니다: {temp_excel}")

        except Exception as e:
            self.log("")
            self.log("=" * 54)
            self.log(f"❌ 오류 발생: {str(e)}")
            self.log("=" * 54)

        finally:
            # 로그 파일 닫기
            if self.log_file_handle:
                try:
                    self.log_file_handle.close()
                    self.log_file_handle = None
                except Exception as e:
                    print(f"로그 파일 닫기 오류: {e}")

            self.run_button.config(state='normal')
            self.file_button.config(state='normal')


def main():
    """메인 함수"""
    config_path = "config.json"
    if not os.path.exists(config_path):
        messagebox.showerror("오류", "config.json 파일을 찾을 수 없습니다.")
        sys.exit(1)

    system = DailyReportSystem(config_path)

    root = tk.Tk()
    app = DailyReportGUI(root, system)
    root.mainloop()


if __name__ == "__main__":
    main()
