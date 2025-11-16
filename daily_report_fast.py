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
from datetime import datetime, date
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

    def get_today_folder_path(self, base_path: str, equipment_id: str) -> Optional[str]:
        """오늘 날짜 폴더 경로 생성 (장비별 폴더 구조에 맞게)"""
        today = self.today

        # 장비별 폴더 구조
        # TOPO: 2025\01\TOPO 01.18
        # ORB: 2025\2025.01\ORB 01.18
        # OCT: 2025\01\18

        if equipment_id == 'TOPO':
            folder = today.strftime("%Y\\%m\\TOPO %m.%d")
        elif equipment_id == 'ORB':
            folder = today.strftime("%Y\\%Y.%m\\ORB %m.%d")
        elif equipment_id == 'OCT':
            folder = today.strftime("%Y\\%m\\%d")
        elif equipment_id == 'OQAS':
            folder = today.strftime("%Y\\%m\\%d.%m")
        else:
            # SP, HFA, IOL700 등은 단일 폴더 구조
            return base_path

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
        """
        equipment = self.config['equipment'][equipment_id]
        base_path = equipment['path']
        pattern = self.compiled_patterns[equipment_id]
        scan_type = equipment['scan_type']

        chart_numbers = set()

        if not os.path.exists(base_path):
            log_callback(f"  ⚠️  경로 없음: {base_path}")
            return chart_numbers

        try:
            # 오늘 날짜 폴더 경로 찾기
            today_folder = self.get_today_folder_path(base_path, equipment_id)

            if today_folder is None:
                # 폴더 구조가 없는 경우 (SP, HFA, IOL700 등) 직접 스캔
                today_folder = base_path
                use_creation_time = equipment.get('use_creation_time', False)
                log_callback(f"     📂 스캔 경로: {today_folder}")
                log_callback(f"     🔍 날짜 확인: {'생성일' if use_creation_time else '파일명'}")

                # 단일 폴더만 스캔 (os.listdir 사용) - 최적화 버전
                if scan_type == 'file':
                    files = os.listdir(today_folder)
                    total_files = len(files)

                    # 확장자 필터링 먼저 (빠른 연산)
                    valid_extensions = self.config['validation']['file_extensions']
                    candidate_files = [f for f in files if any(f.lower().endswith(ext) for ext in valid_extensions)]

                    log_callback(f"     📊 전체: {total_files}개 / 유효 확장자: {len(candidate_files)}개")

                    if not candidate_files:
                        log_callback(f"     ⚠️  유효한 파일 없음")
                        return chart_numbers

                    # 최적화 1: 파일명에 날짜 있는지 먼저 체크
                    today_str = self.today.strftime('%Y%m%d')
                    today_str_dash = self.today.strftime('%Y-%m-%d')
                    today_str_dot = self.today.strftime('%Y.%m.%d')
                    date_patterns = [today_str, today_str_dash, today_str_dot]

                    filename_matched = 0
                    need_ctime_check = []

                    for file_name in candidate_files:
                        # 파일명에 오늘 날짜가 있으면 바로 처리
                        if any(dp in file_name for dp in date_patterns):
                            filename_matched += 1
                            match = pattern.search(file_name)
                            if match:
                                chart_num = match.group(1)
                                if self.is_valid_chart_number(chart_num):
                                    chart_numbers.add(chart_num)
                        elif use_creation_time:
                            need_ctime_check.append(file_name)

                    if filename_matched > 0:
                        log_callback(f"     ⚡ 파일명 날짜 매칭: {filename_matched}개 → {len(chart_numbers)}건")

                    # 최적화 2: 생성일 확인이 필요한 경우 (파일명에 날짜 없음)
                    if need_ctime_check and use_creation_time:
                        log_callback(f"     🔍 생성일 확인 필요: {len(need_ctime_check)}개")
                        log_callback(f"     ⚡ 최적화: 역순 스캔 + 조기 종료 + 병렬 처리")

                        # 역순 정렬 (최신 파일이 보통 끝에 있음)
                        need_ctime_check.sort(reverse=True)

                        def check_file_date(file_name):
                            file_path = os.path.join(today_folder, file_name)
                            try:
                                if not os.path.isfile(file_path):
                                    return None, None
                                ctime = os.path.getctime(file_path)
                                file_date = date.fromtimestamp(ctime)
                                if file_date == self.today:
                                    match = pattern.search(file_name)
                                    if match:
                                        chart_num = match.group(1)
                                        if self.is_valid_chart_number(chart_num):
                                            return chart_num, file_date
                                return None, file_date
                            except:
                                pass
                            return None, None

                        # 배치 처리 (1000개씩)
                        batch_size = 1000
                        total_checked = 0
                        consecutive_old_files = 0
                        ctime_matches = 0

                        for i in range(0, len(need_ctime_check), batch_size):
                            batch = need_ctime_check[i:i+batch_size]

                            # 병렬 처리
                            with ThreadPoolExecutor(max_workers=20) as executor:
                                futures = [executor.submit(check_file_date, f) for f in batch]

                                batch_old_count = 0
                                for future in as_completed(futures):
                                    chart_num, file_date = future.result()
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

                    log_callback(f"     📊 최종 결과: {len(chart_numbers)}건 (중복 제외)")
                return chart_numbers

            # 오늘 폴더와 하위 폴더만 스캔 (os.walk 사용)
            log_callback(f"     📂 스캔 경로: {today_folder}")

            total_files_count = 0
            total_dirs_count = 0

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
                            chart_num = match.group(1)
                            if self.is_valid_chart_number(chart_num):
                                chart_numbers.add(chart_num)

                # 폴더 스캔 (OCT의 경우)
                if scan_type == 'both':
                    for dir_name in dirs:
                        match = pattern.search(dir_name)
                        if match:
                            chart_num = match.group(1)
                            if self.is_valid_chart_number(chart_num):
                                chart_numbers.add(chart_num)

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
        """안저 계산 (FUNDERS + OPTOS 폴더) - 최적화 버전"""
        fundus_charts = set()
        pattern = re.compile(self.config['special_items']['안저']['pattern'])

        # 오늘 날짜 패턴
        today_str = self.today.strftime('%Y%m%d')
        today_str_dash = self.today.strftime('%Y-%m-%d')
        today_str_dot = self.today.strftime('%Y.%m.%d')
        date_patterns = [today_str, today_str_dash, today_str_dot]

        try:
            for folder_str in self.config['special_items']['안저']['folders']:
                if '[TODO' in folder_str or not os.path.exists(folder_str):
                    log_callback(f"  ⚠️  경로 없음 또는 미설정: {folder_str}")
                    continue

                log_callback(f"  📂 스캔: {folder_str}")

                # 오늘 생성된 항목만 - 최적화 버전
                try:
                    items = os.listdir(folder_str)
                    total_items = len(items)

                    # 1단계: 파일명 날짜 패턴 우선 필터링
                    candidates = []
                    filename_matched = 0

                    for item in items:
                        # 파일명에 오늘 날짜가 있는지 먼저 체크
                        has_today_in_name = any(dp in item for dp in date_patterns)

                        if has_today_in_name:
                            filename_matched += 1
                            match = pattern.search(item)
                            if match:
                                chart_num = match.group(1)
                                if self.is_valid_chart_number(chart_num):
                                    fundus_charts.add(chart_num)
                        else:
                            # 생성일 확인 필요
                            candidates.append((item, os.path.join(folder_str, item)))

                    log_callback(f"     전체: {total_items}개 / 파일명 매칭: {filename_matched}개")

                    # 2단계: 나머지는 병렬로 getctime 확인
                    if candidates:
                        log_callback(f"     🔍 생성일 확인: {len(candidates)}개")

                        def check_item_date(item_info):
                            item_name, item_path = item_info
                            try:
                                ctime = os.path.getctime(item_path)
                                file_date = date.fromtimestamp(ctime)
                                if file_date == self.today:
                                    match = pattern.search(item_name)
                                    if match:
                                        chart_num = match.group(1)
                                        if self.is_valid_chart_number(chart_num):
                                            return chart_num
                            except:
                                pass
                            return None

                        # 병렬 처리
                        with ThreadPoolExecutor(max_workers=10) as executor:
                            futures = [executor.submit(check_item_date, info) for info in candidates]
                            for future in as_completed(futures):
                                result = future.result()
                                if result:
                                    fundus_charts.add(result)

                except Exception as e:
                    log_callback(f"  ⚠️  폴더 스캔 오류: {e}")

        except Exception as e:
            log_callback(f"  ❌ 안저 계산 오류: {str(e)}")

        return len(fundus_charts)

    def process_reservation_file(self, file_path: str, log_callback) -> Dict[str, int]:
        """예약 파일 처리 (.xlsx, .xls 모두 지원)"""
        counts = {'verion': 0, 'lensx': 0, 'ex500': 0}
        found_cells = set()

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

                            if "수술방법:" not in cell_value:
                                continue

                            cell_key = f"{sheet.name}_{row_idx}_{col_idx}_{cell_value}"
                            if cell_key in found_cells:
                                continue
                            found_cells.add(cell_key)

                            if any(kw in cell_value for kw in self.config['reservation']['verion_keywords']):
                                counts['verion'] += 1
                            elif any(kw in cell_value for kw in self.config['reservation']['lensx_keywords']):
                                counts['lensx'] += 1
                            elif any(kw in cell_value for kw in self.config['reservation']['ex500_keywords']):
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

                        if "수술방법:" not in cell_value:
                            continue

                        cell_key = f"{sheet.title}_{cell.coordinate}_{cell_value}"
                        if cell_key in found_cells:
                            continue
                        found_cells.add(cell_key)

                        if any(kw in cell_value for kw in self.config['reservation']['verion_keywords']):
                            counts['verion'] += 1
                        elif any(kw in cell_value for kw in self.config['reservation']['lensx_keywords']):
                            counts['lensx'] += 1
                        elif any(kw in cell_value for kw in self.config['reservation']['ex500_keywords']):
                            counts['ex500'] += 1

            wb.close()

        except Exception as e:
            log_callback(f"  ❌ 예약 파일 처리 오류: {str(e)}")

        return counts

    def write_excel(self, output_path: str, staff_selected: List[str],
                   manual_fag: int, manual_glasses: int, manual_lasik: int,
                   reservation_counts: Dict[str, int], log_callback) -> bool:
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
                    ws.cell(cell_info['row'], cell_info['col']).value = len(chart_set)

            # 특수 항목 기입
            glaucoma_count = self.calculate_glaucoma(log_callback)
            glaucoma_cell = self.config['special_items']['녹내장']['cell']
            ws.cell(glaucoma_cell['row'], glaucoma_cell['col']).value = glaucoma_count

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
        self.setup_gui()

    def setup_gui(self):
        """GUI 구성 요소 생성"""
        self.root.title("일일결산 자동화 시스템 (최적화)")
        self.root.geometry("900x700")
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

        # 1. 근무 인원 선택
        staff_label = ttk.Label(left_frame, text="📋 근무 인원", font=("", 12, "bold"))
        staff_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))

        staff_canvas = tk.Canvas(left_frame, height=200)
        staff_scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=staff_canvas.yview)
        staff_scrollable = ttk.Frame(staff_canvas)

        staff_scrollable.bind(
            "<Configure>",
            lambda e: staff_canvas.configure(scrollregion=staff_canvas.bbox("all"))
        )

        staff_canvas.create_window((0, 0), window=staff_scrollable, anchor="nw")
        staff_canvas.configure(yscrollcommand=staff_scrollbar.set)

        staff_canvas.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        staff_scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S), pady=(0, 10))

        # 직원 체크박스 생성
        self.staff_vars = {}
        for i, staff_name in enumerate(self.system.config['staff_list']):
            var = tk.BooleanVar(value=True)
            self.staff_vars[staff_name] = var
            cb = ttk.Checkbutton(staff_scrollable, text=staff_name, variable=var)
            cb.grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)

        # 2. 예약 파일 선택
        ttk.Separator(left_frame, orient='horizontal').grid(row=2, column=0, columnspan=2,
                                                             sticky=(tk.W, tk.E), pady=10)

        reservation_label = ttk.Label(left_frame, text="📁 예약 파일", font=("", 12, "bold"))
        reservation_label.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))

        self.file_button = ttk.Button(left_frame, text="파일 선택...", command=self.select_files)
        self.file_button.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))

        self.file_label = ttk.Label(left_frame, text="선택된 파일: 없음", foreground="gray")
        self.file_label.grid(row=5, column=0, columnspan=2, sticky=tk.W)

        # 3. 수기 입력
        ttk.Separator(left_frame, orient='horizontal').grid(row=6, column=0, columnspan=2,
                                                             sticky=(tk.W, tk.E), pady=10)

        manual_label = ttk.Label(left_frame, text="✍ 수기 입력", font=("", 12, "bold"))
        manual_label.grid(row=7, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))

        lasik_label = ttk.Label(left_frame, text="라식:")
        lasik_label.grid(row=8, column=0, sticky=tk.W, padx=(0, 5))

        self.lasik_entry = ttk.Entry(left_frame, width=10)
        self.lasik_entry.insert(0, "0")
        self.lasik_entry.grid(row=8, column=1, sticky=tk.W, pady=3)

        fag_label = ttk.Label(left_frame, text="FAG:")
        fag_label.grid(row=9, column=0, sticky=tk.W, padx=(0, 5))

        self.fag_entry = ttk.Entry(left_frame, width=10)
        self.fag_entry.insert(0, "0")
        self.fag_entry.grid(row=9, column=1, sticky=tk.W, pady=3)

        glasses_label = ttk.Label(left_frame, text="안경검사:")
        glasses_label.grid(row=10, column=0, sticky=tk.W, padx=(0, 5))

        self.glasses_entry = ttk.Entry(left_frame, width=10)
        self.glasses_entry.insert(0, "0")
        self.glasses_entry.grid(row=10, column=1, sticky=tk.W, pady=3)

        # 4. 실행 버튼
        ttk.Separator(left_frame, orient='horizontal').grid(row=11, column=0, columnspan=2,
                                                             sticky=(tk.W, tk.E), pady=15)

        self.run_button = ttk.Button(left_frame, text="🚀 결산 실행", command=self.run_report)
        self.run_button.grid(row=12, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        # === 우측 영역 구성 ===

        log_label = ttk.Label(right_frame, text="실행 로그", font=("", 12, "bold"))
        log_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        self.log_text = scrolledtext.ScrolledText(right_frame, width=50, height=30,
                                                   state='disabled', wrap=tk.WORD)
        self.log_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    def log(self, message: str):
        """로그 메시지 출력"""
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')
        self.root.update()

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

    def run_report(self):
        """결산 실행"""
        self.run_button.config(state='disabled')
        self.file_button.config(state='disabled')

        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')

        thread = threading.Thread(target=self.process_report, daemon=True)
        thread.start()

    def process_report(self):
        """결산 처리 메인 로직"""
        try:
            self.log("=" * 54)
            self.log(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 결산 시작 (최적화 버전)")
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

            today_str = date.today().strftime('%Y%m%d')
            temp_excel = f"일일결산_{today_str}_temp.xlsx"

            success = self.system.write_excel(
                temp_excel, staff_selected, manual_fag, manual_glasses, manual_lasik,
                reservation_counts, self.log
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
