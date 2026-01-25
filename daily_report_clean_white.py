#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
일일결산 자동화 시스템 (Clean White UI)
안과 검사실의 일일 통계를 자동으로 수집하고 PDF 보고서를 생성하는 프로그램
UI 스타일: Clean White (미니멀 디자인)
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
    """일일결산 시스템의 메인 클래스"""

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
        """정규식 매칭에서 차트번호 추출"""
        groups = match.groups()
        if len(groups) == 2:
            return groups[0] if groups[0] else groups[1]
        return groups[0] if groups else None

    def scan_equipment_folder(self, eq_id: str, target_date: date, log_callback=None) -> Set[str]:
        """장비 폴더 스캔 (최적화 버전)"""
        chart_numbers = set()
        eq_info = self.config['equipment'][eq_id]
        base_path = eq_info['base_path']
        pattern = self.compiled_patterns[eq_id]

        if not os.path.exists(base_path):
            if log_callback:
                log_callback(f"⚠️  {eq_id}: 경로 없음 ({base_path})")
            return chart_numbers

        # 그룹 A vs B 처리
        group_a_equipments = ['SP', 'HFA', 'FUNDUS']

        if eq_id in group_a_equipments:
            # 그룹 A: 실시간 저장 → 저녁 정리
            date_folder = self.build_date_folder_path(base_path, eq_info['folder_structure'], target_date)

            if os.path.exists(date_folder):
                chart_numbers = self.scan_files_in_folder(date_folder, pattern, log_callback, f"{eq_id} (날짜폴더)")
            else:
                chart_numbers = self.scan_files_in_folder(base_path, pattern, log_callback, f"{eq_id} (최상위)")
        else:
            # 그룹 B: 자동 날짜 폴더
            date_folder = self.build_date_folder_path(base_path, eq_info['folder_structure'], target_date)

            if os.path.exists(date_folder):
                chart_numbers = self.scan_files_in_folder(date_folder, pattern, log_callback, f"{eq_id}")
            else:
                if log_callback:
                    log_callback(f"⚠️  {eq_id}: 날짜 폴더 없음")

        return chart_numbers

    def build_date_folder_path(self, base_path: str, folder_structure: str, target_date: date) -> str:
        """날짜 폴더 경로 생성"""
        formatted = folder_structure.replace('YYYY', str(target_date.year))
        formatted = formatted.replace('MM', f"{target_date.month:02d}")
        formatted = formatted.replace('DD', f"{target_date.day:02d}")
        return os.path.join(base_path, formatted)

    def scan_files_in_folder(self, folder: str, pattern, log_callback, label: str) -> Set[str]:
        """폴더 내 파일 스캔"""
        chart_numbers = set()

        try:
            for entry in os.scandir(folder):
                if entry.is_file():
                    match = pattern.search(entry.name)
                    if match:
                        chart_num = self.extract_chart_number(match)
                        if chart_num and self.is_valid_chart_number(chart_num):
                            chart_numbers.add(chart_num)
        except Exception as e:
            if log_callback:
                log_callback(f"⚠️  {label}: 스캔 오류 - {e}")

        return chart_numbers

    def scan_all_equipment(self, target_date: date, log_callback=None) -> Dict[str, Set[str]]:
        """모든 장비 스캔 (병렬 처리)"""
        results = {}

        with ThreadPoolExecutor(max_workers=6) as executor:
            futures = {
                executor.submit(self.scan_equipment_folder, eq_id, target_date, log_callback): eq_id
                for eq_id in self.config['equipment'].keys()
            }

            for future in as_completed(futures):
                eq_id = futures[future]
                try:
                    results[eq_id] = future.result()
                except Exception as e:
                    if log_callback:
                        log_callback(f"⚠️  {eq_id}: 오류 - {e}")
                    results[eq_id] = set()

        return results

    def calculate_glaucoma(self, hfa_set: Set[str], oct_set: Set[str]) -> int:
        """녹내장 = HFA ∩ OCT"""
        return len(hfa_set & oct_set)

    def scan_fundus_folder(self, target_date: date, log_callback=None) -> Set[str]:
        """안저 스캔"""
        all_files = set()

        for folder_id in ['fundus_main', 'fundus_secondary']:
            folder_info = self.config['fundus_folders'].get(folder_id)
            if not folder_info:
                continue

            base_path = folder_info['base_path']
            pattern = re.compile(folder_info['pattern'])

            date_folder = self.build_date_folder_path(base_path, folder_info['folder_structure'], target_date)

            if os.path.exists(date_folder):
                files = self.scan_files_in_folder(date_folder, pattern, log_callback, f"안저-{folder_id}")
                all_files.update(files)

        return all_files

    def scan_reservation_files(self, file_paths: List[str], target_date: date, log_callback=None) -> Dict[str, int]:
        """예약 파일 스캔"""
        counts = {'verion': 0, 'lensx': 0, 'ex500': 0}
        keywords = self.config.get('reservation_keywords', {})

        for file_path in file_paths:
            try:
                df = pd.read_excel(file_path)
                date_str = target_date.strftime('%Y-%m-%d')

                for col in df.columns:
                    for idx, cell in enumerate(df[col]):
                        if pd.notna(cell) and date_str in str(cell):
                            for key, keyword_list in keywords.items():
                                for keyword in keyword_list:
                                    if keyword in str(cell):
                                        counts[key] += 1
            except Exception as e:
                if log_callback:
                    log_callback(f"⚠️  예약 파일 오류: {e}")

        return counts


class CleanWhiteGUI:
    """Clean White UI (미니멀 디자인)"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.system = DailyReportSystem()
        self.reservation_files = []
        self.scan_results = {}
        self.log_file_handle = None

        self.setup_gui()

    def setup_gui(self):
        """GUI 구성"""
        self.root.title("일일결산 자동화 시스템 (Clean White)")
        self.root.geometry("1200x700")
        self.root.configure(bg='#fafafa')

        # 2단 레이아웃
        main_frame = tk.Frame(self.root, bg='#fafafa')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 왼쪽: 입력 영역
        left_frame = tk.Frame(main_frame, bg='#ffffff', relief=tk.FLAT, bd=1)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 10))

        # 오른쪽: 결과 영역
        right_frame = tk.Frame(main_frame, bg='#fafafa')
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # === 왼쪽 패널 ===
        self.create_left_panel(left_frame)

        # === 오른쪽 패널 ===
        self.create_right_panel(right_frame)

    def create_left_panel(self, parent):
        """왼쪽 입력 패널"""
        parent.configure(width=300, bg='#ffffff')

        inner = tk.Frame(parent, bg='#ffffff')
        inner.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)

        # 제목
        title_label = tk.Label(
            inner,
            text="결산 정보",
            font=("Segoe UI", 11, "bold"),
            fg="#11998e",
            bg='#ffffff'
        )
        title_label.pack(anchor='w', pady=(0, 15))

        # 구분선
        sep1 = tk.Frame(inner, height=2, bg='#11998e')
        sep1.pack(fill=tk.X, pady=(0, 20))

        # 날짜
        date_label = tk.Label(inner, text="날짜", font=("Segoe UI", 13), fg="#333", bg='#ffffff')
        date_label.pack(anchor='w', pady=(0, 6))

        self.date_entry = tk.Entry(
            inner,
            font=("Segoe UI", 14),
            relief=tk.FLAT,
            bg='#ffffff',
            fg='#333',
            bd=0,
            highlightthickness=0
        )
        self.date_entry.insert(0, date.today().strftime('%Y-%m-%d'))
        self.date_entry.pack(fill=tk.X, pady=(0, 5))

        date_underline = tk.Frame(inner, height=2, bg='#e0e0e0')
        date_underline.pack(fill=tk.X, pady=(0, 18))

        # 수기 입력 섹션
        manual_title = tk.Label(
            inner,
            text="수기 입력",
            font=("Segoe UI", 11, "bold"),
            fg="#11998e",
            bg='#ffffff'
        )
        manual_title.pack(anchor='w', pady=(10, 15))

        sep2 = tk.Frame(inner, height=2, bg='#11998e')
        sep2.pack(fill=tk.X, pady=(0, 20))

        # 라식
        self.create_input_field(inner, "라식", "lasik_entry")

        # FAG
        self.create_input_field(inner, "FAG", "fag_entry")

        # 안경검사
        self.create_input_field(inner, "안경검사", "glasses_entry")

        # OCTS
        self.create_input_field(inner, "OCTS", "octs_entry")

        # 버튼들
        self.scan_button = tk.Button(
            inner,
            text="스캔 시작",
            font=("Segoe UI", 14, "bold"),
            bg='#11998e',
            fg='white',
            relief=tk.FLAT,
            bd=0,
            cursor='hand2',
            command=self.run_scan
        )
        self.scan_button.pack(fill=tk.X, pady=(30, 10))
        self.scan_button.configure(height=2)

        self.output_button = tk.Button(
            inner,
            text="확정 및 PDF 출력",
            font=("Segoe UI", 14, "bold"),
            bg='#667eea',
            fg='white',
            relief=tk.FLAT,
            bd=0,
            cursor='hand2',
            state='disabled',
            command=self.run_output
        )
        self.output_button.pack(fill=tk.X, pady=(0, 10))
        self.output_button.configure(height=2)

    def create_input_field(self, parent, label_text, attr_name):
        """입력 필드 생성"""
        label = tk.Label(parent, text=label_text, font=("Segoe UI", 13), fg="#333", bg='#ffffff')
        label.pack(anchor='w', pady=(0, 6))

        entry = tk.Entry(
            parent,
            font=("Segoe UI", 14),
            relief=tk.FLAT,
            bg='#ffffff',
            fg='#333',
            bd=0,
            highlightthickness=0
        )
        entry.insert(0, "0")
        entry.pack(fill=tk.X, pady=(0, 5))
        setattr(self, attr_name, entry)

        underline = tk.Frame(parent, height=2, bg='#e0e0e0')
        underline.pack(fill=tk.X, pady=(0, 18))

    def create_right_panel(self, parent):
        """오른쪽 결과 패널"""
        # 제목
        title_label = tk.Label(
            parent,
            text="검사 결과",
            font=("Segoe UI", 11, "bold"),
            fg="#11998e",
            bg='#fafafa'
        )
        title_label.pack(anchor='w', pady=(0, 10))

        # "수정 가능" 배지
        badge = tk.Label(
            parent,
            text="수정 가능",
            font=("Segoe UI", 10, "bold"),
            fg='white',
            bg='#11998e',
            padx=8,
            pady=2
        )
        badge.place(x=250, y=0)

        sep = tk.Frame(parent, height=2, bg='#11998e')
        sep.pack(fill=tk.X, pady=(0, 20))

        # 스크롤 가능한 그리드
        canvas = tk.Canvas(parent, bg='#fafafa', highlightthickness=0)
        scrollbar = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#fafafa')

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 3열 그리드
        self.result_cards = {}
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
            row = idx // 3
            col = idx % 3

            card = self.create_result_card(scrollable_frame, key, label_text)
            card.grid(row=row, column=col, padx=10, pady=10, sticky='nsew')

            # 그리드 가중치 설정
            scrollable_frame.grid_columnconfigure(col, weight=1)

    def create_result_card(self, parent, key, label_text):
        """결과 카드 생성"""
        card = tk.Frame(parent, bg='white', relief=tk.FLAT, bd=1, highlightbackground='#e0e0e0', highlightthickness=1)
        card.configure(width=180, height=100)

        inner = tk.Frame(card, bg='white')
        inner.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 라벨
        label = tk.Label(
            inner,
            text=label_text.upper(),
            font=("Segoe UI", 11),
            fg='#888',
            bg='white'
        )
        label.pack(anchor='w', pady=(0, 8))

        # 입력 필드
        value_entry = tk.Entry(
            inner,
            font=("Segoe UI", 28, "bold"),
            fg='#11998e',
            bg='white',
            relief=tk.FLAT,
            bd=0,
            justify='center',
            state='disabled',
            disabledforeground='#cccccc'
        )
        value_entry.insert(0, "0")
        value_entry.pack(fill=tk.X)

        # 구분선
        underline = tk.Frame(inner, height=2, bg='#11998e')
        underline.pack(fill=tk.X, pady=(5, 0))
        underline.pack_forget()  # 초기에는 숨김

        self.result_cards[key] = {
            'entry': value_entry,
            'underline': underline
        }

        return card

    def run_scan(self):
        """스캔 실행"""
        self.scan_button.config(state='disabled', text='스캔 중...', bg='#999999')
        self.output_button.config(state='disabled')

        # 카드 비활성화
        for key, widgets in self.result_cards.items():
            widgets['entry'].config(state='disabled', disabledforeground='#cccccc')
            widgets['underline'].pack_forget()

        def scan_thread():
            try:
                target_date = datetime.strptime(self.date_entry.get(), '%Y-%m-%d').date()

                # 장비 스캔
                self.system.chart_numbers = self.system.scan_all_equipment(target_date, None)

                # 녹내장 계산
                glaucoma_count = self.system.calculate_glaucoma(
                    self.system.chart_numbers.get('HFA', set()),
                    self.system.chart_numbers.get('OCT', set())
                )

                # 안저 스캔
                fundus_set = self.system.scan_fundus_folder(target_date, None)

                # 예약 파일 스캔
                reservation_counts = {'verion': 0, 'lensx': 0, 'ex500': 0}
                if self.reservation_files:
                    reservation_counts = self.system.scan_reservation_files(
                        self.reservation_files, target_date, None
                    )

                self.scan_results = {
                    'glaucoma_count': glaucoma_count,
                    'fundus_count': len(fundus_set),
                    'reservation_counts': reservation_counts
                }

                # UI 업데이트
                self.root.after(0, self.update_result_cards)

            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("오류", f"스캔 오류: {e}"))
                self.root.after(0, lambda: self.scan_button.config(state='normal', text='스캔 시작', bg='#11998e'))

        threading.Thread(target=scan_thread, daemon=True).start()

    def update_result_cards(self):
        """결과 카드 업데이트"""
        # 값 설정
        values = {
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

        # 카드 업데이트
        for key, value in values.items():
            widgets = self.result_cards[key]
            entry = widgets['entry']
            underline = widgets['underline']

            entry.config(state='normal', disabledforeground='#11998e')
            entry.delete(0, tk.END)
            entry.insert(0, str(value))
            underline.pack(fill=tk.X, pady=(5, 0))

        # 버튼 활성화
        self.scan_button.config(state='normal', text='스캔 시작', bg='#11998e')
        self.output_button.config(state='normal', bg='#667eea')

    def run_output(self):
        """PDF 출력"""
        messagebox.showinfo("안내", "PDF 출력 기능은 원본 파일의 로직을 사용합니다.\n결과 값을 확인하고 수정하셨습니다!")

    def run(self):
        """GUI 실행"""
        self.root.mainloop()


def main():
    root = tk.Tk()
    app = CleanWhiteGUI(root)
    app.run()


if __name__ == "__main__":
    main()
