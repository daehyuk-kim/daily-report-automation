#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
폴더 구조 분석 도구
백업 디렉토리의 폴더 구조를 분석하고 시각화합니다.
"""

import os
import sys
from datetime import datetime, date
from pathlib import Path

class FolderAnalyzer:
    """폴더 구조 분석 클래스"""

    def __init__(self, output_file="folder_structure_report.txt"):
        self.output_file = output_file
        self.today = date.today()
        self.report_lines = []

    def log(self, message, print_console=True):
        """로그 메시지 추가"""
        self.report_lines.append(message)
        if print_console:
            print(message)

    def analyze_directory(self, directory_path, max_depth=3, show_files=False):
        """
        디렉토리 구조 분석

        Args:
            directory_path: 분석할 디렉토리 경로
            max_depth: 최대 탐색 깊이
            show_files: 파일도 표시할지 여부
        """
        self.log("=" * 80)
        self.log(f"폴더 구조 분석 보고서")
        self.log(f"생성 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.log("=" * 80)
        self.log("")

        if not os.path.exists(directory_path):
            self.log(f"❌ 오류: 경로가 존재하지 않습니다: {directory_path}")
            return

        self.log(f"📁 분석 대상: {directory_path}")
        self.log(f"📅 오늘 날짜: {self.today.strftime('%Y-%m-%d')}")
        self.log(f"📊 최대 탐색 깊이: {max_depth}")
        self.log("")

        # 폴더 구조 트리 표시
        self.log("폴더 구조:")
        self.log("-" * 80)
        self._scan_directory(directory_path, depth=0, max_depth=max_depth,
                            prefix="", show_files=show_files)
        self.log("")

        # 오늘 날짜 관련 폴더 찾기
        self.log("=" * 80)
        self.log("오늘 날짜 관련 폴더 검색:")
        self.log("-" * 80)
        self._find_today_folders(directory_path, max_depth)
        self.log("")

        # 날짜별 폴더 패턴 분석
        self.log("=" * 80)
        self.log("폴더 명명 패턴 분석:")
        self.log("-" * 80)
        self._analyze_naming_patterns(directory_path, max_depth)
        self.log("")

    def _scan_directory(self, path, depth, max_depth, prefix, show_files):
        """재귀적으로 디렉토리 스캔"""
        if depth > max_depth:
            return

        try:
            items = sorted(os.listdir(path))
        except PermissionError:
            self.log(f"{prefix}[접근 거부]")
            return
        except Exception as e:
            self.log(f"{prefix}[오류: {e}]")
            return

        # 폴더와 파일 분리
        folders = []
        files = []

        for item in items:
            item_path = os.path.join(path, item)
            if os.path.isdir(item_path):
                folders.append(item)
            elif show_files:
                files.append(item)

        # 폴더 표시
        for i, folder in enumerate(folders):
            is_last_folder = (i == len(folders) - 1) and (not show_files or len(files) == 0)
            connector = "└── " if is_last_folder else "├── "

            folder_path = os.path.join(path, folder)

            # 생성 날짜 확인
            try:
                ctime = os.path.getctime(folder_path)
                folder_date = date.fromtimestamp(ctime)
                date_str = folder_date.strftime('%Y-%m-%d')
                is_today = folder_date == self.today
                date_marker = " 🟢 [오늘]" if is_today else f" ({date_str})"
            except:
                date_marker = ""

            # 하위 항목 개수
            try:
                sub_items = os.listdir(folder_path)
                item_count = len(sub_items)
                count_str = f" [{item_count}개]"
            except:
                count_str = ""

            self.log(f"{prefix}{connector}📁 {folder}{count_str}{date_marker}")

            # 재귀 호출
            if depth < max_depth:
                extension = "    " if is_last_folder else "│   "
                self._scan_directory(folder_path, depth + 1, max_depth,
                                   prefix + extension, show_files)

        # 파일 표시 (선택적)
        if show_files:
            for i, file in enumerate(files):
                is_last = i == len(files) - 1
                connector = "└── " if is_last else "├── "

                file_path = os.path.join(path, file)
                try:
                    file_size = os.path.getsize(file_path)
                    size_str = self._format_size(file_size)

                    ctime = os.path.getctime(file_path)
                    file_date = date.fromtimestamp(ctime)
                    is_today = file_date == self.today
                    date_marker = " 🟢" if is_today else ""

                    self.log(f"{prefix}{connector}📄 {file} ({size_str}){date_marker}")
                except:
                    self.log(f"{prefix}{connector}📄 {file}")

    def _find_today_folders(self, path, max_depth):
        """오늘 날짜 관련 폴더 찾기"""
        today_folders = []

        def search(current_path, depth):
            if depth > max_depth:
                return

            try:
                items = os.listdir(current_path)
            except:
                return

            for item in items:
                item_path = os.path.join(current_path, item)

                if not os.path.isdir(item_path):
                    continue

                # 생성 날짜 확인
                try:
                    ctime = os.path.getctime(item_path)
                    folder_date = date.fromtimestamp(ctime)

                    if folder_date == self.today:
                        rel_path = os.path.relpath(item_path, path)
                        today_folders.append((rel_path, item_path))
                except:
                    pass

                # 폴더명에 오늘 날짜 포함 여부 확인
                today_patterns = [
                    self.today.strftime('%Y%m%d'),
                    self.today.strftime('%Y.%m.%d'),
                    self.today.strftime('%Y-%m-%d'),
                    self.today.strftime('%m.%d'),
                    self.today.strftime('%m-%d'),
                ]

                for pattern in today_patterns:
                    if pattern in item:
                        rel_path = os.path.relpath(item_path, path)
                        if (rel_path, item_path) not in today_folders:
                            today_folders.append((rel_path, item_path))
                        break

                # 재귀 검색
                search(item_path, depth + 1)

        search(path, 0)

        if today_folders:
            self.log(f"발견된 오늘 날짜 폴더: {len(today_folders)}개\n")
            for rel_path, full_path in today_folders:
                try:
                    item_count = len(os.listdir(full_path))
                    self.log(f"  📁 {rel_path}")
                    self.log(f"     전체 경로: {full_path}")
                    self.log(f"     하위 항목: {item_count}개")
                    self.log("")
                except:
                    self.log(f"  📁 {rel_path}")
                    self.log(f"     전체 경로: {full_path}")
                    self.log("")
        else:
            self.log("오늘 날짜 관련 폴더를 찾지 못했습니다.")

    def _analyze_naming_patterns(self, path, max_depth):
        """폴더 명명 패턴 분석"""
        patterns = {
            'YYYY': [],
            'YYYY\\MM': [],
            'YYYY\\MM\\DD': [],
            'YYYY.MM': [],
            'YYYY.MM.DD': [],
            'MM.DD': [],
            '장비명 MM.DD': [],
            '기타': []
        }

        def search(current_path, depth, rel_path=""):
            if depth > max_depth:
                return

            try:
                items = os.listdir(current_path)
            except:
                return

            for item in items:
                item_path = os.path.join(current_path, item)

                if not os.path.isdir(item_path):
                    continue

                current_rel = os.path.join(rel_path, item) if rel_path else item

                # 패턴 매칭
                if item.isdigit() and len(item) == 4:
                    patterns['YYYY'].append(current_rel)
                elif item.isdigit() and len(item) == 2:
                    patterns['YYYY\\MM'].append(current_rel)
                elif '.' in item:
                    parts = item.split('.')
                    if len(parts) == 2 and all(p.isdigit() for p in parts):
                        if len(parts[0]) == 4:
                            patterns['YYYY.MM'].append(current_rel)
                        else:
                            patterns['MM.DD'].append(current_rel)
                    elif len(parts) == 3 and all(p.isdigit() for p in parts):
                        patterns['YYYY.MM.DD'].append(current_rel)
                    else:
                        # 장비명 포함 패턴 (예: TOPO 01.18)
                        if any(char.isalpha() for char in item):
                            patterns['장비명 MM.DD'].append(current_rel)
                        else:
                            patterns['기타'].append(current_rel)
                else:
                    patterns['기타'].append(current_rel)

                # 재귀 검색
                search(item_path, depth + 1, current_rel)

        search(path, 0)

        # 패턴별 통계
        for pattern_name, folders in patterns.items():
            if folders:
                self.log(f"{pattern_name} 패턴: {len(folders)}개")
                for folder in folders[:5]:  # 최대 5개만 표시
                    self.log(f"  - {folder}")
                if len(folders) > 5:
                    self.log(f"  ... 외 {len(folders) - 5}개")
                self.log("")

    def _format_size(self, size_bytes):
        """파일 크기를 읽기 쉽게 포맷"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f}{unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f}TB"

    def save_report(self):
        """보고서를 파일로 저장"""
        try:
            with open(self.output_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(self.report_lines))
            print(f"\n보고서가 저장되었습니다: {self.output_file}")
        except Exception as e:
            print(f"보고서 저장 중 오류: {e}")


def main():
    """메인 함수"""
    print("=" * 80)
    print("폴더 구조 분석 도구")
    print("=" * 80)
    print()

    # 디렉토리 경로 입력
    if len(sys.argv) > 1:
        directory_path = sys.argv[1]
    else:
        print("분석할 디렉토리 경로를 입력하세요:")
        print("예시:")
        print("  - Windows: D:\\BACKUP\\SP")
        print("  - Windows (네트워크): \\\\192.168.0.120\\sp")
        print("  - Mac/Linux: /Users/username/backup")
        print()
        directory_path = input("경로: ").strip()

        # 따옴표 제거
        directory_path = directory_path.strip('"').strip("'")

    if not directory_path:
        print("경로가 입력되지 않았습니다.")
        return

    # 옵션 설정
    print()
    print("분석 옵션:")
    print("1. 최대 탐색 깊이 (기본값: 3)")
    depth_input = input("깊이 (엔터: 기본값): ").strip()
    max_depth = int(depth_input) if depth_input.isdigit() else 3

    print("2. 파일도 표시하시겠습니까? (y/n, 기본값: n)")
    show_files_input = input("선택: ").strip().lower()
    show_files = show_files_input == 'y'

    print("3. 출력 파일명 (기본값: folder_structure_report.txt)")
    output_input = input("파일명: ").strip()
    output_file = output_input if output_input else "folder_structure_report.txt"

    print()
    print("분석을 시작합니다...")
    print()

    # 분석 실행
    analyzer = FolderAnalyzer(output_file)
    analyzer.analyze_directory(directory_path, max_depth, show_files)
    analyzer.save_report()


if __name__ == "__main__":
    main()
