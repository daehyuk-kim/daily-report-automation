#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
디렉토리별 파일명/폴더명 패턴 분석 도구
각 장비 디렉토리의 실제 파일을 샘플링하여 패턴을 추출합니다.
"""

import os
import sys
import json
import re
from datetime import date, datetime
from collections import Counter

def load_config(config_path="config_real.json"):
    """설정 파일 로드"""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"❌ 설정 파일 없음: {config_path}")
        sys.exit(1)

def guess_today_folder(base_path, equipment_id):
    """오늘 날짜 폴더 경로 추측"""
    today = date.today()

    year = today.strftime('%Y')
    month = today.strftime('%m')
    day = today.strftime('%d')

    # 가능한 패턴들
    patterns = []

    if equipment_id == 'TOPO':
        patterns.append(os.path.join(base_path, year, month, f"TOPO {month}.{day}"))
    elif equipment_id == 'ORB':
        patterns.append(os.path.join(base_path, year, f"{year}.{month}", f"ORB {month}.{day}"))
    elif equipment_id == 'OCT':
        patterns.append(os.path.join(base_path, year, month, day))
    elif equipment_id == 'OQAS':
        patterns.append(os.path.join(base_path, year, month, f"{day}.{month}"))
    else:
        # SP, HFA, IOL700 등 단일 폴더
        return base_path

    # 존재하는 경로 찾기
    for path in patterns:
        if os.path.exists(path):
            return path

    return base_path

def analyze_directory(equipment_id, equipment_info, max_samples=30):
    """디렉토리 분석"""
    print("=" * 80)
    print(f"📁 {equipment_info['name']} ({equipment_id})")
    print("=" * 80)

    base_path = equipment_info['path']

    # 경로 존재 확인
    if not os.path.exists(base_path):
        print(f"❌ 경로 없음: {base_path}")
        print()
        return None

    print(f"✅ 기본 경로: {base_path}")

    # 오늘 폴더 추측
    today_folder = guess_today_folder(base_path, equipment_id)

    if today_folder != base_path:
        print(f"📂 오늘 폴더: {today_folder}")
        if not os.path.exists(today_folder):
            print(f"⚠️  오늘 폴더 없음 (기본 경로 사용)")
            today_folder = base_path
    else:
        print(f"📂 단일 폴더 구조")

    print()

    # 샘플 수집
    file_samples = []
    dir_samples = []

    try:
        items = os.listdir(today_folder)

        for item in items[:100]:  # 최대 100개만
            item_path = os.path.join(today_folder, item)

            if os.path.isfile(item_path):
                # 파일 생성 날짜 확인
                try:
                    ctime = os.path.getctime(item_path)
                    file_date = date.fromtimestamp(ctime)
                    file_samples.append({
                        'name': item,
                        'date': file_date.strftime('%Y-%m-%d'),
                        'is_today': file_date == date.today()
                    })
                except:
                    file_samples.append({
                        'name': item,
                        'date': 'Unknown',
                        'is_today': False
                    })

            elif os.path.isdir(item_path):
                dir_samples.append(item)

            if len(file_samples) >= max_samples:
                break

    except Exception as e:
        print(f"❌ 스캔 오류: {e}")
        print()
        return None

    # 파일 분석
    if file_samples:
        print(f"📄 파일 샘플 (총 {len(file_samples)}개)")
        print("-" * 80)

        # 오늘 파일만 필터
        today_files = [f for f in file_samples if f['is_today']]

        if today_files:
            print(f"🟢 오늘 생성된 파일: {len(today_files)}개")
            print()

            for i, file_info in enumerate(today_files[:10], 1):
                print(f"{i:2d}. {file_info['name']}")

            if len(today_files) > 10:
                print(f"    ... 외 {len(today_files) - 10}개")
        else:
            print("⚠️  오늘 생성된 파일 없음. 전체 샘플 표시:")
            print()

            for i, file_info in enumerate(file_samples[:10], 1):
                print(f"{i:2d}. {file_info['name']} ({file_info['date']})")

            if len(file_samples) > 10:
                print(f"    ... 외 {len(file_samples) - 10}개")

        print()

        # 패턴 분석
        print("🔍 차트번호 패턴 분석")
        print("-" * 80)

        patterns_to_test = [
            (r'\s(\d+)_', '공백 + 숫자 + 언더스코어'),
            (r'\s(\d+)-', '공백 + 숫자 + 하이픈'),
            (r'\s(\d+)\s', '공백 + 숫자 + 공백'),
            (r'_(\d+)\.', '언더스코어 + 숫자 + 점'),
            (r'^(\d+)_', '시작 + 숫자 + 언더스코어'),
            (r'__(\\d+)_', '언더스코어2개 + 숫자 + 언더스코어'),
            (r'[a-zA-Z,\s]+\s(\d+)\s', '문자/공백 + 숫자 + 공백'),
        ]

        matched_patterns = []

        for pattern, description in patterns_to_test:
            regex = re.compile(pattern)
            matches = 0
            sample_matches = []

            for file_info in (today_files if today_files else file_samples[:20]):
                match = regex.search(file_info['name'])
                if match:
                    matches += 1
                    chart_num = match.group(1)
                    if len(sample_matches) < 3:
                        sample_matches.append((file_info['name'], chart_num))

            if matches > 0:
                matched_patterns.append({
                    'pattern': pattern,
                    'description': description,
                    'matches': matches,
                    'samples': sample_matches
                })

        if matched_patterns:
            # 매칭 수가 많은 순으로 정렬
            matched_patterns.sort(key=lambda x: x['matches'], reverse=True)

            print(f"발견된 패턴: {len(matched_patterns)}개")
            print()

            for i, p in enumerate(matched_patterns[:3], 1):
                print(f"{i}. 패턴: {p['pattern']}")
                print(f"   설명: {p['description']}")
                print(f"   매칭: {p['matches']}개")
                print(f"   샘플:")
                for fname, chart_num in p['samples']:
                    print(f"     - {fname} → 차트번호: {chart_num}")
                print()
        else:
            print("❌ 차트번호 패턴을 찾지 못했습니다")
            print()

        # 날짜 패턴 분석
        print("📅 날짜 패턴 분석")
        print("-" * 80)

        date_patterns = [
            (r'_(\d{8})\.', 'YYYYMMDD (언더스코어 + 8자리 + 점)'),
            (r'_(\d{6})\.', 'YYMMDD (언더스코어 + 6자리 + 점)'),
            (r'(\d{4}-\d{2}-\d{2})', 'YYYY-MM-DD'),
            (r'(\d{4}\.\d{2}\.\d{2})', 'YYYY.MM.DD'),
        ]

        found_date_pattern = False
        for pattern, description in date_patterns:
            regex = re.compile(pattern)
            matches = 0

            for file_info in file_samples[:20]:
                if regex.search(file_info['name']):
                    matches += 1

            if matches > 0:
                print(f"✓ {description}: {matches}개 파일 매칭")
                found_date_pattern = True

        if not found_date_pattern:
            print("⚠️  파일명에 날짜 패턴 없음 → 파일 생성일 확인 필요")

        print()

    # 폴더 분석 (OCT의 경우)
    if dir_samples and equipment_info.get('scan_type') == 'both':
        print(f"📁 폴더 샘플 (총 {len(dir_samples)}개)")
        print("-" * 80)

        for i, dirname in enumerate(dir_samples[:10], 1):
            print(f"{i:2d}. {dirname}")

        if len(dir_samples) > 10:
            print(f"    ... 외 {len(dir_samples) - 10}개")

        print()

    print("=" * 80)
    print()

    return {
        'equipment_id': equipment_id,
        'name': equipment_info['name'],
        'base_path': base_path,
        'today_folder': today_folder,
        'file_count': len(file_samples),
        'today_file_count': len([f for f in file_samples if f['is_today']]),
        'matched_patterns': matched_patterns if file_samples else [],
        'has_date_in_filename': found_date_pattern if file_samples else False,
    }

def main():
    """메인 함수"""
    print()
    print("=" * 80)
    print("📊 디렉토리별 파일명/폴더명 패턴 분석 도구")
    print("=" * 80)
    print()

    # 설정 파일 로드
    config = load_config()

    print(f"설정 파일: config_real.json")
    print(f"분석 대상: {len(config['equipment'])}개 장비")
    print()

    # 각 장비별 분석
    results = []

    for equipment_id, equipment_info in config['equipment'].items():
        result = analyze_directory(equipment_id, equipment_info)
        if result:
            results.append(result)

    # 결과 요약
    print()
    print("=" * 80)
    print("📋 분석 결과 요약")
    print("=" * 80)
    print()

    for result in results:
        print(f"🔧 {result['name']} ({result['equipment_id']})")
        print(f"   오늘 파일: {result['today_file_count']}개")

        if result['matched_patterns']:
            best = result['matched_patterns'][0]
            print(f"   추천 패턴: {best['pattern']}")
            print(f"   매칭률: {best['matches']}/{result['file_count']}개")

        if result['has_date_in_filename']:
            print(f"   날짜 확인: 파일명")
        else:
            print(f"   날짜 확인: 파일 생성일")

        print()

    # 결과 저장 옵션
    save = input("분석 결과를 파일로 저장하시겠습니까? (y/n): ").strip().lower()
    if save == 'y':
        output_file = f"directory_analysis_{date.today().strftime('%Y%m%d')}.txt"

        try:
            import io
            from contextlib import redirect_stdout

            # 다시 실행해서 파일로 저장
            f = io.StringIO()
            with redirect_stdout(f):
                for equipment_id, equipment_info in config['equipment'].items():
                    analyze_directory(equipment_id, equipment_info)

            with open(output_file, 'w', encoding='utf-8') as file:
                file.write(f.getvalue())

            print(f"✅ 저장 완료: {output_file}")
        except Exception as e:
            print(f"❌ 저장 실패: {e}")

if __name__ == "__main__":
    main()
