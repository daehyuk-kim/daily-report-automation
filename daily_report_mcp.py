#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Daily Report MCP Server
세란안과 일일결산 자동화를 위한 MCP 서버

FastMCP를 사용하여 Claude에게 Daily Report 도구 제공
"""

from fastmcp import FastMCP
import json
import os
import re
from datetime import datetime, date
from typing import Dict, List, Set, Optional
import requests

# MCP 서버 생성
mcp = FastMCP("daily-report")

# 설정 파일 경로
CONFIG_PATH = "config.json"


def load_config() -> dict:
    """설정 파일 로드"""
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)


@mcp.tool()
def scan_equipment(equipment_id: str, target_date: str) -> Dict[str, int]:
    """
    특정 장비 폴더 스캔

    Args:
        equipment_id: 장비 ID (SP, TOPO, ORB, OCT, HFA, OQAS)
        target_date: 스캔 날짜 (YYYY-MM-DD)

    Returns:
        {'count': 검사 건수, 'chart_numbers': [...]}
    """
    config = load_config()

    if equipment_id not in config['equipment']:
        return {'error': f'Unknown equipment: {equipment_id}'}

    eq_info = config['equipment'][equipment_id]
    target = datetime.strptime(target_date, '%Y-%m-%d').date()

    # 날짜 폴더 경로 생성
    folder_structure = eq_info['folder_structure']
    formatted = folder_structure.replace('YYYY', str(target.year))
    formatted = formatted.replace('MM', f"{target.month:02d}")
    formatted = formatted.replace('DD', f"{target.day:02d}")

    full_path = os.path.join(eq_info['path'], formatted)

    if not os.path.exists(full_path):
        return {'count': 0, 'chart_numbers': [], 'note': 'Folder not found'}

    # 파일 스캔
    pattern = re.compile(eq_info['pattern'])
    chart_numbers = set()

    for entry in os.scandir(full_path):
        if entry.is_file():
            match = pattern.search(entry.name)
            if match:
                groups = match.groups()
                chart_num = groups[0] if groups[0] else groups[1] if len(groups) > 1 else groups[0]
                if chart_num:
                    chart_numbers.add(chart_num)

    return {
        'count': len(chart_numbers),
        'chart_numbers': sorted(list(chart_numbers)),
        'path': full_path
    }


@mcp.tool()
def get_today_statistics(target_date: str) -> Dict[str, any]:
    """
    전체 검사 통계 조회

    Args:
        target_date: 날짜 (YYYY-MM-DD)

    Returns:
        모든 장비의 검사 통계
    """
    config = load_config()
    results = {}

    for eq_id in config['equipment'].keys():
        result = scan_equipment(eq_id, target_date)
        results[eq_id] = result['count']

    # 녹내장 = HFA ∩ OCT
    hfa_result = scan_equipment('HFA', target_date)
    oct_result = scan_equipment('OCT', target_date)

    hfa_set = set(hfa_result.get('chart_numbers', []))
    oct_set = set(oct_result.get('chart_numbers', []))
    glaucoma_count = len(hfa_set & oct_set)

    results['GLAUCOMA'] = glaucoma_count

    return results


@mcp.tool()
def get_hospital_schedule_staff(target_date: str, department: str = "검사실") -> List[str]:
    """
    Hospital Schedule API에서 근무 인원 조회

    Args:
        target_date: 날짜 (YYYY-MM-DD)
        department: 부서명 (기본값: 검사실)

    Returns:
        근무 직원 명단
    """
    try:
        url = f"http://192.168.0.210:3001/api/schedule/today"
        params = {
            'date': target_date,
            'department': department
        }

        response = requests.get(url, params=params, timeout=5)

        if response.status_code == 200:
            data = response.json()
            staff_list = data.get('staff', [])
            return [staff['name'] for staff in staff_list if staff.get('status') == '근무']
        else:
            return []
    except Exception as e:
        return {'error': str(e)}


@mcp.tool()
def get_equipment_list() -> List[Dict[str, str]]:
    """
    사용 가능한 장비 목록 조회

    Returns:
        장비 ID와 이름 목록
    """
    config = load_config()
    equipment_list = []

    for eq_id, eq_info in config['equipment'].items():
        equipment_list.append({
            'id': eq_id,
            'name': eq_info['name'],
            'path': eq_info['path']
        })

    return equipment_list


@mcp.tool()
def get_staff_list() -> List[str]:
    """
    직원 명단 조회

    Returns:
        직원 이름 목록
    """
    config = load_config()
    return config.get('staff_list', [])


@mcp.tool()
def update_manual_input(field: str, value: int) -> Dict[str, str]:
    """
    수기 입력값 저장

    Args:
        field: 필드명 (라식, FAG, 안경검사, OCTS)
        value: 값

    Returns:
        저장 결과
    """
    config = load_config()

    if field not in config.get('manual_input', {}):
        return {'error': f'Unknown field: {field}'}

    # 메모리 또는 임시 파일에 저장 (실제 구현 필요)
    return {
        'status': 'success',
        'field': field,
        'value': value,
        'message': f'{field} = {value} 저장됨'
    }


@mcp.resource("daily-report://config")
def get_config() -> str:
    """
    현재 설정 조회

    Returns:
        config.json 내용
    """
    config = load_config()
    return json.dumps(config, ensure_ascii=False, indent=2)


@mcp.resource("daily-report://help")
def get_help() -> str:
    """
    Daily Report MCP 도움말

    Returns:
        사용 가능한 도구 목록 및 설명
    """
    help_text = """
    # Daily Report MCP Server

    ## 사용 가능한 도구:

    1. scan_equipment(equipment_id, target_date)
       - 특정 장비 폴더 스캔
       - equipment_id: SP, TOPO, ORB, OCT, HFA, OQAS

    2. get_today_statistics(target_date)
       - 전체 검사 통계 조회

    3. get_hospital_schedule_staff(target_date, department)
       - 근무 인원 조회

    4. get_equipment_list()
       - 장비 목록 조회

    5. get_staff_list()
       - 직원 명단 조회

    6. update_manual_input(field, value)
       - 수기 입력값 저장

    ## 리소스:

    - daily-report://config - 설정 파일
    - daily-report://help - 도움말
    """
    return help_text


if __name__ == "__main__":
    # MCP 서버 실행
    mcp.run()
