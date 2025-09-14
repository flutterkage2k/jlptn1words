#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
엑셀 파일을 JSON으로 자동 변환하는 프로그램
여러 개의 엑셀 파일을 한 번에 처리할 수 있습니다.
"""

import pandas as pd
import json
import os
import glob
from pathlib import Path

def convert_excel_to_json(excel_file_path, output_folder=None):
    """
    하나의 엑셀 파일을 JSON으로 변환하는 함수
    
    Args:
        excel_file_path (str): 엑셀 파일 경로
        output_folder (str): 출력 폴더 (None이면 원본 파일과 같은 폴더)
    
    Returns:
        bool: 성공하면 True, 실패하면 False
    """
    try:
        # 파일 경로 정보
        file_path = Path(excel_file_path)
        file_name = file_path.stem  # 확장자 제외한 파일명
        
        # 출력 폴더 설정
        if output_folder is None:
            output_folder = file_path.parent
        else:
            output_folder = Path(output_folder)
            output_folder.mkdir(exist_ok=True)
        
        # 엑셀 파일 읽기
        print(f"📖 읽는 중: {file_path.name}")
        
        # 엑셀 파일의 모든 시트 읽기
        excel_data = pd.read_excel(excel_file_path, sheet_name=None)
        
        # 시트가 하나인 경우와 여러 개인 경우 구분
        if len(excel_data) == 1:
            # 시트가 하나인 경우
            sheet_name = list(excel_data.keys())[0]
            df = excel_data[sheet_name]
            
            # 빈 행 제거
            df = df.dropna(how='all')
            
            # JSON으로 변환 (한국어 컬럼명 처리)
            json_data = df.to_dict('records')
            
            # JSON 파일로 저장
            output_file = output_folder / f"{file_name}.json"
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            
            print(f"✅ 완료: {output_file.name} (데이터 {len(json_data)}개)")
            
        else:
            # 시트가 여러 개인 경우
            for sheet_name, df in excel_data.items():
                # 빈 행 제거
                df = df.dropna(how='all')
                
                # JSON으로 변환
                json_data = df.to_dict('records')
                
                # 각 시트별로 별도 JSON 파일 생성
                output_file = output_folder / f"{file_name}_{sheet_name}.json"
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(json_data, f, ensure_ascii=False, indent=2)
                
                print(f"✅ 완료: {output_file.name} (데이터 {len(json_data)}개)")
        
        return True
        
    except Exception as e:
        print(f"❌ 오류 발생 ({excel_file_path}): {str(e)}")
        return False

def batch_convert_excel_to_json(input_folder=".", output_folder=None, file_pattern="*.xlsx"):
    """
    폴더 내의 모든 엑셀 파일을 JSON으로 변환
    
    Args:
        input_folder (str): 엑셀 파일이 있는 폴더
        output_folder (str): JSON 파일을 저장할 폴더
        file_pattern (str): 파일 패턴 (기본값: *.xlsx)
    """
    
    print("🚀 엑셀 → JSON 변환 프로그램 시작!")
    print("=" * 50)
    
    # 입력 폴더에서 엑셀 파일 찾기
    input_path = Path(input_folder)
    excel_files = list(input_path.glob(file_pattern))
    
    # .xls 파일도 함께 찾기
    if file_pattern == "*.xlsx":
        excel_files.extend(list(input_path.glob("*.xls")))
    
    if not excel_files:
        print(f"📂 '{input_folder}' 폴더에서 엑셀 파일을 찾을 수 없습니다.")
        return
    
    print(f"📋 발견된 파일 {len(excel_files)}개:")
    for file in excel_files:
        print(f"   • {file.name}")
    print()
    
    # 변환 시작
    success_count = 0
    fail_count = 0
    
    for excel_file in excel_files:
        if convert_excel_to_json(excel_file, output_folder):
            success_count += 1
        else:
            fail_count += 1
    
    # 결과 요약
    print("\n" + "=" * 50)
    print(f"🎉 변환 완료!")
    print(f"✅ 성공: {success_count}개")
    if fail_count > 0:
        print(f"❌ 실패: {fail_count}개")
    print("=" * 50)

def main():
    """메인 함수 - 사용자 인터페이스"""
    
    print("📊 엑셀 → JSON 변환기")
    print("=" * 30)
    
    # 사용자에게 옵션 제공
    print("사용 방법을 선택하세요:")
    print("1. 현재 폴더의 모든 엑셀 파일 변환")
    print("2. 특정 폴더의 모든 엑셀 파일 변환")
    print("3. 특정 파일 하나만 변환")
    
    try:
        choice = input("\n선택 (1/2/3): ").strip()
        
        if choice == "1":
            # 현재 폴더의 모든 엑셀 파일 변환
            batch_convert_excel_to_json()
            
        elif choice == "2":
            # 특정 폴더 지정
            folder = input("엑셀 파일이 있는 폴더 경로: ").strip()
            if not folder:
                folder = "."
            
            output = input("JSON 파일을 저장할 폴더 (엔터: 원본과 같은 폴더): ").strip()
            if not output:
                output = None
                
            batch_convert_excel_to_json(folder, output)
            
        elif choice == "3":
            # 특정 파일 하나만 변환
            file_path = input("엑셀 파일 경로: ").strip()
            if os.path.exists(file_path):
                convert_excel_to_json(file_path)
            else:
                print("❌ 파일을 찾을 수 없습니다.")
        else:
            print("올바른 번호를 선택해주세요.")
            
    except KeyboardInterrupt:
        print("\n\n프로그램을 종료합니다.")
    except Exception as e:
        print(f"오류가 발생했습니다: {e}")

if __name__ == "__main__":
    main()
