# 엑셀 파일로 정리된 일본어를 json 형태로 가공하기 입니다. 

## 🔧 단계별 해결 방법

1단계: 현재 폴더에서 가상 환경 만들기
폴더로 이동
python3 -m venv venv

2단계: 가상 환경 활성화
source venv/bin/activate
성공하면 터미널 앞에 (venv) 가 표시됩니다!

3단계: 필요한 라이브러리 설치
pip install pandas openpyxl

4단계: 프로그램 실행
python excel_to_json_converter.py


## 확인하기

성공하면 이런 화면이 나올 예요:
📊 엑셀 → JSON 변환기
==============================
사용 방법을 선택하세요:
1. 현재 폴더의 모든 엑셀 파일 변환
2. 특정 폴더의 모든 엑셀 파일 변환
3. 특정 파일 하나만 변환

선택 (1/2/3):


## 다음에 사용할 때
나중에 다시 사용하려면:  
```
해당폴더로 이동해서  
source venv/bin/activate  
python excel_to_json_converter.py  
```
