# 엑셀 파일을 JSON 형태로 변환하기

## 🔧 단계별 해결 방법

### 1단계: 현재 폴더에서 가상 환경 만들기
```
1. 폴더로 이동
python3 -m venv venv
```

### 2단계: 가상 환경 활성화
```
source venv/bin/activate
```
성공하면 터미널 앞에 **(venv)** 표시가 나옵니다.

### 3단계: 필요한 라이브러리 설치
```
pip install pandas openpyxl
```

### 4단계: 프로그램 실행

```
python excel_to_json_converter.py
```

---

## 확인하기

성공하면 이런 화면이 출력됩니다:

```
📊 엑셀 → JSON 변환기
==============================
사용 방법을 선택하세요:
1. 현재 폴더의 모든 엑셀 파일 변환
2. 특정 폴더의 모든 엑셀 파일 변환
3. 특정 파일 하나만 변환
선택 (1/2/3):
```

---

## 다음에 사용할 때
다음 번 실행 시에는 아래와 같이 합니다:

```
1. 해당 폴더로 이동해서
2. source venv/bin/activate
3. python excel_to_json_converter.py
```
