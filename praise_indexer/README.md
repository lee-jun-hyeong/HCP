# 찬양 검색 및 PPT 생성 프로그램

찬양 PPTX 파일을 검색하고, 선택한 찬양들을 하나의 PPT로 합쳐주는 프로그램입니다.

## 주요 기능

- 🔍 **찬양 검색**: 제목, 가사, 전체 검색
- 📝 **JSON 기반 인덱싱**: 빠른 검색을 위한 JSON 데이터베이스
- 🎨 **PPT 생성**: 선택한 찬양들을 하나의 PPT로 자동 생성
- 🎯 **템플릿 지원**: temp.pptx의 스타일을 그대로 적용
- 🗑️ **파일 관리**: PPTX 파일 추가/삭제 기능

## 설치 방법

### 1. Python 설치
Python 3.8 이상 필요

### 2. 의존성 설치
```bash
pip install -r requirements.txt
```

필요한 패키지:
- customtkinter
- python-pptx

### 3. 프로그램 실행
```bash
python json_gui.py
```

## 사용 방법

### 1. 인덱싱
- "인덱싱" 버튼 클릭
- Praise_PPT 폴더의 모든 PPTX 파일을 JSON으로 인덱싱

### 2. 검색
- 검색어 입력
- 검색 타입 선택 (제목/가사/전체)
- 검색 결과에서 "선택" 버튼 클릭

### 3. PPT 생성
- 선택된 찬양 목록에서 순서 조정 (드래그 앤 드롭)
- "PPT 생성" 버튼 클릭
- 저장 위치 선택

## 파일 구조

```
praise_indexer/
├── json_gui.py              # 메인 GUI 프로그램
├── json_indexer.py          # JSON 인덱싱 엔진
├── json_ppt_generator_fixed.py  # PPT 생성기
├── config.json              # 설정 파일
├── praise_index.json        # 인덱스 데이터 (자동 생성)
├── temp.pptx               # PPT 템플릿
└── Praise_PPT/             # 찬양 PPTX 파일들
```

## 기능 상세

### 텍스트 정규화
- 특수 제어 문자 자동 정리 (`_x000B_` 등)
- 반복 가사 보존
- 줄바꿈 정규화

### 파일 관리
- 새 PPTX 파일 추가
- 파일 삭제 (휴지통 버튼)
- 선택 목록 유지

### PPT 생성
- 템플릿 스타일 자동 적용
- 슬라이드별 분할
- 구분 슬라이드 자동 추가

## 빌드 방법

PyInstaller로 실행 파일 생성:
```bash
pyinstaller --onefile --windowed --add-data "praise_index.json;." --add-data "temp.pptx;." json_gui.py
```

## 라이센스

MIT License

## 작성자

lee-jun-hyeong
