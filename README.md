# 재무 자동화 시스템 (Financial Automation System)

> SAP GUI 자동화를 활용한 재무/회계 업무 자동화 도구

## 📋 프로젝트 개요

재무팀의 반복적인 업무를 자동화하여 효율성을 극대화하는 시스템입니다.
- **매출/비용 분석기**: Excel 데이터 자동 수집 및 트렌드 분석
- **재무제표 자동 생성**: 시산표 기반 재무제표 자동 변환 및 전년 비교

## 🏗️ 프로젝트 구조

```
financial-automation/
├── README.md                 # 프로젝트 문서
├── requirements.txt          # 의존성 패키지
├── config.py                # 설정 파일
├── NEO_SAP.py               # SAP GUI 자동화 모듈
├── modules/
│   ├── __init__.py
│   ├── sales_analyzer.py    # 매출/비용 분석기
│   ├── financial_statements.py  # 재무제표 생성기
│   └── data_processor.py    # 데이터 처리 유틸리티
├── templates/
│   ├── dashboard.html       # 리포트 템플릿
│   └── financial_report.html
├── data/
│   ├── input/              # 입력 Excel 파일들
│   ├── output/             # 생성된 리포트들
│   └── temp/               # 임시 파일들
├── tests/
│   ├── test_sales_analyzer.py
│   └── test_financial_statements.py
└── main.py                 # 메인 실행 파일
```

## 🚀 주요 기능

### A. 매출/비용 분석기 (`sales_analyzer.py`)
- **Excel 파일 자동 수집**: 지정 폴더의 모든 Excel 파일 통합
- **월별/분기별 트렌드 분석**: 시계열 데이터 분석 및 시각화
- **예산 대비 실적 비교**: 실적과 예산 간 차이 분석
- **자동 리포트 생성**: HTML 대시보드 형태로 결과 출력

### B. 재무제표 자동 생성 (`financial_statements.py`)
- **시산표 → 재무제표 변환**: SAP 시산표를 표준 재무제표 형식으로 변환
- **전년 동기 비교**: 당기 vs 전기 비교 분석
- **재무비율 자동 계산**: 유동비율, ROE, ROA 등 주요 지표 산출
- **컴플라이언스 체크**: 재무제표 검증 및 오류 탐지

## 🛠️ 설치 및 실행

### 간단 실행 (권장)
```bash
# 1. install.bat 실행 (자동 설치)
install.bat

# 2. run.bat 실행 (프로그램 시작)
run.bat
```

### 수동 설치
```bash
# 가상환경 생성
python -m venv venv
venv\Scripts\activate

# 의존성 설치
pip install -r requirements.txt

# 프로그램 실행
python main.py
```

## ⚙️ 설정

`config.py` 파일에서 설정 변경:
```python
회사코드 = "h182"          # SAP 회사코드
결산월 = "2025.05"         # 분석할 월
파일저장경로 = "C:\\"      # 결과 저장 경로
```

## 🎯 사용법

1. **Excel 파일 준비**: `data/input/` 폴더에 매출/손익 Excel 파일 저장
2. **프로그램 실행**: `run.bat` 또는 `python main.py`
3. **메뉴 선택**:
   - `1`: 매출/비용 분석만 실행
   - `2`: 재무제표 생성만 실행  
   - `3`: 전체 실행 (추천)
4. **결과 확인**: `data/output/` 폴더에서 생성된 파일 확인

## 📊 출력 결과

- **HTML 대시보드**: `sales_analysis_dashboard.html`
- **재무제표**: `재무제표_2025.05.xlsx`
- **시산표**: `시산표_2025.05.xlsx`

## 🔧 추가 도구

### VBS → Python 변환기 (`converter.py`)
기존 SAP VBS 스크립트를 Python으로 자동 변환

### 웹 실행기 (`wrapper.py`)
브라우저에서 스크립트를 실행할 수 있는 웹 인터페이스

## 📖 문서

- `사용방법.md`: 상세한 사용 가이드
- `사용팁.md`: 효율적인 활용 팁
- `체크리스트.md`: 설치 및 실행 체크리스트

## 🚨 문제 해결

### 자주 발생하는 오류
1. **SAP 연결 실패**: SAP GUI 실행 및 로그인 확인
2. **파일을 찾을 수 없음**: `data/input/` 폴더에 Excel 파일 배치
3. **권한 오류**: 관리자 권한으로 실행

자세한 해결 방법은 `사용방법.md` 참조

## 📄 라이선스

MIT License - 자유롭게 사용, 수정, 배포 가능

---

💡 **재무팀 업무 효율성을 혁신적으로 개선하는 자동화 솔루션**