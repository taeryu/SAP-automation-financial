# modules/__init__.py

"""
재무 자동화 시스템 모듈

이 패키지는 SAP GUI 자동화를 통한 재무/회계 업무 자동화를 위한 모듈들을 포함합니다.

모듈:
- sales_analyzer: 매출/비용 분석 및 트렌드 분석
- financial_statements: 재무제표 자동 생성 및 전년 비교
"""

from .sales_analyzer import SalesAnalyzer
from .financial_statements import FinancialStatements

__version__ = "1.0.0"
__author__ = "Financial Automation Team"

__all__ = [
    'SalesAnalyzer',
    'FinancialStatements'
]