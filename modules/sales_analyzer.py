# modules/sales_analyzer.py

import pandas as pd
import numpy as np
import os
import glob
from datetime import datetime, timedelta
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import json
import sys
sys.path.append('..')
from NEO_SAP import SAPAutomation
from config import 회사코드, 결산월, 파일저장경로

class SalesAnalyzer:
    def __init__(self):
        self.sap = SAPAutomation()
        self.data = {}
        self.budget_data = None
        self.analysis_results = {}
        
    def collect_excel_files(self, input_folder="../data/input/"):
        """지정 폴더의 모든 Excel 파일 자동 수집 및 통합"""
        print(f"📁 {input_folder}에서 Excel 파일 수집 중...")
        
        # Excel 파일 패턴 검색
        excel_patterns = [
            "*.xlsx", "*.xls", "*매출*.xlsx", "*손익*.xlsx", "*비용*.xlsx"
        ]
        
        all_files = []
        for pattern in excel_patterns:
            files = glob.glob(os.path.join(input_folder, pattern))
            all_files.extend(files)
        
        # 중복 제거
        all_files = list(set(all_files))
        
        print(f"📊 발견된 파일: {len(all_files)}개")
        
        for file_path in all_files:
            try:
                # 파일명에서 월별 정보 추출
                filename = os.path.basename(file_path)
                month_key = self.extract_month_from_filename(filename)
                
                # Excel 파일 읽기 (여러 시트 지원)
                xl_file = pd.ExcelFile(file_path)
                
                for sheet_name in xl_file.sheet_names:
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    
                    # 데이터 정제
                    df = self.clean_data(df)
                    
                    # 데이터 저장
                    key = f"{month_key}_{sheet_name}"
                    self.data[key] = df
                    
                print(f"✅ {filename} 처리 완료")
                
            except Exception as e:
                print(f"❌ {filename} 처리 실패: {e}")
        
        return self.data
    
    def extract_month_from_filename(self, filename):
        """파일명에서 년월 정보 추출"""
        import re
        
        # 패턴: 2025.05, 202505, 2025_05 등
        patterns = [
            r'(\d{4})[.\-_](\d{2})',
            r'(\d{4})(\d{2})',
            r'(\d{2})[.\-_](\d{4})',  # 05.2025 형태
        ]
        
        for pattern in patterns:
            match = re.search(pattern, filename)
            if match:
                year, month = match.groups()
                if len(year) == 2:  # 05.2025 → 2025.05로 변환
                    year, month = month, year
                return f"{year}.{month.zfill(2)}"
        
        # 패턴이 없으면 현재 월 사용
        return 결산월
    
    def clean_data(self, df):
        """데이터 정제 및 표준화"""
        # 빈 행/열 제거
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # 숫자 컬럼 변환
        for col in df.columns:
            if df[col].dtype == 'object':
                # 숫자로 변환 가능한 컬럼 찾기
                df[col] = pd.to_numeric(df[col], errors='ignore')
        
        return df
    
    def analyze_trends(self):
        """월별/분기별 트렌드 분석"""
        print("📈 트렌드 분석 실행 중...")
        
        # 월별 매출 트렌드
        monthly_sales = self.calculate_monthly_sales()
        
        # 분기별 트렌드
        quarterly_trends = self.calculate_quarterly_trends(monthly_sales)
        
        # 성장률 계산
        growth_rates = self.calculate_growth_rates(monthly_sales)
        
        # 계절성 분석
        seasonality = self.analyze_seasonality(monthly_sales)
        
        self.analysis_results.update({
            'monthly_sales': monthly_sales,
            'quarterly_trends': quarterly_trends,
            'growth_rates': growth_rates,
            'seasonality': seasonality
        })
        
        print("✅ 트렌드 분석 완료")
        return self.analysis_results
    
    def calculate_monthly_sales(self):
        """월별 매출 계산"""
        monthly_data = {}
        
        for key, df in self.data.items():
            month = key.split('_')[0]
            
            # '매출' 키워드가 포함된 컬럼 찾기
            sales_columns = [col for col in df.columns if '매출' in str(col) or 'SALES' in str(col).upper()]
            
            if sales_columns:
                total_sales = df[sales_columns].sum().sum()
                monthly_data[month] = total_sales
        
        return monthly_data
    
    def generate_dashboard(self, output_file="../data/output/sales_analysis_dashboard.html"):
        """인터랙티브 HTML 대시보드 생성"""
        print("🎨 대시보드 생성 중...")
        
        # 차트 데이터 준비
        chart_data = {
            'monthly_sales': self.analysis_results.get('monthly_sales', {}),
            'growth_rates': self.analysis_results.get('growth_rates', {}),
        }
        
        # HTML 템플릿
        html_content = f'''
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{결산월} 매출 분석 대시보드</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }}
        .container {{ max-width: 1200px; margin: 0 auto; }}
        .header {{ background: #2c3e50; color: white; padding: 20px; border-radius: 10px; }}
        .summary-card {{ background: white; padding: 20px; margin: 20px 0; border-radius: 10px; }}
        .chart-container {{ background: white; padding: 20px; margin: 20px 0; border-radius: 10px; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>{결산월} 매출 분석 대시보드</h1>
            <p>자동 생성된 매출 분석 리포트</p>
        </div>
        
        <div class="summary-card">
            <h3>분석 요약</h3>
            <p><strong>회사코드:</strong> {회사코드}</p>
            <p><strong>분석 기간:</strong> {결산월}</p>
            <p><strong>데이터 수집:</strong> {len(self.data)}개 파일</p>
        </div>
        
        <div class="chart-container">
            <h3>매출 트렌드</h3>
            <div id="trendChart"></div>
        </div>
    </div>
    
    <script>
        const salesData = {json.dumps(chart_data['monthly_sales'], ensure_ascii=False)};
        
        const months = Object.keys(salesData);
        const sales = Object.values(salesData);
        
        const trace = {{
            x: months,
            y: sales,
            type: 'scatter',
            mode: 'lines+markers',
            name: '월별 매출'
        }};
        
        const layout = {{
            title: '월별 매출 트렌드',
            xaxis: {{ title: '월' }},
            yaxis: {{ title: '매출액 (원)' }}
        }};
        
        Plotly.newPlot('trendChart', [trace], layout);
    </script>
</body>
</html>
        '''
        
        # 파일 저장
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"✅ 대시보드 생성 완료: {output_file}")
        return output_file

# CLI 실행 지원
if __name__ == "__main__":
    analyzer = SalesAnalyzer()
    
    print("🚀 매출/비용 분석기 실행!")
    
    # 1. Excel 파일 수집
    analyzer.collect_excel_files()
    
    # 2. 트렌드 분석
    analyzer.analyze_trends()
    
    # 3. 대시보드 생성
    analyzer.generate_dashboard()
    
    print("✅ 분석 완료!")