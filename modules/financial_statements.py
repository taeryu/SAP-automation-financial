# modules/financial_statements.py

import pandas as pd
import numpy as np
import os
import time
from datetime import datetime, timedelta
import json
import sys
sys.path.append('..')
from NEO_SAP import SAPAutomation
from config import 회사코드, 결산월, 파일저장경로

class FinancialStatements:
    def __init__(self):
        self.sap = SAPAutomation()
        self.trial_balance = None
        self.previous_year_data = None
        self.statements = {}
        self.ratios = {}
        
    def load_trial_balance(self, file_path=None):
        """SAP 시산표 데이터 로드"""
        if file_path is None:
            # SAP에서 직접 시산표 추출
            file_path = self.extract_trial_balance_from_sap()
        
        try:
            self.trial_balance = pd.read_excel(file_path)
            print(f"✅ 시산표 로드 완료: {file_path}")
            
            # 데이터 정제
            self.trial_balance = self.clean_trial_balance(self.trial_balance)
            
        except Exception as e:
            print(f"❌ 시산표 로드 실패: {e}")
    
    def extract_trial_balance_from_sap(self):
        """SAP에서 시산표 직접 추출"""
        print("📊 SAP에서 시산표 추출 중...")
        
        try:
            # F.01 (시산표) T-code 실행
            self.sap.enter_the_wutang("F.01")
            
            # 조회 조건 설정
            year, month = 결산월.split(".")
            self.sap.find("wnd[0]/usr/ctrlCOMPANY_CODE/txtS_BUKRS-LOW").text = 회사코드
            self.sap.find("wnd[0]/usr/ctrlFISCAL_YEAR/txtS_GJAHR-LOW").text = year
            self.sap.find("wnd[0]/usr/ctrlPERIOD/txtS_MONAT-LOW").text = month
            
            # 실행 및 Excel 내보내기
            self.sap.실행()
            time.sleep(3)
            self.sap.export_to_excel()
            
            # 파일 경로 반환
            file_path = f"{파일저장경로}시산표_{결산월}.xlsx"
            return file_path
            
        except Exception as e:
            print(f"❌ SAP 시산표 추출 실패: {e}")
            return None
    
    def clean_trial_balance(self, df):
        """시산표 데이터 정제"""
        # 빈 행 제거
        df = df.dropna(how='all')
        
        # 계정과목, 차변, 대변 컬럼 식별
        df.columns = df.columns.astype(str)
        
        # 계정과목 컬럼 찾기
        account_cols = [col for col in df.columns if any(keyword in col.lower() for keyword in ['계정', 'account', '과목'])]
        debit_cols = [col for col in df.columns if any(keyword in col.lower() for keyword in ['차변', 'debit', 'dr'])]
        credit_cols = [col for col in df.columns if any(keyword in col.lower() for keyword in ['대변', 'credit', 'cr'])]
        
        # 표준 컬럼명으로 변경
        if account_cols:
            df = df.rename(columns={account_cols[0]: '계정과목'})
        if debit_cols:
            df = df.rename(columns={debit_cols[0]: '차변'})
        if credit_cols:
            df = df.rename(columns={credit_cols[0]: '대변'})
        
        # 숫자 컬럼 변환
        for col in ['차변', '대변']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        return df
    
    def generate_statements(self):
        """재무제표 자동 생성"""
        if self.trial_balance is None:
            print("❌ 시산표 데이터가 없습니다.")
            return
        
        print("📋 재무제표 생성 중...")
        
        # 1. 재무상태표 생성
        self.statements['재무상태표'] = self.create_balance_sheet()
        
        # 2. 손익계산서 생성
        self.statements['손익계산서'] = self.create_income_statement()
        
        print("✅ 재무제표 생성 완료")
        return self.statements
    
    def create_balance_sheet(self):
        """재무상태표 생성"""
        balance_sheet = {}
        
        # 계정과목별 잔액 계산
        df = self.trial_balance.copy()
        df['잔액'] = df['차변'] - df['대변']
        
        # 자산 항목
        balance_sheet['자산'] = {
            '유동자산': {
                '현금및현금성자산': self.get_account_balance(df, ['현금', '보통예금', '당좌예금']),
                '매출채권': self.get_account_balance(df, ['매출채권', '받을어음']),
                '재고자산': self.get_account_balance(df, ['재고자산', '상품', '제품', '원재료']),
            },
            '비유동자산': {
                '유형자산': self.get_account_balance(df, ['토지', '건물', '기계장치', '차량운반구', '비품']),
                '무형자산': self.get_account_balance(df, ['영업권', '특허권', '소프트웨어']),
            }
        }
        
        # 부채 항목
        balance_sheet['부채'] = {
            '유동부채': {
                '매입채무': self.get_account_balance(df, ['매입채무', '지급어음']),
                '단기차입금': self.get_account_balance(df, ['단기차입금', '운전자금대출']),
                '미지급금': self.get_account_balance(df, ['미지급금', '미지급비용']),
            },
            '비유동부채': {
                '장기차입금': self.get_account_balance(df, ['장기차입금', '사채']),
            }
        }
        
        # 자본 항목
        balance_sheet['자본'] = {
            '자본금': self.get_account_balance(df, ['자본금', '출자금']),
            '이익잉여금': self.get_account_balance(df, ['이익잉여금', '미처분이익잉여금']),
            '당기순이익': self.get_account_balance(df, ['당기순이익'])
        }
        
        return balance_sheet
    
    def create_income_statement(self):
        """손익계산서 생성"""
        income_statement = {}
        
        df = self.trial_balance.copy()
        df['잔액'] = df['차변'] - df['대변']
        
        # 수익 항목 (대변 잔액)
        income_statement['수익'] = {
            '매출액': abs(self.get_account_balance(df, ['매출', '상품매출', '제품매출'])),
            '기타수익': abs(self.get_account_balance(df, ['잡수익', '이자수익', '임대수익']))
        }
        
        # 비용 항목 (차변 잔액)
        income_statement['비용'] = {
            '매출원가': self.get_account_balance(df, ['매출원가', '상품매출원가']),
            '판매비와관리비': {
                '급여': self.get_account_balance(df, ['급여', '임금']),
                '임차료': self.get_account_balance(df, ['임차료', '지급임차료']),
                '감가상각비': self.get_account_balance(df, ['감가상각비']),
                '기타판관비': self.get_account_balance(df, ['광고선전비', '접대비', '통신비'])
            },
            '금융비용': self.get_account_balance(df, ['이자비용', '차입금이자'])
        }
        
        # 손익 계산
        income_statement['매출총이익'] = income_statement['수익']['매출액'] - income_statement['비용']['매출원가']
        
        total_selling_admin = sum(income_statement['비용']['판매비와관리비'].values())
        income_statement['영업이익'] = income_statement['매출총이익'] - total_selling_admin
        
        income_statement['법인세비용차감전순이익'] = (income_statement['영업이익'] + 
                                              income_statement['수익']['기타수익'] - 
                                              income_statement['비용']['금융비용'])
        
        income_statement['법인세비용'] = self.get_account_balance(df, ['법인세비용'])
        income_statement['당기순이익'] = income_statement['법인세비용차감전순이익'] - income_statement['법인세비용']
        
        return income_statement
    
    def get_account_balance(self, df, keywords):
        """계정과목 키워드로 잔액 합계 계산"""
        total = 0
        
        for keyword in keywords:
            mask = df['계정과목'].str.contains(keyword, na=False, case=False)
            matched_rows = df[mask]
            if not matched_rows.empty:
                total += matched_rows['잔액'].sum()
        
        return total
    
    def calculate_financial_ratios(self):
        """재무비율 계산"""
        print("📈 재무비율 계산 중...")
        
        bs = self.statements.get('재무상태표', {})
        is_data = self.statements.get('손익계산서', {})
        
        # 유동성 비율
        current_assets = bs.get('자산', {}).get('유동자산', {})
        current_liabilities = bs.get('부채', {}).get('유동부채', {})
        
        ca_total = sum(current_assets.values()) if current_assets else 0
        cl_total = sum(current_liabilities.values()) if current_liabilities else 0
        
        self.ratios['유동비율'] = (ca_total / cl_total * 100) if cl_total > 0 else 0
        
        # 안정성 비율
        total_debt = sum(bs.get('부채', {}).get('유동부채', {}).values()) + sum(bs.get('부채', {}).get('비유동부채', {}).values())
        total_equity = sum(bs.get('자본', {}).values())
        
        self.ratios['부채비율'] = (total_debt / total_equity * 100) if total_equity > 0 else 0
        
        # 수익성 비율
        sales = is_data.get('수익', {}).get('매출액', 0)
        operating_income = is_data.get('영업이익', 0)
        net_income = is_data.get('당기순이익', 0)
        
        self.ratios['영업이익률'] = (operating_income / sales * 100) if sales > 0 else 0
        self.ratios['순이익률'] = (net_income / sales * 100) if sales > 0 else 0
        self.ratios['ROE'] = (net_income / total_equity * 100) if total_equity > 0 else 0
        
        print("✅ 재무비율 계산 완료")
        return self.ratios
    
    def export_to_excel(self, output_file=None):
        """Excel 파일로 내보내기"""
        if output_file is None:
            output_file = f"{파일저장경로}재무제표_{결산월}.xlsx"
        
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # 재무상태표
                if '재무상태표' in self.statements:
                    bs_df = self.convert_to_dataframe(self.statements['재무상태표'], '재무상태표')
                    bs_df.to_excel(writer, sheet_name='재무상태표', index=False)
                
                # 손익계산서
                if '손익계산서' in self.statements:
                    is_df = self.convert_to_dataframe(self.statements['손익계산서'], '손익계산서')
                    is_df.to_excel(writer, sheet_name='손익계산서', index=False)
                
                # 재무비율
                if self.ratios:
                    ratio_df = pd.DataFrame(list(self.ratios.items()), columns=['비율명', '값'])
                    ratio_df.to_excel(writer, sheet_name='재무비율', index=False)
            
            print(f"✅ Excel 내보내기 완료: {output_file}")
            
        except Exception as e:
            print(f"❌ Excel 내보내기 실패: {e}")
    
    def convert_to_dataframe(self, data, statement_type):
        """재무제표 데이터를 DataFrame으로 변환"""
        rows = []
        
        def flatten_dict(d, parent_key=''):
            for k, v in d.items():
                new_key = f"{parent_key}.{k}" if parent_key else k
                if isinstance(v, dict):
                    flatten_dict(v, new_key)
                else:
                    rows.append({'항목': new_key, '금액': v})
        
        flatten_dict(data)
        return pd.DataFrame(rows)

# CLI 실행 지원
if __name__ == "__main__":
    fs = FinancialStatements()
    
    print("🚀 재무제표 자동 생성기 실행!")
    
    # 1. 시산표 로드
    fs.load_trial_balance()
    
    # 2. 재무제표 생성
    fs.generate_statements()
    
    # 3. 재무비율 계산
    fs.calculate_financial_ratios()
    
    # 4. Excel 내보내기
    fs.export_to_excel()
    
    print("✅ 재무제표 생성 완료!")