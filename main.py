# main.py

import os
import sys
from datetime import datetime

# 모듈 임포트
from modules.sales_analyzer import SalesAnalyzer
from modules.financial_statements import FinancialStatements

def create_directories():
    """필요한 디렉토리 생성"""
    directories = [
        "data/input",
        "data/output", 
        "data/temp",
        "templates"
    ]
    
    for directory in directories:
        os.makedirs(directory, exist_ok=True)
    
    print("📁 디렉토리 구조 생성 완료")

def print_banner():
    """프로그램 시작 배너"""
    print("=" * 60)
    print("🏢 재무 자동화 시스템 (Financial Automation System)")
    print("=" * 60)
    print(f"📅 실행 일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("🔧 기능: 매출분석 + 재무제표 자동생성")
    print("-" * 60)

def show_menu():
    """메뉴 표시"""
    print("\n📋 메뉴를 선택하세요:")
    print("1. 📊 매출/비용 분석기 (Sales Analyzer)")
    print("2. 📋 재무제표 자동 생성 (Financial Statements)")
    print("3. 🚀 전체 실행 (All-in-One)")
    print("4. ⚙️  설정 확인")
    print("5. 🚪 종료")
    print("-" * 40)

def run_sales_analyzer():
    """매출/비용 분석기 실행"""
    print("\n🚀 매출/비용 분석기 시작!")
    print("-" * 40)
    
    try:
        analyzer = SalesAnalyzer()
        
        # 단계별 실행
        print("1️⃣ Excel 파일 수집 중...")
        analyzer.collect_excel_files("data/input/")
        
        print("2️⃣ 트렌드 분석 중...")
        analyzer.analyze_trends()
        
        print("3️⃣ 대시보드 생성 중...")
        dashboard_file = analyzer.generate_dashboard("data/output/sales_analysis_dashboard.html")
        
        print(f"✅ 매출 분석 완료!")
        print(f"📊 결과: {dashboard_file}")
        
        return True
        
    except Exception as e:
        print(f"❌ 매출 분석 실패: {e}")
        return False

def run_financial_statements():
    """재무제표 생성기 실행"""
    print("\n🚀 재무제표 자동 생성기 시작!")
    print("-" * 40)
    
    try:
        fs = FinancialStatements()
        
        print("1️⃣ SAP 시산표 추출 중...")
        fs.load_trial_balance()
        
        print("2️⃣ 재무제표 생성 중...")
        fs.generate_statements()
        
        print("3️⃣ 재무비율 계산 중...")
        fs.calculate_financial_ratios()
        
        print("4️⃣ Excel 파일 생성 중...")
        fs.export_to_excel()
        
        print("✅ 재무제표 생성 완료!")
        return True
        
    except Exception as e:
        print(f"❌ 재무제표 생성 실패: {e}")
        return False

def run_all_in_one():
    """전체 프로세스 실행"""
    print("\n🚀 전체 프로세스 시작!")
    print("=" * 50)
    
    results = {
        'sales_analysis': False,
        'financial_statements': False
    }
    
    # 1. 매출 분석
    print("\n🔹 STEP 1: 매출/비용 분석")
    results['sales_analysis'] = run_sales_analyzer()
    
    # 2. 재무제표 생성
    print("\n🔹 STEP 2: 재무제표 생성")
    results['financial_statements'] = run_financial_statements()
    
    # 결과 요약
    print("\n" + "=" * 50)
    print("📊 전체 실행 결과 요약")
    print("=" * 50)
    
    if results['sales_analysis']:
        print("✅ 매출/비용 분석: 성공")
    else:
        print("❌ 매출/비용 분석: 실패")
    
    if results['financial_statements']:
        print("✅ 재무제표 생성: 성공")
    else:
        print("❌ 재무제표 생성: 실패")
    
    # 성공률 계산
    success_rate = sum(results.values()) / len(results) * 100
    print(f"\n🎯 전체 성공률: {success_rate:.0f}%")
    
    # 출력 파일 안내
    from config import 결산월
    print(f"\n📁 결과 파일 위치:")
    print(f"   - 매출 분석: data/output/sales_analysis_dashboard.html")
    print(f"   - 재무제표: data/output/재무제표_{결산월}.xlsx")

def check_settings():
    """설정 확인"""
    from config import 회사코드, 결산월, 파일저장경로
    
    print("\n⚙️ 현재 설정 확인")
    print("-" * 30)
    print(f"🏢 회사코드: {회사코드}")
    print(f"📅 결산월: {결산월}")
    print(f"💾 저장경로: {파일저장경로}")
    
    # 디렉토리 존재 확인
    required_dirs = ["data/input", "data/output", "data/temp"]
    print(f"\n📁 디렉토리 확인:")
    
    for directory in required_dirs:
        exists = "✅" if os.path.exists(directory) else "❌"
        print(f"   {exists} {directory}")
    
    # SAP 연결 테스트
    print(f"\n🔗 SAP 연결 테스트:")
    try:
        from NEO_SAP import SAPAutomation
        sap = SAPAutomation()
        session_info = sap.get_session_info()
        if session_info:
            print(f"   ✅ 연결 성공 - 사용자: {session_info.get('user', 'Unknown')}")
        else:
            print(f"   ❌ 연결 실패")
    except Exception as e:
        print(f"   ❌ 연결 오류: {e}")

def main():
    """메인 함수"""
    # 초기 설정
    create_directories()
    print_banner()
    
    while True:
        show_menu()
        
        try:
            choice = input("선택 (1-5): ").strip()
            
            if choice == '1':
                run_sales_analyzer()
                
            elif choice == '2':
                run_financial_statements()
                
            elif choice == '3':
                run_all_in_one()
                
            elif choice == '4':
                check_settings()
                
            elif choice == '5':
                print("👋 프로그램을 종료합니다.")
                sys.exit(0)
                
            else:
                print("❌ 잘못된 선택입니다. 1-5 중에서 선택해주세요.")
            
            # 계속 여부 확인
            continue_choice = input("\n계속 실행하시겠습니까? (y/n): ").strip().lower()
            if continue_choice not in ['y', 'yes', '예', 'ㅇ']:
                print("👋 프로그램을 종료합니다.")
                break
                
        except KeyboardInterrupt:
            print("\n\n👋 사용자에 의해 중단되었습니다.")
            break
        except Exception as e:
            print(f"\n❌ 오류 발생: {e}")
            print("다시 시도해주세요.")

if __name__ == "__main__":
    main()