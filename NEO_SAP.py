import win32com.client
import time
import sys
from config import 회사코드, 결산월, 파일저장경로

class SAPAutomation:
    def __init__(self, connection_index=0, session_index=0):
        self.session = None
        self.connection_index = connection_index
        self.session_index = session_index
        self.회사코드 = 회사코드
        self.결산월 = 결산월  
        self.파일저장경로 = 파일저장경로
        self.connect_to_sap()
    
    def connect_to_sap(self):
        """SAP GUI에 연결 (다중 세션 지원)"""
        try:
            # SAP GUI 연결
            sapgui = win32com.client.GetObject("SAPGUI")
            application = sapgui.GetScriptingEngine
            connection = application.Children(self.connection_index)
            self.session = connection.Children(self.session_index)
            print(f"SAP 연결 성공! (Connection: {self.connection_index}, Session: {self.session_index})")
        except Exception as e:
            print(f"SAP 연결 실패: {e}")
            print(f"Connection Index: {self.connection_index}, Session Index: {self.session_index}")
            sys.exit(1)
    
    @classmethod
    def create_multiple_sessions(cls, count=3):
        """여러 세션 인스턴스 생성"""
        sessions = []
        for i in range(count):
            try:
                sap_session = cls(connection_index=0, session_index=i)
                sessions.append(sap_session)
                print(f"세션 {i} 생성 완료!")
            except Exception as e:
                print(f"세션 {i} 생성 실패: {e}")
                break
        return sessions
    
    def new_session(self):
        """새로운 세션 생성"""
        try:
            self.session.createSession()
            print("새 세션 생성 완료!")
        except Exception as e:
            print(f"새 세션 생성 실패: {e}")
    
    def get_session_info(self):
        """현재 세션 정보 확인"""
        try:
            session_info = {
                'connection_index': self.connection_index,
                'session_index': self.session_index,
                'session_id': self.session.id,
                'user': self.session.info.user,
                'client': self.session.info.client,
                'language': self.session.info.language
            }
            return session_info
        except Exception as e:
            print(f"세션 정보 확인 실패: {e}")
            return None
    
    def enter_the_wutang(self, tcode):
        """T-code로 화면 이동 (Enter the Wu-Tang!)"""
        try:
            self.find("wnd[0]/tbar[0]/okcd").text = f"/n{tcode}"
            self.find("wnd[0]").sendVKey(0)  # Enter
            time.sleep(1)
            print(f"Wu-Tang {tcode} ain't nuthing ta f*** wit! 🔥")
        except Exception as e:
            print(f"Wu-Tang clan entry failed: {e}")
    
    def find(self, id):
        """SAP 객체 찾기 헬퍼 함수"""
        return self.session.findById(id)
    
    def 실행(self):
        """F8 실행"""
        try:
            self.session.sendVKey(8)
            time.sleep(1)
            print("실행 완료! 🚀")
        except Exception as e:
            print(f"실행 실패: {e}")
    
    def 저장(self):
        """Ctrl+S 저장"""
        try:
            self.session.sendVKey(11)
            time.sleep(1)
            print("저장 완료! 💾")
        except Exception as e:
            print(f"저장 실패: {e}")
    
    def 뒤로가기(self):
        """F3 뒤로가기"""
        try:
            self.session.sendVKey(3)
            time.sleep(0.5)
            print("뒤로가기 완료! ⬅️")
        except Exception as e:
            print(f"뒤로가기 실패: {e}")
    
    def 엔터(self):
        """Enter 키"""
        try:
            self.session.sendVKey(0)
            time.sleep(0.5)
            print("엔터 완료! ⏎")
        except Exception as e:
            print(f"엔터 실패: {e}")
    
    def 새로고침(self):
        """F5 새로고침"""
        try:
            self.session.sendVKey(5)
            time.sleep(1)
            print("새로고침 완료! 🔄")
        except Exception as e:
            print(f"새로고침 실패: {e}")
    
    def export_to_excel(self):
        """조회 결과를 Excel로 내보내기"""
        try:
            # Ctrl+Shift+F9 (Excel 내보내기)
            self.session.findById("wnd[0]").sendVKey(9, "ctrl+shift")
            time.sleep(1)
            
            # 파일 경로 지정 (config에서 가져옴)
            file_path = f"{self.파일저장경로}sap_export_{self.결산월}.xlsx"
            self.find("wnd[1]/usr/ctrlSSLN_EXPORT/txtDY_PATH").text = file_path
            self.find("wnd[1]/tbar[0]/btn[11]").press()  # 확인
            
            print(f"Excel 내보내기 완료: {file_path}")
            
        except Exception as e:
            print(f"Excel 내보내기 실패: {e}")

# 사용 예시
if __name__ == "__main__":

    # 일반적인 경우
    sap = SAPAutomation()
    session = sap.session

    # 두 개 세션 생성
    sap1 = SAPAutomation(connection_index=0, session_index=0)  # 첫 번째 세션
    sap2 = SAPAutomation(connection_index=0, session_index=1)  # 두 번째 세션
    
    # 각 세션에서 다른 작업 실행
    sap1.enter_the_wutang("FB01")      # 첫 번째 세션: 전표입력
    sap2.enter_the_wutang("zcoqr451d") # 두 번째 세션: 관리손익계산서 조회
    
    print("멀티 세션 작업 완료!")