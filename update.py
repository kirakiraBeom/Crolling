import os
import shutil
from webdriver_manager.chrome import ChromeDriverManager
import ctypes
import sys

# 드라이버가 저장될 경로
TARGET_DIR = r"C:\Program Files\SeleniumBasic"

def is_admin():
    """현재 프로세스가 관리자 권한인지 확인"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def show_message(title, text):
    """확인 메시지 창 표시"""
    ctypes.windll.user32.MessageBoxW(0, text, title, 0x40)  # 0x40: 정보 아이콘

def update_chrome_driver():
    try:
        # webdriver-manager로 크롬 드라이버 다운로드
        driver_path = ChromeDriverManager().install()
        print(f"다운로드된 드라이버 위치: {driver_path}")

        # 지정된 경로로 복사 준비
        if not os.path.exists(TARGET_DIR):
            os.makedirs(TARGET_DIR)

        target_path = os.path.join(TARGET_DIR, "chromedriver.exe")

        # 기존 파일이 있으면 삭제
        if os.path.exists(target_path):
            os.remove(target_path)
            print(f"기존 드라이버 파일을 삭제했습니다: {target_path}")

        # 새 드라이버 복사
        shutil.copy(driver_path, target_path)
        print(f"크롬 드라이버가 성공적으로 업데이트되었습니다! 위치: {target_path}")

        # 확인 메시지 표시
        show_message("업데이트 완료", f"크롬 드라이버가 성공적으로 업데이트되었습니다.\n위치: {target_path}")

    except Exception as e:
        error_message = f"드라이버 업데이트 중 오류가 발생했습니다:\n{e}"
        print(error_message)
        # 오류 메시지 표시
        show_message("업데이트 오류", error_message)

if __name__ == "__main__":
    if not is_admin():
        # 관리자 권한으로 스크립트를 재실행
        print("관리자 권한이 필요합니다. 프로그램을 다시 실행합니다...")
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, __file__, None, 1
        )
        sys.exit()

    # 드라이버 업데이트 실행
    update_chrome_driver()
