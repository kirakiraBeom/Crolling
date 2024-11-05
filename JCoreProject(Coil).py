import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import datetime

def select_first_file():
    global first_file_path
    first_file_path = filedialog.askopenfilename(title="첫 번째 파일 선택", filetypes=[("Excel files", "*.xlsx *.xls")])
    print(f"첫 번째 파일 경로: {first_file_path}")

def select_second_file():
    second_file_path = filedialog.askopenfilename(title="두 번째 파일 선택", filetypes=[("Excel files", "*.xlsx *.xls")])
    print(f"두 번째 파일 경로: {second_file_path}")
    if first_file_path and second_file_path:
        process_files(first_file_path, second_file_path)

def process_files(first_file, second_file):
    # 엑셀 파일을 로드합니다
    df1 = pd.read_excel(first_file, sheet_name="20241104")
    df2 = pd.read_excel(second_file)

    # 첫 번째 파일에서 필요한 열만 선택하여 두 번째 파일에 매핑합니다
    df2['성명'] = df1['받는분성함']
    df2['차종/연식/색상'] = df1['옵션']
    df2['주소'] = df1['받는분주소']
    df2['연락처'] = df1['반는분연락처']
    df2['기타연락처'] = df1['받는분기타연락처']
    df2['배송메세지'] = df1['배송메시지']

    # 현재 날짜로 수정된 파일을 바탕화면에 저장합니다
    current_date = datetime.datetime.now().strftime("%Y%m%d")
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")  # 바탕화면 경로
    output_filename = os.path.join(desktop_path, f"코일매트_{current_date}.xlsx")
    
    df2.to_excel(output_filename, index=False)  # 엑셀 파일로 저장
    print(f"파일이 저장되었습니다: {output_filename}")

# Tkinter GUI 설정
root = tk.Tk()
root.title("J CORE PROJECT")  # 창 제목
root.geometry("400x250")  # 창 크기
root.configure(bg="#2c3e50")  # 배경 색상 설정

label = tk.Label(root, text="엑셀 파일을 선택하세요:", font=("Arial", 16), bg="#2c3e50", fg="white")
label.pack(pady=20)  # 레이블 배치

first_button = tk.Button(root, text="첫 번째 파일 선택", font=("Arial", 14), bg="#3498db", fg="white", command=select_first_file)
first_button.pack(pady=10)  # 첫 번째 버튼 배치

second_button = tk.Button(root, text="두 번째 파일 선택", font=("Arial", 14), bg="#e74c3c", fg="white", command=select_second_file)
second_button.pack(pady=10)  # 두 번째 버튼 배치

root.mainloop()  # GUI 메인 루프 실행
