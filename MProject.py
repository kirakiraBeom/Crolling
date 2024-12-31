import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import msoffcrypto
from io import BytesIO
import os
import json
from openpyxl.styles import PatternFill
from openpyxl import Workbook
import time

# 키워드 파일 경로
KEYWORD_FILE = "keywords.json"

# 키워드와 해당 상품명 매핑
if os.path.exists(KEYWORD_FILE):
    with open(KEYWORD_FILE, "r", encoding="utf-8") as f:
        keyword_mapping = json.load(f)
else:
    keyword_mapping = {}

def show_keywords():
    def add_keyword():
        keyword = keyword_entry.get().strip()
        new_name = name_entry.get().strip()

        if keyword and new_name:
            keyword_mapping[keyword] = new_name
            save_keywords()
            refresh_list()
        else:
            messagebox.showwarning("입력 필요", "키워드와 상품명을 모두 입력하세요.")

    def delete_keyword():
        selected = listbox.curselection()
        if selected:
            keyword = listbox.get(selected[0]).split(" -> ")[0].strip()
            if keyword in keyword_mapping:
                del keyword_mapping[keyword]
                save_keywords()
                refresh_list()
            else:
                messagebox.showwarning("키워드 없음", "해당 키워드가 존재하지 않습니다.")
        else:
            messagebox.showwarning("선택 필요", "삭제할 키워드를 선택하세요.")

    def refresh_list():
        # 기존 Listbox 내용 삭제
        listbox.delete(0, tk.END)

        # 상단 고정 텍스트를 별도의 Label로 추가
        if not hasattr(refresh_list, "header_label"):
            refresh_list.header_label = tk.Label(
                keyword_window,
                text="키워드 -> 바뀔 상품명",
                bg="#f0f8ff",
                fg="#000000",
                font=("Arial", 12, "bold")
            )
            refresh_list.header_label.pack(pady=(0, 5))

        # Listbox에 나머지 키워드 매핑 삽입
        for k, v in keyword_mapping.items():
            listbox.insert(tk.END, f"{k} -> {v}")

    def show_overview():
        # 키워드 별로 그룹화
        grouped_mapping = {}
        for keyword, product_name in keyword_mapping.items():
            if product_name not in grouped_mapping:
                grouped_mapping[product_name] = []
            grouped_mapping[product_name].append(keyword)

        # 결과 출력
        overview_window = tk.Toplevel()
        overview_window.title("키워드 한눈에 보기")
        overview_window.geometry("400x500")
        overview_window.configure(bg="#f0f8ff")
        overview_window.resizable(False, False)

        text_box = tk.Text(overview_window, wrap="word", bg="#ffffff", fg="#000000", font=("Arial", 10))
        text_box.pack(expand=True, fill="both", padx=10, pady=10)

        for product_name, keywords in grouped_mapping.items():
            text_box.insert(tk.END, f"{product_name}:\n")
            for keyword in keywords:
                text_box.insert(tk.END, f"  - {keyword}\n")
            text_box.insert(tk.END, "\n")

    def search_keyword():
        search_window = tk.Toplevel()
        search_window.title("키워드 검색")
        search_window.geometry("300x100")
        search_window.configure(bg="#f0f8ff")
        search_window.resizable(False, False)

        tk.Label(search_window, text="찾을 키워드:", bg="#f0f8ff", fg="#000000").pack(pady=5)
        search_entry = tk.Entry(search_window)
        search_entry.pack(pady=5)

        # 검색 위치를 기억하기 위한 인덱스
        search_index = {'current': 0}

        def perform_search(event=None):
            search_text = search_entry.get().strip()
            if not search_text:
                messagebox.showwarning("입력 필요", "검색할 키워드를 입력하세요.")
                return

            # 다음 검색을 위해 현재 인덱스를 순환하면서 찾기
            total_items = listbox.size()
            start = search_index['current']
            found = False

            for i in range(total_items):
                idx = (start + i) % total_items  # 리스트 순환
                item_text = listbox.get(idx)
                if search_text.lower() in item_text.lower():
                    listbox.selection_clear(0, tk.END)  # 이전 선택 해제
                    listbox.selection_set(idx)         # 항목 선택
                    listbox.see(idx)                   # 항목으로 이동
                    search_index['current'] = idx + 1  # 다음 검색 시작 위치
                    found = True
                    break

            if not found:
                messagebox.showinfo("검색 결과", f"'{search_text}'에 해당하는 키워드를 찾을 수 없습니다.")
                search_index['current'] = 0  # 검색 인덱스 초기화

        # 검색 버튼
        search_button = tk.Button(search_window, text="검색", command=perform_search, bg="#add8e6", fg="#000000")
        search_button.pack(pady=5)

        # 엔터 키 이벤트 바인딩
        search_entry.bind('<Return>', perform_search)

    keyword_window = tk.Toplevel()
    keyword_window.title("키워드 관리")
    keyword_window.geometry("400x450")
    keyword_window.configure(bg="#f0f8ff")
    keyword_window.resizable(False, False)

    listbox = tk.Listbox(keyword_window, width=50, height=15, bg="#ffffff", fg="#000000")
    listbox.pack(pady=10)

    refresh_list()

    tk.Label(keyword_window, text="키워드:", bg="#f0f8ff", fg="#000000").pack()
    keyword_entry = tk.Entry(keyword_window)
    keyword_entry.pack(pady=5)

    tk.Label(keyword_window, text="상품명:", bg="#f0f8ff", fg="#000000").pack()
    name_entry = tk.Entry(keyword_window)
    name_entry.pack(pady=5)

    button_frame = tk.Frame(keyword_window, bg="#f0f8ff")
    button_frame.pack(pady=10)

    add_button = tk.Button(button_frame, text="추가", command=add_keyword, bg="#add8e6", fg="#000000")
    add_button.grid(row=0, column=0, padx=5)

    delete_button = tk.Button(button_frame, text="삭제", command=delete_keyword, bg="#add8e6", fg="#000000")
    delete_button.grid(row=0, column=1, padx=5)

    overview_button = tk.Button(button_frame, text="한눈에 보기", command=show_overview, bg="#add8e6", fg="#000000")
    overview_button.grid(row=0, column=2, padx=5)

    # Enter 키로 추가 버튼 클릭
    keyword_window.bind('<Return>', lambda event: add_keyword())

    # Ctrl + F 단축키로 검색 창 열기
    keyword_window.bind('<Control-f>', lambda event: search_keyword())

# 키워드 저장 함수
def save_keywords():
    with open(KEYWORD_FILE, "w", encoding="utf-8") as f:
        json.dump(keyword_mapping, f, ensure_ascii=False, indent=4)

# 파일 선택 및 처리 함수
def select_and_process_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if not file_path:
        return

    try:
        # 파일 암호화 확인 및 해제
        password = "0000"  # 비밀번호를 설정합니다
        decrypted_file = decrypt_excel(file_path, password)
        if decrypted_file is None:
            return

        # 엑셀 파일 불러오기
        df = pd.read_excel(decrypted_file, engine='openpyxl')

        def update_product_name(row):
            option_keywords = ["소형", "중형", "대형", "기본형", "고급형", "스몰", "라지", "올인원"]

            # NaN 처리: NaN이면 빈 문자열로 대체
            product_name = str(row['상품명']) if pd.notna(row['상품명']) else ""
            option_info = str(row['옵션정보']) if pd.notna(row['옵션정보']) else ""
            
            combined_text = f"{product_name} {option_info}"

            # 키워드 매핑 적용
            for keyword, mapped_name in keyword_mapping.items():
                if keyword in combined_text:
                    product_name = mapped_name
                    break

            # 옵션정보에서 특정 키워드를 상품명 뒤에 붙이기
            for keyword in option_keywords:
                if keyword in option_info:
                    product_name += f" {keyword}"

            # 상품명이나 옵션정보에 "1+1", "1set+1set"이 있으면 상품명 앞에 추가
            if ("1+1" in combined_text or "1set+1set" in combined_text) and "미니 브러쉬" not in combined_text:
                product_name = f"1+1 {product_name}"

            if "듀얼팩" in combined_text:
                product_name = f"{product_name} 듀얼팩"

            return product_name

            


        if "상품명" in df.columns and "옵션정보" in df.columns:
            df["상품명"] = df.apply(update_product_name, axis=1)
        else:
            messagebox.showerror("열 누락", "필수 열(상품명, 옵션정보)이 파일에 없습니다.")
            return

        # 필요한 열만 선택
        columns_to_keep = ["주문일시", "상품명", "수량", "구매자명"]
        if not all(col in df.columns for col in columns_to_keep):
            missing_cols = [col for col in columns_to_keep if col not in df.columns]
            messagebox.showerror(
                "열 누락",
                f"다음 필수 열이 파일에 없습니다: {', '.join(missing_cols)}",
            )
            return

        df_filtered = df[columns_to_keep]
        
        # 상품명 기준 정렬
        df_filtered = df_filtered.sort_values(by=["상품명"], ascending=True)

        # 처리된 파일 저장 (중복 처리)
        base_name = os.path.splitext(file_path)[0]
        save_path = f"{base_name}_processed.xlsx"
        counter = 1
        while os.path.exists(save_path):
            save_path = f"{base_name}_processed_{counter}.xlsx"
            counter += 1

        # 엑셀 파일 저장 및 열 너비 및 스타일 설정
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            df_filtered.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']
            worksheet.column_dimensions['A'].width = 19
            worksheet.column_dimensions['B'].width = 93
            worksheet.column_dimensions['C'].width = 8
            worksheet.column_dimensions['D'].width = 15

            # 1행 배경색 설정
            fill = PatternFill(start_color="CCCCFF", end_color="CCCCFF", fill_type="solid")
            for cell in worksheet[1]:
                cell.fill = fill

        messagebox.showinfo("성공", f"파일이 처리되어 저장되었습니다: {save_path}")

    except Exception as e:
        messagebox.showerror("오류", f"오류가 발생했습니다: {str(e)}")

# 암호화된 엑셀 파일 해제 함수
def decrypt_excel(file_path, password):
    try:
        decrypted = BytesIO()
        with open(file_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
        return decrypted
    except Exception as e:
        messagebox.showerror("오류", f"암호화된 파일을 해제하는 중 오류가 발생했습니다: {str(e)}")
        return None

def setup_gui():
    root = tk.Tk()
    root.title("엑셀 파일 처리기")
    root.geometry("300x200")  # 창 크기 설정
    root.configure(bg="#f0f8ff")  # 배경색 설정
    root.resizable(False, False)  # 창 크기 고정

    # 키워드 관리 버튼
    keyword_button = tk.Button(
        root,
        text="키워드 보기",
        command=show_keywords,
        bg="#add8e6",
        fg="#000000",
        font=("Arial", 12, "bold")
    )
    keyword_button.pack(pady=20)

    # 파일 처리 섹션
    file_frame = tk.Frame(root, bg="#f0f8ff")
    file_frame.pack(pady=20)

    label = tk.Label(
        file_frame,
        text="처리할 엑셀 파일을 선택하세요:",
        bg="#f0f8ff",
        fg="#000000",
        font=("Arial", 12)
    )
    label.pack(pady=10)

    select_button = tk.Button(
        file_frame,
        text="파일 선택",
        command=select_and_process_file,
        bg="#87cefa",
        fg="#000000",
        font=("Arial", 12, "bold")
    )
    select_button.pack(pady=10)

    root.mainloop()

# 프로그램 실행
if __name__ == "__main__":
    setup_gui()
