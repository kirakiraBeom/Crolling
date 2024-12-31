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
KEYWORD_FILE = "C:\\Temp\\JSON\\JSON\\keywords.json"

# 키워드와 해당 상품명 매핑
if os.path.exists(KEYWORD_FILE):
    with open(KEYWORD_FILE, "r", encoding="utf-8") as f:
        keyword_mapping = json.load(f)
else:
    keyword_mapping = {}  

def show_keywords():
    def add_keyword():
        global keyword_mapping  # 전역 변수 선언
        
        # 현재 선택된 항목의 인덱스를 가져오기
        selected_index = listbox.curselection()
        keyword = keyword_entry.get().strip()
        new_name = name_entry.get().strip()

        if keyword and new_name:
            # 키워드 추가 로직
            keyword_mapping[keyword] = new_name

            # 선택된 인덱스 아래에 추가하기
            if selected_index:
                insert_position = selected_index[0] + 1  # 선택된 항목의 다음 위치
            else:
                insert_position = listbox.size()  # 선택된 항목이 없으면 맨 아래 추가

            # 리스트박스에 새 항목 삽입
            listbox.insert(insert_position, f"{keyword} -> {new_name}")

            # 키워드 매핑을 정렬된 순서로 업데이트
            updated_mapping = {}
            for i in range(listbox.size()):
                item = listbox.get(i)
                k, v = item.split(" -> ")
                updated_mapping[k.strip()] = v.strip()
            keyword_mapping = updated_mapping

            # JSON 파일에 저장
            save_keywords()

            # 추가된 항목으로 스크롤 및 선택
            listbox.selection_clear(0, tk.END)  # 이전 선택 해제
            listbox.selection_set(insert_position)  # 새 항목 선택
            listbox.see(insert_position)  # 추가된 항목 위치로 이동

            # 메시지 표시
            messagebox.showinfo("추가 완료", f"키워드 '{keyword}'가 추가되었습니다.", parent=keyword_window)

            # 입력 필드 비우기 및 키워드 필드로 포커스 이동
            keyword_entry.delete(0, tk.END)
            name_entry.delete(0, tk.END)
            keyword_entry.focus_set()
        else:
            messagebox.showwarning("입력 필요", "키워드와 상품명을 모두 입력하세요.", parent=keyword_window)

    def delete_keyword():
        selected = listbox.curselection()
        if selected:
            # 선택된 항목의 인덱스
            selected_index = selected[0]
            keyword = listbox.get(selected_index).split(" -> ")[0].strip()

            # 키워드 삭제
            if keyword in keyword_mapping:
                del keyword_mapping[keyword]
                save_keywords()
                refresh_list()
                messagebox.showinfo("삭제 완료", f"키워드 '{keyword}'가 삭제되었습니다.", parent=keyword_window)
                # 삭제 후에도 리스트박스 위치 유지
                if selected_index < listbox.size():  # 삭제된 항목이 마지막이 아닐 때
                    listbox.selection_set(selected_index)  # 다음 항목 선택
                    listbox.see(selected_index)  # 선택한 항목이 보이게 유지
                elif listbox.size() > 0:  # 삭제된 항목이 마지막이고 리스트에 다른 항목이 있을 때
                    listbox.selection_set(selected_index - 1)  # 이전 항목 선택
                    listbox.see(selected_index - 1)  # 이전 항목 위치로 유지
            else:
                messagebox.showwarning("키워드 없음", "해당 키워드가 존재하지 않습니다.", parent=keyword_window)
        else:
            messagebox.showwarning("선택 필요", "삭제할 키워드를 선택하세요.", parent=keyword_window)

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
            
    def listbox_deselect(event):
        if event.widget not in (keyword_entry, name_entry,add_button, delete_button, overview_button, listbox):
            listbox.selection_clear(0, tk.END)

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

        # 텍스트를 읽기 전용으로 설정
        text_box.config(state="disabled")  # 읽기 전용 모드

        # Ctrl + F 키로 찾기 기능 추가
        def search_text(event=None):
            search_window = tk.Toplevel()
            search_window.title("한눈에 찾기")
            search_window.geometry("300x100")
            search_window.configure(bg="#f0f8ff")
            search_window.resizable(False, False)

            tk.Label(search_window, text="찾을 단어:", bg="#f0f8ff", fg="#000000").pack(pady=5)
            search_entry = tk.Entry(search_window)
            search_entry.pack(pady=5)
            search_entry.focus_set()  # 입력창에 포커스 설정

            # 검색 위치를 기억하기 위한 변수
            search_index = {"current": "1.0"}

            def perform_search():
                search_query = search_entry.get().strip()
                if not search_query:
                    messagebox.showwarning("입력 필요", "검색어를 입력하세요.", parent=search_window)
                    return

                # 현재 위치에서 다음 검색 시작
                pos = text_box.search(search_query, search_index["current"], tk.END)
                if pos:  
                    # 이전 강조 제거 후 현재 검색어 강조
                    text_box.tag_remove("highlight", "1.0", tk.END)
                    text_box.tag_add("highlight", pos, f"{pos} + {len(search_query)}c")
                    text_box.tag_config("highlight", background="yellow", foreground="black")
                    text_box.see(pos)

                    # 다음 검색을 위해 인덱스 갱신
                    search_index["current"] = f"{pos} + 1c"
                else:
                    search_index["current"] = "1.0"  # 검색이 끝나면 처음부터 다시 시작

            search_button = tk.Button(search_window, text="검색", command=perform_search, bg="#add8e6", fg="#000000")
            search_button.pack(pady=5)

            # 엔터 키 이벤트 바인딩
            search_entry.bind("<Return>", lambda event: perform_search())

        overview_window.bind("<Control-f>", search_text)  # Ctrl + f 바인딩 추가
        overview_window.bind("<Control-F>", search_text)  # Ctrl + F 바인딩 추가

    def search_keyword():
        search_window = tk.Toplevel()
        search_window.title("관리 단어찾기")
        search_window.geometry("300x100")
        search_window.configure(bg="#f0f8ff")
        search_window.resizable(False, False)

        tk.Label(search_window, text="찾을 키워드:", bg="#f0f8ff", fg="#000000").pack(pady=5)
        search_entry = tk.Entry(search_window)
        search_entry.pack(pady=5)
        search_entry.focus_set()  # 입력창에 포커스 설정

        # 검색 위치를 기억하기 위한 인덱스
        search_index = {'current': 0}  

        def perform_search(event=None):
            search_text = search_entry.get().strip()
            if not search_text:
                messagebox.showwarning("입력 필요", "검색할 키워드를 입력하세요.", parent=search_window)
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
                messagebox.showinfo("검색 결과", f"'{search_text}'에 해당하는 키워드를 찾을 수 없습니다.", parent=search_window)
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
    
    # Tab 키 이벤트 설정
    def setup_tab_navigation():
        def custom_tab_handler(event):
            focused_widget = event.widget
            if focused_widget == keyword_entry:
                name_entry.focus_set()
            elif focused_widget == name_entry:
                keyword_entry.focus_set()
            return "break"

        keyword_entry.bind("<Tab>", custom_tab_handler)
        name_entry.bind("<Tab>", custom_tab_handler)

    # Tab 키 동작 설정 함수 호출
    setup_tab_navigation()
    
    keyword_window.bind("<Button-1>", listbox_deselect)
    # Enter 키로 추가 버튼 클릭
    keyword_window.bind('<Return>', lambda event: add_keyword())

    # Ctrl + F 단축키로 검색 창 열기
    keyword_window.bind('<Control-f>', lambda event: search_keyword())
    keyword_window.bind('<Control-F>', lambda event: search_keyword())

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

            # 옵션정보에서 특정 키워드를 상품명 뒤에 붙이기 (단, product_name에 "코일매트"가 없을 때만)
            if "코일" not in product_name:
                for keyword in option_keywords:
                    if keyword in option_info:
                        product_name += f" {keyword}"

            if "듀얼팩" in combined_text:
                product_name = f"{product_name} 듀얼팩"
            
            if "주차번호판 x 2개" in combined_text:
                product_name = f"{product_name} x 2개"
            
            if "사이드미러형" in combined_text:
                product_name = f"{product_name} 사이드미러형"
                
            if "블랙박스형" in combined_text:
                product_name = f"{product_name} 블랙박스형"
            
            if "（우드）" in combined_text:
                product_name = f"{product_name} 우드"
            
            if "1+1+1+1" in combined_text:
                product_name = f"{product_name} 1+1+1+1"
                
            elif "1+1+1" in combined_text:
                product_name = f"{product_name} 1+1+1" 
            # 상품명이나 옵션정보에 "1+1", "1set+1set"이 있으면 상품명 앞에 추가
            elif ("1+1" in combined_text or "1set+1set" in combined_text) and "미니 브러쉬" not in combined_text:
                product_name = f"{product_name} 1+1"
            
            if "1+2" in combined_text and "코일" not in combined_text and "엣지형 변경" not in combined_text:
                product_name = f"{product_name} 1+2"
            
            if "2+1" in combined_text:
                product_name = f"{product_name} 2+1"
                
            if "2+2" in combined_text:
                product_name = f"{product_name} 2+2"
                
            if "3+1" in combined_text:
                product_name = f"{product_name} 3+1"
            
            if "3+2" in combined_text:
                product_name = f"{product_name} 3+2"
            
            if "3+3" in combined_text:
                product_name = f"{product_name} 3+3"
                
            if "4+1" in combined_text:
                product_name = f"{product_name} 4+1"
                
            if "4+2" in combined_text:
                product_name = f"{product_name} 4+2"
                
            if "4+3" in combined_text:
                product_name = f"{product_name} 4+3"
                
            if "4+4" in combined_text:
                product_name = f"{product_name} 4+4"
                
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
        base_name = os.path.splitext(os.path.basename(file_path))[0]  # 파일명만 추출
        directory = os.path.dirname(file_path)  # 파일 경로 추출
        save_path = os.path.join(directory, f"res_{base_name}.xlsx")  # 경로와 파일명 결합
        counter = 1
        while os.path.exists(save_path):
            save_path = os.path.join(directory, f"res_{base_name}_{counter}.xlsx")  # 중복 시 번호 추가
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