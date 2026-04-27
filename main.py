import pandas as pd
import os
import customtkinter as ctk
from tkinter import filedialog, messagebox


from utils.delete_column import delete_column
from utils.modify_rows import modify_rows
from utils.add_columns import add_columns
from utils.apply_border import apply_border
from utils.split_by_name import split_by_name
from utils.insert_blank_rows import insert_blank_rows
from utils.insert_formula import insert_formula

selected_file_path = ""


def select_file():
    global selected_file_path
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        selected_file_path = path
        label.configure(text=f"선택됨: {os.path.basename(path)}")

        # #파일명만 추출해서 라벨에 표시
        # filename = os.path.basename(path)
        # label.configure(text=f"선택된 파일: {filename}")
        # print(f"선택된 경로: {selected_file_path}")



########################


def run_automation():
    global selected_file_path

    if not selected_file_path:
        messagebox.showwarning("경고", "먼저 엑셀 파일을 선택해 주세요")
        return
    
    btn_run.configure(state="disabled")  # 중복 실행 방지
    try:
        #1. 파일 불러오기
        df = pd.read_excel(selected_file_path)
    
        #2. pandas 로직
        #2-1. 열삭제
        df = delete_column(df)
        #2-2. CEO 행 삭제
        df = modify_rows(df)
        #2-3. 열 삽입
        df = add_columns(df)
        
        
        #3. 결과 저장
        base, ext = os.path.splitext(selected_file_path)
        if base.endswith("_결과"):
            save_path = f"{base}{ext}"
        else:
            save_path = f"{base}_결과{ext}"
        
        df.to_excel(save_path, index=False)


        #4. openpyxl 처리
        #4-1. 전체 시트 테두리 적용
        apply_border(save_path, skip_title=None)
        #4-2. 시트 분리 (이름별)
        split_by_name(save_path)
        #4-3. 개인 시트 테두리 적용
        apply_border(save_path, skip_title='전체')  # 전체 시트 건너뜀
        #4-4. 빈 행 삽입 (일요일 기준) + 노란색
        insert_blank_rows(save_path)
        #4-5. 수식 삽입 (저장 후 openpyxl로 후처리)
        insert_formula(save_path)

        label.configure(text=f"작업 완료! 결과 파일 생성됨")
        messagebox.showinfo("성공", f"파일이 성공적으로 저장되었습니다.\n{save_path}")
        
    except Exception as e:
        messagebox.showerror("에러", f"오류가 발생했습니다: {e}")
    finally:
        btn_run.configure(state="normal")

#UI 설정
app = ctk.CTk()
app.title("근태 정리 자동화 프로그램")
app.geometry("500x300")

label = ctk.CTkLabel(app, text="엑셀 파일을 선택하세요", font=("Pretendard", 14))
label.pack(pady=30)

#버튼 1: 파일 선택
btn_select = ctk.CTkButton(app, text="파일 불러오기", command=select_file)
btn_select.pack(pady=10)

#버튼 2: 실행
btn_run = ctk.CTkButton(app, text="자동화 실행", command=run_automation, fg_color="green", hover_color="darkgreen")
btn_run.pack(pady=10)

app.mainloop()