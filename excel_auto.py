import pandas as pd
import os
import customtkinter as ctk
from tkinter import filedialog, messagebox

selected_file_path = ""

def select_file():
    global selected_file_path
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

    if path:
        selected_file_path = path
        #파일명만 추출해서 라벨에 표시
        filename = os.path.basename(path)
        label.configure(text=f"선택된 파일: {filename}")
        print(f"선택된 경로: {selected_file_path}")

def run_automation():
    global selected_file_path

    if not selected_file_path:
        messagebox.showwarning("경고", "먼저 엑셀 파일을 선택해 주세요")
        return
    
    try:
        # 1. 파일 불러오기
        df = pd.read_excel(selected_file_path)
    
        # -------- 자동화 로직 추가 위치 --------
        
        #임시로직
        df['완료여부'] = '확인됨'

        # ------------------------------------
        # 2. os.path.splitext를 이용한 저장 경로 설정
        base, ext = os.path.splitext(selected_file_path)
        save_path = f"{base}_결과{ext}"

        # 3. 저장
        df.to_excel(save_path, index=False)
        label.configure(text=f"작업 완료! 결과 파일 생성됨")
        messagebox.showinfo("성공", f"파일이 성공적으로 저장되었습니다.\n{save_path}")
        
    except Exception as e:
        messagebox.showerror("에러", f"오류가 발생했습니다: {e}")


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