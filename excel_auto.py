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
        label.configure(text=f"선택됨: {os.path.basename(path)}")

        # #파일명만 추출해서 라벨에 표시
        # filename = os.path.basename(path)
        # label.configure(text=f"선택된 파일: {filename}")
        # print(f"선택된 경로: {selected_file_path}")

def delete_column(df):
    # 1. 열 삭제
    cols_to_delete = ['사번', '출근스케줄', '퇴근스케줄', '휴게시간', '근무시간', '제외시간', '연장근무시간(신청)', '야간근무시간(신청)', 
                      '야간근무시간(실제)', '외근-간주시간(신청)', '제외시간(분)', '연장근무시간(신청)(분)', '야간근무시간(신청)(분)',
                      '외근-간주시간(신청)(분)', '출근입력시간', '퇴근입력시간', '자리비움시간(RAW)', '외근-간주시간']
    
    existing_cols = [col for col in cols_to_delete if col in df.columns]
    if existing_cols:
        df = df.drop(columns=existing_cols)
    
    return df

def modify_rows(df):
    #특정 조건 행 삭제(CEO)
    #1. "a"열 값이 CEO가 아닌 데이터만 추출
    if 'Team' in df.columns:
        df=df[df['Team'] != 'CEO']
        
    return df

def add_columns(df):
    #근무시간(시간) 열 뒤에 새로운 열 2개 추가

    target_col='근무시간(분)'

    if target_col in df.columns:
        #1. 대상 열 위치 찾기
        idx = df.columns.get_loc(target_col)

        #2. 첫 번째 열 삽입 (대상 열 바로 뒤이므로 idx+1)
        df.insert(idx+1, '실제근무시간(K-O-Q)','')

        #3. 두 번째 열 삽입
        df.insert(idx+2, '실제근무시간(시.분)','')

    else:
        print(f"경고: '{target_col}'열을 찾을수 없습니다.")

    return df

def run_automation():
    global selected_file_path

    if not selected_file_path:
        messagebox.showwarning("경고", "먼저 엑셀 파일을 선택해 주세요")
        return
    
    try:
        #1. 파일 불러오기
        df = pd.read_excel(selected_file_path)
    
        #2. 자동화 로직
        #열삭제
        df = delete_column(df)
        #CEO 행 삭제
        df = modify_rows(df)
        #행 삽입
        df = add_columns(df)
        
        #3. 결과 저장
        base, ext = os.path.splitext(selected_file_path)
        save_path = f"{base}_결과{ext}"
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