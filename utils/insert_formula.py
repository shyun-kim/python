import openpyxl
from openpyxl.utils import get_column_letter

def insert_formula(save_path):
    wb = openpyxl.load_workbook(save_path)

    for ws in wb.worksheets:
        if ws.title == '전체':
            continue

        # 1. 헤더(1행)에서 필요한 열 인덱스 동적 추출
        header = [cell.value for cell in ws[1]]

        try:
            # 기존 뺄셈 수식용 열 찾기
            col_K = header.index('근무시간(분)') + 1
            col_O = header.index('저녁시간(분)') + 1
            col_P = header.index('출근시간전(분)') + 1
            col_result = header.index('실제근무시간(K-O-Q)') + 1

            # 주차별 SUM 수식을 입력할 J~Q열의 시작/끝 인덱스 찾기
            sum_start_col = header.index('휴게시간(분)') + 1   # J열 역할
            sum_end_col = header.index('퇴근시간후(분)') + 1     # Q열 역할
        except ValueError as e:
            print(f"[{ws.title}] 열을 찾을 수 없습니다: {e}")
            continue        

        k = get_column_letter(col_K)
        o = get_column_letter(col_O)
        p = get_column_letter(col_P)
        r = get_column_letter(col_result)

        # 2. 행 탐색 및 수식 집어넣기
        last_yellow_row = 1  # 주차 구분 기준점 (헤더인 1행부터 시작)

        for row in range(2, ws.max_row + 1):
            # [A] 노란색 합계 행인 경우 (Team/A열 데이터가 비어있는 행)
            if ws[f'A{row}'].value in (None, ''):
                
                # 첫 노란 행일 때: 2행이 헤더/텍스트면 데이터 시작은 3행부터
                if last_yellow_row == 1:
                    val_check = ws[f'{k}2'].value
                    start_row = 2 if (val_check in (None, '') or isinstance(val_check, str)) else 2
                else:
                    start_row = last_yellow_row + 1

                end_row = row - 1

                # 해당 주차 범위(J~Q열)에 SUM 수식 일괄 기입
                if start_row <= end_row:
                    for col_idx in range(sum_start_col, sum_end_col + 1):
                        col_letter = get_column_letter(col_idx)
                        ws[f'{col_letter}{row}'] = f"=SUM({col_letter}{start_row}:{col_letter}{end_row})"

                # 다음 주차 계산을 위해 기준점 업데이트
                last_yellow_row = row

            # [B] 일반 데이터 행인 경우 -> 기존 실제근무시간 뺄셈 수식 작성
            else:
                # 근무시간이 비어있는 휴일 등은 계산 생략
                if ws[f'{k}{row}'].value not in (None, ''):
                    ws[f'{r}{row}'] = f'={k}{row}-({o}{row}+{p}{row})'

        # 3. 틀고정 설정
        ws.freeze_panes = 'K2'

    wb.save(save_path)
    wb.close()