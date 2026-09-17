def add_columns(df):
    #근무시간(시간) 열 뒤에 새로운 열 2개 추가

    target_col='근무시간(분)'

    if target_col in df.columns:
        #1. 대상 열 위치 찾기
        idx = df.columns.get_loc(target_col)

        #2. 열 삽입 (대상 열 바로 뒤이므로 idx+1)
        if '실제근무시간(K-O-Q)' not in df.columns:
            df.insert(idx+1, '실제근무시간(K-O-Q)', '')
        if '실제근무시간(시.분)' not in df.columns:
            df.insert(idx+2, '실제근무시간(시.분)', '')

    else:
        print(f"경고: '{target_col}'열을 찾을수 없습니다.")

    return df