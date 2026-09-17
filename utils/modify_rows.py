def modify_rows(df):
    #특정 조건 행 삭제(CEO)
    #"a"열 값이 CEO가 아닌 데이터만 추출
    if 'Team' in df.columns:
        df=df[df['Team'] != 'CEO']
        
    return df