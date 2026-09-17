def delete_column(df):
    # 1. 열 삭제
    cols_to_delete = ['사번', '출근스케줄', '퇴근스케줄', '휴게시간', '근무시간', '제외시간', '연장근무시간(신청)', '야간근무시간(신청)', 
                      '야간근무시간(실제)', '외근-간주시간(신청)', '제외시간(분)', '연장근무시간(신청)(분)', '야간근무시간(신청)(분)',
                      '외근-간주시간(신청)(분)', '출근입력시간', '퇴근입력시간', '자리비움시간(RAW)', '외근-간주시간']
    
    existing_cols = [col for col in cols_to_delete if col in df.columns]
    if existing_cols:
        df = df.drop(columns=existing_cols)
    
    return df