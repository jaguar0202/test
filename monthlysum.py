import streamlit as st
import pandas as pd
from io import BytesIO
import re

# "요약"에서 단어수와 기준 언어 추출
def extract_info_from_summary(요약):
    if pd.isna(요약) or not isinstance(요약, str):  # NaN 체크 및 문자열 확인
        return pd.Series([None, None])
    
    match = re.search(r'\[(\d+)\s*([A-Za-z]+)\]', 요약)  # 단어수와 기준 언어 추출
    if match:
        word_count = match.group(1)  # 숫자 (단어수)
        language = match.group(2)  # 영어 코드 (기준 언어)
        return pd.Series([word_count, language])
    return pd.Series([None, None])

# Streamlit UI 구성
st.title("📊 Jira CSV 데이터 추출기")
st.write("CSV 파일을 업로드하면 필요한 데이터만 추출하여 엑셀로 다운로드할 수 있습니다.")

# 파일 업로드
uploaded_file = st.file_uploader("📂 CSV 파일을 업로드하세요", type=["csv"])

if uploaded_file is not None:
    # CSV 읽기
    df = pd.read_csv(uploaded_file)

    # 필요한 열만 선택 (해당 열이 없으면 기본값 사용)
    selected_columns = ['프로젝트 이름', '요약', '기한', '생성일']
    for col in selected_columns:
        if col not in df.columns:
            df[col] = None  # 없는 열은 빈 값 생성

    df_filtered = df[selected_columns].copy()

    # "기한"과 "생성일" 열에서 처음 10개의 문자만 추출
    df_filtered['기한'] = df_filtered['기한'].astype(str).str[:10]
    df_filtered['생성일'] = df_filtered['생성일'].astype(str).str[:10]

    # "요약"에서 단어수와 기준 언어 추출
    df_filtered[['단어수', '기준 언어']] = df_filtered['요약'].apply(extract_info_from_summary)

    # '단어수' 열을 숫자로 변환
    df_filtered['단어수'] = pd.to_numeric(df_filtered['단어수'], errors='coerce')

    # 프로젝트별로 항목 개수 및 단어수 합계 계산
    project_summary_df = df_filtered.groupby('프로젝트 이름').agg(
        요청수=('요약', 'count'),
        단어수_합계=('단어수', 'sum')
    ).reset_index()

    # 요청수 열 기준으로 내림차순 정렬
    project_summary_df = project_summary_df.sort_values(by='요청수', ascending=False)

    # 합계 행 추가
    total_row = project_summary_df[['요청수', '단어수_합계']].sum().to_frame().T
    total_row['프로젝트 이름'] = '합계'
    project_summary_df = pd.concat([project_summary_df, total_row], ignore_index=True)

    # 데이터 미리보기
    st.write("📋 **원본 데이터**")
    st.dataframe(df_filtered.head())  # 상위 5개 데이터 미리 보기

    # 엑셀로 변환
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 원본 데이터 시트
        df_filtered.to_excel(writer, index=False, sheet_name="원본 데이터")
        
        # 프로젝트별 요약 시트
        project_summary_df.to_excel(writer, index=False, sheet_name="프로젝트별 요약")
        
        writer.close()
    output.seek(0)

    # 다운로드 버튼
    st.download_button(
        label="📥 엑셀 파일 다운로드",
        data=output,
        file_name="project_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
