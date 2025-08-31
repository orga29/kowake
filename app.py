import streamlit as st
import pandas as pd 
from kowake import create_repacking_priority_list_from_excel 
# import datetime # サブタイトル削除のため不要に

# --- Streamlit アプリケーションのUI設定 ---
st.set_page_config(page_title="小分け作業用", layout="wide") # page_titleも変更

st.title("小分け作業用") # タイトルを変更

# サブタイトル「（日付） 小分け作成メモ」を削除
# current_date_mmdd = datetime.datetime.now().strftime("%m月%d日")
# st.markdown(f"### {current_date_mmdd} 小分け作成メモ") 

st.markdown("---") # 区切り線

# ファイルアップローダーの設置
uploaded_file = st.file_uploader("処理対象のExcelファイル（.xlsx）をアップロードしてください", type=["xlsx"])

if uploaded_file is not None:
    st.markdown("---")
    st.info(f"アップロードファイル名: {uploaded_file.name}")

    # 処理実行ボタン
    if st.button("処理実行"):
        with st.spinner("処理中です... しばらくお待ちください。"):
            success, message, output_filename, excel_data = create_repacking_priority_list_from_excel(uploaded_file)

        if success:
            if output_filename and excel_data:
                st.success(f"{message}")
                st.download_button(
                    label=f"{output_filename} をダウンロード",
                    data=excel_data,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info(message)
        else:
            st.error(message)
else:
    st.info("上記よりExcelファイルをアップロードして処理を開始してください。")

st.markdown("---")
st.markdown("#### 補足説明")
st.markdown("""
- Excelファイルをアップロードしたら、「処理実行」ボタンを押してください。
- 処理結果は、ファイルとしてダウンロードできます。
- 「充足率」は「納品数」に対する「昨日残数」の割合です。（昨日残数÷納品数×100）
""")
