import streamlit as st
import pandas as pd
from io import BytesIO

# ページ設定
st.set_page_config(page_title="卓球部練習スケジュール最適化", layout="wide")

st.title("🏓 練習スケジュール最適化ツール")

# --- Excel 読み込み ---
uploaded_file = st.file_uploader("📂 Excelファイルをアップロードしてください", type=["xlsx"])

if uploaded_file is not None:
    book = pd.ExcelFile(uploaded_file)

    # r_timeシート表示・編集
    st.subheader("🗓️ 可用性（r_time）")
    r_time = pd.read_excel(book, sheet_name="r_time")
    edited_r_time = st.data_editor(r_time, num_rows="dynamic", key="r_time_edit")

    # day_limitsシート表示・編集
    st.subheader("⚙️ 曜日ごとの人数制約（day_limits）")
    day_limits = pd.read_excel(book, sheet_name="day_limits")
    edited_day_limits = st.data_editor(day_limits, num_rows="dynamic", key="day_limits_edit")

    # チアの日を選択
    st.subheader("🎽 チアの日設定")
    cheer_days = st.multiselect("チアのある曜日を選択", ["火", "水", "木", "金"], default=["火", "金"])

    # 最適化実行ボタン
    if st.button("🚀 最適化を実行"):
    info = run_optimization_from_workbook(book)

        st.write("最適化を実行中...")

        # 仮の結果表示
        result = pd.DataFrame({
            "曜日": ["火", "水", "木", "金"],
            "人数": [8, 12, 10, 7],
        })
        st.success("最適化が完了しました ✅")
        st.dataframe(result)

else:
    st.info("👆 Excelファイルをアップロードしてください。")

