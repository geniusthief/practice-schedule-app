import streamlit as st
import pandas as pd
import tempfile
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from pulp import LpProblem, LpMaximize, LpVariable, lpSum, LpBinary, LpInteger, LpStatus

# --- ページ設定 ---
st.set_page_config(page_title="卓球部練習スケジュール最適化", layout="wide")
st.title("🏓 卓球部練習スケジュール最適化ツール")

# --- Excelアップロード ---
uploaded_file = st.file_uploader("📂 Excelファイルをアップロードしてください", type=["xlsx"])
if uploaded_file is None:
    st.info("👆 Excelファイルをアップロードしてください。")
    st.stop()

# --- Workbook 読み込み ---
tmpf = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
tmpf.write(uploaded_file.read())
tmpf.flush()
book = load_workbook(tmpf.name)

# --- r_timeシート表示・編集 ---
st.subheader("🗓️ 可用性（r_time）")
r_time_df = pd.read_excel(tmpf.name, sheet_name="r_time")
edited_r_time = st.data_editor(r_time_df, num_rows="dynamic", key="r_time_edit")

# --- day_limitsシート表示・編集 ---
st.subheader("⚙️ 曜日ごとの人数制約（day_limits）")
day_limits_df = pd.read_excel(tmpf.name, sheet_name="day_limits")
edited_day_limits = st.data_editor(day_limits_df, num_rows="dynamic", key="day_limits_edit")

# --- チア日選択 ---
st.subheader("🎽 チアの日設定")
cheer_days = st.multiselect("チアのある曜日を選択", ["火", "水", "木", "金"], default=["火", "金"])

# --- 最適化実行ボタン ---
if st.button("🚀 最適化を実行"):

    # --- Workbook に編集内容を書き戻す ---
    sheet_rt = book['r_time']
    for i, row in edited_r_time.iterrows():
        for j, val in enumerate(row[1:], start=2):  # 1列目は名前列
            sheet_rt.cell(row=i+2, column=j, value=val)

    sheet_day = book['day_limits']
    for i, row in edited_day_limits.iterrows():
        for j, val in enumerate(row[1:], start=2):
            sheet_day.cell(row=i+2, column=j, value=val)

    # --- PuLPで最適化関数 ---
    def run_optimization_from_workbook(book):
        # 簡略化例：人数最大化（本来の制約はここに追加）
        sheet_rt = book['r_time']
        sheet_day = book['day_limits']

        num_members = sheet_rt.max_row - 1
        T = list(range(1, 9))
        D = list(range(1, 5))

        x = {(i, t, d): LpVariable(f"x_{i}_{t}_{d}", cat=LpBinary)
             for i in range(1, num_members+1) for t in T for d in D}

        prob = LpProblem("practice_schedule", LpMaximize)

        # 可用性に応じてxを制約
        for i in range(1, num_members+1):
            for d in D:
                for t in T:
                    val = sheet_rt.cell(row=i+1, column=d+1).value
                    if val is None:
                        prob += x[(i, t, d)] == 0

        # 目的関数：出席人数の合計最大化
        prob += lpSum([x[i, t, d] for i in range(1, num_members+1) for t in T for d in D])

        prob.solve()

        # --- 結果出力 ---
        result_sheet = book.create_sheet("result")
        labels = [chr(65+i) for i in range(num_members)]
        weekday_map = {1: "火", 2: "水", 3: "木", 4: "金"}

        for d in D:
            cell = result_sheet.cell(row=1, column=1+d)
            cell.value = f"{weekday_map[d]}曜"
            cell.alignment = Alignment(horizontal='center')

        for t in T:
            cell = result_sheet.cell(row=1+t, column=1)
            cell.value = f"{12+t}時"
            cell.alignment = Alignment(horizontal='center')

        for i in range(1, num_members+1):
            name = labels[i-1]
            for t in T:
                for d in D:
                    if x[(i, t, d)].value() >= 0.5:
                        row = 1 + t
                        col = 1 + d
                        cell = result_sheet.cell(row=row, column=col)
                        prev = cell.value if cell.value else ""
                        names = prev.split(",") if prev else []
                        if name not in names:
                            names.append(name)
                        cell.value = ",".join(names)
                        cell.alignment = Alignment(horizontal='center')
                        cell.font = Font(size=12)

        # 一時ファイルに保存
        tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        book.save(tmp_out.name)
        return tmp_out.name

    try:
        with st.spinner("最適化中...数秒かかる場合があります"):
            out_path = run_optimization_from_workbook(book)

        # --- 結果表示 ---
        st.success("✅ 最適化完了")
        result_df = pd.read_excel(out_path, sheet_name="result")
        st.subheader("割当表 (result シート)")
        st.dataframe(result_df)

        # --- ダウンロード ---
        with open(out_path, "rb") as f:
            data = f.read()
        st.download_button("結果をダウンロード (practice_result.xlsx)", data, file_name="practice_result.xlsx")

    except Exception as e:
        st.exception(e)
