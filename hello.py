import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
import tempfile
from pulp import LpProblem, LpMaximize, LpVariable, lpSum, LpBinary, LpInteger, LpStatus

# --- ページ設定 ---
st.set_page_config(page_title="卓球部練習スケジュール最適化", layout="wide")
st.title("🏓 卓球部 練習シフト最適化ツール (完全版)")

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

# --- 重み設定（任意） ---
with st.sidebar:
    st.header("重みパラメータ")
    w1 = st.number_input("授業直後スコア (w1)", value=100.0)
    w2 = st.number_input("連続練習スコア (w2)", value=0.5)
    w3 = st.number_input("人数スコア (w3)", value=1.0)

# --- 最適化関数 ---
def run_optimization_from_workbook(book, cheer_days, w1, w2, w3):
    # --- 部員数・曜日・時間帯設定 ---
    sheet_rt = book['r_time']
    sheet_len = book['w_len']
    sheet_day = book['day_limits']

    num_members = sheet_rt.max_row - 1
    I = list(range(1, num_members+1))
    T = list(range(1, 9))  # 13-20時
    D = list(range(1, 5))  # 火-金
    L = [1, 2, 3]          # 連続練習の長さ候補
    L_s = {s: [l for l in L if s + l - 1 <= 8] for s in T}  # 開始時刻ごとの可能長さ

    labels = [chr(65+i) for i in range(num_members)]

    # --- r_time マップ（Excel値→開始時間） ---
    r_map = {2: 1, 3: 3, 4: 5, 5: 7}
    r_time = {}
    for i in I:
        for d in D:
            val = sheet_rt.cell(row=i+1, column=d+1).value
            r_time[i,d] = r_map.get(val, None)

    # --- 可用性マップ a[i,t,d] ---
    a = {}
    for i in I:
        for d in D:
            for t in T:
                a[i,t,d] = 0
            start = r_time[i,d]
            if start:
                for t in range(start, 9):
                    a[i,t,d] = 1

    # --- w_len 読み込み ---
    w_len = {l: sheet_len.cell(row=l+1, column=2).value for l in L}

    # --- day_limits 読み込み（チアの日判定） ---
    day_chia = {}
    weekday_map = {1:'火',2:'水',3:'木',4:'金'}
    for d in D:
        day_name = weekday_map[d]
        day_chia[d] = day_name in cheer_days

    # --- 人数制約 min/max ---
    day_min = {}
    day_max = {}
    for d in D:
        if day_chia[d]:
            day_min[d] = 3
            day_max[d] = 8
        else:
            day_min[d] = 3
            day_max[d] = 16

    # --- 人数評価 w_num ---
    N_range = list(range(3, num_members+1))
    w_num = {d:{} for d in D}
    for d in D:
        ideal = (day_min[d]+day_max[d])//2 + 1
        for n in N_range:
            w_num[d][n] = max(0.0, 1.0 - 0.1 * abs(n - ideal))

    # --- PuLP 変数定義 ---
    prob = LpProblem("practice_schedule", LpMaximize)
    x = {(i,t,d): LpVariable(f"x_{i}_{t}_{d}", cat=LpBinary) for i in I for t in T for d in D}
    y = {(t,d): LpVariable(f"y_{t}_{d}", cat=LpBinary) for t in T for d in D}
    forbidden_start = [4,6,8]
    z = {}
    for i in I:
        for d in D:
            for s in T:
                if s in forbidden_start: continue
                for l in L_s[s]:
                    z[(i,s,d,l)] = LpVariable(f"z_{i}_{s}_{d}_{l}", cat=LpBinary)
    v = {(t,d,n): LpVariable(f"v_{t}_{d}_{n}", cat=LpBinary) for t in T for d in D for n in N_range}
    num_td = {(t,d): LpVariable(f"num_{t}_{d}", lowBound=0, cat=LpInteger) for t in T for d in D}

    # --- 制約 ---
    # 可用性
    for i in I:
        for t in T:
            for d in D:
                if a[i,t,d]==0:
                    prob += x[i,t,d]==0

    # 週3回以上
    for i in I:
        prob += lpSum([x[i,t,d] for t in T for d in D]) >= 3

    # 一日の最大3時間
    for i in I:
        for d in D:
            prob += lpSum([x[i,t,d] for t in T]) <= 3

    # 人数制約
    for t in T:
        for d in D:
            prob += num_td[t,d] == lpSum([x[i,t,d] for i in I])
            prob += num_td[t,d] >= day_min[d]
            prob += num_td[t,d] <= day_max[d]

    # zの一意性
    for i in I:
        for d in D:
            prob += lpSum([z[(i,s,d,l)] for s in T if s not in forbidden_start for l in L_s[s] if (i,s,d,l) in z]) <= 1

    # z->x
    for i in I:
        for d in D:
            for t in T:
                prob += x[i,t,d] == lpSum([z[(i,s,d,l)] for s in T if s not in forbidden_start for l in L_s[s] if (i,s,d,l) in z and s <= t < s+l])

    # 飛び飛び禁止
    for i in I:
        for d in D:
            for t in T[1:-1]:
                prob += x[i,t,d] <= x[i,t-1,d] + x[i,t+1,d]

    # v一意 + num_tdに対応
    M = len(I)
    for t in T:
        for d in D:
            prob += lpSum([v[t,d,n] for n in N_range]) == 1
            for n in N_range:
                prob += num_td[t,d]-n <= (1-v[t,d,n])*M
                prob += n-num_td[t,d] <= (1-v[t,d,n])*M

    # --- 目的関数 ---
    term1 = lpSum([x[i,r_time[i,d],d] for i in I for d in D if r_time[i,d] is not None and r_time[i,d] in T])
    term2 = lpSum([w_len[l]*z[(i,s,d,l)] for i in I for d in D for s in T if s not in forbidden_start for l in L_s[s] if (i,s,d,l) in z])
    term3 = lpSum([w_num[d][n]*v[(t,d,n)] for t in T for d in D for n in N_range])
    prob += w1*term1 + w2*term2 + w3*term3

    # --- solve ---
    prob.solve()

    # --- 結果出力 ---
    if 'result' in book.sheetnames:
        book.remove(book['result'])
    result_sheet = book.create_sheet('result')

    for d in D:
        cell = result_sheet.cell(row=1, column=1+d)
        cell.value = f"{weekday_map[d]}曜"
        cell.alignment = Alignment(horizontal='center')

    for t in T:
        cell = result_sheet.cell(row=1+t, column=1)
        cell.value = f"{12+t}時"
        cell.alignment = Alignment(horizontal='center')

    for i in I:
        name = labels[i-1]
        for t in T:
            for d in D:
                if x[(i,t,d)].value() is not None and x[(i,t,d)].value()>=0.5:
                    row = 1+t
                    col = 1+d
                    cell = result_sheet.cell(row=row, column=col)
                    prev = cell.value if cell.value else ''
                    names = prev.split(',') if prev else []
                    if name not in names:
                        names.append(name)
                    cell.value = ','.join(names)
                    cell.alignment = Alignment(horizontal='center')
                    cell.font = Font(size=12)

    # 一時ファイルに保存
    tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    book.save(tmp_out.name)
    return tmp_out.name, LpStatus[prob.status]

# --- 実行 ---
try:
    with st.spinner("最適化中...（数秒〜数分かかる場合があります）"):
        out_path, status = run_optimization_from_workbook(book, cheer_days, w1, w2, w3)

    st.success(f"✅ 最適化完了（ステータス: {status}）")
    result_df = pd.read_excel(out_path, sheet_name='result')
    st.subheader("割当表 (result シート)")
    st.dataframe(result_df)

    with open(out_path, 'rb') as f:
        data = f.read()
    st.download_button("結果をダウンロード (practice_result.xlsx)", data, file_name="practice_result.xlsx")

except Exception as e:
    st.exception(e)
