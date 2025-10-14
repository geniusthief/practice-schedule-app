import streamlit as st
from mip import Model, xsum, BINARY,maximize, INTEGER, OptimizationStatus
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
import string
import tempfile
import pandas as pd
import io

st.set_page_config(page_title="Practice Schedule Optimizer", layout="wide")
st.title("卓球部 練習シフト最適化 (Streamlit)")

st.markdown("アップロードした Excel ファイル（`r_time`, `w_len`, `day_limits` シートが必要）を読み込み、最適化を実行します。")

uploaded = st.file_uploader("Excelファイルをアップロード（Book2.xlsx を使う場合は空のまま実行可）", type=["xlsx"])
use_default = False
if uploaded is None:
    st.info("アップロードがない場合、同じフォルダにある 'Book2.xlsx' を探します。")
    use_default = True

with st.sidebar:
    st.header("重みパラメータ (任意)")
    w1 = st.number_input("授業後スコア (w1)", value=100.0)
    w2 = st.number_input("連続練習スコア (w2)", value=0.5)
    w3 = st.number_input("人数スコア (w3)", value=1.0)

run_button = st.button("最適化を実行")

# helper: run optimization given workbook path-like object
def run_optimization_from_workbook(book):
    # --- 以下はユーザー提供コードを Streamlit 向けに関数化したもの ---
    sheet_rt = book['r_time']
    sheet_len = book['w_len']
    sheet_day = book['day_limits']

    # 部員数を自動で取得
    num_members = 0
    for i in range(1, 100):
        if sheet_rt.cell(row=i + 1, column=1).value is None:
            break
        num_members += 1

    I = [i + 1 for i in range(num_members)]
    T = [i + 1 for i in range(8)]   # 時間帯 13〜20時 → 1〜8
    D = [i + 1 for i in range(4)]   # 曜日 火〜金 → 1〜4
    L = [1, 2, 3]
    L_s = {s: [l for l in L if s + l - 1 <= 8] for s in T}
    labels = list(string.ascii_uppercase)

    # r_time マップ
    r_map = {2: 1, 3: 3, 4: 5, 5: 7}
    r_time = {}
    for i in I:
        for d in D:
            val = sheet_rt.cell(row=i + 1, column=d + 1).value
            if val in r_map:
                r_time[i, d] = r_map[val]
            else:
                r_time[i, d] = None

    # availability の自動生成
    a = {}
    for i in I:
        for d in D:
            start = r_time[i, d]
            for t in T:
                a[i, t, d] = 0
            if start:
                # start から 2 つの時間枠分を可用とする（元コードの意図を維持）
                for t in range(start, 9):
                    a[i, t, d] = 1

    # w_len 読み込み
    w_len = {l: sheet_len.cell(row=l + 1, column=2).value for l in L}

    # day_limits の読み込み（チアの有無）
    day_chia = {}
    for d in D:
        day_chia[d] = sheet_day.cell(row=d+1, column=2).value

    # min/max
    day_min = {}
    day_max = {}
    for d in D:
        if day_chia[d]:
            day_min[d] = 3
            day_max[d] = 8
        else:
            day_min[d] = 3
            day_max[d] = 16

    # ideal / w_num
    ideal = {}
    w_num = {d: {} for d in D}
    N_range = range(3, num_members + 1)
    for d in D:
        if day_chia[d] is not None:
            ideal[d] = (day_min[d] + day_max[d]) // 2 + 1
            for n in N_range:
                w_num[d][n] = max(0.0, 1.0 - 0.1 * abs(n - ideal[d]))
        else:
            ideal[d] = None
            for n in N_range:
                w_num[d][n] = 1.0

    # モデル作成
    model = Model("practice_schedule")
    forbidden_start = [4, 6, 8]  # 16時,18時,20時

    x = {(i, t, d): model.add_var(name=f"x_{i},{t},{d}", var_type=BINARY) for i in I for t in T for d in D}
    y = {(t, d): model.add_var(var_type=BINARY) for t in T for d in D}
    z = {}
    for i in I:
        for d in D:
            for s in T:
                if s in forbidden_start:
                    continue
                for l in L_s[s]:
                    z[i, s, d, l] = model.add_var(var_type=BINARY)
    v = {(t, d, n): model.add_var(var_type=BINARY) for t in T for d in D for n in N_range}
    num_td = {(t, d): model.add_var(var_type=INTEGER) for t in T for d in D}

    # 制約: x <= a
    for i in I:
        for t in T:
            for d in D:
                model += x[i, t, d] <= a[i, t, d]

    # 週3回以上
    for i in I:
        model += xsum(x[i, t, d] for t in T for d in D) >= 3

    # 一日の最大3時間
    for i in I:
        for d in D:
            model += xsum(x[i, t, d] for t in T) <= 3

    # 人数制約
    for t in T:
        for d in D:
            available = sum(a[i, t, d] for i in I)
            if available < day_min[d]:
                for i in I:
                    model += x[i, t, d] == 0
                continue
            model += num_td[t, d] == xsum(x[i, t, d] for i in I)
            model += num_td[t, d] >= day_min[d]
            model += num_td[t, d] <= day_max[d]

    # z 一意
    for i in I:
        for d in D:
            model += xsum(z[i, s, d, l] for s in T if s not in forbidden_start for l in L_s[s]) <= 1

    # z->x, x->z
    for i in I:
        for d in D:
            for t in T:
                model += x[i, t, d] == xsum(
                    z[i, s, d, l]
                    for s in T if s not in forbidden_start
                    for l in L_s.get(s, [])
                    if s <= t < s + l
                )

    # 飛び飛び禁止
    for i in I:
        for d in D:
            for t in T[1:-1]:
                model += x[i, t, d] <= x[i, t - 1, d] + x[i, t + 1, d]
            t_first = T[0]
            if len(T) >= 2:
                model += x[i, t_first, d] <= x[i, t_first + 1, d]
            t_last = T[-1]
            if len(T) >= 2:
                model += x[i, t_last, d] <= x[i, t_last - 1, d]

    # v の一意性 + 人数に一致
    M = len(I)
    for t in T:
        for d in D:
            model += xsum(v[t, d, n] for n in N_range) == 1
            for n in N_range:
                model += num_td[t, d] - n <= (1 - v[t, d, n]) * M
                model += n - num_td[t, d] <= (1 - v[t, d, n]) * M

    # 目的関数
    term1 = xsum(
        x[i, r_time[i, d], d]
        for i in I
        for d in D
        if r_time[i, d] is not None and r_time[i, d] in T
    )
    term2 = xsum(w_len[l] * z[i, s, d, l] for i in I for d in D for s in T if s not in forbidden_start for l in L_s.get(s, []))
    term3 = xsum(w_num[d][n] * v[t, d, n] for t in T for d in D for n in N_range)
    model.objective = maximize(w1 * term1 + w2 * term2 + w3 * term3)

    # solve
    model.write("model.lp")
    status = model.optimize()

    result_info = {"status": model.status}

    # 結果出力
    if status in (OptimizationStatus.OPTIMAL, OptimizationStatus.FEASIBLE):
        if 'result' in book.sheetnames:
            book.remove(book['result'])
        result_sheet = book.create_sheet('result')

        # スコア集計
        score1 = sum(x[i, r_time[i, d], d].x for i in I for d in D if r_time[i, d] is not None and r_time[i, d] in T)
        score2 = sum(w_len[l] * z[i, s, d, l].x for i in I for d in D for s in T if s not in forbidden_start for l in L_s.get(s, []))
        score3 = sum(w_num[d][n] * v[t, d, n].x for t in T for d in D for n in N_range)
        weighted1 = w1 * score1
        weighted2 = w2 * score2
        weighted3 = w3 * score3
        total_score = weighted1 + weighted2 + weighted3

        result_info.update({
            'score1': score1, 'score2': score2, 'score3': score3,
            'weighted1': weighted1, 'weighted2': weighted2, 'weighted3': weighted3,
            'total_score': total_score
        })

        weekday_map = {1: '火', 2: '水', 3: '木', 4: '金'}
        for d in D:
            cell = result_sheet.cell(row=1, column=1 + d)
            cell.value = f"{weekday_map[d]}曜"
            cell.alignment = Alignment(horizontal='center')
            result_sheet.column_dimensions[cell.column_letter].width = 20
        for t in T:
            cell = result_sheet.cell(row=1 + t, column=1)
            cell.value = f"{12 + t}時"
            cell.alignment = Alignment(horizontal='center')
            result_sheet.column_dimensions['A'].width = 12

        for i in I:
            name = labels[i - 1]
            for t in T:
                for d in D:
                    if x[i, t, d].x >= 0.99:
                        row = 1 + t
                        col = 1 + d
                        cell = result_sheet.cell(row=row, column=col)
                        prev = cell.value if cell.value else ''
                        names = prev.split(',') if prev else []
                        if name not in names:
                            names.append(name)
                        cell.value = ",".join(names)
                        cell.alignment = Alignment(wrap_text=True, horizontal='center')
                        cell.font = Font(size=12)

        # 一時ファイルへ保存
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        book.save(tmp.name)
        result_info['output_path'] = tmp.name

    else:
        result_info['output_path'] = None

    return result_info


if run_button:
    try:
        if use_default:
            book = load_workbook('Book2.xlsx')
        else:
            # アップロードファイルを一時保存して読み込む
            tmpf = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
            tmpf.write(uploaded.getvalue())
            tmpf.flush()
            book = load_workbook(tmpf.name)

        with st.spinner('最適化モデルを作成・解いています...（数秒〜数分かかる場合があります）'):
            info = run_optimization_from_workbook(book)

        st.subheader('最適化結果')
        st.write('モデルステータス:', info.get('status'))
        if info.get('output_path'):
            st.metric('合計スコア', f"{info.get('total_score'):.2f}")
            st.write('目的関数内訳:')
            st.write(f"授業直後スコア: {info.get('weighted1'):.2f}")
            st.write(f"連続練習スコア: {info.get('weighted2'):.2f}")
            st.write(f"人数スコア: {info.get('weighted3'):.2f}")

            # 結果シートを DataFrame 化して表示
            df = pd.read_excel(info['output_path'], sheet_name='result', index_col=None)
            st.subheader('割当表 (result シート)')
            st.dataframe(df)

            # ダウンロードボタン
            with open(info['output_path'], 'rb') as f:
                data = f.read()
            st.download_button('結果（practice_result.xlsx）をダウンロード', data, file_name='practice_result.xlsx')
        else:
            st.error('実行可能な解が見つかりませんでした。')

    except Exception as e:
        st.exception(e)

else:
    st.info('準備ができたら「最適化を実行」ボタンを押してください。')
