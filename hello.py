import streamlit as st
import pandas as pd
from openpyxl import Workbook

st.title("Streamlit + Pandas + openpyxl テスト")

# 簡単なデータを作成
data = {
    "名前": ["Aさん", "Bさん", "Cさん"],
    "得点": [85, 92, 78]
}
df = pd.DataFrame(data)
st.write("📊 pandasのデータフレーム：")
st.dataframe(df)

# Excelに一時保存して読み込むテスト
excel_path = "sample.xlsx"
df.to_excel(excel_path, index=False)

# 読み込み確認
df_loaded = pd.read_excel(excel_path)
st.write("📖 openpyxl経由でExcelを読み込みました：")
st.dataframe(df_loaded)

st.success("openpyxl が正常に動作しました！")
