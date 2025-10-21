import streamlit as st
import pandas as pd

st.title("Streamlit + Pandas テスト")

# 簡単なデータフレームを作成
data = {
    "名前": ["Aさん", "Bさん", "Cさん"],
    "得点": [85, 92, 78]
}

df = pd.DataFrame(data)

# 表示
st.dataframe(df)

st.success("pandas と streamlit の動作確認ができました！")
