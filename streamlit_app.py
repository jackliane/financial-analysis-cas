import streamlit as st
import pandas as pd

st.title("🎈 该网站主要用于对财务数据进行分析")

uploaded_file = st.file_uploader(
    "上传 Excel 文件",
    type=["xlsx"],
    help="支持 .xlsx 格式的文件"
)

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    st.success(f"✅ 文件上传成功！共 {len(df)} 行，{len(df.columns)} 列")
    st.dataframe(df)
else:
    st.info("👆 请上传一个 .xlsx 文件")
