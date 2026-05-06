import streamlit as st

from kaizhan.page_wrapper import render_kaizhan_page
from kuorong.app_sdr_expansion import render_expansion_page


st.set_page_config(page_title="4G宏站工具主页", layout="centered")

st.title("4G宏站工具主页")
st.markdown("---")

tool_name = st.radio(
    "请选择工具",
    ["4g宏站开站", "4g宏站扩容"],
    horizontal=True,
)

st.markdown("---")

if tool_name == "4g宏站开站":
    render_kaizhan_page()
else:
    render_expansion_page()
