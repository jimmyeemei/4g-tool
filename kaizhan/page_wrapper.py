from pathlib import Path
import runpy

import streamlit as st


def render_kaizhan_page():
    script_path = Path(__file__).with_name("app_SDR_FDD_gongxiang.py")
    original_set_page_config = st.set_page_config
    st.set_page_config = lambda *args, **kwargs: None
    try:
        runpy.run_path(str(script_path), run_name="__main__")
    finally:
        st.set_page_config = original_set_page_config
