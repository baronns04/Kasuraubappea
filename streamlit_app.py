import json
import streamlit as st
import streamlit.components.v1 as components
from app import DEFAULT_SHEET_ID, HTML_TEMPLATE, fetch_spreadsheet_data

st.set_page_config(
    page_title="Syura Bappeda - Buku Kas & Progres Keuangan",
    page_icon="🕌",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Sembunyikan header bawaan Streamlit agar tampilan penuh seperti web asli
st.markdown(
    """
    <style>
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        .block-container {padding: 0 !important; max-width: 100% !important;}
    </style>
    """,
    unsafe_allow_html=True,
)

data = fetch_spreadsheet_data(DEFAULT_SHEET_ID)
rendered_html = HTML_TEMPLATE.replace("{{ initial_data | safe }}", json.dumps(data))
components.html(rendered_html, height=1350, scrolling=True)
