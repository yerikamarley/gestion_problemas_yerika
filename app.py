import streamlit as st

from app_ui import run_app


st.set_page_config(
    page_title="Control de casos e incidentes",
    page_icon=":material/fact_check:",
    layout="wide",
)

run_app()
