import streamlit as st


st.set_page_config(
    page_title="Control de casos e incidentes",
    page_icon=":material/fact_check:",
    layout="wide",
)

from app_ui import run_app


run_app()
