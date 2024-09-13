import streamlit as st
from io import BytesIO

st.set_page_config(page_title='🏠 TST Optibus Scripts')
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)