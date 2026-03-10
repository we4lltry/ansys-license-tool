import streamlit as st

st.set_page_config(
    page_title="Ansys License Tool",
    page_icon="📄",
    layout="centered",
)

pg = st.navigation([
    st.Page("pages/라이선스_인증서.py", title="라이선스 인증서", icon="📄"),
    st.Page("pages/견적서.py",          title="견적서",          icon="📊"),
])
pg.run()
