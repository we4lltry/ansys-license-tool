import streamlit as st

DARK_CSS = """
  html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; background-color: #05080f !important; color: #c5c8d8 !important; }
  .stApp { background-color: #05080f !important; }
  .stApp::before { content:''; display:block; height:3px; background:linear-gradient(90deg,#b8922a,#c9a23a 60%,#b8922a); position:fixed; top:0; left:0; right:0; z-index:9999; }
  .page-title { color: #c5c8d8; }
  .page-title em { font-style: normal; color: #b8922a; }
  .page-tag { display: inline-flex; align-items: center; gap: 8px; background: rgba(184,146,42,.10); border: 1px solid rgba(184,146,42,.35); color: #b8922a; font-family: 'IBM Plex Mono', monospace; font-size: .7rem; letter-spacing: .12em; text-transform: uppercase; padding: 5px 14px; border-radius: 20px; margin-bottom: 14px; }
  .page-sub { color: #4a5070; }
  .section-label { color: #4a5070; }
  .section-label .sn { background: #b8922a; color: #05080f; }
  [data-testid="stFileUploader"] { background: #090d1e !important; border: 2px dashed #141d40 !important; }
  [data-testid="stFileUploader"]:hover { border-color: #b8922a !important; }
  [data-testid="stDataFrame"] { border: 1px solid #141d40; }
  thead tr th { background: #0e1535 !important; color: #b8922a !important; }
  tbody tr:hover td { background: rgba(184,146,42,.04) !important; }
  [data-testid="stForm"] { background: #090d1e !important; border: 1px solid #141d40 !important; box-shadow: 0 4px 32px rgba(0,0,0,.4) !important; }
  input, textarea, select, [data-baseweb="select"] div { background: #080c1a !important; border-color: #141d40 !important; color: #c5c8d8 !important; }
  [data-baseweb="input"] { background: #080c1a !important; border-color: #141d40 !important; overflow: hidden !important; }
  [data-baseweb="input"] button, [data-baseweb="input"] [role="button"] { background: #080c1a !important; color: #b8922a !important; border: none !important; }
  input:focus, select:focus { border-color: #b8922a !important; box-shadow: 0 0 0 2px rgba(184,146,42,.12) !important; }
  label { color: #b8922a !important; }
  [data-testid="stFormSubmitButton"] button, .stButton button { background: #b8922a !important; color: #05080f !important; }
  [data-testid="stFormSubmitButton"] button:hover { background: #c9a23a !important; }
  [data-testid="stDownloadButton"] button { background: #0e1535 !important; color: #b8922a !important; border: 1.5px solid rgba(184,146,42,.4) !important; }
  [data-testid="stDownloadButton"] button:hover { background: rgba(184,146,42,.10) !important; }
  [data-testid="stSuccess"] { background: rgba(0,168,120,.08) !important; border: 1.5px solid rgba(0,168,120,.3) !important; color: #00a878 !important; }
  [data-testid="stError"]   { background: rgba(255,80,80,.08) !important; border: 1.5px solid rgba(255,80,80,.3) !important; }
  [data-testid="stWarning"] { background: rgba(184,146,42,.08) !important; border: 1.5px solid rgba(184,146,42,.3) !important; color: #b8922a !important; }
  hr { border-color: #141d40 !important; }
  [data-baseweb="select"] { background: #080c1a !important; }
  [data-baseweb="popover"] { background: #0e1535 !important; border: 1px solid #141d40 !important; }
  .stSpinner > div { border-top-color: #b8922a !important; }
  details { background: #090d1e !important; border: 1px solid #141d40 !important; }
  [data-testid="stSidebar"] { background: #07091a !important; border-right: 1px solid #141d40 !important; }
  [data-testid="stSidebar"] * { color: #c5c8d8 !important; }
  [data-testid="stSidebar"] a { color: #b8922a !important; }
  [data-testid="stSidebarNav"] a[aria-current="page"] { background: rgba(184,146,42,.12) !important; border-left: 3px solid #b8922a !important; }
  /* 테마 토글 버튼 */
  [data-testid="stHorizontalBlock"]:first-of-type [data-testid="stButton"] button { background: rgba(255,255,255,.04) !important; color: #7a82a0 !important; border: 1.5px solid #1e2750 !important; border-radius: 50px !important; font-size: .8rem !important; font-weight: 600 !important; letter-spacing: .06em !important; padding: .42rem 1.4rem !important; }
  [data-testid="stHorizontalBlock"]:first-of-type [data-testid="stButton"] button:hover { background: rgba(184,146,42,.08) !important; border-color: rgba(184,146,42,.5) !important; color: #b8922a !important; transform: none !important; }
  [data-testid="stHorizontalBlock"]:first-of-type [data-testid="column"] { display:flex !important; justify-content:center !important; }
"""

LIGHT_CSS = """
  html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; background-color: #f0f2f8 !important; color: #1a1e3d !important; }
  .stApp { background-color: #f0f2f8 !important; }
  .stApp::before { content:''; display:block; height:3px; background:linear-gradient(90deg,#8a6d1e,#a07f24 60%,#8a6d1e); position:fixed; top:0; left:0; right:0; z-index:9999; }
  .page-title { color: #1a1e3d; }
  .page-title em { font-style: normal; color: #8a6d1e; }
  .page-tag { display: inline-flex; align-items: center; gap: 8px; background: rgba(138,109,30,.10); border: 1px solid rgba(138,109,30,.35); color: #8a6d1e; font-family: 'IBM Plex Mono', monospace; font-size: .7rem; letter-spacing: .12em; text-transform: uppercase; padding: 5px 14px; border-radius: 20px; margin-bottom: 14px; }
  .page-sub { color: #6b748a; }
  .section-label { color: #6b748a; }
  .section-label .sn { background: #8a6d1e; color: #ffffff; }
  [data-testid="stFileUploader"] { background: #ffffff !important; border: 2px dashed #cdd1e4 !important; }
  [data-testid="stFileUploader"]:hover { border-color: #8a6d1e !important; }
  [data-testid="stDataFrame"] { border: 1px solid #cdd1e4; }
  thead tr th { background: #eef0f8 !important; color: #8a6d1e !important; }
  tbody tr:hover td { background: rgba(138,109,30,.04) !important; }
  [data-testid="stForm"] { background: #ffffff !important; border: 1px solid #dde1f0 !important; box-shadow: 0 2px 20px rgba(0,0,0,.07) !important; }
  input, textarea, select, [data-baseweb="select"] div { background: #f8f9fc !important; border-color: #cdd1e4 !important; color: #1a1e3d !important; }
  [data-baseweb="input"] { background: #f8f9fc !important; border-color: #cdd1e4 !important; overflow: hidden !important; }
  [data-baseweb="input"] button, [data-baseweb="input"] [role="button"] { background: #f8f9fc !important; color: #8a6d1e !important; border: none !important; }
  input:focus, select:focus { border-color: #8a6d1e !important; box-shadow: 0 0 0 2px rgba(138,109,30,.12) !important; }
  label { color: #8a6d1e !important; }
  [data-testid="stFormSubmitButton"] button, .stButton button { background: #8a6d1e !important; color: #ffffff !important; }
  [data-testid="stFormSubmitButton"] button:hover { background: #a07f24 !important; }
  [data-testid="stDownloadButton"] button { background: #eef0f8 !important; color: #8a6d1e !important; border: 1.5px solid rgba(138,109,30,.4) !important; }
  [data-testid="stDownloadButton"] button:hover { background: rgba(138,109,30,.10) !important; }
  [data-testid="stSuccess"] { background: rgba(0,168,120,.08) !important; border: 1.5px solid rgba(0,168,120,.3) !important; color: #007a58 !important; }
  [data-testid="stError"]   { background: rgba(220,50,50,.08) !important; border: 1.5px solid rgba(220,50,50,.3) !important; }
  [data-testid="stWarning"] { background: rgba(138,109,30,.08) !important; border: 1.5px solid rgba(138,109,30,.3) !important; color: #8a6d1e !important; }
  hr { border-color: #dde1f0 !important; }
  [data-baseweb="select"] { background: #f8f9fc !important; }
  [data-baseweb="popover"] { background: #ffffff !important; border: 1px solid #cdd1e4 !important; }
  .stSpinner > div { border-top-color: #8a6d1e !important; }
  details { background: #ffffff !important; border: 1px solid #cdd1e4 !important; }
  [data-testid="stSidebar"] { background: #e8eaf4 !important; border-right: 1px solid #cdd1e4 !important; }
  [data-testid="stSidebar"] * { color: #1a1e3d !important; }
  [data-testid="stSidebarNav"] a[aria-current="page"] { background: rgba(138,109,30,.10) !important; border-left: 3px solid #8a6d1e !important; }
  /* 테마 토글 버튼 */
  [data-testid="stHorizontalBlock"]:first-of-type [data-testid="stButton"] button { background: rgba(0,0,0,.03) !important; color: #8a90a8 !important; border: 1.5px solid #cdd1e4 !important; border-radius: 50px !important; font-size: .8rem !important; font-weight: 600 !important; letter-spacing: .06em !important; padding: .42rem 1.4rem !important; }
  [data-testid="stHorizontalBlock"]:first-of-type [data-testid="stButton"] button:hover { background: rgba(138,109,30,.06) !important; border-color: rgba(138,109,30,.4) !important; color: #8a6d1e !important; transform: none !important; }
  [data-testid="stHorizontalBlock"]:first-of-type [data-testid="column"] { display:flex !important; justify-content:center !important; }
"""

BASE_CSS = """
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;600;700;900&family=IBM+Plex+Mono:wght@400;600&display=swap');
  .block-container {{ padding-top: 4rem !important; max-width: 780px !important; }}
  .page-title {{ font-size: 2.2rem; font-weight: 900; line-height: 1.15; letter-spacing: -.03em; margin-bottom: .4rem; }}
  .section-label {{ display: flex; align-items: center; gap: 10px; margin: 1.8rem 0 .9rem; font-family: 'IBM Plex Mono', monospace; font-size: .72rem; letter-spacing: .12em; text-transform: uppercase; }}
  .section-label .sn {{ width: 28px; height: 28px; border-radius: 7px; display: inline-flex; align-items: center; justify-content: center; font-size: .72rem; font-weight: 700; }}
  .page-sub {{ font-size: .9rem; margin-bottom: 2rem; }}
  [data-testid="stFileUploader"] {{ border-radius: 14px !important; padding: 1rem !important; }}
  [data-testid="stDataFrame"] {{ border-radius: 12px; overflow: hidden; }}
  thead tr th {{ font-family: 'IBM Plex Mono', monospace !important; font-size: .72rem !important; letter-spacing: .1em !important; text-transform: uppercase !important; }}
  [data-testid="stForm"] {{ border-radius: 16px !important; padding: 1.6rem !important; }}
  input, textarea, select, [data-baseweb="select"] div {{ border-radius: 9px !important; font-family: 'Noto Sans KR', sans-serif !important; }}
  label {{ font-family: 'IBM Plex Mono', monospace !important; font-size: .72rem !important; letter-spacing: .08em !important; text-transform: uppercase !important; }}
  [data-testid="stFormSubmitButton"] button, .stButton button {{ border: none !important; border-radius: 10px !important; font-weight: 900 !important; font-size: 1rem !important; padding: .85rem 2rem !important; width: 100% !important; transition: all .2s !important; }}
  [data-testid="stFormSubmitButton"] button:hover {{ transform: translateY(-2px) !important; }}
  [data-testid="stDownloadButton"] button {{ border-radius: 10px !important; font-weight: 700 !important; width: 100% !important; }}
  [data-testid="stSuccess"], [data-testid="stError"], [data-testid="stWarning"] {{ border-radius: 10px !important; }}
  [data-baseweb="select"] {{ border-radius: 9px !important; }}
  details {{ border-radius: 10px !important; }}
"""


def init_theme():
    """세션 상태에서 테마 초기화"""
    if "theme" not in st.session_state:
        st.session_state.theme = "dark"


def render_theme_css():
    """현재 테마에 맞는 CSS를 페이지에 주입"""
    theme_css = DARK_CSS if st.session_state.theme == "dark" else LIGHT_CSS
    st.markdown(f"<style>{BASE_CSS}{theme_css}</style>", unsafe_allow_html=True)


def render_theme_toggle():
    """중앙 정렬 테마 토글 버튼 렌더링"""
    _, _tc, _ = st.columns([3, 2, 3])
    with _tc:
        _is_dark = st.session_state.theme == "dark"
        label = "☀️  라이트 모드" if _is_dark else "🌙  다크 모드"
        if st.button(label, key="theme_toggle", use_container_width=True):
            st.session_state.theme = "light" if _is_dark else "dark"
            st.rerun()


def accent_color():
    """현재 테마의 액센트 컬러 반환"""
    return "#b8922a" if st.session_state.theme == "dark" else "#8a6d1e"


def muted_color():
    """현재 테마의 뮤트 텍스트 컬러 반환"""
    return "#4a5070" if st.session_state.theme == "dark" else "#6b748a"
