import streamlit as st
import re
import os
import io
import pandas as pd
from datetime import date
from copy import deepcopy
from pptx import Presentation
from pptx.util import Pt as _Pt, Cm as _Cm

@st.cache_data
def parse_license_text(content: str):
    """라이선스 텍스트 정규식 파싱을 캐싱하여 성능 최적화"""
    pattern = re.compile(
        r'#?\s*(\d+)\.\s+([\w\s\(\)\/\-]+?):\s+(\d+)\s+task\(s\).*?expiring\s+([\d\-a-zA-Z]+).*?Customer\s*#\s*(\d+)',
        re.IGNORECASE | re.DOTALL
    )
    return pattern.findall(content)

@st.cache_data
def load_template_bytes():
    """PPTX 템플릿 로드를 캐싱하여 매 버튼 클릭 시 발생하던 디스크 I/O 최적화"""
    template_candidates = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "라이선스_확인서_템플릿.pptx"),
        "라이선스_확인서_템플릿.pptx",
    ]
    for tp in template_candidates:
        if os.path.exists(tp):
            with open(tp, "rb") as f:
                return f.read()
    return None



# ─── 페이지 설정 ────────────────────────────────────────────────
st.set_page_config(
    page_title="Ansys 라이선스 확인서 생성기",
    page_icon="📄",
    layout="centered"
)

# ─── 커스텀 CSS (다크 / 라이트 모드 자동 대응) ──────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;600;700;900&family=IBM+Plex+Mono:wght@400;600&display=swap');

  /* ── 색상 토큰: 다크 모드 (기본) ── */
  :root {
    --bg:       #05080f;
    --surface:  #090d1e;
    --surface2: #0e1535;
    --border:   #141d40;
    --text:     #c5c8d8;
    --subtext:  #4a5070;
    --accent:   #b8922a;
    --accent-h: #c9a23a;
    --input-bg: #080c1a;
    --sn-fg:    #05080f;
    --a04: rgba(184,146,42,.04);
    --a08: rgba(184,146,42,.08);
    --a10: rgba(184,146,42,.10);
    --a12: rgba(184,146,42,.12);
    --a30: rgba(184,146,42,.30);
    --a35: rgba(184,146,42,.35);
    --a40: rgba(184,146,42,.40);
  }

  /* ── 색상 토큰: 라이트 모드 ── */
  @media (prefers-color-scheme: light) {
    :root {
      --bg:       #f5f7ff;
      --surface:  #ffffff;
      --surface2: #eef1fb;
      --border:   #d8dced;
      --text:     #1a1d2e;
      --subtext:  #6b7194;
      --accent:   #7a5c10;
      --accent-h: #8e6c14;
      --input-bg: #f8f9fd;
      --sn-fg:    #ffffff;
      --a04: rgba(122,92,16,.04);
      --a08: rgba(122,92,16,.08);
      --a10: rgba(122,92,16,.10);
      --a12: rgba(122,92,16,.12);
      --a30: rgba(122,92,16,.30);
      --a35: rgba(122,92,16,.35);
      --a40: rgba(122,92,16,.40);
    }
  }

  /* 기본 메뉴, 헤더, 푸터 숨기기 */
  #MainMenu {visibility: hidden;}
  header {visibility: hidden;}
  footer {visibility: hidden;}

  html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; background-color: var(--bg) !important; color: var(--text) !important; }
  .stApp { background-color: var(--bg) !important; }
  .block-container { padding-top: 3.5rem !important; max-width: 780px !important; }
  .page-title { font-size: 2.2rem; font-weight: 900; line-height: 1.15; letter-spacing: -.03em; margin-bottom: .4rem; }
  .page-title em { font-style: normal; color: var(--accent); }
  .page-tag { display: inline-flex; align-items: center; gap: 8px; background: var(--a10); border: 1px solid var(--a35); color: var(--accent); font-family: 'IBM Plex Mono', monospace; font-size: .7rem; letter-spacing: .12em; text-transform: uppercase; padding: 5px 14px; border-radius: 20px; margin-bottom: 14px; }
  .page-sub { color: var(--subtext); font-size: .9rem; margin-bottom: 2rem; }
  .section-label { display: flex; align-items: center; gap: 10px; margin: 1.8rem 0 .9rem; font-family: 'IBM Plex Mono', monospace; font-size: .72rem; letter-spacing: .12em; text-transform: uppercase; color: var(--subtext); }
  .section-label .sn { width: 28px; height: 28px; background: var(--accent); border-radius: 7px; display: inline-flex; align-items: center; justify-content: center; font-size: .72rem; font-weight: 700; color: var(--sn-fg); }
  [data-testid="stFileUploader"] { background: var(--surface) !important; border: 2px dashed var(--border) !important; border-radius: 14px !important; padding: 1rem !important; }
  [data-testid="stFileUploader"]:hover { border-color: var(--accent) !important; }
  [data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; border: 1px solid var(--border); }
  thead tr th { background: var(--surface2) !important; color: var(--accent) !important; font-family: 'IBM Plex Mono', monospace !important; font-size: .72rem !important; letter-spacing: .1em !important; text-transform: uppercase !important; }
  tbody tr:hover td { background: var(--a04) !important; }
  [data-testid="stForm"] { background: var(--surface) !important; border: 1px solid var(--border) !important; border-radius: 16px !important; padding: 1.6rem !important; }
  input, textarea, select, [data-baseweb="select"] div { background: var(--input-bg) !important; border-color: var(--border) !important; color: var(--text) !important; border-radius: 9px !important; font-family: 'Noto Sans KR', sans-serif !important; }
  input:focus, select:focus { border-color: var(--accent) !important; box-shadow: 0 0 0 2px var(--a12) !important; }
  label { color: var(--accent) !important; font-family: 'IBM Plex Mono', monospace !important; font-size: .72rem !important; letter-spacing: .08em !important; text-transform: uppercase !important; }
  [data-testid="stFormSubmitButton"] button, .stButton button { background: var(--accent) !important; color: var(--sn-fg) !important; border: none !important; border-radius: 10px !important; font-weight: 900 !important; font-size: 1rem !important; padding: .85rem 2rem !important; width: 100% !important; transition: all .2s !important; }
  [data-testid="stFormSubmitButton"] button:hover { background: var(--accent-h) !important; transform: translateY(-2px) !important; }
  [data-testid="stDownloadButton"] button { background: var(--surface2) !important; color: var(--accent) !important; border: 1.5px solid var(--a40) !important; border-radius: 10px !important; font-weight: 700 !important; width: 100% !important; }
  [data-testid="stDownloadButton"] button:hover { background: var(--a10) !important; }
  [data-testid="stSuccess"] { background: rgba(0,168,120,.08) !important; border: 1.5px solid rgba(0,168,120,.3) !important; border-radius: 10px !important; color: #00a878 !important; }
  [data-testid="stError"]   { background: rgba(255,80,80,.08) !important; border: 1.5px solid rgba(255,80,80,.3) !important; border-radius: 10px !important; }
  [data-testid="stWarning"] { background: var(--a08) !important; border: 1.5px solid var(--a30) !important; border-radius: 10px !important; color: var(--accent) !important; }
  hr { border-color: var(--border) !important; }
  [data-baseweb="select"] { background: var(--input-bg) !important; border-radius: 9px !important; }
  [data-baseweb="popover"] { background: var(--surface2) !important; border: 1px solid var(--border) !important; }
  .stSpinner > div { border-top-color: var(--accent) !important; }
  details { background: var(--surface) !important; border: 1px solid var(--border) !important; border-radius: 10px !important; }
</style>
""", unsafe_allow_html=True)

# ─── 타이틀 ─────────────────────────────────────────────────────
st.markdown('<div class="page-tag">Ansys License Tool</div>', unsafe_allow_html=True)
st.markdown('<div class="page-title">라이선스 확인서 <em>자동 생성기</em></div>', unsafe_allow_html=True)
st.markdown('<div class="page-sub">라이선스 .txt 파일을 업로드하면 제품 목록을 자동 추출하고<br>공식 확인서 PDF 및 PPT를 즉시 다운로드할 수 있습니다.</div>', unsafe_allow_html=True)

# ─── STEP 1: 파일 업로드 ────────────────────────────────────────
st.markdown('<div class="section-label"><span class="sn">01</span> 라이선스 파일 업로드</div>', unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "라이선스 .txt 파일을 여기에 드래그하거나 클릭하여 업로드하세요",
    type=["txt"], label_visibility="collapsed"
)

df_licenses = pd.DataFrame()

# ─── STEP 2: 파싱 & 테이블 ─────────────────────────────────────
st.markdown('<div class="section-label"><span class="sn">02</span> 추출된 라이선스 목록</div>', unsafe_allow_html=True)

extracted_customer_no = None
extracted_warranty_end = None

MAX_FILE_SIZE = 5 * 1024 * 1024  # 5 MB

if uploaded_file:
    raw_bytes = uploaded_file.getvalue()
    if len(raw_bytes) > MAX_FILE_SIZE:
        st.error("❌ 파일 크기가 너무 큽니다. 5 MB 이하 파일만 허용됩니다.")
        uploaded_file = None
if uploaded_file:
    content = raw_bytes.decode("utf-8", errors="ignore")
    matches = parse_license_text(content)
    if matches:
        df_licenses = pd.DataFrame(matches, columns=["No", "Software (제품명)", "QTY (수량)", "ExpireDate", "CustomerNo"])
        if not df_licenses.empty:
            extracted_customer_no = df_licenses.iloc[-1]["CustomerNo"]
            extracted_warranty_end = df_licenses.iloc[-1]["ExpireDate"]
        df_display = df_licenses[["No", "Software (제품명)", "QTY (수량)"]]
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        st.success(f"✅ 총 {len(df_licenses)}개 라이선스 항목이 감지되었습니다.")
        df_licenses = df_display
    else:
        st.warning("⚠️ 파일에서 라이선스 패턴을 찾지 못했습니다.")
        with st.expander("파일 내용 미리보기"):
            st.text(content[:2000])
else:
    st.info("① 위에서 파일을 업로드하면 여기에 라이선스 목록이 표시됩니다.")

# ─── STEP 3: 확인서 폼 ──────────────────────────────────────────
st.markdown('<div class="section-label"><span class="sn">03</span> 확인서 정보 입력</div>', unsafe_allow_html=True)

with st.form("cert_form"):
    col1, col2 = st.columns(2)
    with col1:
        customer_name    = st.text_input("고 객 명",    placeholder="예) 한국 컴퍼니")
        customer_number  = st.text_input("고 객 번 호", value=extracted_customer_no or "", placeholder="예) 1213401")
    with col2:
        install_location = st.text_input("설 치 장 소",    placeholder="예) 서울시 강남구 테헤란로 123")
        warranty_start   = st.date_input("라이선스 보증기간 시작", value=date.today())

    warranty_end_default = date.today()
    if extracted_warranty_end:
        try:
            warranty_end_default = pd.to_datetime(extracted_warranty_end).date()
        except Exception:
            pass

    warranty_end = st.date_input("라이선스 보증기간 끝", value=warranty_end_default)

    license_type = st.selectbox(
        "라이선스 유형",
        ["Permanent License / LAN", "Lease License / WAN",
         "Maintenance License / LAN", "Permanent License / Regional WAN"]
    )
    issue_date = st.date_input("발행 일자", value=date.today())
    submitted  = st.form_submit_button("📋  확인서 생성 (PPT)", use_container_width=True)




# ══════════════════════════════════════════════════════════
#  PPTX 생성 — 원본 템플릿 기반
# ══════════════════════════════════════════════════════════
def create_pptx_from_template(customer, customer_no, location, lic_type,
                               warranty, iss_date, df, template_bytes):
    prs = Presentation(io.BytesIO(template_bytes))
    slide = prs.slides[0]

    ph_map = {
        10: customer,
        11: location,
        12: warranty,
        14: customer_no,
        18: str(iss_date.year),
        19: f"{iss_date.month:02d}",
        20: f"{iss_date.day:02d}",
        21: lic_type,
    }

    for shape in slide.shapes:
        try:
            ph = shape.placeholder_format
            if ph and ph.idx in ph_map:
                tf = shape.text_frame
                para = tf.paragraphs[0]
                if para.runs:
                    para.runs[0].text = ph_map[ph.idx]
                else:
                    para.text = ph_map[ph.idx]
        except Exception:
            pass

    if not df.empty:
        items = [(str(r["No"]), str(r["Software (제품명)"]),
                  str(r["QTY (수량)"]) + " task(s)")
                 for _, r in df.iterrows()]
    else:
        items = [("1", "라이선스 정보 없음", "-")]

    # 행수에 따라 폰트 자동 축소 ── 7행 이하: 10pt, 8~12: 9pt, 13+: 8pt
    n_items = len(items)
    cell_pt = 10 if n_items <= 7 else 9 if n_items <= 12 else 8

    max_data_h = _Cm(5.1)  # 테이블 전체 허용 최대 높이

    for shape in slide.shapes:
        if shape.has_table:
            tbl = shape.table
            while len(tbl.rows) < len(items):
                last_tr = tbl.rows[len(tbl.rows) - 1]._tr
                new_tr = deepcopy(last_tr)
                tbl._tbl.append(new_tr)
            
            # 행 개수가 많아 오버플로우 발생 시 테이블 높이 압축
            if len(tbl.rows) * _Cm(0.85) > max_data_h:
                new_h = int(max_data_h / len(tbl.rows))
                for i in range(len(tbl.rows)):
                    tbl.rows[i].height = new_h

            for r_idx, (no, sw, qty) in enumerate(items):
                row = tbl.rows[r_idx]
                for c_idx, val in enumerate([no, sw, qty]):
                    cell = row.cells[c_idx]
                    tf = cell.text_frame
                    para = tf.paragraphs[0]
                    if para.runs:
                        para.runs[0].text = val
                        para.runs[0].font.size = _Pt(cell_pt)
                    else:
                        para.text = val
            break

    # 슬라이드 1개만 남기기
    sld_id_list = prs.slides._sldIdLst
    sld_ids = list(sld_id_list)
    for i in range(1, len(sld_ids)):
        sld_id_list.remove(sld_ids[i])

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ─── 생성 버튼 처리 ─────────────────────────────────────────────
if submitted:
    if not customer_name or not install_location:
        st.error("❌ 고객명과 설치 장소를 모두 입력해 주세요.")
    else:
        # 템플릿 파일 로드 (캐싱된 함수 사용)
        template_bytes = load_template_bytes()

        with st.spinner("확인서 생성 중..."):
            safe_customer = re.sub(r'[\\/*?:"<>|]', '_', customer_name).strip()
            fname_base = f"Ansys_라이선스확인서_{safe_customer}_{issue_date.strftime('%Y%m%d')}"
            warranty_period_formatted = (
                f"{warranty_start.year}. {warranty_start.month:02d}. {warranty_start.day:02d}"
                f" ~ {warranty_end.year}. {warranty_end.month:02d}. {warranty_end.day:02d}"
            )

            # PPTX 생성
            _loc = install_location or "-"
            pptx_buf  = None
            pptx_err  = None
            if template_bytes:
                try:
                    pptx_buf = create_pptx_from_template(
                        customer_name, customer_number or "-",
                        _loc, license_type,
                        warranty_period_formatted,
                        issue_date, df_licenses,
                        template_bytes
                    )
                except Exception:
                    pptx_err = "PPT 생성 중 오류가 발생했습니다. 템플릿 파일 형식을 확인해 주세요."
            else:
                pptx_err = "템플릿 파일(라이선스_확인서_템플릿.pptx)이 앱 폴더에 없습니다."

        if pptx_buf:
            st.success("✅ 확인서가 생성되었습니다! 아래 버튼으로 다운로드하세요.")
            st.download_button(
                label="⬇️  PPT 템플릿 다운로드",
                data=pptx_buf,
                file_name=fname_base + ".pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )
        else:
            st.error(f"❌ PPT 오류: {pptx_err}")

st.markdown("---")
st.markdown(
    "<center style='color:var(--subtext); font-size:.75rem; font-family:monospace;'>"
    "Ansys License Certificate Generator · Dark &amp; Light Edition"
    "</center>",
    unsafe_allow_html=True
)
